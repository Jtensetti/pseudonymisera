import os
import re
import json
import tempfile
import requests
import subprocess
import logging
import sys
import webbrowser

import gradio as gr
import uvicorn
from docx import Document

OLLAMA_MODEL = "gemma3:4b"

# -------------------------
# Hjälpfunktioner: konsol-wait
# -------------------------

def wait_for_enter():
    """
    Vänta på ENTER om programmet körs i en riktig terminal.
    I PyInstaller --noconsole / GUI-läge finns ofta ingen stdin.
    """
    try:
        if sys.stdin and sys.stdin.isatty():
            input("\nTryck ENTER för att stänga fönstret...")
    except Exception:
        # Ignorera alla fel här – vi vill inte krascha bara för att stdin saknas
        pass


# -------------------------
# Ollama-modellhantering
# -------------------------

def ensure_ollama_model(model=OLLAMA_MODEL):
    """
    Säkerställ att Ollama är igång, att modellen finns lokalt
    och värm upp modellen så att första riktiga anropet inte time:ar.
    """
    # 1. Kolla att Ollama svarar
    try:
        resp = requests.get("http://localhost:11434/api/tags", timeout=10)
        resp.raise_for_status()
    except Exception as e:
        raise RuntimeError(
            "Ollama verkar inte vara igång.\n"
            "Starta Ollama-appen (eller Ollama-tjänsten) och försök igen."
        ) from e

    # 2. Kolla om modellen redan finns
    try:
        data = resp.json()
        models = data.get("models", [])
        names = [m.get("name") for m in models if isinstance(m, dict)]
    except Exception:
        names = []

    if model not in names:
        print(f"🧩 Modellen '{model}' saknas. Försöker ladda ner med 'ollama pull {model}'...")
        try:
            # Detta kräver att 'ollama' finns i PATH på datorn
            subprocess.run(
                ["ollama", "pull", model],
                check=True,
            )
            print(f"✅ Klar: modellen '{model}' nedladdad.")
        except Exception as e:
            raise RuntimeError(
                f"Kunde inte ladda ner modellen '{model}'.\n"
                f"Kontrollera att Ollama är korrekt installerat och att kommandot "
                f"'ollama pull {model}' fungerar.\n"
                f"Teknisk info: {e}"
            ) from e

    # 3. Warmup – litet generate-anrop så modellen laddas i RAM innan första riktiga jobbet
    print("🔥 Värmer upp modellen (första gången kan detta ta en stund)...")
    try:
        warmup_resp = requests.post(
            "http://localhost:11434/api/generate",
            json={
                "model": model,
                "prompt": "Värm upp modellen kort.",
                "stream": False,
                "options": {
                    "num_predict": 1  # bara en token, vi bryr oss inte om svaret
                },
            },
            timeout=900,
        )
        warmup_resp.raise_for_status()
        print("✅ Modellen är uppvärmd och redo.\n")
    except Exception as e:
        raise RuntimeError(
            "Kunde prata med Ollama men misslyckades med att värma upp modellen.\n"
            "Testa att köra 'ollama run gemma3:12b' manuellt en gång och prova sedan igen.\n"
            f"Teknisk info: {e}"
        ) from e


# -------------------------
# Hjälpfunktioner för text
# -------------------------

def read_document(file_obj):
    """Läs text ur .txt eller .docx."""
    filename = file_obj.name
    ext = os.path.splitext(filename)[1].lower()

    if ext == ".txt":
        with open(filename, "r", encoding="utf-8", errors="ignore") as f:
            return f.read(), ".txt"
    elif ext == ".docx":
        doc = Document(filename)
        text = "\n".join(p.text for p in doc.paragraphs)
        return text, ".docx"
    else:
        raise ValueError("Endast .txt och .docx stöds i denna version.")


def merge_persons(persons):
    """
    Slår ihop personposter som sannolikt är samma person,
    baserat på canonical_name och överlappande mentions.
    """
    merged = []
    for p in persons:
        canonical = (p.get("canonical_name") or "").strip()
        mentions = [
            m.strip()
            for m in (p.get("mentions") or [])
            if isinstance(m, str) and m.strip()
        ]

        if not canonical and not mentions:
            continue

        placed = False
        for m in merged:
            m_canonical = m.get("canonical_name") or ""
            m_mentions = m.get("mentions") or []

            # Samma canonical_name
            if canonical and m_canonical and canonical == m_canonical:
                all_mentions = set(m_mentions) | set(mentions)
                if canonical:
                    all_mentions.add(canonical)
                m["mentions"] = list(all_mentions)
                placed = True
                break

            # Överlappande mentions
            if set(mentions) & set(m_mentions):
                all_mentions = set(m_mentions) | set(mentions)
                if not m_canonical and canonical:
                    m["canonical_name"] = canonical
                if canonical:
                    all_mentions.add(canonical)
                m["mentions"] = list(all_mentions)
                placed = True
                break

        if not placed:
            merged.append(
                {
                    "canonical_name": canonical,
                    "mentions": mentions if mentions else ([canonical] if canonical else []),
                }
            )

    return merged


def call_ollama_extract(text, model=OLLAMA_MODEL):
    """
    Anropar Ollama för att:
    - hitta personer och alla sätt de omnämns på
    - hitta personnummer/födelsedatum
    - hitta telefonnummer

    Returnerar ett dict:
    {
      "persons": [...],
      "personnummer": [...],
      "phones": [...]
    }
    """
    # Begränsa längden för säkerhets skull (kan justeras upp vid behov)
    limited_text = text[:8000]

    prompt = f"""
Du analyserar svensk löptext och extraherar personuppgifter (PII).

Din uppgift är att hitta:
- alla personer och alla sätt de omnämns på
- alla personnummer/födelsedatum kopplade till personer
- alla telefonnummer

Viktiga regler:

1. Personer
- Identifiera alla personer i texten, även om de bara nämns med förnamn.
- "canonical_name" ska vara det mest kompletta namnet som finns i texten för personen
  (t.ex. "Emma Sandberg" om både "Emma" och "Emma Sandberg" förekommer,
   annars bara "Emma" om förnamnet är det enda som finns).
- Om "canonical_name" saknas för en person men det finns "mentions",
  ska du låta den längsta mention-strängen representera personen.
- "mentions" ska innehålla alla exakta textsträngar som syftar på personen,
  t.ex. ["Emma Sandberg", "Emma"]. Använd exakt stavning från texten.
- Tänk på att namn kan ha ett genitiv-s "Emmas", det ska ändå ingå i "mentions" för den personen.
- Ta inte med pronomen (han, hon, de, honom, henne, etc.).
- Ta med smeknamn och namn i citat, t.ex. "Oskar" i "den där killen som jobbar med API-nycklarna, Oskar".
- Om texten innehåller en beskrivande fras som tydligt syftar på en specifik person,
  t.ex. "den där killen som jobbar med API-nycklarna" eller "chefen som alltid ringer vid sju",
  ska hela frasen också ingå i "mentions" för den personen.
- Om en fras både innehåller en beskrivning och ett namn, t.ex.
  "den där killen som jobbar med API-nycklarna, Oskar, tror jag?",
  ska både hela frasen och namnet ("Oskar") finnas med i "mentions".
- Hitta inte på egna namn som inte finns i texten.

2. Personnummer och födelsedatum
- Lista alla strängar som ser ut som svenska personnummer eller födelsedatum kopplade till en person.
- Ta med format som:
  - YYYYMMDD-XXXX
  - YYYYMMDDXXXX
  - YYMMDD-XXXX
  - YYMMDDXXXX
  - YYYY-MM-DD (om det är tydligt kopplat till en person, t.ex. står efter ett namn eller inom parentes).
- Returnera strängarna exakt som de står i texten.

3. Telefonnummer
- Lista alla strängar som ser ut som telefonnummer, t.ex.:
  - 07XXXXXXXX
  - 0X-XXX XX XX
  - 0XX-XXXXXXX
  - +46 7X XXX XX XX
  - +467XXXXXXXX
- Inkludera varianter med mellanslag och bindestreck.
- Exkludera:
  - klockslag (t.ex. "07.43", "16.28")
  - fakturanummer
  - versionsnummer (t.ex. "3.7.14")
  - andra rena nummer som inte tydligt är telefonnummer.

4. Allmänt
- Alla värden i "mentions", "personnummer" och "phones" ska vara exakta substrängar från texten.
- Hitta så många korrekta träffar som möjligt. Missa hellre inget än att vara överförsiktig.
- Svara med välformad JSON som går att parsa direkt.

Text:
\"\"\"{limited_text}\"\"\"


Svara ENBART med giltig JSON på formen:

{{
  "persons": [
    {{
      "canonical_name": "Anna Björk",
      "mentions": ["Anna Björk", "Anna"]
    }}
  ],
  "personnummer": ["1986-05-19", "YYYYMMDD-XXXX"],
  "phones": ["+46 70 123 45 67", "070-123 45 67"]
}}

Inga kommentarer, ingen extra text, endast JSON.
"""

    resp = requests.post(
        "http://localhost:11434/api/generate",
        json={
            "model": model,
            "prompt": prompt,
            "stream": False,
        },
        timeout=900,
    )
    resp.raise_for_status()
    data = resp.json()
    raw = data.get("response", "").strip()

    # Enkel städning om modellen skulle råka lägga till ```json ``` runt
    if raw.startswith("```"):
        raw = raw.strip("`")
        # ta bort ev. "json" i första raden
        if raw.lower().startswith("json"):
            raw = raw[4:].strip()
        # hitta första { och sista }
        start = raw.find("{")
        end = raw.rfind("}")
        if start != -1 and end != -1:
            raw = raw[start:end+1].strip()

    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError:
        # Om modellen spårar ur, returnera tom struktur
        return {"persons": [], "personnummer": [], "phones": []}

    persons = parsed.get("persons", []) or []
    personnummer = parsed.get("personnummer", []) or []
    phones = parsed.get("phones", []) or []

    # Städa upp personer
    clean_persons = []
    for p in persons:
        canonical = (p.get("canonical_name") or "").strip()
        mentions = p.get("mentions") or []
        mentions = [m.strip() for m in mentions if isinstance(m, str) and m.strip()]

        # Om canonical saknas men vi har mentions → välj längsta mention som canonical
        if not canonical and mentions:
            canonical = max(mentions, key=len)

        if not canonical and not mentions:
            continue

        if canonical and canonical not in mentions:
            mentions.insert(0, canonical)

        clean_persons.append(
            {
                "canonical_name": canonical,
                "mentions": mentions,
            }
        )

    # Slå ihop dubbletter av samma person (t.ex. "Oskar" + "Oskar Nyblom")
    clean_persons = merge_persons(clean_persons)

    personnummer = [
        s.strip() for s in personnummer if isinstance(s, str) and s.strip()
    ]
    phones = [s.strip() for s in phones if isinstance(s, str) and s.strip()]

    return {
        "persons": clean_persons,
        "personnummer": personnummer,
        "phones": phones,
    }


def regex_find_entities(text):
    """Komplettera med regex för personnummer/telefonnummer."""
    entities = []

    # Svenska personnummer: 6+4 eller 8+4 samt födelsedatum YYYY-MM-DD
    pnr_pattern = r"\b\d{6}[-+]\d{4}\b|\b\d{8}-\d{4}\b|\b\d{4}-\d{2}-\d{2}\b"
    for match in re.findall(pnr_pattern, text):
        entities.append({"type": "personnummer", "value": match})

    # Telefonnr: grovt mönster för +46 / 0 följt av siffror och mellanslag
    phone_pattern = r"\b(?:\+46|0)\d[\d\s\-]{6,}\d\b"
    for match in re.findall(phone_pattern, text):
        entities.append({"type": "phone", "value": match})

    return entities


def deduplicate_entities(entities):
    """Ta bort dubbletter (type + value)."""
    seen = set()
    unique = []
    for e in entities:
        key = (e["type"], e["value"])
        if key not in seen:
            seen.add(key)
            unique.append(e)
    return unique


# -------------------------
# Pseudonymisering
# -------------------------

def build_mapping_from_table(table_rows):
    """
    Tar in tabellen från Gradio (kan vara en pandas DataFrame eller en lista av listor)
    och bygger en mapping {original: pseudonym}.

    - Kolumner förväntas vara: [Typ, Original, Pseudonym]
    - Om Pseudonym är tom genereras den automatiskt.
    """

    # 1. Konvertera ev. DataFrame -> lista av listor
    if hasattr(table_rows, "values"):  # pandas DataFrame
        rows = table_rows.values.tolist()
    else:
        rows = table_rows or []

    mapping = {}
    counters = {"name": 1, "personnummer": 1, "phone": 1}

    if rows is None or len(rows) == 0:
        return mapping

    for row in rows:
        if not row:
            continue

        # Se till att vi har minst 2 kolumner (typ + original)
        # Pseudonymkolumnen kan saknas, då blir den tom.
        t = row[0] if len(row) > 0 else ""
        original = row[1] if len(row) > 1 else ""
        pseudo = row[2] if len(row) > 2 else ""

        # Om något av dessa är NaN/None etc, gör dem till tomma strängar
        t = "" if t is None else str(t).strip()
        original = "" if original is None else str(original).strip()
        pseudo = "" if pseudo is None else str(pseudo).strip()

        if not original:
            continue  # hoppa rader utan originalvärde

        if not t:
            t = "name"  # default

        # Auto-generera pseudonym om tom
        if not pseudo:
            if t == "name":
                pseudo = f"Person_{counters['name']:03d}"
                counters["name"] += 1
            elif t == "personnummer":
                pseudo = f"PNR_{counters['personnummer']:03d}"
                counters["personnummer"] += 1
            elif t == "phone":
                pseudo = f"TEL_{counters['phone']:03d}"
                counters["phone"] += 1
            else:
                pseudo = f"X_{counters['name']:03d}"
                counters["name"] += 1

        mapping[original] = pseudo

    return mapping


def pseudonymize_text(text, mapping):
    """
    Ersätt alla förekomster av original med pseudonym.
    Lite försiktigare för namn (ordgränser).
    """
    new_text = text

    # Sortera så att längre strängar ersätts först (minskar risk för delmatchning)
    items = sorted(mapping.items(), key=lambda kv: len(kv[0]), reverse=True)

    for original, pseudo in items:
        # Heuristik: om det ser ut som ett personnummer eller telefon, byt rakt av
        if re.match(r"^\d{6}[-+]\d{4}$|^\d{8}-\d{4}$|^\d{4}-\d{2}-\d{2}$", original):
            pattern = re.escape(original)
        elif re.match(r"^(?:\+46|0)\d", original.replace(" ", "").replace("-", "")):
            pattern = re.escape(original)
        else:
            # Behandla som namn/fras – omge med ordgränser
            pattern = r"\b" + re.escape(original) + r"\b"

        new_text = re.sub(pattern, pseudo, new_text)

    return new_text


def write_temp_file(text, ext=".txt"):
    """Skriv text till temporär fil och returnera sökvägen."""
    fd, path = tempfile.mkstemp(suffix=ext)
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        f.write(text)
    return path


# -------------------------
# Gradio-callbacks
# -------------------------

def analyze_document(file_obj):
    if file_obj is None:
        return [], "Ingen fil uppladdad ännu."

    try:
        text, ext = read_document(file_obj)
    except Exception as e:
        return [], f"Fel vid läsning av dokument: {e}"

    # Hämta strukturerad info från Ollama
    ai_result = call_ollama_extract(text)
    persons = ai_result.get("persons", [])
    pnr_list = ai_result.get("personnummer", []) or []
    phone_list = ai_result.get("phones", []) or []

    # Komplettera personnummer/telefoner med regex (säkerhetsbälte)
    rx_entities = regex_find_entities(text)
    for e in rx_entities:
        if e["type"] == "personnummer":
            pnr_list.append(e["value"])
        elif e["type"] == "phone":
            phone_list.append(e["value"])

    rows = []
    status_parts = []

    # 1. Personer: en pseudonym per person, samma för alla omnämnanden
    person_counter = 1
    total_mentions = 0
    for person in persons:
        pseudo = f"Person_{person_counter:03d}"
        person_counter += 1

        mentions = person.get("mentions", [])
        added = set()
        for mention in mentions:
            if not mention:
                continue
            if mention in added:
                continue
            added.add(mention)
            rows.append(["name", mention, pseudo])
            total_mentions += 1

    if persons:
        status_parts.append(
            f"hittade {len(persons)} person(er) med totalt {total_mentions} namn-omnämnanden"
        )

    # 2. Personnummer – en rad per unikt nummer (pseudonym tom -> auto-genereras senare)
    unique_pnr = sorted(set(pnr_list))
    for p in unique_pnr:
        rows.append(["personnummer", p, ""])
    if unique_pnr:
        status_parts.append(f"{len(unique_pnr)} personnummer")

    # 3. Telefoner – en rad per unikt nummer
    unique_phones = sorted(set(phone_list))
    for t in unique_phones:
        rows.append(["phone", t, ""])
    if unique_phones:
        status_parts.append(f"{len(unique_phones)} telefonnummer")

    if not rows:
        msg = (
            "Inga namn, personnummer eller telefonnummer hittades automatiskt. "
            "Du kan ändå lägga till rader manuellt."
        )
    else:
        msg = (
            "Träffar: "
            + ", ".join(status_parts)
            + ". Justera listan vid behov och klicka sedan på 'Skapa pseudonymiserat dokument'."
        )

    return rows, msg


def run_pseudonymization(file_obj, table_rows):
    if file_obj is None:
        return "Ingen fil.", None

    try:
        text, ext = read_document(file_obj)
    except Exception as e:
        return f"Fel vid läsning av dokument: {e}", None

    # Bygg mapping från tabellen (oavsett om det är lista eller DataFrame)
    mapping = build_mapping_from_table(table_rows)

    if not mapping:
        return (
            "Ingen mapping kunde skapas från tabellen. "
            "Kontrollera att du klickat på 'Analysera dokument' först och att tabellen inte är helt tom.",
            None,
        )

    new_text = pseudonymize_text(text, mapping)
    out_path = write_temp_file(new_text, ext=".txt")
    preview = new_text[:1000]

    return preview, out_path


# -------------------------
# Bygg Gradio-gränssnitt
# -------------------------

with gr.Blocks(title="Lokal Pseudonymiserare") as demo:
    gr.Markdown(
        "# Lokal pseudonymiserare\n"
        "Dra in ett dokument, låt en lokal AI (Ollama) hitta namn/personnummer/telefonnummer, "
        "justera listan och generera ett pseudonymiserat dokument."
    )

    with gr.Row():
        file_in = gr.File(
            label="Ladda upp dokument (.txt eller .docx)",
            file_types=[".txt", ".docx"],
        )
        analyze_btn = gr.Button("Analysera dokument")

    entities_table = gr.Dataframe(
        headers=["Typ", "Original", "Pseudonym"],
        datatype=["str", "str", "str"],
        row_count=(0, "dynamic"),
        interactive=True,
        label="Identifierade (och redigerbara) entiteter",
    )

    status_box = gr.Markdown("Statusmeddelanden visas här.")

    analyze_btn.click(
        fn=analyze_document,
        inputs=file_in,
        outputs=[entities_table, status_box],
    )

    gr.Markdown("### Steg 2: Pseudonymisera dokumentet")
    pseud_btn = gr.Button("Skapa pseudonymiserat dokument")

    preview_box = gr.Textbox(
        label="Förhandsgranskning (första 1000 tecken)",
        lines=15,
    )
    download_out = gr.File(label="Ladda ner pseudonymiserad textfil")

    pseud_btn.click(
        fn=run_pseudonymization,
        inputs=[file_in, entities_table],
        outputs=[preview_box, download_out],
    )


# -------------------------
# Uvicorn-loggpatch för PyInstaller-miljö
# -------------------------

def patch_uvicorn_logging():
    """
    PyInstaller + uvicorn kan krascha när uvicorns DefaultFormatter försöker
    använda extra kwargs (t.ex. use_colors) mot logging.Formatter.
    Vi ersätter formattern 'default' med logging.Formatter och rensar bort
    alla nycklar som inte stöds.
    """
    try:
        from uvicorn.config import LOGGING_CONFIG
        fmt_default = LOGGING_CONFIG.get("formatters", {}).get("default")
        if isinstance(fmt_default, dict):
            # Använd standard-Formatter istället för Uvicorns egen
            fmt_default["()"] = "logging.Formatter"

            # Endast tillåtna nycklar för logging.Formatter via dictConfig
            allowed_keys = {"()", "fmt", "datefmt", "style", "validate"}

            for key in list(fmt_default.keys()):
                if key not in allowed_keys:
                    fmt_default.pop(key, None)
    except Exception as e:
        # Om patchen misslyckas vill vi inte krascha programmet
        print(f"Kunde inte patcha uvicorn-logging: {e}")


if __name__ == "__main__":
    patch_uvicorn_logging()
    print("🚀 Startar Lokal Pseudonymiserare...")
    print("Öppna webbläsaren och gå till http://127.0.0.1:7860 om den inte öppnas automatiskt.")
    print("Stäng inte detta fönster förrän du laddat ner den pseudonymiserade filen.\n")

    try:
        ensure_ollama_model()
    except RuntimeError as e:
        print(e)
        wait_for_enter()
    else:
        try:
            # Gradio startar servern och öppnar webbläsaren
            demo.launch(
                server_name="127.0.0.1",
                server_port=7860,
                inbrowser=True,   # öppna webbläsare automatiskt
                share=False,
            )
        except Exception:
            import traceback
            print("Ett oväntat fel uppstod i pseudonymiseraren:\n")
            traceback.print_exc()
            wait_for_enter()
