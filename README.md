# Lokal Pseudonymiserare

⚡ AI-drivet program för att pseudonymisera dokument **helt lokalt** på din dator.

Detta verktyg:

* läser `.txt`, `.docx` eller `.pdf`
* hittar namn, personnummer och telefonnummer automatiskt
* låter dig granska och justera innan pseudonymisering
* kräver **ingen** molnkoppling eller konto
* kör **Gemma 3 4B modellen helt lokalt**

---

## 🧩 För användare

### 🔽 Steg 1 — Installera Ollama

Ladda ner och installera Ollama *innan* du startar programmet:

👉 [https://ollama.com/download](https://ollama.com/download)

Starta Ollama (den brukar starta automatiskt i bakgrunden).

---

### 🔽 Steg 2 — Ladda ner programmet

Ladda ner den senaste `.exe`-filen här:

👉 [https://github.com/Jtensetti/pseudonymisera/releases/latest](https://github.com/Jtensetti/pseudonymisera/releases/latest)

Fil:

```
pseudonymiserare.exe
```

---

### 🔽 Steg 3 — Starta programmet

Dubbelklicka på `.exe`-filen.

Första gången kan Ollama behöva ladda modellen:

```
gemma3:4b
```

Den är stor och kan ta några minuter att "värma upp".
När webbläsaren öppnas → du är igång! 🎉

---

### 🧠 Systemkrav

Minsta rekommendation:

* Windows 10/11
* 16 GB RAM (gärna 32 GB för snabbare modelluppvärmning)
* ~8 GB ledigt diskutrymme för modellfilen

Stöder **endast Windows** i nuvarande version.

---

### 🔒 Allt sker lokalt

✔ Ingen data skickas till internet
✔ Ingen uppladdning till chattbotar eller moln
✔ Tryggt för kommunal användning och känsliga dokument

---

## 🧑‍💻 För utvecklare

### Klona projektet

```bash
git clone https://github.com/Jtensetti/pseudonymisera.git
cd pseudonymisera
```

### Installera beroenden

```bash
pip install -r requirements.txt
```

(Se till att Ollama är installerat och igång.)

### Kör direkt i Python

```bash
python pseudonymiserare.py
```

### Bygg egen .exe (PyInstaller)

```bash
pyinstaller --noconfirm --clean pseudonymiserare.spec
```

Resultatet hamnar i:

```
build/pseudonymiserare/
```

---

## 📝 Licens

Apache 2.0 — fritt att använda och vidareutveckla!

---

## ❤️ Bidrag

PR:er, förbättringar och issue-rapporter välkomnas!
