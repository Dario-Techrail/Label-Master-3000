# Label Master 3000

**Breve descrizione**
Label Master 3000 è un'applicazione desktop per la gestione e stampa di etichette (serial numbers, componenti, preset) sviluppata in Python.

**Requisiti**
- Python 3.8+
- Dipendenze elencate in `requirements.txt`

**Installazione & avvio (Windows, cmd.exe)**
1. Apri il terminale nella cartella del progetto:
```cmd
cd "C:\Users\d.zanzonelli\Desktop\SOFTWARE VOULT\Label Master 3000"
```
2. (Opzionale) crea e attiva un virtualenv:
```cmd
python -m venv .venv
.\.venv\Scripts\activate
```
3. Installa le dipendenze:
```cmd
python -m pip install -r requirements.txt
```
4. Avvia l'app:
```cmd
python main.py
```

**Struttura principale**
- `main.py` – entrypoint dell'app
- `business_logic.py` – logica applicativa
- `ui_modules.py` – interfaccia utente
- `DB/` – dati e preset
- `Resources/` – risorse statiche

**Contribuire**
Apri un issue o crea una pull request; manteniamo il repository semplice e leggibile.

**Licenza**
Aggiungi qui la licenza (es. MIT) se desideri condividerla pubblicamente.

---
*Generato automaticamente — breve guida per iniziare.*
