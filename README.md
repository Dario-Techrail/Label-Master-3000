# Label Master 3000

**Breve descrizione**
Label Master 3000 è un'applicazione desktop per la gestione e stampa di etichette (serial numbers, componenti, preset) sviluppata in Python.

**Requisiti**
- Python 3.8+
- Dipendenze elencate in `requirements.txt`

**Installazione & avvio (Windows, cmd.exe)**
1. Installa le dipendenze:
```cmd
python -m pip install -r requirements.txt
```
2. Avvia l'app:
```cmd
python main.py
```

**Struttura principale**
- `main.py` – entrypoint dell'app
- `business_logic.py` – logica applicativa
- `ui_modules.py` – interfaccia utente
- `DB/` – dati e preset
- `Resources/` – risorse statiche

