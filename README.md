# riordino_app

## Test locale rapido

1. Crea ambiente virtuale e installa dipendenze:
   - `python -m venv .venv`
   - `source .venv/bin/activate`
   - `pip install -r requirements.txt`
2. Avvia l'app:
   - `streamlit run app.py`
3. Verifica la logica di riordino da terminale:
   - `python -m compileall app.py reorder_engine.py`
   - (opzionale) carica un file Excel di esempio e controlla che il foglio output contenga `riordino`, `scartati`, `summary`.
