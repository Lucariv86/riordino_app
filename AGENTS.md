# Riordino app - regole e obiettivi

## Struttura
- app.py: UI Streamlit
- reorder_engine.py: logiche di parsing e calcolo
- Output Excel: fogli riordino, scartati, summary

## Input (colonne fisse)
A=MARCA, D=CODICE ARTICOLO, F=DESCRIZIONE, G=GRP. MER., H=SCAR. AC, I=U.P.A., L=SCAR. AP, M=GIACENZA

## Regole di business
- Giacenze negative trattate come 0 per il calcolo fabbisogno.
- Venduto rolling 12 mesi stimato:
  venduto_12m = SCAR.AC + (SCAR.AP * (1 - frac_anno_trascorso))
  daily_demand = venduto_12m / 365
- Copertura default 30 giorni, selezionabile fino a 180.
- Target valore ordine (€): se impostato, aumenta qty finché raggiunge il target
  preferendo non superare 180 giorni; se sfora, warning + flag_over_180.
- ASPL: minimo 1 pezzo a stock (se giacenza_effettiva < 1 => ordina fino a 1).
- Output: includere foglio "scartati" con motivo_scarto.

## Performance
- Dataset tipico 500–600 righe: esecuzione in pochi secondi.
- Evitare apply(axis=1) ripetuti e iterrows su tutto il dataset.
