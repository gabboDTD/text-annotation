# Annotatore di titoli — Streamlit + pandas

App Streamlit per caricare un Excel, vedere **un titolo alla volta**, classificare i titoli come **Digitale**, **Non digitale** o lasciarli senza etichetta, aggiungere una nota facoltativa e scaricare l'Excel annotato.

## Avvio con Poetry

Se hai già un `pyproject.toml`:

```bash
poetry add "streamlit>=1.33,<2" "pandas>=2.1,<3" "openpyxl>=3.1,<4"
poetry run streamlit run app.py
```

Se non hai ancora un progetto Poetry, nella cartella estratta avvia prima `poetry init -n`.

## Alternativa con pip

```bash
pip install -r requirements.txt
streamlit run app.py
```

Apri `http://localhost:8501` (o l'indirizzo indicato da Streamlit).

## Utilizzo

1. Carica dall'app un file `.xlsx` con una colonna **`titolo`** (anche su un foglio diverso dal primo).
2. L'app seleziona il primo titolo non ancora etichettato.
3. Seleziona `Digitale`, `Non digitale` o `Non assegnata`, e compila eventualmente la nota.
   Nell’Excel, **Digitale** viene salvato come `1` e **Non digitale** come `0` (la scelta non assegnata resta vuota).
4. Premi **Salva**, **Salva e successivo** o **Precedente** per confermare le modifiche.
5. Scarica l'Excel annotato. Puoi ricaricarlo successivamente per continuare.

Le colonne **`label`** e **`note`** possono già esistere, o saranno create nell'Excel esportato. Le altre colonne e gli altri fogli vengono mantenuti.

## Dettagli tecnici

- Lettura: `pandas.read_excel(..., engine="openpyxl")`. Le label già presenti `1` e `0` vengono mostrate rispettivamente come **Digitale** e **Non digitale**.
- Esportazione: `pandas.ExcelWriter(..., mode="a", if_sheet_exists="overlay")`, con aggiornamento delle sole celle modificate del foglio originale. In questo modo si conservano altre colonne, formule, fogli e formattazione ordinaria (fatte salve le limitazioni del motore openpyxl con specifiche funzionalità Excel avanzate).
- La modifica di un file diverso o la chiusura della sessione può far perdere le annotazioni non esportate: scarica periodicamente una copia.
- Il formato supportato è `.xlsx`. La riga d'intestazione deve trovarsi entro le prime 15 righe, e il nome della colonna deve essere `titolo` (non conta la differenza maiuscole/minuscole).
