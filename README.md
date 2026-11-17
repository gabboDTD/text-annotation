# Annotatore di titoli — Streamlit

App Streamlit per caricare un file Excel e consultare i titoli, **uno alla volta**, assegnare una label **1** o **0**, aggiungere una nota facoltativa e scaricare l'Excel annotato.

## Avvio

1. Estrai lo ZIP e apri un terminale nella cartella `annotatore_titoli_streamlit`.
2. Installa le dipendenze:

   ```bash
   pip install -r requirements.txt
   ```

3. Avvia l'app:

   ```bash
   streamlit run app.py
   ```

4. Apri l'indirizzo mostrato nel terminale (di solito `http://localhost:8501`).

## Come si usa

1. Fai clic su **Carica il file Excel da annotare** e seleziona dal tuo computer un file `.xlsx` con una colonna `titolo`.
2. L'app apre il **primo titolo senza label** (o il primo titolo se tutti sono già etichettati).
3. Seleziona `1`, `0` o `Non assegnata` e, se utile, aggiungi una nota.
4. Premi **Salva e successivo →** per salvare e avanzare; **Salva** per restare sul titolo, **← Precedente** per tornare indietro salvando.
5. Puoi usare la barra laterale per saltare a un numero di titolo o al primo titolo ancora senza label; salva prima eventuali modifiche non confermate.
6. Premi **Scarica Excel annotato** per ottenere un nuovo file con suffisso `_annotato.xlsx`.
7. Per riprendere il lavoro un altro giorno, **carica nell'app l'Excel annotato** scaricato precedentemente.

## Formato dell'Excel

- Formato supportato: **`.xlsx`**.
- Deve esserci una colonna **`titolo`** nell'intestazione delle prime 15 righe di un foglio; non è necessario che sia la prima colonna, e maiuscole/minuscole non contano.
- Le colonne **`label`** e **`note`** possono già esistere o essere assenti: l'app le riconosce o le crea nel file esportato.
- Le label sono salvate come valori numerici **0** o **1**; puoi modificarle anche quando sono già presenti.
- Il file esportato conserva gli altri fogli e le altre colonne del file originale.
- Non serve mettere alcun Excel nella cartella del progetto: **il caricamento avviene direttamente nell'app**.

## Importante: salvataggio

Le annotazioni vengono tenute nella sessione Streamlit. **Scarica periodicamente l'Excel annotato**: chiudere la sessione o caricare un file diverso può far perdere le modifiche non esportate. Inoltre l'esportazione comprende soltanto le modifiche confermate con i pulsanti di salvataggio.
