"""App Streamlit per annotare, un titolo alla volta, un file Excel.

Avvio: streamlit run app.py
Dipendenze: streamlit, openpyxl
"""

from __future__ import annotations

from io import BytesIO
from pathlib import Path
import hashlib
import html

import streamlit as st
from openpyxl import load_workbook


EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

st.set_page_config(
    page_title="Annotatore di titoli",
    page_icon="🏷️",
    layout="centered",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
      .block-container { max-width: 900px; padding-top: 2rem; }
      .title-panel {
        background: rgba(104, 140, 193, .08);
        border: 1px solid rgba(104, 140, 193, .26);
        border-radius: 16px;
        padding: 1.55rem 1.7rem;
        margin: .6rem 0 1.4rem;
      }
      .title-eyebrow { font-size: .82rem; font-weight: 700; opacity: .7;
        text-transform: uppercase; letter-spacing: .07em; margin-bottom: .7rem; }
      .title-content { font-size: 1.4rem; font-weight: 600; line-height: 1.5;
        overflow-wrap: anywhere; }
      .help-text { font-size: .87rem; opacity: .72; margin-top: .7rem; }
      div[data-testid="stMetric"] { padding: .15rem .35rem; }
    </style>
    """,
    unsafe_allow_html=True,
)


def normalize_label(value: object) -> int | None:
    """Accetta label preesistenti numeriche o stringa 0/1."""
    if isinstance(value, bool):
        return None
    if value in (0, 1, "0", "1"):
        return int(value)
    return None


def read_records(raw_bytes: bytes) -> tuple[str, int, dict[str, int], list[dict]]:
    """Trova 'titolo' nell'intestazione e conserva gli indici di riga Excel."""
    workbook = load_workbook(BytesIO(raw_bytes), read_only=True, data_only=False)
    try:
        selected = None
        for worksheet in workbook.worksheets:
            for line in worksheet.iter_rows(min_row=1, max_row=min(15, worksheet.max_row)):
                names = {
                    str(cell.value).strip().casefold(): cell.column
                    for cell in line
                    if isinstance(cell.value, str)
                }
                if "titolo" in names:
                    selected = (worksheet, line[0].row, names)
                    break
            if selected:
                break

        if selected is None:
            raise ValueError(
                "Non trovo una colonna chiamata 'titolo' nelle prime 15 righe "
                "di nessun foglio. Controlla l'intestazione dell'Excel."
            )

        sheet, header_row, names = selected
        title_col = names["titolo"]
        # Il file fornito contiene già 'label' e 'note'. Per altri file,
        # le colonne mancanti verranno create soltanto durante l'esportazione.
        label_col = names.get("label", sheet.max_column + 1)
        note_col = names.get("note", max(sheet.max_column, label_col) + 1)
        cols = {"titolo": title_col, "label": label_col, "note": note_col}

        records = []
        for row in sheet.iter_rows(min_row=header_row + 1):
            row_idx = row[0].row
            title = row[title_col - 1].value
            if title is None or not str(title).strip():
                continue
            raw_label = row[label_col - 1].value if label_col <= len(row) else None
            raw_note = row[note_col - 1].value if note_col <= len(row) else None
            raw_id = row[0].value if title_col != 1 else None
            records.append(
                {
                    "excel_row": row_idx,
                    "id": raw_id,
                    "titolo": str(title),
                    "label": normalize_label(raw_label),
                    "note": "" if raw_note is None else str(raw_note),
                }
            )
        if not records:
            raise ValueError("La colonna 'titolo' è presente, ma non contiene titoli da annotare.")
        return sheet.title, header_row, cols, records
    finally:
        workbook.close()


def export_excel() -> bytes:
    """Scrive soltanto le righe modificate; preserva le altre celle e i fogli."""
    changed = st.session_state["changed_rows"]
    source_bytes = st.session_state["source_bytes"]
    if not changed:
        return source_bytes

    workbook = load_workbook(BytesIO(source_bytes))
    worksheet = workbook[st.session_state["sheet_name"]]
    columns = st.session_state["columns"]
    header_row = st.session_state["header_row"]

    if worksheet.cell(header_row, columns["label"]).value is None:
        worksheet.cell(header_row, columns["label"], "label")
    if worksheet.cell(header_row, columns["note"]).value is None:
        worksheet.cell(header_row, columns["note"], "note")

    for item in st.session_state["records"]:
        if item["excel_row"] not in changed:
            continue
        excel_row = item["excel_row"]
        worksheet.cell(excel_row, columns["label"]).value = item["label"]
        note_cell = worksheet.cell(excel_row, columns["note"])
        note_cell.value = item["note"] or None
        # Una nota che inizia con '=' non deve diventare una formula Excel.
        if isinstance(note_cell.value, str) and note_cell.value.startswith(("=", "+", "-", "@")):
            note_cell.data_type = "s"

    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def save_record(index: int, label_text: str, note_text: str) -> None:
    record = st.session_state["records"][index]
    record["label"] = None if label_text == "Non assegnata" else int(label_text)
    record["note"] = note_text.strip()
    st.session_state["changed_rows"].add(record["excel_row"])


def first_to_annotate(records: list[dict]) -> int:
    return next((i for i, record in enumerate(records) if record["label"] is None), 0)


st.title("🏷️ Annotazione titoli")
st.caption("Carica un file Excel, assegna una label 1 o 0 a ogni titolo e aggiungi una nota facoltativa.")

uploaded = st.file_uploader(
    "📂 Carica il file Excel da annotare",
    type=["xlsx"],
    help="Seleziona un file .xlsx con una colonna 'titolo'. Puoi ricaricare un file già annotato per riprendere il lavoro.",
)
if uploaded is None:
    st.info("Carica un file Excel (.xlsx) per iniziare. Deve contenere una colonna chiamata 'titolo'.")
    st.stop()

raw_data = uploaded.getvalue()
active_filename = Path(uploaded.name).name
if not raw_data:
    st.error("Il file caricato è vuoto. Seleziona un Excel valido.")
    st.stop()

fingerprint = hashlib.sha256(raw_data).hexdigest()
if st.session_state.get("source_hash") != fingerprint:
    try:
        sheet_name, header_row, columns, records = read_records(raw_data)
    except Exception as exc:
        st.error(f"Impossibile leggere l'Excel: {exc}")
        st.stop()
    st.session_state.update(
        {
            "source_hash": fingerprint,
            "source_bytes": raw_data,
            "source_filename": active_filename,
            "sheet_name": sheet_name,
            "header_row": header_row,
            "columns": columns,
            "records": records,
            "position": first_to_annotate(records),
            "changed_rows": set(),
        }
    )

records = st.session_state["records"]
current = st.session_state["position"]
total = len(records)
ones = sum(item["label"] == 1 for item in records)
zeros = sum(item["label"] == 0 for item in records)
labeled = ones + zeros
missing = total - labeled

with st.sidebar:
    st.header("File attivo")
    st.caption(st.session_state["source_filename"])
    st.divider()
    st.subheader("Avanzamento")
    st.progress(labeled / total, text=f"{labeled} su {total} titoli etichettati ({labeled / total:.0%})")
    col1, col2 = st.columns(2)
    col1.metric("Label 1", ones)
    col2.metric("Label 0", zeros)
    st.metric("Da etichettare", missing)
    st.caption(f"Foglio: {st.session_state['sheet_name']}")
    if st.button("Vai al primo senza label", use_container_width=True, disabled=missing == 0):
        st.session_state["position"] = first_to_annotate(records)
        st.rerun()

    with st.form("go_to_row", border=False):
        destination = st.number_input(
            "Vai al titolo numero", min_value=1, max_value=total, value=current + 1, step=1
        )
        go = st.form_submit_button("Vai", use_container_width=True)
    if go:
        st.session_state["position"] = int(destination) - 1
        st.rerun()
    st.caption("Prima di spostarti con questi comandi, salva eventuali modifiche al titolo corrente.")

    st.divider()
    st.caption("Le modifiche rimangono nella sessione corrente. Scarica periodicamente l'Excel per conservarle.")

st.progress((current + 1) / total, text=f"Titolo {current + 1} di {total}")
item = records[current]
identifier = "" if item["id"] is None else f" · ID {html.escape(str(item['id']))}"
st.markdown(
    f"""
    <div class="title-panel">
      <div class="title-eyebrow">Titolo {current + 1}/{total}{identifier}</div>
      <div class="title-content">{html.escape(item['titolo'])}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

selected = "Non assegnata" if item["label"] is None else str(item["label"])
choices = ["Non assegnata", "1", "0"]

with st.form(key=f"edit_record_{item['excel_row']}", clear_on_submit=False):
    label = st.radio(
        "Label del titolo",
        options=choices,
        index=choices.index(selected),
        horizontal=True,
        help="1 e 0 sono salvati come valori numerici nell'Excel. 'Non assegnata' lascia la cella vuota.",
    )
    note = st.text_area(
        "Nota (facoltativa)",
        value=item["note"],
        placeholder="Aggiungi un'osservazione, una motivazione o un dubbio...",
        height=120,
    )
    back_col, save_col, next_col = st.columns([1, 1, 1.7])
    with back_col:
        previous = st.form_submit_button("← Precedente", use_container_width=True, disabled=current == 0)
    with save_col:
        save = st.form_submit_button("Salva", use_container_width=True)
    with next_col:
        next_one = st.form_submit_button("Salva e successivo →", type="primary", use_container_width=True)

if previous or save or next_one:
    save_record(current, label, note)
    if previous:
        st.session_state["position"] = max(0, current - 1)
    elif next_one:
        st.session_state["position"] = min(total - 1, current + 1)
    st.rerun()

st.divider()
st.subheader("Esporta annotazioni")
st.write("Scarica il file Excel con le label e le note aggiornate. Tutte le colonne originali sono mantenute.")
try:
    export_bytes = export_excel()
    export_name = f"{Path(st.session_state['source_filename']).stem}_annotato.xlsx"
    st.download_button(
        "⬇️ Scarica Excel annotato",
        data=export_bytes,
        file_name=export_name,
        mime=EXCEL_MIME,
        type="primary",
        use_container_width=True,
    )
except Exception as exc:
    st.error(f"Errore durante la preparazione dell'Excel: {exc}")

st.caption("L'esportazione include le modifiche salvate con i pulsanti sopra. Il download non modifica il file originale.")
