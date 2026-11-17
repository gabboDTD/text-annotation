"""App Streamlit per annotare titoli caricati da Excel (lettura con pandas)."""
from __future__ import annotations

from pathlib import Path
import hashlib
import html

import streamlit as st

from excel_io import export_annotations, read_records


EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

st.set_page_config(
    page_title="Annotatore di titoli", page_icon="🏷️", layout="centered",
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
    </style>
    """,
    unsafe_allow_html=True,
)


def first_unlabeled(records: list[dict]) -> int:
    return next((i for i, r in enumerate(records) if r["label"] is None), 0)


def save_record(index: int, label: str, note: str) -> None:
    item = st.session_state["records"][index]
    item["label"] = None if label == "Non assegnata" else int(label)
    item["note"] = note.strip()
    st.session_state["changed_rows"].add(item["excel_row"])


st.title("🏷️ Annotazione titoli")
st.caption("Carica un Excel, classifica un titolo alla volta con 1 o 0 e aggiungi note facoltative.")

uploaded = st.file_uploader(
    "📂 Carica il file Excel da annotare", type=["xlsx"],
    help="È sufficiente che un foglio contenga la colonna 'titolo'. Puoi ricaricare un Excel già annotato.",
)
if uploaded is None:
    st.info("Carica un file .xlsx con una colonna 'titolo' per iniziare.")
    st.stop()

raw_data = uploaded.getvalue()
if not raw_data:
    st.error("Il file caricato è vuoto.")
    st.stop()

fingerprint = hashlib.sha256(raw_data).hexdigest()
if st.session_state.get("source_hash") != fingerprint:
    try:
        sheet_name, header_index, columns, records = read_records(raw_data)
    except Exception as exc:
        st.error(f"Impossibile leggere l'Excel: {exc}")
        st.stop()
    st.session_state.update({
        "source_hash": fingerprint,
        "source_bytes": raw_data,
        "source_filename": Path(uploaded.name).name,
        "sheet_name": sheet_name,
        "header_index": header_index,
        "columns": columns,
        "records": records,
        "position": first_unlabeled(records),
        "changed_rows": set(),
    })

records = st.session_state["records"]
current = st.session_state["position"]
total = len(records)
ones = sum(r["label"] == 1 for r in records)
zeros = sum(r["label"] == 0 for r in records)
labeled = ones + zeros
missing = total - labeled

with st.sidebar:
    st.header("File attivo")
    st.caption(st.session_state["source_filename"])
    st.divider()
    st.subheader("Avanzamento")
    st.progress(labeled / total, text=f"{labeled} su {total} classificati ({labeled / total:.0%})")
    col1, col2 = st.columns(2)
    col1.metric("Label 1", ones)
    col2.metric("Label 0", zeros)
    st.metric("Da etichettare", missing)
    st.caption(f"Foglio: {st.session_state['sheet_name']}")
    if st.button("Vai al primo senza label", use_container_width=True, disabled=missing == 0):
        st.session_state["position"] = first_unlabeled(records)
        st.rerun()

    with st.form("go_to_row", border=False):
        destination = st.number_input("Vai al titolo numero", min_value=1, max_value=total, value=current + 1)
        go = st.form_submit_button("Vai", use_container_width=True)
    if go:
        st.session_state["position"] = int(destination) - 1
        st.rerun()
    st.caption("Salva il titolo corrente prima di spostarti dalla barra laterale.")
    st.divider()
    st.caption("Le modifiche restano nella sessione. Scarica l'Excel per conservarle.")

st.progress((current + 1) / total, text=f"Titolo {current + 1} di {total}")
item = records[current]
identifier = f" · ID {html.escape(str(item['id']))}" if item["id"] else ""
st.markdown(
    f'<div class="title-panel"><div class="title-eyebrow">Titolo {current + 1}/{total}'
    f'{identifier}</div><div class="title-content">{html.escape(item["titolo"])}</div></div>',
    unsafe_allow_html=True,
)

selected = "Non assegnata" if item["label"] is None else str(item["label"])
choices = ["Non assegnata", "1", "0"]
with st.form(key=f"edit_record_{item['excel_row']}", clear_on_submit=False):
    label = st.radio("Label", choices, index=choices.index(selected), horizontal=True)
    note = st.text_area("Nota (facoltativa)", value=item["note"], height=120,
                        placeholder="Motivazione, osservazione, dubbio...")
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
st.write("Scarica un Excel con label e note salvate, conservando gli altri dati e fogli.")
try:
    export_bytes = export_annotations(
        st.session_state["source_bytes"], st.session_state["sheet_name"],
        st.session_state["header_index"], st.session_state["columns"],
        records, st.session_state["changed_rows"],
    )
    export_name = f"{Path(st.session_state['source_filename']).stem}_annotato.xlsx"
    st.download_button(
        "⬇️ Scarica Excel annotato", data=export_bytes, file_name=export_name,
        mime=EXCEL_MIME, type="primary", use_container_width=True,
    )
except Exception as exc:
    st.error(f"Errore durante l'esportazione: {exc}")

st.caption("Il download include le annotazioni confermate con Salva o Salva e successivo.")
