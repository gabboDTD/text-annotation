"""Lettura Excel con pandas e salvataggio mirato delle annotazioni.

Pandas gestisce la lettura dei DataFrame, evitando accessi diretti a EmptyCell.
ExcelWriter usa il motore openpyxl soltanto per aggiornare alcune celle senza
ricreare l'intero workbook (e dunque senza perdere fogli e formattazione).
"""
from __future__ import annotations

from io import BytesIO
from typing import Any

import pandas as pd


def normalize_label(value: Any) -> int | None:
    if isinstance(value, bool):
        return None
    if value in (0, 1, "0", "1", 0.0, 1.0):
        return int(value)
    if isinstance(value, str) and value.strip() in ("0", "1"):
        return int(value.strip())
    return None


def _text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if bool(pd.isna(value)):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value)


def read_records(raw_bytes: bytes) -> tuple[str, int, dict[str, int | None], list[dict]]:
    """Restituisce (nome foglio, indice header 0-based, colonne 1-based, record)."""
    with pd.ExcelFile(BytesIO(raw_bytes), engine="openpyxl") as excel:
        for sheet_name in excel.sheet_names:
            preview = pd.read_excel(
                excel, sheet_name=sheet_name, header=None, nrows=15,
                dtype=object, keep_default_na=False,
            )
            for header_idx, cells in preview.iterrows():
                headers = [_text(v).strip().casefold() for v in cells.tolist()]
                if "titolo" not in headers:
                    continue

                # I dati sono letti in DataFrame: le righe completamente vuote
                # intermedie mantengono la loro posizione rispetto all'Excel.
                frame = pd.read_excel(
                    excel, sheet_name=sheet_name, header=int(header_idx),
                    dtype=object, keep_default_na=False,
                )
                names = [_text(col).strip().casefold() for col in frame.columns]
                title_index = names.index("titolo")
                label_index = names.index("label") if "label" in names else None
                note_index = names.index("note") if "note" in names else None
                columns = {
                    "titolo": title_index + 1,
                    "label": (label_index + 1) if label_index is not None else None,
                    "note": (note_index + 1) if note_index is not None else None,
                }

                records: list[dict] = []
                for pos in range(len(frame)):
                    title = _text(frame.iat[pos, title_index]).strip()
                    if not title:
                        continue
                    label_raw = frame.iat[pos, label_index] if label_index is not None else ""
                    note_raw = frame.iat[pos, note_index] if note_index is not None else ""
                    record = {
                        "excel_row": int(header_idx) + 2 + pos,
                        "id": (_text(frame.iat[pos, 0]) if title_index != 0 else ""),
                        "titolo": title,
                        "label": normalize_label(label_raw),
                        "note": _text(note_raw),
                    }
                    records.append(record)

                if not records:
                    raise ValueError("La colonna 'titolo' è presente ma non contiene titoli da annotare.")
                return sheet_name, int(header_idx), columns, records

    raise ValueError(
        "Non trovo una colonna chiamata 'titolo' nelle prime 15 righe di alcun foglio."
    )


def export_annotations(
    source_bytes: bytes,
    sheet_name: str,
    header_index: int,
    columns: dict[str, int | None],
    records: list[dict],
    changed_rows: set[int],
) -> bytes:
    """Conserva il file originale, scrivendo solo label/note delle righe salvate."""
    output = BytesIO(source_bytes)
    with pd.ExcelWriter(output, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
        sheet = writer.book[sheet_name]
        # Se mancanti, aggiunge label e note senza sovrascrivere colonne esistenti.
        target_columns = dict(columns)
        for name in ("label", "note"):
            if target_columns[name] is None:
                col = sheet.max_column + 1
                sheet.cell(row=header_index + 1, column=col, value=name)
                target_columns[name] = col

        for item in records:
            if item["excel_row"] not in changed_rows:
                continue
            row_num = item["excel_row"]
            sheet.cell(row=row_num, column=target_columns["label"]).value = item["label"]
            note_cell = sheet.cell(row=row_num, column=target_columns["note"])
            note_cell.value = item["note"] or None
            # Evita che una nota iniziata da '=' diventi una formula Excel.
            if isinstance(note_cell.value, str) and note_cell.value.startswith("="):
                note_cell.data_type = "s"

    return output.getvalue()
