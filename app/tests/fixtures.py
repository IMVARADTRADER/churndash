"""Build test files in memory from the original Problem Statement 2.xlsx: a reworded-notes file and broken files.

Nothing here is saved unless make_demo_files.py is run (it writes to a folder you choose, kept out of git).
"""
import io
import json
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[2]
ORIGINAL = ROOT / "Problem Statement 2.xlsx"
TEST_SET = ROOT / "scripts" / "note_test_set.json"


def original_bytes() -> bytes:
    return ORIGINAL.read_bytes()


def _edit(fn) -> bytes:
    wb = openpyxl.load_workbook(io.BytesIO(original_bytes()))
    fn(wb)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def rewordings() -> dict:
    """Original note -> (first reworded version, the user's reviewed tag), from the 94-note tagging test set."""
    notes = json.load(open(TEST_SET, encoding="utf-8"))["notes"]
    originals = {n["source"]: n for n in notes if n["group"] == "original"}
    out = {}
    for n in notes:
        if n["group"] == "reworded" and originals[n["source"]]["text"] not in out:
            out[originals[n["source"]]["text"]] = (n["text"], n["label"])
    return out


def reworded_bytes() -> tuple[bytes, dict]:
    """Every account manager note replaced by a rewording. Returns the file and {new note: correct tag}."""
    rw = rewordings()

    def fn(wb):
        ws = wb["P2_Merchants"]
        head = [c.value for c in ws[1]]
        col = head.index("account_manager_note") + 1
        for r in range(2, ws.max_row + 1):
            v = ws.cell(r, col).value
            if v in rw:
                ws.cell(r, col).value = rw[v][0]
    return _edit(fn), {new: tag for new, tag in rw.values()}


PLANTED_ID = "C2004"  # on the CSM list: Pricing note, renews in 39 days
PLANTED_NOTE = "Wants us to offer them a discount before they renew"


def planted_bytes() -> bytes:
    """Test 3: one CSM-list merchant's note replaced with a request the rules forbid promising."""
    def fn(wb):
        ws = wb["P2_Merchants"]
        head = [c.value for c in ws[1]]
        col = head.index("account_manager_note") + 1
        for r in range(2, ws.max_row + 1):
            if ws.cell(r, 1).value == PLANTED_ID:
                ws.cell(r, col).value = PLANTED_NOTE
    return _edit(fn)


def _set(sheet_fn):
    return _edit(sheet_fn)


def broken_files() -> dict:
    """name -> (bytes, a phrase the error message must contain)."""
    def rename_sheet(wb):
        wb["P2_Merchants"].title = "Merchants"

    def drop_column(wb):
        ws = wb["P2_Merchants"]
        head = [c.value for c in ws[1]]
        ws.delete_cols(head.index("unresolved_tickets") + 1)

    def text_in_number(wb):
        ws = wb["P2_Merchants"]
        head = [c.value for c in ws[1]]
        ws.cell(5, head.index("tpv_m6_inr_lakh") + 1).value = "twelve"

    def blank_required(wb):
        ws = wb["P2_Merchants"]
        head = [c.value for c in ws[1]]
        ws.cell(8, head.index("days_to_contract_renewal") + 1).value = None

    def dup_id(wb):
        ws = wb["P2_Merchants"]
        ws.cell(3, 1).value = ws.cell(2, 1).value

    def bad_csat(wb):
        ws = wb["P2_Merchants"]
        head = [c.value for c in ws[1]]
        ws.cell(2, head.index("last_csat_1_to_5") + 1).value = 9

    def headers_only(wb):
        ws = wb["P2_Merchants"]
        ws.delete_rows(2, ws.max_row)

    return {
        "not_excel.xlsx": (b"merchant_id,business_name\nC1,Test\n", "isn't an Excel .xlsx file"),
        "empty.xlsx": (b"", "empty"),
        "wrong_sheet.xlsx": (_set(rename_sheet), "no sheet called 'P2_Merchants'"),
        "missing_column.xlsx": (_set(drop_column), "Missing column: unresolved_tickets"),
        "text_in_number.xlsx": (_set(text_in_number), "'tpv_m6_inr_lakh' must be a number; check Excel row 5"),
        "blank_required.xlsx": (_set(blank_required), "'days_to_contract_renewal' is blank in Excel row 8"),
        "duplicate_id.xlsx": (_set(dup_id), "Merchant IDs repeat: C2001"),
        "csat_out_of_range.xlsx": (_set(bad_csat), "must be 1 to 5 or blank; check Excel row 2"),
        "headers_only.xlsx": (_set(headers_only), "headers but no merchants"),
    }
