"""Saved-workbook invariants for TRQ sheets.

This is the executable version of the manual checklist in
docs/known-issues.md (Bug 1, "Prevention"): contiguous product blocks, no
duplicate countries, no formulas in country data cells, and every TOTAL SUM()
spanning exactly its product's country rows — in EVERY date column, not just
the newest one (the newest column is always written last from fresh ranges,
so checking only it hides insert-shift corruption).
"""
import re

SUM_RE = re.compile(r"^=SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)$")


def product_blocks(ws):
    """Return {product: {"countries": [(row, name)], "total_row": int|None}}
    from the sheet's current physical layout, asserting block contiguity."""
    blocks: dict[str, dict] = {}
    order: list[str] = []
    for row in range(3, ws.max_row + 1):
        prod = ws.cell(row=row, column=1).value
        country = ws.cell(row=row, column=2).value
        if not prod or not country:
            continue
        prod, country = str(prod).strip(), str(country).strip()
        if prod not in blocks:
            blocks[prod] = {"countries": [], "total_row": None}
            order.append(prod)
        assert prod == order[-1], (
            f"non-contiguous block: {prod!r} reappears at row {row} after {order[-1]!r}")
        if country == "TOTAL":
            assert blocks[prod]["total_row"] is None, f"{prod}: duplicate TOTAL row {row}"
            blocks[prod]["total_row"] = row
        else:
            assert blocks[prod]["total_row"] is None, (
                f"{prod}: country {country!r} at row {row} AFTER its TOTAL row")
            blocks[prod]["countries"].append((row, country))
    return blocks


def assert_trq_sheet_invariants(ws, first_date_col, last_date_col):
    blocks = product_blocks(ws)
    assert blocks, f"sheet {ws.title!r}: no product blocks found"
    for prod, b in blocks.items():
        names = [c for _, c in b["countries"]]
        assert len(names) == len(set(names)), f"{prod}: duplicate countries {names}"
        total_row = b["total_row"]
        assert total_row is not None, f"{prod}: no TOTAL row"

        for col in range(first_date_col, last_date_col + 1):
            # Country data cells must hold numbers or be blank — never formulas.
            for row, country in b["countries"]:
                v = ws.cell(row=row, column=col).value
                assert v is None or isinstance(v, (int, float)), (
                    f"{prod}/{country} r{row}c{col}: data cell holds {v!r}")

            v = ws.cell(row=total_row, column=col).value
            if not b["countries"]:
                assert v == 0.0, (
                    f"{prod} r{total_row}c{col}: zero-country TOTAL must be "
                    f"literal 0.0, got {v!r}")
                continue
            assert isinstance(v, str), (
                f"{prod} r{total_row}c{col}: TOTAL must be a SUM formula, got {v!r}")
            m = SUM_RE.match(v)
            assert m, f"{prod} r{total_row}c{col}: unexpected TOTAL formula {v!r}"
            c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
            assert c1 == c2, f"{prod}: TOTAL sums across columns: {v!r}"
            first = b["countries"][0][0]
            last = b["countries"][-1][0]
            assert (r1, r2) == (first, last), (
                f"{prod} r{total_row}c{col}: TOTAL range {v!r} != country rows "
                f"{first}..{last} — stale/shifted formula")
            assert r2 < total_row, f"{prod}: TOTAL formula includes itself: {v!r}"
