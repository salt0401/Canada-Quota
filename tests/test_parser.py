"""parse_trq_csv: both CSV format variants, the Part A/Part B positional
contract, and the fail-loud guard on structural changes."""
import json
import os

import pytest

import canada_trq_tracker as t
from conftest import FIXTURES

# 23 products; the 8 tracked ones sit at their real item numbers.
PRODUCTS = [
    (1, "Energy Tubular"), (2, "Line Pipe"), (3, "Hot-Rolled Sheet"),
    (4, "Steel Plate"), (5, "Cold-Rolled Sheet"), (6, "Coated Steel Sheet"),
    (7, "Rebar"), (8, "Hot-Rolled Bar"), (9, "Wire Rod"),
    (10, "Structural Steel"), (11, "Other A"), (12, "Other B"),
    (13, "Other C"), (14, "Other D"), (15, "Other E"), (16, "Other F"),
    (17, "Other G"), (18, "Other H"), (19, "Other I"), (20, "Other J"),
    (21, "Other K"), (22, "Other L"), (23, "Other M"),
]

# product -> (max_quota, max_share%, [(country, kgm, share%)]) ; [] = zero util
DATA = {
    "Hot-Rolled Sheet": (8_569_800, "71.00%", [("Japan", "2,712,750", "31.65%"),
                                               ("Korea, Republic of", "5,857,050", "68.35%")]),
    "Steel Plate": (937_768, "33.00%", [("Germany", "297,418", "10.68%")]),
    "Wire Rod": (1_000_000, "25.00%", []),          # zero utilization
    "Rebar": (2_000_000, "41.00%", [("Max", "0", "Max Utilized")]),  # capped country
}


def _rows_for(name):
    quota, share, rows = DATA.get(name, (500_000, "10.00%", [("France", "1,000", "0.20%")]))
    total_kgm = sum(int(k.replace(",", "")) for _, k, _ in rows)
    util_pct = "Max Utilized" if any(p == "Max Utilized" for _, _, p in rows) else (
        "0.00%" if not rows else "50.00%")
    return quota, share, rows, total_kgm, util_pct


def build_new_format(rename_products=False, mismatch_products=frozenset()) -> str:
    """rename_products: Part A product labels change (simulates a source-side
    rename — no tracked product should be found).
    mismatch_products: these products' Part B totals disagree with Part A."""
    lines = ["Textbox137,Textbox143,Textbox155", "meta,meta,meta", "",
             "Textboxb6db,Textbox0874,Textbox38"]
    hdr6 = ("Product class,Maximum quota (KGM),Maximum country share (%),"
            "Current utilization (KGM),Current utilization (%),Remaining quota (KGM)")
    for item, name in PRODUCTS:
        quota, share, rows, total_kgm, util_pct = _rows_for(name)
        label = f"Renamed {name}" if rename_products else name
        lines.append(f'{hdr6},{item},{label},"{quota:,}",{share},"{total_kgm:,}",{util_pct},"0"')
    lines.append("")
    for item, name in PRODUCTS:
        _, _, rows, total_kgm, _ = _rows_for(name)
        if name in mismatch_products:
            total_kgm += 5000
        lines.append(f"Textboxb6db{item},Textbox6defe{item},Textbox81837{item}")
        for country, kgm, pct in rows:
            c = f'"{country}"' if "," in country else country
            lines.append(f'Country,Current utilization (KGM),Share of total utilization (%),'
                         f'{c},"{kgm}",{pct},"{total_kgm:,}",50.00%')
        lines.append("")
    return "\n".join(lines)


def build_old_format() -> str:
    lines = ["ExecutionTime,2026-07-05", "meta", "", "hdr"]
    for item, name in PRODUCTS:
        quota, share, rows, total_kgm, util_pct = _rows_for(name)
        lines.append(f'{item},{name},"{quota:,}",{share},"{total_kgm:,}",{util_pct},"0"')
    lines.append("")
    for item, name in PRODUCTS:
        _, _, rows, total_kgm, _ = _rows_for(name)
        lines.append(f"CONTROL_ITEMS_Level_1_{item},x,y")
        for country, kgm, pct in rows:
            c = f'"{country}"' if "," in country else country
            lines.append(f'{c},"{kgm}",{pct},"{total_kgm:,}",50.00%')
        lines.append("")
    return "\n".join(lines)


@pytest.mark.parametrize("builder", [build_new_format, build_old_format],
                         ids=["new-Textbox", "old-ExecutionTime"])
def test_parse_both_formats(builder):
    result = t.parse_trq_csv(builder())
    assert set(result) == set(t.TRACKED_PRODUCTS)

    hrs = result["Hot-Rolled Sheet"]
    assert hrs["item_number"] == 3
    assert hrs["max_quota"] == 8_569_800
    assert hrs["max_share"] == pytest.approx(0.71)
    assert hrs["countries"]["Japan"] == pytest.approx(0.3165)
    assert hrs["countries"]["Korea, Republic of"] == pytest.approx(0.6835)

    # Zero-utilization product: present, empty countries.
    assert result["Wire Rod"]["countries"] == {}
    assert result["Wire Rod"]["total_util_pct"] == 0.0

    # "Max Utilized" maps to 100% (product) / 100% is not forced on countries.
    assert result["Rebar"]["total_util_pct"] == 1.0
    assert result["Rebar"]["countries"]["Max"] == pytest.approx(1.0)


def test_positional_alignment_not_shifted():
    """Countries land on THEIR OWN product, not a neighbour's (the silent
    desync failure class)."""
    result = t.parse_trq_csv(build_new_format())
    assert "Germany" in result["Steel Plate"]["countries"]
    assert "Germany" not in result["Cold-Rolled Sheet"]["countries"]
    assert "Japan" not in result["Steel Plate"]["countries"]


@pytest.mark.parametrize("n_items", [22, 24])
def test_wrong_part_a_count_fails_loud(n_items):
    global PRODUCTS
    orig = PRODUCTS
    try:
        PRODUCTS = orig[:n_items] if n_items < len(orig) else orig + [(24, "Extra")]
        csv_text = build_new_format()
    finally:
        PRODUCTS = orig
    with pytest.raises(ValueError, match="expected 23"):
        t.parse_trq_csv(csv_text)


def test_malformed_part_a_line_fails_loud():
    """A short Part A row is dropped by the field guard, so the count assert
    aborts the run instead of letting Part B desynchronize."""
    lines = build_new_format().split("\n")
    lines[8] = "garbage,short"   # a Part A row (index 4..26)
    with pytest.raises(ValueError, match="expected 23"):
        t.parse_trq_csv("\n".join(lines))


def test_empty_and_unknown_csv():
    with pytest.raises(ValueError):
        t.parse_trq_csv("")
    with pytest.raises(ValueError, match="Unrecognized"):
        t.parse_trq_csv("hello,world\n1,2")


def test_all_tracked_products_missing_fails_loud():
    """A wholesale label rename must abort, not publish an all-blank week."""
    with pytest.raises(ValueError, match="NONE of the tracked products"):
        t.parse_trq_csv(build_new_format(rename_products=True))


def test_widespread_kgm_mismatch_fails_loud():
    """>= _PART_AB_MISMATCH_LIMIT Part A/B total disagreements = positional
    desync -> abort; a couple of mismatches stay warn-only (deliberate)."""
    many = {"Hot-Rolled Sheet", "Steel Plate", "Other A", "Other B", "Other C"}
    with pytest.raises(ValueError, match="desynchronized"):
        t.parse_trq_csv(build_new_format(mismatch_products=many))
    # one mismatch: warn-only, parse succeeds
    result = t.parse_trq_csv(build_new_format(mismatch_products={"Steel Plate"}))
    assert "Steel Plate" in result


def test_real_fixture_matches_golden():
    """Live-captured government CSVs (2026-07-05) parse to the exact output
    the pre-refactor parser produced (regression net for parser changes)."""
    with open(os.path.join(FIXTURES, "TRQ_sample_golden.json"), encoding="utf-8") as f:
        golden = json.load(f)
    for partner, fixture in [("FTA", "TRQ_FTA_sample.csv"), ("NFTA", "TRQ_NFTA_sample.csv")]:
        with open(os.path.join(FIXTURES, fixture), encoding="utf-8") as f:
            parsed = json.loads(json.dumps(t.parse_trq_csv(f.read())))
        assert parsed == golden[partner], f"{partner} fixture parse drifted from golden"
