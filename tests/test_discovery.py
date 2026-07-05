"""Landing-page discovery: label parsing, current-quarter selection, and the
fail-loud guards (malformed link, off-domain link, FTA/NFTA disagreement)."""
from datetime import date

import pytest

import canada_trq_tracker as t


# --- _parse_quarter_label ---------------------------------------------------

def test_label_en_dash_and_inferred_start_year():
    q, s, e, ss, es = t._parse_quarter_label("Quarter 1: June 28 – September 29, 2026")
    assert (q, s, e) == ("Q1", date(2026, 6, 28), date(2026, 9, 29))
    assert (ss, es) == ("June 28, 2026", "September 29, 2026")


def test_label_year_wrap_infers_previous_year():
    q, s, e, _, _ = t._parse_quarter_label("Quarter 3: December 26 – March 25, 2026")
    assert (q, s, e) == ("Q3", date(2025, 12, 26), date(2026, 3, 25))


def test_label_explicit_both_years():
    q, s, e, _, _ = t._parse_quarter_label("Quarter 3: December 26, 2025 – March 25, 2026")
    assert (q, s, e) == ("Q3", date(2025, 12, 26), date(2026, 3, 25))


@pytest.mark.parametrize("text", [
    "Period 2: July 1 – July 31, 2026",     # sub-period rows must not match
    "Quarter 1: whenever",
    "not a label at all",
])
def test_label_rejects_non_quarter_rows(text):
    assert t._parse_quarter_label(text) is None


# --- discover_current_reports ------------------------------------------------

BASE = "https://www.international.gc.ca/trade-commerce/controls-controles/steel-acier/"


def landing_html(rows):
    anchors = []
    for stem, label, csv_href in rows:
        anchors.append(f'<a href="{stem}.htm">{label}</a>')
        anchors.append(f'<a href="{csv_href}">[.CSV]</a>')
    return f"<html><body>{''.join(anchors)}</body></html>"


class FakeResponse:
    def __init__(self, text):
        self.text = text
        self.encoding = None

    def raise_for_status(self):
        pass


def patch_landing(monkeypatch, html):
    class FakeSession:
        def get(self, url, timeout=None):
            return FakeResponse(html)
    monkeypatch.setattr(t, "SESSION", FakeSession())


GOOD_ROWS = [
    ("TRQ_FTA-Q4", "Quarter 4: March 26 – June 27, 2026",
     "https://www.eics-scei.gc.ca/report-rapport/TRQ_FTA-Q4.csv"),
    ("TRQ_NFTA-Q4", "Quarter 4: March 26 – June 27, 2026",
     "https://www.eics-scei.gc.ca/report-rapport/TRQ_NFTA-q4.csv"),
    ("TRQ_FTA-Y2Q1", "Quarter 1: June 28 – September 29, 2026",
     "https://www.eics-scei.gc.ca/report-rapport/TRQ_FTA-Y2Q1.csv"),
    ("TRQ_NFTA-Y2Q1", "Quarter 1: June 28 – September 29, 2026",
     "https://www.eics-scei.gc.ca/report-rapport/TRQ_nFTA-Y2Q1.csv"),
]


def test_discovery_picks_quarter_containing_today(monkeypatch):
    patch_landing(monkeypatch, landing_html(GOOD_ROWS))
    reports = t.discover_current_reports(date(2026, 7, 6))
    assert reports["FTA"]["quarter"] == "Q1"
    assert reports["FTA"]["csv_url"].endswith("TRQ_FTA-Y2Q1.csv")
    assert reports["NFTA"]["csv_url"].endswith("TRQ_nFTA-Y2Q1.csv")
    assert reports["FTA"]["start_str"] == "June 28, 2026"


def test_discovery_gap_day_takes_most_recently_started(monkeypatch):
    """2026-06-27: Q4 ended today, Y2Q1 starts tomorrow -> stay on Q4."""
    patch_landing(monkeypatch, landing_html(GOOD_ROWS))
    assert t.discover_current_reports(date(2026, 6, 27))["FTA"]["quarter"] == "Q4"


def test_discovery_resolves_relative_hrefs(monkeypatch):
    rows = [(s, l, "TRQ_FTA-Y2Q1.csv" if s == "TRQ_FTA-Y2Q1" else u)
            for s, l, u in GOOD_ROWS]
    patch_landing(monkeypatch, landing_html(rows))
    url = t.discover_current_reports(date(2026, 7, 6))["FTA"]["csv_url"]
    assert url == BASE + "TRQ_FTA-Y2Q1.csv"   # absolute, resolved against the page


def test_discovery_rejects_malformed_csv_link(monkeypatch):
    rows = [(s, l, u.replace("TRQ_FTA-Y2Q1.csv", "TRQ_FTA-Y2Q1.htm"))
            for s, l, u in GOOD_ROWS]
    patch_landing(monkeypatch, landing_html(rows))
    with pytest.raises(RuntimeError, match="not a .csv"):
        t.discover_current_reports(date(2026, 7, 6))


def test_discovery_rejects_off_domain_link(monkeypatch):
    rows = [(s, l, u.replace("www.eics-scei.gc.ca", "evil.example.com"))
            for s, l, u in GOOD_ROWS]
    patch_landing(monkeypatch, landing_html(rows))
    with pytest.raises(RuntimeError, match="gc.ca"):
        t.discover_current_reports(date(2026, 7, 6))


def test_discovery_rejects_quarter_disagreement(monkeypatch):
    rows = [r for r in GOOD_ROWS if "NFTA-Y2Q1" not in r[0]]  # NFTA stuck on Q4
    patch_landing(monkeypatch, landing_html(rows))
    with pytest.raises(RuntimeError, match="different quarters"):
        t.discover_current_reports(date(2026, 7, 6))


def test_discovery_fails_without_reports(monkeypatch):
    patch_landing(monkeypatch, "<html><body>maintenance</body></html>")
    with pytest.raises(RuntimeError, match="no FTA"):
        t.discover_current_reports(date(2026, 7, 6))


# --- should_fetch_b1 ----------------------------------------------------------

@pytest.mark.parametrize("quarter,month,expected", [
    ("Q1", 7, True), ("Q1", 9, True), ("Q1", 6, False),
    ("Q3", 1, True), ("Q3", 12, False), ("Q4", 4, True), ("Q4", 3, False),
])
def test_should_fetch_b1(quarter, month, expected):
    assert t.should_fetch_b1(quarter, date(2026, month, 15)) is expected
