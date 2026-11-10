import csv
import unittest
from datetime import date, datetime
from decimal import Decimal
from io import BytesIO, StringIO
from pathlib import Path
from unittest.mock import patch
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl import Workbook
from streamlit.testing.v1 import AppTest

from tiliraportti_analyysi.analysis import analyze
from tiliraportti_analyysi.app import csv_bytes

ROOT = Path(__file__).resolve().parents[1]
HEADERS = ["Organisaatio", "Tilityksen viitenumero", "Maksun tunniste",
           "Joukkueet", "Maksuosuuden tyyppi", "Maksuosuus"]


def workbook_bytes(rows, headers=HEADERS):
    book = Workbook()
    sheet = book.active
    sheet.title = "Tilitysrivit"
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    output = BytesIO()
    book.save(output)
    book.close()
    return output.getvalue()


def example_bytes():
    """Generate synthetic data without requiring private settlement reports."""
    rows = []
    for team, gross, service, commission in [
        ("Joukkue A", 1260, "-11.64", "-5.88"),
        ("Joukkue B", 80, "-0.65", "-0.42"),
        ("Joukkue C", 320, "-4.20", "-1.46"),
    ]:
        rows.extend([
            ["Testiorganisaatio", "1", team, team, "Maksu", gross],
            ["Testiorganisaatio", "1", team, None, "Palvelumaksu", service],
            ["Testiorganisaatio", "1", team, None, "Komissio", commission],
        ])
    return workbook_bytes(rows)


def multiple_days_bytes():
    return workbook_bytes([
        ["Org", "1", "id", None, "Komissio", -2, "04.11.2026"],
        ["Org", "1", "id", "A", "Maksu", 200, datetime(2026, 11, 4, 9)],
        ["Org", "1", "id", "A", "Maksu", 100, "03.11.2026"],
        ["Org", "1", "id", None, "Komissio", -1, "2026-11-03T15:30:00"],
        ["Org", "1", "id", "B", "Maksu", 50, "03.11.2026 12:00"],
    ], HEADERS + ["Tilitysaika"])


def replace_part(data, name, transform):
    output = BytesIO()
    with ZipFile(BytesIO(data)) as source, ZipFile(output, "w", ZIP_DEFLATED) as target:
        for member in source.infolist():
            content = source.read(member)
            target.writestr(member.filename, transform(content) if member.filename == name else content)
    return output.getvalue()


class SecurityTests(unittest.TestCase):
    def report(self):
        return workbook_bytes([["Org", "1", "id", "A", "Maksu", 10]])

    def test_upload_and_zip_expansion_limits(self):
        with self.assertRaisesRegex(ValueError, "5 MiB"):
            analyze(b"x" * (5 * 1024 * 1024 + 1))
        output = BytesIO()
        with ZipFile(output, "w", ZIP_DEFLATED) as archive:
            archive.writestr("xl/bomb.xml", b"x" * (20 * 1024 * 1024 + 1))
        with self.assertRaises(ValueError):
            analyze(output.getvalue())

    def test_dtd_and_entities_are_rejected(self):
        data = replace_part(self.report(), "xl/worksheets/sheet1.xml", lambda xml:
                            b'<!DOCTYPE worksheet [<!ENTITY payload "expanded">]>' + xml)
        with self.assertRaisesRegex(ValueError, "XML"):
            analyze(data)

    def test_inflated_row_and_column_coordinates(self):
        for old, new in [(b'A1:F2', b'A1:XFD1048576'),
                         (b'r="2"', b'r="1048576"'),
                         (b'r="A2"', b'r="XFD2"')]:
            with self.subTest(new=new):
                data = replace_part(self.report(), "xl/worksheets/sheet1.xml",
                                    lambda xml: xml.replace(old, new))
                with self.assertRaises(ValueError):
                    analyze(data)

    def test_forged_small_dimension_does_not_hide_rows(self):
        data = replace_part(self.report(), "xl/worksheets/sheet1.xml",
                            lambda xml: xml.replace(b'A1:F2', b'A1:A1'))
        self.assertEqual(len(analyze(data).details), 1)

    def test_malformed_worksheet_is_a_validation_error(self):
        data = replace_part(self.report(), "xl/worksheets/sheet1.xml", lambda xml: xml[:-10])
        with self.assertRaises(ValueError):
            analyze(data)

    def test_resource_limits(self):
        for setting in ["MAX_ARCHIVE_MEMBERS", "MAX_CELLS", "MAX_XML_ELEMENTS", "MAX_ROWS",
                        "MAX_COLUMNS"]:
            with self.subTest(setting=setting), patch(
                "tiliraportti_analyysi.analysis." + setting, 0
            ), self.assertRaises(ValueError):
                analyze(self.report())

    def test_extreme_amounts_are_rejected(self):
        for amount in ["1e999999", "1e1000000", "1e-999999", "0e999999", "NaN", "Infinity"]:
            with self.subTest(amount=amount), self.assertRaises(ValueError):
                analyze(workbook_bytes([["Org", "1", "id", "A", "Maksu", amount]]))

    def test_csv_formula_injection(self):
        for value in ['=HYPERLINK("https://example.invalid")', "+1+1", "-1+1", "@SUM(A1)",
                      "  =1+1", "\t=1+1", "\r=1+1", "\n=1+1", "\ufeff=1+1"]:
            with self.subTest(value=value):
                data = csv_bytes([{"Joukkue": value, "Summa": Decimal("-10.25")}])
                row = next(csv.DictReader(StringIO(data.decode("utf-8-sig")), delimiter=";"))
                self.assertEqual(row["Joukkue"], "'" + value)
                self.assertEqual(row["Summa"], "-10,25")


class AnalysisTests(unittest.TestCase):
    def test_groups_by_team_and_calendar_day(self):
        result = analyze(multiple_days_bytes())
        self.assertEqual([(row["Joukkue"], row["Tilityspäivä"]) for row in result.summary],
                         [("A", date(2026, 11, 3)), ("B", date(2026, 11, 3)),
                          ("Kohdistamatta", date(2026, 11, 3)), ("A", date(2026, 11, 4))])
        self.assertEqual(result.summary[3]["Nettotilitys (€)"], 198)
        self.assertEqual(result.summary[3]["Kulut (%)"], 1)
        self.assertEqual(result.summary[3]["Maksuja (kpl)"], 1)
        self.assertEqual(len(result.warnings), 1)
        self.assertEqual(result.details[0]["Joukkue"], "A")
        exported = list(csv.DictReader(
            StringIO(csv_bytes(result.summary).decode("utf-8-sig")), delimiter=";"))
        self.assertEqual(exported[3]["Tilityspäivä"], "2026-11-04")

    def test_single_day_and_missing_dates(self):
        for value in [date(2026, 11, 3), None]:
            with self.subTest(value=value):
                result = analyze(workbook_bytes([
                    ["Org", "1", "id", "A", "Maksu", 100, value],
                    ["Org", "1", "id", None, "Komissio", -1, value],
                ], HEADERS + ["Tilitysaika"]))
                self.assertEqual(len(result.summary), 1)
                self.assertEqual(result.summary[0]["Tilityspäivä"], value)
                self.assertEqual(result.summary[0]["Nettotilitys (€)"], 99)
                self.assertEqual(result.warnings, [])

    def test_invalid_settlement_date(self):
        with self.assertRaisesRegex(ValueError, "Rivi 2: Tilitysaika"):
            analyze(workbook_bytes([
                ["Org", "1", "id", "A", "Maksu", 100, "31.11.2026"],
            ], HEADERS + ["Tilitysaika"]))

    def test_example_totals_and_attribution(self):
        result = analyze(example_bytes())
        self.assertEqual(result.warnings, [])
        self.assertEqual(len(result.details), 9)
        self.assertEqual(len(result.summary), 3)
        for field, expected in [("Maksut (€)", "1660"), ("Palvelumaksut (€)", "16.49"),
                                ("Komissiot (€)", "7.76"), ("Nettotilitys (€)", "1635.75")]:
            self.assertEqual(sum(row[field] for row in result.summary), Decimal(expected))
        teams = {row["Joukkue"]: row for row in result.summary}
        self.assertEqual(teams["Joukkue A"]["Maksut (€)"], Decimal("1260"))
        self.assertEqual(teams["Joukkue A"]["Palvelumaksut (€)"], Decimal("11.64"))
        self.assertEqual(teams["Joukkue B"]["Komissiot (€)"], Decimal("0.42"))

    def test_cost_before_payment_and_separate_settlements(self):
        result = analyze(workbook_bytes([
            ["Org", "1", "id", None, "Komissio", "-0,42"],
            ["Org", "2", "id", "B", "Maksu", 50],
            ["Org", "1", "id", "A", "Maksu", 100],
        ]))
        self.assertEqual(result.warnings, [])
        self.assertEqual(result.summary[0]["Komissiot (€)"], Decimal("0.42"))
        self.assertEqual(result.summary[1]["Komissiot (€)"], 0)

    def test_ambiguous_and_missing_identifiers_remain_unallocated(self):
        result = analyze(workbook_bytes([
            ["Org", "1", "id", "A", "Maksu", 10],
            ["Org", "1", "id", "B", "Maksu", 20],
            ["Org", "1", "id", None, "Komissio", -1],
            ["Org", "1", None, None, "Palvelumaksu", -2],
        ]))
        self.assertEqual(len(result.warnings), 2)
        unallocated = next(row for row in result.summary if row["Joukkue"] == "Kohdistamatta")
        self.assertEqual(unallocated["Kulut yhteensä (€)"], 3)
        self.assertIsNone(unallocated["Kulut (%)"])

    def test_refund_keeps_sign(self):
        result = analyze(workbook_bytes([
            ["Org", "1", "id", "A", "Maksu", -10],
            ["Org", "1", "id", None, "Komissio", 1],
        ]))
        self.assertEqual(result.summary[0]["Nettotilitys (€)"], -9)
        self.assertEqual(result.summary[0]["Komissiot (€)"], -1)

    def test_invalid_input(self):
        for data in [b"not excel", workbook_bytes([], ["Joukkueet"]),
                     workbook_bytes([]),
                     workbook_bytes([["Org", "1", "id", "A", "Maksu", None]]),
                     workbook_bytes([["Org", "1", "id", "A", "Muu", 10]])]:
            with self.subTest(data_length=len(data)), self.assertRaises(ValueError):
                analyze(data)


class InterfaceTests(unittest.TestCase):
    def app(self):
        return AppTest.from_file(str(ROOT / "src/tiliraportti_analyysi/app.py"), default_timeout=15)

    def test_empty_upload(self):
        app = self.app().run()
        self.assertFalse(app.exception)
        self.assertEqual(len(app.info), 1)

    def test_report_and_team_filter(self):
        with patch("streamlit.file_uploader", return_value=BytesIO(example_bytes())):
            app = self.app().run()
            self.assertFalse(app.exception)
            self.assertEqual(app.metric[0].value, "1 660,00 €")
            self.assertEqual(app.metric[3].label, "Kulut yhteensä")
            self.assertEqual(app.metric[3].value, "24,25 €")
            self.assertEqual(len(app.dataframe[0].value), 3)
            app.multiselect[0].set_value(["Joukkue B"]).run()
            self.assertFalse(app.exception)
            self.assertEqual(app.metric[0].value, "80,00 €")
            self.assertEqual(app.metric[3].value, "1,07 €")
            self.assertEqual(len(app.dataframe[0].value), 1)

    def test_multiple_days_and_team_filter(self):
        with patch("streamlit.file_uploader", return_value=BytesIO(multiple_days_bytes())):
            app = self.app().run()
            self.assertFalse(app.exception)
            self.assertEqual(app.multiselect[0].options, ["A", "B", "Kohdistamatta"])
            self.assertEqual(len(app.dataframe[0].value), 4)
            app.multiselect[0].set_value(["A"]).run()
            self.assertFalse(app.exception)
            self.assertEqual(len(app.dataframe[0].value), 2)
            self.assertEqual(app.metric[0].value, "300,00 €")
            self.assertIn("Tilityspäivä", app.dataframe[0].value.columns)

    def test_invalid_upload_displays_error(self):
        with patch("streamlit.file_uploader", return_value=BytesIO(b"broken")):
            app = self.app().run()
            self.assertFalse(app.exception)
            self.assertEqual(len(app.error), 1)


if __name__ == "__main__":
    unittest.main()
