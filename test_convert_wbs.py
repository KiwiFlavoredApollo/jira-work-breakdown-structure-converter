"""소수 가중치, 달력 경계 및 상단 양식 유지 회귀 테스트."""

import tempfile
import unittest
import xml.etree.ElementTree as ET
from datetime import date, timedelta
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

from convert_wbs import Issue, convert, tag


class ConversionTest(unittest.TestCase):
    def test_weights_calendar_boundaries_and_template_preservation(self) -> None:
        template = Path(__file__).with_name("work-breakdown-structure-template.xlsx")
        with ZipFile(template) as archive:
            original = ET.fromstring(archive.read("xl/worksheets/sheet2.xml"))
        data = original.find(tag("sheetData"))
        assert data is not None
        header = next(row for row in data if row.get("r") == "7")
        first_cell = next(cell for cell in header if cell.get("r") == "K7")
        first = date(1899, 12, 30) + timedelta(days=float(first_cell.findtext(tag("v"))))
        start = first + timedelta(days=1)
        end = first + timedelta(days=2)
        issues = [
            Issue("T-1", "Epic", "Epic", "", "", start, end, 0.125, False, ""),
            Issue("T-2", "Story", "Story", "T-1", "", start, end, 1.75, True, "", 1),
            Issue("T-3", "Subtask", "Subtask", "T-2", "", start, start, 2.5, False, "", 2),
            Issue("T-4", "Task", "Task", "", "", start, end, 0.5, False, ""),
            Issue("T-5", "에픽", "에픽", "", "", start, end, 1.0, False, ""),
            Issue("T-6", "스토리", "스토리", "T-5", "", start, end, 1.0, False, "", 1),
        ]
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "result.xlsx"
            with patch("convert_wbs.read_issues", return_value=issues):
                convert(Path(directory) / "input.csv", template, output)
                convert(Path(directory) / "input.csv", template, output)
            with ZipFile(output) as archive:
                result = ET.fromstring(archive.read("xl/worksheets/sheet2.xml"))
                styles = ET.fromstring(archive.read("xl/styles.xml"))
                self.assertIsNone(archive.testzip())
        result_data = result.find(tag("sheetData"))
        assert result_data is not None
        before = [ET.tostring(row) for row in data if int(row.get("r")) <= 10]
        after = [ET.tostring(row) for row in result_data if int(row.get("r")) <= 10]
        self.assertEqual(before, after)
        self.assertEqual(ET.tostring(original.find(tag("mergeCells"))), ET.tostring(result.find(tag("mergeCells"))))
        cells = {cell.get("r"): cell for row in result_data for cell in row}
        xfs = styles.find(tag("cellXfs"))
        fills = styles.find(tag("fills"))
        assert xfs is not None and fills is not None

        def color(reference: str) -> str | None:
            style = xfs[int(cells[reference].get("s", "0"))]
            foreground = fills[int(style.get("fillId", "0"))].find(f".//{tag('fgColor')}")
            return foreground.get("rgb") if foreground is not None else None

        for row, issue in enumerate(issues, 11):
            for letter in "ABCDEFGHIJ":
                expected = None
                if issue.kind in {"Epic", "에픽"}:
                    expected = "FFBDD7EE"
                elif issue.kind in {"Story", "스토리"}:
                    expected = "FFEDF2FA"
                elif letter in "FJ":
                    expected = "FFFFF2CC"
                self.assertEqual(color(f"{letter}{row}"), expected, f"{letter}{row}")
            weight = cells[f"I{row}"]
            self.assertEqual(float(weight.findtext(tag("v"))), issue.weight)
            self.assertEqual(xfs[int(weight.get("s"))].get("numFmtId"), "0")
            self.assertNotIn(color(f"K{row}"), {"FFB4C7E7", "FFDDEBF7"})
            self.assertNotIn(color(f"N{row}"), {"FFB4C7E7", "FFDDEBF7"})
        for row in (11, 12, 15, 16):
            self.assertEqual(color(f"L{row}"), "FFB4C7E7")
            self.assertEqual(color(f"M{row}"), "FFB4C7E7")
        self.assertEqual(color("L13"), "FFDDEBF7")
        self.assertNotEqual(color("M13"), "FFDDEBF7")
        self.assertIsNone(color("L14"))


if __name__ == "__main__":
    unittest.main()
