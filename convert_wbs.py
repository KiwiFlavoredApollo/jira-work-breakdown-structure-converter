"""Jira CSV를 WBS로 변환합니다. Python 3.10+, 외부 패키지 불필요.

실행: python convert_wbs.py [Jira.csv] [--template 양식.xlsx] [-o 결과.xlsx]
시작일/기한이 없는 상위 이슈는 하위 이슈의 일정 범위를 사용합니다.
양쪽 일정이 없는 말단 이슈는 날짜와 간트 영역을 비워 둡니다.
"""

from __future__ import annotations

import argparse
import csv
import io
import math
import re
import sys
import xml.etree.ElementTree as ET
from collections import defaultdict
from copy import deepcopy
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path
from zipfile import ZipFile


NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
ET.register_namespace("", NS)
ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")


def tag(name: str) -> str:
    return f"{{{NS}}}{name}"


def column(number: int) -> str:
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def parse_date(value: str) -> date | None:
    if not value.strip():
        return None
    match = re.match(r"(\d{1,2})/(\d{1,2})월/(\d{2,4})", value)
    if match:
        day, month, year = map(int, match.groups())
        return date(year + 2000 if year < 100 else year, month, day)
    for pattern in (r"(\d{4})-(\d{1,2})-(\d{1,2})", r"(\d{4})/(\d{1,2})/(\d{1,2})"):
        match = re.match(pattern, value)
        if match:
            return date(*map(int, match.groups()))
    raise ValueError(f"지원하지 않는 날짜 형식: {value!r}")


@dataclass
class Issue:
    key: str
    summary: str
    kind: str
    parent: str
    assignee: str
    start: date | None
    end: date | None
    weight: float | None
    done: bool
    project: str
    depth: int = 0


def read_issues(path: Path, encoding: str | None) -> list[Issue]:
    raw = path.read_bytes()
    if encoding:
        text = raw.decode(encoding)
    else:
        try:
            text = raw.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = raw.decode("cp949")
    reader = csv.DictReader(io.StringIO(text))
    required = {"이슈 키", "요약", "이슈 유형", "상태"}
    if not required.issubset(reader.fieldnames or []):
        raise ValueError(f"CSV 필수 열 누락: {', '.join(sorted(required - set(reader.fieldnames or [])))}")
    records = list(reader)
    ids = {r["이슈 ID"]: r["이슈 키"] for r in records if r.get("이슈 ID")}
    issues: dict[str, Issue] = {}
    for row in records:
        key = row["이슈 키"].strip()
        if not key or key in issues:
            raise ValueError(f"비어 있거나 중복된 이슈 키: {key!r}")
        try:
            point = row.get("사용자정의 필드 (Story point estimate)", "").strip()
            weight = float(point) if point else None
            if weight is not None and (not math.isfinite(weight) or weight < 0):
                raise ValueError("스토리 포인트는 0 이상의 유한한 숫자여야 합니다.")
            issues[key] = Issue(
                key, row["요약"], row["이슈 유형"],
                row.get("상위 키", "").strip() or ids.get(row.get("상위 항목", ""), ""),
                row.get("담당자", ""), parse_date(row.get("사용자정의 필드 (시작 날짜)", "")),
                parse_date(row.get("기한", "")), weight,
                row.get("상태 범주", "").strip().lower() in {"완료", "done"}
                or row["상태"].strip().lower() in {"완료", "done", "closed", "resolved"},
                row.get("프로젝트 이름", ""),
            )
        except ValueError as error:
            raise ValueError(f"{key}: {error}") from error
    if not issues:
        raise ValueError("CSV에 이슈가 없습니다.")
    children: dict[str, list[Issue]] = defaultdict(list)
    for issue in issues.values():
        if issue.parent and issue.parent not in issues:
            print(f"경고: {issue.key}의 상위 이슈 {issue.parent}가 CSV에 없어 최상위로 배치합니다.", file=sys.stderr)
        children[issue.parent if issue.parent in issues else ""].append(issue)
    ordered: list[Issue] = []
    visiting: set[str] = set()
    visited: set[str] = set()

    def visit(issue: Issue, depth: int) -> None:
        if issue.key in visiting:
            raise ValueError(f"상위 항목 순환 참조: {issue.key}")
        if issue.key in visited:
            return
        visiting.add(issue.key)
        issue.depth = depth
        ordered.append(issue)
        for child in sorted(children[issue.key], key=sort_key):
            visit(child, depth + 1)
        descendants = children[issue.key]
        issue.start = issue.start or min((c.start for c in descendants if c.start), default=None)
        issue.end = issue.end or max((c.end for c in descendants if c.end), default=None)
        if issue.start and issue.end and issue.start > issue.end:
            print(f"경고: {issue.key}의 시작일이 종료일보다 늦어 간트 막대를 표시하지 않습니다.", file=sys.stderr)
        visiting.remove(issue.key)
        visited.add(issue.key)

    for root in sorted(children[""], key=sort_key):
        visit(root, 0)
    for issue in issues.values():
        if issue.key not in visited:
            visit(issue, 0)  # 루트가 없는 순환도 검출합니다.
    return ordered


def sort_key(issue: Issue) -> tuple[int, str, int]:
    rank = {"에픽": 0, "epic": 0, "스토리": 1, "story": 1, "subtask": 3, "sub-task": 3, "하위 작업": 3}
    prefix, _, suffix = issue.key.rpartition("-")
    return rank.get(issue.kind.lower(), 2), prefix, int(suffix) if suffix.isdigit() else 0


def write_cell(cell: ET.Element, value: str | float | int | None) -> None:
    for child in list(cell):
        cell.remove(child)
    cell.attrib.pop("t", None)
    if value is None:
        return
    if isinstance(value, str):
        cell.set("t", "inlineStr")
        text = ET.SubElement(ET.SubElement(cell, tag("is")), tag("t"))
        text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        text.text = value
    else:
        ET.SubElement(cell, tag("v")).text = str(value)


def convert(csv_path: Path, template: Path, output: Path, encoding: str | None = None) -> int:
    if output.resolve() in {csv_path.resolve(), template.resolve()}:
        raise ValueError("출력 경로는 CSV 및 양식 파일과 달라야 합니다.")
    issues = read_issues(csv_path, encoding)
    with ZipFile(template) as archive:
        parts = {entry.filename: archive.read(entry) for entry in archive.infolist()}
    workbook = ET.fromstring(parts["xl/workbook.xml"])
    rels = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
    sheet = next((s for s in workbook.find(tag("sheets")) if s.get("name") == "프로젝트 WBS"), None)
    if sheet is None:
        raise ValueError("양식에 '프로젝트 WBS' 시트가 없습니다.")
    relationship = sheet.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
    target = next(r.get("Target", "") for r in rels if r.get("Id") == relationship)
    sheet_path = target.lstrip("/") if target.startswith("/") else "xl/" + target
    root = ET.fromstring(parts[sheet_path])
    data = root.find(tag("sheetData"))
    assert data is not None
    rows = {int(row.get("r", "0")): row for row in data}
    if 11 not in rows:
        raise ValueError("양식에 데이터 서식 행(11행)이 없습니다.")
    prototype = deepcopy(rows[11])
    # 상단 달력은 수정하지 않고 양식 7행의 Excel 날짜를 읽습니다.
    properties = workbook.find(tag("workbookPr"))
    use_1904 = properties is not None and properties.get("date1904") in {"1", "true"}
    epoch = date(1904, 1, 1) if use_1904 else date(1899, 12, 30)
    calendar: dict[int, date] = {}
    for cell in rows.get(7, []):
        letters = re.sub(r"\d", "", cell.get("r", ""))
        index = 0
        for letter in letters:
            index = index * 26 + ord(letter) - ord("A") + 1
        value = cell.find(tag("v"))
        if index >= 11 and value is not None and cell.get("t") not in {"s", "inlineStr"}:
            calendar[index] = epoch + timedelta(days=float(value.text or "0"))
    if not calendar:
        raise ValueError("양식 7행에 달력 날짜가 없습니다.")
    if len(issues) > 1048566:
        raise ValueError("Excel 최대 행 수를 초과합니다.")
    last_column = max(calendar)
    last_col = column(last_column)
    last_row = 10 + len(issues)
    styles = ET.fromstring(parts["xl/styles.xml"])
    fills = styles.find(tag("fills"))
    xfs = styles.find(tag("cellXfs"))
    assert fills is not None and xfs is not None
    fill_ids: dict[str, int] = {}
    for color in ("B4C7E7", "DDEBF7"):
        fill_ids[color] = len(fills)
        fill = ET.SubElement(fills, tag("fill"))
        pattern = ET.SubElement(fill, tag("patternFill"), patternType="solid")
        ET.SubElement(pattern, tag("fgColor"), rgb="FF" + color)
        ET.SubElement(pattern, tag("bgColor"), indexed="64")
    fills.set("count", str(len(fills)))
    style_cache: dict[tuple[str, str], str] = {}
    # 가중치의 원래 소수값을 표시합니다. 양식의 백분율/정수 서식은 사용하지 않습니다.
    weight_base = next(c for c in prototype if c.get("r") == "I11")
    weight_style = deepcopy(xfs[int(weight_base.get("s", "0"))])
    weight_style.set("numFmtId", "0")  # General: 0.5를 50%나 1로 표시하지 않습니다.
    weight_style.set("applyNumberFormat", "1")
    weight_style_id = str(len(xfs))
    xfs.append(weight_style)

    def colored_style(base: str, color: str) -> str:
        key = base, color
        if key not in style_cache:
            xf = deepcopy(xfs[int(base)])
            xf.set("fillId", str(fill_ids[color]) if color else "0")
            xf.set("applyFill", "1")
            style_cache[key] = str(len(xfs))
            xfs.append(xf)
        return style_cache[key]

    for row in list(data):
        if int(row.get("r", "0")) >= 11:
            data.remove(row)
    for number, issue in enumerate(issues, 11):
        row = ET.SubElement(data, tag("row"), r=str(number), ht="30", customHeight="1", outlineLevel=str(min(issue.depth, 7)))
        category = issue.summary if issue.kind.strip().lower() in {"에픽", "epic"} else None
        values = [category, issue.key, issue.summary, issue.kind, None, issue.assignee,
                  (issue.start - epoch).days if issue.start else None,
                  (issue.end - epoch).days if issue.end else None, issue.weight, int(issue.done)]
        for index in range(1, last_column + 1):
            source = prototype[index - 1] if index <= len(prototype) else prototype[-1]
            cell = deepcopy(source)
            cell.set("r", f"{column(index)}{number}")
            write_cell(cell, values[index - 1] if index <= 10 else None)
            if index == 9:
                cell.set("s", weight_style_id)
            # 양식 행의 배경색을 복사하지 않고 달력의 해당 기간만 칠합니다.
            cell.set("s", colored_style(cell.get("s", "0"), ""))
            if index in calendar and issue.start and issue.end:
                # 시작일과 종료일을 모두 포함하며 해당 기간 밖은 칠하지 않습니다.
                if issue.start <= calendar[index] <= issue.end and sort_key(issue)[0] in {0, 1, 3}:
                    color = "DDEBF7" if sort_key(issue)[0] == 3 else "B4C7E7"
                    cell.set("s", colored_style(cell.get("s", "0"), color))
            row.append(cell)
    for conditional in list(root.findall(tag("conditionalFormatting"))):
        root.remove(conditional)  # 양식의 진행률 색상 규칙 대신 이슈 유형 색상을 사용합니다.
    root.find(tag("dimension")).set("ref", f"A1:{last_col}{last_row}")
    for defined in workbook.findall(f"{tag('definedNames')}/{tag('definedName')}"):
        if defined.get("name") == "_xlnm.Print_Area" and "프로젝트 WBS" in (defined.text or ""):
            defined.text = f"'프로젝트 WBS'!$A$1:${last_col}${last_row}"
    xfs.set("count", str(len(xfs)))
    parts[sheet_path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    parts["xl/styles.xml"] = ET.tostring(styles, encoding="utf-8", xml_declaration=True)
    parts["xl/workbook.xml"] = ET.tostring(workbook, encoding="utf-8", xml_declaration=True)
    # 같은 이름의 결과 파일이 있으면 덮어씁니다.
    with output.open("wb") as stream, ZipFile(stream, "w") as archive:
        for name, content in parts.items():
            archive.writestr(name, content, compress_type=8)
    return len(issues)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("csv", nargs="?", type=Path, default=Path("Jira.csv"))
    parser.add_argument("--template", type=Path, default=Path(__file__).with_name("work-breakdown-structure-template.xlsx"))
    parser.add_argument("-o", "--output", type=Path, default=Path(f"work-breakdown-structure-{date.today():%Y-%m-%d}.xlsx"))
    parser.add_argument("--encoding", help="CSV 인코딩 지정 (기본: UTF-8, CP949 자동 감지)")
    args = parser.parse_args()
    try:
        count = convert(args.csv, args.template, args.output, args.encoding)
    except (OSError, ValueError, csv.Error) as error:
        parser.exit(1, f"변환 실패: {error}\n")
    print(f"완료: {args.output} ({count}개 이슈)")


if __name__ == "__main__":
    main()
