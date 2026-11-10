from contextlib import contextmanager
from copy import deepcopy
from datetime import date
from decimal import Decimal, ROUND_CEILING
from io import BytesIO
from pathlib import Path
from typing import Annotated, Literal
from zipfile import BadZipFile, ZipFile

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, Response
from lxml.etree import XMLSyntaxError
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

ROOT = Path(__file__).resolve().parent
UPLOAD_LIMIT = 10 * 1024 * 1024
PASS_SCORE = "RequiredExamScore"
ALIASES = {
    **{f"AA{i}": f"MA{i}" for i in range(1, 8)},
    "MAAve": "MActAve",
    "MAAve25": "MActAve25",
    "FAAve": "FActAve",
    "FAAve25": "FActAve25",
}
app = FastAPI(title="EtoGradeMo")


def report_pages(document, report, show_pass_score):
    pages = []
    for element in document.element.body:
        if element.tag == qn("w:sectPr"):
            continue
        if "".join(element.xpath(".//w:t/text()")) == "Date: (Input)":
            pages.append([])
        pages[-1].append(element)

    selected = {"midterm": pages[:1], "finals": pages[1:], "full": pages}[report]
    for page in selected:
        if not show_pass_score:
            page[:] = [
                element for element in page
                if not "".join(element.xpath(".//w:t/text()")).startswith(
                    ("Final Exam Needed to Pass", "*Grade Needed to Pass")
                )
            ]
        while page and page[-1].tag == qn("w:p") and not page[-1].xpath(".//w:t"):
            page.pop()
    return selected


def score_fields(pages):
    for page in pages:
        for table in (element for element in page if element.tag == qn("w:tbl")):
            rows = table.xpath("./w:tr")
            if len(rows) == 1:
                yield rows[0].xpath("./w:tc")[1], PASS_SCORE
                continue
            for cell in rows[1].xpath("./w:tc"):
                token = "".join("".join(cell.xpath(".//w:t/text()")).split())
                if not token:
                    continue
                # The finals activities row repeats the midterm weight placeholder.
                if token == "MAAve25" and "FAAve" in "".join(rows[1].xpath(".//w:t/text()")).replace(" ", ""):
                    token = "FAAve25"
                yield cell, ALIASES.get(token, token)


def upload_contents(workbook):
    if Path(workbook.filename or "").suffix.lower() not in {".xlsx", ".xlsm"}:
        raise HTTPException(400, "Upload an Excel workbook (.xlsx or .xlsm).")
    contents = workbook.file.read(UPLOAD_LIMIT + 1)
    if len(contents) > UPLOAD_LIMIT:
        raise HTTPException(413, "Upload a workbook no larger than 10 MB.")
    return contents


@contextmanager
def open_class_record(contents):
    with ZipFile(BytesIO(contents)) as archive:
        if sum(member.file_size for member in archive.infolist()) > 50 * 1024 * 1024:
            raise ValueError("The expanded workbook exceeds 50 MB.")
    book = load_workbook(BytesIO(contents), read_only=True, data_only=True)
    try:
        yield book
    finally:
        book.close()


def read_students(contents, required, sheet_name=None):
    with open_class_record(contents) as book:
        sheets = {sheet.title: sheet for sheet in book.worksheets}
        if sheet_name is not None:
            if sheet_name not in sheets:
                raise ValueError("The selected worksheet is not in this workbook.")
            sheet = sheets[sheet_name]
        elif len(sheets) == 1:
            sheet = next(iter(sheets.values()))
        else:
            raise ValueError("Select a worksheet before generating the report.")

        header = next(sheet.iter_rows(max_row=1, max_col=256, values_only=True), ())
        names = [str(value).strip() for value in header]
        if not required.issubset(names):
            raise ValueError("The selected worksheet is missing required class-record headers in row 1.")
        if any(names.count(name) != 1 for name in required):
            raise ValueError("The class record contains duplicate required headers.")
        columns = {name: names.index(name) for name in required}
        students = []
        for number, row in enumerate(
            sheet.iter_rows(min_row=2, max_col=max(columns.values()) + 1), start=2
        ):
            name = row[columns["StudentName"]].value
            if not isinstance(name, str) or not name.strip():
                continue
            student = {}
            for field, index in columns.items():
                cell = row[index]
                if cell.data_type == "e":
                    raise ValueError(f"Excel error in {field}, row {number}.")
                student[field] = cell.value
            student["StudentName"] = name.strip()
            students.append(student)
            if len(students) > 100:
                raise ValueError("Upload a class record with at most 100 students.")
        if not students:
            raise ValueError("The class record contains no student names.")
        return students


def required_exam_score(student):
    components = [student["MidtermGrade40"], student["FClassPerf20"]]
    if not all(isinstance(value, (int, float)) for value in components):
        return None
    # Snap cached Excel arithmetic to 10 decimal places before rounding upward.
    remaining = Decimal(str(round(75 - sum(components), 10)))
    return (remaining / Decimal("0.4")).quantize(Decimal("0.01"), rounding=ROUND_CEILING)


def display(value):
    if value is None:
        return ""
    if isinstance(value, (int, float, Decimal)):
        return f"{value:.2f}".rstrip("0").rstrip(".")
    return str(value)


def fill_text(element, value):
    nodes = element.xpath(".//w:t")
    nodes[0].text = value
    for node in nodes[1:]:
        node.text = ""


def build_report(contents, fields, report, show_pass_score, sheet_name=None):
    document = Document(ROOT / "Template.docx")
    pages = report_pages(document, report, show_pass_score)
    required = {"StudentName", *(field for _, field in score_fields(pages))}
    required.discard(PASS_SCORE)
    students = read_students(contents, required, sheet_name)
    body = document.element.body
    for element in list(body):
        if element.tag != qn("w:sectPr"):
            body.remove(element)

    for student in students:
        for template in pages:
            page = deepcopy(template)
            for paragraph in page:
                if paragraph.tag != qn("w:p"):
                    continue
                for node in paragraph.xpath(".//w:t"):
                    if node.text == "StudentName":
                        node.text = student["StudentName"]
                    elif node.text in fields:
                        node.text = fields[node.text]
                if "".join(paragraph.xpath(".//w:t/text()")).startswith("*Grade Needed to Pass"):
                    fill_text(
                        paragraph,
                        "*The score assumes all other grades are complete. "
                        "A score above 100 means the final exam alone cannot bring the course grade to 75.",
                    )
                for marker in paragraph.xpath(".//w:lastRenderedPageBreak"):
                    marker.getparent().remove(marker)
            for cell, field in list(score_fields([page])):
                value = required_exam_score(student) if field == PASS_SCORE else student[field]
                fill_text(cell, display(value))
            if len(body) > 1:
                Paragraph(page[0], document).paragraph_format.page_break_before = True
            for element in page:
                body.sectPr.addprevious(element)
    output = BytesIO()
    document.save(output)
    return output.getvalue()


@app.get("/", response_class=HTMLResponse)
def index():
    return (ROOT / "index.html").read_text(encoding="utf-8")


@app.get("/download-class-record")
def download_class_record():
    filename = "with_macros_-_2026_blank_grading_sheet_with_autotransmuter.xlsm"
    return FileResponse(
        ROOT / filename,
        filename=filename,
        media_type="application/vnd.ms-excel.sheet.macroEnabled.12",
    )


@app.post("/worksheets")
def worksheets(workbook: Annotated[UploadFile, File()]):
    contents = upload_contents(workbook)
    try:
        with open_class_record(contents) as book:
            names = [sheet.title for sheet in book.worksheets]
            if not names:
                raise ValueError("The workbook does not contain any worksheets.")
            return {"worksheets": names}
    except (BadZipFile, InvalidFileException, XMLSyntaxError, KeyError, OSError) as error:
        raise HTTPException(400, "This file could not be read as an Excel workbook.") from error
    except ValueError as error:
        raise HTTPException(400, str(error)) from error


@app.post("/generate")
def generate(
    workbook: Annotated[UploadFile, File()],
    report_date: Annotated[date, Form()],
    section: Annotated[str, Form(min_length=1, max_length=60)],
    course_code: Annotated[str, Form(min_length=1, max_length=30)],
    course_title: Annotated[str, Form(min_length=1, max_length=100)],
    report: Annotated[Literal["midterm", "finals", "full"], Form()],
    show_pass_score: Annotated[bool, Form()] = False,
    sheet_name: Annotated[str | None, Form()] = None,
):
    contents = upload_contents(workbook)
    if not all(value.strip() for value in (section, course_code, course_title)):
        raise HTTPException(400, "Enter the section, course code, and course title.")

    fields = {
        "Date: (Input)": f"Date: {report_date:%B} {report_date.day}, {report_date.year}",
        "Section: Section (Input)": f"Section: {section.strip()}",
        "Course Code: Course Code (Input)": f"Course Code: {course_code.strip()}",
        "Course Title: Course Title (Input)": f"Course Title: {course_title.strip()}",
    }
    try:
        result = build_report(contents, fields, report, show_pass_score, sheet_name)
    except (BadZipFile, InvalidFileException, XMLSyntaxError, KeyError, OSError) as error:
        raise HTTPException(400, "This file could not be read as an Excel workbook.") from error
    except ValueError as error:
        raise HTTPException(400, str(error)) from error
    return Response(
        result,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="{report}-grade-breakdown.docx"'},
    )
