from __future__ import annotations

import argparse
from pathlib import Path
import re
import os

import win32com.client as win32


ROOT = Path(__file__).resolve().parent
POWERPOINT_REPORT = ROOT / "Microsoft Office 2024 - PowerPoint Win32.docx"
WORD_REPORT = ROOT / "Microsoft Office 2024 - Word Win32.docx"
OUTPUT = ROOT / "Microsoft Office 2024 - PowerPoint and Word Win32_combined.docx"

HEADING_LEVELS = {
    "测试用例": 1,
    "一级标题": 2,
    "二级标题": 3,
    "三级标题": 4,
    "四级标题": 5,
    "五级标题": 6,
}


def copy_list_level(source, target) -> None:
    for property_name in (
        "Alignment",
        "NumberFormat",
        "NumberPosition",
        "NumberStyle",
        "ResetOnHigher",
        "StartAt",
        "TabPosition",
        "TextPosition",
        "TrailingCharacter",
    ):
        setattr(target, property_name, getattr(source, property_name))


def paragraph_style_name(paragraph) -> str:
    style = paragraph.Style
    return getattr(style, "NameLocal", str(style))


def heading_paragraphs(document):
    headings = []
    for style in HEADING_LEVELS:
        search = document.Content.Duplicate
        search.Find.ClearFormatting()
        search.Find.Text = ""
        search.Find.Style = style
        search.Find.Format = True
        search.Find.Forward = True
        search.Find.Wrap = win32.constants.wdFindStop
        while search.Find.Execute():
            paragraph = search.Paragraphs.Item(1)
            if paragraph.Range.Text.strip() and not paragraph.Range.Information(win32.constants.wdWithInTable):
                headings.append(paragraph)
            end = paragraph.Range.End
            if end >= document.Content.End:
                break
            search.SetRange(end, document.Content.End)
    yield from sorted(headings, key=lambda paragraph: paragraph.Range.Start)


def append_document_in_chunks(target_document, source_document, chunk_size: int = 100) -> None:
    paragraph_count = source_document.Paragraphs.Count
    source_headings = [
        (paragraph.Range.Start, paragraph_style_name(paragraph))
        for paragraph in heading_paragraphs(source_document)
    ]
    for first in range(1, paragraph_count + 1, chunk_size):
        last = min(first + chunk_size - 1, paragraph_count)
        start = source_document.Paragraphs.Item(first).Range.Start
        end = source_document.Paragraphs.Item(last).Range.End
        if last == paragraph_count:
            end = min(end, source_document.Content.End - 1)
        headings = [(position - start, style) for position, style in source_headings if start <= position < end]
        source_document.Range(start, end).Copy()
        insertion_start = target_document.Content.End - 1
        insertion = target_document.Range(
            insertion_start,
            insertion_start,
        )
        insertion.PasteAndFormat(win32.constants.wdFormatOriginalFormatting)
        for offset, style in headings:
            heading = target_document.Range(insertion_start + offset, insertion_start + offset).Paragraphs.Item(1)
            heading.Range.Style = style
            heading.OutlineLevel = HEADING_LEVELS[style]


def strip_heading_digits(document) -> None:
    for paragraph in heading_paragraphs(document):
        text_range = paragraph.Range.Duplicate
        text_range.MoveEnd(win32.constants.wdCharacter, -1)
        original = text_range.Text
        cleaned = re.sub(r"\d+", "", original)
        cleaned = re.sub(r"[ \t]{2,}", " ", cleaned).strip()
        if cleaned != original:
            text_range.Text = cleaned


def continue_heading_numbering(document):
    source_template = document.Paragraphs.Item(1).Range.ListFormat.ListTemplate
    template = document.ListTemplates.Add(True, "MergedReportHeadings")
    for level in range(1, 7):
        copy_list_level(source_template.ListLevels.Item(level), template.ListLevels.Item(level))
    for index, paragraph in enumerate(heading_paragraphs(document)):
        level = HEADING_LEVELS[paragraph_style_name(paragraph)]
        paragraph.Range.ListFormat.ApplyListTemplateWithLevel(
            template,
            index > 0,
            win32.constants.wdListApplyToSelection,
            win32.constants.wdWord10ListBehavior,
            level,
        )
        paragraph.OutlineLevel = level
    return template


def collect_test_result_groups(document):
    headings = []
    for paragraph in heading_paragraphs(document):
        level = HEADING_LEVELS[paragraph_style_name(paragraph)]
        if level <= 3:
            headings.append(
                (
                    paragraph.Range.Start,
                    level,
                    paragraph.Range.ListFormat.ListString,
                    paragraph.Range.Text.rstrip("\r\x07").strip(),
                )
            )
    headings.sort(key=lambda item: item[0])

    products = []
    product = None
    category = None
    for _, level, number, text in headings:
        if level == 1:
            if "测试用例" not in text:
                continue
            result_name = "PowerPoint测试结果" if "PowerPoint" in text else "Word测试结果"
            product = {"name": result_name, "categories": []}
            products.append(product)
            category = None
        elif level == 2 and product is not None:
            category = {"number": number, "name": text, "cases": []}
            product["categories"].append(category)
        elif level == 3 and category is not None:
            category["cases"].append((number, text))
    return products


def append_outline_heading(document, outline_template, text: str, style: str, level: int) -> None:
    insertion_point = document.Content.End - 1
    heading_range = document.Range(insertion_point, insertion_point)
    heading_range.Text = text
    heading_range.Style = style
    heading_range.ListFormat.ApplyListTemplateWithLevel(
        outline_template,
        True,
        win32.constants.wdListApplyToSelection,
        win32.constants.wdWord10ListBehavior,
        level,
    )
    heading_range.InsertParagraphAfter()
    trailing = document.Range(document.Content.End - 1, document.Content.End - 1)
    trailing.Style = win32.constants.wdStyleNormal
    trailing.ListFormat.RemoveNumbers()
    trailing.ParagraphFormat.OutlineLevel = win32.constants.wdOutlineLevelBodyText


def add_result_table(document, product) -> None:
    categories = [category for category in product["categories"] if category["cases"]]
    row_count = 1 + sum(len(category["cases"]) for category in categories)
    insertion_point = document.Content.End - 1
    table_range = document.Range(insertion_point, insertion_point)
    table = document.Tables.Add(table_range, row_count, 3)

    page_setup = table.Range.Sections.Item(1).PageSetup
    usable_width = page_setup.PageWidth - page_setup.LeftMargin - page_setup.RightMargin - page_setup.Gutter - 12
    table.AllowAutoFit = False
    table.AutoFitBehavior(win32.constants.wdAutoFitFixed)
    table.PreferredWidthType = win32.constants.wdPreferredWidthPoints
    table.PreferredWidth = usable_width
    table.Rows.LeftIndent = 0
    table.Columns.Item(1).Width = usable_width * 0.18
    table.Columns.Item(2).Width = usable_width * 0.49
    table.Columns.Item(3).Width = usable_width * 0.33
    table.Borders.Enable = True
    table.Rows.AllowBreakAcrossPages = True

    table.Cell(1, 1).Merge(table.Cell(1, 2))
    table.Cell(1, 1).Range.Text = "测试用例"
    table.Cell(1, 2).Range.Text = "测试结果"
    table.Rows.Item(1).HeadingFormat = True
    table.Rows.Item(1).Range.Bold = True
    table.Rows.Item(1).Range.ParagraphFormat.Alignment = win32.constants.wdAlignParagraphCenter
    table.Rows.Item(1).Shading.BackgroundPatternColor = 12632256

    row_index = 2
    for category in categories:
        first_row = row_index
        for case_number, case_name in category["cases"]:
            table.Cell(row_index, 2).Range.Text = f"{case_number}  {case_name}"
            table.Cell(row_index, 3).Range.Text = "通过"
            row_index += 1
        last_row = row_index - 1
        if first_row != last_row:
            table.Cell(first_row, 1).Merge(table.Cell(last_row, 1))
        table.Cell(first_row, 1).Range.Text = f"{category['number']}  {category['name']}"

    table.Range.Cells.VerticalAlignment = win32.constants.wdCellAlignVerticalCenter
    table.Range.Style = win32.constants.wdStyleNormal
    table.Range.ListFormat.RemoveNumbers()
    table.Range.ParagraphFormat.OutlineLevel = win32.constants.wdOutlineLevelBodyText
    table.Range.ParagraphFormat.LeftIndent = 0
    table.Range.ParagraphFormat.RightIndent = 0
    table.Range.ParagraphFormat.FirstLineIndent = 0
    table.Range.ParagraphFormat.SpaceBefore = 0
    table.Range.ParagraphFormat.SpaceAfter = 0
    table.Range.Font.Size = 10.5
    document.Range(document.Content.End - 1, document.Content.End - 1).InsertParagraphAfter()
    document.Range(table.Range.End, table.Range.End).Paragraphs.Item(1).Range.ListFormat.RemoveNumbers()


def append_test_results(document, outline_template) -> None:
    products = collect_test_result_groups(document)
    if not products:
        raise ValueError("No test-case directory hierarchy was found")

    insertion = document.Range(document.Content.End - 1, document.Content.End - 1)
    insertion.InsertBreak(win32.constants.wdPageBreak)
    append_outline_heading(document, outline_template, "测试结果", "测试用例", 1)

    for product in products:
        append_outline_heading(document, outline_template, product["name"], "一级标题", 2)
        add_result_table(document, product)


def merge_reports(input1: Path, input2: Path, output: Path) -> None:
    word = win32.gencache.EnsureDispatch(win32.DispatchEx("Word.Application"))
    word.Visible = False
    word.DisplayAlerts = 0
    document = None
    source_document = None
    try:
        document = word.Documents.Open(str(input1), False, True)
        if output.exists():
            os.remove(OUTPUT)
        document.SaveAs2(str(output), 16)
        document.TrackRevisions = False

        insertion = document.Range(document.Content.End - 1, document.Content.End - 1)
        insertion.InsertBreak(win32.constants.wdSectionBreakNextPage)
        source_document = word.Documents.Open(str(input2), False, True)
        append_document_in_chunks(document, source_document)
        source_document.Close(False)
        source_document = None
        document.Save()

        strip_heading_digits(document)
        outline_template = continue_heading_numbering(document)
        append_test_results(document, outline_template)
        document.Fields.Update()
        document.Save()
    finally:
        if source_document is not None:
            source_document.Close(False)
        if document is not None:
            document.Close(False)
        word.Quit()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Merge two Word reports.")
    parser.add_argument("--input1", type=Path, default=POWERPOINT_REPORT, help="First Word report")
    parser.add_argument("--input2", type=Path, default=WORD_REPORT, help="Second Word report")
    parser.add_argument("--output", type=Path, default=OUTPUT, help="Output Word report (overwritten if it exists)")
    args = parser.parse_args()
    args.input1 = args.input1.resolve()
    args.input2 = args.input2.resolve()
    args.output = args.output.resolve()
    for path in (args.input1, args.input2):
        if not path.is_file():
            parser.error(f"Input file does not exist: {path}")
    if args.output in (args.input1, args.input2):
        parser.error("Output must not be the same as either input file.")
    return args


if __name__ == "__main__":
    args = parse_args()
    args.output.unlink(missing_ok=True)
    merge_reports(args.input1, args.input2, args.output)
    print(f"Created: {args.output}")