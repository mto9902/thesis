#!/usr/bin/env python3
"""Verify the direct-replication package against the published source PDF."""

from __future__ import annotations

import ast
import hashlib
import re
import sys
import zipfile
from pathlib import Path

import pdfplumber
import pypdfium2 as pdfium
from PIL import Image, ImageChops
from docx import Document


ROOT = Path(__file__).resolve().parents[1]
BUILD_SCRIPT = ROOT / "tools" / "build_direct_replication.py"
SOURCE_PDF = ROOT / "source" / "Steenkamp_2024_Digital_Badges.pdf"
MASTER_DOCX = ROOT / "outputs" / "Master_Thesis_Digital_Badges_Thai_IT_Students.docx"
EVIDENCE_DOCX = ROOT / "outputs" / "Evidence_Pack_Direct_Framework_Hypotheses_Questionnaire.docx"

SOURCE_IMAGES = (
    (10, ROOT / "source" / "steenkamp_page_10_model.png", (105, 100, 1175, 1205), ROOT / "outputs" / "Published_Framework_Steenkamp_Figure_1_Direct_Crop.png"),
    (28, ROOT / "source" / "steenkamp_page_28_appendix.png", (110, 115, 1120, 1510), ROOT / "outputs" / "Published_Questionnaire_Steenkamp_Appendix_p28_Direct_Crop.png"),
    (29, ROOT / "source" / "steenkamp_page_29_appendix.png", (110, 115, 1120, 1135), ROOT / "outputs" / "Published_Questionnaire_Steenkamp_Appendix_p29_Direct_Crop.png"),
)


def load_literal(name: str):
    tree = ast.parse(BUILD_SCRIPT.read_text(encoding="utf-8"))
    for node in tree.body:
        if isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name) and target.id == name:
                    return ast.literal_eval(node.value)
    raise RuntimeError(f"Could not find literal {name} in {BUILD_SCRIPT}")


def normalize(text: str) -> str:
    text = text.casefold().replace("’", "'").replace("‘", "'")
    return re.sub(r"[^a-z0-9]+", "", text)


def extract_pdf_text(path: Path) -> str:
    with pdfplumber.open(path) as pdf:
        return "\n".join(
            page.extract_text(x_tolerance=1, y_tolerance=3) or "" for page in pdf.pages
        )


def extract_appendix_item_column(path: Path) -> str:
    with pdfplumber.open(path) as pdf:
        pages = []
        for page_number in (28, 29):
            page = pdf.pages[page_number]
            item_column = page.crop((135, 70, 350, page.height - 20))
            pages.append(item_column.extract_text(x_tolerance=1, y_tolerance=3) or "")
        return "\n".join(pages)


def extract_docx_text(path: Path) -> str:
    doc = Document(path)
    parts = [paragraph.text for paragraph in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                parts.append(cell.text)
    return "\n".join(parts)


def same_pixels(first: Image.Image, second: Image.Image) -> bool:
    return first.size == second.size and ImageChops.difference(
        first.convert("RGB"), second.convert("RGB")
    ).getbbox() is None


def docx_contains_exact_image(docx_path: Path, image_path: Path) -> bool:
    expected = hashlib.sha256(image_path.read_bytes()).digest()
    with zipfile.ZipFile(docx_path) as package:
        for name in package.namelist():
            if name.startswith("word/media/"):
                if hashlib.sha256(package.read(name)).digest() == expected:
                    return True
    return False


def require_contains(haystack: str, needle: str, label: str, errors: list[str]) -> None:
    if normalize(needle) not in haystack:
        errors.append(label)


def main() -> int:
    hypotheses = load_literal("HYPOTHESES")
    constructs = load_literal("CONSTRUCTS")
    items = load_literal("ITEMS")

    errors: list[str] = []
    if len(hypotheses) != 13:
        errors.append(f"Expected 13 hypotheses; found {len(hypotheses)}")
    if len(constructs) != 11:
        errors.append(f"Expected 11 constructs; found {len(constructs)}")
    if len(items) != 40:
        errors.append(f"Expected 40 items; found {len(items)}")

    codes = [item[0] for item in items]
    if len(codes) != len(set(codes)):
        errors.append("Questionnaire item codes are not unique")

    source_text = normalize(extract_pdf_text(SOURCE_PDF))
    source_item_text = normalize(extract_appendix_item_column(SOURCE_PDF))
    master_text = normalize(extract_docx_text(MASTER_DOCX))
    evidence_text = normalize(extract_docx_text(EVIDENCE_DOCX))

    for code, statement in hypotheses:
        require_contains(source_text, f"{code} {statement}", f"Source hypothesis {code}", errors)
        master_statement = statement.replace(
            "Perceptions of external control has", "Perceptions of external control have"
        )
        require_contains(master_text, f"{code} {master_statement}", f"Master hypothesis {code}", errors)
        require_contains(evidence_text, f"{code} {statement}", f"Evidence hypothesis {code}", errors)

    for name, code, _source_location, _item_codes, _count in constructs:
        source_name = "Behavioral Intention" if code == "BI" else name
        master_name = "Behavioural Intention" if code == "BI" else name
        require_contains(source_text, source_name, f"Source construct {code}", errors)
        require_contains(master_text, master_name, f"Master construct {code}", errors)
        require_contains(evidence_text, name, f"Evidence construct {code}", errors)

    for code, _construct, source_wording, _prior_source in items:
        require_contains(source_item_text, f"{code} {source_wording}", f"Source item {code}", errors)
        require_contains(evidence_text, source_wording, f"Evidence item {code}", errors)
        master_wording = source_wording
        if code == "BI3":
            master_wording = source_wording.replace("use the when", "use them when")
        require_contains(master_text, master_wording, f"Master item {code}", errors)

    for label, text in (("master", master_text), ("evidence", evidence_text)):
        for forbidden in ("preparedfordr", "preparedforprofessor", "professorkimi", "drqizhengu"):
            if forbidden in text:
                errors.append(f"Forbidden addressee wording in {label}: {forbidden}")

    for forbidden in (
        "directmodelreplication",
        "contextualreplication",
        "sourceaudit",
        "auditmaster",
        "sourceanomaly",
        "draftstatus",
        "noresultsareavailable",
        "inserthere",
        "correctionrequiresapproval",
    ):
        if forbidden in master_text:
            errors.append(f"Meta/process wording remains in master thesis: {forbidden}")

    source_hash = hashlib.sha256(SOURCE_PDF.read_bytes()).hexdigest()

    source_pdf = pdfium.PdfDocument(SOURCE_PDF)
    for page_number, page_image_path, crop_box, crop_path in SOURCE_IMAGES:
        rendered = source_pdf[page_number].render(scale=2.4).to_pil()
        saved_page = Image.open(page_image_path)
        saved_crop = Image.open(crop_path)
        if not same_pixels(rendered, saved_page):
            errors.append(f"Source page image does not match PDF page index {page_number}")
        if not same_pixels(saved_page.crop(crop_box), saved_crop):
            errors.append(f"Published crop does not match source page image: {crop_path.name}")

    model_crop = SOURCE_IMAGES[0][3]
    if not docx_contains_exact_image(MASTER_DOCX, model_crop):
        errors.append("Master thesis does not contain the exact source framework crop")
    for image_path in (SOURCE_IMAGES[0][3], SOURCE_IMAGES[1][3], SOURCE_IMAGES[2][3]):
        if not docx_contains_exact_image(EVIDENCE_DOCX, image_path):
            errors.append(f"Evidence pack does not contain exact image: {image_path.name}")

    if errors:
        print("DIRECT-SOURCE AUDIT: FAIL")
        for error in errors:
            print(f"- {error}")
        return 1

    print("DIRECT-SOURCE AUDIT: PASS")
    print("- Source framework and Appendix images match direct PDF renders pixel-for-pixel")
    print("- Exact source Figure 1 crop embedded in both documents")
    print("- 13/13 hypotheses found in the source PDF, master thesis, and evidence pack")
    print("- 11/11 constructs retained")
    print("- 40/40 coded source items found in the source PDF and evidence pack")
    print("- 40/40 source items found in the master thesis; BI3 contains only the field-ready typo correction")
    print("- Supervisor-facing manuscript contains none of the prohibited audit/drafting phrases")
    print(f"- Source PDF SHA-256: {source_hash}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
