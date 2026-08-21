from __future__ import annotations

import argparse
from pathlib import Path

from PIL import Image
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

import build_documents as base


PROJECT_ROOT = Path(__file__).resolve().parents[1]
OUTPUTS = PROJECT_ROOT / "outputs"
SOURCE_DIR = PROJECT_ROOT / "source"

MASTER_OUT = OUTPUTS / "Master_Thesis_Digital_Badges_Thai_IT_Students.docx"
EVIDENCE_OUT = OUTPUTS / "Evidence_Pack_Direct_Framework_Hypotheses_Questionnaire.docx"

SOURCE_MODEL_PAGE = SOURCE_DIR / "steenkamp_page_10_model.png"
SOURCE_APPENDIX_28 = SOURCE_DIR / "steenkamp_page_28_appendix.png"
SOURCE_APPENDIX_29 = SOURCE_DIR / "steenkamp_page_29_appendix.png"

MODEL_CROP = OUTPUTS / "Published_Framework_Steenkamp_Figure_1_Direct_Crop.png"
APPENDIX_28_CROP = OUTPUTS / "Published_Questionnaire_Steenkamp_Appendix_p28_Direct_Crop.png"
APPENDIX_29_CROP = OUTPUTS / "Published_Questionnaire_Steenkamp_Appendix_p29_Direct_Crop.png"

TITLE = (
    "Factors Influencing Thai IT Students' Intentions to Use University-Issued "
    "Digital Badges for IT Micro-Credentials in Job Applications"
)
SUBTITLE = "A Quantitative Study Using an Extended Technology Acceptance Model"

BLUE = RGBColor(31, 78, 121)
GRAY = RGBColor(89, 89, 89)


HYPOTHESES = [
    (
        "H1",
        "Perceived usefulness has a positive direct effect on students' intention to use "
        "digital badges in applying for jobs.",
    ),
    (
        "H2",
        "Perceived ease of use has a positive direct effect on students' intention to use "
        "digital badges in applying for jobs.",
    ),
    (
        "H3",
        "Perceived ease of use has a positive direct effect on students' perceived usefulness "
        "of digital badges in applying for jobs.",
    ),
    (
        "H4",
        "Subjective norm has a positive direct effect on students' intention to use digital "
        "badges in applying for jobs.",
    ),
    (
        "H5",
        "Subjective norm has a positive direct effect on students' perceived usefulness of "
        "digital badges in applying for jobs.",
    ),
    ("H6", "Subjective norm has a positive direct effect on image."),
    (
        "H7",
        "Image has a positive direct effect on perceived usefulness of digital badges in "
        "applying for jobs.",
    ),
    (
        "H8",
        "Job application relevance has a positive direct effect on perceived usefulness of "
        "digital badges in applying for jobs.",
    ),
    (
        "H9",
        "Result demonstrability has a positive direct effect on perceived usefulness of digital "
        "badges in applying for jobs.",
    ),
    (
        "H10",
        "Computer self-efficacy has a positive direct effect on perceived ease of use of digital "
        "badges in applying for jobs.",
    ),
    (
        "H11",
        "Perceptions of external control has a positive direct effect on perceived ease of use "
        "of digital badges in applying for jobs.",
    ),
    (
        "H12",
        "Computer anxiety has a negative direct effect on perceived ease of use of digital badges "
        "in applying for jobs.",
    ),
    (
        "H13",
        "Computer playfulness has a positive direct effect on perceived ease of use of digital "
        "badges in applying for jobs.",
    ),
]


CONSTRUCTS = [
    (
        "Perceived Usefulness",
        "PU",
        "Figure 1 and hypothesis development (pp. 10-11); Appendix (p. 28)",
        "PU1-PU4",
        "4",
    ),
    (
        "Perceived Ease of Use",
        "PEOU",
        "Figure 1 and hypothesis development (pp. 10-11); Appendix (p. 28)",
        "PEOU1-PEOU4",
        "4",
    ),
    (
        "Computer Self-Efficacy",
        "CSE",
        "Figure 1 and hypothesis development (pp. 12-13); Appendix (p. 28)",
        "CSE1-CSE3",
        "3",
    ),
    (
        "Perceptions of External Control",
        "PEC",
        "Figure 1 and hypothesis development (pp. 12-13); Appendix (p. 28)",
        "PEC1-PEC4",
        "4",
    ),
    (
        "Computer Playfulness",
        "CPLAY",
        "Figure 1 and hypothesis development (pp. 12-13); Appendix (p. 28)",
        "CPLAY1-CPLAY4",
        "4",
    ),
    (
        "Computer Anxiety",
        "CANX",
        "Figure 1 and hypothesis development (pp. 12-13); Appendix (p. 28)",
        "CANX1-CANX4",
        "4",
    ),
    (
        "Subjective Norm",
        "SN",
        "Figure 1 and hypothesis development (pp. 11-12); Appendix (pp. 28-29)",
        "SN1-SN4",
        "4",
    ),
    (
        "Image",
        "IMG",
        "Figure 1 and hypothesis development (pp. 11-12); Appendix (p. 29)",
        "IMG1-IMG3",
        "3",
    ),
    (
        "Job Application Relevance",
        "REL",
        "Figure 1 and hypothesis development (p. 12); Appendix (p. 29)",
        "REL1-REL3",
        "3",
    ),
    (
        "Result Demonstrability",
        "RES",
        "Figure 1 and hypothesis development (p. 12); Appendix (p. 29)",
        "RES1-RES4",
        "4",
    ),
    (
        "Behavioural Intention to Use DB",
        "BI",
        "Figure 1 and hypothesis development (pp. 10-11); Appendix (p. 29)",
        "BI1-BI3",
        "3",
    ),
]


ITEMS = [
    ("PU1", "Perceived Usefulness", "Using university-issued digital badges would increase my employability when I apply for jobs.", "Davis (1989); Venkatesh & Bala (2008)"),
    ("PU2", "Perceived Usefulness", "Using university-issued digital badges would increase the quality of my job applications.", "Davis (1989); Venkatesh & Bala (2008)"),
    ("PU3", "Perceived Usefulness", "Using university-issued digital badges would enhance my effectiveness when I apply for jobs.", "Davis (1989); Venkatesh & Bala (2008)"),
    ("PU4", "Perceived Usefulness", "I would find using university-issued digital badges useful when applying for jobs.", "Davis (1989); Venkatesh & Bala (2008)"),
    ("PEOU1", "Perceived Ease of Use", "If I were to obtain university-issued digital badges, my interaction with them would be clear and understandable.", "Davis (1989); Venkatesh & Bala (2008)"),
    ("PEOU2", "Perceived Ease of Use", "If I were to obtain university-issued digital badges, interacting with them would not require a lot of my mental effort.", "Davis (1989); Venkatesh & Bala (2008)"),
    ("PEOU3", "Perceived Ease of Use", "If I were to obtain university-issued digital badges, I would find them easy to use.", "Davis (1989); Venkatesh & Bala (2008)"),
    ("PEOU4", "Perceived Ease of Use", "If I were to obtain university-issued digital badges, I would find it easy to get them to do what I want them to do.", "Davis (1989); Venkatesh & Bala (2008)"),
    ("CSE1", "Computer Self-Efficacy", "I could complete a job or task using a computer if there was no one around to tell me what to do.", "Sykes et al. (2014)"),
    ("CSE2", "Computer Self-Efficacy", "I could complete a job or task using a computer if I could contact someone if I got stuck.", "Sykes et al. (2014)"),
    ("CSE3", "Computer Self-Efficacy", "I could complete a job or task using a computer if I had a lot of time to complete the job for which the software was provided.", "Sykes et al. (2014)"),
    ("PEC1", "Perceptions of External Control", "I would have control over my university-issued digital badges.", "Venkatesh & Bala (2008)"),
    ("PEC2", "Perceptions of External Control", "I would have the resources necessary to use university-issued digital badges.", "Venkatesh & Bala (2008)"),
    ("PEC3", "Perceptions of External Control", "Given the resources, opportunities and knowledge it takes to use university-issued digital badges, it would be easy for me to use university-issued digital badges.", "Venkatesh & Bala (2008)"),
    ("PEC4", "Perceptions of External Control", "University-issued digital badges are not compatible with other means I use for communicating my skills and competencies to potential employers.", "Venkatesh & Bala (2008)"),
    ("CPLAY1", "Computer Playfulness", "I would characterise myself as spontaneous when using computers.", "Venkatesh & Bala (2008)"),
    ("CPLAY2", "Computer Playfulness", "I would characterise myself as creative when using computers.", "Venkatesh & Bala (2008)"),
    ("CPLAY3", "Computer Playfulness", "I would characterise myself as playful when using computers.", "Venkatesh & Bala (2008)"),
    ("CPLAY4", "Computer Playfulness", "I would characterise myself as unoriginal when using computers.", "Venkatesh & Bala (2008)"),
    ("CANX1", "Computer Anxiety", "Computers do not scare me at all.", "Venkatesh & Bala (2008)"),
    ("CANX2", "Computer Anxiety", "Working with a computer makes me nervous.", "Venkatesh & Bala (2008)"),
    ("CANX3", "Computer Anxiety", "Computers make me feel uncomfortable.", "Venkatesh & Bala (2008)"),
    ("CANX4", "Computer Anxiety", "Computers make me feel uneasy.", "Venkatesh & Bala (2008)"),
    ("SN1", "Subjective Norm", "People who influence my behaviour think that I should use university-issued digital badges when applying for jobs.", "Venkatesh & Bala (2008); Venkatesh et al. (2003)"),
    ("SN2", "Subjective Norm", "People who are important to me think that I should use university-issued digital badges when applying for jobs.", "Venkatesh & Bala (2008); Venkatesh et al. (2003)"),
    ("SN3", "Subjective Norm", "My university would be helpful in using university-issued digital badges for job applications.", "Venkatesh & Bala (2008); Venkatesh et al. (2003)"),
    ("SN4", "Subjective Norm", "In general, my university would support the use of university-issued digital badges for job applications.", "Venkatesh & Bala (2008); Venkatesh et al. (2003)"),
    ("IMG1", "Image", "Students who use university-issued digital badges when applying for jobs are likely to have more prestige than those who do not.", "Venkatesh & Bala (2008)"),
    ("IMG2", "Image", "Students who use university-issued digital badges when applying for jobs will have a high profile.", "Venkatesh & Bala (2008)"),
    ("IMG3", "Image", "Being able to display university-issued digital badges when applying for jobs would be a status symbol for students.", "Venkatesh & Bala (2008)"),
    ("REL1", "Job Application Relevance", "When job searching, usage of university-issued digital badges is important.", "Venkatesh & Bala (2008)"),
    ("REL2", "Job Application Relevance", "When job searching, usage of university-issued digital badges is relevant.", "Venkatesh & Bala (2008)"),
    ("REL3", "Job Application Relevance", "The use of university-issued digital badges is pertinent to my job searching activities.", "Venkatesh & Bala (2008)"),
    ("RES1", "Result Demonstrability", "I would have no difficulty telling others about the results of using university-issued digital badges.", "Venkatesh & Bala (2008)"),
    ("RES2", "Result Demonstrability", "I believe I could communicate to others the consequences for me of using university-issued digital badges.", "Venkatesh & Bala (2008)"),
    ("RES3", "Result Demonstrability", "The results of using university-issued digital badges would be apparent to me.", "Venkatesh & Bala (2008)"),
    ("RES4", "Result Demonstrability", "I would have difficulty explaining why using university-issued digital badges for job searching may or may not be beneficial.", "Venkatesh & Bala (2008)"),
    ("BI1", "Behavioral Intention", "Assuming I had access to university-issued digital badges I intend to use them when I apply for jobs.", "Venkatesh & Bala (2008); Venkatesh & Davis (2000)"),
    ("BI2", "Behavioral Intention", "Given that I had access to university-issued digital badges, I predict that I would use them when I apply for jobs.", "Venkatesh & Bala (2008); Venkatesh & Davis (2000)"),
    ("BI3", "Behavioral Intention", "If I had access to university-issued digital badges, I would plan to use the when I apply for jobs in the next 12 months.", "Venkatesh & Bala (2008); Venkatesh & Davis (2000)"),
]


REFERENCES = [
    "Brislin, R. W. (1970). Back-translation for cross-cultural research. Journal of Cross-Cultural Psychology, 1(3), 185-216. https://doi.org/10.1177/135910457000100301",
    "Council of the European Union. (2022). Council Recommendation of 16 June 2022 on a European approach to micro-credentials for lifelong learning and employability (2022/C 243/02). Official Journal of the European Union, C 243, 10-25.",
    "Davis, F. D. (1989). Perceived usefulness, perceived ease of use, and user acceptance of information technology. MIS Quarterly, 13(3), 319-340. https://doi.org/10.2307/249008",
    "Hair, J. F., Hult, G. T. M., Ringle, C. M., and Sarstedt, M. (2022). A primer on partial least squares structural equation modeling (PLS-SEM) (3rd ed.). Sage.",
    "Hertzog, M. A. (2008). Considerations in determining sample size for pilot studies. Research in Nursing & Health, 31(2), 180-191. https://doi.org/10.1002/nur.20247",
    "Johanson, G. A., and Brooks, G. P. (2010). Initial scale development: Sample size for pilot studies. Educational and Psychological Measurement, 70(3), 394-400. https://doi.org/10.1177/0013164409355692",
    "Kiiskila, P., Kukkonen, A., and Pirkkalainen, H. (2023). Are micro-credentials valuable for students? Perspective on verifiable digital credentials. SN Computer Science, 4, Article 366. https://doi.org/10.1007/s42979-023-01797-y",
    "Krejcie, R. V., and Morgan, D. W. (1970). Determining sample size for research activities. Educational and Psychological Measurement, 30(3), 607-610. https://doi.org/10.1177/001316447003000308",
    "Miao, M., Ahmed, M., Ahsan, N., and Qamar, B. (2024). Intention to use technology for micro-credential programs: Evidence from technology acceptance and self-determination model. International Journal of Educational Management, 38(4), 948-977. https://doi.org/10.1108/IJEM-02-2023-0066",
    "Steenkamp, N., Fisher, R., and Nesbit, T. (2024). Understanding accounting students' intentions to use digital badges to showcase employability skills. Accounting Education, 33(6), 906-934. https://doi.org/10.1080/09639284.2023.2276200",
    "Sykes, T. A., Venkatesh, V., and Johnson, J. L. (2014). Enterprise system implementation and employee job performance: Understanding the role of advice networks. MIS Quarterly, 38(1), 51-72. https://doi.org/10.25300/MISQ/2014/38.1.03",
    "UNESCO. (2022). Towards a common definition of micro-credentials. UNESCO. https://unesdoc.unesco.org/ark:/48223/pf0000381668",
    "Venkatesh, V. (2000). Determinants of perceived ease of use: Integrating control, intrinsic motivation, and emotion into the technology acceptance model. Information Systems Research, 11(4), 342-365. https://doi.org/10.1287/isre.11.4.342.11872",
    "Venkatesh, V., and Bala, H. (2008). Technology acceptance model 3 and a research agenda on interventions. Decision Sciences, 39(2), 273-315. https://doi.org/10.1111/j.1540-5915.2008.00192.x",
    "Venkatesh, V., and Davis, F. D. (2000). A theoretical extension of the technology acceptance model: Four longitudinal field studies. Management Science, 46(2), 186-204. https://doi.org/10.1287/mnsc.46.2.186.11926",
    "Venkatesh, V., Morris, M. G., Davis, G. B., and Davis, F. D. (2003). User acceptance of information technology: Toward a unified view. MIS Quarterly, 27(3), 425-478. https://doi.org/10.2307/30036540",
]


assert len(ITEMS) == 40


OPERATIONAL_DEFINITIONS = {
    "PU": (
        "The extent to which a student believes that using a university-issued digital badge "
        "will improve the effectiveness and quality of a job application."
    ),
    "PEOU": (
        "The extent to which a student believes that using a university-issued digital badge "
        "will be clear, understandable, and free of effort."
    ),
    "CSE": "A student's belief in their ability to complete computer-based tasks.",
    "PEC": (
        "A student's perception that the resources, knowledge, opportunities, and support needed "
        "to use digital badges are available."
    ),
    "CPLAY": "The degree of cognitive spontaneity a student experiences when interacting with computers.",
    "CANX": "A student's apprehension or uneasiness when faced with using computers.",
    "SN": (
        "The extent to which a student perceives that important people and the university expect "
        "or support the use of digital badges in job applications."
    ),
    "IMG": (
        "The extent to which a student believes that using digital badges will enhance status, "
        "prestige, or profile."
    ),
    "REL": (
        "The extent to which a student believes that digital badges are applicable and pertinent "
        "to job-search and job-application activities."
    ),
    "RES": (
        "The extent to which a student believes that the results of using digital badges are "
        "tangible, observable, and communicable."
    ),
    "BI": (
        "A student's stated intention or plan to use university-issued digital badges when "
        "applying for jobs."
    ),
}


CONSTRUCT_THEORY_SOURCES = {
    "PU": "Davis (1989); Venkatesh and Bala (2008); Steenkamp et al. (2024)",
    "PEOU": "Davis (1989); Venkatesh and Bala (2008); Steenkamp et al. (2024)",
    "CSE": "Venkatesh (2000); Sykes et al. (2014); Steenkamp et al. (2024)",
    "PEC": "Venkatesh (2000); Venkatesh and Bala (2008); Steenkamp et al. (2024)",
    "CPLAY": "Venkatesh (2000); Venkatesh and Bala (2008); Steenkamp et al. (2024)",
    "CANX": "Venkatesh et al. (2003); Venkatesh and Bala (2008); Steenkamp et al. (2024)",
    "SN": "Venkatesh and Davis (2000); Venkatesh and Bala (2008); Steenkamp et al. (2024)",
    "IMG": "Venkatesh and Davis (2000); Venkatesh and Bala (2008); Steenkamp et al. (2024)",
    "REL": "Venkatesh and Bala (2008); Steenkamp et al. (2024)",
    "RES": "Venkatesh and Bala (2008); Steenkamp et al. (2024)",
    "BI": "Venkatesh and Davis (2000); Venkatesh and Bala (2008); Steenkamp et al. (2024)",
}


THESIS_HYPOTHESES = [
    (code, statement.replace("Perceptions of external control has", "Perceptions of external control have"))
    for code, statement in HYPOTHESES
]


FIELD_ITEM_CORRECTIONS = {
    "BI3": (
        "If I had access to university-issued digital badges, I would plan to use them when I "
        "apply for jobs in the next 12 months."
    )
}


def prepare_source_excerpts() -> None:
    OUTPUTS.mkdir(parents=True, exist_ok=True)
    required = [SOURCE_MODEL_PAGE, SOURCE_APPENDIX_28, SOURCE_APPENDIX_29]
    missing = [str(path) for path in required if not path.exists()]
    if missing:
        raise FileNotFoundError("Missing source render(s): " + ", ".join(missing))

    Image.open(SOURCE_MODEL_PAGE).crop((105, 100, 1175, 1205)).save(MODEL_CROP)
    Image.open(SOURCE_APPENDIX_28).crop((110, 115, 1120, 1510)).save(APPENDIX_28_CROP)
    Image.open(SOURCE_APPENDIX_29).crop((110, 115, 1120, 1135)).save(APPENDIX_29_CROP)


def configure_document(doc: Document, running_label: str) -> None:
    base.configure_document(doc)
    section = doc.sections[0]
    header = section.header.paragraphs[0]
    header.paragraph_format.first_line_indent = None
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = header.add_run(running_label)
    run.font.name = "Times New Roman"
    run.font.size = Pt(9)
    run.font.color.rgb = GRAY


def title_page(doc: Document, title: str, subtitle: str, label: str) -> None:
    for _ in range(4):
        doc.add_paragraph()
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = None
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(label.upper())
    r.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = BLUE

    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = None
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(title)
    r.bold = True
    r.font.size = Pt(17)

    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = None
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(subtitle)
    r.italic = True
    r.font.size = Pt(13)
    r.font.color.rgb = GRAY

    for _ in range(5):
        doc.add_paragraph()
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = None
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Min Thiha Oo")
    r.bold = True
    r.font.size = Pt(13)
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = None
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("August 2026").font.size = Pt(12)
    doc.add_page_break()


def add_body(doc: Document, text: str, *, no_indent: bool = False) -> None:
    base.add_body(doc, text, no_indent=no_indent)


def add_bullets(doc: Document, entries: list[str]) -> None:
    for entry in entries:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.first_line_indent = None
        p.paragraph_format.left_indent = Inches(0.4)
        p.paragraph_format.space_after = Pt(2)
        p.add_run(" " + entry)


def add_numbered(doc: Document, entries: list[str]) -> None:
    numbering = doc.part.numbering_part.element
    abstract_ids = [
        int(node.get(qn("w:abstractNumId")))
        for node in numbering.findall(qn("w:abstractNum"))
    ]
    num_ids = [
        int(node.get(qn("w:numId")))
        for node in numbering.findall(qn("w:num"))
    ]
    abstract_id = max(abstract_ids, default=-1) + 1
    num_id = max(num_ids, default=0) + 1

    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(abstract_id))
    multi_level = OxmlElement("w:multiLevelType")
    multi_level.set(qn("w:val"), "singleLevel")
    abstract.append(multi_level)

    level = OxmlElement("w:lvl")
    level.set(qn("w:ilvl"), "0")
    start = OxmlElement("w:start")
    start.set(qn("w:val"), "1")
    level.append(start)
    num_format = OxmlElement("w:numFmt")
    num_format.set(qn("w:val"), "decimal")
    level.append(num_format)
    level_text = OxmlElement("w:lvlText")
    level_text.set(qn("w:val"), "%1.")
    level.append(level_text)
    alignment = OxmlElement("w:lvlJc")
    alignment.set(qn("w:val"), "left")
    level.append(alignment)
    paragraph_properties = OxmlElement("w:pPr")
    indentation = OxmlElement("w:ind")
    indentation.set(qn("w:left"), "576")
    indentation.set(qn("w:hanging"), "288")
    paragraph_properties.append(indentation)
    level.append(paragraph_properties)
    abstract.append(level)
    numbering.append(abstract)

    number = OxmlElement("w:num")
    number.set(qn("w:numId"), str(num_id))
    abstract_ref = OxmlElement("w:abstractNumId")
    abstract_ref.set(qn("w:val"), str(abstract_id))
    number.append(abstract_ref)
    numbering.append(number)

    for entry in entries:
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = None
        p.paragraph_format.space_after = Pt(2)
        p_pr = p._p.get_or_add_pPr()
        num_pr = OxmlElement("w:numPr")
        level_ref = OxmlElement("w:ilvl")
        level_ref.set(qn("w:val"), "0")
        num_ref = OxmlElement("w:numId")
        num_ref.set(qn("w:val"), str(num_id))
        num_pr.extend([level_ref, num_ref])
        p_pr.append(num_pr)
        p.add_run(entry)


def add_note(doc: Document, text: str) -> None:
    base.add_note(doc, text)
    base.prevent_row_split(doc.tables[-1].rows[0])


def fix_table_geometry(table, widths: list[float]) -> None:
    total_dxa = round(sum(widths) * 1440)
    tbl_pr = table._tbl.tblPr

    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(total_dxa))
    tbl_w.set(qn("w:type"), "dxa")

    tbl_ind = tbl_pr.find(qn("w:tblInd"))
    if tbl_ind is None:
        tbl_ind = OxmlElement("w:tblInd")
        tbl_pr.append(tbl_ind)
    tbl_ind.set(qn("w:w"), "120")
    tbl_ind.set(qn("w:type"), "dxa")

    grid = table._tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for width in widths:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(round(width * 1440)))
        grid.append(col)

    table.alignment = WD_TABLE_ALIGNMENT.LEFT


def add_table(
    doc: Document,
    rows: list[list[str]],
    widths: list[float],
    *,
    caption: str,
    font_size: float = 9.0,
) -> None:
    if abs(sum(widths) - 6.2) > 0.01:
        raise ValueError(f"Table widths must total 6.2 inches: {widths}")
    base.add_table(doc, rows, widths, caption=caption, font_size=font_size)
    fix_table_geometry(doc.tables[-1], widths)


def add_figure(
    doc: Document,
    path: Path,
    caption: str,
    *,
    width: float = 5.9,
    alt_text: str | None = None,
) -> None:
    base.add_figure(doc, path, caption, width=width)
    if alt_text:
        doc_pr = doc.inline_shapes[-1]._inline.docPr
        doc_pr.set("title", "Conceptual framework")
        doc_pr.set("descr", alt_text)


def add_source_note(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = None
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_after = Pt(5)
    r = p.add_run(text)
    r.italic = True
    r.font.size = Pt(9.5)
    r.font.color.rgb = GRAY


def add_hypotheses(doc: Document, hypotheses=HYPOTHESES) -> None:
    for code, statement in hypotheses:
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = None
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.hanging_indent = Inches(0.25)
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(f"{code}: ")
        r.bold = True
        p.add_run(statement)


def add_references(doc: Document, *, font_size: float = 10.5, space_after: float = 3) -> None:
    for reference in REFERENCES:
        p = doc.add_paragraph(reference)
        p.paragraph_format.first_line_indent = Inches(-0.3)
        p.paragraph_format.left_indent = Inches(0.3)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.space_after = Pt(space_after)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for run in p.runs:
            run.font.size = Pt(font_size)


def narrative_source_list(source_text: str) -> str:
    sources = source_text.split("; ")
    if len(sources) == 1:
        return sources[0]
    if len(sources) == 2:
        return " and ".join(sources)
    return ", ".join(sources[:-1]) + ", and " + sources[-1]


def questionnaire_item_text(code: str, source_text: str) -> str:
    return FIELD_ITEM_CORRECTIONS.get(code, source_text)


def add_questionnaire(doc: Document, *, include_sources: bool = True) -> None:
    scale_rows = [
        ["1", "2", "3", "4", "5", "6", "7"],
        [
            "Strongly disagree",
            "Moderately disagree",
            "Somewhat disagree",
            "Neutral",
            "Somewhat agree",
            "Moderately agree",
            "Strongly agree",
        ],
    ]
    add_table(
        doc,
        scale_rows,
        [0.8854166667, 0.8854166667, 0.8854166667, 0.8854166667, 0.8854166667, 0.8854166667, 0.8875],
        caption="Table 4.2: Seven-point agreement scale",
        font_size=7.4,
    )

    grouped: dict[str, list[tuple[str, str, str]]] = {}
    for code, construct, source_text, prior_source in ITEMS:
        grouped.setdefault(construct, []).append((code, source_text, prior_source))

    for table_number, (construct, group) in enumerate(grouped.items(), start=3):
        heading = "Behavioural Intention" if construct == "Behavioral Intention" else construct
        doc.add_heading(heading, level=3)
        rows = [["Code", "Questionnaire statement", "Measurement source"]]
        for code, source_text, prior_source in group:
            source = f"Steenkamp et al. (2024); {prior_source}" if include_sources else ""
            rows.append([code, questionnaire_item_text(code, source_text), source])
        add_table(
            doc,
            rows,
            [0.55, 4.1, 1.55],
            caption=f"Table 4.{table_number}: {heading} measurement items",
            font_size=8.0,
        )


def add_toc_page(doc: Document) -> None:
    doc.add_heading("Table of Contents", level=1)
    for label, page in [
        ("CHAPTER 1: INTRODUCTION", 4),
        ("CHAPTER 2: LITERATURE REVIEW", 7),
        ("CHAPTER 3: RESEARCH FRAMEWORK", 11),
        ("CHAPTER 4: RESEARCH METHODOLOGY", 15),
        ("References", 22),
    ]:
        base.add_contents_entry(doc, label, page, bold=True)
    doc.add_page_break()


def build_master() -> None:
    doc = Document()
    configure_document(doc, "Digital badges for IT micro-credentials | Thai IT students")
    title_page(doc, TITLE, SUBTITLE, "Master's Thesis")

    doc.add_heading("Abstract", level=1)
    add_body(
        doc,
        "University-issued digital badges can provide students with a verifiable way to present learning "
        "achievements from IT micro-credentials in job applications. Their practical value, however, "
        "depends partly on whether students intend to use them. This study examines the factors associated "
        "with Thai IT students' intentions to use university-issued digital badges when applying for jobs. "
        "The conceptual framework and questionnaire are adopted from the extended Technology Acceptance "
        "Model developed for digital badges by Steenkamp, Fisher, and Nesbit (2024). The model contains 11 "
        "constructs and 13 hypothesised relationships involving perceived usefulness, perceived ease of use, "
        "social influence, job relevance, facilitating conditions, computer-related beliefs, and behavioural "
        "intention. A quantitative cross-sectional survey will be conducted with students aged 18 years or "
        "older who are enrolled in IT-related university programmes in Thailand. Following a pilot study of "
        "40 eligible students, at least 384 usable responses will be collected for the main study. The "
        "measurement and structural models will be assessed using partial least squares structural equation "
        "modelling. The findings are expected to extend evidence on student adoption of digital credentials "
        "and inform Thai universities considering IT micro-credential badge initiatives.",
        no_indent=True,
    )
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = None
    p.add_run("Keywords: ").bold = True
    p.add_run(
        "digital badges, micro-credentials, technology acceptance model, behavioural intention, "
        "Thai IT students"
    )
    doc.add_page_break()
    add_toc_page(doc)

    doc.add_heading("CHAPTER 1: INTRODUCTION", level=1)
    doc.add_heading("1.1 Background of the Study", level=2)
    add_body(
        doc,
        "Micro-credentials certify learning outcomes acquired through a comparatively small volume of "
        "learning and are intended to support flexible learning and employability (Council of the European "
        "Union, 2022; UNESCO, 2022). A digital badge is a visual and potentially verifiable representation "
        "of an achievement. It can contain information about the issuer, recipient, assessment criteria, "
        "and learning outcome (Kiiskila et al., 2023; Steenkamp et al., 2024). Although the terms are related, "
        "the micro-credential is the record of learning, whereas the badge is one way of representing and "
        "sharing that record.",
    )
    add_body(
        doc,
        "For university students, a badge may provide an additional way to communicate skills in online "
        "profiles and job-application materials. The availability of a badge does not necessarily mean that "
        "students will use it. Adoption may depend on whether students regard the badge as useful and easy to "
        "use, whether important people and institutions support its use, whether it is relevant to job "
        "applications, and whether students have the confidence and resources needed to use it (Steenkamp et "
        "al., 2024).",
    )
    add_body(
        doc,
        "The Technology Acceptance Model (TAM) explains technology use through beliefs about usefulness and "
        "ease of use (Davis, 1989). Later extensions identify social influence, task relevance, perceived "
        "control, computer self-efficacy, computer anxiety, and computer playfulness as additional determinants "
        "of acceptance (Venkatesh, 2000; Venkatesh & Bala, 2008). Steenkamp et al. (2024) applied these "
        "relationships specifically to university-issued digital badges used in job applications. The present "
        "study applies that established model to IT students in Thailand.",
    )

    doc.add_heading("1.2 Statement of the Problem", level=2)
    add_body(
        doc,
        "Universities may invest in IT micro-credentials and digital-badge systems without knowing whether "
        "students will use the badges when seeking employment. Existing research has examined digital-badge "
        "adoption among accounting students in New Zealand (Steenkamp et al., 2024) and technology use for "
        "micro-credential programmes in other national settings (Miao et al., 2024). Evidence remains limited "
        "for students in IT-related programmes in Thailand. Without such evidence, Thai universities have "
        "little empirical guidance on the beliefs and conditions most closely associated with students' "
        "intention to use university-issued badges in job applications. This study addresses that gap by "
        "testing an established digital-badge acceptance model in the Thai IT-student context.",
    )

    doc.add_heading("1.3 Research Objectives", level=2)
    add_numbered(
        doc,
        [
            "To assess Thai IT students' intention to use university-issued digital badges when applying for jobs.",
            "To examine the effects of perceived usefulness, perceived ease of use, and subjective norm on behavioural intention.",
            "To examine the effects of perceived ease of use, subjective norm, image, job application relevance, and result demonstrability on perceived usefulness.",
            "To examine the effects of computer self-efficacy, perceptions of external control, computer anxiety, and computer playfulness on perceived ease of use.",
        ],
    )

    doc.add_heading("1.4 Research Questions", level=2)
    add_numbered(
        doc,
        [
            "To what extent do Thai IT students intend to use university-issued digital badges to showcase employability skills in job applications?",
            "How do perceived usefulness, perceived ease of use, and subjective norm influence behavioural intention?",
            "How do perceived ease of use, subjective norm, image, job application relevance, and result demonstrability influence perceived usefulness?",
            "How do computer self-efficacy, perceptions of external control, computer anxiety, and computer playfulness influence perceived ease of use?",
        ],
    )

    doc.add_heading("1.5 Scope of the Research", level=2)
    add_bullets(
        doc,
        [
            "Population: university students aged 18 or above who are enrolled in an IT-related program in Thailand.",
            "Focal object: a university-issued digital badge representing completion of an IT micro-credential and usable in job applications.",
            "Variables: 11 constructs from the extended Technology Acceptance Model applied by Steenkamp et al. (2024).",
            "Instrument: 40 construct items measured on a seven-point Likert agreement scale.",
            "Method: a pilot study followed by a cross-sectional online survey and PLS-SEM analysis.",
        ],
    )

    doc.add_heading("1.6 Research Limitations", level=2)
    add_body(
        doc,
        "The cross-sectional, self-report design can estimate associations and predictive paths but cannot "
        "prove actual badge use, employer recognition, employment, or causal effects. Non-probability "
        "recruitment may limit representativeness. Students may also have limited prior experience with "
        "digital badges; therefore a neutral definition will precede the construct items. Translation can "
        "introduce semantic differences, which will be managed through forward translation, independent "
        "back-translation, and the required pilot.",
    )

    doc.add_heading("1.7 Significance of the Study", level=2)
    add_body(
        doc,
        "The study extends the application of TAM3-related constructs to digital credentials in a Thai "
        "higher-education setting. It also provides practical evidence about the factors associated with "
        "students' willingness to use badges issued for IT micro-credentials. The findings may assist "
        "universities in planning badge communication, student support, training, and integration with "
        "career-development activities. The outcome measured is behavioural intention rather than actual "
        "employment or employer evaluation.",
    )

    doc.add_heading("1.8 Definition of Key Terms", level=2)
    definition_rows = [["Term", "Working definition"]]
    definition_rows += [
        ["Micro-credential", "A record of learning outcomes acquired through a small volume of learning (Council of the European Union, 2022; UNESCO, 2022)."],
        ["University-issued digital badge", "A verifiable visual token issued by a university to represent an achievement; in this study, completion of an IT micro-credential."],
        ["Perceived usefulness", OPERATIONAL_DEFINITIONS["PU"] + " (Davis, 1989)."],
        ["Perceived ease of use", OPERATIONAL_DEFINITIONS["PEOU"] + " (Davis, 1989)."],
        ["Behavioural intention", OPERATIONAL_DEFINITIONS["BI"] + " (Venkatesh & Davis, 2000)."],
    ]
    add_table(doc, definition_rows, [1.65, 4.55], caption="Table 1.1: Working definitions", font_size=9.4)

    doc.add_page_break()
    doc.add_heading("CHAPTER 2: LITERATURE REVIEW", level=1)
    doc.add_heading("2.1 Applied Theory: Technology Acceptance Model and TAM3", level=2)
    add_body(
        doc,
        "The Technology Acceptance Model proposes that perceived usefulness and perceived ease of use are "
        "central beliefs underlying a person's intention to use a technology (Davis, 1989). Perceived ease of "
        "use may also influence perceived usefulness because a technology that requires less effort can be "
        "more useful in practice (Venkatesh & Davis, 2000). Behavioural intention is consequently treated as "
        "the immediate outcome of the acceptance process.",
    )
    add_body(
        doc,
        "TAM3 extends the original model by explaining how social influence, task characteristics, control "
        "beliefs, intrinsic motivation, and anxiety contribute to usefulness and ease-of-use perceptions "
        "(Venkatesh, 2000; Venkatesh & Bala, 2008). Steenkamp et al. (2024) applied this extended structure to "
        "students' use of university-issued digital badges in job applications. Their model provides the "
        "theoretical framework for the present study because it addresses the same technology, decision, and "
        "intended-use outcome.",
    )

    doc.add_heading("2.2 Digital Badges and IT Micro-Credentials", level=2)
    add_body(
        doc,
        "Micro-credentials are short, assessed learning experiences that document specific learning outcomes "
        "(Council of the European Union, 2022; UNESCO, 2022). Digital badges can represent such achievements "
        "in a portable and verifiable format. Student perceptions of verifiable digital credentials include "
        "their usefulness for displaying competence, their credibility, and the ease with which they can be "
        "shared (Kiiskila et al., 2023). Research on micro-credential programmes also indicates that technology "
        "acceptance provides a suitable basis for studying student intention (Miao et al., 2024). In this study, "
        "the focal technology is a university-issued badge representing completion of an IT micro-credential."
    )

    doc.add_heading("2.3 Variable Definitions", level=2)
    for index, (name, code, _, _, _) in enumerate(CONSTRUCTS, start=1):
        display_name = "Behavioural Intention" if code == "BI" else name
        doc.add_heading(f"2.3.{index} {display_name}", level=3)
        add_body(
            doc,
            f"{display_name} is defined in this study as {OPERATIONAL_DEFINITIONS[code][0].lower() + OPERATIONAL_DEFINITIONS[code][1:]} "
            f"This definition follows {narrative_source_list(CONSTRUCT_THEORY_SOURCES[code])}.",
        )

    doc.add_heading("2.4 Relationships Between Variables", level=2)
    doc.add_heading("2.4.1 Determinants of Behavioural Intention", level=3)
    add_body(
        doc,
        "TAM proposes that perceived usefulness and perceived ease of use influence behavioural intention "
        "(Davis, 1989). Subjective norm can also influence intention when important social actors encourage "
        "technology use (Venkatesh & Davis, 2000). In the digital-badge study by Steenkamp et al. (2024), "
        "perceived usefulness and subjective norm had significant positive effects on behavioural intention, "
        "whereas the direct effect of perceived ease of use was not significant. All three theoretically "
        "specified paths are examined in the Thai context.",
    )

    doc.add_heading("2.4.2 Determinants of Perceived Usefulness", level=3)
    add_body(
        doc,
        "Perceived ease of use can increase perceived usefulness when lower effort improves the practical "
        "value of a system (Venkatesh & Davis, 2000). TAM3 further proposes that subjective norm and image "
        "represent social-influence processes, while job relevance and result demonstrability represent "
        "cognitive instrumental processes (Venkatesh & Bala, 2008). Steenkamp et al. (2024) found significant "
        "positive effects for perceived ease of use and job application relevance on perceived usefulness, "
        "and a significant effect of subjective norm on image. The remaining proposed relationships were not "
        "significant in their sample and are retested among Thai IT students.",
    )

    doc.add_heading("2.4.3 Determinants of Perceived Ease of Use", level=3)
    add_body(
        doc,
        "Early perceptions of ease of use may be anchored in computer self-efficacy, perceptions of external "
        "control, computer anxiety, and computer playfulness (Venkatesh, 2000; Venkatesh & Bala, 2008). "
        "Self-efficacy, external control, and playfulness are expected to increase perceived ease of use, "
        "whereas anxiety is expected to reduce it. Steenkamp et al. (2024) found a significant positive effect "
        "only for perceptions of external control. Testing all four relationships permits assessment of whether "
        "the pattern differs among Thai IT students.",
    )

    doc.add_heading("2.5 Previous Studies", level=2)
    prior_rows = [
        ["Study", "Context and method", "Main finding", "Relevance"],
        ["Davis (1989)", "Technology acceptance; scale development and validation.", "Perceived usefulness and perceived ease of use are central determinants of user acceptance.", "Establishes the core TAM constructs."],
        ["Venkatesh and Bala (2008)", "Longitudinal research integrating determinants of usefulness and ease of use.", "TAM3 explains acceptance through social influence, cognitive processes, control, motivation, and emotion.", "Provides the extended theoretical structure and many measurement scales."],
        ["Kiiskila et al. (2023)", "Higher-education students; qualitative study of verifiable digital credentials.", "Students identified value in verifiable and shareable credentials while also raising implementation concerns.", "Supports the digital-credential and student context."],
        ["Miao et al. (2024)", "University students; quantitative PLS-SEM study of technology use for micro-credential programmes.", "A TAM-based model was used to explain intention in a micro-credential setting.", "Supports the use of technology-acceptance theory for micro-credentials."],
        ["Steenkamp et al. (2024)", "New Zealand accounting students; online survey; extended TAM3; PLS-SEM; n = 57.", "Six of 13 hypothesised relationships were supported, including usefulness and subjective norm as predictors of intention.", "Provides the conceptual framework, hypotheses, and questionnaire used in this study."],
    ]
    add_table(doc, prior_rows, [1.2, 1.75, 1.75, 1.5], caption="Table 2.1: Summary of previous studies", font_size=7.8)

    doc.add_page_break()
    doc.add_heading("CHAPTER 3: RESEARCH FRAMEWORK", level=1)
    doc.add_heading("3.1 Theoretical Framework", level=2)
    add_body(
        doc,
        "This study is grounded in TAM and TAM3. The core TAM relationships connect perceived usefulness and "
        "perceived ease of use with behavioural intention (Davis, 1989). TAM3 explains these beliefs through "
        "social influence, cognitive instrumental processes, control beliefs, intrinsic motivation, and "
        "computer anxiety (Venkatesh, 2000; Venkatesh & Bala, 2008). The combined framework is appropriate "
        "because the decision under investigation is whether students intend to use a digital technology for "
        "a specific task: presenting university-issued badges in job applications.",
    )
    add_body(
        doc,
        "Within this framework, perceived usefulness and perceived ease of use are the principal acceptance "
        "beliefs. Subjective norm and image represent social-influence processes, while job application "
        "relevance and result demonstrability represent cognitive judgements about the value of the badge for "
        "the task. Computer self-efficacy, external control, anxiety, and playfulness explain students' initial "
        "ease-of-use perceptions (Venkatesh, 2000; Venkatesh & Bala, 2008).",
    )
    add_body(
        doc,
        "Behavioural intention is used as the dependent variable because the proposed badge is not assumed to "
        "be available to every respondent at the time of data collection. The framework therefore explains "
        "students' stated intention to use a badge, rather than actual badge use or employment outcomes."
    )

    doc.add_heading("3.2 Conceptual Framework", level=2)
    add_body(
        doc,
        "The conceptual framework is adopted from the digital-badge acceptance model developed and tested by "
        "Steenkamp et al. (2024). It includes the same 11 constructs and 13 directional relationships. The "
        "population examined in the present study is Thai students enrolled in IT-related university "
        "programmes, and the badge represents completion of an IT micro-credential.",
    )
    add_figure(
        doc,
        MODEL_CROP,
        "Figure 3.1: Conceptual framework for students' intention to use digital badges",
        width=5.75,
        alt_text=(
            "Conceptual framework adopted from Steenkamp et al. (2024), showing 11 constructs and "
            "13 directional relationships predicting perceived usefulness, perceived ease of use, "
            "image, and behavioural intention to use digital badges."
        ),
    )
    add_source_note(
        doc,
        "Source: Steenkamp, Fisher, and Nesbit (2024, Figure 1, p. 10).",
    )
    add_body(
        doc,
        "Perceived usefulness, perceived ease of use, and subjective norm are positioned as predictors of "
        "behavioural intention. Perceived ease of use, subjective norm, image, job application relevance, and "
        "result demonstrability are positioned as predictors of perceived usefulness. Computer self-efficacy, "
        "perceptions of external control, computer anxiety, and computer playfulness are positioned as "
        "predictors of perceived ease of use. Subjective norm is also linked to image. In Figure 3.1, DB "
        "denotes digital badges.",
    )

    doc.add_heading("3.3 Research Hypotheses", level=2)
    add_body(
        doc,
        "Based on the theoretical relationships described above and the digital-badge model of Steenkamp et "
        "al. (2024, pp. 11-13), the following hypotheses are proposed:",
    )
    add_hypotheses(doc, THESIS_HYPOTHESES)

    heading = doc.add_heading("3.4 Operationalization", level=2)
    heading.paragraph_format.page_break_before = True
    op_rows = [["Variable", "Operational definition", "Items", "Scale", "Source/reference"]]
    for name, code, _, codes, _ in CONSTRUCTS:
        display_name = "Behavioural Intention" if code == "BI" else name
        op_rows.append(
            [
                display_name,
                OPERATIONAL_DEFINITIONS[code],
                codes,
                "7-point Likert",
                CONSTRUCT_THEORY_SOURCES[code],
            ]
        )
    add_table(
        doc,
        op_rows,
        [1.15, 2.35, 0.75, 0.75, 1.2],
        caption="Table 3.1: Operationalization of the study variables",
        font_size=7.4,
    )
    add_source_note(
        doc,
        "The complete questionnaire statements and their measurement sources are presented in Section 4.3.",
    )

    doc.add_page_break()
    doc.add_heading("CHAPTER 4: RESEARCH METHODOLOGY", level=1)
    doc.add_heading("4.1 Research Design", level=2)
    add_body(
        doc,
        "The study will use a quantitative, explanatory, cross-sectional survey design. This design is "
        "appropriate for measuring students' perceptions and testing the hypothesised relationships among "
        "the 11 latent variables at one point in time. The construct indicators will be specified as "
        "reflective, consistent with Steenkamp et al. (2024), and the model will be estimated using PLS-SEM.",
    )

    doc.add_heading("4.2 Sampling Procedure", level=2)
    doc.add_heading("4.2.1 Target Population", level=3)
    add_body(
        doc,
        "The target population consists of students aged 18 years or older who are currently enrolled in an "
        "IT-related undergraduate or postgraduate programme at a university in Thailand. Relevant fields may "
        "include information technology, computer science, software engineering, information systems, data "
        "science, cybersecurity, and closely related programmes.",
    )
    doc.add_heading("4.2.2 Sampling Method and Sample Size", level=3)
    add_body(
        doc,
        "Non-probability convenience and snowball sampling will be used because a complete national sampling "
        "frame of Thai IT students is not available. The online questionnaire will be distributed through "
        "IT-related university programmes, student groups, and academic networks. The Krejcie and Morgan "
        "(1970) table identifies 384 as a conventional sample-size benchmark for a large population under "
        "probability-sampling assumptions. Because the present study uses non-probability sampling, this "
        "benchmark does not imply a 5% margin of error for the achieved sample. It is used as a conservative "
        "target that also exceeds the minimum normally required for a PLS-SEM model of this complexity (Hair "
        "et al., 2022). The main study will therefore require at least 384 usable responses, with a recruitment "
        "target of 400 to allow for incomplete or ineligible responses.",
    )
    doc.add_heading("4.2.3 Inclusion and Exclusion Criteria", level=3)
    add_body(
        doc,
        "Respondents will be included if they provide informed consent, are at least 18 years old, are "
        "currently enrolled in an IT-related university programme in Thailand, and complete the construct "
        "measures. Responses will be excluded for failed eligibility screening, duplicate submission, "
        "substantial missing data, or implausibly short completion time. Pilot participants will not be "
        "included in the main-study dataset.",
    )

    doc.add_heading("4.3 Research Instrument and Questionnaire Design", level=2)
    add_body(
        doc,
        "Data will be collected using a structured online questionnaire. The conceptual measures and 40 "
        "construct items are adopted from Steenkamp et al. (2024, Appendix, pp. 28-29), who drew the measures "
        "from established TAM, TAM3, UTAUT, and computer self-efficacy scales. The focal referent throughout "
        "the construct section is a university-issued digital badge representing completion of an IT "
        "micro-credential and used in a job application.",
    )
    add_body(
        doc,
        "BI3 contains a grammatical correction from 'use the when' in the published Appendix to 'use them "
        "when'; the construct meaning, referent, and 12-month timeframe are unchanged.",
    )
    instrument_rows = [
        ["Section", "Content", "Purpose"],
        ["A", "Participant information, consent, age, enrolment status, and field of study", "Confirm consent and eligibility"],
        ["B", "Neutral definition of an IT micro-credential and a university-issued digital badge", "Provide a common referent"],
        ["C", "Age group, gender, study level, field, year of study, institution type, and prior badge awareness or use", "Describe the sample"],
        ["D", "Forty statements measuring the 11 model constructs", "Test the measurement and structural models"],
    ]
    add_table(doc, instrument_rows, [0.65, 3.65, 1.9], caption="Table 4.1: Structure of the questionnaire", font_size=8.6)
    add_body(
        doc,
        "Construct items will use a seven-point Likert agreement scale ranging from 1 (strongly disagree) to "
        "7 (strongly agree). The full English questionnaire statements and their published measurement sources "
        "are presented below.",
    )
    add_questionnaire(doc)

    doc.add_heading("4.4 Translation and Pilot Study", level=2)
    add_body(
        doc,
        "The questionnaire will be translated from English into Thai and independently back-translated into "
        "English. Differences will be reconciled to preserve the intended meaning of each construct (Brislin, "
        "1970). Before the main survey, the Thai questionnaire will be pilot tested with 40 eligible Thai IT "
        "students. A sample of this size is appropriate for an initial assessment of item comprehension and "
        "scale performance (Hertzog, 2008; Johanson & Brooks, 2010). The pilot will assess eligibility logic, "
        "clarity of instructions and items, completion time, missing responses, and preliminary internal "
        "consistency. Pilot participants will be excluded from the main survey.",
    )

    doc.add_heading("4.5 Data Collection", level=2)
    add_body(
        doc,
        "Following academic and ethical approval, an online survey link will be distributed through the "
        "identified university and student channels. The first page will provide the participant information "
        "and consent statement. Participation will be voluntary, and respondents may leave the survey before "
        "submission. No directly identifying information will be required. Responses will be stored securely "
        "and used only for the purposes described in the participant information.",
    )

    doc.add_heading("4.6 Statistical Treatment of Data", level=2)
    add_numbered(
        doc,
        [
            "Screen records for consent, eligibility, duplicates, missingness, completion quality, and coding errors; reverse-code negatively phrased indicators before construct analysis.",
            "Describe respondent characteristics using frequencies and percentages, and summarize construct responses using means and standard deviations.",
            "Assess the reflective measurement model using indicator loadings, Cronbach's alpha, composite reliability, average variance extracted, and discriminant validity. All 40 indicators will be administered; any indicator removal will be based on the Thai-sample results and theoretical content and will be reported.",
            "Assess structural-model collinearity using variance inflation factors and evaluate the explanatory power of endogenous constructs using R-squared values.",
            "Estimate the 13 hypothesised paths in SmartPLS using 5,000 bootstrap samples and report path coefficients, confidence intervals, p-values, and effect sizes (Hair et al., 2022).",
            "Determine support for each hypothesis at the 5% significance level and compare the resulting pattern with the findings of Steenkamp et al. (2024).",
        ],
    )

    doc.add_heading("4.7 Ethical Considerations", level=2)
    add_body(
        doc,
        "The study will follow the university's research-ethics requirements. Participants will receive "
        "information about the study purpose, eligibility requirements, voluntary participation, "
        "confidentiality, data use, and withdrawal before submission. Only respondents who provide informed "
        "consent will proceed. Results will be reported in aggregate form.",
    )

    doc.add_page_break()
    doc.add_heading("References", level=1)
    add_references(doc)

    doc.save(MASTER_OUT)


def build_evidence_pack() -> None:
    doc = Document()
    configure_document(doc, "Evidence pack | direct source audit")
    title_page(
        doc,
        "Published Framework, Hypotheses, Variables, and Questionnaire Evidence",
        "Direct-model replication source pack",
        "Evidence Pack",
    )

    doc.add_heading("1. Decision", level=1)
    add_note(
        doc,
        "Pass under the strict source-fidelity rule. One published article supplies the actual model, "
        "all 13 hypotheses, all 11 constructs, the 40 coded construct items, the seven-point scale, and "
        "the PLS-SEM procedure. The Thai population and IT micro-credential badge context are disclosed "
        "replication changes.",
    )
    add_body(
        doc,
        "Primary source: Steenkamp, N., Fisher, R., and Nesbit, T. (2024), Understanding "
        "accounting students' intentions to use digital badges to showcase employability skills, "
        "Accounting Education, 33(6), 906-934. DOI: 10.1080/09639284.2023.2276200.",
    )

    doc.add_heading("2. Strict Source Audit", level=1)
    audit_rows = [
        ["Required component", "Published location", "Status", "Current use"],
        ["Actual framework", "Figure 1, article p. 10", "Direct", "Complete figure reproduced; no redrawing or path selection."],
        ["Variables", "Figure 1, Appendix, Table 2", "Direct", "All 11 source constructs and codes retained."],
        ["Hypotheses", "H1-H13, article pp. 11-13", "Direct", "All paths and directions reproduced."],
        ["Questionnaire", "Appendix, article pp. 28-29", "Direct", "All 40 coded source items transcribed."],
        ["Scale", "Appendix note, article p. 29", "Direct", "Seven-point agreement scale retained."],
        ["Analysis", "Method/results, article pp. 13-17", "Direct-method guide", "Reflective PLS-SEM and 5,000 bootstrap samples."],
    ]
    add_table(doc, audit_rows, [1.25, 1.55, 0.85, 2.55], caption="Table 2.1: Pass/fail audit", font_size=8.3)

    doc.add_page_break()
    doc.add_heading("3. Actual Published Framework", level=1)
    add_body(
        doc,
        "The image below is a direct crop from the source PDF. It is not a reconstruction. The source "
        "boxes, labels, arrows, and caption are unchanged.",
    )
    add_figure(
        doc,
        MODEL_CROP,
        "Figure 3.1: Direct source reproduction of Steenkamp et al. (2024, Figure 1, article p. 10).",
        width=5.8,
    )

    doc.add_heading("4. Exact Published Hypotheses", level=1)
    add_body(
        doc,
        "The statements below reproduce the source hypotheses. H11's grammar is retained to avoid "
        "silently rewriting the published statement.",
    )
    add_hypotheses(doc)

    doc.add_heading("5. Exact Variable and Item-Count Match", level=1)
    count_rows = [["Published construct", "Code", "Published codes", "Count", "Thesis match"]]
    count_rows += [[name, code, codes, count, "Exact"] for name, code, _, codes, count in CONSTRUCTS]
    count_rows.append(["Total", "", "", "40", "Exact"])
    add_table(doc, count_rows, [1.8, 0.55, 1.55, 0.55, 1.75], caption="Table 5.1: Construct and item-count audit", font_size=8.4)

    doc.add_heading("6. Actual Published Questionnaire Appendix", level=1)
    add_body(
        doc,
        "The next two images are direct crops of the source Appendix pages. They show the item codes, "
        "wording, prior sources identified by Steenkamp et al., and seven-point response scale.",
    )
    add_figure(
        doc,
        APPENDIX_28_CROP,
        "Figure 6.1: Direct source reproduction of the first construct-item Appendix page, Steenkamp et al. (2024, article p. 28).",
        width=5.35,
    )
    doc.add_page_break()
    add_figure(
        doc,
        APPENDIX_29_CROP,
        "Figure 6.2: Direct source reproduction of the second construct-item Appendix page, Steenkamp et al. (2024, article p. 29).",
        width=6.0,
    )

    doc.add_page_break()
    doc.add_heading("7. Item-Level Provenance", level=1)
    add_body(
        doc,
        "All 40 coded statements below are verbatim from the source Appendix, including BI3's apparent "
        "typo; any correction requires approval before fielding.",
    )
    provenance_rows = [["Code", "Construct", "Exact source wording", "Master status"]]
    for code, construct, source_text, _ in ITEMS:
        status = "Exact"
        if code == "BI3":
            status = "Exact; typo noted"
        provenance_rows.append([code, construct, source_text, status])
    add_table(doc, provenance_rows, [0.55, 1.35, 3.55, 0.75], caption="Table 7.1: Forty-item direct provenance matrix", font_size=7.6)

    doc.add_heading("8. Source Anomalies and Decisions", level=1)
    anomaly_rows = [
        ["Source detail", "Evidence", "Decision"],
        ["Unnumbered employer sentence after SN4", "The methods state 40 items; the Appendix has 40 coded items; Table 2 reports SN1-SN4 only.", "Do not invent SN5 or score the unnumbered sentence."],
        ["PEC4, CPLAY4, RES4 removed in source analysis", "Source measurement-model section reports outer loadings below .40.", "Administer all 40 source items; make any Thai-sample deletion from Thai data and report it."],
        ["BI3 source typo", "Appendix prints 'use the when'.", "Preserve source wording in evidence; seek approval for the field correction 'them'."],
        ["Behavioural/Behavioral label variation", "Figure and Appendix use different English spelling/label length.", "Retain the figure label in the model and the Appendix label in the item evidence; treat both as BI."],
    ]
    add_table(doc, anomaly_rows, [1.35, 2.55, 2.3], caption="Table 8.1: No-hidden-decisions register", font_size=8.4)

    doc.add_heading("9. What Changes and What Does Not", level=1)
    change_rows = [
        ["Element", "Source study", "Thai study", "Classification"],
        ["Population", "Accounting students at one New Zealand university", "Students in IT-related programs in Thailand", "Disclosed context change"],
        ["Badge context", "University-issued badges for employability skills", "University-issued badges representing IT micro-credentials for employability skills", "Disclosed context specification"],
        ["Model, constructs, paths", "Figure 1; 11 constructs; 13 paths", "Same complete model", "No change"],
        ["Construct items", "40 coded Appendix items", "Same English source master", "No conceptual change"],
        ["Response scale", "Seven-point agreement", "Same", "No change"],
        ["Language", "English", "English audit master plus approved Thai translation", "Necessary cross-language procedure"],
        ["Pilot", "Not reported as a separate pilot", "Required local pilot of 40", "Local process requirement"],
    ]
    add_table(doc, change_rows, [1.15, 1.85, 2.05, 1.15], caption="Table 9.1: Replication boundary", font_size=8.2)

    doc.add_heading("10. References", level=1)
    add_references(doc, font_size=9, space_after=1)
    doc.save(EVIDENCE_OUT)


def main() -> None:
    parser = argparse.ArgumentParser(description="Build the current thesis documents.")
    parser.add_argument(
        "--master-only",
        action="store_true",
        help="Build only the supervisor-facing thesis manuscript.",
    )
    args = parser.parse_args()

    prepare_source_excerpts()
    build_master()
    print(MASTER_OUT)
    if not args.master_only:
        build_evidence_pack()
        print(EVIDENCE_OUT)
    print(MODEL_CROP)
    print(APPENDIX_28_CROP)
    print(APPENDIX_29_CROP)


if __name__ == "__main__":
    main()
