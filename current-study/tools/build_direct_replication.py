from __future__ import annotations

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

MASTER_OUT = OUTPUTS / "Master_Thesis_Direct_Replication_Digital_Badges_Thai_IT_Students.docx"
EVIDENCE_OUT = OUTPUTS / "Evidence_Pack_Direct_Framework_Hypotheses_Questionnaire.docx"

SOURCE_MODEL_PAGE = SOURCE_DIR / "steenkamp_page_10_model.png"
SOURCE_APPENDIX_28 = SOURCE_DIR / "steenkamp_page_28_appendix.png"
SOURCE_APPENDIX_29 = SOURCE_DIR / "steenkamp_page_29_appendix.png"

MODEL_CROP = OUTPUTS / "Published_Framework_Steenkamp_Figure_1_Direct_Crop.png"
APPENDIX_28_CROP = OUTPUTS / "Published_Questionnaire_Steenkamp_Appendix_p28_Direct_Crop.png"
APPENDIX_29_CROP = OUTPUTS / "Published_Questionnaire_Steenkamp_Appendix_p29_Direct_Crop.png"

TITLE = (
    "Understanding Thai IT Students' Intentions to Use University-Issued "
    "Digital Badges to Showcase Employability Skills"
)
SUBTITLE = "A Direct-Model Replication of Steenkamp, Fisher, and Nesbit (2024)"

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
        p.add_run(entry)


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


def add_figure(doc: Document, path: Path, caption: str, *, width: float = 5.9) -> None:
    base.add_figure(doc, path, caption, width=width)


def add_source_note(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = None
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_after = Pt(5)
    r = p.add_run(text)
    r.italic = True
    r.font.size = Pt(9.5)
    r.font.color.rgb = GRAY


def add_hypotheses(doc: Document) -> None:
    for code, statement in HYPOTHESES:
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


def source_master_item(source_text: str) -> str:
    return source_text


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
    add_table(doc, scale_rows, [0.885, 0.885, 0.885, 0.885, 0.885, 0.885, 0.89], caption="Seven-point response scale reproduced from Steenkamp et al. (2024)", font_size=7.4)

    grouped: dict[str, list[tuple[str, str, str]]] = {}
    for code, construct, source_text, prior_source in ITEMS:
        grouped.setdefault(construct, []).append((code, source_text, prior_source))

    for construct, group in grouped.items():
        doc.add_heading(construct, level=2)
        if include_sources:
            add_source_note(
                doc,
                "Direct item source: Steenkamp et al. (2024), Appendix, pp. 28-29. "
                f"The source article identifies: {group[0][2]}.",
            )
        for code, source_text, _ in group:
            wording = source_master_item(source_text)
            p = doc.add_paragraph()
            p.paragraph_format.first_line_indent = None
            p.paragraph_format.left_indent = Inches(0.25)
            p.paragraph_format.hanging_indent = Inches(0.25)
            p.paragraph_format.space_after = Pt(5)
            r = p.add_run(f"{code}. ")
            r.bold = True
            p.add_run(wording)
            if code == "BI3":
                r = p.add_run(" [Source wording retained verbatim; correction requires approval.]")
                r.italic = True
                r.font.size = Pt(9)


def add_toc_page(doc: Document) -> None:
    doc.add_heading("Table of Contents", level=1)
    for label, page in [
        ("CHAPTER 1: INTRODUCTION", 4),
        ("CHAPTER 2: LITERATURE REVIEW", 7),
        ("CHAPTER 3: RESEARCH FRAMEWORK", 9),
        ("CHAPTER 4: RESEARCH METHODOLOGY", 12),
        ("CHAPTER 5: RESEARCH RESULTS", 14),
        ("CHAPTER 6: DISCUSSION AND CONCLUSION", 15),
        ("References", 16),
        ("Appendix A: English Source Questionnaire for Pilot Review", 17),
        ("Appendix B: Source Measurement Decisions", 21),
    ]:
        base.add_contents_entry(doc, label, page, bold=True)
    doc.add_page_break()


def build_master() -> None:
    doc = Document()
    configure_document(doc, "Direct-model replication | Thai IT students")
    title_page(doc, TITLE, SUBTITLE, "Master's Thesis Working Draft")

    doc.add_heading("Draft Status", level=1)
    add_note(
        doc,
        "Proposal-stage draft. Chapters 1-4 and the English source questionnaire are prepared for "
        "review. Chapters 5-6 contain reporting structures only. No approval, pilot result, main-study "
        "result, or employment outcome is claimed.",
    )
    doc.add_heading("Abstract", level=1)
    add_body(
        doc,
        "This study proposes a contextual replication of Steenkamp, Fisher, and Nesbit's published "
        "model of university students' intentions to use digital badges in job applications. The "
        "population is Thai university students in IT-related programs, and the focal badge is the "
        "university-issued digital representation of an IT micro-credential. The complete source model, "
        "its 13 hypotheses, 11 constructs, 40 coded questionnaire items, and seven-point response scale "
        "are retained. After a required pilot, the main cross-sectional survey will be analysed using "
        "reflective PLS-SEM. The study is designed to test whether the source model transfers to a new "
        "disciplinary and national setting; it will not establish actual employability or employer behaviour.",
    )
    doc.add_page_break()
    add_toc_page(doc)

    doc.add_heading("CHAPTER 1: INTRODUCTION", level=1)
    doc.add_heading("1.1 Background of the Study", level=2)
    add_body(
        doc,
        "Micro-credentials record learning outcomes gained through a small volume of learning, while "
        "digital badges are visual and verifiable tokens that can represent learning and achievement "
        "(Council of the European Union, 2022; UNESCO, 2022; Steenkamp et al., 2024). The two terms are "
        "related but not interchangeable: a digital badge can be used as the portable representation of "
        "a micro-credential and can contain information about the issuer, recipient, criteria, and "
        "achievement. This distinction allows the present study to remain in the IT micro-credential "
        "context while measuring students' acceptance of the university-issued badge used to communicate "
        "that credential.",
    )
    add_body(
        doc,
        "For students, the potential benefit is practical. A verifiable badge can be displayed through "
        "online profiles or job-application materials so that skills not fully visible on a traditional "
        "transcript can be communicated to potential employers. That benefit depends on student adoption. "
        "If students do not regard badges as useful, relevant, manageable, or socially supported, the "
        "credential's signalling function may not be used even when the badge is technically available "
        "(Steenkamp et al., 2024).",
    )
    add_body(
        doc,
        "Steenkamp et al. (2024) developed and tested an extended, context-specific Technology "
        "Acceptance Model for university-issued digital badges. Their article is unusually suitable as a "
        "replication source because it publishes the complete research model, every directional hypothesis, "
        "the construct-item Appendix, and the analysis procedure. The present study applies that model to "
        "Thai students in IT-related programs rather than assembling a new framework from partly matching "
        "sources.",
    )

    doc.add_heading("1.2 Statement of the Problem", level=2)
    add_body(
        doc,
        "The practical problem is that issuing an IT micro-credential and its digital badge does not by "
        "itself ensure that students will use the badge in job applications. Universities need evidence "
        "about the beliefs and conditions associated with intended use. The reviewed source study provides "
        "that evidence for accounting students at one New Zealand university, but it does not establish "
        "whether the same model operates among Thai IT students. A contextual replication can test the "
        "transferability of the published relationships while preserving a fully traceable framework and "
        "questionnaire.",
    )

    doc.add_heading("1.3 Research Objectives", level=2)
    add_numbered(
        doc,
        [
            "To describe Thai IT students' perceived usefulness and perceived ease of use of university-issued digital badges for job applications.",
            "To assess Thai IT students' intention to use university-issued digital badges when applying for jobs.",
            "To test the 13 direct relationships published in Steenkamp et al.'s research model using a Thai IT-student sample.",
            "To evaluate whether the source study's reflective measurement model demonstrates acceptable reliability and validity in the Thai context.",
        ],
    )

    doc.add_heading("1.4 Research Questions", level=2)
    add_numbered(
        doc,
        [
            "What are Thai IT students' perceptions of the usefulness and ease of use of university-issued digital badges designed to showcase employability skills?",
            "What general factors influence Thai IT students' perceptions of the usefulness and ease of use of those digital badges?",
            "To what extent do Thai IT students intend to use university-issued digital badges to showcase employability skills in job applications?",
            "Does the complete published Steenkamp et al. model demonstrate comparable measurement and structural relationships in the Thai IT-student context?",
        ],
    )
    add_source_note(
        doc,
        "RQ1-RQ3 retain the content of Steenkamp et al.'s first three research questions while changing "
        "the population from accounting students to Thai IT students. RQ4 states the replication objective.",
    )

    doc.add_heading("1.5 Scope of the Research", level=2)
    add_bullets(
        doc,
        [
            "Population: university students aged 18 or above who are enrolled in an IT-related program in Thailand.",
            "Focal object: a university-issued digital badge representing completion of an IT micro-credential and usable in job applications.",
            "Variables: the complete 11-construct model published by Steenkamp et al. (2024).",
            "Instrument: the 40 coded construct items printed in the source Appendix, using its seven-point agreement scale.",
            "Method: required pilot followed by a cross-sectional questionnaire survey and PLS-SEM.",
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
        "The study contributes by testing a complete published digital-badge acceptance model in a new "
        "national and disciplinary population. Practically, it can help Thai universities understand "
        "whether usefulness, ease of use, social influence, job relevance, institutional support, and "
        "computer-related beliefs are associated with intended badge use. The study does not claim that "
        "badges improve employability; it examines students' intention to use them as a way of showcasing "
        "skills.",
    )

    doc.add_heading("1.8 Definition of Key Terms", level=2)
    definition_rows = [["Term", "Working definition"]]
    definition_rows += [
        ["Micro-credential", "A record of learning outcomes acquired through a small volume of learning (Council of the European Union, 2022; UNESCO, 2022)."],
        ["University-issued digital badge", "A verifiable visual token issued by a university to represent an achievement; in this study, completion of an IT micro-credential."],
        ["Employability skills", "Skills and competencies that a student may communicate in job applications; no actual employment outcome is measured."],
        ["Contextual replication", "Use of the same published model, paths, constructs, and source instrument in a different population and setting, with all context changes disclosed."],
    ]
    add_table(doc, definition_rows, [1.65, 4.55], caption="Table 1.1: Working definitions", font_size=9.4)

    doc.add_page_break()
    doc.add_heading("CHAPTER 2: LITERATURE REVIEW", level=1)
    doc.add_heading("2.1 Digital Badges and IT Micro-Credentials", level=2)
    add_body(
        doc,
        "Steenkamp et al. (2024) distinguish a micro-credential, which records learning outcomes from a "
        "small volume of learning, from a digital badge, which is a visual token acknowledging learning or "
        "achievement. A badge may be applied to a micro-credential and can contain verifiable metadata. "
        "Kiiskila et al. (2023) likewise examine verifiable digital credentials from a student perspective, "
        "while Miao et al. (2024) demonstrate that student intention can be investigated directly in a "
        "micro-credential-program context. These studies support the context; Steenkamp et al. alone supply "
        "the model and questionnaire used in this thesis.",
    )

    doc.add_heading("2.2 Technology Acceptance Model and TAM3", level=2)
    add_body(
        doc,
        "The Technology Acceptance Model identifies perceived usefulness and perceived ease of use as "
        "central beliefs associated with technology acceptance (Davis, 1989). TAM3 extends this logic by "
        "specifying determinants of those beliefs, including social-influence processes, cognitive "
        "instrumental processes, control beliefs, computer anxiety, and computer playfulness (Venkatesh, "
        "2000; Venkatesh & Bala, 2008). Steenkamp et al. (2024) apply an extended, context-specific version "
        "of TAM3 to university-issued digital badges for job applications. The current study retains that "
        "application rather than using TAM as a general label for a newly assembled model.",
    )

    doc.add_heading("2.3 Variables in the Borrowed Model", level=2)
    construct_rows = [["Published construct", "Code", "Direct source location", "Published items"]]
    construct_rows += [[name, code, source_location, items] for name, code, source_location, items, _ in CONSTRUCTS]
    add_table(doc, construct_rows, [1.55, 0.55, 3.35, 0.75], caption="Table 2.1: Exact constructs retained from the source model", font_size=8.4)

    doc.add_heading("2.4 Relationships in the Borrowed Model", level=2)
    add_body(
        doc,
        "Perceived usefulness and perceived ease of use are modelled as direct determinants of behavioural "
        "intention, and ease of use is also modelled as a determinant of usefulness. Subjective norm is "
        "linked to intention, usefulness, and image. Image, job application relevance, and result "
        "demonstrability are linked to usefulness. Computer self-efficacy, perceptions of external control, "
        "computer anxiety, and computer playfulness are linked to ease of use. These are not relationships "
        "created for the Thai study; they are the complete 13-path structure in Steenkamp et al. (2024, "
        "Figure 1 and H1-H13).",
    )

    doc.add_heading("2.5 Previous Studies", level=2)
    prior_rows = [
        ["Study", "Context and method", "Use in this thesis"],
        ["Steenkamp et al. (2024)", "New Zealand accounting students; online survey; extended TAM3; PLS-SEM.", "Direct source of Figure 1, H1-H13, 11 constructs, 40 coded items, and analysis pattern."],
        ["Miao et al. (2024)", "University students; technology use for micro-credential programs; TAM and self-determination model.", "Same-field evidence that intention can be studied in a micro-credential-program context; not the item source."],
        ["Kiiskila et al. (2023)", "Higher-education students; qualitative study of verifiable digital credentials and micro-credentials.", "Supports student-value and digital-credential context; not the current quantitative framework."],
        ["Davis (1989); Venkatesh & Bala (2008)", "Foundational TAM and TAM3 theory and measures.", "Theoretical foundation and prior measurement sources identified by Steenkamp et al."],
    ]
    add_table(doc, prior_rows, [1.5, 2.35, 2.35], caption="Table 2.2: Most relevant previous studies", font_size=8.8)

    doc.add_page_break()
    doc.add_heading("CHAPTER 3: RESEARCH FRAMEWORK", level=1)
    doc.add_heading("3.1 Borrowed Theoretical and Research Framework", level=2)
    add_body(
        doc,
        "The framework is borrowed directly from Steenkamp et al. (2024). The source article identifies "
        "TAM3 as its theoretical basis and publishes the research model as Figure 1. The image below is a "
        "direct crop from the source PDF. No node or arrow has been generated, renamed, added, removed, or "
        "redirected.",
    )
    add_figure(
        doc,
        MODEL_CROP,
        "Figure 3.1: Research model reproduced directly from Steenkamp, Fisher, and Nesbit (2024, Figure 1, article p. 10).",
        width=5.75,
    )
    add_source_note(
        doc,
        "Source article license: CC BY-NC-ND 4.0. The figure content is unaltered; only surrounding page material is omitted in the direct crop.",
    )

    doc.add_heading("3.2 Application to the Thai IT-Student Context", level=2)
    add_body(
        doc,
        "The model itself is unchanged. The contextual replication changes the population from accounting "
        "students at a New Zealand university to students in IT-related programs in Thailand. The survey "
        "introduction defines the university-issued digital badge as the verifiable representation awarded "
        "for completing an IT micro-credential. The 40 construct statements continue to use the source "
        "referent, 'university-issued digital badges.'",
    )

    doc.add_heading("3.3 Research Hypotheses", level=2)
    add_body(
        doc,
        "The following hypotheses are reproduced from Steenkamp et al. (2024, pp. 11-13). The source's "
        "construct wording and direction are retained. H11 therefore also retains the source's grammatical "
        "form.",
    )
    add_hypotheses(doc)

    doc.add_heading("3.4 Operationalization", level=2)
    op_rows = [["Construct", "Code", "Source item codes", "Count", "Scale"]]
    op_rows += [[name, code, codes, count, "7-point Likert"] for name, code, _, codes, count in CONSTRUCTS]
    add_table(doc, op_rows, [1.8, 0.6, 1.55, 0.55, 1.7], caption="Table 3.1: Operationalization of the exact source constructs", font_size=8.8)
    add_source_note(
        doc,
        "All item wordings appear in Appendix A. Direct source: Steenkamp et al. (2024), Appendix, article pp. 28-29.",
    )

    doc.add_page_break()
    doc.add_heading("CHAPTER 4: RESEARCH METHODOLOGY", level=1)
    doc.add_heading("4.1 Research Design", level=2)
    add_body(
        doc,
        "The study uses a quantitative, cross-sectional online questionnaire and a contextual-replication "
        "design. It tests the complete published model rather than comparing experimental scenarios. No "
        "interviews or open-ended research questions are included. All model indicators are treated as "
        "reflective, matching the source study.",
    )

    doc.add_heading("4.2 Population and Sampling", level=2)
    add_body(
        doc,
        "Eligible participants will be aged 18 or above and currently enrolled in an IT-related university "
        "program in Thailand. Participants in the pilot will not enter the main-study dataset. The proposed "
        "minimum is 384 usable main-study responses, with a recruitment target of 400. Recruitment channels "
        "and participating institutions remain subject to professor and university approval. The final "
        "report will state the sampling method and response flow without implying probability sampling if "
        "non-probability recruitment is used.",
    )

    doc.add_heading("4.3 Required Pilot", level=2)
    add_body(
        doc,
        "A pilot with 40 eligible Thai IT students will occur before the main survey. It will check "
        "eligibility logic, comprehension, completion time, missing responses, scale use, and preliminary "
        "measurement performance. Pilot participants will be excluded from the main sample. The English "
        "source master and item-level provenance will remain unchanged; any Thai-language or item decision "
        "after the pilot will be documented and approved rather than silently incorporated.",
    )

    doc.add_heading("4.4 Questionnaire Design", level=2)
    add_numbered(
        doc,
        [
            "Participant information, consent, and eligibility screening.",
            "A neutral explanation of a university-issued digital badge representing an IT micro-credential, including that it can be verified and displayed in job applications.",
            "Demographic and context questions, such as age band, field of study, study level, and prior awareness or use of digital badges.",
            "The 40 coded construct items from Steenkamp et al. (2024), retaining the source's seven-point agreement scale.",
        ],
    )
    add_note(
        doc,
        "Source anomaly: the Appendix contains an unnumbered employer sentence after SN4. Because the "
        "article states that 40 items were analysed and Table 2 reports SN1-SN4 only, the present instrument "
        "does not invent SN5. BI3's apparent typo is retained in the English source master; any correction "
        "requires approval before fielding.",
    )

    doc.add_heading("4.5 Translation", level=2)
    add_body(
        doc,
        "The English source instrument will be the audit master. After the supervisor approves the English "
        "instrument, it will be translated into Thai and independently back-translated into English. "
        "Differences will be reconciled against construct meaning and the exact source statement (Brislin, "
        "1970). Both approved language versions will be retained in the thesis appendix.",
    )

    doc.add_heading("4.6 Data Collection and Ethics", level=2)
    add_body(
        doc,
        "The questionnaire will be distributed only after the required academic and ethical permissions. "
        "Participation will be voluntary. The information sheet will explain the study, inclusion criteria, "
        "estimated completion time, withdrawal conditions, confidentiality, and data handling. No claim of "
        "institutional approval is made in this draft.",
    )

    doc.add_heading("4.7 Statistical Treatment", level=2)
    add_numbered(
        doc,
        [
            "Screen records for consent, eligibility, duplicates, missingness, completion quality, and coding errors.",
            "Describe the sample and item distributions using frequencies, means, and standard deviations.",
            "Assess reflective measurement quality using outer loadings, Cronbach's alpha, composite reliability, average variance extracted, and discriminant validity.",
            "Assess structural-model collinearity using variance inflation factors.",
            "Estimate the 13 published paths and R-squared values using PLS-SEM with 5,000 bootstrap samples, matching the source study's general procedure.",
            "Report all retained and removed items from the Thai data. Do not automatically copy the source sample's deletion of PEC4, CPLAY4, and RES4.",
        ],
    )

    doc.add_heading("4.8 Claim Boundaries", level=2)
    add_body(
        doc,
        "The primary findings will concern measured perceptions and behavioural intention. The analysis may "
        "support or fail to support replicated paths, but it will not prove causal effects, actual badge use, "
        "employer recognition, improved employability, or employment outcomes.",
    )

    heading = doc.add_heading("CHAPTER 5: RESEARCH RESULTS", level=1)
    heading.paragraph_format.page_break_before = True
    add_note(doc, "No results are available. Complete this chapter only after the required pilot and main-study analysis.")
    for heading in [
        "5.1 Response Flow and Data Screening",
        "5.2 Participant Characteristics",
        "5.3 Descriptive Statistics",
        "5.4 Measurement Model",
        "5.5 Structural Model and Hypothesis Tests",
        "5.6 Summary of Results",
    ]:
        doc.add_heading(heading, level=2)
        add_body(doc, "Insert verified results, tables, and interpretation here after analysis.")

    doc.add_page_break()
    doc.add_heading("CHAPTER 6: DISCUSSION AND CONCLUSION", level=1)
    add_note(doc, "No discussion or conclusion is claimed before real results exist.")
    for heading in [
        "6.1 Discussion by Replicated Hypothesis",
        "6.2 Comparison with Steenkamp et al. (2024)",
        "6.3 Theoretical Implications",
        "6.4 Practical Implications",
        "6.5 Limitations",
        "6.6 Future Research",
        "6.7 Conclusion",
    ]:
        doc.add_heading(heading, level=2)
        add_body(doc, "Complete this section after the empirical findings are available.")

    doc.add_page_break()
    doc.add_heading("References", level=1)
    add_references(doc)

    doc.add_page_break()
    doc.add_heading("Appendix A: English Source Questionnaire for Pilot Review", level=1)
    add_body(
        doc,
        "Instruction: Please indicate how strongly you agree or disagree with each statement. The "
        "construct-item wording and response scale are reproduced from Steenkamp et al. (2024).",
        no_indent=True,
    )
    add_questionnaire(doc)

    doc.add_heading("Appendix B: Source Measurement Decisions", level=1)
    decision_rows = [
        ["Issue", "Published source", "Current treatment"],
        ["Item count", "40 coded construct items", "All 40 coded items administered in the pilot."],
        ["Unnumbered employer sentence", "Printed after SN4 but not coded; absent from the 40-item measurement table.", "Not assigned a code and not scored."],
        ["PEC4, CPLAY4, RES4", "Administered, then removed for source-sample loadings below .40.", "Administered; any Thai-sample removal will be reported from Thai data."],
        ["BI3", "Source prints 'use the when'.", "Retained verbatim in the English source master; any correction requires approval before fielding."],
    ]
    add_table(doc, decision_rows, [1.25, 2.45, 2.5], caption="Table B.1: Transparent handling of source-instrument details", font_size=8.8)

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
    prepare_source_excerpts()
    build_master()
    build_evidence_pack()
    print(MASTER_OUT)
    print(EVIDENCE_OUT)
    print(MODEL_CROP)
    print(APPENDIX_28_CROP)
    print(APPENDIX_29_CROP)


if __name__ == "__main__":
    main()
