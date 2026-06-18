from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import weasyprint
import os

# ── CV DATA ────────────────────────────────────────────────────────────────────

NAME = "IBRAHEEM MUSTAFA"
TITLE = "Information Security Engineer – GRC"
CONTACT = [
    "ibraheem.mustafa00@gmail.com",
    "+92-3018442580",
    "Lahore, Pakistan",
    "LinkedIn: Ibraheem Mustafa",
]

SUMMARY = (
    "GRC-focused Information Security Engineer with 2+ years of experience delivering "
    "ISO 27001:2022 audit engagements, ISMS implementations, and compliance assessments "
    "across local and international organizations. Conducted end-to-end internal and "
    "external audits for 5 organizations including Kualitatem, Virtual Force, and "
    "international clients, driving ISMS certification and compliance outcomes. Certified "
    "ISO 27001:2022 Lead Auditor (CQI|IRCA & Mastermind Assurance) with hands-on "
    "expertise in SAMA CSF, SOC 2, GDPR, Vanta, and Sahl."
)

SKILLS = (
    "ISO 27001:2022  |  Internal Audit  |  External Audit  |  ISMS Implementation  |  "
    "SAMA CSF  |  SOC 2  |  GDPR  |  Risk Assessment  |  Risk Treatment Plan  |  "
    "Gap Analysis  |  Statement of Applicability (SoA)  |  Non-Conformity Reporting  |  "
    "Audit Planning & Execution  |  Compliance Reporting  |  Vanta  |  Sahl  |  "
    "Top Management Presentations"
)

EXPERIENCE = [
    {
        "company": "Kualitatem Inc.",
        "title": "Information Security Engineer – Information Security Department",
        "dates": "December 2023 – Present",
        "location": "Lahore, Pakistan",
        "bullets": [
            "Strengthened Kualitatem's ISO 27001:2022 compliance posture, as measured by a successful external certification audit with zero major non-conformities, by leading end-to-end internal audit activities covering planning, evidence collection, controls assessment, and non-conformity reporting to top management.",
            "Maintained ISO 27001:2022 certification for Virtual Force, as measured by a passed surveillance audit, by executing a full audit lifecycle including scope review, controls re-evaluation, and verification of previously identified non-conformity closures.",
            "Achieved ISO 27001:2022 certification readiness for 3 clients (2 local, 1 international), as measured by successful audit outcomes across all engagements, by delivering structured end-to-end internal and external audit programs, identifying control gaps, and presenting prioritized findings to client management.",
            "Achieved ISO 27001:2022 certification for a client from ground zero, as measured by full certification attainment, by leading end-to-end ISMS implementation including risk assessment, SoA development, control design, and pre-certification audit preparation.",
            "Renewed SOC 2 certification for Kualitatem, as measured by continued compliance with all Trust Service Criteria, by coordinating cross-functional evidence collection via Vanta, liaising with external auditors, and remediating identified control gaps ahead of assessment.",
            "Enhanced client compliance posture against SAMA CSF, as measured by detailed gap assessment reports delivered to senior management, by conducting structured stakeholder interviews, analyzing control deficiencies via Sahl, and producing prioritized remediation roadmaps.",
            "Supported organization-wide GDPR compliance program, as measured by documented data processing assessments, by collaborating with the Internal Audit team to review data handling practices and implement corrective control measures.",
        ],
    }
]

EDUCATION = [
    ("Masters in Cyber Security – CGPA 3.57", "Bahcesehir University, Turkiye", "2023"),
    ("Bachelors in Computer Science – CGPA 3.40", "Bahauddin Zakariya University, Pakistan", "2018"),
]

CERTIFICATIONS = [
    ("ISO 27001:2022 Lead Auditor", "CQI|IRCA"),
    ("ISO/IEC 27001:2022 Lead Auditor", "Mastermind Assurance"),
    ("ISO 27001:2022 INFOSEC Associate", "SkillFront"),
    ("eJPT – Junior Penetration Tester", "INE eLearnSecurity"),
]

LANGUAGES = [
    ("English", "Full Professional Proficiency"),
    ("Urdu", "Native / Bilingual Proficiency"),
    ("Turkish", "Elementary Proficiency"),
]

# ── DOCX GENERATION ────────────────────────────────────────────────────────────

def set_cell_bg(cell, color_hex):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def add_hrule(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '2C5282')
    pBdr.append(bottom)
    pPr.append(pBdr)

def section_heading(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text.upper())
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0x2C, 0x52, 0x82)
    add_hrule(doc)

def build_docx(path):
    doc = Document()

    # Page margins
    for section in doc.sections:
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)

    # Default style
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(10)

    # ── Header ──
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(NAME)
    run.bold = True
    run.font.size = Pt(22)
    run.font.color.rgb = RGBColor(0x1A, 0x36, 0x5C)

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_after = Pt(4)
    run2 = p2.add_run(TITLE)
    run2.font.size = Pt(12)
    run2.font.color.rgb = RGBColor(0x2C, 0x52, 0x82)
    run2.bold = True

    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.space_after = Pt(6)
    contact_str = "  |  ".join(CONTACT)
    run3 = p3.add_run(contact_str)
    run3.font.size = Pt(9)
    run3.font.color.rgb = RGBColor(0x44, 0x44, 0x44)

    add_hrule(doc)

    # ── Summary ──
    section_heading(doc, "Professional Summary")
    p = doc.add_paragraph(SUMMARY)
    p.paragraph_format.space_after = Pt(4)

    # ── Skills ──
    section_heading(doc, "Core Skills")
    p = doc.add_paragraph(SKILLS)
    p.paragraph_format.space_after = Pt(4)

    # ── Experience ──
    section_heading(doc, "Work Experience")
    for job in EXPERIENCE:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(job["company"])
        r.bold = True
        r.font.size = Pt(11)

        p2 = doc.add_paragraph()
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(0)
        r2 = p2.add_run(job["title"])
        r2.bold = True
        r2.font.size = Pt(10)

        p3 = doc.add_paragraph()
        p3.paragraph_format.space_before = Pt(0)
        p3.paragraph_format.space_after = Pt(4)
        r3 = p3.add_run(f"{job['dates']}  |  {job['location']}")
        r3.italic = True
        r3.font.size = Pt(9)
        r3.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

        for bullet in job["bullets"]:
            p = doc.add_paragraph(style='List Bullet')
            p.paragraph_format.left_indent = Inches(0.25)
            p.paragraph_format.space_after = Pt(3)
            run = p.add_run(bullet)
            run.font.size = Pt(10)

    # ── Education ──
    section_heading(doc, "Education")
    for degree, uni, year in EDUCATION:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(degree)
        r.bold = True
        p2 = doc.add_paragraph()
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(2)
        r2 = p2.add_run(f"{uni}  |  {year}")
        r2.font.size = Pt(9)
        r2.italic = True
        r2.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    # ── Certifications ──
    section_heading(doc, "Certifications")
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    hdr = table.rows[0].cells
    hdr[0].text = "Certification"
    hdr[1].text = "Issuing Body"
    for cell in hdr:
        set_cell_bg(cell, "2C5282")
        for para in cell.paragraphs:
            for run in para.runs:
                run.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                run.font.size = Pt(10)
    for cert, issuer in CERTIFICATIONS:
        row = table.add_row().cells
        row[0].text = cert
        row[1].text = issuer
        for cell in row:
            for para in cell.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(10)

    doc.add_paragraph()

    # ── Languages ──
    section_heading(doc, "Languages")
    for lang, prof in LANGUAGES:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r1 = p.add_run(f"{lang}: ")
        r1.bold = True
        p.add_run(prof)

    doc.save(path)
    print(f"DOCX saved: {path}")


# ── HTML / PDF GENERATION ──────────────────────────────────────────────────────

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<style>
  @page {{ margin: 1.5cm 2cm; }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: Calibri, Arial, sans-serif; font-size: 10pt; color: #222; line-height: 1.4; }}
  .header {{ text-align: center; border-bottom: 3px solid #2C5282; padding-bottom: 8px; margin-bottom: 10px; }}
  .header h1 {{ font-size: 22pt; color: #1A365C; letter-spacing: 1px; margin-bottom: 2px; }}
  .header .title {{ font-size: 12pt; color: #2C5282; font-weight: bold; margin-bottom: 4px; }}
  .header .contact {{ font-size: 9pt; color: #444; }}
  .section-title {{ font-size: 11pt; font-weight: bold; color: #2C5282; text-transform: uppercase;
                    letter-spacing: 0.5px; border-bottom: 1.5px solid #2C5282;
                    margin-top: 12px; margin-bottom: 5px; padding-bottom: 2px; }}
  p {{ margin-bottom: 4px; }}
  .skills {{ margin-bottom: 4px; }}
  .job-company {{ font-size: 11pt; font-weight: bold; margin-top: 8px; margin-bottom: 1px; }}
  .job-title {{ font-weight: bold; font-size: 10pt; margin-bottom: 1px; }}
  .job-meta {{ font-size: 9pt; color: #666; font-style: italic; margin-bottom: 4px; }}
  ul {{ margin-left: 18px; margin-bottom: 4px; }}
  li {{ margin-bottom: 3px; }}
  .edu-entry {{ margin-bottom: 6px; }}
  .edu-degree {{ font-weight: bold; }}
  .edu-uni {{ font-size: 9pt; color: #555; font-style: italic; }}
  table {{ width: 100%; border-collapse: collapse; margin-top: 4px; font-size: 10pt; }}
  th {{ background: #2C5282; color: #fff; font-weight: bold; padding: 5px 8px; text-align: left; }}
  td {{ border: 1px solid #ccc; padding: 4px 8px; }}
  tr:nth-child(even) td {{ background: #f5f8ff; }}
  .lang-row {{ display: inline-block; margin-right: 20px; margin-bottom: 3px; }}
  .lang-name {{ font-weight: bold; }}
</style>
</head>
<body>

<div class="header">
  <h1>{name}</h1>
  <div class="title">{title}</div>
  <div class="contact">{contact}</div>
</div>

<div class="section-title">Professional Summary</div>
<p>{summary}</p>

<div class="section-title">Core Skills</div>
<p class="skills">{skills}</p>

<div class="section-title">Work Experience</div>
{experience}

<div class="section-title">Education</div>
{education}

<div class="section-title">Certifications</div>
<table>
  <tr><th>Certification</th><th>Issuing Body</th></tr>
  {cert_rows}
</table>

<div class="section-title">Languages</div>
<p>{languages}</p>

</body>
</html>"""


def build_pdf(html_path, pdf_path):
    weasyprint.HTML(filename=html_path).write_pdf(pdf_path)
    print(f"PDF saved: {pdf_path}")


def build_html():
    contact_str = "  |  ".join(CONTACT)

    exp_html = ""
    for job in EXPERIENCE:
        bullets = "".join(f"<li>{b}</li>" for b in job["bullets"])
        exp_html += f"""
        <div class="job-company">{job['company']}</div>
        <div class="job-title">{job['title']}</div>
        <div class="job-meta">{job['dates']}  |  {job['location']}</div>
        <ul>{bullets}</ul>
        """

    edu_html = ""
    for degree, uni, year in EDUCATION:
        edu_html += f"""
        <div class="edu-entry">
          <div class="edu-degree">{degree}</div>
          <div class="edu-uni">{uni}  |  {year}</div>
        </div>
        """

    cert_rows = "".join(f"<tr><td>{c}</td><td>{i}</td></tr>" for c, i in CERTIFICATIONS)
    lang_str = "  |  ".join(f"<span class='lang-name'>{l}</span>: {p}" for l, p in LANGUAGES)

    return HTML_TEMPLATE.format(
        name=NAME,
        title=TITLE,
        contact=contact_str,
        summary=SUMMARY,
        skills=SKILLS,
        experience=exp_html,
        education=edu_html,
        cert_rows=cert_rows,
        languages=lang_str,
    )


if __name__ == "__main__":
    out_dir = "/home/user/CV-Creator---Claude-AI/output"
    os.makedirs(out_dir, exist_ok=True)

    docx_path = f"{out_dir}/Ibraheem_Mustafa_GRC_CV.docx"
    html_path = f"{out_dir}/Ibraheem_Mustafa_GRC_CV.html"
    pdf_path  = f"{out_dir}/Ibraheem_Mustafa_GRC_CV.pdf"

    build_docx(docx_path)

    html = build_html()
    with open(html_path, "w") as f:
        f.write(html)
    print(f"HTML saved: {html_path}")

    build_pdf(html_path, pdf_path)
    print("Done.")
