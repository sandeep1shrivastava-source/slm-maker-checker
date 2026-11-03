import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

def generate_updated_manual():
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(0.8)
        s.bottom_margin = Inches(0.8)
        s.left_margin = Inches(0.8)
        s.right_margin = Inches(0.8)

    p_title = doc.add_paragraph()
    r_title = p_title.add_run("ACADEMIC CONTENT & AUDIT PIPELINE (MAKER–CHECKER–AUDITOR)")
    r_title.bold = True
    r_title.font.size = Pt(18)
    r_title.font.color.rgb = RGBColor(0, 51, 102)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p_sub = doc.add_paragraph()
    r_sub = p_sub.add_run("Standard Operating Procedure & User Manual (v2.0 - Comprehensive Release)")
    r_sub.font.size = Pt(11)
    r_sub.font.italic = True
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph("―" * 60)

    doc.add_heading("1. Executive Overview & Pipeline Architecture", level=1)
    doc.add_paragraph(
        "The Academic Content & Audit Pipeline is an institutional-grade platform designed to automate the authoring, "
        "syllabus benchmarking, academic quality assurance, and prompt fidelity auditing of higher-education learning "
        "materials. Hosted at slm-maker-checker-new-ad99.streamlit.app, the platform operates across three sequential, "
        "interlocked modules with complete role-based governance:"
    )

    p_flow = doc.add_paragraph()
    p_flow.add_run("• Stage 1 (AI Maker): ").bold = True
    p_flow.add_run("Ingests syllabus and benchmark references to generate full-length instructional manuscripts. Exports deliverable DOCX.\n")
    p_flow.add_run("• Stage 2 (Quality Checker - 12 QA Framework): ").bold = True
    p_flow.add_run("Audits draft against 12 pedagogical dimensions, syllabus coverage %, and factual accuracy. Generates both an executive Word report (.DOCX) and a comprehensive 3-sheet audit workbook (.XLSX).\n")
    p_flow.add_run("• Stage 3 (Prompt Meta-Auditor): ").bold = True
    p_flow.add_run("Performs deep prompt guardrail verification, Bloom's taxonomy action verb auditing, and structural drift analysis. Exports certified accreditation dossier (.DOCX).")

    doc.add_heading("2. Role-Based Access Control (RBAC) & Governance", level=1)
    doc.add_paragraph("Access is strictly restricted via the institutional access manager. Execution buttons dynamically activate based on verified user credentials:")

    t_rbac = doc.add_table(rows=1, cols=4)
    t_rbac.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_rbac.style = "Table Grid"
    headers_rbac = ["Role Title", "Operational Mandate", "Pipeline Authority", "Registered Email"]
    for i, h in enumerate(headers_rbac):
        cell = t_rbac.rows[0].cells[i]
        cell.text = h
        for p in cell.paragraphs:
            for r in p.runs:
                r.bold = True

    roles_data = [
        ("Lead Administrator", "Full system governance, cross-stage auditing, university registry", "Stages 1, 2, and 3", "admin@university.edu"),
        ("Content Author (Maker)", "Course syllabus ingestion, prompt tuning, initial manuscript drafting", "Stage 1 Only", "author@university.edu"),
        ("SME Quality Reviewer", "Pedagogical verification, 12-QA scoring, critical defect logging", "Stage 2 Only", "sme@university.edu"),
        ("Prompt Compliance Auditor", "Prompt guardrail compliance, regulatory oversight, accreditation dossiers", "Stage 3 Only", "auditor@university.edu")
    ]
    for r in roles_data:
        row = t_rbac.add_row().cells
        row[0].text = r[0]
        row[1].text = r[1]
        row[2].text = r[2]
        row[3].text = r[3]

    doc.add_paragraph("")

    doc.add_heading("3. The 12-Parameter Academic QA Framework", level=1)
    doc.add_paragraph("Stage 2 evaluates generated content against 12 standardized higher-education quality dimensions:")

    t_12qa = doc.add_table(rows=1, cols=4)
    t_12qa.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_12qa.style = "Table Grid"
    headers_12qa = ["#", "QA Dimension", "Weight", "Pedagogical Benchmark & Scope"]
    for i, h in enumerate(headers_12qa):
        cell = t_12qa.rows[0].cells[i]
        cell.text = h
        for p in cell.paragraphs:
            for r in p.runs:
                r.bold = True

    qa_params = [
        ("1", "Syllabus & Curriculum Mapping", "10%", "100% alignment of topics, units, and learning objectives to prescribed outline."),
        ("2", "Bloom's Taxonomy & Action Verbs", "10%", "Consistent application of measurable cognitive verbs (Analyze, Evaluate, Synthesize)."),
        ("3", "Conceptual Depth & Rigor", "10%", "Postgraduate/Undergraduate level thoroughness, theoretical depth, and academic sufficiency."),
        ("4", "Real-world Case Studies & Applications", "8%", "Inclusion of illustrative industry case studies, practical scenarios, and applied exercises."),
        ("5", "Quadrant-II Structural Architecture", "8%", "Compliance with UGC/DEB self-learning material flow (Objectives, Content, SAQs, Summary)."),
        ("6", "Self-Assessment Questions (SAQs)", "8%", "Presence of formative MCQs with distractors, review questions, and conceptual prompts."),
        ("7", "Factual & Theoretical Accuracy", "10%", "Verification of equations, domain theorems, factual claims, and analytical data."),
        ("8", "Originality & Non-Hallucination", "10%", "Zero phantom citations, unattributed external data, or fabricated assertions."),
        ("9", "Academic Tone & Scholarly Voice", "6%", "Formal objective instructional tone; complete absence of conversational AI boilerplate."),
        ("10", "Key Takeaways & Summary Synthesis", "6%", "Clear conceptual recap, unit review checklists, and actionable concept synthesis."),
        ("11", "Benchmark & Reference Adherence", "8%", "Direct anchoring to prescribed reference textbooks and institutional curricula."),
        ("12", "Scaffolding, Tables & Typography", "6%", "Effective visual hierarchy, summary comparison tables, and pedagogical callout boxes.")
    ]
    for q in qa_params:
        row = t_12qa.add_row().cells
        row[0].text = q[0]
        row[1].text = q[1]
        row[2].text = q[2]
        row[3].text = q[3]

    doc.add_paragraph("")

    doc.add_heading("4. End-to-End Operational Workflow", level=1)

    doc.add_heading("Step 1: Authentication & Institutional Scoping", level=2)
    doc.add_paragraph(
        "1. Open the portal in your browser.\n"
        "2. In the sidebar under 'User Authentication', enter your registered institutional email (e.g., admin@university.edu).\n"
        "3. Once verified, configure the 'Institutional Scope': University Name, Program/Discipline, Deliverable Format, and Academic Level."
    )

    doc.add_heading("Step 2: Source File Ingestion", level=2)
    doc.add_paragraph(
        "1. Panel 1 (Syllabus Outline): Upload a .docx or .txt file containing the unit syllabus, or paste text manually.\n"
        "2. Panel 2 (Reference Material): Upload prescribed textbook extracts, course notes, or case studies.\n"
        "3. Verify that green confirmation badges appear confirming clean parsing."
    )

    doc.add_heading("Step 3: Stage 1 Execution (AI Maker)", level=2)
    doc.add_paragraph(
        "1. Click 'Step 1: Run AI Maker'.\n"
        "2. Streaming tokens process over an extended 600-second window. Do not refresh the page while the spinner runs.\n"
        "3. Review the preview box, then click '📥 Download [Format] Manuscript (.DOCX)' to export the complete deliverable."
    )

    doc.add_heading("Step 4: Stage 2 Execution (Quality Checker & 12 QA)", level=2)
    doc.add_paragraph(
        "1. Click 'Step 2: Run Academic Audit' (automatically unlocks once Stage 1 completes).\n"
        "2. Review the top scorecard: Audit Score / 100, Syllabus Coverage %, Critical Errors, and Verification Decision.\n"
        "3. Expand the '12-Parameter Academic Quality Matrix' to inspect dimensional compliance.\n"
        "4. Download Deliverables:\n"
        "   • '📄 Download Stage 2 QA Review Report (.DOCX)': Executive Word report with scorecard, matrix, and SME sign-off footer.\n"
        "   • '📊 Download 12-QA Evaluation Report (.XLSX)': Formatted 3-sheet Excel workbook (Executive Summary, 12 QA Matrix, and Error Register)."
    )

    doc.add_heading("Step 5: Stage 3 Execution (Prompt Meta-Auditor)", level=2)
    doc.add_paragraph(
        "1. Click 'Step 3: Run Prompt Audit'.\n"
        "2. Evaluates prompt fidelity, negative constraint violations, and structural drift.\n"
        "3. Click '📄 Download Stage 3 Prompt Audit Report (.DOCX)' to export the certified meta-audit dossier."
    )

    doc.add_heading("5. Troubleshooting & Operational Guardrails", level=1)
    doc.add_paragraph(
        "• Unclickable Buttons: Stage 2 and Stage 3 buttons remain disabled until Stage 1 has generated a manuscript in session memory. Ensure your verified role has permission for the stage.\n"
        "• Long Generation Times: Large documents (3,500+ words) take 90 to 180 seconds to stream. Network sockets remain open via token streaming to prevent 504 timeouts.\n"
        "• Session Resets: Refreshing the page (F5) clears the temporary session state. Ensure you download generated DOCX/XLSX reports before navigating away."
    )

    doc.add_paragraph("\n" + "―" * 60)
    p_foot = doc.add_paragraph("University Academic AI Quality Board • Standard Operating Procedure v2.0")
    p_foot.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_foot.style.font.italic = True
    p_foot.style.font.size = Pt(9)

    doc.save("AI_Pipeline_User_Manual.docx")
    print("AI_Pipeline_User_Manual.docx generated successfully!")

if __name__ == "__main__":
    generate_updated_manual()