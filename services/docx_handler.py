import io
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def extract_text_from_docx(file_bytes: bytes) -> str:
    """Reads an uploaded DOCX file and extracts clean paragraph and table text."""
    doc = docx.Document(io.BytesIO(file_bytes))
    full_text = []

    # Extract paragraphs
    for para in doc.paragraphs:
        if para.text.strip():
            full_text.append(para.text.strip())

    # Extract table rows
    for table in doc.tables:
        for row in table.rows:
            row_data = [cell.text.strip() for cell in row.cells if cell.text.strip()]
            if row_data:
                full_text.append(" | ".join(row_data))

    return "\n\n".join(full_text)

def generate_slm_docx(title: str, content_markdown: str, metadata: dict = None) -> bytes:
    """Transforms Markdown-formatted generation into a university-formatted DOCX."""
    doc = docx.Document()

    # Document Header / Title
    title_p = doc.add_paragraph()
    title_run = title_p.add_run(title)
    title_run.font.name = "Calibri"
    title_run.font.size = Pt(20)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(31, 78, 120)

    if metadata:
        meta_p = doc.add_paragraph()
        meta_p.paragraph_format.space_after = Pt(12)
        meta_run = meta_p.add_run(
            f"Deliverable Type: {metadata.get('deliverable_type', 'SLM Unit')} | "
            f"Academic Level: {metadata.get('academic_level', 'N/A')}\n"
            f"Generated via: AI Maker Engine (Gemini Production)"
        )
        meta_run.font.size = Pt(9.5)
        meta_run.font.italic = True
        meta_run.font.color.rgb = RGBColor(110, 110, 110)

    doc.add_paragraph("―" * 50)

    # Parse basic Markdown headers and body text
    for line in content_markdown.split("\n"):
        line_clean = line.strip()
        if not line_clean:
            continue

        if line_clean.startswith("# "):
            p = doc.add_paragraph()
            run = p.add_run(line_clean.replace("# ", ""))
            run.font.size = Pt(16)
            run.font.bold = True
            run.font.color.rgb = RGBColor(31, 78, 120)
        elif line_clean.startswith("## "):
            p = doc.add_paragraph()
            run = p.add_run(line_clean.replace("## ", ""))
            run.font.size = Pt(13)
            run.font.bold = True
            run.font.color.rgb = RGBColor(65, 113, 156)
        elif line_clean.startswith("### "):
            p = doc.add_paragraph()
            run = p.add_run(line_clean.replace("### ", ""))
            run.font.size = Pt(11.5)
            run.font.bold = True
        elif line_clean.startswith("- ") or line_clean.startswith("* "):
            p = doc.add_paragraph(style='List Bullet')
            p.add_run(line_clean[2:])
        else:
            p = doc.add_paragraph()
            p.add_run(line_clean)
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(4)

    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()