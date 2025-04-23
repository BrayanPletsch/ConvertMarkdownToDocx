from pathlib import Path
from docx import Document
from convert_markdown.docx_generator import generate_docx


def test_generate_docx_file_created(tmp_path):
    out_file = tmp_path / "test_output.docx"
    markdown = "Just a simple paragraph."
    generate_docx(markdown, str(out_file))
    assert out_file.exists(), "O arquivo DOCX não foi criado."


def test_generate_docx_content(tmp_path):
    markdown = """
# Título 1
Parágrafo normal.
- Item 1
1. Primeiro
**Negrito** no meio do texto.
"""
    out_file = tmp_path / "content_test.docx"
    generate_docx(markdown, str(out_file))

    doc = Document(str(out_file))
    paras = list(doc.paragraphs)

    heading_paras = [p for p in paras if p.style.name.lower().startswith('heading')]
    assert heading_paras, "Deve haver ao menos um parágrafo de heading"
    assert heading_paras[0].text == "Título 1"

    idx_heading = paras.index(heading_paras[0])
    assert paras[idx_heading + 1].text == "Parágrafo normal."

    bullet_paras = [p for p in paras if p.style.name == 'List Bullet']
    assert bullet_paras and bullet_paras[0].text == "Item 1"

    num_paras = [p for p in paras if p.style.name == 'List Number']
    assert num_paras and num_paras[0].text == "Primeiro"

    bold_paras = [p for p in paras if any(r.bold for r in p.runs)]
    bold_paras = [p for p in bold_paras if p.text.startswith('Negrito')]
    assert bold_paras, "Deve haver pelo menos um parágrafo com run bold"
    runs = bold_paras[0].runs
    bold_runs = [r for r in runs if r.bold]
    assert bold_runs[0].text == "Negrito"
