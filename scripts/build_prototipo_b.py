"""Gera os arquivos da Versão B (planilha-índice e ficha-modelo) da Entrega 5.

Requer openpyxl e python-docx. Os arquivos ficam em
entregas/05-prototipar-e-comparar/prototipos/versao_b/.
"""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn as dqn
from docx.shared import Cm, Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

OUT = Path(__file__).resolve().parents[1] / "entregas/05-prototipar-e-comparar/prototipos/versao_b"
OUT.mkdir(parents=True, exist_ok=True)
NAVY = "15365F"

# ---------------- Planilha-índice ----------------
wb = Workbook()
ws = wb.active
ws.title = "Índice"
ws["A1"] = "Caderno do grupo — Cálculo I · Grupo da turma 2"
ws["A1"].font = Font(name="Liberation Sans", size=16, bold=True, color=NAVY)
ws["A2"] = "Onde paramos: Lista 3, exercício 6  (atualizar ao fim de cada sessão)"
ws["A2"].font = Font(name="Liberation Sans", size=13, bold=True, color="B45309")
ws.merge_cells("A2:G2")
ws["A2"].fill = PatternFill("solid", fgColor="FDF1E3")

head = ["Lista", "Ex.", "Assunto", "Situação", "Quem registrou", "Dica-chave ou dúvida", "Ficha"]
rows = [
    ("3", "1", "substituição", "resolvido", "Alex", "Troque o que está dentro da raiz por u.", "Ficha_L3_E1"),
    ("3", "2", "substituição", "resolvido", "Duda", "Confira se du aparece na integral.", "Ficha_L3_E2"),
    ("3", "3", "substituição", "resolvido", "Bia", "Ajuste a constante antes de integrar.", "Ficha_L3_E3"),
    ("3", "4", "partes duas vezes", "com dúvida", "Caio", "Dúvida: quando aplicar partes duas vezes?", "Ficha_L3_E4"),
    ("3", "5", "por partes", "resolvido", "Bia", "Chame de u o que fica mais simples ao derivar.", "Ficha_L3_E5"),
    ("3", "6", "frações parciais", "pendente", "", "", ""),
    ("2", "7", "por partes", "resolvido", "Alex", "Em ln(x), use u = ln(x) e dv = dx.", "Ficha_L2_E7"),
]
thin = Side(style="thin", color="CBD5E1")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
for j, h in enumerate(head, start=1):
    c = ws.cell(row=4, column=j, value=h)
    c.font = Font(name="Liberation Sans", size=12, bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=NAVY)
    c.alignment = Alignment(vertical="center")
    c.border = border
fills = {"resolvido": "E6F4F1", "com dúvida": "FDF1E3", "pendente": "F1F5F9"}
for i, r in enumerate(rows, start=5):
    for j, v in enumerate(r, start=1):
        c = ws.cell(row=i, column=j, value=v)
        c.font = Font(name="Liberation Sans", size=12,
                      color="1D4ED8" if j == 7 else "1E293B", underline="single" if j == 7 and v else None)
        c.alignment = Alignment(vertical="center", wrap_text=True)
        c.border = border
    ws.cell(row=i, column=4).fill = PatternFill("solid", fgColor=fills[r[3]])
    ws.row_dimensions[i].height = 22
for col, w in zip("ABCDEFG", (7, 5, 18, 13, 17, 50, 14)):
    ws.column_dimensions[col].width = w
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_options.gridLines = False
wb.save(OUT / "Indice_Calculo_I.xlsx")

# ---------------- Ficha-modelo ----------------
doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(16), Cm(19)
for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
    setattr(sec, side, Cm(1.2))
style = doc.styles["Normal"]
style.font.name = "Liberation Sans"
style.font.size = Pt(12)

t = doc.add_paragraph()
run = t.add_run("Ficha de resolução")
run.bold = True
run.font.size = Pt(18)
run.font.color.rgb = RGBColor.from_string(NAVY)
meta = doc.add_paragraph("Cálculo I · Lista 3 · Ex. 5 · registrada por Bia, terça")
meta.runs[0].font.color.rgb = RGBColor.from_string("64748B")


def shade(cell, color):
    tc = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(dqn("w:val"), "clear")
    shd.set(dqn("w:fill"), color)
    tc.append(shd)


fields = [
    ("Enunciado", "Calcule ∫ x·eˣ dx."),
    ("Foto da folha", "[colar aqui a foto da resolução no papel]"),
    ("Passos", "1. Escolher u = x e dv = eˣ dx\n2. Aplicar ∫u dv = uv − ∫v du\n3. Somar a constante C"),
    ("Dica-chave", "Chame de u o que fica mais simples ao derivar."),
    ("Dúvida que ficou", "Nenhuma."),
]
table = doc.add_table(rows=0, cols=2)
table.style = "Table Grid"
table.autofit = False
for label, value in fields:
    row = table.add_row()
    a, b = row.cells
    a.width, b.width = Cm(4), Cm(9.6)
    a.text = label
    a.paragraphs[0].runs[0].bold = True
    a.paragraphs[0].runs[0].font.color.rgb = RGBColor.from_string(NAVY)
    shade(a, "EEF3F9")
    b.text = value
    if label == "Foto da folha":
        b.paragraphs[0].runs[0].italic = True
        b.paragraphs[0].runs[0].font.color.rgb = RGBColor.from_string("94A3B8")
        row.height = Cm(1.6)
for col, w in zip(table.columns, (Cm(3.8), Cm(9.8))):
    col.width = w
    for cell in col.cells:
        cell.width = w
foot = doc.add_paragraph()
foot.alignment = WD_ALIGN_PARAGRAPH.LEFT
foot.add_run("Ao salvar: colar o link na planilha-índice e atualizar “Onde paramos”.").italic = True
doc.save(OUT / "Ficha_L3_E5.docx")
print("salvo em", OUT)


# ---------------- Imagem da Versão B para o caderno ----------------
def compor_imagem(pdf_dir: Path, destino: Path) -> None:
    """Junta planilha, ficha e estrutura de pastas numa imagem 2520x1400.

    Requer pymupdf e Pillow; os PDFs vêm da conversão dos arquivos pelo LibreOffice.
    """
    import pymupdf
    from PIL import Image, ImageChops, ImageDraw, ImageFont

    def render(pdf):
        pix = pymupdf.open(pdf)[0].get_pixmap(dpi=220)
        im = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        bbox = ImageChops.difference(im, Image.new("RGB", im.size, "white")).getbbox()
        return im.crop((bbox[0] - 12, bbox[1] - 12, bbox[2] + 12, bbox[3] + 12))

    font = "/usr/share/fonts/truetype/liberation/LiberationSans-%s.ttf"
    bold = ImageFont.truetype(font % "Bold", 40)
    reg = ImageFont.truetype(font % "Regular", 34)
    small = ImageFont.truetype(font % "Regular", 30)
    navy, muted, line = (21, 54, 95), (100, 116, 139), (203, 213, 225)

    canvas = Image.new("RGB", (2520, 1400), "white")
    d = ImageDraw.Draw(canvas)

    def badge(x, y, n):
        d.ellipse((x, y, x + 48, y + 48), fill=navy)
        d.text((x + 24, y + 24), str(n), font=bold, fill="white", anchor="mm")

    # 1 · planilha-índice, largura total
    planilha = render(pdf_dir / "Indice_Calculo_I.pdf")
    planilha.thumbnail((2480, 640))
    badge(20, 10, 1)
    d.text((84, 14), "Planilha-índice: faz o papel da tela inicial", font=bold, fill=navy)
    canvas.paste(planilha, (20, 76))
    d.rectangle((20, 76, 20 + planilha.width, 76 + planilha.height), outline=line, width=3)

    # 2 · ficha-modelo, à direita embaixo
    y0 = 76 + planilha.height + 34
    ficha = render(pdf_dir / "Ficha_L3_E5.pdf")
    ficha.thumbnail((1420, 1400 - y0 - 76))
    fx = 2520 - ficha.width - 20
    badge(fx, y0, 2)
    d.text((fx + 64, y0 + 4), "Ficha-modelo: uma por exercício", font=bold, fill=navy)
    canvas.paste(ficha, (fx, y0 + 66))
    d.rectangle((fx, y0 + 66, fx + ficha.width, y0 + 66 + ficha.height), outline=line, width=3)

    # 3 · estrutura de pastas, à esquerda embaixo
    badge(20, y0, 3)
    d.text((84, y0 + 4), "Pasta compartilhada do grupo", font=bold, fill=navy)
    arvore = [
        ("Cálculo I", 0, True),
        ("Indice_Calculo_I.xlsx", 1, False),
        ("Lista 3", 1, True),
        ("Ficha_L3_E1.docx", 2, False),
        ("…  Ficha_L3_E5.docx", 2, False),
        ("Lista 2", 1, True),
    ]
    y = y0 + 80
    for nome, nivel, pasta in arvore:
        x = 40 + nivel * 60
        if pasta:
            d.rectangle((x, y + 8, x + 40, y + 36), fill=(226, 232, 240), outline=muted, width=2)
            d.rectangle((x, y + 2, x + 18, y + 10), fill=(226, 232, 240), outline=muted, width=2)
        else:
            d.rectangle((x + 6, y + 2, x + 34, y + 38), fill="white", outline=muted, width=2)
        d.text((x + 56, y + 2), nome, font=reg, fill=(30, 41, 59))
        y += 54
    d.text((40, y + 20), "Arquivos reais (.xlsx e .docx), sem programação.", font=small, fill=muted)
    canvas.save(destino)


if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1:
        compor_imagem(Path(sys.argv[1]), OUT.parent / "versao_b.png")
        print("imagem salva em", OUT.parent / "versao_b.png")
