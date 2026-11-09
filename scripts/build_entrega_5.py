"""Preenche uma cópia do caderno da Entrega 5 (V1). Requer python-pptx.

As imagens das versões vêm de entregas/05-prototipar-e-comparar/prototipos/:
versao_a.png (renderizada de versao_a.html no Google Chrome) e versao_b.png
(gerada por scripts/build_prototipo_b.py).
"""

import copy
from pathlib import Path

from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Emu

ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT / "entregas/05-prototipar-e-comparar"
SOURCE = BASE / "caderno.pptx"
OUTPUT = BASE / "Entrega_5_Collabora.AI_v1.pptx"
IMG_A = BASE / "prototipos/versao_a.png"
IMG_B = BASE / "prototipos/versao_b.png"
INK = "23334C"

prs = Presentation(SOURCE)
S = list(prs.slides)


def shape(slide, sid):
    for sh in slide.shapes:
        if sh.name.split(";")[1] == str(sid):
            return sh
    raise KeyError(sid)


def set_text(sh, paras, size=None, ink=True, space_after=None):
    """paras: lista de str ou de [(texto, negrito), ...]."""
    body = sh.text_frame._txBody
    ps = body.findall(qn("a:p"))
    p0 = ps[0]
    r0 = p0.find(qn("a:r"))
    rpr = copy.deepcopy(r0.find(qn("a:rPr"))) if r0 is not None else None
    if rpr is None:
        rpr = copy.deepcopy(p0.find(qn("a:endParaRPr")))
        rpr.tag = qn("a:rPr")
    ppr = p0.find(qn("a:pPr"))
    for p in ps:
        body.remove(p)
    if ink:
        rpr.set("i", "0")
        fill = rpr.find(qn("a:solidFill"))
        if fill is not None:
            fill[0].set("val", INK)
    if size:
        rpr.set("sz", str(round(size * 100)))
    for para in paras:
        if isinstance(para, str):
            para = [(para, None)]
        p = body.makeelement(qn("a:p"), {})
        if ppr is not None:
            pp = copy.deepcopy(ppr)
            if space_after is not None:
                sa = pp.find(qn("a:spcAft"))
                if sa is not None:
                    sa[0].set("val", str(round(space_after * 100)))
            p.append(pp)
        for text, bold in para:
            r = p.makeelement(qn("a:r"), {})
            rp = copy.deepcopy(rpr)
            if bold is not None:
                rp.set("b", "1" if bold else "0")
            r.append(rp)
            t = r.makeelement(qn("a:t"), {})
            t.text = text
            r.append(t)
            p.append(r)
        body.append(p)


def geom(sh, top=None, height=None, left=None, width=None):
    for attr, val in (("top", top), ("height", height), ("left", left), ("width", width)):
        if val is not None:
            setattr(sh, attr, Emu(int(val)))


def anchor(sh, where="t"):
    sh.text_frame._txBody.find(qn("a:bodyPr")).set("anchor", where)


def remove(sh):
    sh._element.getparent().remove(sh._element)


def put_picture(slide, frame, path, pad=60000):
    """Insere a imagem centralizada dentro do retângulo `frame`, mantendo a proporção."""
    from PIL import Image

    w, h = Image.open(path).size
    fw, fh = frame.width - 2 * pad, frame.height - 2 * pad
    scale = min(fw / w, fh / h)
    pw, ph = int(w * scale), int(h * scale)
    slide.shapes.add_picture(str(path), Emu(frame.left + pad + (fw - pw) // 2),
                             Emu(frame.top + pad + (fh - ph) // 2), Emu(pw), Emu(ph))


# ---------------- Slide 38: planejamento ----------------
s = S[1]
g = lambda i: shape(s, i)
plano = {
    49: "Alex e estudantes do 1º e 2º semestres que estudam com colegas. Usam o caderno ao fim de uma sessão de "
        "listas, para registrar, e ao voltar a estudar ou na semana da prova, para retomar.",
    52: "Registrar a resolução de um exercício (foto da folha, passos e dica-chave) e, em outra sessão, ver onde "
        "o grupo parou e reencontrar um exercício já resolvido.",
    55: "Reencontrar rápido o que o grupo resolveu e explicou, sem refazer exercícios nem procurar em folhas, "
        "arquivos e mensagens (proposta de valor da Entrega 4).",
    58: "Média na versão A (telas com conteúdo realista, sem funcionamento) e baixa na B (arquivos-modelo). Em uma "
        "semana, o objetivo é comparar formas de mostrar as duas funções essenciais, não o acabamento visual.",
}
for sid, txt in plano.items():
    set_text(g(sid), [txt], size=10)
    anchor(g(sid))
for box, txt in ((48, 49), (51, 52), (54, 55), (57, 58)):
    geom(g(box), height=830000)
    geom(g(txt), height=740000)
for lab in (53, 56):
    geom(g(lab), top=3220000)
for box, txt in ((54, 55), (57, 58)):
    geom(g(box), top=3458000)
    geom(g(txt), top=3500000)
set_text(g(61), ["• Indicação de onde o grupo parou",
                 "• Lista de exercícios com situação: resolvido, com dúvida, pendente",
                 "• Registro de um exercício com foto, passos, dica-chave e dúvida",
                 "• Busca por disciplina, lista ou termo, e quem registrou cada resolução"], size=10)
set_text(g(64), ["Registrar um exercício leva menos de dois minutos. Foto, passos e dica bastam para refazê-lo "
                 "depois. Ver onde o grupo parou ajuda a retomar. Um registro comum evita dúvidas repetidas."], size=9.5)
for lab in (59, 62):
    geom(g(lab), top=4440000)
for box in (60, 63):
    geom(g(box), top=4678000, height=1020000)
for txt in (61, 64):
    geom(g(txt), top=4730000, height=930000)
    anchor(g(txt))
geom(g(65), top=5860000, height=400000)
geom(g(66), top=5950000)


# ---------------- Slides 39 e 40: versões A e B ----------------
def versao(slide, base, img, campos):
    g = lambda i: shape(slide, i)
    put_picture(slide, g(base), img)
    remove(g(base + 1))
    remove(g(base + 2))
    for sid, (txt, size) in campos.items():
        set_text(g(base + sid), [txt], size=size)
        anchor(g(base + sid))
    # campo "o que demonstra" mais alto
    geom(g(base + 10), height=1250000)
    geom(g(base + 11), height=1160000)


versao(S[2], 80, IMG_A, {
    5: ("HTML e CSS, renderizados no Google Chrome", 10.5),
    8: ("Wireframe de aplicativo de celular, média fidelidade (3 telas)", 10.5),
    11: ("As duas funções essenciais da Entrega 4: retomar pela tela inicial (onde paramos e situação da "
         "lista), registrar um exercício e reencontrar uma dica pela busca. Nomes e exercícios são exemplos.", 9.5),
    14: ("Aplicativo próprio, organizado por disciplina → lista → exercício. “Onde paramos” é a primeira coisa que "
         "aparece. O registro parte da foto da folha, para aceitar papel e contas. Cada exercício tem situação.", 9.5),
})
versao(S[3], 108, IMG_B, {
    5: ("LibreOffice Calc e LibreOffice Writer (arquivos .xlsx e .docx, abertos também no Google Planilhas e no Google Docs)", 9.5),
    8: ("Protótipo sem código, baixa fidelidade: planilha-índice, ficha-modelo e pastas", 10),
    11: ("O mesmo caderno montado com arquivos comuns: a planilha faz o papel da tela inicial (onde paramos e "
         "situação) e a ficha registra um exercício com foto, passos e dica.", 9.5),
    14: ("Nenhum aplicativo novo: tudo fica numa pasta compartilhada. O índice e o “onde paramos” são atualizados "
         "à mão. A busca se limita ao nome dos arquivos e ao texto da planilha.", 9.5),
})
for slide, base in ((S[2], 80), (S[3], 108)):
    g = lambda i: shape(slide, i)
    geom(g(base + 12), top=5720000)
    geom(g(base + 13), top=5958000, height=440000)
    geom(g(base + 14), top=5995000, height=370000)

# ---------------- Slide 41: comparação ----------------
s = S[4]
g = lambda i: shape(s, i)
linhas = [
    ("App de celular com 3 telas; “onde paramos” em destaque na abertura.",
     "Planilha, ficha em documento e pastas; “onde paramos” é uma linha no topo da planilha.",
     "Em A, a pessoa vê em segundos onde o grupo está. Em B, precisa abrir a planilha e ler a tabela."),
    ("Abrir → continuar daqui → registrar em uma tela → buscar por termo.",
     "Abrir a pasta → copiar a ficha → preencher → colar o link na planilha → atualizar “onde paramos”.",
     "B tem dois passos manuais a mais por exercício. Se alguém esquecer, o índice fica desatualizado e o trabalho "
     "tende a cair em quem organiza (E2)."),
    ("Registro com foto, passos, dica e dúvida; situação por exercício; onde paramos; busca no conteúdo.",
     "Mesma ficha e situação; onde paramos manual; busca só no nome dos arquivos e na planilha.",
     "B retira a busca no conteúdo e a atualização automática, justamente o que serve para reencontrar e retomar."),
    ("Exige programar um app ou site e guardar fotos num servidor: semanas de trabalho e algum custo.",
     "Montada em uma tarde, com ferramentas gratuitas que a turma já usa.",
     "B é mais viável hoje. A depende de um desenvolvimento que a equipe ainda precisa dimensionar."),
    ("Cumpre os três critérios “deve” da Entrega 3 e o “retomar rápido”.",
     "Cumpre registro comum e caminho da resolução; reencontrar fica lento e repete o limite de documentos para contas e imagens (E2).",
     "A responde melhor aos critérios. B arrisca virar mais um lugar onde o material se espalha (E4, E5)."),
]
top, bottom = 2400300, 5900000
h = (bottom - top) / len(linhas)
for k, cells in enumerate(linhas):
    base = 144 + 8 * k
    y = top + k * h
    for j in range(4):
        geom(g(base + 2 * j), top=y, height=h)
        geom(g(base + 2 * j + 1), top=y + 45000, height=h - 90000)
        anchor(g(base + 2 * j + 1), "ctr")
    for j, txt in enumerate(cells, start=1):
        set_text(g(base + 2 * j + 1), [txt], size=8.5)
geom(g(184), top=5975000)
geom(g(185), top=6195000, height=320000)
geom(g(186), top=6225000, height=260000)
set_text(g(186), ["Retomar e reencontrar sem trabalho manual: só a A mostra onde o grupo parou e busca dentro dos "
                  "registros (função essencial 2 da Entrega 4)."], size=9.5)

# ---------------- Slide 42: decisão final ----------------
s = S[5]
g = lambda i: shape(s, i)
put_picture(s, g(200), IMG_A, pad=50000)
remove(g(201))
set_text(g(204), ["☒ A   ☐ B   ☐ versão combinada"], ink=False)
set_text(g(207), ["HTML e CSS, renderizados no Google Chrome"], size=10.5)
set_text(g(210), ["Só a versão A retoma e reencontra sem passos manuais, que é o benefício central da proposta de valor. "
                  "Ela cumpre os três critérios “deve” da Entrega 3 e responde à frustração da persona Alex de não achar "
                  "mais como fez. A B é mais barata, mas depende de atualizar o índice à mão."], size=9.5)
set_text(g(213), ["Em três telas, mostra o caminho completo: abrir o caderno, continuar de onde o grupo parou, registrar "
                  "um exercício com foto e dica e reencontrá-lo depois pela busca."], size=10)
set_text(g(216), ["Telas estáticas, sem salvar nem buscar de verdade, com dados de exemplo. Não foi testado com estudantes. "
                  "Custo de construir e de guardar fotos segue incerto. Fotos e nomes de estudantes exigem cuidado com a LGPD."],
         size=9.5)
for sid in (210, 213, 216):
    anchor(g(sid))
# coluna direita mais compacta, motivo mais alto
geom(g(205), top=2800000)
geom(g(206), top=3038000, height=560000)
geom(g(207), top=3090000, height=470000)
geom(g(208), top=3720000)
geom(g(209), top=3958000, height=1130000)
geom(g(210), top=4010000, height=1040000)
# linha de baixo mais alta
for lab in (211, 214):
    geom(g(lab), top=5150000)
for box in (212, 215):
    geom(g(box), top=5388000, height=860000)
for txt in (213, 216):
    geom(g(txt), top=5440000, height=780000)

# ---------------- Slide 43: revisão ----------------
s = S[6]
for cid in (231, 234, 237, 240, 243):
    set_text(shape(s, cid), ["☒"], ink=False)
set_text(shape(s, 251), ["Entrega_5_Collabora.AI.pdf"], ink=False)

# ---------------- Slide 44: participação e IA ----------------
s = S[7]
g = lambda i: shape(s, i)
team = [("Cauê Paiva Lira", "14675416"), ("Enzo Tonon Morente", "14568476"),
        ("Letícia Barbosa Neves", "14588659"), ("Lázaro Pereira Vinaud Neto", "14675396"),
        ("Vitor Thompson Borges", "13682941")]
for k, (name, nusp) in enumerate(team):
    set_text(g(272 + 6 * k), [name])
    set_text(g(274 + 6 * k), [nusp])
    set_text(g(276 + 6 * k), ["TODO: descrever a contribuição nesta entrega."])
for k in range(len(team), 7):
    for off in (272, 274, 276):
        set_text(g(off + 6 * k), [""])
set_text(g(315), ["☒"], ink=False)
set_text(g(316), ["Este trabalho incorpora o uso de IA generativa (Claude, da Anthropic) para auxiliar no plano do protótipo, "
                  "na construção das telas da versão A e dos arquivos da versão B, na comparação entre as versões e na "
                  "montagem do arquivo. O conteúdo final foi revisado criticamente, editado e reflete a compreensão e "
                  "análise da equipe."], size=10)

prs.save(OUTPUT)
print("salvo", OUTPUT)
