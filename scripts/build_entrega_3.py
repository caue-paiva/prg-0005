"""Preenche uma cópia do caderno da Entrega 3 (V2). Requer python-pptx."""

import copy
from pathlib import Path

from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Emu

ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT / "entregas/03-persona-e-desafio-refinado"
SOURCE = BASE / "caderno.pptx"
OUTPUT = BASE / "Entrega_3_Collabora.AI_v2.pptx"
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


def geom(sh, top=None, height=None):
    if top is not None:
        sh.top = Emu(int(top))
    if height is not None:
        sh.height = Emu(int(height))


def anchor_top(sh):
    sh.text_frame._txBody.find(qn("a:bodyPr")).set("anchor", "t")


# ---------------- Slide 21: persona ----------------
s = S[2]
g = lambda i: shape(s, i)
set_text(g(69), ["Origem dos traços"], ink=False)
set_text(g(70), ["Síntese de E1, E2, E3 e E5. Os códigos indicam de qual entrevista vem cada traço. "
                  "Quem estuda sozinho com pouco tempo (E4) fica como perfil secundário."], size=10)
set_text(g(73), ["Alex"], size=11)
set_text(g(76), ["19 anos"], size=10)
set_text(g(79), ["2º semestre da graduação na USP, cursando as disciplinas iniciais do curso (E1, E2, E3, E5)"], size=8.5)
for box, txt in ((78, 79), (75, 76), (72, 73)):
    geom(g(box), height=640000)
    geom(g(txt), top=2190000, height=560000)

persona = {
    82: ["Dar conta das disciplinas iniciais, que considera pesadas (E3, E5). Aprender com colegas que cursam as mesmas "
         "disciplinas (E1, E5). Entender para que servem as matérias (E3)."],
    85: ["Chegar à prova sabendo refazer os exercícios que o grupo resolveu (E1, E5). Retomar com o grupo o ponto "
         "em que pararam sem refazer o que já foi resolvido (E5)."],
    88: ["Resolve listas com amigos na biblioteca, em folhas soltas (E1). Faz ou recebe resumos no Google Docs (E2). "
         "Guarda anotações em cadernos, arquivos e mensagens (E4, E5). Pergunta a veteranos para que serve uma matéria (E3)."],
    91: ["Dicas e resoluções do grupo se perdem até a prova (E1). Cada colega segue um material e dúvidas já resolvidas "
         "voltam (E5). Perde tempo procurando a versão usada pelo grupo (E4, E5). Montar resumos com contas e imagens cansa (E2)."],
    94: ["Gosta de estudar com colegas e de explicar o que entendeu (E2, E5). Às vezes estuda em grupos grandes da "
         "mesma turma, com conversas paralelas que atrapalham acompanhar um exercício (E5). Também estuda sozinho em "
         "parte da semana (E3)."],
}
for sid, txt in persona.items():
    set_text(g(sid), txt, size=9.5)
    anchor_top(g(sid))
# linha motivações/objetivos
geom(g(80), top=2860000); geom(g(83), top=2860000)
for box, txt in ((81, 82), (84, 85)):
    geom(g(box), top=3090000, height=900000)
    geom(g(txt), top=3150000, height=800000)
# linha comportamentos/problemas/particularidades
for lab in (86, 89, 92):
    geom(g(lab), top=4090000)
for box, txt in ((87, 88), (90, 91), (93, 94)):
    geom(g(box), top=4320000, height=1150000)
    geom(g(txt), top=4370000, height=1060000)
geom(g(95), top=5570000)
geom(g(96), top=5800000, height=560000)
geom(g(97), top=5850000, height=460000)
set_text(g(97), ["“A gente resolve a lista junto e entende na hora, mas quando chega a prova ninguém acha mais as dicas nem lembra como fez.”"], size=11)

# ---------------- Slide 22: confronto ----------------
s = S[3]
g = lambda i: shape(s, i)
rows = [
    ("Estudantes universitários do 1º e 2º semestres que estudam em grupo para provas e atividades.",
     "Cinco estudantes de exatas e computação da USP (ICMC e Poli). E1, E2, E3 e E5 no 2º semestre. E4 relembra o 1º e o 2º. Ninguém do 1º semestre.",
     "Manter universitários do 1º e 2º semestres, com foco em quem estuda em grupo. Ouvir outras áreas e o 1º semestre segue como dúvida. E4 fica como perfil secundário."),
    ("Transformar o tempo de estudo em grupo em compreensão real do conteúdo.",
     "Guardar e recuperar o que foi produzido: dicas (E1), explicações e resoluções do grupo (E5), materiais em pouco tempo (E4).",
     "Alterar. A compreensão segue como objetivo, mas a necessidade que se repetiu aparece depois da sessão: reencontrar o que foi construído (E1, E4, E5)."),
    ("Níveis de preparo diferentes, pouca prática em organizar o estudo e dúvidas que o grupo não resolve no momento.",
     "Conhecimento disperso em folhas, cadernos, arquivos e mensagens (E1, E4, E5). Nenhum relato citou diferença de nível ou dúvida sem resposta.",
     "Alterar para dispersão e perda do que o grupo resolveu. Diferença de nível e dúvidas sem resposta voltam a ser dúvidas a investigar."),
    ("Mensagens, chamadas, biblioteca, monitorias, vídeos, IA usada sozinho, veteranos (percepção da equipe).",
     "Folhas soltas (E1), Google Docs (E2), veteranos (E3), pastas e arquivos (E4), cadernos e mensagens (E5). Monitoria e IA não apareceram.",
     "Já usam ferramentas gerais. Falta um registro comum e fácil de reencontrar, que conviva com papel, contas e imagens (E1, E2)."),
    ("Sessões dispersas, estudo na véspera, dúvidas acumuladas, desempenho baixo, desmotivação e ansiedade.",
     "Frustração com dicas perdidas (E1). Exercícios que não conseguem refazer (E1, E5). Dúvidas repetidas e retrabalho (E5). Tempo gasto em busca e cansaço (E2, E4).",
     "Manter o risco para as provas (E1). Retirar estudo na véspera e ansiedade, que não apareceram. Incluir retrabalho e tempo perdido."),
]
top, bottom = 2400300, 6020000
h = (bottom - top) / len(rows)
for k, cells in enumerate(rows):
    base = 118 + 8 * k
    y = top + k * h
    for j in range(4):
        geom(g(base + 2 * j), top=y, height=h)
        geom(g(base + 2 * j + 1), top=y + 50000, height=h - 100000)
    for j, txt in enumerate(cells, start=1):
        set_text(g(base + 2 * j + 1), [txt], size=8.5)
geom(g(158), top=6110000)
set_text(g(158), ["Coluna Entrega 2: relatos, citados por código. Coluna Decisão: interpretação da equipe. Pivotamos a dificuldade e mantivemos o objetivo."], ink=False)

# ---------------- Slide 23: declaração refinada ----------------
s = S[4]
g = lambda i: shape(s, i)
set_text(g(175), ["Estudantes universitários do 1º e 2º semestres que estudam com colegas para provas e atividades."], size=10)
set_text(g(178), ["Guardar e reencontrar as dicas, explicações e resoluções construídas durante o estudo."], size=10)
set_text(g(181), ["Esse conhecimento fica espalhado em folhas, cadernos, arquivos e mensagens e se perde entre uma sessão e outra."], size=10)
set_text(g(184), ["Estudantes universitários do 1º e 2º semestres que estudam com colegas para provas e atividades "
                  "precisam de uma forma de guardar e reencontrar as dicas, explicações e resoluções construídas durante o estudo, "
                  "porque esse conhecimento fica espalhado em folhas, cadernos, arquivos e mensagens e se perde entre uma sessão e outra, "
                  "o que leva a dúvidas repetidas e a exercícios que não conseguem refazer na revisão para as provas."], size=11)
geom(g(185), top=5470000)
geom(g(186), top=5700000, height=640000)
geom(g(187), top=5740000, height=560000)
set_text(g(187), ["E1: dicas em folhas soltas se perdem. E5: registros dispersos, dúvidas repetidas e perda do ponto em que pararam. "
                  "E4: procurar materiais consome o tempo de estudo. E2: registrar contas e imagens exige esforço. "
                  "Confirma a implicação da síntese da Entrega 2: recuperar conhecimento como necessidade central. "
                  "Hipótese mais arriscada: a dor ser forte o bastante para mudar o jeito de estudar."], size=9.5)

# ---------------- Slide 24: pergunta e critérios ----------------
s = S[5]
g = lambda i: shape(s, i)
set_text(g(204), ["Como poderíamos ajudar estudantes universitários do 1º e 2º semestres a retomar e reaproveitar "
                  "o que aprenderam estudando com colegas, diante de explicações e resoluções que hoje ficam espalhadas e se perdem "
                  "entre as sessões de estudo?"], size=11)
criteria = {
    208: ["• Permitir reencontrar depois o que o grupo resolveu e explicou (E1, E5).",
          "• Deixar claro qual é o registro comum do grupo (E5).",
          "• Preservar o caminho da resolução, não só a resposta (E1, E5)."],
    211: ["• Permitir retomar o estudo rápido, mesmo em intervalos curtos (E4, E5).",
          "• Aceitar contas, desenhos e o que hoje fica no papel (E1, E2).",
          "• Servir também a quem estuda sozinho (E3, E4)."],
    214: ["• Virar mais um lugar onde o material se espalha (E4, E5).",
          "• Concentrar o trabalho de organizar em uma só pessoa (E2).",
          "• Depender de que todos usem a mesma ferramenta para participar (E1, E2, E5)."],
}
for box, txt in ((207, 208), (210, 211), (213, 214)):
    geom(g(box), height=1550000)
    geom(g(txt), height=1420000)
    set_text(g(txt), criteria[txt], size=10, space_after=4)
    anchor_top(g(txt))

# ---------------- Slide 25: revisão ----------------
s = S[6]
set_text(shape(s, 248), ["Entrega_3_Collabora.AI.pdf"], ink=False)
for cid in (228, 231, 234, 237, 240):
    set_text(shape(s, cid), ["☒"], ink=False)

# ---------------- Slide 26: participação e IA ----------------
s = S[7]
g = lambda i: shape(s, i)
team = [("Cauê Paiva Lira", "14675416"), ("Enzo Tonon Morente", "14568476"),
        ("Letícia Barbosa Neves", "14588659"), ("Lázaro Pereira Vinaud Neto", "14675396"),
        ("Vitor Thompson Borges", "13682941")]
for k, (name, nusp) in enumerate(team):
    set_text(g(269 + 6 * k), [name])
    set_text(g(271 + 6 * k), [nusp])
    set_text(g(273 + 6 * k), ["TODO: descrever a contribuição nesta entrega."])
for k in range(len(team), 7):
    for off in (269, 271, 273):
        set_text(g(off + 6 * k), [""])
set_text(g(312), ["☒"], ink=False)
set_text(g(313), ["Este trabalho incorpora o uso de IA generativa (Claude, da Anthropic) para auxiliar na comparação entre as "
                  "Entregas 1 e 2, no rascunho da persona, do confronto, da declaração, da pergunta e dos critérios, e na montagem "
                  "do arquivo. O conteúdo final foi revisado criticamente, editado e reflete a compreensão e análise da equipe."], size=10)

# ---------------- Slide 22A: matriz CSD atualizada (cópia do slide 20) ----------------
src = S[1]
csd = prs.slides.add_slide(src.slide_layout)
for ph in list(csd.placeholders):
    ph._element.getparent().remove(ph._element)
for el in src.shapes._spTree:
    if el.tag == qn("p:sp"):
        csd.shapes._spTree.append(copy.deepcopy(el))
c = lambda i: shape(csd, i)
next_id = 900
for sid in (50, 51, 52, 53):  # terceiro cartão a partir do segundo
    el = copy.deepcopy(c(sid)._element)
    cnv = el.find(qn("p:nvSpPr")).find(qn("p:cNvPr"))
    cnv.set("id", str(next_id)); cnv.set("name", f"Google Shape;{sid + 1000};p2")
    next_id += 1
    csd.shapes._spTree.append(el)
W = (10820400 - 2 * 228600) // 3
cards = [(46, 47, 48, 49), (50, 51, 52, 53), (1050, 1051, 1052, 1053)]
for k, (bg, bar, title, body) in enumerate(cards):
    x = 685800 + k * (W + 228600)
    for sid, dx, w in ((bg, 0, W), (bar, 0, 66675), (title, 266700, W - 400000), (body, 266700, W - 400000)):
        sh = c(sid); sh.left = Emu(x + dx); sh.width = Emu(w)
    geom(c(bg), top=1950000, height=3330000); geom(c(bar), top=1950000, height=3330000)
    geom(c(title), top=2100000); geom(c(body), top=2500000, height=2700000)
set_text(c(39), ["ENTREGA 3 · ANÁLISE"], ink=False)
set_text(c(40), ["Matriz CSD atualizada com as entrevistas"], ink=False)
set_text(c(42), ["Revisão da matriz CSD da Entrega 1: o que as entrevistas confirmaram, o que segue sem evidência e o que ainda precisamos descobrir."], ink=False)
set_text(c(45), ["22A"], ink=False)
csd_txt = [
    ("Certezas", ["• O material do estudo se espalha em folhas, cadernos, arquivos e mensagens (E1, E4, E5).",
                  "• Dicas e resoluções se perdem e colegas não conseguem refazer exercícios (E1, E5).",
                  "• Estudar com colegas ajuda em disciplinas iniciais consideradas pesadas (E3, E5)."]),
    ("Suposições", ["• A diferença de nível atrapalha o grupo (nenhum relato).",
                    "• Dúvidas ficam sem resposta durante o estudo (nenhum relato).",
                    "• Muitos preferem colegas à monitoria (monitoria não apareceu).",
                    "• Grupos grandes perdem tempo com coordenação (só E5)."]),
    ("Dúvidas", ["• O padrão se repete no 1º semestre e em cursos fora de exatas e computação?",
                 "• Como os estudantes usam IA no estudo hoje?",
                 "• A falta de visão geral das disciplinas (E3) é outro problema?",
                 "• O que o grupo já tentou para guardar o que resolve?"]),
]
for (bg, bar, title, body), (tt, items) in zip(cards, csd_txt):
    set_text(c(title), [tt], ink=False)
    set_text(c(body), items, size=11, space_after=6)
geom(c(54), top=5450000, height=700000); geom(c(55), top=5560000, height=480000)
set_text(c(55), [[("Hipótese mais arriscada: ", True),
                  ("perder o que o grupo resolveu incomoda o bastante para os estudantes mudarem o jeito de estudar.", None)]], size=12)
ids = prs.slides._sldIdLst
el = ids[-1]; ids.remove(el); ids.insert(4, el)

prs.save(OUTPUT)
print("salvo", OUTPUT)
