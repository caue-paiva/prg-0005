"""Preenche uma cópia do caderno da Entrega 2 sem dependências externas."""

from __future__ import annotations

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
import xml.etree.ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT / "entregas/02-entrevistas-e-mapa-de-empatia"
SOURCE = BASE / "caderno.pptx"
OUTPUT = BASE / "Entrega_2_Collabora.AI_v2.pptx"

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"a": A, "p": P, "r": R}
ET.register_namespace("a", A)
ET.register_namespace("p", P)
ET.register_namespace("r", R)


def xml_bytes(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def replace_shape(
    root: ET.Element,
    shape_id: int,
    value: str,
    *,
    points: float | None = None,
    ink: bool = True,
) -> None:
    for shape in root.findall(".//p:sp", NS):
        props = shape.find("p:nvSpPr/p:cNvPr", NS)
        if props is None or props.get("id") != str(shape_id):
            continue
        texts = shape.findall(".//a:t", NS)
        if not texts:
            raise ValueError(f"Shape {shape_id} não tem texto")
        texts[0].text = value
        for text in texts[1:]:
            text.text = ""
        for style in shape.findall(".//a:rPr", NS):
            if points is not None:
                style.set("sz", str(round(points * 100)))
            if ink:
                style.set("i", "0")
                color = style.find("a:solidFill/a:srgbClr", NS)
                if color is not None:
                    color.set("val", "23334C")
        return
    raise ValueError(f"Shape {shape_id} não encontrado")


def add_header_text(
    root: ET.Element,
    shape_id: int,
    value: str,
    x: float,
    y: float,
    width: float,
    height: float,
    points: float,
    *,
    transparent: bool = False,
) -> None:
    """Adiciona texto legível sobre os campos pequenos da imagem do mapa."""
    tree = root.find("p:cSld/p:spTree", NS)
    assert tree is not None
    shape = ET.SubElement(tree, f"{{{P}}}sp")
    nv = ET.SubElement(shape, f"{{{P}}}nvSpPr")
    ET.SubElement(nv, f"{{{P}}}cNvPr", {"id": str(shape_id), "name": f"Metadado do mapa {shape_id}"})
    ET.SubElement(nv, f"{{{P}}}cNvSpPr")
    ET.SubElement(nv, f"{{{P}}}nvPr")
    props = ET.SubElement(shape, f"{{{P}}}spPr")
    trans = ET.SubElement(props, f"{{{A}}}xfrm")
    ET.SubElement(trans, f"{{{A}}}off", {"x": str(round(x * 914400)), "y": str(round(y * 914400))})
    ET.SubElement(trans, f"{{{A}}}ext", {"cx": str(round(width * 914400)), "cy": str(round(height * 914400))})
    geom = ET.SubElement(props, f"{{{A}}}prstGeom", {"prst": "rect"})
    ET.SubElement(geom, f"{{{A}}}avLst")
    if transparent:
        ET.SubElement(props, f"{{{A}}}noFill")
    else:
        fill = ET.SubElement(props, f"{{{A}}}solidFill")
        ET.SubElement(fill, f"{{{A}}}srgbClr", {"val": "FFFFFF"})
    line = ET.SubElement(props, f"{{{A}}}ln")
    ET.SubElement(line, f"{{{A}}}noFill")
    body = ET.SubElement(shape, f"{{{P}}}txBody")
    ET.SubElement(body, f"{{{A}}}bodyPr", {
        "lIns": "18288", "rIns": "18288", "tIns": "0", "bIns": "0", "anchor": "ctr"
    })
    ET.SubElement(body, f"{{{A}}}lstStyle")
    paragraph = ET.SubElement(body, f"{{{A}}}p")
    run = ET.SubElement(paragraph, f"{{{A}}}r")
    style = ET.SubElement(run, f"{{{A}}}rPr", {"lang": "pt-BR", "sz": str(round(points * 100))})
    text_fill = ET.SubElement(style, f"{{{A}}}solidFill")
    ET.SubElement(text_fill, f"{{{A}}}srgbClr", {"val": "174A90"})
    ET.SubElement(style, f"{{{A}}}latin", {"typeface": "Calibri"})
    ET.SubElement(run, f"{{{A}}}t").text = value


MAPS = [
    {
        "code": "E1",
        "author": "Cauê Paiva Lira",
        "date": "06/10/2026",
        "profile": "Computação · 2º sem.",
        "fields": {
            170: "Estudante de Computação no 2º sem. Estuda com amigos.",
            173: "Resolver listas e guardar dicas úteis para provas.",
            176: "Listas de cálculo em folhas soltas.",
            179: "Dicas surgem na prática e depois se perdem.",
            182: "Resolve exercícios com amigos na biblioteca.",
            185: "Ouve que colegas acertam exercícios, mas depois esquecem como refazê-los.",
            188: "Sente frustração: dicas perdidas poderiam ajudar nas provas.",
            195: "Dicas criadas pelo grupo se perdem antes das provas.",
        },
    },
    {
        "code": "E2",
        "author": "Enzo Tonon Morente",
        "date": "06/10/2026",
        "profile": "Computação · 2º sem.",
        "fields": {
            170: "Estudante de Computação no 2º sem. Faz resumos para o grupo.",
            173: "Explicar conteúdos com imagens, vídeos e contas.",
            176: "Usa Google Docs.",
            179: "O Docs limita vídeos, imagens e resoluções.",
            182: "Produz resumos para explicar aos colegas.",
            185: "Ouve que os resumos são bons, mas às vezes maçantes.",
            188: "Preparar resumos demora e a pessoa se sente cansada.",
            195: "O Docs dificulta registrar explicações com mídia e contas.",
        },
    },
    {
        "code": "E3",
        "author": "Letícia Barbosa Neves",
        "date": "06/10/2026",
        "profile": "Estatística · 2º sem.",
        "fields": {
            170: "Estudante de Estatística no 2º sem. Alterna estudo individual e em grupo.",
            173: "Entender a importância das matérias e seu uso profissional.",
            176: "Matérias parecem desconexas no início.",
            179: "Falta visão geral do semestre e da utilidade.",
            182: "Consulta veteranos sobre as matérias.",
            185: "Ouve que a carga inicial de disciplinas é pesada.",
            188: "Quer saber como o semestre a prepara para desafios.",
            195: "Sem uma visão geral da importância das matérias, pergunta a veteranos para que servem as matérias no contexto profissional.",
        },
    },
    {
        "code": "E4",
        "author": "Lázaro Pereira Vinaud Neto",
        "date": "07/10/2026",
        "profile": "SI noturno · 4º sem.",
        "fields": {
            170: "Trabalha e cursa SI à noite. Relata o 1º e o 2º semestres.",
            173: "Retomar matérias em intervalos de 30 a 60 min.",
            176: "Anotações e exercícios de várias matérias.",
            179: "Buscar materiais reduzia a produtividade.",
            182: "Estudava sozinho entre trabalho e aulas.",
            185: "Colegas: aprenderá mais no mercado que na faculdade.",
            188: "Sente-se perdido e cansado ao mudar de contexto.",
            195: "Localizar materiais consome parte do pouco tempo de estudo.",
        },
    },
    {
        "code": "E5",
        "author": "Vitor Thompson Borges",
        "date": "07/10/2026",
        "profile": "Eng. Comp. · 2º sem.",
        "fields": {
            170: "Estudante de Engenharia no 2º sem. Estuda em grupo.",
            173: "Aprender com colegas e manter um registro comum.",
            176: "Soluções em cadernos, arquivos e mensagens.",
            179: "Grupo grande sem material central.",
            182: "Estuda em grupo. Refaz soluções e repete dúvidas.",
            185: "Ouve que colegas acertam exercícios, mas depois esquecem como refazê-los.",
            188: "Grupo ajuda, mas conversas paralelas e desorganização atrapalham.",
            195: "Sem registro comum, o grupo perde conclusões e repete dúvidas.",
        },
    },
]


def fill_map(source: bytes, data: dict[str, object]) -> bytes:
    root = ET.fromstring(source)
    code = data["code"]
    replace_shape(root, 159, f"ENTREGA 2 · MAPA DE EMPATIA · {code}", ink=False)
    replace_shape(root, 160, f"Mapa de empatia — {code}", ink=False)
    replace_shape(root, 166, str(code), points=9.75, ink=False)
    replace_shape(root, 167, str(data["profile"]), points=8.1, ink=False)
    for shape_id, content in data["fields"].items():
        size = 8.2 if shape_id in (176, 179, 185, 188) else 8.8
        if shape_id == 195:
            size = 9.4 if code == "E3" else 10.6
        replace_shape(root, shape_id, content, points=size)
    replace_shape(root, 196, "Relatos em paráfrase. Necessidade e insight: interpretações da equipe.", points=9.0)
    add_header_text(root, 500, str(data["date"]), 7.30, 1.61, 0.55, 0.13, 6.7)
    add_header_text(root, 501, str(data["author"]).split()[0], 7.19, 1.785, 0.62, 0.19, 7.6, transparent=True)
    return xml_bytes(root)


def build() -> None:
    with ZipFile(SOURCE) as original:
        entries = {name: original.read(name) for name in original.namelist()}

    slide2 = ET.fromstring(entries["ppt/slides/slide2.xml"])
    replace_shape(slide2, 48, "E1, E2 e E5 estudam em grupo no 2º semestre. E3, também no 2º, alterna estudo individual e em grupo. E4 está no 4º e relata estudo individual no 1º e no 2º semestres.", points=9.8)
    replace_shape(slide2, 51, "Listas em papel (E1), resumos digitais (E2), estudo individual e em grupo (E3), blocos curtos no 1º e no 2º semestres (E4) e grupos grandes (E5).", points=10.4)
    replace_shape(slide2, 54, "Cauê fez E1, Enzo E2, Letícia E3, Lázaro E4 e Vitor E5. Os registros usam códigos. A equipe reuniu cinco mapas e uma síntese sem identificar entrevistados.", points=10.2)
    entries["ppt/slides/slide2.xml"] = xml_bytes(slide2)

    slide3 = ET.fromstring(entries["ppt/slides/slide3.xml"])
    replace_shape(slide3, 69, "Usem as perguntas como guia. Façam perguntas de aprofundamento quando a pessoa trouxer exemplos relevantes.", ink=False)
    replace_shape(slide3, 103, "Conte sobre uma dúvida recente ao estudar com colegas: como tentaram resolvê-la e o que aconteceu depois?", points=11.5)
    entries["ppt/slides/slide3.xml"] = xml_bytes(slide3)

    source_map = entries["ppt/slides/slide5.xml"]
    for index, data in enumerate(MAPS):
        number = 5 if index == 0 else 8 + index
        entries[f"ppt/slides/slide{number}.xml"] = fill_map(source_map, data)
        if index:
            entries[f"ppt/slides/_rels/slide{number}.xml.rels"] = entries["ppt/slides/_rels/slide5.xml.rels"]

    slide6 = ET.fromstring(entries["ppt/slides/slide6.xml"])
    replace_shape(slide6, 205, "Um padrão exige recorrência entre relatos. Identifiquem as interpretações da equipe.", ink=False)
    table = [
        (
            "E1, E2 e E5 buscam materiais úteis ao estudo em grupo.",
            "E3 busca a utilidade das matérias. E4 precisa retomar o estudo em tempo curto.",
            "E1: quer guardar dicas. E2: quer incluir mídia e contas. E5: falta material central.",
            "Investigar quais objetivos se repetem entre os estudantes.",
        ),
        (
            "E1, E2 e E5 relatam problemas com materiais.",
            "E1 perde dicas. E2 limitações do Docs. E5 registros dispersos. E3/E4 trazem outras dores.",
            "E1: dicas esquecidas. E2: Docs limita formatos. E5: materiais dispersos.",
            "Investigar como registrar e recuperar explicações.",
        ),
        (
            "E1, E2, E3 e E5 estudam com colegas.",
            "E3 alterna com estudo individual. E4 estudava sozinho em blocos curtos.",
            "E1: listas em grupo. E2: resumos ao grupo. E3: estudo misto. E4: estuda só. E5: grupo.",
            "Não generalizar os modos de estudar.",
        ),
        (
            "E1 e E2 usam registros próprios durante o estudo.",
            "E1 usa papel. E2 usa Docs. E3 consulta veteranos. E5 relata materiais dispersos.",
            "E1: folhas soltas. E2: Docs. E3: veteranos. E5: cadernos, arquivos e mensagens.",
            "Entender práticas atuais antes de definir ferramenta.",
        ),
        (
            "Não surgiu um desejo recorrente entre os relatos.",
            "E1 quer guardar dicas para provas. E3 busca a utilidade profissional das matérias.",
            "E1: quer rever dicas para provas. E3: busca uso profissional das matérias.",
            "Verificar se esses resultados importam a outros estudantes.",
        ),
    ]
    for row, content in enumerate(table):
        for col, value in enumerate(content):
            replace_shape(slide6, 222 + row * 10 + col * 2, value, points=8.4)
    replace_shape(slide6, 269, "Padrões e implicações são interpretações da equipe apoiadas nos relatos E1–E5.", points=10.5)
    entries["ppt/slides/slide6.xml"] = xml_bytes(slide6)

    slide7 = ET.fromstring(entries["ppt/slides/slide7.xml"])
    for shape_id in (283, 286, 289, 292, 295):
        replace_shape(slide7, shape_id, "☒", ink=False)
    replace_shape(slide7, 303, "Entrega_2_Collabora.AI.pdf", points=12)
    entries["ppt/slides/slide7.xml"] = xml_bytes(slide7)

    slide8 = ET.fromstring(entries["ppt/slides/slide8.xml"])
    members = [
        ("Cauê Paiva Lira", "14675416", "Conduziu E1."),
        ("Enzo Tonon Morente", "14568476", "Conduziu E2."),
        ("Letícia Barbosa Neves", "14588659", "Conduziu E3."),
        ("Lázaro Pereira Vinaud Neto", "14675396", "Conduziu E4."),
        ("Vitor Thompson Borges", "13682941", "Conduziu E5."),
    ]
    for row in range(7):
        values = members[row] if row < len(members) else ("", "", "")
        for col, value in enumerate(values):
            replace_shape(slide8, 324 + row * 6 + col * 2, value, points=10.3)
    replace_shape(slide8, 367, "☒", points=18, ink=False)
    replace_shape(slide8, 368, "IA generativa auxiliou na organização dos relatos, na redação dos mapas e na síntese das entrevistas.", points=10.7)
    entries["ppt/slides/slide8.xml"] = xml_bytes(slide8)

    # Mantém os namespaces padrão dos arquivos de controle do caderno. O
    # LibreOffice não carrega a cópia se ElementTree reserializa esses arquivos.
    presentation = entries["ppt/presentation.xml"].decode("utf-8")
    relations = entries["ppt/_rels/presentation.xml.rels"].decode("utf-8")
    types = entries["[Content_Types].xml"].decode("utf-8")
    extra_ids = "".join(
        f'<p:sldId id="{264 + index}" r:id="rId{18 + index}"/>'
        for index in range(4)
    )
    presentation = presentation.replace(
        '<p:sldId id="260" r:id="rId6"/>',
        '<p:sldId id="260" r:id="rId6"/>' + extra_ids,
    )
    for index, number in enumerate(range(9, 13)):
        relations = relations.replace(
            "</Relationships>",
            f'<Relationship Id="rId{18 + index}" Type="{R}/slide" '
            f'Target="slides/slide{number}.xml"/></Relationships>',
        )
        types = types.replace(
            "</Types>",
            f'<Override PartName="/ppt/slides/slide{number}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
            "</Types>",
        )
    entries["ppt/presentation.xml"] = presentation.encode("utf-8")
    entries["ppt/_rels/presentation.xml.rels"] = relations.encode("utf-8")
    entries["[Content_Types].xml"] = types.encode("utf-8")

    with ZipFile(OUTPUT, "w", compression=ZIP_DEFLATED) as result:
        for name, contents in entries.items():
            result.writestr(name, contents)
    print(OUTPUT)


if __name__ == "__main__":
    build()
