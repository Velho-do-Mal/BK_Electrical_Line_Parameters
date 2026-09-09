"""
equations.py — BK Electrical Line Parameters

Construtor de equações NATIVAS do Word (OMML — Office Math Markup Language),
o mesmo formato usado pelo recurso "Inserir Equação" do Word. Não depende de
imagens nem de fórmulas em texto simples: o engenheiro pode clicar na
equação dentro do Word e editá-la com o editor de equações nativo.

python-docx não tem API para matemática, então este módulo monta o XML OMML
diretamente (mesma técnica recomendada pela skill "docx" para conteúdo que
python-docx não cobre) e injeta no parágrafo via oxml.

Nenhuma lógica de cálculo de engenharia mora aqui — apenas apresentação
tipográfica das fórmulas já usadas em calculos.py / relatorio.py.
"""

from __future__ import annotations

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

_XML_ESCAPE = {"&": "&amp;", "<": "&lt;", ">": "&gt;"}


def _esc(text: str) -> str:
    for a, b in _XML_ESCAPE.items():
        text = text.replace(a, b)
    return text


# ---------------------------------------------------------------------------
# Blocos OMML básicos (cada função retorna um fragmento XML <m:...>)
# ---------------------------------------------------------------------------

def run(text: str, italic: bool = True) -> str:
    """Texto simples dentro da equação (variável ou operador)."""
    sty = "" if italic else '<m:sty m:val="p"/>'
    return (
        f'<m:r><m:rPr>{sty}</m:rPr>'
        f'<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>'
        f'<m:t xml:space="preserve">{_esc(text)}</m:t></m:r>'
    )


def frac(num: str, den: str) -> str:
    return f'<m:f><m:fPr/><m:num>{num}</m:num><m:den>{den}</m:den></m:f>'


def sup(base: str, exp: str) -> str:
    return f'<m:sSup><m:sSupPr/><m:e>{base}</m:e><m:sup>{exp}</m:sup></m:sSup>'


def sub(base: str, subscript: str) -> str:
    return f'<m:sSub><m:sSubPr/><m:e>{base}</m:e><m:sub>{subscript}</m:sub></m:sSub>'


def subsup(base: str, subscript: str, exp: str) -> str:
    return (f'<m:sSubSup><m:sSubSupPr/><m:e>{base}</m:e>'
            f'<m:sub>{subscript}</m:sub><m:sup>{exp}</m:sup></m:sSubSup>')


def sqrt(content: str) -> str:
    return f'<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>{content}</m:e></m:rad>'


def nthroot(content: str, degree: str) -> str:
    return f'<m:rad><m:radPr/><m:deg>{degree}</m:deg><m:e>{content}</m:e></m:rad>'


def group(*parts: str) -> str:
    """Concatena fragmentos em sequência (ex.: vários m:r/m:f seguidos)."""
    return "".join(parts)


def paren(*parts: str) -> str:
    return group(run("("), *parts, run(")"))


def txt(*fragments: str, italic: bool = True) -> str:
    """Atalho: concatena vários textos simples em runs (útil p/ nomes com
    vários caracteres, ex.: 'sen', 'ln', 'cosh')."""
    return group(*[run(f, italic=italic) for f in fragments])


# ---------------------------------------------------------------------------
# Montagem final e inserção no documento
# ---------------------------------------------------------------------------

def _omath_para_xml(equation_xml: str) -> str:
    return (
        f'<m:oMathPara xmlns:m="{M_NS}" xmlns:w="{W_NS}">'
        f'<m:oMath>{equation_xml}</m:oMath></m:oMathPara>'
    )


def add_equation(document, equation_xml: str, center: bool = True):
    """Insere uma equação OMML como parágrafo próprio (estilo "display",
    igual ao Word ao inserir uma nova equação em linha própria)."""
    p = document.add_paragraph()
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    element = parse_xml(_omath_para_xml(equation_xml))
    p._p.append(element)
    return p


# ---------------------------------------------------------------------------
# Fórmulas específicas do modelo de linha longa (BK Electrical Line Parameters)
# Cada função devolve o fragmento OMML pronto para add_equation().
# ---------------------------------------------------------------------------

def eq_deq():
    # Deq = (DAB . DBC . DCA)^(1/3)
    base = paren(sub(run("D"), run("AB", italic=False)), run(" . "),
                 sub(run("D"), run("BC", italic=False)), run(" . "),
                 sub(run("D"), run("CA", italic=False)))
    return group(sub(run("D"), run("eq", italic=False)), run(" = "), sup(base, frac(run("1"), run("3"))))


def eq_req_duplo():
    base = paren(run("r"), run(" . "), run("d"))
    return group(sub(run("r"), run("eq", italic=False)), run(" = "), sup(base, frac(run("1"), run("2"))))


def eq_req_quadruplo():
    base = paren(run("r"), run(" . "), sup(run("d"), run("3")), run(" . "), sqrt(run("2")))
    return group(sub(run("r"), run("eq", italic=False)), run(" = "), sup(base, frac(run("1"), run("4"))))


def eq_resistencia():
    return group(run("R"), run(" = "), frac(sub(run("R"), run("cat", italic=False)), run("n")))


def eq_indutancia():
    return group(
        run("L"), run(" = 2×"), sup(run("10"), run("-4")), run(" . ln"),
        paren(frac(sub(run("D"), run("eq", italic=False)), sub(run("D"), run("s,eq", italic=False))))
    )


def eq_capacitancia():
    return group(
        run("C"), run(" = "),
        frac(group(run("2π . "), run("ε")),
             group(run("ln"), paren(frac(sub(run("D"), run("eq", italic=False)),
                                          sub(run("r"), run("eq", italic=False)))))),
        run(" × "), sup(run("10"), run("3")),
    )


def eq_condutancia():
    return group(run("G"), run(" = "), frac(group(run("C"), run(" . "), run("σ")), run("ε")))


def eq_impedancia_serie():
    return group(run("Z"), run(" = R + j"), run("ω"), run("L"))


def eq_admitancia_shunt():
    return group(run("Y"), run(" = G + j"), run("ω"), run("C"))


def eq_zo():
    return group(sub(run("Z"), run("o", italic=False)), run(" = "), sqrt(frac(run("Z"), run("Y"))))


def eq_gamma():
    return group(run("γ"), run(" = "), sqrt(group(run("Z"), run(" . "), run("Y"))),
                 run(" = "), run("α"), run(" + j"), run("β"))


def eq_ro():
    return group(sub(run("R"), run("o", italic=False)), run(" = "), sqrt(frac(run("L"), run("C"))))


def eq_zs_pi():
    return group(sub(run("Z"), run("s", italic=False)), run(" = "), sub(run("Z"), run("o", italic=False)),
                 run(" . senh"), paren(run("γ"), run("l")))


def eq_yp_pi():
    return group(
        sub(run("Y"), run("p", italic=False)), run(" = "),
        frac(run("1"), sub(run("Z"), run("o", italic=False))),
        run(" . tanh"), paren(frac(group(run("γ"), run("l")), run("2"))),
    )


def eq_vs_caso1():
    return group(
        sub(run("V"), run("s", italic=False)), run(" = "), sub(run("V"), run("c", italic=False)),
        run(" . cosh"), paren(run("γ"), run("l")), run(" + "),
        sub(run("I"), run("c", italic=False)), run(" . "), sub(run("Z"), run("o", italic=False)),
        run(" . senh"), paren(run("γ"), run("l")),
    )


def eq_is_caso1():
    return group(
        sub(run("I"), run("s", italic=False)), run(" = "),
        frac(sub(run("V"), run("c", italic=False)), sub(run("Z"), run("o", italic=False))),
        run(" . senh"), paren(run("γ"), run("l")), run(" + "),
        sub(run("I"), run("c", italic=False)), run(" . cosh"), paren(run("γ"), run("l")),
    )


def eq_sil():
    return group(
        run("SIL"), run(" = "),
        frac(sup(run("VCL"), run("2")), sub(run("R"), run("o", italic=False))),
    )


def eq_campo_eletrico():
    return group(
        run("E"), run(" = "),
        frac(run("Q"), group(run("2π . "), run("ε"), run(" . "), sub(run("r"), run("eq", italic=False)))),
        run("     ,     "), run("Q"), run(" = "), run("C"), run(" . "), sub(run("V"), run("c", italic=False)),
    )
