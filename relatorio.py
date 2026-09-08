"""
relatorio.py — BK Electrical Line Parameters
Geração da Memória de Cálculo em Word (.docx), no padrão documental BK
Engenharia e Tecnologia (referência: Modelo_Memoria_Calculo.docx do projeto).

Também monta as ETAPAS DA MEMÓRIA DE CÁLCULO (fórmula → onde → dados →
substituição numérica → resultado → interpretação), reaproveitadas tanto
pela aba "5. Memória de Cálculo" do app quanto pelo relatório Word.

Nenhum cálculo de engenharia é feito aqui — apenas leitura de
`calculos.ResultadoLT` já calculado e formatação/apresentação.
"""

from __future__ import annotations

import io
import math
import os
from dataclasses import dataclass
from datetime import date

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor, Inches

import calculos as c
import dados

TEMPLATES_DIR = os.path.join(os.path.dirname(__file__), "templates")
# Observação: o relatório Word (memória de cálculo) não exibe logo — ver
# _cabecalho_documento(). O logo BK continua sendo usado apenas na interface
# do aplicativo (app.py).

BK_AZUL = RGBColor(0x15, 0x65, 0xA8)
BK_AZUL_ESCURO = RGBColor(0x0D, 0x3D, 0x63)
BK_CINZA = RGBColor(0x4A, 0x4A, 0x4A)


# ---------------------------------------------------------------------------
# MEMÓRIA DE CÁLCULO (etapas fórmula -> substituição -> resultado)
# ---------------------------------------------------------------------------

@dataclass
class EtapaCalculo:
    titulo: str
    formula: str
    onde: str
    dados: str
    substituicao: str
    resultado: str
    interpretacao: str = ""


def _fmt(v: float, casas: int = 4) -> str:
    return f"{v:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _fasor(z: complex, casas: int = 4) -> str:
    mod = abs(z)
    ang = math.degrees(math.atan2(z.imag, z.real))
    return f"{mod:.{casas}f} ∠ {ang:.{casas}f}°"


def montar_memoria_calculo(r: c.ResultadoLT) -> list[EtapaCalculo]:
    e = r.entrada
    g = r.geometria
    cond = r.condutor
    rl = r.rlgc
    ll = r.linha_longa
    pi = r.circuito_pi
    c1 = r.caso1
    reg = r.regulacao
    sil = r.sil
    ce = r.campo_eletrico

    etapas: list[EtapaCalculo] = []

    etapas.append(EtapaCalculo(
        "1. Distâncias entre fases (DAB, DBC, DCA)",
        "D = [(h₁-h₂)² + (v₁-v₂)²]^0,5",
        "h = deslocamento horizontal da fase (m); v = altura da fase (m)",
        f"Estrutura tipo {e.tipo_estrutura}: A(h={g.h_a:.3f}; v={g.v_a:.3f}), "
        f"B(h={g.h_b:.3f}; v={g.v_b:.3f}), C(h={g.h_c:.3f}; v={g.v_c:.3f}) m",
        f"DAB={g.dab_m:.4f} m; DBC={g.dbc_m:.4f} m; DCA={g.dca_m:.4f} m",
        f"DAB={g.dab_m:.4f} m; DBC={g.dbc_m:.4f} m; DCA={g.dca_m:.4f} m",
        "Distâncias geométricas entre os centros das fases, na estrutura selecionada.",
    ))
    etapas.append(EtapaCalculo(
        "2. Distância Média Geométrica entre fases (GMD / Deq)",
        "Deq = (DAB · DBC · DCA)^(1/3)",
        "DAB, DBC, DCA em metros",
        f"DAB={g.dab_m:.4f} m; DBC={g.dbc_m:.4f} m; DCA={g.dca_m:.4f} m",
        f"Deq = ({g.dab_m:.4f} × {g.dbc_m:.4f} × {g.dca_m:.4f})^(1/3)",
        f"Deq = {g.gmd_m:.4f} m",
        "Distância equivalente usada nas fórmulas de L e C da linha transposta.",
    ))
    etapas.append(EtapaCalculo(
        "3. Condutor e feixe de subcondutores",
        "req = (r · d)^0,5  [feixe duplo]   |   req = (r · d³ · √2)^0,25  [feixe quádruplo]",
        "r = raio do condutor unitário (m); d = espaçamento entre subcondutores (m)",
        f"Condutor {cond.catalogo} {cond.nome_comercial} (bitola {cond.bitola}); "
        f"diâmetro={cond.diametro_mm:.3f} mm; RMG={cond.rmg_m:.5f} m; "
        f"R={cond.r_ohm_km:.4f} Ω/km; nº subcondutores={cond.n_subcondutores}"
        + (f"; d={cond.espacamento_feixe_cm:.2f} cm" if cond.n_subcondutores > 1 else " (feixe simples)"),
        f"raio equivalente={cond.r_eq_m:.6f} m; RMG equivalente={cond.ds_eq_m:.6f} m",
        f"r_eq={cond.r_eq_m:.6f} m; Ds_eq={cond.ds_eq_m:.6f} m; R_fase={cond.r_fase_ohm_km:.5f} Ω/km",
        "Para feixe simples, r_eq e Ds_eq são os do condutor unitário do catálogo.",
    ))
    etapas.append(EtapaCalculo(
        "4. Resistência série (R)",
        "R = R_catálogo / nº subcondutores",
        "R_catálogo = resistência CA do condutor unitário (Ω/km, catálogo do fabricante)",
        f"R_catálogo={cond.r_ohm_km:.4f} Ω/km; nº subcondutores={cond.n_subcondutores}",
        f"R = {cond.r_ohm_km:.4f} / {cond.n_subcondutores}",
        f"R = {rl.r_ohm_km:.5f} Ω/km",
    ))
    etapas.append(EtapaCalculo(
        "5. Indutância série (L)",
        "L = 2×10⁻⁴ · ln(Deq / Ds_eq)",
        "Deq = GMD entre fases (m); Ds_eq = RMG equivalente do feixe (m)",
        f"Deq={g.gmd_m:.4f} m; Ds_eq={cond.ds_eq_m:.6f} m",
        f"L = 2×10⁻⁴ × ln({g.gmd_m:.4f} / {cond.ds_eq_m:.6f})",
        f"L = {rl.l_h_km*1000:.5f} mH/km",
    ))
    etapas.append(EtapaCalculo(
        "6. Capacitância (C) e condutância (G)",
        "C = 2π·ε / ln(Deq / r_eq) × 1000     |     G = C · σ / ε",
        "ε = εR·ε0 (ε0 ≈ 1/(36π)×10⁻⁹ F/m — aproximação de engenharia da planilha original); "
        "r_eq = raio equivalente do feixe (m); σ = condutividade do meio (S/m)",
        f"εR={e.eps_r:.4f}; Deq={g.gmd_m:.4f} m; r_eq={cond.r_eq_m:.6f} m; σ={e.sigma_s_m:.6f} S/m",
        f"C = 2π×{rl.eps_meio:.4e} / ln({g.gmd_m:.4f} / {cond.r_eq_m:.6f}) × 1000",
        f"C = {rl.c_f_km*1e9:.5f} nF/km; G = {rl.g_s_km:.3e} S/km",
    ))
    etapas.append(EtapaCalculo(
        "7. Impedância série e admitância shunt por km",
        "Z = R + jωL     |     Y = G + jωC     (ω = 2πf)",
        f"f={rl.freq_hz:.0f} Hz → ω={rl.omega_rad_s:.4f} rad/s",
        f"R={rl.r_ohm_km:.5f} Ω/km; L={rl.l_h_km*1000:.5f} mH/km; G={rl.g_s_km:.3e} S/km; C={rl.c_f_km*1e9:.5f} nF/km",
        f"Z = {rl.r_ohm_km:.5f} + j{rl.omega_rad_s*rl.l_h_km:.5f} Ω/km",
        f"Z = {_fasor(ll.z_km)} Ω/km; Y = {_fasor(ll.y_km*1000, 5)} mS/km",
    ))
    etapas.append(EtapaCalculo(
        "8. Impedância característica (Zo) e constante de propagação (γ)",
        "Zo = √(Z/Y)     |     γ = √(Z·Y) = α + jβ",
        "Z, Y por km",
        f"Z={_fasor(ll.z_km)} Ω/km; Y={_fasor(ll.y_km*1000,5)} mS/km",
        f"Zo = √(Z/Y); γ = √(Z·Y)",
        f"Zo = {_fasor(ll.zo_ohm)} Ω; γ = {ll.gamma_km.real:.6e} + j{ll.gamma_km.imag:.6e} (1/km)",
    ))
    etapas.append(EtapaCalculo(
        "9. Impedância de surto (Ro) e comprimento da linha (γ·l)",
        "Ro = √(L/C)     |     γ·l",
        "L, C por km; l = comprimento da linha (km)",
        f"L={rl.l_h_km*1000:.5f} mH/km; C={rl.c_f_km*1e9:.5f} nF/km; l={ll.comprimento_km:.3f} km",
        f"Ro = √({rl.l_h_km:.4e} / {rl.c_f_km:.4e})",
        f"Ro = {ll.ro_ohm:.4f} Ω; γ·l = {ll.gl.real:.6f} + j{ll.gl.imag:.6f}",
        "Ro é a impedância de surto (real, sem perdas) — usada no cálculo do SIL.",
    ))
    etapas.append(EtapaCalculo(
        "10. Circuito π equivalente",
        "Zs = Zo·senh(γl)     |     Yp = (1/Zo)·tanh(γl/2)",
        "Zo, γl calculados nas etapas 8 e 9",
        f"Zo={_fasor(ll.zo_ohm)} Ω; γl={ll.gl.real:.5f}+j{ll.gl.imag:.5f}",
        f"Zs = Zo × senh(γl); Yp = (1/Zo) × tanh(γl/2)",
        f"Zs = {_fasor(pi.zs_ohm)} Ω; Yp = {abs(pi.yp_siemens)*1000:.5f} mS ∠ "
        f"{math.degrees(math.atan2(pi.yp_siemens.imag,pi.yp_siemens.real)):.3f}°",
    ))
    etapas.append(EtapaCalculo(
        "11. Tensão e corrente no receptor (Caso 1 — carga plena)",
        "Vc = (VCL·1000/√3) ∠θv     |     Ic = conj(S / (3·Vc)),  S = SC·1000 ∠φ",
        "VCL = tensão de linha no receptor (kV); SC = potência aparente no receptor (kVA); "
        "φ = ± acos(cosφ) (sinal positivo para fator de potência atrasado)",
        f"VCL={e.vcl_kv:.4f} kV; SC={e.sc_kva:.4f} kVA; cosφ={e.cos_phi:.4f} "
        f"({'atrasado' if e.atrasado else 'adiantado'})",
        f"Vc = ({e.vcl_kv:.4f}×1000/√3) ∠{e.ang_vcl_deg:.2f}°",
        f"Vc = {_fasor(c1.vc_fasor_v, 2)} V (fase); Ic = {_fasor(c1.ic_fasor_a)} A",
    ))
    etapas.append(EtapaCalculo(
        "12. Tensão e corrente no emissor (Caso 1)",
        "Vs = Vc·cosh(γl) + Ic·Zo·senh(γl)     |     Is = (Vc/Zo)·senh(γl) + Ic·cosh(γl)",
        "Equações gerais do modelo de linha longa (parâmetros distribuídos)",
        f"Vc={_fasor(c1.vc_fasor_v,2)} V; Ic={_fasor(c1.ic_fasor_a)} A; Zo={_fasor(ll.zo_ohm)} Ω",
        "Vs = Vc·cosh(γl) + Ic·Zo·senh(γl)",
        f"|VoL| = {c1.vol_emissor_kv:.4f} kV ∠{c1.ang_vol_emissor_deg:.4f}°; "
        f"|Io| = {c1.io_emissor_a:.4f} A ∠{c1.ang_io_emissor_deg:.4f}°",
    ))
    etapas.append(EtapaCalculo(
        "13. Potências, perdas e regulação",
        "So = 3·Vs·Is*     |     Po=|So|·cos(∠So)     |     Pc=SC·cosφ     |     "
        "Perdas=Po−Pc     |     Regulação=(V_vazio−VCL)/VCL,  V_vazio=|Vs/cosh(γl)|",
        "Vs, Is do emissor (etapa 12); Vc do receptor em vazio obtida com a mesma "
        "tensão de emissor do Caso 1 (Ic=0)",
        f"So={c1.so_emissor_kva:.4f} kVA∠{c1.ang_so_emissor_deg:.3f}°; Pc={c1.pc_receptor_kw:.4f} kW",
        f"Perdas = {c1.po_emissor_kw:.4f} − {c1.pc_receptor_kw:.4f}",
        f"Po={c1.po_emissor_kw:.4f} kW; Perdas={c1.perdas_kw:.4f} kW ({c1.perdas_pct*100:.4f}%); "
        f"Regulação={reg.regulacao_pct*100:.4f}%",
    ))
    etapas.append(EtapaCalculo(
        "14. Potência natural (SIL) e carregamento",
        "SIL = VCL² / Ro     |     %SIL = Pc / SIL",
        "VCL em kV; Ro = impedância de surto (Ω)",
        f"VCL={e.vcl_kv:.4f} kV; Ro={ll.ro_ohm:.4f} Ω; Pc={c1.pc_receptor_kw:.4f} kW",
        f"SIL = {e.vcl_kv:.4f}² / {ll.ro_ohm:.4f} × 1000",
        f"SIL = {sil.sil_kw:.2f} kW; %SIL = {sil.pct_sil*100:.2f}%",
        "SIL (Surge Impedance Loading) é o carregamento natural da linha, referência de "
        "carregabilidade — carregamentos muito abaixo de 1×SIL indicam linha sobredimensionada "
        "para a demanda; muito acima, possível necessidade de compensação reativa.",
    ))
    etapas.append(EtapaCalculo(
        "15. Campo elétrico superficial do condutor (indicador de corona)",
        "Q = C·Vc (C em F/m)     |     E = Q / (2π·ε·r_eq)     |     E[kV/cm] = E[V/m]×10⁻⁵",
        "C = capacitância (F/m); Vc = tensão fase-neutro no receptor (V); r_eq = raio equivalente do feixe (m)",
        f"C={rl.c_f_km/1000:.4e} F/m; |Vc|={abs(c1.vc_fasor_v):.2f} V; r_eq={cond.r_eq_m:.6f} m",
        f"E = ({rl.c_f_km/1000:.4e} × {abs(c1.vc_fasor_v):.2f}) / (2π × {rl.eps_meio:.4e} × {cond.r_eq_m:.6f})",
        f"E = {ce.campo_kv_cm:.4f} kV/cm",
        f"Valor de referência informativo para disrupção do ar seco: ~{ce.limite_referencia_kv_cm:.1f} kV/cm "
        f"(não é um limite normativo aplicado automaticamente pela planilha original — "
        f"comparação apenas orientativa).",
    ))
    return etapas


# ---------------------------------------------------------------------------
# GERAÇÃO DO DOCUMENTO WORD
# ---------------------------------------------------------------------------

def nome_arquivo_relatorio(identificacao: dict) -> str:
    num = (identificacao.get("numero_documento") or "SN").strip() or "SN"
    nome = (identificacao.get("nome_documento") or "Memoria_Calculo").strip()
    rev = (identificacao.get("revisao") or "0").strip()
    nome = "".join(ch if ch.isalnum() or ch in " _-" else "_" for ch in nome).replace(" ", "_")
    num = "".join(ch if ch.isalnum() or ch in "_-" else "_" for ch in num)
    return f"{num}_{nome}_Rev_{rev}.docx"


def _set_cell_bg(cell, hex_color: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.makeelement(qn('w:shd'), {qn('w:val'): 'clear', qn('w:color'): 'auto', qn('w:fill'): hex_color})
    tc_pr.append(shd)


def _cell_text(cell, text, bold=False, size=9, color=None, align=None):
    cell.text = ""
    p = cell.paragraphs[0]
    if align:
        p.alignment = align
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = color


def _cabecalho_documento(doc: Document, ident: dict, titulo_doc: str):
    tbl = doc.add_table(rows=2, cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    widths = [Cm(3), Cm(10), Cm(3)]
    for row in tbl.rows:
        for cell, w in zip(row.cells, widths):
            cell.width = w

    # Linha 1: (célula reservada, sem logo) | razão social | espaço cliente
    # Mantida em branco propositalmente — não deve exibir logo no relatório
    # (mesmo padrão da célula do modelo Modelo_Memoria_Calculo.docx, que também
    # não traz uma imagem embutida nessa posição, apenas o rótulo de referência).
    cel_logo = tbl.cell(0, 0)
    cel_logo.merge(tbl.cell(1, 0))

    cel_empresa = tbl.cell(0, 1)
    cel_empresa.merge(tbl.cell(1, 1))
    cel_empresa.text = ""
    p = cel_empresa.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p.add_run("BK ENGENHARIA E TECNOLOGIA")
    r1.bold = True
    r1.font.size = Pt(13)
    r1.font.color.rgb = BK_AZUL_ESCURO
    p2 = cel_empresa.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run(titulo_doc)
    r2.font.size = Pt(10)
    r2.font.color.rgb = BK_CINZA

    cel_cliente = tbl.cell(0, 2)
    cel_cliente.merge(tbl.cell(1, 2))
    _cell_text(cel_cliente, ident.get("cliente", "") or "-", bold=True, size=9,
               align=WD_ALIGN_PARAGRAPH.CENTER)

    doc.add_paragraph()

    # Tabela de identificação documental
    tbl2 = doc.add_table(rows=2, cols=5)
    tbl2.style = "Table Grid"
    hdrs = ["Código", "Revisão", "Documento", "Aprovação", "Pág."]
    for i, h in enumerate(hdrs):
        _cell_text(tbl2.cell(0, i), h, bold=True, size=8, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_bg(tbl2.cell(0, i), "1565A8")
        tbl2.cell(0, i).paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    vals = [ident.get("numero_documento", "") or "-", ident.get("revisao", "0"),
            ident.get("nome_documento", ""), "", "1"]
    for i, v in enumerate(vals):
        _cell_text(tbl2.cell(1, i), str(v), size=8, align=WD_ALIGN_PARAGRAPH.CENTER)

    doc.add_paragraph()


def _titulo(doc: Document, texto: str, nivel: int = 1):
    h = doc.add_heading(texto, level=nivel)
    for run in h.runs:
        run.font.color.rgb = BK_AZUL_ESCURO
    return h


def _paragrafo(doc: Document, texto: str, negrito=False, tamanho=10):
    p = doc.add_paragraph()
    run = p.add_run(texto)
    run.bold = negrito
    run.font.size = Pt(tamanho)
    return p


def _tabela_resultados(doc: Document, titulo: str, linhas: list[tuple[str, str, str]]):
    _paragrafo(doc, titulo, negrito=True)
    tbl = doc.add_table(rows=1 + len(linhas), cols=3)
    tbl.style = "Table Grid"
    for i, h in enumerate(["Item", "Valor", "Unidade"]):
        _cell_text(tbl.cell(0, i), h, bold=True, size=9)
        _set_cell_bg(tbl.cell(0, i), "D9E6F2")
    for r_i, (item, valor, unidade) in enumerate(linhas, start=1):
        _cell_text(tbl.cell(r_i, 0), item, size=9)
        _cell_text(tbl.cell(r_i, 1), valor, size=9, align=WD_ALIGN_PARAGRAPH.RIGHT)
        _cell_text(tbl.cell(r_i, 2), unidade, size=9)
    doc.add_paragraph()


def _rodape_padrao(doc: Document):
    section = doc.sections[0]
    footer = section.footer
    p = footer.paragraphs[0]
    p.text = ""
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(
        "Este documento é de propriedade da BK ENGENHARIA E TECNOLOGIA, não é permitida sua "
        "reprodução ou comunicação a terceiros sem prévia autorização."
    )
    run.italic = True
    run.font.size = Pt(7)
    run.font.color.rgb = BK_CINZA
    p2 = footer.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = p2.add_run("www.bk-engenharia.com")
    run2.font.size = Pt(7)
    run2.font.color.rgb = BK_AZUL


def gerar_relatorio_docx(r: c.ResultadoLT, ident: dict) -> bytes:
    """Gera a Memória de Cálculo em .docx e retorna os bytes do arquivo."""
    doc = Document()

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10)

    titulo_doc = ident.get("nome_documento", "Memória de Cálculo — Parâmetros Elétricos de LT")
    _cabecalho_documento(doc, ident, titulo_doc)
    _rodape_padrao(doc)

    _titulo(doc, titulo_doc, nivel=0)
    _paragrafo(doc, f"Projeto: {ident.get('projeto','-')}", negrito=True)
    _paragrafo(doc, f"Cliente: {ident.get('cliente','-')}")
    _paragrafo(doc, f"Documento nº: {ident.get('numero_documento','-')}    "
                     f"Revisão: {ident.get('revisao','0')}    Data: {ident.get('data','-')}")
    _paragrafo(doc, f"Responsável técnico: {ident.get('responsavel_tecnico','-')}")
    doc.add_page_break()

    # 1. OBJETIVO
    _titulo(doc, "1. Objetivo")
    _paragrafo(doc, f"Este relatório técnico apresenta a memória de cálculo dos parâmetros elétricos "
                     f"da linha de transmissão \"{r.entrada.nome_linha}\", incluindo geometria, "
                     f"parâmetros unitários (R, L, G, C), impedância característica, constante de "
                     f"propagação, circuito π equivalente, condições de carga plena (Caso 1), "
                     f"regulação de tensão, potência natural (SIL) e indicador de campo elétrico "
                     f"superficial (corona), calculados pelo software BK Electrical Line Parameters.")

    # 2. DOCUMENTOS DE REFERÊNCIA
    _titulo(doc, "2. Documentos de Referência")
    for ref in [
        "Planilha de cálculo elétrico de LTs original da BK Engenharia e Tecnologia "
        "(CELT2IPOG_calculo_eletrico_de_LTs.xlsm) — fonte de verdade da metodologia reproduzida.",
        "Catálogos de condutores CA (AAC), CAA (ACSR) e CAL (AAAC) — fabricantes/normas de mercado.",
        "Fuchs, R.D. — Transmissão de Energia Elétrica: Linhas Aéreas (modelo de linha longa / "
        "parâmetros distribuídos, referência clássica da metodologia empregada).",
        "ABNT NBR 5422 — Projeto de linhas aéreas de transmissão de energia elétrica "
        "(referência normativa geral de projeto de LT — não auditada formula-a-fórmula neste "
        "software; ver seção de premissas).",
    ]:
        doc.add_paragraph(ref, style="List Bullet")

    # 3. METODOLOGIA
    _titulo(doc, "3. Metodologia de Cálculo")
    _paragrafo(doc, "Modelo de linha longa (parâmetros distribuídos), com representação da linha por "
                     "impedância série Z=R+jωL e admitância shunt Y=G+jωC por unidade de comprimento, "
                     "impedância característica Zo=√(Z/Y) e constante de propagação γ=√(Z·Y)=α+jβ. "
                     "A relação entre grandezas no emissor e no receptor é dada pelas equações gerais "
                     "da linha longa (funções hiperbólicas de γ·l). Circuito π equivalente derivado de "
                     "Zo e γ·l. SIL calculado a partir da impedância de surto Ro=√(L/C).")
    _paragrafo(doc, "Premissa de engenharia herdada da planilha original: a permissividade do vácuo "
                     "é adotada pela aproximação clássica ε0 ≈ 1/(36π)×10⁻⁹ F/m (≈8,8419×10⁻¹² F/m), "
                     "e não o valor CODATA de precisão (8,8542×10⁻¹² F/m) — diferença ≈0,14%, "
                     "propagada a C, Zo, Ro e SIL. Mantida por fidelidade à metodologia original "
                     "(validado — ver seção 9).")

    # 4. DADOS DE ENTRADA
    _titulo(doc, "4. Dados de Entrada")
    e = r.entrada
    linhas_entrada = [
        ("Linha", e.nome_linha, "-"),
        ("Frequência", f"{e.frequencia_hz:.0f}", "Hz"),
        ("Comprimento", f"{e.comprimento_km:.3f}", "km"),
        ("Tensão de linha no receptor (VCL)", f"{e.vcl_kv:.4f}", "kV"),
        ("Potência aparente no receptor (SC)", f"{e.sc_kva:.4f}", "kVA"),
        ("Fator de potência (cosφ)", f"{e.cos_phi:.4f} ({'atrasado' if e.atrasado else 'adiantado'})", "-"),
        ("Condutor", f"{r.condutor.catalogo} — {r.condutor.nome_comercial} ({r.condutor.bitola})", "-"),
        ("Subcondutores por fase", f"{r.condutor.n_subcondutores}", "-"),
        ("Estrutura", f"Tipo {e.tipo_estrutura}" if e.tipo_estrutura else "Customizada", "-"),
        ("Ambiente (εR / σ)", f"{e.eps_r:.3f} / {e.sigma_s_m:.6f}", "- / S/m"),
    ]
    _tabela_resultados(doc, "Dados de entrada", linhas_entrada)

    # 5. CARACTERÍSTICAS DOS CONDUTORES E ESTRUTURA
    _titulo(doc, "5. Características dos Condutores e da Estrutura")
    cond = r.condutor
    g = r.geometria
    linhas_cond = [
        ("Diâmetro do condutor", f"{cond.diametro_mm:.3f}", "mm"),
        ("RMG do condutor (Ds)", f"{cond.rmg_m:.5f}", "m"),
        ("Ampacidade (catálogo)", f"{cond.ampacidade_a:.0f}", "A"),
        ("Resistência CA (catálogo, condutor unitário)", f"{cond.r_ohm_km:.4f}", "Ω/km"),
        ("Raio equivalente do feixe", f"{cond.r_eq_m:.6f}", "m"),
        ("RMG equivalente do feixe", f"{cond.ds_eq_m:.6f}", "m"),
        ("DAB", f"{g.dab_m:.4f}", "m"),
        ("DBC", f"{g.dbc_m:.4f}", "m"),
        ("DCA", f"{g.dca_m:.4f}", "m"),
        ("GMD (Deq)", f"{g.gmd_m:.4f}", "m"),
    ]
    _tabela_resultados(doc, "Condutor e geometria", linhas_cond)

    # 6. DESENVOLVIMENTO DOS CÁLCULOS (memória)
    _titulo(doc, "6. Desenvolvimento dos Cálculos")
    for et in montar_memoria_calculo(r):
        _paragrafo(doc, et.titulo, negrito=True)
        _paragrafo(doc, f"Fórmula: {et.formula}")
        if et.onde:
            _paragrafo(doc, f"Onde: {et.onde}")
        _paragrafo(doc, f"Dados utilizados: {et.dados}")
        _paragrafo(doc, f"Substituição numérica: {et.substituicao}")
        p = _paragrafo(doc, f"Resultado: {et.resultado}", negrito=True)
        if et.interpretacao:
            ip = doc.add_paragraph()
            ir = ip.add_run(et.interpretacao)
            ir.italic = True
            ir.font.size = Pt(9)
            ir.font.color.rgb = BK_CINZA
        doc.add_paragraph()

    # 7. RESULTADOS
    doc.add_page_break()
    _titulo(doc, "7. Resultados")
    ll, pi_eq, c1, reg, sil, ce = r.linha_longa, r.circuito_pi, r.caso1, r.regulacao, r.sil, r.campo_eletrico
    ang_zo = math.degrees(math.atan2(ll.zo_ohm.imag, ll.zo_ohm.real))
    linhas_param = [
        ("R", f"{r.rlgc.r_ohm_km:.5f}", "Ω/km"),
        ("L", f"{r.rlgc.l_h_km*1000:.5f}", "mH/km"),
        ("C", f"{r.rlgc.c_f_km*1e9:.5f}", "nF/km"),
        ("G", f"{r.rlgc.g_s_km:.3e}", "S/km"),
        ("|Zo|", f"{abs(ll.zo_ohm):.4f}", "Ω"),
        ("∠Zo", f"{ang_zo:.4f}", "°"),
        ("Ro (impedância de surto)", f"{ll.ro_ohm:.4f}", "Ω"),
        ("αl", f"{ll.gl.real:.6f}", "-"),
        ("βl", f"{ll.gl.imag:.6f}", "-"),
    ]
    _tabela_resultados(doc, "Parâmetros e características da linha", linhas_param)

    linhas_pi = [
        ("|Zs| (série do π)", f"{abs(pi_eq.zs_ohm):.4f}", "Ω"),
        ("|Yp| (shunt do π)", f"{abs(pi_eq.yp_siemens)*1000:.5f}", "mS"),
    ]
    _tabela_resultados(doc, "Circuito π equivalente", linhas_pi)

    linhas_fonte = [
        ("|VoL| (emissor)", f"{c1.vol_emissor_kv:.4f}", "kV"),
        ("∠VoL (emissor)", f"{c1.ang_vol_emissor_deg:.4f}", "°"),
        ("|Io| (emissor)", f"{c1.io_emissor_a:.4f}", "A"),
        ("|So| (emissor)", f"{c1.so_emissor_kva:.4f}", "kVA"),
        ("Po (emissor)", f"{c1.po_emissor_kw:.4f}", "kW"),
    ]
    _tabela_resultados(doc, "Fonte (emissor) — Caso 1 (carga plena)", linhas_fonte)

    linhas_carga = [
        ("VCL (receptor)", f"{e.vcl_kv:.4f}", "kV"),
        ("SC (receptor)", f"{e.sc_kva:.4f}", "kVA"),
        ("Pc (receptor)", f"{c1.pc_receptor_kw:.4f}", "kW"),
    ]
    _tabela_resultados(doc, "Carga plena (receptor) — Caso 1", linhas_carga)

    linhas_desempenho = [
        ("Perdas", f"{c1.perdas_kw:.4f}", "kW"),
        ("Perdas", f"{c1.perdas_pct*100:.4f}", "%"),
        ("Regulação de tensão", f"{reg.regulacao_pct*100:.4f}", "%"),
        ("SIL (Potência Natural)", f"{sil.sil_kw:.2f}", "kW"),
        ("%SIL (carregamento)", f"{sil.pct_sil*100:.2f}", "%"),
        ("Campo elétrico superficial", f"{ce.campo_kv_cm:.4f}", "kV/cm"),
    ]
    _tabela_resultados(doc, "Desempenho elétrico", linhas_desempenho)

    if r.caso2_manual is not None:
        c2 = r.caso2_manual
        linhas_c2 = [
            ("Tensão no receptor (vazio)", f"{c2.vol_receptor_kv:.4f}", "kV"),
            ("Corrente no receptor (vazio)", f"{c2.io_receptor_a:.4f}", "A"),
            ("Potência no receptor (vazio)", f"{c2.sc_receptor_kva:.4f}", "kVA"),
        ]
        _paragrafo(doc, "Caso 2 — Tensão de emissor especificada (PREMISSA NÃO IDENTIFICADA — "
                         "ver observações, item 10).", negrito=True)
        _tabela_resultados(doc, "Caso 2 — resultados no receptor", linhas_c2)

    # 8. QUADRO DE REFERÊNCIA DE CARREGABILIDADE
    _titulo(doc, "8. Quadro de Referência — Carregabilidade Típica")
    _paragrafo(doc, "Tabela de referência sem vínculo de fórmula direto com esta linha (extraída da "
                     "planilha original, aba \"Análises\"), apresentada apenas como comparação informativa "
                     "entre carregamento (%SIL) e valores típicos de regulação/perdas.")
    linhas_ref = [(f"{row['pct_sil']*100:.0f}% SIL", f"R≈{row['regulacao_pct']*100:.2f}% / "
                   f"P≈{row['perdas_pct']*100:.2f}%", "") for row in dados.TABELA_CARREGABILIDADE_REFERENCIA]
    _tabela_resultados(doc, "Regulação e perdas típicas em função do %SIL", linhas_ref)

    # 9. VALIDAÇÃO EXCEL x PYTHON (metodologia)
    _titulo(doc, "9. Validação da Metodologia (Excel × Python)")
    _paragrafo(doc, "O motor de cálculo deste software foi validado célula a célula contra o "
                     "caso-exemplo de fábrica da planilha original (condutor CA Tulip 336 MCM, "
                     "estrutura típica 3, l=100 km, VCL=72 kV, SC=1000 kVA, cosφ=1). Os 24 resultados "
                     "comparados (geometria, RLGC, impedância característica, propagação, Caso 1, "
                     "regulação, SIL e campo elétrico) reproduziram os valores da planilha original "
                     "com diferença inferior a 0,01% (resíduo apenas de arredondamento de exibição do "
                     "Excel). Ver arquivo de testes automatizados (tests/test_calculos.py) para o "
                     "detalhamento completo.")

    # 10. OBSERVAÇÕES / PREMISSAS NÃO IDENTIFICADAS
    _titulo(doc, "10. Observações e Premissas a Validar")
    for obs in [
        "εR/σ (ambiente): parâmetros avançados, mantidos em ar seco por padrão (εR=1, σ=0) salvo "
        "indicação em contrário do usuário.",
        "Campo elétrico superficial: indicador informativo (não é uma verificação normativa "
        "automática); comparação com limite de referência de ~21,1 kV/cm (ar seco) é apenas orientativa.",
        "Caso 2 (tensão de emissor especificada): quando utilizado, reproduz a metodologia exata da "
        "planilha original (energização com a impedância de carga do Caso 1 — não necessariamente "
        "um receptor em circuito aberto). Interpretação de engenharia pendente de validação formal "
        "com a equipe responsável antes de uso em verificações de sobretensão (efeito Ferranti).",
        "Estruturas 13, 14 e 15 do catálogo não possuem geometria cadastrada na planilha original — "
        "indisponíveis até cadastro futuro ou uso de geometria customizada.",
    ]:
        doc.add_paragraph(obs, style="List Bullet")

    # 11. CONCLUSÃO
    _titulo(doc, "11. Conclusão")
    _paragrafo(doc, f"Os parâmetros elétricos calculados para a linha \"{e.nome_linha}\" indicam "
                     f"perdas de {c1.perdas_pct*100:.2f}% e regulação de tensão de "
                     f"{reg.regulacao_pct*100:.2f}% nas condições especificadas, com carregamento de "
                     f"{sil.pct_sil*100:.1f}% do SIL. Os resultados foram obtidos por reprodução fiel "
                     f"da metodologia da planilha de cálculo elétrico de LTs da BK Engenharia e "
                     f"Tecnologia, validada conforme seção 9.")

    doc.add_paragraph()
    p_aprov = doc.add_paragraph()
    p_aprov.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_aprov = p_aprov.add_run("APROVADO")
    r_aprov.bold = True
    r_aprov.font.size = Pt(12)
    r_aprov.font.color.rgb = BK_AZUL_ESCURO

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()
