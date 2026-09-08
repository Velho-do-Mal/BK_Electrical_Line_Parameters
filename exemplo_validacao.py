"""
exemplo_validacao.py — BK Electrical Line Parameters

Caso-exemplo extraído da própria planilha original
`CELT2IPOG_calculo_eletrico_de_LTs.xlsm` (condições preenchidas de fábrica),
usado para a validação obrigatória Excel × Python (seção 16/26/28 do
briefing do projeto). Os valores em `VALORES_EXCEL` foram lidos diretamente
das células da planilha (não recalculados) — ver
`Mapa_Calculos_LT_CELT2IPOG.md` para a origem célula-a-célula de cada um.

Condição: condutor CA Tulip 336 MCM (catálogo, item 15), 1 subcondutor/fase,
estrutura típica 3, l=100 km, f=60 Hz, VCL=72 kV, SC=1000 kVA, cos(phi)=1
(atrasado), ambiente ar seco (eps_r=1, sigma=0).
"""

from __future__ import annotations

import math
from dataclasses import dataclass

import calculos as c

ENTRADA_EXEMPLO = c.EntradaLT(
    nome_linha="Exemplo (validação de fábrica da planilha)",
    catalogo_condutor="CA", item_condutor=15, nome_condutor=None,
    n_subcondutores=1, espacamento_feixe_cm=0.0,
    tipo_estrutura=3, geometria_custom=None,
    frequencia_hz=60.0, comprimento_km=100.0,
    vcl_kv=72.0, ang_vcl_deg=0.0,
    sc_kva=1000.0, cos_phi=1.0, atrasado=True,
    eps_r=1.0, sigma_s_m=0.0,
)

# (rótulo, célula de origem, valor Excel, unidade, função extratora do ResultadoLT)
VALORES_EXCEL: list[tuple[str, str, float, str]] = [
    ("DAB", "Estruturas!AE6", 3.310966626228661, "m"),
    ("DBC", "Estruturas!AF6", 1.7, "m"),
    ("DCA", "Estruturas!AG6", 3.310966626228661, "m"),
    ("GMD (Deq)", "Calc!D7", 2.6512631927148687, "m"),
    ("R", "Calc!A10", 0.206, "Ω/km"),
    ("L", "Calc!B10", 1.2052986983073337, "mH/km"),
    ("C", "Calc!D10", 9.664146672262786, "nF/km"),
    ("|Zo|", "Calc!A28", 370.05000742361005, "Ω"),
    ("∠Zo", "Calc!A28", -12.193774960379088, "°"),
    ("Ro", "Calc!A30", 353.15518545167816, "Ω"),
    ("αl", "Calc!E30", 0.028476539440712996, "-"),
    ("βl", "Calc!F30", 0.13177852577023302, "-"),
    ("|VoL| emissor", "Principal!L21", 71.69485765295626, "kV"),
    ("∠VoL emissor", "Principal!M21", 0.7185867345733055, "°"),
    ("|Io| emissor", "Principal!L23", 17.086642487949455, "A"),
    ("|So| emissor", "Principal!L25", 2121.804502940562, "kVA"),
    ("Po emissor", "Principal!O21", 1008.67064902142, "kW"),
    ("Pc receptor", "Principal!O24", 1000.0, "kW"),
    ("Perdas", "Principal!U4", 8.67064902141999, "kW"),
    ("Perdas %", "Principal!U5", 0.8596115124230072, "%"),
    ("V vazio (regulação)", "Principal!U7", 72.29207718338081, "kV"),
    ("Regulação %", "Principal!U9", 0.4056627546954728, "%"),
    ("SIL", "Principal!U13", 14679.099199321628, "kW"),
    ("Campo elétrico superficial", "Calc!H15", 8.557584138229295, "kV/cm"),
]


@dataclass
class LinhaValidacao:
    item: str
    origem: str
    excel: float
    python: float
    unidade: str
    diferenca_abs: float
    diferenca_pct: float
    status: str


def _extrair(resultado: c.ResultadoLT, item: str) -> float:
    r = resultado
    mapa = {
        "DAB": r.geometria.dab_m,
        "DBC": r.geometria.dbc_m,
        "DCA": r.geometria.dca_m,
        "GMD (Deq)": r.geometria.gmd_m,
        "R": r.rlgc.r_ohm_km,
        "L": r.rlgc.l_h_km * 1000,
        "C": r.rlgc.c_f_km * 1e9,
        "|Zo|": abs(r.linha_longa.zo_ohm),
        "∠Zo": math.degrees(__import__("cmath").phase(r.linha_longa.zo_ohm)),
        "Ro": r.linha_longa.ro_ohm,
        "αl": r.linha_longa.gl.real,
        "βl": r.linha_longa.gl.imag,
        "|VoL| emissor": r.caso1.vol_emissor_kv,
        "∠VoL emissor": r.caso1.ang_vol_emissor_deg,
        "|Io| emissor": r.caso1.io_emissor_a,
        "|So| emissor": r.caso1.so_emissor_kva,
        "Po emissor": r.caso1.po_emissor_kw,
        "Pc receptor": r.caso1.pc_receptor_kw,
        "Perdas": r.caso1.perdas_kw,
        "Perdas %": r.caso1.perdas_pct * 100,
        "V vazio (regulação)": r.regulacao.vol_vazio_kv,
        "Regulação %": r.regulacao.regulacao_pct * 100,
        "SIL": r.sil.sil_kw,
        "Campo elétrico superficial": r.campo_eletrico.campo_kv_cm,
    }
    return mapa[item]


def gerar_tabela_validacao(tolerancia_pct: float = 0.05) -> list[LinhaValidacao]:
    """Executa o caso-exemplo e monta a tabela de validação Excel × Python
    (seção 16 do briefing). Tolerância default 0,05% (arredondamentos de
    exibição da planilha original explicam qualquer resíduo abaixo disso)."""
    resultado = c.calcular_linha_completa(ENTRADA_EXEMPLO)
    linhas = []
    for item, origem, excel_val, unidade in VALORES_EXCEL:
        py_val = _extrair(resultado, item)
        diff_abs = py_val - excel_val
        diff_pct = (diff_abs / excel_val * 100) if excel_val != 0 else 0.0
        status = "OK" if abs(diff_pct) <= tolerancia_pct else "DIVERGE"
        linhas.append(LinhaValidacao(item, origem, excel_val, py_val, unidade, diff_abs, diff_pct, status))
    return linhas


if __name__ == "__main__":
    linhas = gerar_tabela_validacao()
    largura = max(len(l.item) for l in linhas)
    print(f"{'Item':{largura}s}  {'Excel':>16}  {'Python':>16}  {'Dif %':>10}  Status")
    for l in linhas:
        print(f"{l.item:{largura}s}  {l.excel:16.6f}  {l.python:16.6f}  {l.diferenca_pct:10.6f}  {l.status}")
    n_diverge = sum(1 for l in linhas if l.status == "DIVERGE")
    print(f"\n{len(linhas)} itens comparados, {n_diverge} divergência(s) acima da tolerância.")
