"""
calculos.py — BK Electrical Line Parameters
Motor de cálculo de parâmetros elétricos de Linha de Transmissão (LT).

Reproduz fielmente a metodologia da planilha original
`CELT2IPOG_calculo_eletrico_de_LTs.xlsm` (modelo de linha longa —
parâmetros distribuídos, γ, Zc, ABCD). Nenhuma fórmula foi substituída
por formulação "mais moderna" sem autorização — ver
`Mapa_Calculos_LT_CELT2IPOG.md` para o rastreamento célula-a-célula.

Convenções:
  - Nenhum arredondamento interno. Arredondamento só na apresentação
    (app.py / relatorio.py).
  - Fasores representados como `complex` do Python (módulo em unidade
    de engenharia indicada em cada função; ângulos internos SEMPRE em
    radianos — graus só na interface/relatório).
  - Todas as grandezas "por km" referem-se a valores unitários da
    linha; multiplicação pelo comprimento (l, em km) é explícita.

Este módulo NÃO deve conter nenhuma chamada Streamlit (st.*). Toda a
lógica de interface fica em app.py.
"""

from __future__ import annotations

import cmath
import math
from dataclasses import dataclass, field
from typing import Literal, Optional

import dados

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

EPS0 = dados.PERMISSIVIDADE_VACUO  # F/m
MU0 = dados.PERMEABILIDADE_VACUO   # H/m


class ErroCalculo(Exception):
    """Erro de engenharia (situação fisicamente inconsistente ou premissa
    não identificada exigindo validação do usuário) — distinto de erro de
    entrada (tratado em validacoes.py)."""


# ---------------------------------------------------------------------------
# 1. GEOMETRIA
# ---------------------------------------------------------------------------

@dataclass
class ResultadoGeometria:
    tipo_estrutura: Optional[int]
    h_a: float
    h_b: float
    h_c: float
    v_a: float
    v_b: float
    v_c: float
    dab_m: float
    dbc_m: float
    dca_m: float
    gmd_m: float  # Deq — Distância Média Geométrica entre fases


def _distancia(h1: float, v1: float, h2: float, v2: float) -> float:
    """Distância euclidiana entre duas fases (m). Origem: Estruturas!AE/AF/AG."""
    return math.hypot(h1 - h2, v1 - v2)


def geometria_estrutura(tipo_estrutura: Optional[int] = None,
                         h_a: Optional[float] = None, h_b: Optional[float] = None, h_c: Optional[float] = None,
                         v_a: Optional[float] = None, v_b: Optional[float] = None, v_c: Optional[float] = None
                         ) -> ResultadoGeometria:
    """Calcula DAB, DBC, DCA e GMD (Deq) a partir de uma estrutura do
    catálogo (tipo_estrutura 1-15) ou de geometria customizada (h_a..v_c).

    Origem: Estruturas!AE4:AG18 (fórmulas de distância), Calc!D7 (GMD).
    """
    if tipo_estrutura is not None:
        est = dados.ESTRUTURAS_POR_TIPO.get(tipo_estrutura)
        if est is None:
            raise ErroCalculo(f"Estrutura tipo {tipo_estrutura} não existe no catálogo (1-15).")
        if est['h_A'] is None:
            raise ErroCalculo(
                f"Estrutura tipo {tipo_estrutura} não possui geometria cadastrada na planilha "
                f"original (PREMISSA NÃO IDENTIFICADA). Informe geometria customizada (h/v) "
                f"ou escolha outra estrutura."
            )
        h_a, h_b, h_c = est['h_A'], est['h_B'], est['h_C']
        v_a, v_b, v_c = est['v_A'], est['v_B'], est['v_C']
    else:
        if None in (h_a, h_b, h_c, v_a, v_b, v_c):
            raise ErroCalculo("Geometria customizada exige h_a,h_b,h_c,v_a,v_b,v_c.")

    dab = _distancia(h_a, v_a, h_b, v_b)
    dbc = _distancia(h_b, v_b, h_c, v_c)
    dca = _distancia(h_c, v_c, h_a, v_a)
    gmd = (dab * dbc * dca) ** (1 / 3)  # Calc!D7

    return ResultadoGeometria(tipo_estrutura, h_a, h_b, h_c, v_a, v_b, v_c, dab, dbc, dca, gmd)


# ---------------------------------------------------------------------------
# 2. CONDUTOR E FEIXE (SUBCONDUTORES POR FASE)
# ---------------------------------------------------------------------------

TipoCatalogo = Literal['CA', 'CAA', 'CAL']


@dataclass
class ResultadoCondutor:
    catalogo: TipoCatalogo
    item: int
    nome_comercial: str
    bitola: str
    ampacidade_a: float
    diametro_mm: float
    rmg_m: float               # Ds do condutor unitário (self-GMD)
    r_ohm_km: float            # resistência CA do condutor unitário
    n_subcondutores: int       # 1, 2 ou 4 (Calc!G4/G5/G6)
    espacamento_feixe_cm: float
    r_eq_m: float              # raio equivalente do feixe (usado na capacitância/campo elétrico)
    ds_eq_m: float             # RMG equivalente do feixe (usado na indutância)
    r_fase_ohm_km: float       # resistência por fase (dividida pelo nº de subcondutores)


def buscar_condutor(catalogo: TipoCatalogo, item: Optional[int] = None,
                     nome_comercial: Optional[str] = None) -> dict:
    """Busca um condutor no catálogo pelo número do item ou pelo nome
    comercial (bitola). Origem: abas CA/CAA/CAL."""
    tabela = dados.CATALOGOS_CONDUTORES.get(catalogo)
    if tabela is None:
        raise ErroCalculo(f"Catálogo de condutor desconhecido: {catalogo!r} (use CA, CAA ou CAL).")
    if item is not None:
        for row in tabela:
            if row['item'] == item:
                return row
        raise ErroCalculo(f"Item {item} não encontrado no catálogo {catalogo}.")
    if nome_comercial is not None:
        for row in tabela:
            if str(row['tipo']).strip().lower() == nome_comercial.strip().lower():
                return row
        raise ErroCalculo(f"Condutor '{nome_comercial}' não encontrado no catálogo {catalogo}.")
    raise ErroCalculo("Informe item ou nome_comercial do condutor.")


def calcular_feixe_2(raio_m: float, espacamento_m: float) -> tuple[float, float]:
    """Feixe duplo (2 subcondutores). Origem: Calc!H51 (req), Calc!H52 (Ds_eq).
    req  = (r * d)^0.5
    Dseq = (Ds * d)^0.5
    Retorna (raio_equivalente_m, ) — usar com ds do condutor à parte.
    """
    return (raio_m * espacamento_m) ** 0.5


def calcular_feixe_4(raio_m: float, espacamento_m: float) -> float:
    """Feixe quádruplo. Origem: Calc!I51 (req), Calc!I52 (Ds_eq).
    req = (r * d^3 * sqrt(2))^0.25
    """
    return (raio_m / 1 * espacamento_m ** 3 * 2 ** 0.5) ** 0.25 if False else (raio_m * espacamento_m ** 3 * 2 ** 0.5) ** 0.25


def resolver_condutor(catalogo: TipoCatalogo, n_subcondutores: int,
                       item: Optional[int] = None, nome_comercial: Optional[str] = None,
                       espacamento_feixe_cm: float = 0.0) -> ResultadoCondutor:
    """Resolve o condutor + feixe completo (item 2 do mapa de cálculos)."""
    if n_subcondutores not in (1, 2, 4):
        raise ErroCalculo(
            "A planilha original suporta apenas feixe simples (1), duplo (2) ou "
            "quádruplo (4) subcondutores por fase (Calc!G4:G6). "
            "PREMISSA NÃO IDENTIFICADA para outras configurações."
        )
    row = buscar_condutor(catalogo, item=item, nome_comercial=nome_comercial)

    diametro_mm = row['diametro_mm']
    rmg_m = row['rmg_m']
    r_ohm_km = row.get('rca_ohm_km') or row.get('rca_60hz_75c_ohm_km')
    if r_ohm_km is None:
        raise ErroCalculo(f"Condutor {row['tipo']} sem resistência CA cadastrada no catálogo.")

    raio_m = (diametro_mm / 1000.0) / 2.0

    if n_subcondutores == 1:
        r_eq_m = raio_m
        ds_eq_m = rmg_m
        espacamento_feixe_cm = 0.0
    else:
        d_m = espacamento_feixe_cm / 100.0
        if d_m <= 0:
            raise ErroCalculo("Espaçamento entre subcondutores do feixe (d, em cm) é obrigatório para feixe duplo/quádruplo.")
        if n_subcondutores == 2:
            r_eq_m = calcular_feixe_2(raio_m, d_m)
            ds_eq_m = calcular_feixe_2(rmg_m, d_m)
        else:  # 4
            r_eq_m = calcular_feixe_4(raio_m, d_m)
            ds_eq_m = calcular_feixe_4(rmg_m, d_m)

    r_fase_ohm_km = r_ohm_km / n_subcondutores  # Calc!J5/J6 = J4/G

    return ResultadoCondutor(
        catalogo=catalogo, item=row['item'], nome_comercial=row['tipo'], bitola=str(row['bitola']),
        ampacidade_a=row['ampacidade'], diametro_mm=diametro_mm, rmg_m=rmg_m, r_ohm_km=r_ohm_km,
        n_subcondutores=n_subcondutores, espacamento_feixe_cm=espacamento_feixe_cm,
        r_eq_m=r_eq_m, ds_eq_m=ds_eq_m, r_fase_ohm_km=r_fase_ohm_km,
    )


# ---------------------------------------------------------------------------
# 3. PARÂMETROS ELÉTRICOS POR KM (R, L, G, C)
# ---------------------------------------------------------------------------

@dataclass
class ResultadoRLGC:
    r_ohm_km: float
    l_h_km: float
    g_s_km: float
    c_f_km: float
    eps_meio: float
    sigma_meio: float
    freq_hz: float
    omega_rad_s: float


def calcular_parametros_rlgc(r_fase_ohm_km: float, gmd_m: float, ds_eq_m: float, r_eq_m: float,
                              freq_hz: float = 60.0, eps_r: float = 1.0, sigma: float = 0.0
                              ) -> ResultadoRLGC:
    """Calcula R, L, G, C por km. Origem: Calc!A10, B10, C10, D10, E7, F10.

    L [H/km] = 2e-4 * ln(GMD / Ds_eq)          (Calc!B10)
    eps [F/m] = eps_r * eps0                    (Calc!E7)
    C [F/km] = 2*pi*eps / ln(GMD / r_eq) * 1000  (Calc!D10, usando 2*GMD/diâmetro = GMD/raio)
    G [S/km] = C * sigma / eps                   (Calc!C10; nulo para ar seco, sigma=0)
    """
    omega = 2 * math.pi * freq_hz  # Calc!F10
    eps = eps_r * EPS0             # Calc!E7 (Principal!G26 * 1e-9/(36*pi) == eps_r*EPS0)

    l_h_km = 2e-4 * math.log(gmd_m / ds_eq_m)          # Calc!B10
    c_f_km = (2 * math.pi * eps / math.log(gmd_m / r_eq_m)) * 1000  # Calc!D10
    g_s_km = c_f_km * sigma / eps if eps > 0 else 0.0  # Calc!C10

    return ResultadoRLGC(
        r_ohm_km=r_fase_ohm_km, l_h_km=l_h_km, g_s_km=g_s_km, c_f_km=c_f_km,
        eps_meio=eps, sigma_meio=sigma, freq_hz=freq_hz, omega_rad_s=omega,
    )


# ---------------------------------------------------------------------------
# 4. MODELO DE LINHA LONGA (PARÂMETROS DISTRIBUÍDOS)
# ---------------------------------------------------------------------------

@dataclass
class ResultadoLinhaLonga:
    comprimento_km: float
    z_km: complex          # Ω/km
    y_km: complex          # S/km
    zo_ohm: complex        # impedância característica
    gamma_km: complex      # constante de propagação (1/km)
    alpha_np_km: float
    beta_rad_km: float
    ro_ohm: float           # impedância de surto (real, sqrt(L/C))
    gl: complex
    cosh_gl: complex
    sinh_gl: complex
    velocidade_m_s: float
    velocidade_sem_perdas_m_s: float
    comprimento_onda_km: float


def calcular_linha_longa(rlgc: ResultadoRLGC, comprimento_km: float) -> ResultadoLinhaLonga:
    """Modelo de linha longa completo. Origem: Calc!A26,C26,A28,E28,A30,E30/F30,A32:D33,H16:H19."""
    z_km = complex(rlgc.r_ohm_km, rlgc.omega_rad_s * rlgc.l_h_km)       # Calc!A26
    y_km = complex(rlgc.g_s_km, rlgc.omega_rad_s * rlgc.c_f_km)         # Calc!C26

    zo = cmath.sqrt(z_km / y_km)     # Calc!A28
    gamma = cmath.sqrt(z_km * y_km)  # Calc!E28
    alpha = gamma.real               # Calc!C30
    beta = gamma.imag                # Calc!D30

    ro = math.sqrt(rlgc.l_h_km / rlgc.c_f_km)  # Calc!A30

    gl = gamma * comprimento_km      # Calc!E30/F30 combinados (complexo)
    cosh_gl = cmath.cosh(gl)         # Calc!A32:D33
    sinh_gl = cmath.sinh(gl)

    v = (rlgc.omega_rad_s / beta) * 1000 if beta != 0 else float('inf')      # Calc!H16 — m/s
    v_sp = 1 / math.sqrt(rlgc.l_h_km * 0.001 * rlgc.c_f_km * 0.001)          # Calc!H17 — m/s (sem perdas)
    comprimento_onda_km = v / rlgc.freq_hz / 1000 if beta != 0 else float('inf')  # Calc!H18

    return ResultadoLinhaLonga(
        comprimento_km=comprimento_km, z_km=z_km, y_km=y_km, zo_ohm=zo, gamma_km=gamma,
        alpha_np_km=alpha, beta_rad_km=beta, ro_ohm=ro, gl=gl, cosh_gl=cosh_gl, sinh_gl=sinh_gl,
        velocidade_m_s=v, velocidade_sem_perdas_m_s=v_sp, comprimento_onda_km=comprimento_onda_km,
    )


# ---------------------------------------------------------------------------
# 5. CIRCUITO PI EQUIVALENTE
# ---------------------------------------------------------------------------

@dataclass
class ResultadoCircuitoPi:
    zs_ohm: complex     # impedância série do pi (linha longa)
    yp_siemens: complex  # admitância shunt do pi (linha longa)
    zs_curta_ohm: complex  # referência: modelo de linha curta (R+jwL)*l


def calcular_circuito_pi(linha: ResultadoLinhaLonga, rlgc: ResultadoRLGC) -> ResultadoCircuitoPi:
    """Origem: PI-calc!B6 (Zs), PI-calc!B13 (Yp), PI-calc!B20 (Zs linha curta, referência)."""
    zs = linha.zo_ohm * linha.sinh_gl                           # PI-calc!B6
    yp = (1 / linha.zo_ohm) * cmath.tanh(linha.gl / 2)           # PI-calc!B13
    zs_curta = complex(rlgc.r_ohm_km, rlgc.omega_rad_s * rlgc.l_h_km) * linha.comprimento_km  # PI-calc!B20
    return ResultadoCircuitoPi(zs_ohm=zs, yp_siemens=yp, zs_curta_ohm=zs_curta)


# ---------------------------------------------------------------------------
# 6. CASO 1 — LINHA EM CARGA (dados no receptor: VCL, SC, cos phi)
# ---------------------------------------------------------------------------

@dataclass
class ResultadoCaso1:
    vc_fasor_v: complex
    ic_fasor_a: complex
    vs_fasor_v: complex
    is_fasor_a: complex
    vol_emissor_kv: float
    ang_vol_emissor_deg: float
    io_emissor_a: float
    ang_io_emissor_deg: float
    so_emissor_kva: float
    ang_so_emissor_deg: float
    po_emissor_kw: float
    pc_receptor_kw: float
    perdas_kw: float
    perdas_pct: float


def calcular_caso1(vcl_kv: float, ang_vcl_deg: float, sc_kva: float, cos_phi: float, atrasado: bool,
                    linha: ResultadoLinhaLonga) -> ResultadoCaso1:
    """Caso 1 — carga plena no receptor. Origem: Calc!C17,C18/A19,A38,A40,A42:E44,
    Principal!O21,O24,U4,U5."""
    if not (-1.0 <= cos_phi <= 1.0) or cos_phi == 0:
        raise ErroCalculo("cos(phi) deve estar no intervalo (-1, 1] e diferente de zero.")

    phi = math.acos(cos_phi) if atrasado else -math.acos(cos_phi)  # Calc!H9

    vc = cmath.rect(vcl_kv * 1000 / math.sqrt(3), math.radians(ang_vcl_deg))  # Calc!C17
    s_receptor = cmath.rect(sc_kva * 1000, phi)                                # Calc!A17 (aplicado ao receptor)
    ic = (s_receptor / (3 * vc)).conjugate()                                   # Calc!C18

    vs = vc * linha.cosh_gl + ic * linha.zo_ohm * linha.sinh_gl                # Calc!A38
    is_ = (vc / linha.zo_ohm) * linha.sinh_gl + ic * linha.cosh_gl             # Calc!A40

    vol_emissor_kv = abs(vs) * math.sqrt(3) / 1000                              # Calc!A42
    ang_vol_emissor = math.degrees(cmath.phase(vs))                             # Calc!B42
    io_emissor = abs(is_)                                                       # Calc!C42
    ang_io_emissor = math.degrees(cmath.phase(is_))                             # Calc!D42

    so = 3 * vs * is_.conjugate()                                               # Calc!E42
    so_kva = abs(so) / 1000                                                     # Calc!A44
    ang_so = math.degrees(cmath.phase(so))                                      # Calc!B44

    po_kw = so_kva * math.cos(math.radians(ang_so))                             # Principal!O21
    pc_kw = sc_kva * math.cos(phi)                                              # Principal!O24
    perdas_kw = po_kw - pc_kw                                                   # Principal!U4
    perdas_pct = (po_kw - pc_kw) / po_kw if po_kw != 0 else 0.0                 # Principal!U5

    return ResultadoCaso1(
        vc_fasor_v=vc, ic_fasor_a=ic, vs_fasor_v=vs, is_fasor_a=is_,
        vol_emissor_kv=vol_emissor_kv, ang_vol_emissor_deg=ang_vol_emissor,
        io_emissor_a=io_emissor, ang_io_emissor_deg=ang_io_emissor,
        so_emissor_kva=so_kva, ang_so_emissor_deg=ang_so,
        po_emissor_kw=po_kw, pc_receptor_kw=pc_kw, perdas_kw=perdas_kw, perdas_pct=perdas_pct,
    )


# ---------------------------------------------------------------------------
# 7. REGULAÇÃO (auto — energização em vazio a partir da tensão de emissor do Caso 1)
#    e CASO 2 manual (tensão de emissor especificada pelo usuário, opcional)
# ---------------------------------------------------------------------------

@dataclass
class ResultadoRegulacao:
    vol_vazio_kv: float
    ang_vol_vazio_deg: float
    regulacao_pct: float


def calcular_regulacao(caso1: ResultadoCaso1, vcl_kv: float, linha: ResultadoLinhaLonga) -> ResultadoRegulacao:
    """Regulação de tensão (%): compara a tensão que apareceria no receptor em
    vazio — mantendo a MESMA tensão de emissor do Caso 1 — com a tensão nominal
    em carga plena. Origem: Principal!U7,U9 / Calc!G30,G34 (ramo C12=0, automático).

    V_vazio = Vs(Caso1) / cosh(gamma*l)
    Regulação % = (|V_vazio_L| - VCL) / VCL
    """
    v_vazio = caso1.vs_fasor_v / linha.cosh_gl                         # Calc!G30 (ramo automático)
    vol_vazio_kv = abs(v_vazio) * math.sqrt(3) / 1000                   # Calc!G34
    ang_vol_vazio = math.degrees(cmath.phase(v_vazio))
    regulacao_pct = (vol_vazio_kv - vcl_kv) / vcl_kv                    # Principal!U9
    return ResultadoRegulacao(vol_vazio_kv, ang_vol_vazio, regulacao_pct)


@dataclass
class ResultadoCaso2Manual:
    vs_fasor_v: complex
    zl_ohm: complex       # impedância de carga equivalente do receptor (Caso 1: Vc/Ic)
    zin_ohm: complex      # impedância de entrada da linha terminada em ZL
    is_fasor_a: complex
    vr_fasor_v: complex
    ir_fasor_a: complex
    vol_receptor_kv: float
    ang_vol_receptor_deg: float
    io_receptor_a: float
    ang_io_receptor_deg: float
    sc_receptor_kva: float
    ang_sc_receptor_deg: float


def calcular_caso2_manual(vo_kv: float, ang_vo_deg: float, caso1: ResultadoCaso1,
                           linha: ResultadoLinhaLonga) -> ResultadoCaso2Manual:
    """Caso 2 (opcional/avançado) — energização com tensão de emissor
    ESPECIFICADA PELO USUÁRIO (Vo), mantendo a mesma impedância de carga
    equivalente do receptor do Caso 1 (ZL = Vc/Ic). Origem: Calc!A21 (Zin),
    Calc!A46:F55 (ramo M7<>0).

    ATENÇÃO — PREMISSA NÃO IDENTIFICADA NECESSITANDO VALIDAÇÃO: a planilha
    original usa a impedância de carga do Caso 1 (não um receptor em aberto)
    como terminação para este cálculo. Reproduzido fielmente aqui; confirmar
    com o usuário o significado de engenharia pretendido antes de usar este
    resultado para verificação de sobretensão de Ferranti (que classicamente
    assume receptor em circuito aberto).
    """
    zl = caso1.vc_fasor_v / caso1.ic_fasor_a  # Calc!G36 (impedância de carga do receptor)

    tanh_gl = cmath.tanh(linha.gl)
    zo = linha.zo_ohm
    zin = zo * (zl + zo * tanh_gl) / (zo + zl * tanh_gl)  # Calc!A21 (impedância de entrada terminada em ZL)

    vs = cmath.rect(vo_kv * 1000 / math.sqrt(3), math.radians(ang_vo_deg))
    is_ = vs / zin  # Calc!E14

    # Propagação inversa (emissor -> receptor), consistente com Calc!A49/A51
    vr = vs * linha.cosh_gl - is_ * zo * linha.sinh_gl
    ir = is_ * linha.cosh_gl - (vs / zo) * linha.sinh_gl

    vol_r_kv = abs(vr) * math.sqrt(3) / 1000
    ang_vr = math.degrees(cmath.phase(vr))
    io_r = abs(ir)
    ang_ir = math.degrees(cmath.phase(ir))

    sc = 3 * vr * ir.conjugate()
    sc_kva = abs(sc) / 1000
    ang_sc = math.degrees(cmath.phase(sc))

    return ResultadoCaso2Manual(
        vs_fasor_v=vs, zl_ohm=zl, zin_ohm=zin, is_fasor_a=is_, vr_fasor_v=vr, ir_fasor_a=ir,
        vol_receptor_kv=vol_r_kv, ang_vol_receptor_deg=ang_vr, io_receptor_a=io_r,
        ang_io_receptor_deg=ang_ir, sc_receptor_kva=sc_kva, ang_sc_receptor_deg=ang_sc,
    )


# ---------------------------------------------------------------------------
# 8. SIL (POTÊNCIA NATURAL) E CARREGAMENTO
# ---------------------------------------------------------------------------

@dataclass
class ResultadoSIL:
    sil_kw: float
    pct_sil: float


def calcular_sil(vcl_kv: float, ro_ohm: float, pc_receptor_kw: float) -> ResultadoSIL:
    """SIL = VCL^2 / Ro. Origem: Principal!U13 (SIL), V13 (%SIL = Pc/SIL)."""
    sil_kw = (vcl_kv ** 2 / ro_ohm) * 1000
    pct_sil = pc_receptor_kw / sil_kw if sil_kw != 0 else 0.0
    return ResultadoSIL(sil_kw=sil_kw, pct_sil=pct_sil)


# ---------------------------------------------------------------------------
# 9. CAMPO ELÉTRICO SUPERFICIAL (INDICADOR DE CORONA)
# ---------------------------------------------------------------------------

@dataclass
class ResultadoCampoEletrico:
    carga_c_m: float
    campo_v_m: float
    campo_kv_cm: float
    limite_referencia_kv_cm: float
    dentro_do_limite_referencia: bool


def calcular_campo_eletrico_superficial(vc_fasor_v: complex, c_f_km: float, r_eq_m: float,
                                         eps_meio: float) -> ResultadoCampoEletrico:
    """Campo elétrico na superfície do condutor. Origem: Calc!H13, H14, H15.

    Q = C(F/m) * Vfase(V)         (Calc!H13; H12=C em F/m = C_f_km/1000)
    E [V/m] = Q / (2*pi*eps*r_eq)  (Calc!H14; r_eq = raio equivalente do feixe)
    E [kV/cm] = E[V/m] * 1e-5
    """
    c_f_m = c_f_km / 1000.0
    q_c_m = c_f_m * abs(vc_fasor_v)                       # Calc!H13
    campo_v_m = q_c_m / (2 * math.pi * eps_meio * r_eq_m)  # Calc!H14
    campo_kv_cm = campo_v_m * 1e-5                          # Calc!H15
    limite = dados.CAMPO_ELETRICO_CRITICO_REFERENCIA_KV_CM
    return ResultadoCampoEletrico(
        carga_c_m=q_c_m, campo_v_m=campo_v_m, campo_kv_cm=campo_kv_cm,
        limite_referencia_kv_cm=limite, dentro_do_limite_referencia=campo_kv_cm < limite,
    )


# ---------------------------------------------------------------------------
# 10. RESULTADO COMPLETO (agregador)
# ---------------------------------------------------------------------------

@dataclass
class EntradaLT:
    # Identificação
    nome_linha: str
    # Condutor
    catalogo_condutor: TipoCatalogo
    item_condutor: Optional[int]
    nome_condutor: Optional[str]
    n_subcondutores: int
    espacamento_feixe_cm: float
    # Estrutura
    tipo_estrutura: Optional[int]
    geometria_custom: Optional[dict]  # {'h_a':...,'h_b':...,'h_c':...,'v_a':...,'v_b':...,'v_c':...}
    # Linha / operação
    frequencia_hz: float
    comprimento_km: float
    vcl_kv: float
    ang_vcl_deg: float
    sc_kva: float
    cos_phi: float
    atrasado: bool
    # Ambiente (avançado, opcional)
    eps_r: float = 1.0
    sigma_s_m: float = 0.0
    # Caso 2 manual (avançado, opcional)
    vo_kv: Optional[float] = None
    ang_vo_deg: float = 0.0


@dataclass
class ResultadoLT:
    entrada: EntradaLT
    geometria: ResultadoGeometria
    condutor: ResultadoCondutor
    rlgc: ResultadoRLGC
    linha_longa: ResultadoLinhaLonga
    circuito_pi: ResultadoCircuitoPi
    caso1: ResultadoCaso1
    regulacao: ResultadoRegulacao
    sil: ResultadoSIL
    campo_eletrico: ResultadoCampoEletrico
    caso2_manual: Optional[ResultadoCaso2Manual] = None


def calcular_linha_completa(entrada: EntradaLT) -> ResultadoLT:
    """Orquestra o cálculo completo da LT, na ordem:
    Geometria -> Condutor/Feixe -> RLGC -> Linha longa -> Circuito pi ->
    Caso 1 -> Regulação -> SIL -> Campo elétrico -> (Caso 2 manual, opcional).
    """
    if entrada.tipo_estrutura is not None:
        geo = geometria_estrutura(tipo_estrutura=entrada.tipo_estrutura)
    else:
        g = entrada.geometria_custom or {}
        geo = geometria_estrutura(h_a=g.get('h_a'), h_b=g.get('h_b'), h_c=g.get('h_c'),
                                   v_a=g.get('v_a'), v_b=g.get('v_b'), v_c=g.get('v_c'))

    cond = resolver_condutor(
        entrada.catalogo_condutor, entrada.n_subcondutores,
        item=entrada.item_condutor, nome_comercial=entrada.nome_condutor,
        espacamento_feixe_cm=entrada.espacamento_feixe_cm,
    )

    rlgc = calcular_parametros_rlgc(
        cond.r_fase_ohm_km, geo.gmd_m, cond.ds_eq_m, cond.r_eq_m,
        freq_hz=entrada.frequencia_hz, eps_r=entrada.eps_r, sigma=entrada.sigma_s_m,
    )

    linha = calcular_linha_longa(rlgc, entrada.comprimento_km)
    pi_eq = calcular_circuito_pi(linha, rlgc)

    caso1 = calcular_caso1(entrada.vcl_kv, entrada.ang_vcl_deg, entrada.sc_kva,
                            entrada.cos_phi, entrada.atrasado, linha)

    regulacao = calcular_regulacao(caso1, entrada.vcl_kv, linha)
    sil = calcular_sil(entrada.vcl_kv, linha.ro_ohm, caso1.pc_receptor_kw)
    campo = calcular_campo_eletrico_superficial(caso1.vc_fasor_v, rlgc.c_f_km, cond.r_eq_m, rlgc.eps_meio)

    caso2 = None
    if entrada.vo_kv is not None:
        caso2 = calcular_caso2_manual(entrada.vo_kv, entrada.ang_vo_deg, caso1, linha)

    return ResultadoLT(
        entrada=entrada, geometria=geo, condutor=cond, rlgc=rlgc, linha_longa=linha,
        circuito_pi=pi_eq, caso1=caso1, regulacao=regulacao, sil=sil, campo_eletrico=campo,
        caso2_manual=caso2,
    )
