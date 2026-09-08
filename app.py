"""
app.py — BK Electrical Line Parameters
Interface Streamlit. Nenhum cálculo de engenharia acontece aqui — toda a
lógica fica em calculos.py; validações em validacoes.py; relatório em
relatorio.py. Este arquivo cuida apenas de: navegação, entrada de dados,
apresentação de resultados e acionamento dos módulos acima.
"""

from __future__ import annotations

import base64
import io
import math
import os
from datetime import date

import streamlit as st

import calculos as c
import dados
import validacoes as v
from exemplo_validacao import ENTRADA_EXEMPLO, gerar_tabela_validacao

try:
    import relatorio
    RELATORIO_DISPONIVEL = True
except Exception:
    RELATORIO_DISPONIVEL = False

# ---------------------------------------------------------------------------
# CONFIGURAÇÃO DA PÁGINA / IDENTIDADE BK ENGINEERING TOOLS
# ---------------------------------------------------------------------------

APP_NOME = "BK Electrical Line Parameters"
APP_VERSAO = "1.0.0"
APP_DESCRICAO = "Parâmetros elétricos de Linhas de Transmissão — modelo de linha longa"

LOGO_PATH = os.path.join(os.path.dirname(__file__), "templates", "logo_bk.jpeg")

st.set_page_config(page_title=APP_NOME, page_icon=LOGO_PATH if os.path.exists(LOGO_PATH) else "⚡",
                    layout="wide")

BK_AZUL = "#1565A8"
BK_AZUL_ESCURO = "#0D3D63"
BK_CINZA = "#4A4A4A"

CSS = f"""
<style>
    .bk-header {{
        display: flex; align-items: center; gap: 16px;
        padding: 10px 0 18px 0; border-bottom: 3px solid {BK_AZUL};
        margin-bottom: 20px;
    }}
    .bk-header h1 {{
        color: {BK_AZUL_ESCURO}; font-size: 1.6rem; margin: 0; font-weight: 700;
    }}
    .bk-header p {{ color: {BK_CINZA}; margin: 0; font-size: 0.85rem; }}
    .bk-badge {{
        background: {BK_AZUL}; color: white; padding: 2px 10px; border-radius: 4px;
        font-size: 0.72rem; font-weight: 600; letter-spacing: 0.03em;
    }}
    div[data-testid="stMetricValue"] {{ font-size: 1.35rem; color: {BK_AZUL_ESCURO}; }}
    .stButton > button[kind="primary"] {{ background-color: {BK_AZUL}; }}
    .bk-secao {{
        color: {BK_AZUL_ESCURO}; font-weight: 700; font-size: 1.05rem;
        border-left: 4px solid {BK_AZUL}; padding-left: 10px; margin: 18px 0 8px 0;
    }}
    .bk-premissa {{
        background: #FFF6E5; border-left: 4px solid #E0A429; padding: 8px 12px;
        font-size: 0.85rem; border-radius: 4px; margin: 6px 0;
    }}
    .bk-footer {{ color: #999; font-size: 0.75rem; text-align: center; margin-top: 40px; }}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


def cabecalho():
    col_logo, col_txt = st.columns([1, 10])
    with col_logo:
        if os.path.exists(LOGO_PATH):
            st.image(LOGO_PATH, width=64)
    with col_txt:
        st.markdown(
            f"""<div class="bk-header" style="border-bottom:none;padding-bottom:0;">
            <div><h1>{APP_NOME} <span class="bk-badge">v{APP_VERSAO}</span></h1>
            <p>{APP_DESCRICAO} — BK Engenharia e Tecnologia</p></div>
            </div>""",
            unsafe_allow_html=True,
        )
    st.markdown(f'<hr style="border:1px solid {BK_AZUL};margin-top:0;">', unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# ESTADO DE SESSÃO (por usuário — nunca variáveis globais mutáveis)
# ---------------------------------------------------------------------------

def _init_state():
    defaults = {
        "identificacao": {"cliente": "", "projeto": "", "nome_documento": "Parâmetros Elétricos de LT",
                           "numero_documento": "", "revisao": "0", "data": date.today().isoformat(),
                           "responsavel_tecnico": "Márcio Nunes Knopp"},
        "resultado": None,
        "erros_validacao": [],
        "usar_geometria_custom": False,
        "usar_caso2": False,
    }
    for k, val in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = val


_init_state()


# ---------------------------------------------------------------------------
# BARRA LATERAL — NAVEGAÇÃO
# ---------------------------------------------------------------------------

with st.sidebar:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=56)
    st.markdown(f"### {APP_NOME}")
    st.caption("BK Engineering Tools")
    st.divider()
    pagina = st.radio(
        "Navegação",
        ["1. Identificação", "2. Entrada de Dados", "3. Cálculo", "4. Resultados",
         "5. Memória de Cálculo", "6. Relatório Word", "7. Validação Excel × Python"],
        label_visibility="collapsed",
    )
    st.divider()
    if st.button("🔄 Novo cálculo (limpar)", use_container_width=True):
        st.session_state["resultado"] = None
        st.rerun()
    st.caption(f"© BK Engenharia e Tecnologia — v{APP_VERSAO}")


# ---------------------------------------------------------------------------
# PÁGINA 1 — IDENTIFICAÇÃO
# ---------------------------------------------------------------------------

if pagina == "1. Identificação":
    cabecalho()
    st.markdown('<div class="bk-secao">Identificação do Documento</div>', unsafe_allow_html=True)
    ident = st.session_state["identificacao"]
    col1, col2 = st.columns(2)
    with col1:
        ident["cliente"] = st.text_input("Cliente", ident["cliente"])
        ident["projeto"] = st.text_input("Projeto", ident["projeto"])
        ident["nome_documento"] = st.text_input("Nome do documento", ident["nome_documento"])
    with col2:
        ident["numero_documento"] = st.text_input("Número do documento", ident["numero_documento"])
        ident["revisao"] = st.text_input("Revisão", ident["revisao"])
        ident["data"] = st.text_input("Data", ident["data"])
    ident["responsavel_tecnico"] = st.text_input("Responsável técnico", ident["responsavel_tecnico"])
    st.session_state["identificacao"] = ident
    st.info("Preencha a identificação e siga para **2. Entrada de Dados**.")


# ---------------------------------------------------------------------------
# PÁGINA 2 — ENTRADA DE DADOS
# ---------------------------------------------------------------------------

elif pagina == "2. Entrada de Dados":
    cabecalho()

    st.markdown('<div class="bk-secao">Linha</div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    with col1:
        nome_linha = st.text_input("Nome da linha", "Exemplo")
        frequencia_hz = st.selectbox("Frequência (Hz)", [60.0, 50.0], index=0)
    with col2:
        comprimento_km = st.number_input("Comprimento da linha, l (km)", min_value=0.001, value=100.0, step=1.0)
        vcl_kv = st.number_input("Tensão de linha no receptor, VCL (kV)", min_value=0.001, value=72.0)
    with col3:
        sc_kva = st.number_input("Potência aparente no receptor, SC (kVA)", min_value=0.001, value=1000.0)
        cos_phi = st.number_input("Fator de potência, cos(φ)", min_value=-1.0, max_value=1.0, value=1.0, step=0.01)
    atrasado = st.radio("Natureza do fator de potência (receptor)", ["Atrasado (indutivo)", "Adiantado (capacitivo)"],
                         horizontal=True) == "Atrasado (indutivo)"
    ang_vcl_deg = st.number_input("Ângulo de referência de VCL (°)", value=0.0,
                                   help="Normalmente 0° — referência angular do sistema.")

    st.markdown('<div class="bk-secao">Condutor</div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    with col1:
        catalogo = st.selectbox("Catálogo", list(dados.NOMES_CATALOGOS.keys()),
                                 format_func=lambda k: dados.NOMES_CATALOGOS[k])
    tabela_cat = dados.CATALOGOS_CONDUTORES[catalogo]
    opcoes_condutor = {f"{r['item']:>3} — {r['tipo']} ({r['bitola']})": r['item'] for r in tabela_cat}
    with col2:
        escolha = st.selectbox("Bitola / nome comercial", list(opcoes_condutor.keys()))
        item_condutor = opcoes_condutor[escolha]
    condutor_sel = next(r for r in tabela_cat if r['item'] == item_condutor)
    with col3:
        st.metric("Ampacidade catálogo", f"{condutor_sel['ampacidade']:.0f} A")

    col1, col2 = st.columns(2)
    with col1:
        n_subcondutores = st.selectbox("Nº de subcondutores por fase (feixe)", [1, 2, 4],
                                        format_func=lambda n: {1: "1 — Simples", 2: "2 — Duplo", 4: "4 — Quádruplo"}[n])
    with col2:
        espacamento_feixe_cm = 0.0
        if n_subcondutores > 1:
            espacamento_feixe_cm = st.number_input("Espaçamento entre subcondutores do feixe, d (cm)",
                                                     min_value=0.0, value=45.0)

    st.markdown('<div class="bk-secao">Estrutura (geometria)</div>', unsafe_allow_html=True)
    usar_custom = st.checkbox("Usar geometria customizada (em vez do catálogo de 15 estruturas)",
                               value=st.session_state["usar_geometria_custom"])
    st.session_state["usar_geometria_custom"] = usar_custom
    tipo_estrutura = None
    geometria_custom = None
    if not usar_custom:
        estruturas_validas = [e for e in dados.ESTRUTURAS if e['h_A'] is not None]
        opcoes_est = {f"Estrutura {e['tipo']} (Dab={e['dab']:.2f} / Dbc={e['dbc']:.2f} / Dca={e['dca']:.2f} m)": e['tipo']
                      for e in estruturas_validas}
        escolha_est = st.selectbox("Tipo de estrutura (catálogo)", list(opcoes_est.keys()), index=2)
        tipo_estrutura = opcoes_est[escolha_est]
    else:
        st.caption("Coordenadas (m) — referência horizontal (h) e vertical (v) de cada fase.")
        col1, col2, col3 = st.columns(3)
        with col1:
            h_a = st.number_input("h — Fase A (m)", value=-1.6)
            v_a = st.number_input("v — Fase A (m)", value=0.85, min_value=0.0)
        with col2:
            h_b = st.number_input("h — Fase B (m)", value=1.6)
            v_b = st.number_input("v — Fase B (m)", value=1.7, min_value=0.0)
        with col3:
            h_c = st.number_input("h — Fase C (m)", value=1.6)
            v_c = st.number_input("v — Fase C (m)", value=0.0, min_value=0.0)
        geometria_custom = {"h_a": h_a, "h_b": h_b, "h_c": h_c, "v_a": v_a, "v_b": v_b, "v_c": v_c}

    with st.expander("Ambiente (avançado — opcional, padrão = ar seco)"):
        col1, col2 = st.columns(2)
        with col1:
            eps_r = st.number_input("Permissividade relativa do meio, εR", min_value=1.0, value=1.0)
        with col2:
            sigma_s_m = st.number_input("Condutividade do meio, σ (S/m)", min_value=0.0, value=0.0, format="%.6f")

    with st.expander("Caso 2 — energização com tensão de emissor especificada (avançado — opcional)"):
        st.markdown(
            '<div class="bk-premissa">⚠️ PREMISSA NÃO IDENTIFICADA — NECESSITA VALIDAÇÃO: este cálculo '
            'reproduz fielmente a planilha original (energização com a mesma impedância de carga do '
            'Caso 1, alimentada por uma tensão de emissor diferente). Confirme com a engenharia responsável '
            'se este é o comportamento esperado antes de usar para verificação de sobretensão de Ferranti.</div>',
            unsafe_allow_html=True,
        )
        usar_caso2 = st.checkbox("Calcular Caso 2 (tensão de emissor especificada)",
                                  value=st.session_state["usar_caso2"])
        st.session_state["usar_caso2"] = usar_caso2
        vo_kv, ang_vo_deg = None, 0.0
        if usar_caso2:
            col1, col2 = st.columns(2)
            with col1:
                vo_kv = st.number_input("Tensão de linha no emissor, Vo (kV)", min_value=0.001, value=vcl_kv)
            with col2:
                ang_vo_deg = st.number_input("Ângulo de Vo (°)", value=0.0)

    entrada = c.EntradaLT(
        nome_linha=nome_linha, catalogo_condutor=catalogo, item_condutor=item_condutor, nome_condutor=None,
        n_subcondutores=n_subcondutores, espacamento_feixe_cm=espacamento_feixe_cm,
        tipo_estrutura=tipo_estrutura, geometria_custom=geometria_custom,
        frequencia_hz=frequencia_hz, comprimento_km=comprimento_km,
        vcl_kv=vcl_kv, ang_vcl_deg=ang_vcl_deg, sc_kva=sc_kva, cos_phi=cos_phi, atrasado=atrasado,
        eps_r=eps_r, sigma_s_m=sigma_s_m, vo_kv=vo_kv, ang_vo_deg=ang_vo_deg,
    )
    st.session_state["entrada_lt"] = entrada

    st.divider()
    erros = v.validar_entrada_completa(entrada)
    erros += v.validar_identificacao(st.session_state["identificacao"]["cliente"],
                                      st.session_state["identificacao"]["projeto"],
                                      st.session_state["identificacao"]["nome_documento"])
    if erros:
        for e in erros:
            st.warning(e)

    col1, col2 = st.columns([1, 4])
    with col1:
        calcular = st.button("⚡ CALCULAR", type="primary", use_container_width=True, disabled=bool(erros))
    if calcular:
        try:
            st.session_state["resultado"] = c.calcular_linha_completa(entrada)
            st.success("Cálculo executado com sucesso. Acesse **4. Resultados**.")
        except c.ErroCalculo as ex:
            st.error(f"Erro de engenharia: {ex}")


# ---------------------------------------------------------------------------
# PÁGINA 3 — CÁLCULO (status)
# ---------------------------------------------------------------------------

elif pagina == "3. Cálculo":
    cabecalho()
    st.markdown('<div class="bk-secao">Executar Cálculo</div>', unsafe_allow_html=True)
    entrada = st.session_state.get("entrada_lt")
    if entrada is None:
        st.warning("Preencha a **2. Entrada de Dados** primeiro.")
    else:
        st.write("Entradas atuais:")
        st.json({
            "linha": entrada.nome_linha, "condutor": f"{entrada.catalogo_condutor} item {entrada.item_condutor}",
            "n_subcondutores": entrada.n_subcondutores, "estrutura": entrada.tipo_estrutura,
            "comprimento_km": entrada.comprimento_km, "vcl_kv": entrada.vcl_kv, "sc_kva": entrada.sc_kva,
            "cos_phi": entrada.cos_phi,
        })
        if st.button("⚡ CALCULAR", type="primary"):
            erros = v.validar_entrada_completa(entrada)
            if erros:
                for e in erros:
                    st.error(e)
            else:
                try:
                    st.session_state["resultado"] = c.calcular_linha_completa(entrada)
                    st.success("Cálculo executado com sucesso.")
                except c.ErroCalculo as ex:
                    st.error(f"Erro de engenharia: {ex}")


# ---------------------------------------------------------------------------
# PÁGINA 4 — RESULTADOS
# ---------------------------------------------------------------------------

elif pagina == "4. Resultados":
    cabecalho()
    r: c.ResultadoLT | None = st.session_state.get("resultado")
    if r is None:
        st.warning("Nenhum resultado calculado ainda. Vá em **2. Entrada de Dados** e clique em CALCULAR.")
    else:
        st.markdown('<div class="bk-secao">Geometria</div>', unsafe_allow_html=True)
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("DAB", f"{r.geometria.dab_m:.4f} m")
        col2.metric("DBC", f"{r.geometria.dbc_m:.4f} m")
        col3.metric("DCA", f"{r.geometria.dca_m:.4f} m")
        col4.metric("GMD (Deq)", f"{r.geometria.gmd_m:.4f} m")

        try:
            import matplotlib.pyplot as plt
            fig, ax = plt.subplots(figsize=(5, 3.2))
            xs = [r.geometria.h_a, r.geometria.h_b, r.geometria.h_c]
            ys = [r.geometria.v_a, r.geometria.v_b, r.geometria.v_c]
            labels = ["A", "B", "C"]
            ax.scatter(xs, ys, s=180, color=BK_AZUL, zorder=3)
            for x, y, lab in zip(xs, ys, labels):
                ax.annotate(lab, (x, y), textcoords="offset points", xytext=(0, 12),
                            ha="center", fontsize=11, fontweight="bold", color=BK_AZUL_ESCURO)
            ax.axhline(0, color="#888", linewidth=1)
            ax.set_xlabel("Distância horizontal (m)")
            ax.set_ylabel("Altura (m)")
            ax.set_title("Geometria da estrutura (silhueta)", fontsize=10)
            ax.grid(alpha=0.3)
            ax.set_ylim(bottom=-0.5)
            st.pyplot(fig, use_container_width=False)
        except Exception:
            pass

        st.markdown('<div class="bk-secao">Condutor Selecionado</div>', unsafe_allow_html=True)
        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("Condutor", r.condutor.nome_comercial)
        col2.metric("Bitola", r.condutor.bitola)
        col3.metric("Diâmetro", f"{r.condutor.diametro_mm:.2f} mm")
        col4.metric("Ampacidade", f"{r.condutor.ampacidade_a:.0f} A")
        col5.metric("Subcondutores/fase", r.condutor.n_subcondutores)

        st.markdown('<div class="bk-secao">Parâmetros Elétricos (por km)</div>', unsafe_allow_html=True)
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("R", f"{r.rlgc.r_ohm_km:.5f} Ω/km")
        col2.metric("L", f"{r.rlgc.l_h_km*1000:.5f} mH/km")
        col3.metric("C", f"{r.rlgc.c_f_km*1e9:.5f} nF/km")
        col4.metric("G", f"{r.rlgc.g_s_km:.2e} S/km")

        st.markdown('<div class="bk-secao">Características da Linha</div>', unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)
        col1.metric("|Zo| (impedância característica)", f"{abs(r.linha_longa.zo_ohm):.3f} Ω",
                    help=f"∠ {math.degrees(math.atan2(r.linha_longa.zo_ohm.imag, r.linha_longa.zo_ohm.real)):.3f}°")
        col2.metric("Ro (impedância de surto)", f"{r.linha_longa.ro_ohm:.3f} Ω")
        col3.metric("γ·l (propagação)", f"{r.linha_longa.gl.real:.5f} + j{r.linha_longa.gl.imag:.5f}")

        st.markdown('<div class="bk-secao">Circuito π Equivalente</div>', unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        col1.metric("Zs (série)", f"{abs(r.circuito_pi.zs_ohm):.3f} Ω ∠ "
                    f"{math.degrees(math.atan2(r.circuito_pi.zs_ohm.imag, r.circuito_pi.zs_ohm.real)):.2f}°")
        col2.metric("Yp (shunt)", f"{abs(r.circuito_pi.yp_siemens)*1000:.5f} mS")

        st.markdown('<div class="bk-secao">Fonte (Emissor) — Caso 1</div>', unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)
        col1.metric("|VoL|", f"{r.caso1.vol_emissor_kv:.4f} kV", help=f"∠ {r.caso1.ang_vol_emissor_deg:.4f}°")
        col2.metric("|Io|", f"{r.caso1.io_emissor_a:.4f} A", help=f"∠ {r.caso1.ang_io_emissor_deg:.4f}°")
        col3.metric("|So|", f"{r.caso1.so_emissor_kva:.4f} kVA", help=f"∠ {r.caso1.ang_so_emissor_deg:.4f}°")

        st.markdown('<div class="bk-secao">Carga Plena (Receptor) — Caso 1</div>', unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)
        col1.metric("VCL", f"{r.entrada.vcl_kv:.4f} kV")
        col2.metric("SC", f"{r.entrada.sc_kva:.4f} kVA")
        col3.metric("Pc", f"{r.caso1.pc_receptor_kw:.4f} kW")

        st.markdown('<div class="bk-secao">Desempenho Elétrico</div>', unsafe_allow_html=True)
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Po (emissor)", f"{r.caso1.po_emissor_kw:.4f} kW")
        col2.metric("Perdas", f"{r.caso1.perdas_kw:.4f} kW", delta=f"{r.caso1.perdas_pct*100:.4f} %",
                    delta_color="inverse")
        col3.metric("Regulação de tensão", f"{r.regulacao.regulacao_pct*100:.4f} %")
        col4.metric("SIL", f"{r.sil.sil_kw:.2f} kW")

        col1, col2 = st.columns(2)
        col1.metric("%SIL (carregamento)", f"{r.sil.pct_sil*100:.2f} %")
        with col2:
            campo_status = "✅ OK" if r.campo_eletrico.dentro_do_limite_referencia else "⚠️ ACIMA DA REFERÊNCIA"
            st.metric("Campo elétrico superficial", f"{r.campo_eletrico.campo_kv_cm:.4f} kV/cm", help=campo_status)

        with st.expander("📊 Quadro de referência de carregabilidade típica (informativo)"):
            st.caption("Tabela de referência sem vínculo direto com esta linha — apenas comparação "
                       "com valores típicos de regulação/perdas em função do carregamento (%SIL).")
            st.table([
                {"%SIL": f"{row['pct_sil']*100:.0f}%", "Regulação (%) típica": f"{row['regulacao_pct']*100:.2f}%",
                 "Perdas (%) típica": f"{row['perdas_pct']*100:.2f}%"}
                for row in dados.TABELA_CARREGABILIDADE_REFERENCIA
            ])

        if r.caso2_manual is not None:
            st.markdown('<div class="bk-secao">Caso 2 — Tensão de Emissor Especificada</div>', unsafe_allow_html=True)
            st.markdown(
                '<div class="bk-premissa">⚠️ Ver observação de premissa não identificada na aba de entrada.</div>',
                unsafe_allow_html=True)
            col1, col2, col3 = st.columns(3)
            col1.metric("Tensão no receptor", f"{r.caso2_manual.vol_receptor_kv:.4f} kV")
            col2.metric("Corrente no receptor", f"{r.caso2_manual.io_receptor_a:.4f} A")
            col3.metric("Potência no receptor", f"{r.caso2_manual.sc_receptor_kva:.4f} kVA")


# ---------------------------------------------------------------------------
# PÁGINA 5 — MEMÓRIA DE CÁLCULO
# ---------------------------------------------------------------------------

elif pagina == "5. Memória de Cálculo":
    cabecalho()
    r = st.session_state.get("resultado")
    if r is None:
        st.warning("Nenhum resultado calculado ainda. Vá em **2. Entrada de Dados** e clique em CALCULAR.")
    else:
        if RELATORIO_DISPONIVEL:
            etapas = relatorio.montar_memoria_calculo(r)
            for et in etapas:
                with st.expander(f"**{et.titulo}**"):
                    st.markdown(f"**Fórmula:** {et.formula}")
                    if et.onde:
                        st.markdown(f"**Onde:** {et.onde}")
                    st.markdown(f"**Dados utilizados:** {et.dados}")
                    st.markdown(f"**Substituição numérica:** {et.substituicao}")
                    st.markdown(f"**Resultado:** {et.resultado}")
                    if et.interpretacao:
                        st.caption(et.interpretacao)
        else:
            st.error("Módulo de relatório indisponível.")


# ---------------------------------------------------------------------------
# PÁGINA 6 — RELATÓRIO WORD
# ---------------------------------------------------------------------------

elif pagina == "6. Relatório Word":
    cabecalho()
    r = st.session_state.get("resultado")
    if r is None:
        st.warning("Nenhum resultado calculado ainda. Vá em **2. Entrada de Dados** e clique em CALCULAR.")
    elif not RELATORIO_DISPONIVEL:
        st.error("Módulo de relatório indisponível (verifique instalação de python-docx).")
    else:
        st.markdown('<div class="bk-secao">Gerar Memória de Cálculo em Word</div>', unsafe_allow_html=True)
        ident = st.session_state["identificacao"]
        st.write("Documento a ser gerado com os dados de identificação preenchidos na página 1.")
        st.table([ident])
        if st.button("📄 GERAR RELATÓRIO WORD", type="primary"):
            with st.spinner("Gerando documento..."):
                doc_bytes = relatorio.gerar_relatorio_docx(r, ident)
            nome_arquivo = relatorio.nome_arquivo_relatorio(ident)
            st.success("Relatório gerado com sucesso.")
            st.download_button("⬇️ Baixar .docx", data=doc_bytes, file_name=nome_arquivo,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                type="primary")


# ---------------------------------------------------------------------------
# PÁGINA 7 — VALIDAÇÃO EXCEL × PYTHON
# ---------------------------------------------------------------------------

elif pagina == "7. Validação Excel × Python":
    cabecalho()
    st.markdown('<div class="bk-secao">Validação Excel × Python</div>', unsafe_allow_html=True)
    st.write(
        "Reproduz o caso-exemplo de fábrica da planilha original "
        "(`CELT2IPOG_calculo_eletrico_de_LTs.xlsm`) e compara cada resultado calculado "
        "pelo Python com o valor lido diretamente da célula do Excel."
    )
    if st.button("▶️ Rodar validação"):
        linhas = gerar_tabela_validacao()
        n_ok = sum(1 for l in linhas if l.status == "OK")
        st.metric("Itens validados", f"{n_ok} / {len(linhas)}")
        st.table([
            {"Item": l.item, "Origem (Excel)": l.origem, "Excel": f"{l.excel:.6f}", "Python": f"{l.python:.6f}",
             "Unidade": l.unidade, "Diferença %": f"{l.diferenca_pct:.6f}", "Status": l.status}
            for l in linhas
        ])
        if n_ok == len(linhas):
            st.success("Todos os itens dentro da tolerância — software valida corretamente a metodologia original.")
        else:
            st.error("Há divergências acima da tolerância — revisar antes de liberar o software.")


st.markdown(f'<div class="bk-footer">BK Engenharia e Tecnologia · www.bk-engenharia.com · '
            f'{APP_NOME} v{APP_VERSAO}</div>', unsafe_allow_html=True)
