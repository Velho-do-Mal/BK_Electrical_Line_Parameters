"""
Testes automatizados — BK Electrical Line Parameters

1) Validação Excel × Python: reproduz o caso-exemplo da planilha original
   e compara cada resultado com o valor lido da célula do Excel (tolerância
   0,05% — cobre apenas resíduo de arredondamento de exibição do Excel).
2) Testes unitários de funções isoladas do motor de cálculo.
3) Testes de robustez (validações de entrada / exceções esperadas).
"""

import cmath
import math
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import calculos as c
import dados
from exemplo_validacao import gerar_tabela_validacao, ENTRADA_EXEMPLO


# ---------------------------------------------------------------------------
# 1. Validação Excel × Python (caso-exemplo de fábrica da planilha)
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("linha", gerar_tabela_validacao(), ids=lambda l: l.item)
def test_validacao_excel_python(linha):
    assert linha.status == "OK", (
        f"{linha.item}: Excel={linha.excel} Python={linha.python} "
        f"diff={linha.diferenca_pct:.4f}% (origem {linha.origem})"
    )


# ---------------------------------------------------------------------------
# 2. Testes unitários — geometria
# ---------------------------------------------------------------------------

def test_geometria_estrutura_catalogo():
    g = c.geometria_estrutura(tipo_estrutura=3)
    assert g.dab_m == pytest.approx(3.310966626228661)
    assert g.dbc_m == pytest.approx(1.7)
    assert g.dca_m == pytest.approx(3.310966626228661)
    assert g.gmd_m == pytest.approx((g.dab_m * g.dbc_m * g.dca_m) ** (1 / 3))


def test_geometria_estrutura_sem_cadastro_leva_a_erro():
    with pytest.raises(c.ErroCalculo):
        c.geometria_estrutura(tipo_estrutura=13)  # sem geometria cadastrada na planilha original


def test_geometria_custom():
    g = c.geometria_estrutura(h_a=-1.6, h_b=1.6, h_c=1.6, v_a=0.85, v_b=1.7, v_c=0)
    assert g.dab_m == pytest.approx(3.310966626228661)


# ---------------------------------------------------------------------------
# 3. Testes unitários — condutor / feixe
# ---------------------------------------------------------------------------

def test_buscar_condutor_por_item():
    row = c.buscar_condutor('CA', item=15)
    assert row['tipo'] == 'Tulip'
    assert row['bitola'] == 336


def test_buscar_condutor_por_nome():
    row = c.buscar_condutor('CAA', nome_comercial='Hawk')
    assert row['item'] == 23
    assert row['bitola'] == 477


def test_resolver_condutor_simples():
    cond = c.resolver_condutor('CA', 1, item=15)
    assert cond.r_eq_m == pytest.approx(0.0169 / 2)
    assert cond.ds_eq_m == pytest.approx(0.0064)
    assert cond.r_fase_ohm_km == pytest.approx(0.206)


def test_resolver_condutor_feixe_duplo():
    cond = c.resolver_condutor('CA', 2, item=15, espacamento_feixe_cm=45.0)
    raio = (0.0169 / 2)
    d = 0.45
    assert cond.r_eq_m == pytest.approx((raio * d) ** 0.5)
    assert cond.ds_eq_m == pytest.approx((0.0064 * d) ** 0.5)
    assert cond.r_fase_ohm_km == pytest.approx(0.206 / 2)


def test_feixe_invalido_gera_erro():
    with pytest.raises(c.ErroCalculo):
        c.resolver_condutor('CA', 3, item=15)  # 3 subcondutores não suportado pela planilha original


def test_feixe_sem_espacamento_gera_erro():
    with pytest.raises(c.ErroCalculo):
        c.resolver_condutor('CA', 2, item=15, espacamento_feixe_cm=0)


# ---------------------------------------------------------------------------
# 4. Teste de consistência física — linha longa
# ---------------------------------------------------------------------------

def test_zo_zero_gamma_para_linha_sem_perdas_reduz_a_ro():
    """Sem perdas (R=0, G=0), Zo deve ser puramente real e igual a Ro."""
    rlgc = c.calcular_parametros_rlgc(r_fase_ohm_km=0.0, gmd_m=2.6512631927148687,
                                       ds_eq_m=0.0064, r_eq_m=0.00845, freq_hz=60.0)
    linha = c.calcular_linha_longa(rlgc, comprimento_km=100.0)
    assert linha.zo_ohm.imag == pytest.approx(0.0, abs=1e-9)
    assert abs(linha.zo_ohm) == pytest.approx(linha.ro_ohm, rel=1e-9)


def test_linha_completa_end_to_end_nao_lanca_excecao():
    resultado = c.calcular_linha_completa(ENTRADA_EXEMPLO)
    assert resultado.sil.sil_kw > 0
    assert resultado.caso1.po_emissor_kw > resultado.caso1.pc_receptor_kw  # perdas positivas


# ---------------------------------------------------------------------------
# 5. Caso 1 — validações de entrada
# ---------------------------------------------------------------------------

def test_caso1_cos_phi_invalido():
    linha = c.calcular_linha_longa(
        c.calcular_parametros_rlgc(0.206, 2.65, 0.0064, 0.00845, 60.0), 100.0
    )
    with pytest.raises(c.ErroCalculo):
        c.calcular_caso1(72, 0, 1000, cos_phi=0, atrasado=True, linha=linha)


if __name__ == "__main__":
    sys.exit(pytest.main([__file__, "-v"]))
