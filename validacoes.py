"""
validacoes.py — BK Electrical Line Parameters

Validação das entradas do usuário ANTES de acionar calculos.py.
Não contém lógica de cálculo de engenharia nem de interface (Streamlit).

Retorna sempre uma lista de mensagens de erro (vazia = entradas válidas).
Mensagens escritas para o engenheiro usuário final, sem jargão de Python.
"""

from __future__ import annotations

from typing import Optional

import dados


def validar_identificacao(cliente: str, projeto: str, nome_documento: str) -> list[str]:
    erros = []
    if not projeto or not projeto.strip():
        erros.append("Informe o nome do projeto.")
    if not nome_documento or not nome_documento.strip():
        erros.append("Informe o nome do documento.")
    return erros


def validar_condutor(catalogo: str, item: Optional[int], n_subcondutores: int,
                      espacamento_feixe_cm: float) -> list[str]:
    erros = []
    if catalogo not in dados.CATALOGOS_CONDUTORES:
        erros.append(f"Catálogo de condutor inválido: {catalogo!r}. Use CA, CAA ou CAL.")
        return erros
    if item is None:
        erros.append("Selecione o condutor (bitola) no catálogo.")
    else:
        tabela = dados.CATALOGOS_CONDUTORES[catalogo]
        if not any(row['item'] == item for row in tabela):
            erros.append(f"Item {item} não existe no catálogo {catalogo}.")
    if n_subcondutores not in (1, 2, 4):
        erros.append("Número de subcondutores por fase deve ser 1 (simples), 2 (duplo) ou 4 (quádruplo) "
                      "— única configuração suportada pela metodologia original.")
    if n_subcondutores in (2, 4):
        if espacamento_feixe_cm is None or espacamento_feixe_cm <= 0:
            erros.append("Informe o espaçamento entre subcondutores do feixe (d, em cm) — obrigatório "
                          "para feixe duplo ou quádruplo.")
        elif espacamento_feixe_cm > 200:
            erros.append("Espaçamento entre subcondutores do feixe parece incompatível "
                          "(> 200 cm) — verifique a unidade (cm).")
    return erros


def validar_estrutura(tipo_estrutura: Optional[int], geometria_custom: Optional[dict]) -> list[str]:
    erros = []
    if tipo_estrutura is None and geometria_custom is None:
        erros.append("Selecione uma estrutura do catálogo ou informe geometria customizada.")
        return erros
    if tipo_estrutura is not None:
        est = dados.ESTRUTURAS_POR_TIPO.get(tipo_estrutura)
        if est is None:
            erros.append(f"Estrutura tipo {tipo_estrutura} não existe (use 1 a 15).")
        elif est['h_A'] is None:
            erros.append(f"Estrutura tipo {tipo_estrutura} não possui geometria cadastrada na planilha "
                         f"original. Escolha outra estrutura ou informe geometria customizada.")
    if geometria_custom is not None:
        campos = ['h_a', 'h_b', 'h_c', 'v_a', 'v_b', 'v_c']
        faltantes = [c for c in campos if geometria_custom.get(c) is None]
        if faltantes:
            erros.append(f"Geometria customizada incompleta — faltam: {', '.join(faltantes)}.")
        else:
            va, vb, vc = geometria_custom['v_a'], geometria_custom['v_b'], geometria_custom['v_c']
            if any(v < 0 for v in (va, vb, vc)):
                erros.append("Alturas das fases (v) não podem ser negativas.")
            ha, hb, hc = geometria_custom['h_a'], geometria_custom['h_b'], geometria_custom['h_c']
            if ha == hb == hc and va == vb == vc:
                erros.append("Geometria inconsistente: as três fases não podem ocupar exatamente "
                              "a mesma posição (distância nula entre fases).")
    return erros


def validar_linha(frequencia_hz: float, comprimento_km: float, vcl_kv: float, sc_kva: float,
                   cos_phi: float) -> list[str]:
    erros = []
    if frequencia_hz is None or frequencia_hz <= 0:
        erros.append("Frequência deve ser positiva (tipicamente 50 ou 60 Hz).")
    elif frequencia_hz not in (50.0, 60.0):
        erros.append("Frequência incomum (diferente de 50/60 Hz) — confirme se é intencional.")
    if comprimento_km is None or comprimento_km <= 0:
        erros.append("Comprimento da linha (l) deve ser maior que zero.")
    if vcl_kv is None or vcl_kv <= 0:
        erros.append("Tensão de linha no receptor (VCL) deve ser maior que zero.")
    if sc_kva is None or sc_kva <= 0:
        erros.append("Potência aparente no receptor (SC) deve ser maior que zero.")
    if cos_phi is None or not (-1.0 <= cos_phi <= 1.0) or cos_phi == 0:
        erros.append("Fator de potência cos(φ) deve estar entre -1 e 1, e diferente de zero.")
    return erros


def validar_ambiente(eps_r: float, sigma_s_m: float) -> list[str]:
    erros = []
    if eps_r is None or eps_r < 1.0:
        erros.append("Permissividade relativa do meio (εR) deve ser ≥ 1 (1 = ar).")
    if sigma_s_m is None or sigma_s_m < 0:
        erros.append("Condutividade do meio (σ) não pode ser negativa.")
    return erros


def validar_caso2_manual(vo_kv: Optional[float], ang_vo_deg: Optional[float]) -> list[str]:
    erros = []
    if vo_kv is not None and vo_kv <= 0:
        erros.append("Tensão de emissor especificada (Caso 2) deve ser maior que zero.")
    return erros


def validar_entrada_completa(entrada) -> list[str]:
    """Valida um objeto calculos.EntradaLT completo. Retorna lista de erros
    (vazia = pode calcular)."""
    erros: list[str] = []
    erros += validar_condutor(entrada.catalogo_condutor, entrada.item_condutor,
                               entrada.n_subcondutores, entrada.espacamento_feixe_cm)
    erros += validar_estrutura(entrada.tipo_estrutura, entrada.geometria_custom)
    erros += validar_linha(entrada.frequencia_hz, entrada.comprimento_km, entrada.vcl_kv,
                            entrada.sc_kva, entrada.cos_phi)
    erros += validar_ambiente(entrada.eps_r, entrada.sigma_s_m)
    erros += validar_caso2_manual(entrada.vo_kv, entrada.ang_vo_deg)
    return erros
