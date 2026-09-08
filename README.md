# BK Electrical Line Parameters

Software de engenharia para cálculo de parâmetros elétricos de Linhas de
Transmissão (LT) pelo **modelo de linha longa** (parâmetros distribuídos),
desenvolvido para a **BK Engenharia e Tecnologia**.

Reproduz fielmente a metodologia da planilha original
`CELT2IPOG_calculo_eletrico_de_LTs.xlsm` (mantida em `referencia/` para
auditoria/rastreabilidade), validada célula a célula — ver seção
**Validação** abaixo.

## O que o software calcula

- Geometria da estrutura (DAB, DBC, DCA, GMD) — catálogo de 15 estruturas
  típicas ou geometria customizada.
- Condutores — catálogos CA (AAC), CAA (ACSR) e CAL (AAAC), com feixe de
  subcondutores (simples, duplo ou quádruplo).
- Parâmetros elétricos por km: R, L, G, C.
- Modelo de linha longa completo: impedância característica (Zo),
  constante de propagação (γ), impedância de surto (Ro).
- Circuito π equivalente.
- Caso 1 — linha em carga plena (tensão/corrente/potência no emissor a
  partir dos dados do receptor).
- Regulação de tensão (%).
- SIL (Potência Natural) e carregamento (%SIL).
- Campo elétrico superficial do condutor (indicador de corona).
- Caso 2 (avançado/opcional) — energização com tensão de emissor
  especificada.
- Memória de cálculo completa (fórmula → substituição numérica →
  resultado) e geração de relatório técnico em Word (.docx), no padrão
  documental BK Engenharia.

## Instalação

Requer Python 3.10+.

```bash
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Executar o aplicativo

```bash
streamlit run app.py
```

Abre automaticamente no navegador (por padrão, http://localhost:8501).

## Rodar os testes / validação Excel × Python

```bash
pytest tests/ -v
```

Ou, para ver a tabela de validação diretamente no terminal:

```bash
python3 exemplo_validacao.py
```

A mesma validação também está disponível dentro do app, na página
**"7. Validação Excel × Python"**.

## Estrutura do projeto

```
app_parametros_lt/
├── app.py                  # Interface Streamlit (navegação, entrada, resultados)
├── calculos.py              # Motor de cálculo (engenharia) — sem dependência de Streamlit
├── dados.py                  # Catálogos de condutores (CA/CAA/CAL) e estruturas (1-15)
├── validacoes.py              # Validação das entradas do usuário
├── relatorio.py                # Geração da memória de cálculo em Word (.docx)
├── exemplo_validacao.py         # Caso-exemplo de fábrica da planilha + tabela de validação
├── requirements.txt
├── templates/
│   └── logo_bk.jpeg              # Logo BK Engenharia (usado na interface e no relatório)
├── tests/
│   └── test_calculos.py            # Testes automatizados (pytest) — inclui validação Excel×Python
└── referencia/
    ├── CELT2IPOG_calculo_eletrico_de_LTs.xlsm   # Planilha original (fonte de verdade)
    └── Mapa_Calculos_LT_CELT2IPOG.md              # Engenharia reversa célula-a-célula
```

## Premissas que necessitam validação da engenharia responsável

Ver seção **10. Observações e Premissas a Validar** do relatório Word
gerado, e a seção final do `Mapa_Calculos_LT_CELT2IPOG.md`. Resumo:

1. **Caso 2 (tensão de emissor especificada)** reproduz exatamente a
   planilha original (energização com a impedância de carga do Caso 1) —
   não é, estritamente, um cálculo de receptor em circuito aberto.
   Confirmar interpretação de engenharia antes de usar para verificação
   de sobretensão (efeito Ferranti).
2. **Campo elétrico superficial**: indicador informativo — a comparação
   com ~21,1 kV/cm é apenas uma referência de literatura, não um limite
   normativo aplicado automaticamente pela planilha original.
3. **Estruturas 13, 14 e 15** do catálogo não têm geometria cadastrada na
   planilha original — indisponíveis até cadastro futuro ou uso de
   geometria customizada.
4. **Quadro de carregabilidade típica** (%SIL × Regulação × Perdas):
   incluído como referência informativa, sem vínculo de fórmula com a
   linha calculada (conforme validado com o solicitante).

## Versão

v1.0.0 — BK Engenharia e Tecnologia — www.bk-engenharia.com
