# Mapa Completo de Cálculos — CELT2IPOG (Cálculo Elétrico de LTs)
Engenharia reversa da planilha `CELT2IPOG_calculo_eletrico_de_LTs.xlsm`

Abas identificadas: **Principal** (entradas/resultados), **Circuito PI** e **PI-calc** (circuito π equivalente),
**Calc** (motor de cálculo, oculta), **CA / CAA / CAL** (catálogos de condutores, ocultas),
**Estruturas** (geometria de 15 silhuetas de estrutura), **Análises** (tabela de referência %SIL x R% x P%),
**Relatório** (quadro-resumo dentro da própria planilha).

---

## 1. GEOMETRIA DA LINHA

| # | Cálculo | Origem (aba!célula) | Fórmula Excel | Variáveis / Unidade | Fórmula Python | Observação |
|---|---|---|---|---|---|---|
|1.1| Seleção da estrutura | `Principal!E10` (input 1–15) | manual | tipo estrutura [-] | `tipo_estrutura: int` | índice na tabela `Estruturas!AE4:AG18` |
|1.2| Catálogo de estruturas (h, v por fase) | `Estruturas!Y4:AD18` | coordenadas cadastradas (m) | h_A,h_B,h_C, v_A,v_B,v_C [m] | tabela estática `ESTRUTURAS` em `dados.py` | 15 silhuetas típicas (A/B/C = fases) |
|1.3| DAB | `Estruturas!AE` | `=((hA-hB)^2+(vA-vB)^2)^0.5` | h,v [m] | `math.hypot(hA-hB, vA-vB)` | distância fase A–B |
|1.4| DBC | `Estruturas!AF` | `=((hB-hC)^2+(vB-vC)^2)^0.5` | idem | idem | distância fase B–C |
|1.5| DCA | `Estruturas!AG` | `=((hC-hA)^2+(vC-vA)^2)^0.5` | idem | idem | distância fase C–A |
|1.6| DAB/DBC/DCA aplicados | `Principal!E15,H17,F19` | `=INDEX(Estruturas!$AE$4:$AG$18,E10,1/2/3)` | [m] | lookup em `ESTRUTURAS[tipo]` | trazem a estrutura selecionada para o cálculo |
|1.7| GMD (Deq) | `Calc!D7` | `=(DAB*DBC*DCA)^(1/3)` | DAB,DBC,DCA [m] | `(dab*dbc*dca)**(1/3)` | Distância Média Geométrica entre fases |

## 2. CONDUTORES (CATÁLOGO E BITOLA SELECIONADA)

| # | Cálculo | Origem | Fórmula | Variáveis | Python | Observação |
|---|---|---|---|---|---|---|
|2.1| Catálogos de condutores | abas `CA` (AAC, 58 itens), `CAA` (ACSR, 76 itens), `CAL` (AAAC, 36 itens) | tabela estática | Bitola, Tipo(nome comercial), Ampacidade[A], Seção[mm²], Diâmetro[mm], RMG(Ds)[m], Rcc[Ω/km], Rca 25/75°C[Ω/km] | 3 tabelas estáticas em `dados.py` | fonte de verdade dos condutores comerciais |
|2.2| Seleção do tipo de condutor | `Calc!M6` / `Principal!C4` (1=CA,2=CAA,3=CAL) | `INDEX($N$3:$U$5,M6,col)` | tipo [-] | lookup | |
|2.3| Seleção do item (bitola) dentro do catálogo | `CA!/CAA!/CAL!  V3:AC3` etc. | `INDEX($A$5:$U$80,item,col)` | item [-] | lookup | número do item no catálogo |
|2.4| Diâmetro do condutor (d) | `Calc!D4=P6/1000` | mm→m | d [m] | `d_mm/1000` | |
|2.5| RMG do condutor (Ds) | `Calc!E4=Q6` | direto do catálogo | Ds [m] | — | Raio Médio Geométrico (self-GMD) |
|2.6| Resistência do condutor | `Calc!F4=R6` | direto do catálogo (Rca) | R [Ω/km] | — | CA 60Hz, referência térmica do catálogo |
|2.7| Nº de subcondutores por fase (feixe) | `Calc!G4/G5/G6` (1,2,4) — `Principal!C22` | manual | n [-] | `n_subcond: int` | feixe simples, duplo ou quádruplo |
|2.8| RMG equivalente do feixe (2 subcond.) | `Calc!H51=(H48/2*H50)^0.5` | `req=(r*d)^0.5` | r=diâmetro/2 [m], d=espaçamento entre subcond. [m] | `sqrt(r*d)` | fórmula clássica de feixe (2 cabos) |
|2.9| RMG equivalente do feixe (4 subcond.) | `Calc!I51=(H48/2*H50^3*2^0.5)^0.25` | `req=(r*d³*√2)^0.25` | idem | `(r*d**3*2**0.5)**0.25` | feixe quádruplo |
|2.10| Ds equivalente do feixe (2 subcond.) | `Calc!H52=(H49*H50)^0.5` | `Ds_eq=(Ds*d)^0.5` | Ds condutor, d espaçamento [m] | `sqrt(Ds*d)` | |
|2.11| Ds equivalente do feixe (4 subcond.) | `Calc!I52=(H49*H50^3*2^0.5)^0.25` | `Ds_eq=(Ds*d³*√2)^0.25` | idem | idem | |
|2.12| R por fase (feixe) | `Calc!J5/J6 = J4/G` | `R_fase=R_cond/n` | R[Ω/km], n | `r_cond/n` | resistência dividida pelo nº de subcondutores |
|2.13| Seleção final r/Ds/R aplicados | `Calc!H7,I7,J7` | `IF(H50=0,H4,INDEX(...))` | | | H50 = espaçamento do feixe (cm/100); se 0 → condutor simples |

## 3. PARÂMETROS ELÉTRICOS POR UNIDADE DE COMPRIMENTO (R, L, G, C)

| # | Cálculo | Origem | Fórmula | Variáveis/Unid. | Python | Observação |
|---|---|---|---|---|---|---|
|3.1| Resistência série | `Calc!A10=J7` | direto | R [Ω/km] | — | |
|3.2| Indutância série | `Calc!B10` | `=2e-7*ln(Deq/RMG_eq)*1000` | Deq,RMGeq[m] | `2e-7*math.log(Deq/rmg)*1000` | H/km — fórmula clássica de LT trifásica transposta |
|3.3| Frequência angular | `Calc!F10` | `=2π*f` | f=60Hz | `2*math.pi*f` | rad/s |
|3.4| Capacitância (fase-neutro) | `Calc!D10` | `=2π*ε/ln(2*Deq/d)*1000` | ε=permissividade meio, d=diâmetro condutor[m] | ver 3.7 | F/km |
|3.5| Condutância | `Calc!C10=D10*σ/ε` | `G=C*σ/ε` | σ=condutividade do meio | `c*sigma/eps` | nula para ar seco (σ=0) — **efeito corona/perdas dielétricas não convencionais; ver dúvida nº2** |
|3.6| Permissividade do meio | `Calc!E7` | `=εR*1e-9/(36π)` | εR (Principal!G26, padrão 1=ar) | `eps_r*8.8419e-12` | ar seco por padrão |
|3.7| Condutividade do meio | `Calc!F7=Principal!H26` | manual (padrão 0) | σ [S/m] | — | |

## 4. MODELO DE LINHA LONGA (PARÂMETROS DISTRIBUÍDOS)

| # | Cálculo | Origem | Fórmula | Variáveis | Python | Observação |
|---|---|---|---|---|---|---|
|4.1| Impedância série por km | `Calc!A26` | `Z=R+jωL` | R[Ω/km], ωL[Ω/km] | `complex(R, w*L)` | |
|4.2| Admitância shunt por km | `Calc!C26` | `Y=G+jωC` | G,ωC [S/km] | `complex(G, w*C)` | |
|4.3| Impedância característica Zo | `Calc!A28` | `Zo=√(Z/Y)` | | `cmath.sqrt(Z/Y)` | Ω |
|4.4| Constante de propagação γ | `Calc!E28` | `γ=√(Z·Y)` | | `cmath.sqrt(Z*Y)` | 1/km, γ=α+jβ |
|4.5| Constante de atenuação α | `Calc!C30` | `α=Re(γ)` | | `gamma.real` | Np/km |
|4.6| Constante de fase β | `Calc!D30` | `β=Im(γ)` | | `gamma.imag` | rad/km |
|4.7| αl, βl | `Calc!E30,F30` | `α·l`, `β·l` | l=comprimento[km] | | adimensional |
|4.8| cosh(γl), senh(γl) | `Calc!A32:D33` | `cosh(αl)cos(βl)+jsenh(αl)sen(βl)` (decomposição retangular) | | `cmath.cosh(gamma*l)` | valida decomposição trigonométrica-hiperbólica |
|4.9| Impedância de surto Ro (real) | `Calc!A30` | `Ro=√(L/C)` | L[H/km], C[F/km] | `sqrt(L/C)` | Ω — usada no SIL |
|4.10| Velocidade de propagação v | `Calc!H16` | `v=ω/β *1000` | | | m/s |
|4.11| Velocidade da luz no vácuo (ref.) | `Calc!H17` | `v_sp=1/√(L·C)` | | | m/s (linha sem perdas) |
|4.12| Relação v/c | `Principal!V25` | `v/3e8` | | | adimensional |
|4.13| Comprimento de onda l | `Calc!H18/H19` | `v/60/1000` | | | km |

## 5. CIRCUITO π EQUIVALENTE (ABCD → π)

| # | Cálculo | Origem | Fórmula | Python | Observação |
|---|---|---|---|---|---|
|5.1| Zs (impedância série do π, linha longa) | `PI-calc!B6` | `Zs=Zo·senh(γl)` | `Zo*cmath.sinh(gamma*l)` | Ω |
|5.2| Yp (admitância paralela do π, linha longa) | `PI-calc!B13` | `Yp=(1/Zo)·tanh(γl/2)` | `(1/Zo)*cmath.tanh(gamma*l/2)` | S |
|5.3| Zs — modelo de linha curta (referência) | `PI-calc!B20` | `Zs=(R+jωL)·l` | `complex(R,w*L)*l` | comparação/short-line |

## 6. CASO 1 — LINHA EM CARGA (dados no receptor: VCL, SC, cos φ)

| # | Cálculo | Origem | Fórmula | Variáveis | Python | Observação |
|---|---|---|---|---|---|---|
|6.1| Tensão fase-neutro no receptor | `Calc!C17` | `VC=VCL·1000/√3 ∠θ` | VCL[kV] (`Principal!Q5`), θ (`Principal!R5`) | | V |
|6.2| Corrente no receptor | `Calc!C18/A19` | `IC=conj(S/(3·VC))` | SC[kVA] (`Principal!Q7`), cosφ | | A |
|6.3| Tensão no emissor (fasor) | `Calc!A38` | `Vs=VC·cosh(γl)+IC·Zo·senh(γl)` | | `Vc*cosh_gl + Ic*Zo*sinh_gl` | equação geral da linha longa |
|6.4| Corrente no emissor (fasor) | `Calc!A40` | `Is=VC/Zo·senh(γl)+IC·cosh(γl)` | | | A |
|6.5| \|VoL\| emissor (linha) | `Calc!A42` | `\|Vs\|·√3/1000` | | | kV |
|6.6| ∠ tensão emissor | `Calc!B42` | `arg(Vs)` | | | ° |
|6.7| \|Io\| emissor | `Calc!C42=\|Is\|` | | | | A |
|6.8| ∠ corrente emissor | `Calc!D42` | `arg(Is)` | | | ° |
|6.9| Potência complexa emissor So | `Calc!E42` | `So=3·Vs·conj(Is)` | | | VA |
|6.10| \|So\|, cos φ emissor | `Calc!A44,B44` | `\|So\|/1000`, `arg(So)` | | | kVA, ° |
|6.11| Potência ativa emissor Po | `Principal!O21` | `\|So\|·cos(φ)` | | | kW |
|6.12| Potência ativa receptor Pc | `Principal!O24` | `\|SC\|·cos(φC)` | | | kW |
|6.13| **Perdas totais** | `Principal!U4` | `Po − Pc` | | `po - pc` | kW |
|6.14| **Perdas %** | `Principal!U5` / `Relatório!E33` | `(Po−Pc)/Po` ou `Perdas/Po` | | | % |
|6.15| **Regulação de tensão %** | `Principal!U9` | `(\|Vs_L\| − VCL)/VCL` | Vs no vazio vs. VCL plena carga | `(v_vazio - vcl)/vcl` | % — ver observação abaixo |

> **Observação sobre regulação (6.15):** a fórmula usa `U7` (=`Calc!G34`, tensão de emissor calculada a partir do CASO 2 — energização em vazio) comparada a `Q21` (VCL plena carga). Ou seja, a Regulação % da planilha compara **a tensão que apareceria no receptor em vazio, para a mesma tensão de emissor do Caso 1, com a tensão nominal em carga plena** — é a definição clássica de regulação de tensão (`(V_vazio - V_plena_carga)/V_plena_carga`). Ver dúvida nº 3 para confirmação do encadeamento exato Caso 1 → Caso 2 usado nesse cálculo.

## 7. CASO 2 — LINHA EM VAZIO / ENERGIZAÇÃO (efeito Ferranti, a partir da tensão de emissor)

| # | Cálculo | Origem | Fórmula | Variáveis | Observação |
|---|---|---|---|---|---|
|7.1| Zin (impedância de entrada) | `Calc!A21` | `Zin=(Zo·cosh(γl))/senh(γl)` (linha em aberto) | | Ω |
|7.2| Tensão de emissor (entrada) | `Principal!M7,N7` (módulo, ângulo) — **campos não preenchidos no exemplo** | manual | VoL[kV], ∠ | **entrada do Caso 2** |
|7.3| Corrente de carregamento (linha aberta) | `Calc!E14` | `Ic=Vs/Zin` | | A |
|7.4| Tensão no receptor em vazio (Ferranti) | `Calc!A49→A53` | `Vr=Vs/cosh(γl)` (receptor em aberto) | | kV — **sobretensão de Ferranti** |
|7.5| \|VCL\|, \|IC\|, SC (vazio) | `Calc!A53,C53,E53` | | | resultado do Caso 2 |
|7.6| Zc aplicado (impedância característica exibida) | `Principal!U16/U17` | `\|Zc\|`, `\|Zin\|` | | Ω, ∠° |
|7.7| Seleção Caso 1 × Caso 2 | `Principal!L21:R25` | `IF(M7=0 AND Q5=0,"",IF(M7=0,<Caso1>,<Caso2>))` | | a app decide o caso pelos campos preenchidos pelo usuário (M7/Vo especificado → Caso 2; Q5/VCL especificado → Caso 1) |

## 8. SIL, CARREGAMENTO E TENSÃO NATURAL

| # | Cálculo | Origem | Fórmula | Python | Observação |
|---|---|---|---|---|---|
|8.1| SIL (Potência Natural) | `Principal!U13` | `SIL=VCL²/Ro` (Ro=impedância de surto, kV,Ω→kW) | `vcl**2/Ro*1000` | kW — carregamento natural da linha |
|8.2| %SIL (carregamento) | `Principal!V13` | `Pc/SIL` | `pc/sil` | adimensional — carga receptora / SIL |
|8.3| \|V∞\| tensão característica (perfil) | `Principal!U7=Calc!G34` | `\|V∞(l)\|·√3/1000` | | kV |

## 9. PERDAS E DESEMPENHO (RESUMO — QUADRO "Perdas [kW]" DO PAINEL PRINCIPAL)

| # | Cálculo | Origem | Fórmula | Observação |
|---|---|---|---|---|
|9.1| Po (potência ativa emissor) | `Principal!U4` termo 1 | `\|So\|·cos(∠So)` | kW |
|9.2| Pc (potência ativa receptor) | `Principal!U4` termo 2 | `\|SC\|·cos(∠SC)` | kW |
|9.3| Perdas totais | `Principal!U4` | `Po − Pc` | kW |
|9.4| Perdas % | `Principal!U5` | `(Po−Pc)/Po` | % |
|9.5| Regulação % | `Principal!U9` | ver item 6.15 | % |

## 10. EFEITO CORONA / CAMPO ELÉTRICO SUPERFICIAL

| # | Cálculo | Origem | Fórmula | Python | Observação |
|---|---|---|---|---|---|
|10.1| Carga por unidade de comprimento (Q) | `Calc!H13` | `Q=C·Vfase` (C em F/m, V em V) | `C_Fm*V` | C/m |
|10.2| Campo elétrico na superfície do condutor (E) | `Calc!H14` | `E=Q/(2π·ε·r)` | `Q/(2*pi*eps*r)` | V/m, r=raio do condutor |
|10.3| Campo elétrico em kV/cm | `Calc!H15` | `E·1e-5` | | kV/cm — comparável a limite crítico de corona (regra prática ~15–17 kV/cm para ar seco) |

## 11. QUADRO-RESUMO / VALIDAÇÃO (aba "Relatório" dentro da planilha)

A aba `Relatório` já materializa o conjunto de resultados que devem obrigatoriamente aparecer no relatório Word:
Cabo (bitola, seção, ampacidade, d, Ds, R), Estrutura (Dab,Dbc,Dca,Deq), Parâmetros (R,G,L,C por km),
Comprimento, Características (γ, Zo, Ro), Fonte (V,I,S,P no emissor), Carga Plena (V,I,S,P,Z no receptor),
Regulação %, Perdas %, SIL %.

## 12. TABELA DE REFERÊNCIA (aba "Análises") — carregabilidade típica

Tabela estática, **sem vínculo de fórmula com as demais abas** (não referenciada por `Principal`/`Calc`/`Relatório`):

| %SIL | R% (regulação) | P% (perdas) |
|---|---|---|
|1.04|4.51%|4.11%|
|1.17|5.10%|4.60%|
|1.30|5.70%|5.08%|
|1.56|6.92%|6.03%|
|1.82|8.15%|6.97%|

Parece ser um **quadro de referência de carregabilidade típica** (regulação/perdas esperadas em função do %SIL), possivelmente para comparação com o resultado calculado da linha em estudo. **Ver dúvida nº 1.**

---

## PREMISSAS NÃO IDENTIFICADAS — NECESSITAM VALIDAÇÃO

1. **Aba "Análises"**: não está referenciada em nenhuma fórmula das demais abas. Presume-se ser um quadro de referência de carregabilidade típica (comparação). Precisa confirmação de uso (incluir no relatório como referência? ignorar?).
2. **Condutância G** (item 3.5): a planilha calcula G a partir de σ do meio (ambiente), zerada por padrão para ar seco — não há modelagem de perdas por corona/efeito coroa via G. O item 10 (campo elétrico superficial) parece ser o proxy usado para avaliação de corona (comparação com limite crítico), mas a planilha não calcula explicitamente "perdas por corona" nem um limite/veredito. Confirmar se isso deve ser tratado apenas como indicador (sem verificação automática de limite) ou se há um critério de aceitação a aplicar.
3. **Encadeamento Caso 1 → Caso 2 na Regulação (item 6.15)**: a interpretação apresentada (regulação = tensão de vazio vs. tensão em carga plena, ambas referenciadas à mesma tensão de emissor) é a mais consistente com as fórmulas, mas os campos de entrada do Caso 2 (`Principal!M7`, `N7` — tensão de emissor) aparecem vazios no exemplo enviado. Confirmar se, na prática, o usuário informa **tensão de emissor (M7/N7)** manualmente para obter a regulação, ou se o software deve calcular isso automaticamente a partir do próprio resultado do Caso 1 (tensão de emissor já obtida em 6.3–6.6), sem exigir nova entrada do usuário.
4. **Ambiente (εR, σ)**: campos `Principal!G25/H25` (εR, s) — presentes mas com uso secundário (afeta C e G apenas se o meio não for ar). Confirmar se serão expostos como entrada avançada opcional ou fixados em ar (εR=1, σ=0) por padrão.
