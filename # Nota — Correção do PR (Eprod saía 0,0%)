# Nota — Correção do PR (Eprod saía 0,0%)

**Arquivo:** `pr_sol_do_agreste (1).py`
**Data:** 2026-07-20
**Sintoma:** PR do dia 2026-07-14 (SDA V) saía **0,0%** mesmo com a planilha lida
(`Eprod_kWh = 0,11` vs `E_teorica_kWh = 235.980`).

## Causa

Não era falha de leitura. A **energia (Eprod)** vinha errada por dois motivos somados:

1. **Medidor errado / circuito incompleto.** `energia_intervalo` pegava
   `g["energia_pmi"][0]` = a **primeira** tag PMI do arquivo
   (`SMF_S01_C1_P...` = medidor da **SDA I**, outra usina, e só 1 circuito),
   enquanto POA/Tmod vinham da SDA V.
2. **Unidade errada.** O acumulador `SMF_..._EneAtvSda` está em **GWh**, não kWh.
   Lê ~4,8 e sobe ~0,2 no dia. Com `fator 1.0`, o delta (~0,11 GWh) virava
   "0,11 kWh". Real: 0,11 GWh = 110.000 kWh.

## Tags SMF (entender antes de mexer)

Formato: `SMF_S05_C1_P_UMF5_Measurements_EneAtvSda`

| Parte  | Significado                                             |
|--------|--------------------------------------------------------|
| `S05`  | Número da UFV (SDA V)                                   |
| `C1/C2`| Circuito / alimentador (a usina tem vários)            |
| `_P`   | Registrador **Principal**                              |
| `_R`   | Registrador **Retaguarda** — ESPELHA o `_P` do circuito |

> ⚠️ Somar `_P` + `_R` **dobra** a energia. Energia total da usina =
> soma de **UM** registrador `_P` por circuito (ex.: `S05_C1_P` + `S05_C2_P`).

## Correção aplicada

- Constante `PMI_ACUM_FATOR_KWH = 1.0e6` (GWh → kWh).
- Função `_medidores_pmi_da_ufv(g, ufv)` — filtra os medidores da UFV detectada,
  1 registrador `_P` por circuito.
- `energia_intervalo(df, g, ufv=None)` — soma os circuitos e escala GWh→kWh.
  Sem `ufv` mantém o comportamento antigo (fallback).
- Chamadas em `validar_dia` e `calcular_pr` passam `ufv`.

## Resultado (SDA V, 2026-07-14)

| Antes         | Depois              |
|---------------|---------------------|
| Eprod = 0,11  | Eprod = 216.238 kWh |
| PR = **0,0%** | PR = **91,63%**     |

E_teorica = 235.980 kWh (coerente).

## Atenção

- **Confirmar a unidade do acumulador (GWh).** Se algum arquivo vier em MWh,
  trocar `PMI_ACUM_FATOR_KWH` para `1.0e3`.
- `CAMINHO_CSV` padrão ainda aponta pro `(2).csv`; o teste usou
  `AnalogReport_20260714 - Sem Correntes_3.csv` (selecionado no diálogo).
- Regex do prefixo Sxx nas tags SMF = `_S0*<num>_` (busca no meio da string);
  o `re.match(r"S(\d+)_")` da irradiância NÃO casa `SMF_S05...`.
- Dados dos inversores (`EneAtv_Anu`) NÃO entram como fonte de energia
  (decisão do procedimento).
