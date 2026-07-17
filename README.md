# Validação de Dias e Cálculo de PR — UFV Sol do Agreste

Script para validar os dias de ensaio e calcular o **Performance Ratio (PR)** a partir dos
relatórios do supervisório (EPM/Elipse), conforme o procedimento
**SDA-SIM-E-PVIC-Q01-0135 (Rev. 00)**.

- **Arquivo:** `pr_sol_do_agreste.py`
- **Entrada:** **qualquer planilha** do AnalogReport — CSV (separador `;`/`,`/tab,
  decimal `,` ou `.`, encoding UTF-8/Windows: tudo detectado automaticamente) ou
  Excel (`.xlsx`/`.xls`) — consolidada em **15 min**.
- **Saída:** dias válidos filtrados + PR por dia e do período, exportados em CSV.

---

## O que o script faz

1. **Lê qualquer planilha** (CSV ou Excel): detecta sozinho o separador, o
   encoding, o formato decimal (`5,45` ou `5.45`), a coluna de data/hora (pelo
   nome: `E3TimeStamp`, `Data`, `Timestamp`…) e aceita datas com ou sem hora.
   Linhas sem data válida (cabeçalho repetido, rodapé) são descartadas.
2. **Classifica as colunas automaticamente pelas tags do EPM** (sufixo `_Measurements_XXX`),
   então funciona para qualquer AnalogReport — não só um arquivo específico:
   - `IRRAD_GHI` → irradiância horizontal
   - `IRRAD_POA` → irradiância no plano (POA)
   - `TEMP_MOD` / `TEMP_AR` → temperatura de módulo / ambiente
   - `PosAngAtual` → posição angular dos trackers
   - `EneAtv_Anu` → energia dos inversores (acumulador)
   - PMI/SMF / `PotAtv` → medidor de faturamento / potência ativa (se existirem)
3. **Detecta a UFV** pelo prefixo da tag (`S01`→SDA I … `S06`→SDA VI) e busca a
   temperatura de célula de referência na Tabela 2 do procedimento.
4. **Valida cada dia** pelos critérios do procedimento (Seção 3 e 8.x).
5. **Mantém apenas os dias válidos** e **calcula o PR** (Seção 9).

---

## Requisitos

```bash
pip install pandas numpy
```

## Como rodar

1. Ajuste o bloco `CONFIGURAÇÃO` no topo do script (ver abaixo).
2. Execute:

```bash
python pr_sol_do_agreste.py
```

## Configuração principal

| Parâmetro | Função |
|---|---|
| `CAMINHO_CSV` | caminho do relatório de entrada |
| `TIPO_TESTE` | `"CAP"` / `"ANO 1"` / `"ANO 2"` (define o Coeficiente – Tabela 1) |
| `UFV_FORCADA` | `None` = detectar pela tag; ou fixar `"SDA V"` etc. |
| `POWER_DC_KW` | potência **DC** nominal por UFV (kW) — valores reais SIMM já aplicados |
| `MODO_FORMULA` | `"fisico"` (dimensionalmente consistente) ou `"verbatim"` (literal do doc) |
| `REF_DIURNO` | `"poa"` (robusto) ou `"ghi"` (literal do procedimento) |
| `ENERGIA_FATOR_KWH` | fator p/ converter o acumulador de energia em kWh |

## Critérios de dia válido (implementados)

- **A** — POA média ≥ 600 W/m² por 3 h consecutivas (12 intervalos de 15 min).
- **B** — Irradiação POA diária ≥ 3 kWh/m².
- **C** — ≤ 20% dos intervalos diurnos previstos descartados (dados essenciais ausentes) — item 8.7.
  **GHI é imprescindível**: intervalo diurno sem leitura de GHI conta como descartado
  (`GHI_OBRIGATORIO = True` no script). GHI não entra na fórmula do PR, mas sem GHI
  íntegro o dia não valida — e o PR do dia não é aceito.
- **D** — Restrição ONS ≤ 20% das horas úteis (sem tag no CSV → assume 0; anexar manualmente).

## Fórmula do PR (Seção 9)

Por intervalo `n`, energia teórica:

```
E_teo_n = (POA_n / G_STC) × Power_DC × Δt × f_temp × f_albedo      (modo "fisico")
f_temp   = 1 − B × (Tcel_n − Tcel_ref) / 100        B = 0,29 %/°C
f_albedo = 1 + (Albedo_n[%] − 20) × Coeficiente
Tcel_n   = Tmod_n + (POA_n / G_STC) × 3
PR = Σ Eprod_n / Σ E_teo_n
```

## Saídas geradas

| Arquivo | Conteúdo |
|---|---|
| `validacao_dias.csv` | veredito e métricas de cada dia (válido/inválido + motivos) |
| `dados_dias_validos.csv` | linhas de dados apenas dos dias aprovados |
| `pr_resultados.csv` | Eprod, E_teórica, Tcel_ref e PR por dia |

---

## ⚠️ Pontos de atenção (ler antes do teste oficial)

1. **Energia ≠ PMI.** Sem medidor de faturamento no CSV, o Eprod é a soma do `EneAtv_Anu`
   dos inversores (proxy). O teste de garantia exige o **PMI/SMF**; o script já prioriza
   essa fonte quando a tag existir.
2. **Unidade de energia.** `EneAtv_Anu` vem rotulado "Wh" no catálogo, mas a magnitude
   indica **kWh** (≈217 MWh/dia numa UFV de ~33 MW). Confirmar e ajustar `ENERGIA_FATOR_KWH`.
3. **`POWER_DC_KW`** já usa os valores **DC reais** informados pela SIMM
   (SDA-I 54,174 · SDA-II 28,687 · SDA-III 28,687 · SDA-IV 33,297 ·
   SDA-V 41,622 · SDA-VI 25,227 MWp).
4. **GHI pode falhar no meio-dia.** No arquivo de referência o piranômetro GHI ficou
   **sem leitura em 100% do pico solar (10h–14h)** — com `GHI_OBRIGATORIO = True` esse
   dia é reprovado no critério C (65% de descarte ≫ 20%). Conferir/consertar o
   piranômetro **antes** do teste oficial, senão nenhum dia valida. O "diurno" usa
   **POA** por padrão; alternar `REF_DIURNO = "ghi"` só com sensor íntegro.
5. **Sem tag de albedo** → termo bifacial neutro (albedo = 20%). Informar o padrão da tag
   caso exista em outros relatórios.
6. **`MODO_FORMULA`.** `"fisico"` produz PR coerente (~84,5% no dia de referência);
   `"verbatim"` (0,20 × kWh) sai ~5× maior. Definir com a SIMM qual adotar oficialmente,
   comparando com o PVsyst de referência.

## Referências

Procedimento SDA-SIM-E-PVIC-Q01-0135 · IEC 61724-1/2/3 · NREL/TP-5200-57991 · Estudos PVsyst.
