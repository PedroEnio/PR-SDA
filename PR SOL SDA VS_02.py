# -*- coding: utf-8 -*-
"""
Validação de dias válidos e cálculo do Performance Ratio (PR)
UFV Sol do Agreste — Procedimento SDA-SIM-E-PVIC-Q01-0135 (Rev. 00)

Leitor GENÉRICO para relatórios do supervisório (EPM/Elipse):
  - Lê QUALQUER planilha: CSV (separador, decimal e encoding detectados
    automaticamente) ou Excel (.xlsx/.xls/.xlsm).
  - Classifica as colunas automaticamente pelas TAGS (sufixo _Measurements_XXX),
    então funciona para qualquer AnalogReport, não só um arquivo específico.
  - Detecta a UFV pelo prefixo da tag (S01..S06 -> SDA I..VI).
  - Agrega múltiplos sensores (média de POA, média de TEMP_MOD, média de GHI).
  - Fonte de energia (Eprod): usa medidor PMI/SMF se existir; senão integra
    potência ativa. Dados do inversor (EneAtv_Anu) NÃO são usados.

Fluxo: ler -> classificar tags -> validar cada dia -> manter só dias válidos -> PR.
Dados esperados: consolidados em 15 min (Intervalo de Cálculo do procedimento).
"""

import os
import re
import sys
import pandas as pd
import numpy as np

# Console do Windows usa cp1252 e quebra nos emojis dos prints (✅/❌).
# Força UTF-8 na saída; se o terminal não suportar, troca o char por "?".
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# =============================================================================
# CONFIGURAÇÃO
# =============================================================================
# Caminho do relatório de entrada. Aceita CSV ou Excel — o formato interno
# (separador, decimal, encoding, coluna de tempo) é detectado sozinho.
CAMINHO_CSV = "AnalogReport_20260714 - Sem Correntes (2).csv"
PASTA_SAIDA = "."

TIPO_TESTE  = "CAP"          # "CAP" | "ANO 1" | "ANO 2"  (Coeficiente – Tabela 1)
UFV_FORCADA = None           # None = detectar pelo prefixo Sxx da tag; ou "SDA V"

# Potência nominal AC por UFV (kW). ⚠ AJUSTE com os valores AC reais do projeto.
# (os valores abaixo são a Pot. AC das tabelas do procedimento como PLACEHOLDER)
POWER_AC_KW = {
    "SDA I": 43290.0, "SDA II": 23310.0, "SDA III": 23310.0,
    "SDA IV": 26640.0, "SDA V": 33300.0, "SDA VI": 19980.0,
}

POWER_DC_KW = {
    "SDA I": 43290.0, "SDA II": 23310.0, "SDA III": 23310.0,
    "SDA IV": 26640.0, "SDA V": 33300.0, "SDA VI": 19980.0,
}


INTERVALO_MIN = 15
INTERVALO_H   = INTERVALO_MIN / 60.0     # 0.25 h — usado p/ converter potência em energia

# --- Constantes da fórmula (Seção 9) ---
# Ajustadas conforme doc "5.13_Anexo_V.13 - Testes de Aceitação e garantia de Performance".
G_STC          = 1.0         # kW/m² (1000 W/m²) — irradiância de referência STC
B_COEF         = 0.34        # %/°C — coeficiente térmico da fórmula
FATOR_FORMULA  = 0.25        # constante "0,25" do documento (só no modo "verbatim")
ALBEDO_REF_PCT = 25.0        # albedo de referência (%) do termo bifacial
# "verbatim": fiel ao doc (0,25 × POA_kWh/G_STC) | "fisico": (POA/G_STC)×PowerAC×horas
# Selecionado em tempo de execução (ver __main__); este é apenas o padrão.
MODO_FORMULA   = "fisico"

# --- Energia / unidades ---
# Fator para converter o acumulador de energia p/ kWh.
# EneAtv_Anu vem rotulado "Wh" no catálogo, mas a magnitude indica kWh -> use 1.0.
ENERGIA_FATOR_KWH = 1.0
# O acumulador do medidor de faturamento (SMF/PMI ..._EneAtvSda) está em GWh
# (lê ~4-5 e sobe ~0,2 no dia p/ uma usina de dezenas de MW). GWh -> kWh = 1e6.
PMI_ACUM_FATOR_KWH = 1.0e6
# Referência para "período diurno": "poa" (robusto; POA costuma estar íntegro)
# ou "ghi" (literal do procedimento, mas quebra se o piranômetro GHI falhar).
REF_DIURNO = "poa"

# --- Limiares dos critérios (procedimento) ---
IRRAD_MIN_DIURNO = 50.0    # W/m² — define intervalo diurno "previsto"
POA_JANELA_MIN   = 600.0   # W/m² — POA média mínima na janela do critério A
POA_JANELA_INT   = 12      # 12 × 15 min = 3 h consecutivas
POA_DIARIA_MIN   = 3.0     # kWh/m² — irradiação POA diária mínima (critério B)
MAX_DESCARTE_PCT = 20.0    # % máx. de intervalos descartados (item 8.7, critério C)
MAX_ONS_PCT      = 20.0    # % máx. de restrição ONS (Seção 3, critério D)

# Se True, o gráfico é gerado E o PR é calculado com TODOS os dias, mesmo os
# inválidos/indisponíveis (ignora o filtro de validação). O gráfico já sai
# sempre; este flag afeta só o PR. Resultado de dia inválido é apenas para
# inspeção — não tem valor de aceitação/garantia.
CALCULAR_PR_SEMPRE = False

# =============================================================================
# TABELAS DO DOCUMENTO
# =============================================================================
# Tabela 1 — Coeficiente do termo de albedo por tipo de teste
COEFICIENTE = {"CAP": 0.00103795, "ANO 1": 0.00110354, "ANO 2": 0.00112148}

TCEL_REF = {  # Tabela 2 — Tcel de referência (°C) por mês (1..12) e UFV
    1:{"SDA I":40.10,"SDA II":40.16,"SDA III":40.16,"SDA IV":40.10,"SDA V":40.08,"SDA VI":40.13},
    2:{"SDA I":40.21,"SDA II":40.28,"SDA III":40.28,"SDA IV":40.20,"SDA V":40.18,"SDA VI":40.24},
    3:{"SDA I":38.51,"SDA II":38.58,"SDA III":38.58,"SDA IV":38.51,"SDA V":38.48,"SDA VI":38.55},
    4:{"SDA I":36.35,"SDA II":36.37,"SDA III":36.37,"SDA IV":36.51,"SDA V":36.28,"SDA VI":36.37},
    5:{"SDA I":35.18,"SDA II":35.18,"SDA III":35.17,"SDA IV":35.12,"SDA V":35.33,"SDA VI":35.05},
    6:{"SDA I":32.69,"SDA II":32.73,"SDA III":32.73,"SDA IV":32.68,"SDA V":32.71,"SDA VI":32.70},
    7:{"SDA I":32.33,"SDA II":32.37,"SDA III":32.37,"SDA IV":32.33,"SDA V":32.30,"SDA VI":32.34},
    8:{"SDA I":33.12,"SDA II":33.03,"SDA III":33.03,"SDA IV":32.97,"SDA V":33.13,"SDA VI":32.99},
    9:{"SDA I":36.02,"SDA II":36.08,"SDA III":36.08,"SDA IV":36.01,"SDA V":35.99,"SDA VI":36.04},
    10:{"SDA I":38.06,"SDA II":38.12,"SDA III":38.12,"SDA IV":38.05,"SDA V":38.03,"SDA VI":38.08},
    11:{"SDA I":40.11,"SDA II":40.14,"SDA III":40.15,"SDA IV":40.32,"SDA V":40.03,"SDA VI":40.15},
    12:{"SDA I":39.88,"SDA II":39.94,"SDA III":39.95,"SDA IV":39.91,"SDA V":39.86,"SDA VI":39.91},
}
# Prefixo numérico da tag (S01..S06) -> nome da UFV
UFV_POR_PREFIXO = {1:"SDA I",2:"SDA II",3:"SDA III",4:"SDA IV",5:"SDA V",6:"SDA VI"}
# Número da UFV a partir do nome (inverso de UFV_POR_PREFIXO)
NUM_POR_UFV = {v: k for k, v in UFV_POR_PREFIXO.items()}

# Piranômetro GHI (estação WSR) mais próximo de cada estação — usado como
# substituto quando o GHI primário está ausente (distâncias medidas em campo):
#   WSR-1 -> WSR-5 (707,85 m) | WSR-2 -> WSR-3 (525,48 m) | WSR-3 -> WSR-2 (525,48 m)
#   WSR-4 -> WSR-1 (1071,60 m) | WSR-5 -> WSR-1 (707,85 m) | WSR-6 -> WSR-3 (711,81 m)
NEAREST_WSR = {1: 5, 2: 3, 3: 2, 4: 1, 5: 1, 6: 3}


# =============================================================================
# 1) LEITURA GENÉRICA + CLASSIFICAÇÃO POR TAG
# =============================================================================
# Extensões tratadas como Excel (exigem `pip install openpyxl` para .xlsx)
EXTENSOES_EXCEL = (".xlsx", ".xls", ".xlsm")
# Palavras-chave (maiúsculas) para localizar a coluna de data/hora pelo nome
PALAVRAS_TEMPO = ("E3TIMESTAMP", "TIMESTAMP", "DATETIME", "DATA", "DATE", "HORA", "TIME")


def _abrir_texto(caminho):
    """Abre o arquivo texto testando encodings comuns (BOM/UTF-8/Windows).

    Retorna (conteudo_da_1a_linha, encoding_que_funcionou).
    """
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            with open(caminho, "r", encoding=enc) as f:
                return f.readline(), enc
        except UnicodeDecodeError:
            continue
    # último recurso: latin-1 nunca falha (aceita qualquer byte)
    with open(caminho, "r", encoding="latin-1") as f:
        return f.readline(), "latin-1"


def _detectar_separador(linha_cabecalho):
    """Escolhe o separador (; , tab |) que mais aparece no CABEÇALHO.

    Usar o cabeçalho evita confundir a vírgula decimal dos dados com separador.
    """
    candidatos = [";", ",", "\t", "|"]
    return max(candidatos, key=linha_cabecalho.count)


def _achar_coluna_tempo(df):
    """Localiza a coluna de data/hora pelo nome; se nada casar, usa a 1ª coluna."""
    for c in df.columns:
        cu = str(c).upper()
        if any(p in cu for p in PALAVRAS_TEMPO):
            return c
    return df.columns[0]


def _converter_numerico(serie):
    """Converte uma coluna texto em número aceitando qualquer convenção:
      - decimal brasileiro:  "1.234,56" / "5,45"  -> 1234.56 / 5.45
      - decimal americano:   "1,234.56" / "5.45"  -> 1234.56 / 5.45
    A convenção é decidida pelo ÚLTIMO separador antes dos dígitos finais
    (esse é sempre o separador decimal; o outro é milhar).
    """
    t = serie.astype(str).str.strip()
    amostra = t[t.str.contains(r"\d", na=False)].head(200)
    ult_sep = amostra.str.extract(r"([.,])(?=\d+$)", expand=False)
    n_virg  = int((ult_sep == ",").sum())
    n_ponto = int((ult_sep == ".").sum())
    if n_virg > n_ponto:
        # brasileiro: remove ponto de milhar, vírgula vira ponto decimal
        t = t.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    else:
        # americano (ou inteiro puro): só remove vírgula de milhar
        t = t.str.replace(",", "", regex=False)
    return pd.to_numeric(t, errors="coerce")


def _parse_datas(serie):
    """Converte texto em datetime aceitando formatos mistos na mesma coluna
    (ex.: "14/07/2026" sem hora na 1ª linha e "14/07/2026 00:15:00" no resto).
    dayfirst=True porque os relatórios usam dia/mês/ano.
    """
    s = serie.astype(str).str.strip()
    try:
        # pandas >= 2.0: format="mixed" resolve linha a linha
        return pd.to_datetime(s, dayfirst=True, format="mixed", errors="coerce")
    except (TypeError, ValueError):
        # pandas antigo não conhece format="mixed"
        return pd.to_datetime(s, dayfirst=True, errors="coerce")


def tratar_dados(caminho):
    """Lê QUALQUER planilha (CSV ou Excel) e devolve um DataFrame numérico
    indexado por data/hora, pronto para a validação e o cálculo do PR.

    Etapas:
      1. Excel -> pd.read_excel; CSV -> detecta encoding e separador.
      2. Tudo é lido como texto (dtype=str) para nós controlarmos a conversão.
      3. Acha a coluna de tempo pelo nome e converte (formatos mistos OK).
      4. Converte cada coluna de dados p/ número, seja decimal "," ou ".".
      5. Descarta linhas sem data válida (remove lixo/preâmbulo automaticamente).
    """
    ext = os.path.splitext(caminho)[1].lower()
    if ext in EXTENSOES_EXCEL:
        df = pd.read_excel(caminho, dtype=str)
    else:
        cabecalho, enc = _abrir_texto(caminho)
        sep = _detectar_separador(cabecalho)
        print(f"CSV detectado: separador '{sep}' | encoding {enc}")
        df = pd.read_csv(caminho, sep=sep, dtype=str, encoding=enc,
                         skip_blank_lines=True)

    df.columns = [str(c).strip() for c in df.columns]   # tira espaços dos nomes
    col_tempo = _achar_coluna_tempo(df)
    df[col_tempo] = _parse_datas(df[col_tempo])

    for col in df.columns:
        if col == col_tempo:
            continue
        df[col] = _converter_numerico(df[col])

    # linhas sem data válida = cabeçalhos repetidos, rodapés, lixo -> fora
    return (df.dropna(subset=[col_tempo])
              .set_index(col_tempo).sort_index())


def classificar_tags(df):
    """Agrupa colunas por grandeza a partir do sufixo _Measurements_XXX da tag EPM.

    Devolve (grupos, prefixos):
      grupos   — dict grandeza -> lista de colunas daquela grandeza
      prefixos — números Sxx encontrados nas tags (p/ detectar a UFV)
    """
    g = {"ghi":[], "poa":[], "tmod":[], "tamb":[], "tracker":[],
         "energia_pmi":[], "energia_inv":[], "potencia":[]}
    prefixos = set()
    for c in df.columns:
        cu = c.upper()
        mS = re.match(r"S(\d+)_", c)          # prefixo S05_... -> UFV 5
        if mS:
            prefixos.add(int(mS.group(1)))
        if "IRRAD_GHI" in cu:                         g["ghi"].append(c)
        elif "IRRAD_POA" in cu:                       g["poa"].append(c)
        elif "TEMP_MOD" in cu:                        g["tmod"].append(c)
        elif "TEMP_AR" in cu:                         g["tamb"].append(c)
        elif "POSANG" in cu:                          g["tracker"].append(c)
        elif ("PMI" in cu or "SMF" in cu or "MEDFAT" in cu) and ("ENE" in cu or "POT" in cu):
            g["energia_pmi"].append(c)                # medidor de faturamento
        elif "ENEATV" in cu or "ENE_ATV" in cu:       g["energia_inv"].append(c)
        elif "POTATV" in cu or "POT_ATV" in cu:       g["potencia"].append(c)
    return g, prefixos


def serie_media(df, cols):
    """Média linha a linha das colunas do grupo (vários sensores -> 1 série).
    Grupo vazio -> série de NaN (mantém o restante do fluxo funcionando)."""
    return df[cols].mean(axis=1) if cols else pd.Series(np.nan, index=df.index)


def _wsr_por_estacao(df, cols_ghi):
    """Agrupa as colunas GHI pelo número da estação WSR (tag ..._WSR_<n>_...).
    Retorna dict {n_estacao: série média das colunas daquele WSR}."""
    por_wsr = {}
    for c in cols_ghi:
        m = re.search(r"WSR_?(\d+)", c.upper())
        n = int(m.group(1)) if m else 0        # 0 = WSR não identificado na tag
        por_wsr.setdefault(n, []).append(c)
    return {n: df[cs].mean(axis=1) for n, cs in por_wsr.items()}


def serie_ghi(df, g, ufv):
    """GHI para a VALIDAÇÃO do dia (procedimento: só GHI decide a validade).

    Regra: usa o piranômetro GHI da própria UFV (estação WSR de mesmo número);
    onde ele estiver ausente (NaN), preenche com o piranômetro mais próximo
    (mapa NEAREST_WSR), encadeando estações até não sobrar buraco ou acabarem
    as estações. Sem nenhum GHI -> série de NaN (dia será INDISPONÍVEL).
    """
    series = _wsr_por_estacao(df, g["ghi"])
    if not series:
        return pd.Series(np.nan, index=df.index)
    # estação primária = a de número da UFV; se não existir, a de menor número
    num_ufv = NUM_POR_UFV.get(ufv)
    primaria = num_ufv if num_ufv in series else min(series)
    ghi = series[primaria].copy()

    # preenche faltas com o WSR mais próximo, seguindo o mapa sem repetir estação
    atual, visitados = primaria, {primaria}
    while ghi.isna().any():
        prox = NEAREST_WSR.get(atual)
        if prox is None or prox in visitados:
            break
        if prox in series:
            ghi = ghi.fillna(series[prox])
        visitados.add(prox)
        atual = prox
    return ghi


def _medidores_pmi_da_ufv(g, ufv):
    """Escolhe as tags de medidor de faturamento (SMF/PMI) da UFV pedida.

    As tags têm o formato SMF_S05_C1_P_UMF5 / SMF_S05_C2_P_UMF7:
      - S05  = número da UFV (aqui SDA V);
      - C1/C2 = circuito (alimentador) — a usina tem vários;
      - _P / _R = registrador Principal / Retaguarda — o _R ESPELHA o _P do
        mesmo circuito; somar os dois DOBRA a energia. Por isso usamos só o _P.
    Energia total da usina = soma de UM registrador (_P) por circuito.
    Se nenhuma tag casar com a UFV, cai p/ todas as PMI (comportamento antigo).
    """
    num = NUM_POR_UFV.get(ufv)
    # tags da UFV: prefixo _S0<num>_ em qualquer posição da tag (SMF_S05_...)
    da_ufv = [c for c in g["energia_pmi"]
              if num is not None and re.search(rf"_S0*{num}_", c.upper())]
    if not da_ufv:
        return g["energia_pmi"]        # sem match -> não filtra (fallback)
    # 1 registrador por circuito: prioriza o Principal (_P_); se um circuito só
    # tiver _R, usa o _R para não perder o circuito.
    por_circuito = {}
    for c in da_ufv:
        mC = re.search(r"_C(\d+)_", c.upper())
        circ = mC.group(1) if mC else c
        principal = re.search(r"_C\d+_P_", c.upper()) is not None
        # guarda o _P; só aceita _R se ainda não houver nada para o circuito
        if circ not in por_circuito or principal:
            por_circuito[circ] = c
    return list(por_circuito.values())


def energia_intervalo(df, g, ufv=None):
    """
    Retorna a energia medida por intervalo (kWh) e a descrição da fonte usada.
    Prioridade: PMI/SMF -> potência ativa.
    Quando `ufv` é informada, usa SÓ os medidores daquela UFV (1 registrador _P
    por circuito) — evita ler o medidor de outra usina ou só metade dos circuitos.
    NOTA: os dados dos inversores (EneAtv_Anu) NÃO são usados como fonte de energia
    (decisão do procedimento — não levar em consideração dados do inversor).
    """
    if g["energia_pmi"]:
        cols = _medidores_pmi_da_ufv(g, ufv) if ufv else [g["energia_pmi"][0]]
        s = df[cols].astype(float)
        # se for acumulador de energia -> diff; heurística: valores sempre crescentes
        e = s.diff()
        # detecta potência vs acumulador pela 1ª coluna (mesma natureza p/ todas)
        if (e[cols[0]].dropna() < 0).mean() > 0.3:   # muitos deltas negativos = é potência
            e = s.sum(axis=1) * INTERVALO_H
            fonte = f"PMI (potência): {', '.join(cols)}"
        else:
            # soma o delta de cada circuito e converte o acumulador (GWh) p/ kWh
            e = e.clip(lower=0).sum(axis=1) * PMI_ACUM_FATOR_KWH
            fonte = f"PMI (acumulador, {len(cols)} circuito(s), GWh->kWh): {', '.join(cols)}"
        return e, fonte
    if g["potencia"]:
        # última opção: integra potência ativa (kW × 0,25 h = kWh)
        e = df[g["potencia"]].sum(axis=1) * INTERVALO_H
        return e, f"Potência ativa (soma de {len(g['potencia'])} tags) × {INTERVALO_H} h"
    return pd.Series(np.nan, index=df.index), "SEM fonte de energia encontrada"


def detectar_ufv(prefixos):
    """UFV_FORCADA vence; senão usa o menor prefixo Sxx visto nas tags."""
    if UFV_FORCADA:
        return UFV_FORCADA
    for p in sorted(prefixos):
        if p in UFV_POR_PREFIXO:
            return UFV_POR_PREFIXO[p]
    return "SDA I"   # fallback conservador se nenhuma tag tiver prefixo


def selecionar_arquivo(padrao=CAMINHO_CSV):
    """Abre um diálogo do Windows p/ escolher a pasta e o arquivo (CSV/Excel).
    Se não houver ambiente gráfico ou o usuário cancelar, usa `padrao`."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk(); root.withdraw()
        root.attributes("-topmost", True)            # traz o diálogo p/ frente
        inicial = os.path.dirname(os.path.abspath(padrao)) or "."
        cam = filedialog.askopenfilename(
            title="Selecione o relatório do supervisório (CSV ou Excel)",
            initialdir=inicial,
            filetypes=[("Planilhas", "*.csv *.xlsx *.xls *.xlsm"),
                       ("CSV", "*.csv"), ("Excel", "*.xlsx *.xls *.xlsm"),
                       ("Todos os arquivos", "*.*")])
        root.destroy()
        if cam:
            return cam
        print("Nenhum arquivo selecionado — usando o padrão.")
    except Exception as e:
        print(f"Diálogo de arquivo indisponível ({e}) — usando o padrão.")
    return padrao


def _abrir_arquivo(caminho):
    """Abre um arquivo (ex.: o PNG do gráfico) no visualizador padrão do SO."""
    try:
        if not os.path.exists(caminho):
            return
        caminho = os.path.abspath(caminho)              # os.startfile exige caminho real
        if sys.platform.startswith("win"):
            os.startfile(caminho)                       # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            import subprocess; subprocess.Popen(["open", caminho])
        else:
            import subprocess; subprocess.Popen(["xdg-open", caminho])
    except Exception as e:
        print(f"Não consegui abrir {caminho} automaticamente ({e}).")


def _rotulo_sensor(tag):
    """Rótulo curto p/ a legenda: mesa + estação (ex.: 'TS1 · WSS_5_2').
    Amarra o piranômetro à mesa (TS) que agrupa os trackers."""
    tu = tag.upper()
    mts = re.search(r"S\d+_(TS\d+)_", tu)
    mws = re.search(r"(WS[SR]_[\dA-Z_]*\d)", tu)
    ts = mts.group(1) if mts else ""
    est = mws.group(1) if mws else tag
    return f"{ts} · {est}" if ts else est


def _mapa_piranometros(g):
    """Retorna dict {TS: {'POA':[estações], 'GHI':[estações]}} p/ mostrar quais
    mesas (grupos de trackers) têm piranômetro."""
    mapa = {}
    for grandeza, cols in (("POA", g["poa"]), ("GHI", g["ghi"])):
        for c in cols:
            mts = re.search(r"S\d+_(TS\d+)_", c.upper())
            ts = mts.group(1) if mts else "?"
            mapa.setdefault(ts, {"POA": [], "GHI": []})[grandeza].append(_rotulo_sensor(c))
    return mapa


def _garantir_matplotlib():
    """Importa matplotlib; se faltar, instala no MESMO interpretador e reimporta.
    Evita o erro 'matplotlib não instalado' quando o script roda num Python
    diferente do que tem a lib. Retorna o módulo ou None se não der."""
    try:
        import matplotlib
        return matplotlib
    except ImportError:
        print("matplotlib não encontrado — instalando no Python atual (pip)...")
        try:
            import subprocess
            subprocess.check_call([sys.executable, "-m", "pip", "install", "matplotlib"])
            import matplotlib
            return matplotlib
        except Exception as e:
            print(f"Falha ao instalar matplotlib ({e}).")
            print(f'Instale manualmente: "{sys.executable}" -m pip install matplotlib')
            return None


def plotar_ghi_poa(df, g, ufv, salvar=True, mostrar=True):
    """Plota a irradiância GHI e POA de cada estação/piranômetro ao longo do tempo.

    Uma linha por sensor: POA (linha cheia) e GHI (tracejada). Salva PNG e,
    se houver ambiente gráfico, mostra a janela.
    """
    matplotlib = _garantir_matplotlib()
    if matplotlib is None:
        print("Gráfico ignorado (matplotlib indisponível).")
        return
    if mostrar:
        # força um backend interativo (janela). Sem isso, alguns ambientes
        # caem no 'Agg' e plt.show() volta na hora, sem abrir nada.
        for backend in ("TkAgg", "QtAgg", "Qt5Agg"):
            try:
                matplotlib.use(backend, force=True)
                break
            except Exception:
                continue
    else:
        matplotlib.use("Agg")               # backend sem tela p/ apenas salvar
    import matplotlib.pyplot as plt
    if mostrar:
        print(f"Backend do gráfico: {matplotlib.get_backend()} "
              "(feche a janela do gráfico para o script continuar)")
    if not g["poa"] and not g["ghi"]:
        print("Sem colunas de POA/GHI no arquivo — gráfico não gerado.")
        return

    # mostra quais mesas (grupos de trackers) têm piranômetro
    mapa = _mapa_piranometros(g)
    print("Piranômetros por mesa (TS):")
    for ts in sorted(mapa):
        info = mapa[ts]
        partes = []
        if info["POA"]: partes.append("POA=" + ",".join(info["POA"]))
        if info["GHI"]: partes.append("GHI=" + ",".join(info["GHI"]))
        print(f"  {ts}: " + " | ".join(partes))

    fig, ax = plt.subplots(figsize=(13, 6))
    for c in g["poa"]:
        ax.plot(df.index, df[c], lw=1.1, label=f"POA · {_rotulo_sensor(c)}")
    for c in g["ghi"]:
        ax.plot(df.index, df[c], lw=1.6, ls="--", label=f"GHI · {_rotulo_sensor(c)}")

    ax.set_title(f"Irradiância — GHI e POA por piranômetro — {ufv}")
    ax.set_xlabel("Data / hora")
    ax.set_ylabel("Irradiância (W/m²)")
    ax.grid(True, alpha=0.3)
    ax.legend(fontsize=8, ncol=2)
    fig.autofmt_xdate()
    fig.tight_layout()

    if salvar:
        out = f"{PASTA_SAIDA}/grafico_ghi_poa.png"
        try:
            fig.savefig(out, dpi=120)
            print(f"Gráfico salvo: {out}")
        except Exception as e:
            print(f"Falha ao salvar o gráfico ({e}).")
    if mostrar:
        # traz a janela do gráfico p/ frente (no Windows o TkAgg costuma abrir
        # atrás do terminal e parece que "não saiu o plot").
        try:
            win = plt.get_current_fig_manager().window
            win.attributes("-topmost", True)
            win.after(700, lambda: win.attributes("-topmost", False))
        except Exception:
            pass
        try:
            plt.show(block=True)
        except Exception as e:
            print(f"Não abriu a janela do gráfico ({e}). Abrindo o PNG salvo...")
    plt.close(fig)
    # garante que a imagem apareça mesmo se a janela interativa não vier ao foco
    if salvar:
        _abrir_arquivo(f"{PASTA_SAIDA}/grafico_ghi_poa.png")


# =============================================================================
# 2) VALIDAÇÃO DO DIA
# =============================================================================
def validar_dia(dfd, g, ufv):
    """Aplica os critérios A–D do procedimento a UM dia de dados.
    Retorna (dia_valido, motivos) — motivos vai virar linha do validacao_dias.csv.
    """
    poa = serie_media(dfd, g["poa"])
    # GHI da UFV (estação WSR própria); onde faltar, usa o piranômetro mais
    # próximo pelo mapa NEAREST_WSR. Sem nenhum GHI -> dia INDISPONÍVEL.
    ghi = serie_ghi(dfd, g, ufv)
    ref = poa if REF_DIURNO == "poa" else ghi
    diurno = ref > IRRAD_MIN_DIURNO          # intervalos diurnos "previstos"
    n_prev = int(diurno.sum())
    mot = {"intervalos_previstos": n_prev, "ref_diurno": REF_DIURNO}
    if n_prev == 0:
        return False, {**mot, "erro": "sem intervalos diurnos"}

    # GATE) Disponibilidade de GHI — a validação do dia exige GHI (topo do procedimento):
    # "o valor para verificação da validação do dia deve ser apenas o GHI; sem GHI,
    #  considerar indisponível". Sem nenhum GHI no período diurno -> dia INDISPONÍVEL.
    ghi_diurno = ghi.where(diurno)
    n_ghi = int(ghi_diurno.notna().sum())
    mot["GHI_cobertura_pct"] = round(100.0 * n_ghi / n_prev, 1)
    mot["GHI_disponivel"] = bool(g["ghi"]) and n_ghi > 0
    if not mot["GHI_disponivel"]:
        mot["INDISPONIVEL"] = True
        mot["DIA_VALIDO"] = False
        return False, {**mot, "erro": "INDISPONÍVEL: sem GHI"}

    # A) POA média >= 600 W/m² por 3 h consecutivas (janela móvel de 12 × 15 min)
    poa_ok = bool((poa.rolling(POA_JANELA_INT).mean() >= POA_JANELA_MIN).any())
    mot["A_poa600_3h"] = poa_ok
    # B) irradiação POA diária >= 3 kWh/m² (integra só o período diurno)
    irr = float(poa.where(diurno, 0).sum() * INTERVALO_H / 1000.0)
    mot["B_poa_diaria_kwh"] = round(irr, 3); mot["B_ok"] = irr >= POA_DIARIA_MIN
    # C) <=20% dos intervalos previstos descartados (grandezas essenciais ausentes)
    tmod = serie_media(dfd, g["tmod"])
    ener, _ = energia_intervalo(dfd, g, ufv)
    falha = poa.isna() | tmod.isna() | ener.isna()   # falta POA, Tmod ou energia
    pct_desc = 100.0 * int((diurno & falha).sum()) / n_prev
    mot["C_pct_descartado"] = round(pct_desc, 1); mot["C_ok"] = pct_desc <= MAX_DESCARTE_PCT
    # D) restrição ONS (sem tag no CSV -> assume 0%; anexar manualmente se houver)
    mot["D_pct_ons"] = 0.0; mot["D_ok"] = True

    ok = bool(poa_ok and mot["B_ok"] and mot["C_ok"] and mot["D_ok"])
    mot["DIA_VALIDO"] = ok
    return ok, mot


def filtrar_dias_validos(df, g, ufv):
    """Roda validar_dia() para cada dia do período.
    Retorna (df só com linhas dos dias aprovados, resumo com veredito por dia)."""
    resumo, idx_val = [], []
    for dia, dfd in df.groupby(df.index.normalize()):   # normalize() agrupa por data
        ok, mot = validar_dia(dfd, g, ufv)
        resumo.append({"dia": dia.date(), **mot})
        if mot.get("INDISPONIVEL"):
            status = "⚠️  INDISPONÍVEL (sem GHI)"
        elif ok:
            status = "✅ VÁLIDO"
        else:
            status = "❌ inválido"
        print(f"{dia.date()} -> {status} "
              f"(GHI_cob={mot.get('GHI_cobertura_pct')}%, POA/3h={mot.get('A_poa600_3h')}, "
              f"POA_dia={mot.get('B_poa_diaria_kwh')} kWh/m², descarte={mot.get('C_pct_descartado')}%)")
        if ok:
            idx_val.append(dfd.index)
    df_val = df.loc[np.concatenate([i.values for i in idx_val])] if idx_val else df.iloc[0:0]
    return df_val, pd.DataFrame(resumo)


# =============================================================================
# 3) CÁLCULO DO PR (Seção 9)
# =============================================================================
def calcular_pr(df_val, g, ufv):
    """Calcula Eprod, E_teórica e PR por dia válido + PR do período.

    Por intervalo n:
      Tcel_n  = Tmod_n + (POA_n/G_STC) × 3
      f_temp  = 1 − B × (Tcel_n − Tcel_ref)/100
      f_alb   = 1 + (Albedo_n − 25) × Coeficiente     (sem sensor -> neutro)
      E_teo_n = (POA_n/G_STC) × Power_DC × Δt × f_temp × f_alb    (modo "fisico")
      PR      = Σ Eprod_n / Σ E_teo_n
    """
    if df_val.empty:
        print("\n(nenhum dia válido — PR não calculado)")
        return pd.DataFrame(), None
    coef = COEFICIENTE[TIPO_TESTE.upper()]
    pdc = POWER_DC_KW[ufv]
    ener_full, fonte = energia_intervalo(df_val, g, ufv)
    print(f"\nFonte de energia (Eprod): {fonte}")
    print(f"UFV: {ufv} | Power_DC: {pdc} kW | Coef: {coef} | Fórmula: {MODO_FORMULA}")

    linhas = []
    for dia, dfd in df_val.groupby(df_val.index.normalize()):
        poa = serie_media(dfd, g["poa"])
        tmod = serie_media(dfd, g["tmod"])
        ref = poa if REF_DIURNO == "poa" else serie_media(dfd, g["ghi"])
        # só entram intervalos diurnos com POA e Tmod presentes
        mask = (ref > IRRAD_MIN_DIURNO) & poa.notna() & tmod.notna()
        d_idx = dfd.index[mask]
        if len(d_idx) == 0:
            continue
        mes = int(dia.month); tcel_ref = TCEL_REF[mes][ufv]   # Tabela 2 (mês × UFV)

        poa_w = poa[mask].astype(float)                       # W/m²
        poa_kwh = poa_w * INTERVALO_H / 1000.0                # kWh/m² no intervalo
        tcel_n = tmod[mask].astype(float) + (poa_w / (G_STC*1000.0)) * 3.0
        albedo = pd.Series(ALBEDO_REF_PCT, index=d_idx)       # sem sensor -> 25% (neutro)
        f_temp = 1.0 - (B_COEF * (tcel_n - tcel_ref)) / 100.0
        f_alb  = 1.0 + (albedo - ALBEDO_REF_PCT) * coef

        if MODO_FORMULA == "verbatim":
            # literal do documento: 0,20 × (POA em kWh/G_STC) × PowerDC × fatores
            e_teo = FATOR_FORMULA * (poa_kwh / G_STC) * pdc * f_temp * f_alb
        else:
            # "fisico": dimensionalmente consistente -> kW × h = kWh
            e_teo = (poa_w / (G_STC*1000.0)) * pdc * INTERVALO_H * f_temp * f_alb

        eprod = float(ener_full.reindex(d_idx).fillna(0).sum())
        e_teo_d = float(e_teo.sum())
        pr = eprod / e_teo_d if e_teo_d else np.nan
        linhas.append({"dia": dia.date(), "intervalos": int(mask.sum()),
                       "Eprod_kWh": round(eprod,2), "E_teorica_kWh": round(e_teo_d,2),
                       "Tcel_ref_C": tcel_ref, "PR_dia_%": round(pr*100,2)})

    df_pr = pd.DataFrame(linhas)
    pr_per = None
    if not df_pr.empty:
        if df_pr["Eprod_kWh"].sum() == 0:
            print("⚠️  Eprod total = 0: sem fonte de energia (PMI/POT ausentes e inversor "
                  "excluído). O PR sai 0% — não reflete a usina. Habilite uma fonte de energia.")
        # PR do período = razão das SOMAS (não é média dos PRs diários)
        pr_per = df_pr["Eprod_kWh"].sum()/df_pr["E_teorica_kWh"].sum()
        print("\n=== RESULTADO PR ===")
        print(df_pr.to_string(index=False))
        print(f"\nDias no cálculo: {len(df_pr)} | PR do período: {pr_per*100:.2f} %")
    return df_pr, pr_per


# =============================================================================
# EXECUÇÃO: ler -> classificar -> validar -> PR -> exportar CSVs
# =============================================================================
def selecionar_modo_formula(padrao=MODO_FORMULA):
    """Solicita ao usuário o tipo de análise (verbatim/fisico).

    'verbatim' = fiel ao documento (0,25 × POA_kWh/G_STC).
    'fisico'   = dimensionalmente consistente ((POA/G_STC) × PowerDC × horas).
    ENTER vazio ou entrada não-interativa mantém o padrão.
    """
    print("\nSelecione o tipo de análise (fórmula da E_teórica):")
    print("  [1] verbatim — fiel ao doc: 0,25 × (POA_kWh/G_STC) × PowerDC × fatores")
    print("  [2] fisico   — (POA/G_STC) × PowerDC × horas × fatores  [RECOMENDADO]")
    try:
        resp = input(f"Opção [1/2] (ENTER = {padrao}): ").strip().lower()
    except EOFError:
        resp = ""
    if resp in ("1", "verbatim"):
        return "verbatim"
    if resp in ("2", "fisico"):
        return "fisico"
    return padrao


def selecionar_pr_sempre(padrao=CALCULAR_PR_SEMPRE):
    """Pergunta se o PR deve ser calculado mesmo em dias INVÁLIDOS/indisponíveis.
    O gráfico é gerado de qualquer forma. ENTER mantém o padrão."""
    dft = "s" if padrao else "n"
    print("\nCalcular o PR mesmo em dias INVÁLIDOS? (o gráfico sai de qualquer forma)")
    print("  [s] sim — usa todos os dias (resultado só p/ inspeção, sem valor contratual)")
    print("  [n] não — apenas dias válidos (procedimento)  [PADRÃO]")
    try:
        resp = input(f"Opção [s/n] (ENTER = {dft}): ").strip().lower()
    except EOFError:
        resp = ""
    if resp in ("s", "sim", "y"):
        return True
    if resp in ("n", "nao", "não"):
        return False
    return padrao


if __name__ == "__main__":
    MODO_FORMULA = selecionar_modo_formula()
    print(f"Modo da fórmula selecionado: {MODO_FORMULA}")
    pr_sempre = selecionar_pr_sempre()
    CAMINHO = selecionar_arquivo(CAMINHO_CSV)     # abre pasta/seleciona o arquivo
    print(f"Lendo: {CAMINHO}")
    df = tratar_dados(CAMINHO)
    g, pref = classificar_tags(df)
    ufv = detectar_ufv(pref)
    print(f"UFV detectada: {ufv} (prefixos {sorted(pref)})")
    print("Tags -> GHI:%d POA:%d TMOD:%d TRK:%d ENE_PMI:%d ENE_INV:%d POT:%d" % (
        len(g["ghi"]),len(g["poa"]),len(g["tmod"]),len(g["tracker"]),
        len(g["energia_pmi"]),len(g["energia_inv"]),len(g["potencia"])))

    print("\n--- Gráfico GHI x POA ---")
    plotar_ghi_poa(df, g, ufv)

    print("\n--- Validação dos dias ---")
    df_val, resumo = filtrar_dias_validos(df, g, ufv)

    if pr_sempre:
        base_pr = df                                  # todos os dias, inclui inválidos
        print("\n--- Cálculo do PR (TODOS os dias — inclui INVÁLIDOS/indisponíveis) ---")
        if df_val.empty:
            print("⚠️  Nenhum dia passou na validação; calculando assim mesmo (só p/ inspeção).")
    else:
        base_pr = df_val                              # comportamento do procedimento
        print("\n--- Cálculo do PR (somente dias válidos) ---")
    df_pr, pr_per = calcular_pr(base_pr, g, ufv)

    # saídas no padrão brasileiro (sep=";" decimal=",") p/ abrir direto no Excel
    resumo.to_csv(f"{PASTA_SAIDA}/validacao_dias.csv", index=False, sep=";", decimal=",")
    if not df_val.empty:
        df_val.to_csv(f"{PASTA_SAIDA}/dados_dias_validos.csv", sep=";", decimal=",")
    if not df_pr.empty:
        df_pr.to_csv(f"{PASTA_SAIDA}/pr_resultados.csv", index=False, sep=";", decimal=",")
    print("\nSalvos: validacao_dias.csv, dados_dias_validos.csv, pr_resultados.csv")
