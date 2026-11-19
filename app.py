# ============================================================
# VHPT 2.0 - Verificador de Hipóteses de Carregamento Mecânico
# ============================================================

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.patches import Ellipse
from io import StringIO, BytesIO
import os

st.set_page_config(
    page_title="VHPT 2.0 - Verificador de Postes",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.title("🏗️ VHPT 2.0 - Verificador de Hipóteses de Carregamento Mecânico")
st.markdown("---")

# ============================================================
# CONSTANTES
# ============================================================
COMERCIAIS = [1500, 2000, 2500, 3000, 4000]
SOBRECARGA_TRANSITORIA = 0.40
ALTURA_REFERENCIA = 0.20

ALTURAS_DISPONIVEIS = [18, 20, 22, 24, 26, 28, 30, 32]
ALTURA_DEFAULT = 22

MAPEAMENTO_SET_PHASE = {
    (1, 1): 'dist_f1', (1, 2): 'dist_f2', (1, 3): 'dist_f3',
    (2, 1): 'dist_f4', (2, 2): 'dist_f5', (2, 3): 'dist_f6',
    (4, 1): 'dist_pr',
    (5, 1): 'dist_pr2',
}

SETS_REQUERIDOS = {
    'CS_sem_PR': {1},
    'CS_com_PR': {1, 4},
    'CD_sem_PR': {1, 2},
    'CD_com_PR': {1, 2, 4},
}

CONFIG_OPCOES = ["CS_sem_PR", "CS_com_PR", "CD_sem_PR", "CD_com_PR"]
CONFIG_LABELS = {
    "CS_sem_PR": "CS sem PR", "CS_com_PR": "CS com PR",
    "CD_sem_PR": "CD sem PR", "CD_com_PR": "CD com PR",
}

LCS_DISPONIVEIS = [
    {'nome': 'EDS',              'sets': set()},
    {'nome': 'TEMP MÍM',         'sets': set()},
    {'nome': 'VENT MAX',         'sets': set()},
    {'nome': 'QUEBRA C1 FA B',   'sets': {1}},
    {'nome': 'QUEBRA C1 FA A',   'sets': {1}},
    {'nome': 'QUEBRA C1 FB B',   'sets': {1}},
    {'nome': 'QUEBRA C1 FB A',   'sets': {1}},
    {'nome': 'QUEBRA C1 FC B',   'sets': {1}},
    {'nome': 'QUEBRA C1 FC A',   'sets': {1}},
    {'nome': 'QUEBRA C2 FA B',   'sets': {2}},
    {'nome': 'QUEBRA C2 FA A',   'sets': {2}},
    {'nome': 'QUEBRA C2 FB B',   'sets': {2}},
    {'nome': 'QUEBRA C2 FB A',   'sets': {2}},
    {'nome': 'QUEBRA C2 FC B',   'sets': {2}},
    {'nome': 'QUEBRA C2 FC A',   'sets': {2}},
    {'nome': 'QUEBRA PR1 B',     'sets': {4}},
    {'nome': 'QUEBRA PR1 A',     'sets': {4}},
    {'nome': 'QUEBRA PR2 B',     'sets': {5}},
    {'nome': 'QUEBRA PR2 A',     'sets': {5}},
    {'nome': 'MONTAGEM',         'sets': set()},
]
TODAS_LCS = [lc['nome'] for lc in LCS_DISPONIVEIS]

COLUNAS_MASSA = [
    "Sequencial", "EstruturaAltura", "LC_num", "WC_num", "LC",
    "Set", "Phase", "Attach",
    "S_Vert", "S_Trans", "S_Long",
    "B_Vert", "B_Trans", "B_Long",
    "A_Vert", "A_Trans", "A_Long", "Warnings"
]

CAMINHO_ESTRUTURAS = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "estruturas.xlsx"
)

CAMINHO_STRUCTURE_LOADS = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "Structure Loads.xlsx"
)

# ============================================================
# CADASTRO DE ESTRUTURAS
# ============================================================
@st.cache_data
def carregar_estruturas():
    if not os.path.exists(CAMINHO_ESTRUTURAS):
        return []

    def _float_ou_padrao(valor, padrao):
        try:
            if pd.isna(valor):
                return padrao
            texto = str(valor).strip().upper()
            if texto in ["", "N/A", "X", "NONE"]:
                return padrao
            return float(valor)
        except Exception:
            return padrao

    def _int_ou_padrao(valor, padrao):
        try:
            if pd.isna(valor):
                return padrao
            texto = str(valor).strip()
            if texto in ["", "N/A", "X", "NONE"]:
                return padrao
            return int(float(valor))
        except Exception:
            return padrao

    try:
        df = pd.read_excel(CAMINHO_ESTRUTURAS, engine="openpyxl")
        estruturas = []
        for _, row in df.iterrows():
            liberacao = str(row.iloc[1]).strip().upper() if pd.notna(row.iloc[1]) else ""
            poste_cod    = str(row.iloc[2]).strip().upper() if pd.notna(row.iloc[2]) else ""
            circuito_cod = str(row.iloc[4]).strip().upper() if pd.notna(row.iloc[4]) else ""
            num_postes   = _int_ou_padrao(row.iloc[7] if len(row) > 7 else None, 1)
            locacao_raw  = str(row.iloc[9]).strip().upper() if pd.notna(row.iloc[9]) else ""
            pr_cod       = str(row.iloc[10]).strip().upper() if pd.notna(row.iloc[10]) else ""

            tipo_poste = "Circular" if poste_cod == "C" else "Duplo T"
            is_duplo   = (circuito_cod == "CD")
            tem_pr     = (pr_cod == "S")

            if is_duplo and tem_pr:   config = 'CD_com_PR'
            elif is_duplo:            config = 'CD_sem_PR'
            elif tem_pr:              config = 'CS_com_PR'
            else:                     config = 'CS_sem_PR'

            estrutura = {
                "nome": str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else "",
                "liberacao": liberacao,
                "locacao_raw": locacao_raw,
                "tipo_poste": tipo_poste,
                "is_duplo": is_duplo,
                "tem_pr": tem_pr,
                "config": config,
                "num_postes": num_postes,
                "dist_pr":  _float_ou_padrao(row.iloc[12] if len(row) > 12 else None, 0.20),
                "dist_f1":  _float_ou_padrao(row.iloc[13] if len(row) > 13 else None, 0.20),
                "dist_f2":  _float_ou_padrao(row.iloc[14] if len(row) > 14 else None, 2.20),
                "dist_f3":  _float_ou_padrao(row.iloc[15] if len(row) > 15 else None, 4.20),
                "dist_f4":  _float_ou_padrao(row.iloc[16] if len(row) > 16 else None, 0.20),
                "dist_f5":  _float_ou_padrao(row.iloc[17] if len(row) > 17 else None, 2.20),
                "dist_f6":  _float_ou_padrao(row.iloc[18] if len(row) > 18 else None, 4.20),
                "dist_pr2": None,
            }
            if estrutura["nome"]:
                estruturas.append(estrutura)
        return estruturas
    except Exception:
        return []

# ============================================================
# PARSER INDIVIDUAL
# ============================================================
COLUNAS_INTERNAS = [
    "LC", "Set", "Phase", "Attach",
    "S_Vert", "S_Trans", "S_Long",
    "B_Vert", "B_Trans", "B_Long",
    "A_Vert", "A_Trans", "A_Long", "Warnings"
]

def parsear_tabela(texto):
    if not texto or not texto.strip():
        return None, "Nenhum texto colado."
    try:
        df = pd.read_csv(StringIO(texto), sep="\t", header=None, dtype=str, keep_default_na=False)
    except Exception as e:
        return None, f"Falha ao ler tabela: {e}"

    if df.shape[1] < 14:
        return None, f"Esperadas 14 colunas, encontradas {df.shape[1]}."

    df = df.iloc[:, :14].copy()
    df.columns = COLUNAS_INTERNAS

    primeira = str(df.iloc[0, 0]).strip().upper()
    if "LOAD" in primeira and "CASE" in primeira:
        df = df.iloc[1:].reset_index(drop=True)

    colunas_num = ["Set", "Phase", "S_Vert", "S_Trans", "S_Long",
                   "B_Vert", "B_Trans", "B_Long", "A_Vert", "A_Trans", "A_Long"]
    for c in colunas_num:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    df = df[df["LC"].astype(str).str.strip() != ""].reset_index(drop=True)

    obrig = ["Set", "Phase", "B_Vert", "B_Trans", "B_Long",
             "A_Vert", "A_Trans", "A_Long"]
    if df[obrig].isna().any().any():
        linhas_erro = df[df[obrig].isna().any(axis=1)].index.tolist()
        return None, f"Valores numéricos inválidos nas linhas (0-based): {linhas_erro[:5]}..."

    df["Set"] = df["Set"].astype(int)
    df["Phase"] = df["Phase"].astype(int)
    return df, None

def detectar_sets(df):
    try:
        vals = df["Set"].dropna().unique().tolist()
        return set(int(v) for v in vals if v is not None)
    except (ValueError, TypeError):
        return set()

def listar_lcs_disponiveis(sets_presentes):
    disponiveis = []
    for lc in LCS_DISPONIVEIS:
        if lc["sets"].issubset(sets_presentes):
            disponiveis.append(lc["nome"])
    return disponiveis

# ============================================================
# PARSER MASSA
# ============================================================
def parsear_tabela_massa(texto):
    if not texto or not texto.strip():
        return None, "Nenhum texto colado.", 0, 0

    try:
        df = pd.read_csv(StringIO(texto), sep="\t", header=None, dtype=str, keep_default_na=False)
    except Exception as e:
        return None, f"Falha ao ler: {e}", 0, 0

    if df.shape[1] < 18:
        return None, f"Esperadas 18 colunas, encontradas {df.shape[1]}.", 0, 0

    df = df.iloc[:, :18].copy()
    df.columns = COLUNAS_MASSA

    primeira = str(df.iloc[0, 0]).strip().upper()
    if "STR" in primeira:
        df = df.iloc[1:].reset_index(drop=True)

    cols_num = ["Set", "Phase",
                "B_Vert", "B_Trans", "B_Long",
                "A_Vert", "A_Trans", "A_Long"]
    for c in cols_num:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    df = df[df["Sequencial"].astype(str).str.strip() != ""].reset_index(drop=True)

    if df.empty:
        return None, "Nenhuma linha válida encontrada.", 0, 0

    obrig = ["Set", "Phase",
             "B_Vert", "B_Trans", "B_Long",
             "A_Vert", "A_Trans", "A_Long"]
    if df[obrig].isna().any().any():
        linhas_erro = df[df[obrig].isna().any(axis=1)].index.tolist()
        return None, f"Valores numéricos inválidos (linhas 0-based: {linhas_erro[:5]}...)", 0, 0

    df["Set"] = df["Set"].astype(int)
    df["Phase"] = df["Phase"].astype(int)

    n_linhas = len(df)
    n_estruturas = df["Sequencial"].nunique()
    return df, None, n_linhas, n_estruturas

def parsear_estrutura_altura(texto):
    texto = str(texto).strip()
    if "." not in texto:
        return texto, None
    partes = texto.split(".")
    if len(partes) != 2:
        return texto, None
    nome, altura_str = partes[0].strip(), partes[1].strip()
    if len(altura_str) != 3 or not altura_str.isdigit():
        return nome, None
    return nome, int(altura_str) / 10

def locacao_para_display(tipo_poste, locacao_raw):
    if tipo_poste == "Circular":
        return "N/A"
    if locacao_raw in ["L1", "L3"]:
        return "Topo para bissetriz"
    if locacao_raw in ["L2", "L4"]:
        return "Gaveta para bissetriz"
    return "N/A"

# ============================================================
# NÚCLEO DE CÁLCULO
# ============================================================
def calcular_altura_aerea(altura):
    enterrado = 0.10 * altura + 0.60
    return altura - enterrado

def calcular_K(altura_aerea, dist_cabo):
    X = altura_aerea - ALTURA_REFERENCIA
    if X <= 0:
        return None
    return (altura_aerea - dist_cabo) / X

def fator_correcao_duplo_t(theta_res, theta_face):
    delta = abs(theta_res - theta_face)
    if delta > 90:
        delta = 180 - delta
    rad = np.radians(delta)
    a, b = 1.0, 0.5
    return 1.0 / np.sqrt((np.cos(rad) / a) ** 2 + (np.sin(rad) / b) ** 2)

def calcular_forca_vento_poste(tipo_poste, altura, pressao_vento):
    h = altura
    if tipo_poste == "Duplo T":
        num = h**2 * (0.34524 + 0.01134 * h) * (1.1592 + 0.0252 * h)
        den = (2.3016 + 0.0756 * h) * (0.9 * h - 0.8)
        return 1.62 * pressao_vento * (num / den)
    else:
        num = h**2 * (0.42345 + 0.006075 * h) * (1.416 + 0.0135 * h)
        den = (2.823 + 0.0405 * h) * (0.9 * h - 0.8)
        return 1.08 * pressao_vento * (num / den)

def escolher_poste(R_final):
    if R_final <= 0:
        return COMERCIAIS[0], False
    for p in COMERCIAIS:
        if p >= R_final:
            return p, False
    return COMERCIAIS[-1], True

def processar_lc(nome_lc, df_lc, estrutura, altura, fator_cagaco_pct,
                 theta_face, considerar_vento, pressao_vento):
    altura_aerea = calcular_altura_aerea(altura)

    resultados_por_cabo = {}
    for _, row in df_lc.iterrows():
        try:
            chave = (int(row["Set"]), int(row["Phase"]))
        except (ValueError, TypeError):
            continue
        campo_alt = MAPEAMENTO_SET_PHASE.get(chave)
        if campo_alt is None:
            continue
        dist_cabo = estrutura.get(campo_alt)
        if dist_cabo is None:
            continue
        K = calcular_K(altura_aerea, dist_cabo)
        if K is None:
            continue

        resultados_por_cabo[chave] = {
            "K": K,
            "T": float(row["B_Trans"]) + float(row["A_Trans"]),
            "L": float(row["B_Long"])  + float(row["A_Long"]),
            "V": float(row["B_Vert"])  + float(row["A_Vert"]),
            "B_Trans": float(row["B_Trans"]),
            "A_Trans": float(row["A_Trans"]),
            "B_Long":  float(row["B_Long"]),
            "A_Long":  float(row["A_Long"]),
        }

    if not resultados_por_cabo:
        return None

    F_T = sum(r["K"] * r["T"] for r in resultados_por_cabo.values())
    F_L = sum(r["K"] * r["L"] for r in resultados_por_cabo.values())
    V_total = sum(r["V"] for r in resultados_por_cabo.values())

    F_vento_poste = 0.0
    if considerar_vento and nome_lc == "VENT MAX":
        F_vento_poste = calcular_forca_vento_poste(
            estrutura["tipo_poste"], altura, pressao_vento
        )
        if F_T < 0:
            F_vento_poste = -F_vento_poste
        F_T += F_vento_poste

    R = np.hypot(F_T, F_L)
    theta_R = np.degrees(np.arctan2(F_T, F_L))
    if theta_R < 0:
        theta_R += 360

    if estrutura["tipo_poste"] == "Duplo T" and theta_face is not None:
        fator_DT = fator_correcao_duplo_t(theta_R, theta_face)
        R_corr = R / fator_DT
    else:
        fator_DT = 1.0
        R_corr = R

    R_corr *= (1 + fator_cagaco_pct / 100.0)
    divisor = 1.0 if nome_lc == "EDS" else (1 + SOBRECARGA_TRANSITORIA)
    R_final = R_corr / divisor

    poste, estourou = escolher_poste(R_final)

    T_back  = sum(r["B_Trans"] for r in resultados_por_cabo.values())
    T_ahead = sum(r["A_Trans"] for r in resultados_por_cabo.values())
    L_back  = sum(r["B_Long"]  for r in resultados_por_cabo.values())
    L_ahead = sum(r["A_Long"]  for r in resultados_por_cabo.values())

    return {
        "LC": nome_lc, "V_total": V_total,
        "T_back": T_back, "T_ahead": T_ahead,
        "L_back": L_back, "L_ahead": L_ahead,
        "F_T": F_T, "F_L": F_L,
        "R": R, "theta_R": theta_R,
        "theta_face": theta_face,
        "fator_DT": fator_DT,
        "R_corr": R_corr, "divisor": divisor, "R_final": R_final,
        "poste": poste, "estourou": estourou,
        "F_vento_poste": F_vento_poste,
        "cabos": resultados_por_cabo,
    }

def processar_lc_poste(df_lc_completo, nome_lc, chaves_poste, estrutura, altura,
                       fator_cagaco_pct, theta_face, considerar_vento, pressao_vento):
    if df_lc_completo.empty:
        return None
    mask = df_lc_completo.apply(
        lambda r: (int(r["Set"]), int(r["Phase"])) in chaves_poste, axis=1
    )
    df_filtrado = df_lc_completo[mask]
    if df_filtrado.empty:
        return None
    return processar_lc(nome_lc, df_filtrado, estrutura, altura,
                        fator_cagaco_pct, theta_face, considerar_vento, pressao_vento)

def calcular_3_postes(df_lc, lc_nome, estrutura, altura_central, altura_lateral,
                      fator_cagaco, theta_face, considerar_vento, pressao_vento):
    tem_pr = estrutura["tem_pr"]

    chaves_p1 = {(1, 1)}
    chaves_p2 = {(1, 2)}
    chaves_p3 = {(1, 3)}
    if tem_pr:
        chaves_p2.add((4, 1))

    alt_p1 = altura_lateral if tem_pr else altura_central
    alt_p3 = altura_lateral if tem_pr else altura_central
    alt_p2 = altura_central

    res_p1 = processar_lc_poste(df_lc, lc_nome, chaves_p1, estrutura, alt_p1,
                                fator_cagaco, theta_face, considerar_vento, pressao_vento)
    res_p2 = processar_lc_poste(df_lc, lc_nome, chaves_p2, estrutura, alt_p2,
                                fator_cagaco, theta_face, considerar_vento, pressao_vento)
    res_p3 = processar_lc_poste(df_lc, lc_nome, chaves_p3, estrutura, alt_p3,
                                fator_cagaco, theta_face, considerar_vento, pressao_vento)

    if res_p1 is None and res_p2 is None and res_p3 is None:
        return None

    postes_lateral = [r["poste"] for r in [res_p1, res_p3] if r is not None]
    poste_lateral = max(postes_lateral) if postes_lateral else 0
    poste_central = res_p2["poste"] if res_p2 is not None else 0

    if tem_pr:
        poste_str = f"2× {poste_lateral} + 1× {poste_central}"
        poste_max = max(poste_lateral, poste_central)
    else:
        todos = [p for p in [poste_lateral, poste_central] if p > 0]
        poste_max = max(todos) if todos else 0
        poste_str = f"3× {poste_max}"

    R_final_max = max(
        [r["R_final"] for r in [res_p1, res_p2, res_p3] if r is not None]
    )
    estourou = any(
        r["estourou"] for r in [res_p1, res_p2, res_p3] if r is not None
    )

    return {
        "tipo": "3_postes",
        "LC": lc_nome,
        "res_p1": res_p1, "res_p2": res_p2, "res_p3": res_p3,
        "R_final_max": R_final_max,
        "poste_str": poste_str,
        "poste_max": poste_max,
        "estourou": estourou,
        "tem_pr": tem_pr,
        "poste_lateral": poste_lateral,
        "poste_central": poste_central,
    }

# ============================================================
# VISUALIZAÇÃO
# ============================================================
def plotar_diagrama(resultado, estrutura):
    fig, ax = plt.subplots(figsize=(8, 6))
    F_T, F_L = resultado["F_T"], resultado["F_L"]
    R, theta_R = resultado["R"], resultado["theta_R"]

    max_val = max(abs(F_T), abs(F_L), abs(R), 1e-6)
    scale = max_val / 1.2

    ax.axhline(y=0, color='k', linewidth=0.5, alpha=0.3)
    ax.axvline(x=0, color='k', linewidth=0.5, alpha=0.3)

    ax.arrow(0, 0, F_L / scale, 0, head_width=0.05, head_length=0.05,
             fc='blue', ec='blue', linewidth=1.5,
             label=f'Longitudinal: {F_L:.0f} daN')
    ax.arrow(0, 0, 0, F_T / scale, head_width=0.05, head_length=0.05,
             fc='red', ec='red', linewidth=1.5,
             label=f'Transversal: {F_T:.0f} daN')

    x_res = (R * np.cos(np.radians(theta_R))) / scale
    y_res = (R * np.sin(np.radians(theta_R))) / scale
    ax.arrow(0, 0, x_res, y_res, head_width=0.07, head_length=0.07,
             fc='green', ec='green', linewidth=2.5,
             label=f'Resultante: {R:.0f} daN')

    if estrutura["tipo_poste"] == "Duplo T" and resultado["theta_face"] is not None:
        tf = resultado["theta_face"]
        tg = tf + 90
        x_l = 1.2 * np.cos(np.radians(tf)); y_l = 1.2 * np.sin(np.radians(tf))
        ax.plot([-x_l, x_l], [-y_l, y_l], 'purple', linewidth=1.5,
                linestyle='--', alpha=0.8, label='Face Lisa')
        x_g = 1.2 * np.cos(np.radians(tg)); y_g = 1.2 * np.sin(np.radians(tg))
        ax.plot([-x_g, x_g], [-y_g, y_g], 'orange', linewidth=1.5,
                linestyle='--', alpha=0.8, label='Gaveta')
        a = 1.0 * (max_val / scale) * 0.6
        b = 0.5 * (max_val / scale) * 0.6
        ax.add_patch(Ellipse((0, 0), 2*a, 2*b, angle=tf, facecolor='none',
                             edgecolor='gray', linestyle=':', linewidth=1.5, alpha=0.7))

    ax.set_xlim(-1.35, 1.35); ax.set_ylim(-1.35, 1.35)
    ax.set_aspect('equal'); ax.grid(True, alpha=0.2)
    ax.set_xlabel('Longitudinal (X)', fontsize=10)
    ax.set_ylabel('Transversal (Y)', fontsize=10)
    ax.set_title(f'{resultado["LC"]} — R_final = {resultado["R_final"]:.1f} daN',
                 fontsize=11, fontweight='bold')
    ax.legend(loc='upper left', bbox_to_anchor=(1.02, 1), fontsize=8)
    plt.tight_layout()
    return fig

# ============================================================
# ESTADO DO STREAMLIT
# ============================================================
def inicializar_estado():
    defaults = {
        "textarea_pendente": "",
        "df_dados": None,
        "_erro_parse": None,
        "_limpar_textarea_flag": False,
        "_limpar_tabela_flag": False,
        "_marcar_todos_flag": False,
        "_desmarcar_todos_flag": False,
        "_estrutura_anterior": "",
        "altura": ALTURA_DEFAULT,
        "altura_lateral": ALTURA_DEFAULT,
        "fator_cagaco": 0,
        "pressao_vento": 30.0,
        "considerar_vento": False,
        "topo_gaveta": "Topo para bissetriz",
        "lcs_marcados": {},
        "resultados": None,
        "lc_critico": None,
        "massa_etapa": "colar",
        "massa_textarea": "",
        "massa_erro_parse": None,
        "massa_df_parse": None,
        "massa_linhas_total": 0,
        "massa_estruturas_total": 0,
        "massa_df_conf_source": None,
        "massa_lcs_marcados": {},
        "massa_marcar_flag": False,
        "massa_desmarcar_flag": False,
        "massa_fator_cagaco": 0,
        "massa_pressao_vento": 30.0,
        "massa_considerar_vento": False,
        "massa_resultado_df": None,
        "massa_confirmou_orientacoes": False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

def reset_geral():
    for k in list(st.session_state.keys()):
        del st.session_state[k]
    inicializar_estado()

def reset_massa():
    chaves_massa = [k for k in st.session_state if k.startswith("massa_") or k == "massa_editor"]
    for k in chaves_massa:
        if k in st.session_state:
            del st.session_state[k]
    if "massa_editor" in st.session_state:
        del st.session_state["massa_editor"]
    inicializar_estado()

# ============================================================
# TAB: INDIVIDUAL (INALTERADA)
# ============================================================
def render_individual_tab(estruturas_liberadas):
    if not estruturas_liberadas:
        st.error("Nenhuma estrutura liberada no cadastro.")
        st.stop()

    nomes = [e["nome"] for e in estruturas_liberadas]
    opcoes_estrutura = [""] + nomes

    st.subheader("🏗️ 1. Estrutura")

    if "sel_estrutura" not in st.session_state or st.session_state["sel_estrutura"] not in opcoes_estrutura:
        st.session_state["sel_estrutura"] = ""

    estrutura_nome = st.selectbox(
        "Selecione a estrutura:", opcoes_estrutura,
        key="sel_estrutura"
    )

    estrutura = None
    if estrutura_nome:
        estrutura_base = next(e for e in estruturas_liberadas if e["nome"] == estrutura_nome)
        num_postes = estrutura_base.get("num_postes", 1)

        if st.session_state.get("_estrutura_anterior") != estrutura_nome:
            st.session_state["edit_tipo_poste"] = estrutura_base["tipo_poste"]
            st.session_state["edit_config"] = estrutura_base["config"]
            st.session_state["edit_dist_f1"] = float(estrutura_base["dist_f1"])
            st.session_state["edit_dist_f2"] = float(estrutura_base["dist_f2"])
            st.session_state["edit_dist_f3"] = float(estrutura_base["dist_f3"])
            st.session_state["edit_dist_f4"] = float(estrutura_base["dist_f4"])
            st.session_state["edit_dist_f5"] = float(estrutura_base["dist_f5"])
            st.session_state["edit_dist_f6"] = float(estrutura_base["dist_f6"])
            st.session_state["edit_dist_pr"] = float(estrutura_base["dist_pr"])
            st.session_state["_estrutura_anterior"] = estrutura_nome

        if num_postes == 2:
            st.error(
                "⚠️ Estrutura de **2 postes** ainda não é suportada. "
                "Selecione outra estrutura ou corrija o cadastro."
            )
            st.stop()

        st.markdown("**1.1 Parâmetros da estrutura (editáveis)**")
        st.caption(
            "Os valores abaixo vêm do cadastro (`estruturas.xlsx`), mas podem ser "
            "ajustados aqui. O cálculo usa os valores editados."
        )

        col_a, col_b, col_c = st.columns([1, 1, 1])
        with col_a:
            tipo_poste_edit = st.selectbox(
                "Tipo de poste", ["Circular", "Duplo T"], key="edit_tipo_poste")
        with col_b:
            config_edit = st.selectbox(
                "Configuração", CONFIG_OPCOES, key="edit_config",
                format_func=lambda c: CONFIG_LABELS.get(c, c))
        with col_c:
            st.write("")
            if st.button("🔄 Recarregar estrutura", use_container_width=True):
                st.session_state["_estrutura_anterior"] = ""
                st.rerun()

        tem_pr = config_edit.endswith("_com_PR")
        is_duplo = config_edit.startswith("CD_")

        if num_postes == 3:
            st.info(
                f"📐 **Estrutura de 3 postes** — "
                f"P1: Fase A · P2 (central): Fase B{' + PR' if tem_pr else ''} · P3: Fase C"
            )

        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            dist_f1_edit = st.number_input(
                "Fase 1 — C1 (m)" + (" [altura P1]" if num_postes == 3 else ""),
                min_value=0.0, max_value=30.0, step=0.05,
                format="%.2f", key="edit_dist_f1")
        with col_f2:
            dist_f2_edit = st.number_input(
                "Fase 2 — C1 (m)" + (" [altura P2]" if num_postes == 3 else ""),
                min_value=0.0, max_value=30.0, step=0.05,
                format="%.2f", key="edit_dist_f2")
        with col_f3:
            dist_f3_edit = st.number_input(
                "Fase 3 — C1 (m)" + (" [altura P3]" if num_postes == 3 else ""),
                min_value=0.0, max_value=30.0, step=0.05,
                format="%.2f", key="edit_dist_f3")

        if is_duplo and num_postes == 1:
            col_f4, col_f5, col_f6 = st.columns(3)
            with col_f4:
                dist_f4_edit = st.number_input(
                    "Fase 4 — C2 (m)", min_value=0.0, max_value=30.0, step=0.05,
                    format="%.2f", key="edit_dist_f4")
            with col_f5:
                dist_f5_edit = st.number_input(
                    "Fase 5 — C2 (m)", min_value=0.0, max_value=30.0, step=0.05,
                    format="%.2f", key="edit_dist_f5")
            with col_f6:
                dist_f6_edit = st.number_input(
                    "Fase 6 — C2 (m)", min_value=0.0, max_value=30.0, step=0.05,
                    format="%.2f", key="edit_dist_f6")
        else:
            dist_f4_edit = st.session_state.get("edit_dist_f4", estrutura_base["dist_f4"])
            dist_f5_edit = st.session_state.get("edit_dist_f5", estrutura_base["dist_f5"])
            dist_f6_edit = st.session_state.get("edit_dist_f6", estrutura_base["dist_f6"])

        if tem_pr:
            label_pr = "Para-raios — PR1 (m)"
            if num_postes == 3:
                label_pr += " [altura PR no P2]"
            dist_pr_edit = st.number_input(
                label_pr, min_value=0.0, max_value=30.0, step=0.05,
                format="%.2f", key="edit_dist_pr")
        else:
            dist_pr_edit = st.session_state.get("edit_dist_pr", estrutura_base["dist_pr"])

        estrutura = dict(estrutura_base)
        estrutura["tipo_poste"] = tipo_poste_edit
        estrutura["config"] = config_edit
        estrutura["is_duplo"] = is_duplo
        estrutura["tem_pr"] = tem_pr
        estrutura["num_postes"] = num_postes
        estrutura["dist_f1"] = dist_f1_edit
        estrutura["dist_f2"] = dist_f2_edit
        estrutura["dist_f3"] = dist_f3_edit
        estrutura["dist_f4"] = dist_f4_edit
        estrutura["dist_f5"] = dist_f5_edit
        estrutura["dist_f6"] = dist_f6_edit
        estrutura["dist_pr"] = dist_pr_edit

        sets_req = SETS_REQUERIDOS[estrutura["config"]]
        st.info(
            f"**{estrutura['tipo_poste']}** — Configuração: "
            f"**{CONFIG_LABELS.get(estrutura['config'], estrutura['config'])}** — "
            f"Sets requeridos: **{sorted(sets_req)}**"
        )

    st.markdown("---")
    st.subheader("📋 2. Cole a tabela do PLS-CADD")

    if st.session_state.get("_limpar_textarea_flag"):
        st.session_state["textarea_pendente"] = ""
        st.session_state["_limpar_textarea_flag"] = False

    if st.session_state.get("_limpar_tabela_flag"):
        st.session_state["df_dados"] = None
        st.session_state["_erro_parse"] = None
        if "editor_dados" in st.session_state:
            del st.session_state["editor_dados"]
        st.session_state["_limpar_tabela_flag"] = False

    if estrutura is not None:
        st.caption(
            "Cole apenas os VALORES (sem a linha de cabeçalho). "
            "Ordem: LC | Set | Phase | Attach | S_Vert | S_Trans | S_Long | "
            "B_Vert | B_Trans | B_Long | A_Vert | A_Trans | A_Long | Warnings"
        )
        st.text_area("Cole aqui (Ctrl+V do Excel):", height=200, key="textarea_pendente")

        cbtn_ap, cbtn_limpar_campo = st.columns([1, 1])
        with cbtn_ap:
            if st.button("📥 Aplicar tabela", use_container_width=True):
                df_parsed, erro = parsear_tabela(st.session_state.get("textarea_pendente", ""))
                if erro:
                    st.session_state["_erro_parse"] = erro
                else:
                    st.session_state["_erro_parse"] = None
                    st.session_state["df_dados"] = df_parsed
                    if "editor_dados" in st.session_state:
                        del st.session_state["editor_dados"]
                    st.session_state["resultados"] = None
                st.rerun()
        with cbtn_limpar_campo:
            if st.button("🗑️ Limpar campo", use_container_width=True):
                st.session_state["_limpar_textarea_flag"] = True
                st.session_state["_limpar_tabela_flag"] = True
                st.session_state["resultados"] = None
                st.session_state["lc_critico"] = None
                st.rerun()

        if st.session_state.get("_erro_parse"):
            st.error(f"❌ {st.session_state['_erro_parse']}")

        df = None
        lcs_disponiveis = []

        if st.session_state.get("df_dados") is not None:
            st.markdown("---")
            st.subheader("📝 2.1 Dados colados (editáveis)")
            df_editavel = st.data_editor(
                st.session_state["df_dados"],
                key="editor_dados",
                num_rows="fixed",
                use_container_width=True,
                column_config={
                    "LC": st.column_config.TextColumn("LC"),
                    "Set": st.column_config.NumberColumn("Set", step=1, format="%d", min_value=0),
                    "Phase": st.column_config.NumberColumn("Phase", step=1, format="%d", min_value=0),
                    "Attach": st.column_config.TextColumn("Attach"),
                },
            )
            sets_presentes = detectar_sets(df_editavel)
            sets_req = SETS_REQUERIDOS[estrutura["config"]]
            if sets_presentes != sets_req:
                faltam = sets_req - sets_presentes
                sobram = sets_presentes - sets_req
                msg = "❌ Incompatibilidade de sets após edição."
                if faltam: msg += f"  Faltam: {sorted(faltam)}."
                if sobram: msg += f"  Sobram: {sorted(sobram)}."
                st.error(msg)
            else:
                st.success(f"✅ Tabela válida — {len(df_editavel)} linhas, sets {sorted(sets_presentes)}")
                df = df_editavel
                lcs_disponiveis = listar_lcs_disponiveis(sets_presentes)

        if df is not None and lcs_disponiveis:
            st.markdown("---")
            st.subheader("☑️ 3. Hipóteses a processar")

            if st.session_state.get("_marcar_todos_flag"):
                for lc in lcs_disponiveis:
                    st.session_state[f"chk_{lc}"] = True
                    st.session_state["lcs_marcados"][lc] = True
                st.session_state["_marcar_todos_flag"] = False
            if st.session_state.get("_desmarcar_todos_flag"):
                for lc in lcs_disponiveis:
                    st.session_state[f"chk_{lc}"] = False
                    st.session_state["lcs_marcados"][lc] = False
                st.session_state["_desmarcar_todos_flag"] = False

            cb1, cb2 = st.columns(2)
            with cb1:
                if st.button("✓ Marcar todos", use_container_width=True):
                    st.session_state["_marcar_todos_flag"] = True
                    st.rerun()
            with cb2:
                if st.button("✗ Desmarcar todos", use_container_width=True):
                    st.session_state["_desmarcar_todos_flag"] = True
                    st.rerun()

            cols = st.columns(4)
            for i, lc in enumerate(lcs_disponiveis):
                with cols[i % 4]:
                    marcado = st.checkbox(lc, value=st.session_state["lcs_marcados"].get(lc, False),
                                          key=f"chk_{lc}")
                    st.session_state["lcs_marcados"][lc] = marcado

            st.markdown("---")
            st.subheader("⚙️ 4. Configurações")

            num_postes = estrutura.get("num_postes", 1)
            tem_pr = estrutura["tem_pr"]

            if num_postes == 3 and tem_pr:
                c1, c2 = st.columns(2)
                with c1:
                    if st.session_state["altura"] not in ALTURAS_DISPONIVEIS:
                        st.session_state["altura"] = ALTURA_DEFAULT
                    altura_central = st.selectbox(
                        "Altura do poste central (m)", ALTURAS_DISPONIVEIS,
                        key="sel_altura")
                    st.session_state["altura"] = altura_central
                with c2:
                    if st.session_state["altura_lateral"] not in ALTURAS_DISPONIVEIS:
                        st.session_state["altura_lateral"] = ALTURA_DEFAULT
                    altura_lateral = st.selectbox(
                        "Altura dos postes laterais (m)", ALTURAS_DISPONIVEIS,
                        key="sel_altura_lateral")
                    st.session_state["altura_lateral"] = altura_lateral
            else:
                if st.session_state["altura"] not in ALTURAS_DISPONIVEIS:
                    st.session_state["altura"] = ALTURA_DEFAULT
                altura = st.selectbox(
                    "Altura do poste (m)", ALTURAS_DISPONIVEIS,
                    key="sel_altura")
                st.session_state["altura"] = altura
                st.session_state["altura_lateral"] = altura

            c1, c2, c3 = st.columns(3)
            with c1:
                fator_cagaco = st.number_input("Fator cagaço (%)", min_value=0, max_value=20,
                                                value=st.session_state["fator_cagaco"], step=1,
                                                key="inp_cagaco")
                st.session_state["fator_cagaco"] = fator_cagaco
            with c2:
                pressao_vento = st.number_input("Pressão de vento (kgf/m²)", min_value=0.0,
                                                 max_value=200.0,
                                                 value=st.session_state["pressao_vento"], step=1.0,
                                                 key="inp_vento")
                st.session_state["pressao_vento"] = pressao_vento
            with c3:
                st.write("")

            considerar_vento = st.checkbox(
                "Considerar vento no fuste (apenas VENT MAX) — INATIVO por enquanto",
                value=st.session_state["considerar_vento"], key="chk_vento")
            st.session_state["considerar_vento"] = considerar_vento

            theta_face = None
            if estrutura["tipo_poste"] == "Duplo T":
                topo_gaveta = st.selectbox(
                    "Montagem Duplo T",
                    ["Topo para bissetriz", "Gaveta para bissetriz"],
                    index=0 if st.session_state["topo_gaveta"] == "Topo para bissetriz" else 1,
                    key="sel_topo")
                st.session_state["topo_gaveta"] = topo_gaveta
                theta_face = 90.0 if topo_gaveta == "Topo para bissetriz" else 0.0

            st.markdown("---")
            cbtn1, cbtn2, cbtn3 = st.columns(3)
            with cbtn1:
                botao_calcular = st.button(
                    "🔍 Calcular", type="primary", use_container_width=True,
                    disabled=(not any(st.session_state["lcs_marcados"].values())))
                if botao_calcular:
                    lcs_sel = [lc for lc, v in st.session_state["lcs_marcados"].items() if v]
                    resultados = []
                    for lc_nome in lcs_sel:
                        df_lc = df[df["LC"].astype(str).str.strip() == lc_nome]
                        if df_lc.empty: continue

                        if num_postes == 3:
                            res = calcular_3_postes(
                                df_lc, lc_nome, estrutura,
                                st.session_state["altura"],
                                st.session_state["altura_lateral"],
                                st.session_state["fator_cagaco"],
                                theta_face,
                                st.session_state["considerar_vento"],
                                st.session_state["pressao_vento"])
                        else:
                            res_lc = processar_lc(lc_nome, df_lc, estrutura,
                                                   st.session_state["altura"],
                                                   st.session_state["fator_cagaco"],
                                                   theta_face,
                                                   st.session_state["considerar_vento"],
                                                   st.session_state["pressao_vento"])
                            res = None
                            if res_lc is not None:
                                res = {
                                    "tipo": "1_poste",
                                    "LC": lc_nome,
                                    "res_unico": res_lc,
                                    "R_final_max": res_lc["R_final"],
                                    "poste_str": str(res_lc["poste"]),
                                    "poste_max": res_lc["poste"],
                                    "estourou": res_lc["estourou"],
                                }

                        if res is not None:
                            resultados.append(res)

                    st.session_state["resultados"] = resultados
                    if resultados:
                        crit = max(resultados, key=lambda x: x["R_final_max"])
                        st.session_state["lc_critico"] = crit["LC"]
                        st.session_state["lc_grafico"] = crit["LC"]
            with cbtn2:
                if st.button("🧹 Limpar tabela e resultados", use_container_width=True):
                    st.session_state["_limpar_textarea_flag"] = True
                    st.session_state["_limpar_tabela_flag"] = True
                    st.session_state["resultados"] = None
                    st.session_state["lc_critico"] = None
                    st.rerun()
            with cbtn3:
                if st.button("🔄 Reset geral", use_container_width=True):
                    reset_geral()
                    st.rerun()

    if st.session_state["resultados"]:
        st.markdown("---")
        st.subheader("📊 Resultados")
        resultados = st.session_state["resultados"]
        num_postes = estrutura.get("num_postes", 1) if estrutura else 1

        if num_postes == 1:
            linhas = []
            for r in resultados:
                res = r["res_unico"]
                linhas.append({
                    "LC": res["LC"],
                    "V_total": f"{res['V_total']:.1f}",
                    "T_back": f"{res['T_back']:.1f}", "T_ahead": f"{res['T_ahead']:.1f}",
                    "L_back": f"{res['L_back']:.1f}", "L_ahead": f"{res['L_ahead']:.1f}",
                    "F_T (transf.)": f"{res['F_T']:.1f}", "F_L (transf.)": f"{res['F_L']:.1f}",
                    "R (daN)": f"{res['R']:.1f}", "θ_R (°)": f"{res['theta_R']:.1f}",
                    "θ_face (°)": f"{res['theta_face']:.1f}" if res['theta_face'] is not None else "—",
                    "Fator DT": f"{res['fator_DT']:.3f}",
                    "R_corr (daN)": f"{res['R_corr']:.1f}",
                    "Divisor": f"{res['divisor']:.2f}",
                    "R_final (daN)": f"{res['R_final']:.1f}",
                    "Poste": res["poste"],
                })
            st.dataframe(pd.DataFrame(linhas), use_container_width=True, hide_index=True)

            poste_final = max(r["poste_max"] for r in resultados)
            crit = next(r for r in resultados if r["LC"] == st.session_state["lc_critico"])
            st.markdown("### 🎯 Resumo")
            cA, cB, cC = st.columns(3)
            with cA: st.metric("Poste recomendado", f"{poste_final} daN")
            with cB: st.metric("LC crítico", crit["LC"])
            with cC: st.metric("R_final do crítico", f"{crit['R_final_max']:.1f} daN")

            if any(r["estourou"] for r in resultados):
                st.warning("⚠️ Pelo menos um LC ultrapassou 4000 daN.")

            st.markdown("### 📈 Diagrama vetorial")
            opcoes = [r["LC"] for r in resultados]
            if "sel_lc_grafico" not in st.session_state or st.session_state["sel_lc_grafico"] not in opcoes:
                st.session_state["sel_lc_grafico"] = opcoes[0] if opcoes else ""
            lc_graf = st.selectbox("LC para visualizar:", opcoes, key="sel_lc_grafico")
            res_graf = next(r for r in resultados if r["LC"] == lc_graf)
            fig = plotar_diagrama(res_graf["res_unico"], estrutura)
            st.pyplot(fig)
            plt.close(fig)

        elif num_postes == 3:
            tem_pr = estrutura["tem_pr"]

            linhas = []
            for r in resultados:
                linha = {
                    "LC": r["LC"],
                    "Poste lateral": r["poste_lateral"],
                    "Poste central": r["poste_central"],
                    "Consolidado": r["poste_str"],
                    "R_final máx (daN)": f"{r['R_final_max']:.1f}",
                }
                linhas.append(linha)
            st.dataframe(pd.DataFrame(linhas), use_container_width=True, hide_index=True)

            if tem_pr:
                poste_lateral_final = max(r["poste_lateral"] for r in resultados)
                poste_central_final = max(r["poste_central"] for r in resultados)
                poste_str_final = f"2× {poste_lateral_final} + 1× {poste_central_final}"
            else:
                poste_final = max(r["poste_max"] for r in resultados)
                poste_str_final = f"3× {poste_final}"

            crit = next(r for r in resultados if r["LC"] == st.session_state["lc_critico"])
            st.markdown("### 🎯 Resumo")
            cA, cB, cC = st.columns(3)
            with cA: st.metric("Poste recomendado", poste_str_final)
            with cB: st.metric("LC crítico", crit["LC"])
            with cC: st.metric("R_final do crítico", f"{crit['R_final_max']:.1f} daN")

            if any(r["estourou"] for r in resultados):
                st.warning("⚠️ Pelo menos um LC ultrapassou 4000 daN.")

            st.markdown("### 📈 Diagramas vetoriais (LC crítico)")
            crit_res = next(r for r in resultados if r["LC"] == st.session_state["lc_critico"])

            abas = []
            if crit_res["res_p1"]: abas.append(("P1 (lateral esq.)", crit_res["res_p1"]))
            if crit_res["res_p2"]: abas.append(("P2 (central)", crit_res["res_p2"]))
            if crit_res["res_p3"]: abas.append(("P3 (lateral dir.)", crit_res["res_p3"]))

            if abas:
                tab_objs = st.tabs([a[0] for a in abas])
                for i, (label, res) in enumerate(abas):
                    with tab_objs[i]:
                        fig = plotar_diagrama(res, estrutura)
                        st.pyplot(fig)
                        plt.close(fig)

# ============================================================
# TAB: MASSA
# ============================================================
def _calcular_altura_lateral_inicial(cadastro, altura_central):
    num_postes = cadastro.get("num_postes", 1)
    tem_pr = cadastro.get("tem_pr", False)
    if num_postes != 3:
        return "N/A"
    if tem_pr:
        return ""
    if altura_central is None or pd.isna(altura_central):
        return "N/A"
    try:
        return str(float(altura_central))
    except (ValueError, TypeError):
        return "N/A"

def _aplicar_regra_altura_lateral(df_editado, estruturas_dict):
    df = df_editado.copy()
    for idx in df.index:
        nome = df.at[idx, "Estrutura"]
        cadastro = estruturas_dict.get(nome)
        if cadastro is None:
            df.at[idx, "Altura lateral"] = "N/A"
            continue
        num_postes = cadastro.get("num_postes", 1)
        tem_pr = cadastro.get("tem_pr", False)
        if num_postes != 3:
            df.at[idx, "Altura lateral"] = "N/A"
        elif not tem_pr:
            val_central = df.at[idx, "Altura"]
            if pd.isna(val_central):
                df.at[idx, "Altura lateral"] = "N/A"
            else:
                try:
                    df.at[idx, "Altura lateral"] = str(float(val_central))
                except (ValueError, TypeError):
                    df.at[idx, "Altura lateral"] = "N/A"
    return df

def construir_conferencia(df_parse, estruturas_dict):
    sequenciais = []
    for s in df_parse["Sequencial"]:
        if s not in sequenciais:
            sequenciais.append(s)

    rows = []
    for seq in sequenciais:
        df_grupo = df_parse[df_parse["Sequencial"] == seq]
        primeiro = df_grupo.iloc[0]
        nome, altura = parsear_estrutura_altura(primeiro["EstruturaAltura"])
        cadastro = estruturas_dict.get(nome)

        if cadastro is None:
            rows.append({
                "Sequencial": seq, "Estrutura": nome,
                "Altura": altura if altura is not None else 0,
                "Altura lateral": "N/A",
                "Tipo": "N/A", "Config": "N/A", "Locacao": "N/A",
                "dist_f1": 0.0, "dist_f2": 0.0, "dist_f3": 0.0,
                "dist_f4": 0.0, "dist_f5": 0.0, "dist_f6": 0.0, "dist_pr": 0.0,
                "Estado": "❌ Não homologada",
                "Motivo": f"Estrutura '{nome}' não encontrada no cadastro",
            })
            continue

        tipo = cadastro["tipo_poste"]
        config = cadastro["config"]
        locacao = locacao_para_display(tipo, cadastro["locacao_raw"])

        altura_lateral = _calcular_altura_lateral_inicial(cadastro, altura)

        if cadastro["liberacao"] != "SIM":
            estado, motivo = "⚠️ Não liberada", "Estrutura existe mas não está liberada"
        elif altura is None or altura not in ALTURAS_DISPONIVEIS:
            estado, motivo = "🔶 Altura incompatível", f"Altura {altura} fora da lista"
        elif tipo == "Duplo T" and locacao == "N/A":
            estado, motivo = "🟣 Locação não definida", f"Locacao '{cadastro['locacao_raw']}' não mapeada"
        else:
            estado, motivo = "✅ Calculada", ""

        rows.append({
            "Sequencial": seq, "Estrutura": nome,
            "Altura": altura if altura is not None else 0,
            "Altura lateral": altura_lateral,
            "Tipo": tipo, "Config": config, "Locacao": locacao,
            "dist_f1": cadastro["dist_f1"], "dist_f2": cadastro["dist_f2"],
            "dist_f3": cadastro["dist_f3"], "dist_f4": cadastro["dist_f4"],
            "dist_f5": cadastro["dist_f5"], "dist_f6": cadastro["dist_f6"],
            "dist_pr": cadastro["dist_pr"],
            "Estado": estado, "Motivo": motivo,
        })

    return pd.DataFrame(rows)

def recomputar_estado(row, estruturas_dict):
    nome = row["Estrutura"]
    altura = row["Altura"]
    tipo = row["Tipo"]
    locacao = row["Locacao"]
    altura_lateral = row.get("Altura lateral", None)

    cadastro = estruturas_dict.get(nome)
    if cadastro is None:
        return "❌ Não homologada", f"Estrutura '{nome}' não encontrada no cadastro"
    if cadastro["liberacao"] != "SIM":
        return "⚠️ Não liberada", "Estrutura existe mas não está liberada"
    try:
        altura_f = float(altura)
    except (ValueError, TypeError):
        return "🔶 Altura incompatível", f"Altura inválida: {altura}"
    if altura_f not in ALTURAS_DISPONIVEIS:
        return "🔶 Altura incompatível", f"Altura {altura_f} fora da lista"
    if tipo == "Duplo T" and locacao not in ["Topo para bissetriz", "Gaveta para bissetriz"]:
        return "🟣 Locação não definida", f"Locacao '{locacao}' não mapeada"

    if cadastro.get("num_postes", 1) == 3 and cadastro.get("tem_pr", False):
        al_txt = "" if altura_lateral is None else str(altura_lateral).strip()
        if al_txt == "" or al_txt.upper() == "N/A":
            return "🔶 Altura lateral pendente", "Necessário informar altura do poste lateral"
        try:
            al_f = float(al_txt.replace(",", "."))
        except (ValueError, TypeError):
            return "🔶 Altura lateral pendente", f"Altura lateral inválida: {altura_lateral}"
        if al_f not in ALTURAS_DISPONIVEIS:
            return "🔶 Altura lateral pendente", f"Altura lateral {al_f} fora da lista"

    return "✅ Calculada", ""

def render_massa_tab(estruturas_dict):
    etapa = st.session_state.get("massa_etapa", "colar")

    if etapa == "colar":
        _massa_etapa_colar()
    elif etapa == "confirmar":
        _massa_etapa_confirmar()
    elif etapa == "conferir":
        _massa_etapa_conferir(estruturas_dict)
    elif etapa == "resultado":
        _massa_etapa_resultado()

def _massa_etapa_colar():
    st.subheader("📋 1. Orientações de uso e colagem da tabela")

    with st.expander("📖 Leia atentamente as orientações de uso antes de prosseguir", expanded=True):
        st.markdown("""
**1. Cadastre todas as 20 hipóteses de carregamento.**
Hipóteses faltantes não são processadas e o dimensionamento refletirá apenas as que vierem — o que tende a **subestimar** o poste necessário.

**2. Use exatamente estes nomes.**
O VHPT 2.0 identifica cada hipótese **pelo nome exato**, sem tolerância a diferenças de maiúsculas/minúsculas, acentos, espaços ou pontuação. Nomes diferentes fazem a hipótese ser **ignorada** no cálculo.

~~~
EDS
TEMP MÍM
VENT MAX
QUEBRA C1 FA B
QUEBRA C1 FA A
QUEBRA C1 FB B
QUEBRA C1 FB A
QUEBRA C1 FC B
QUEBRA C1 FC A
QUEBRA C2 FA B
QUEBRA C2 FA A
QUEBRA C2 FB B
QUEBRA C2 FB A
QUEBRA C2 FC B
QUEBRA C2 FC A
QUEBRA PR1 B
QUEBRA PR1 A
QUEBRA PR2 B
QUEBRA PR2 A
MONTAGEM
~~~

**3. Cadastre as hipóteses no PLS-CADD conforme o arquivo padrão.**
O arquivo **`Structure Loads.xlsx`** está disponível no mesmo diretório do VHPT 2.0 e é o **padrão obrigatório** de cadastro. Importe-o no PLS-CADD pelo menu **`Criteria` → `Structure Loads (Methods 3,4)`**. Não altere nomes nem crie hipóteses fora desse arquivo — as consequências estão nos itens 2 e 3.
""")

        if os.path.exists(CAMINHO_STRUCTURE_LOADS):
            with open(CAMINHO_STRUCTURE_LOADS, "rb") as _f:
                st.download_button(
                    "📥 Baixar Structure Loads.xlsx (padrão obrigatório)",
                    _f.read(),
                    "Structure Loads.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="massa_dl_structure_loads",
                )
        else:
            st.warning(
                "⚠️ Arquivo `Structure Loads.xlsx` não encontrado no diretório do app. "
                "Contate o administrador."
            )

        st.markdown("""
**4. Formato do `Str. Name`.**
Formato obrigatório: **`NomeEstrutura.AlturaEmDecímetros`**
- `NomeEstrutura` = nome **exato** cadastrado em `estruturas.xlsx`
- `AlturaEmDecímetros` = **3 dígitos**, altura × 10 (ex.: `220` = 22,0 m)

Exemplo: `UA-GV-PR.220`

Nome divergente → **❌ Não homologada**. Altura fora da lista (`18, 20, 22, 24, 26, 28, 30, 32`) → **🔶 Altura incompatível**.
""")

        if os.path.exists(CAMINHO_ESTRUTURAS):
            with open(CAMINHO_ESTRUTURAS, "rb") as _f:
                st.download_button(
                    "📥 Baixar estruturas.xlsx (conferir nomes homologados)",
                    _f.read(),
                    "estruturas.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="massa_dl_estruturas",
                )
        else:
            st.warning(
                "⚠️ Planilha `estruturas.xlsx` não encontrada no diretório do app. "
                "Contate o administrador."
            )

        st.markdown("""

**5. Não misture trechos com configurações de circuito diferentes no mesmo lote** (`CD_com_PR` com Sets {1, 2, 4} + `CS_com_PR` com Sets {1, 4}). Processe em lotes separados — o cálculo em massa não valida compatibilidade de sets e pode **subestimar** os esforços.

**6. PLS-CADD exporta apenas a altura do poste central.** Para estruturas de **3 postes com PR**, a altura dos laterais precisa ser informada manualmente na coluna **Altura lateral**.

**7. Configuração dos Sets — todas as estruturas, sem exceção:**

| Set | Destinação |
|---|---|
| **1** | Condutores do **circuito 1** |
| **2** | Condutores do **circuito 2** (se houver) |
| **4** | Cabo **para-raios 1** (se houver) |
| **5** | Cabo **para-raios 2** (se houver) |

Sets trocados ou fora dessa numeração fazem o VHPT 2.0 atribuir esforços à estrutura errada — **sub ou superdimensionamento silencioso**.
""")

    st.session_state["massa_confirmou_orientacoes"] = st.checkbox(
        "✅ Li e entendi as orientações de uso acima. Estou ciente de que o uso "
        "incorreto da ferramenta pode gerar dimensionamento inadequado.",
        value=st.session_state.get("massa_confirmou_orientacoes", False),
        key="massa_chk_confirmacao",
    )

    st.markdown("---")
    st.subheader("📋 2. Cole a tabela em massa")
    st.caption(
        "Cole a tabela do PLS-CADD (18 colunas, com Str. No. e Str. Name). "
        "A coluna 'Str. Name' deve vir no formato `Nome.Altura×10` (ex.: `UA-GV-PR.220`)."
    )
    st.text_area("Cole aqui:", height=250, key="massa_textarea")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("📊 Analisar tabela", type="primary", use_container_width=True):
            texto = st.session_state.get("massa_textarea", "")
            df, erro, n_linhas, n_estruturas = parsear_tabela_massa(texto)
            if erro:
                st.session_state["massa_erro_parse"] = erro
            else:
                st.session_state["massa_df_parse"] = df
                st.session_state["massa_linhas_total"] = n_linhas
                st.session_state["massa_estruturas_total"] = n_estruturas
                st.session_state["massa_erro_parse"] = None
                st.session_state["massa_etapa"] = "confirmar"
            st.rerun()
    with col2:
        if st.button("🔄 Limpar", use_container_width=True):
            reset_massa()
            st.rerun()

    if st.session_state.get("massa_erro_parse"):
        st.error(f"❌ {st.session_state['massa_erro_parse']}")

def _massa_etapa_confirmar():
    st.subheader("📊 Confirmação da tabela")
    n_linhas = st.session_state["massa_linhas_total"]
    n_estruturas = st.session_state["massa_estruturas_total"]

    c1, c2 = st.columns(2)
    with c1: st.metric("Linhas totais detectadas", n_linhas)
    with c2: st.metric("Estruturas detectadas", n_estruturas)

    col1, col2 = st.columns(2)
    with col1:
        if st.button("🔄 Voltar e colar novamente", use_container_width=True):
            st.session_state["massa_etapa"] = "colar"
            st.rerun()
    with col2:
        if st.button("✅ Prosseguir", type="primary", use_container_width=True):
            df_conf = construir_conferencia(
                st.session_state["massa_df_parse"],
                st.session_state["_estruturas_dict"]
            )
            st.session_state["massa_df_conf_source"] = df_conf
            st.session_state["massa_etapa"] = "conferir"
            st.rerun()

def _massa_etapa_conferir(estruturas_dict):
    st.subheader("📝 3. Planilha de conferência")

    if st.session_state.get("massa_df_conf_source") is None:
        st.error("Erro interno: planilha não construída.")
        if st.button("Voltar"):
            st.session_state["massa_etapa"] = "colar"
            st.rerun()
        return

    if not st.session_state.get("massa_confirmou_orientacoes", False):
        st.warning(
            "⚠️ **Confirmação de leitura das orientações de uso pendente.** "
            "As orientações estão na **etapa 1** (primeira tela da aba massa). "
            "Marque abaixo para liberar o cálculo, ou volte à etapa 1 para revisá-las."
        )
        st.session_state["massa_confirmou_orientacoes"] = st.checkbox(
            "✅ Li e entendi as orientações de uso. Estou ciente dos riscos de uso incorreto.",
            value=False,
            key="massa_chk_confirmar_em_conferir",
        )
        st.markdown("---")

    st.caption(
        "Os valores vêm do cadastro e da tabela colada. "
        "Editáveis: **Altura, Altura lateral, Tipo, Locação**. Os demais campos estão travados. "
        "`Altura lateral` só é editável em estruturas de **3 postes com PR**."
    )

    colunas_edit = ["Sequencial", "Estrutura", "Altura", "Altura lateral",
                    "Tipo", "Config", "Locacao",
                    "dist_f1", "dist_f2", "dist_f3",
                    "dist_f4", "dist_f5", "dist_f6", "dist_pr"]

    df_editable = st.session_state["massa_df_conf_source"][colunas_edit].copy()

    df_editado = st.data_editor(
        df_editable,
        key="massa_editor",
        num_rows="fixed",
        use_container_width=True,
        disabled=["Sequencial", "Estrutura", "Config",
                  "dist_f1", "dist_f2", "dist_f3",
                  "dist_f4", "dist_f5", "dist_f6", "dist_pr"],
        column_config={
            "Altura": st.column_config.NumberColumn("Altura central (m)", min_value=0, max_value=50, step=1),
            "Altura lateral": st.column_config.TextColumn(
                "Altura lateral (m)",
                help="N/A: não se aplica (1 poste) · Cópia da central: 3 postes sem PR · Manual: 3 postes com PR"
            ),
            "Tipo": st.column_config.SelectboxColumn("Tipo", options=["Circular", "Duplo T"]),
            "Locacao": st.column_config.SelectboxColumn(
                "Locação", options=["Topo para bissetriz", "Gaveta para bissetriz", "N/A"]),
            "dist_f1": st.column_config.NumberColumn("d_f1", format="%.2f"),
            "dist_f2": st.column_config.NumberColumn("d_f2", format="%.2f"),
            "dist_f3": st.column_config.NumberColumn("d_f3", format="%.2f"),
            "dist_f4": st.column_config.NumberColumn("d_f4", format="%.2f"),
            "dist_f5": st.column_config.NumberColumn("d_f5", format="%.2f"),
            "dist_f6": st.column_config.NumberColumn("d_f6", format="%.2f"),
            "dist_pr": st.column_config.NumberColumn("d_pr", format="%.2f"),
        },
    )

    df_editado = _aplicar_regra_altura_lateral(df_editado, estruturas_dict)

    estados, motivos = [], []
    for _, row in df_editado.iterrows():
        estado, motivo = recomputar_estado(row, estruturas_dict)
        estados.append(estado)
        motivos.append(motivo)
    df_editado["Estado"] = estados
    df_editado["Motivo"] = motivos
    st.session_state["massa_df_conf_atual"] = df_editado

    st.markdown("**Estados atuais:**")
    st.dataframe(
        df_editado[["Sequencial", "Estrutura", "Altura", "Altura lateral",
                    "Tipo", "Locacao", "Estado", "Motivo"]],
        use_container_width=True, hide_index=True,
    )

    n_calculaveis = sum(1 for e in estados if e == "✅ Calculada")
    st.info(f"Estruturas calculáveis: **{n_calculaveis}** de **{len(estados)}**")

    n_pendente_lateral = sum(1 for e in estados if e == "🔶 Altura lateral pendente")
    if n_pendente_lateral > 0:
        st.warning(
            f"⚠️ **{n_pendente_lateral}** estrutura(s) de **3 postes com PR** "
            "sem altura lateral informada. Preencha a coluna **Altura lateral** "
            "para que elas sejam calculadas. Enquanto pendentes, essas estruturas "
            "não serão calculadas (nem o poste central, nem os laterais)."
        )

    st.markdown("---")
    st.subheader("☑️ 4. Hipóteses a processar (para todas as estruturas)")

    if st.session_state.get("massa_marcar_flag"):
        for lc in TODAS_LCS:
            st.session_state[f"massa_chk_{lc}"] = True
            st.session_state["massa_lcs_marcados"][lc] = True
        st.session_state["massa_marcar_flag"] = False
    if st.session_state.get("massa_desmarcar_flag"):
        for lc in TODAS_LCS:
            st.session_state[f"massa_chk_{lc}"] = False
            st.session_state["massa_lcs_marcados"][lc] = False
        st.session_state["massa_desmarcar_flag"] = False

    cb1, cb2 = st.columns(2)
    with cb1:
        if st.button("✓ Marcar todos", use_container_width=True, key="massa_btn_marcar"):
            st.session_state["massa_marcar_flag"] = True
            st.rerun()
    with cb2:
        if st.button("✗ Desmarcar todos", use_container_width=True, key="massa_btn_desmarcar"):
            st.session_state["massa_desmarcar_flag"] = True
            st.rerun()

    cols = st.columns(4)
    for i, lc in enumerate(TODAS_LCS):
        with cols[i % 4]:
            marcado = st.checkbox(lc, value=st.session_state["massa_lcs_marcados"].get(lc, False),
                                   key=f"massa_chk_{lc}")
            st.session_state["massa_lcs_marcados"][lc] = marcado

    st.markdown("---")
    st.subheader("⚙️ 5. Configurações globais")

    c1, c2, c3 = st.columns(3)
    with c1:
        st.session_state["massa_fator_cagaco"] = st.number_input(
            "Fator cagaço (%)", min_value=0, max_value=20,
            value=st.session_state["massa_fator_cagaco"], step=1, key="massa_cagaco")
    with c2:
        st.session_state["massa_pressao_vento"] = st.number_input(
            "Pressão de vento (kgf/m²)", min_value=0.0, max_value=200.0,
            value=st.session_state["massa_pressao_vento"], step=1.0, key="massa_vento")
    with c3:
        st.session_state["massa_considerar_vento"] = st.checkbox(
            "Considerar vento no fuste (VENT MAX) — INATIVO",
            value=st.session_state["massa_considerar_vento"], key="massa_chk_vento")

    st.markdown("---")
    btn1, btn2 = st.columns(2)
    with btn1:
        algum_marcado = any(st.session_state["massa_lcs_marcados"].values())
        confirmou = st.session_state.get("massa_confirmou_orientacoes", False)
        pode_executar = algum_marcado and n_calculaveis > 0 and confirmou
        if st.button("🚀 Executar cálculo em massa", type="primary",
                     use_container_width=True, disabled=(not pode_executar)):
            _executar_calculo_massa(df_editado, estruturas_dict)
            st.rerun()
        if not confirmou:
            st.caption("⚠️ Marque a confirmação de leitura das orientações para liberar.")
    with btn2:
        if st.button("🔄 Limpar tudo", use_container_width=True, key="massa_btn_reset"):
            reset_massa()
            st.rerun()

def _executar_calculo_massa(df_conf, estruturas_dict):
    lcs_sel = [lc for lc, v in st.session_state["massa_lcs_marcados"].items() if v]
    df_parse = st.session_state["massa_df_parse"]
    fator_cagaco = st.session_state["massa_fator_cagaco"]
    pressao_vento = st.session_state["massa_pressao_vento"]
    considerar_vento = st.session_state["massa_considerar_vento"]

    resultados = []
    for _, row in df_conf.iterrows():
        seq = row["Sequencial"]
        nome = row["Estrutura"]
        altura = float(row["Altura"])
        tipo = row["Tipo"]
        config = row["Config"]
        locacao = row["Locacao"]

        estado, motivo = recomputar_estado(row, estruturas_dict)

        cadastro = estruturas_dict.get(nome)
        num_postes = cadastro.get("num_postes", 1) if cadastro else 1
        tem_pr = cadastro.get("tem_pr", False) if cadastro else False

        if estado != "✅ Calculada":
            resultados.append({
                "Sequencial": seq, "Estrutura": nome, "Altura": altura,
                "Tipo": tipo, "Config": config,
                "Estado": estado,
                "Poste lateral": "N/A" if num_postes == 1 else "",
                "Poste central": "",
                "Poste consolidado": "",
                "LC crítico": "", "R_final (daN)": "",
                "Ultrapassou 4000": "—",
                "Motivo": motivo,
            })
            continue

        if tipo == "Circular":
            theta_face = None
        elif locacao == "Topo para bissetriz":
            theta_face = 90.0
        else:
            theta_face = 0.0

        estrutura_calc = {
            "tipo_poste": tipo,
            "tem_pr": tem_pr,
            "dist_f1": row["dist_f1"], "dist_f2": row["dist_f2"], "dist_f3": row["dist_f3"],
            "dist_f4": row["dist_f4"], "dist_f5": row["dist_f5"], "dist_f6": row["dist_f6"],
            "dist_pr": row["dist_pr"], "dist_pr2": None,
        }

        altura_lateral = altura
        if num_postes == 3:
            al_raw = row.get("Altura lateral", "")
            if al_raw is not None and str(al_raw).strip() not in ("", "N/A"):
                try:
                    altura_lateral = float(str(al_raw).replace(",", "."))
                except (ValueError, TypeError):
                    altura_lateral = altura

        df_grupo = df_parse[df_parse["Sequencial"] == seq]

        resultados_lc = []
        for lc_nome in lcs_sel:
            df_lc = df_grupo[df_grupo["LC"].astype(str).str.strip() == lc_nome]
            if df_lc.empty:
                continue

            if num_postes == 3:
                res = calcular_3_postes(
                    df_lc, lc_nome, estrutura_calc,
                    altura, altura_lateral,
                    fator_cagaco, theta_face,
                    considerar_vento, pressao_vento,
                )
                if res is not None:
                    resultados_lc.append({
                        "tipo": "3p",
                        "res": res,
                        "R_final_crit": res["R_final_max"],
                        "estourou": res["estourou"],
                        "poste_lat": res["poste_lateral"],
                        "poste_cen": res["poste_central"],
                        "poste_cons": res["poste_str"],
                    })
            else:
                res = processar_lc(lc_nome, df_lc, estrutura_calc, altura,
                                   fator_cagaco, theta_face, considerar_vento, pressao_vento)
                if res is not None:
                    resultados_lc.append({
                        "tipo": "1p",
                        "res": res,
                        "R_final_crit": res["R_final"],
                        "estourou": res["estourou"],
                        "poste_lat": "N/A",
                        "poste_cen": res["poste"],
                        "poste_cons": str(res["poste"]),
                    })

        if not resultados_lc:
            resultados.append({
                "Sequencial": seq, "Estrutura": nome, "Altura": altura,
                "Tipo": tipo, "Config": config,
                "Estado": "⚠️ Sem resultados",
                "Poste lateral": "N/A" if num_postes == 1 else "",
                "Poste central": "",
                "Poste consolidado": "",
                "LC crítico": "", "R_final (daN)": "",
                "Ultrapassou 4000": "—",
                "Motivo": "Nenhum LC processado",
            })
            continue

        crit = max(resultados_lc, key=lambda x: x["R_final_crit"])
        estourou = any(r["estourou"] for r in resultados_lc)

        resultados.append({
            "Sequencial": seq, "Estrutura": nome, "Altura": altura,
            "Tipo": tipo, "Config": config,
            "Estado": "✅ Calculada",
            "Poste lateral": crit["poste_lat"],
            "Poste central": crit["poste_cen"],
            "Poste consolidado": crit["poste_cons"],
            "LC crítico": crit["res"]["LC"],
            "R_final (daN)": f"{crit['R_final_crit']:.1f}",
            "Ultrapassou 4000": "Sim" if estourou else "Não",
            "Motivo": "",
        })

    st.session_state["massa_resultado_df"] = pd.DataFrame(resultados)
    st.session_state["massa_etapa"] = "resultado"

def _massa_etapa_resultado():
    st.subheader("📊 Resultados do dimensionamento em massa")

    df = st.session_state.get("massa_resultado_df")
    if df is None or df.empty:
        st.warning("Sem resultados.")
        if st.button("Voltar"):
            st.session_state["massa_etapa"] = "conferir"
            st.rerun()
        return

    n_total = len(df)
    n_calc = int((df["Estado"] == "✅ Calculada").sum())
    n_nc = n_total - n_calc
    calculadas = df[df["Estado"] == "✅ Calculada"]
    poste_max = calculadas["Poste central"].max() if not calculadas.empty else "—"

    c1, c2, c3, c4 = st.columns(4)
    with c1: st.metric("Total", n_total)
    with c2: st.metric("Calculadas", n_calc)
    with c3: st.metric("Não calculadas", n_nc)
    with c4: st.metric("Poste máx (daN)", poste_max if poste_max != "—" else "—")

    if "Ultrapassou 4000" in df.columns:
        n_estourou = int((df["Ultrapassou 4000"] == "Sim").sum())
        if n_estourou > 0:
            st.warning(
                f"⚠️ **{n_estourou}** estrutura(s) ultrapassou(aram) **4000 daN**. "
                "Poste comercial pode ser insuficiente para essas estruturas."
            )

    st.markdown("---")
    st.dataframe(df, use_container_width=True, hide_index=True)

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Dimensionamento")
    st.download_button(
        "📥 Exportar XLSX",
        buffer.getvalue(),
        "dimensionamento_massa.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.markdown("---")
    if st.button("🔄 Nova análise em massa", use_container_width=True):
        reset_massa()
        st.rerun()

# ============================================================
# MAIN
# ============================================================
inicializar_estado()

estruturas_all = carregar_estruturas()
if not estruturas_all:
    st.error(
        "Nenhuma estrutura carregada. Verifique se `estruturas.xlsx` está na "
        "mesma pasta do `app.py`."
    )
    st.stop()

estruturas_liberadas = [e for e in estruturas_all if e["liberacao"] == "SIM"]
estruturas_dict = {e["nome"]: e for e in estruturas_all}
st.session_state["_estruturas_dict"] = estruturas_dict

tab_ind, tab_massa = st.tabs(["🔍 Teste individual", "📊 Dimensionamento em massa"])

with tab_ind:
    render_individual_tab(estruturas_liberadas)

with tab_massa:
    render_massa_tab(estruturas_dict)
