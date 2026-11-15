# ============================================================
# VHPT 2.0 - Verificador de Hipóteses de Carregamento Mecânico
#            para Postes (baseado em PLS-CADD)
# ============================================================

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.patches import Ellipse
from io import StringIO
import os

# ============================================================
# CONFIGURAÇÃO DA PÁGINA
# ============================================================
st.set_page_config(
    page_title="VHPT 2.0 - Verificador de Postes",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.title("🏗️ VHPT 2.0 - Verificador de Hipóteses de Carregamento Mecânico")
st.markdown("---")

# ============================================================
# CONSTANTES GLOBAIS
# ============================================================
COMERCIAIS = [1500, 2000, 2500, 3000, 4000]
SOBRECARGA_TRANSITORIA = 0.40
ALTURA_REFERENCIA = 0.20

ALTURAS_DISPONIVEIS = [18, 20, 22, 24, 26, 28, 30, 32]
ALTURA_DEFAULT = 22

# Mapeamento Set/Phase -> campo de altura no cadastro
MAPEAMENTO_SET_PHASE = {
    (1, 1): 'dist_f1',
    (1, 2): 'dist_f2',
    (1, 3): 'dist_f3',
    (2, 1): 'dist_f4',
    (2, 2): 'dist_f5',
    (2, 3): 'dist_f6',
    (4, 1): 'dist_pr',
    (5, 1): 'dist_pr2',   # reservado para futuro (não existe no Excel hoje)
}

# Sets requeridos por configuração de estrutura
SETS_REQUERIDOS = {
    'CS_sem_PR': {1},
    'CS_com_PR': {1, 4},
    'CD_sem_PR': {1, 2},
    'CD_com_PR': {1, 2, 4},
}

# Lista de LCs disponíveis (dependência de set)
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

# ============================================================
# CADASTRO DE ESTRUTURAS (mesma leitura do VHPT 1.0)
# ============================================================
@st.cache_data
def carregar_estruturas():
    if not os.path.exists("estruturas.xlsx"):
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

    try:
        df = pd.read_excel("estruturas.xlsx", engine="openpyxl")
        estruturas = []
        for _, row in df.iterrows():
            liberacao = str(row.iloc[1]).strip().upper() if pd.notna(row.iloc[1]) else ""
            if liberacao != "SIM":
                continue

            poste_cod    = str(row.iloc[2]).strip().upper() if pd.notna(row.iloc[2]) else ""
            circuito_cod = str(row.iloc[4]).strip().upper() if pd.notna(row.iloc[4]) else ""
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
                "tipo_poste": tipo_poste,
                "is_duplo": is_duplo,
                "tem_pr": tem_pr,
                "config": config,
                "dist_pr":  _float_ou_padrao(row.iloc[12] if len(row) > 12 else None, 0.20),
                "dist_f1":  _float_ou_padrao(row.iloc[13] if len(row) > 13 else None, 0.20),
                "dist_f2":  _float_ou_padrao(row.iloc[14] if len(row) > 14 else None, 2.20),
                "dist_f3":  _float_ou_padrao(row.iloc[15] if len(row) > 15 else None, 4.20),
                "dist_f4":  _float_ou_padrao(row.iloc[16] if len(row) > 16 else None, 0.20),
                "dist_f5":  _float_ou_padrao(row.iloc[17] if len(row) > 17 else None, 2.20),
                "dist_f6":  _float_ou_padrao(row.iloc[18] if len(row) > 18 else None, 4.20),
                "dist_pr2": None,  # Reservado para PR2 (futuro)
            }
            if estrutura["nome"]:
                estruturas.append(estrutura)
        return estruturas
    except Exception:
        return []

# ============================================================
# PARSER DA TABELA COLADA
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
        df = pd.read_csv(
            StringIO(texto),
            sep="\t",
            header=None,
            dtype=str,
            keep_default_na=False
        )
    except Exception as e:
        return None, f"Falha ao ler tabela: {e}"

    if df.shape[1] < 14:
        return None, f"Esperadas 14 colunas, encontradas {df.shape[1]}."

    df = df.iloc[:, :14].copy()
    df.columns = COLUNAS_INTERNAS

    primeira = str(df.iloc[0, 0]).strip().upper()
    if "LOAD" in primeira and "CASE" in primeira:
        df = df.iloc[1:].reset_index(drop=True)

    colunas_num = ["Set", "Phase",
                   "S_Vert", "S_Trans", "S_Long",
                   "B_Vert", "B_Trans", "B_Long",
                   "A_Vert", "A_Trans", "A_Long"]
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
    return set(df["Set"].unique().tolist())

def listar_lcs_disponiveis(sets_presentes):
    disponiveis = []
    for lc in LCS_DISPONIVEIS:
        if lc["sets"].issubset(sets_presentes):
            disponiveis.append(lc["nome"])
    return disponiveis

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
    Y = altura_aerea - dist_cabo
    return Y / X

def fator_correcao_duplo_t(theta_res, theta_face):
    delta = abs(theta_res - theta_face)
    if delta > 90:
        delta = 180 - delta
    rad = np.radians(delta)
    a, b = 1.0, 0.5
    return 1.0 / np.sqrt((np.cos(rad) / a) ** 2 + (np.sin(rad) / b) ** 2)

def calcular_forca_vento_poste(tipo_poste, altura, pressao_vento):
    """INATIVA no VHPT 2.0 (preparada para uso futuro na VENT MAX)."""
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
        chave = (int(row["Set"]), int(row["Phase"]))
        campo_alt = MAPEAMENTO_SET_PHASE.get(chave)
        if campo_alt is None:
            continue
        dist_cabo = estrutura.get(campo_alt)
        if dist_cabo is None:
            continue
        K = calcular_K(altura_aerea, dist_cabo)
        if K is None:
            continue

        T_cabo = float(row["B_Trans"]) + float(row["A_Trans"])
        L_cabo = float(row["B_Long"])  + float(row["A_Long"])
        V_cabo = float(row["B_Vert"])  + float(row["A_Vert"])

        resultados_por_cabo[chave] = {
            "K": K,
            "T": T_cabo,
            "L": L_cabo,
            "V": V_cabo,
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
        "LC": nome_lc,
        "V_total": V_total,
        "T_back": T_back, "T_ahead": T_ahead,
        "L_back": L_back, "L_ahead": L_ahead,
        "F_T": F_T, "F_L": F_L,
        "R": R, "theta_R": theta_R,
        "theta_face": theta_face,
        "fator_DT": fator_DT,
        "R_corr": R_corr,
        "divisor": divisor,
        "R_final": R_final,
        "poste": poste,
        "estourou": estourou,
        "F_vento_poste": F_vento_poste,
        "cabos": resultados_por_cabo,
    }

# ============================================================
# VISUALIZAÇÃO
# ============================================================
def plotar_diagrama(resultado, estrutura):
    fig, ax = plt.subplots(figsize=(8, 6))

    F_T = resultado["F_T"]
    F_L = resultado["F_L"]
    R = resultado["R"]
    theta_R = resultado["theta_R"]

    max_val = max(abs(F_T), abs(F_L), abs(R), 1e-6)
    scale = max_val / 1.2

    ax.axhline(y=0, color='k', linewidth=0.5, alpha=0.3)
    ax.axvline(x=0, color='k', linewidth=0.5, alpha=0.3)

    ax.arrow(0, 0, F_L / scale, 0,
             head_width=0.05, head_length=0.05,
             fc='blue', ec='blue', linewidth=1.5,
             label=f'Longitudinal: {F_L:.0f} daN')
    ax.arrow(0, 0, 0, F_T / scale,
             head_width=0.05, head_length=0.05,
             fc='red', ec='red', linewidth=1.5,
             label=f'Transversal: {F_T:.0f} daN')

    x_res = (R * np.cos(np.radians(theta_R))) / scale
    y_res = (R * np.sin(np.radians(theta_R))) / scale
    ax.arrow(0, 0, x_res, y_res,
             head_width=0.07, head_length=0.07,
             fc='green', ec='green', linewidth=2.5,
             label=f'Resultante: {R:.0f} daN')

    if estrutura["tipo_poste"] == "Duplo T" and resultado["theta_face"] is not None:
        tf = resultado["theta_face"]
        tg = tf + 90

        x_l = 1.2 * np.cos(np.radians(tf))
        y_l = 1.2 * np.sin(np.radians(tf))
        ax.plot([-x_l, x_l], [-y_l, y_l], 'purple', linewidth=1.5,
                linestyle='--', alpha=0.8, label='Face Lisa')

        x_g = 1.2 * np.cos(np.radians(tg))
        y_g = 1.2 * np.sin(np.radians(tg))
        ax.plot([-x_g, x_g], [-y_g, y_g], 'orange', linewidth=1.5,
                linestyle='--', alpha=0.8, label='Gaveta')

        a = 1.0 * (max_val / scale) * 0.6
        b = 0.5 * (max_val / scale) * 0.6
        el = Ellipse((0, 0), 2 * a, 2 * b, angle=tf,
                     facecolor='none', edgecolor='gray',
                     linestyle=':', linewidth=1.5, alpha=0.7)
        ax.add_patch(el)

    ax.set_xlim(-1.35, 1.35)
    ax.set_ylim(-1.35, 1.35)
    ax.set_aspect('equal')
    ax.grid(True, alpha=0.2)
    ax.set_xlabel('Longitudinal (X)', fontsize=10)
    ax.set_ylabel('Transversal (Y)', fontsize=10)
    ax.set_title(f'{resultado["LC"]} — R_final = {resultado["R_final"]:.1f} daN',
                 fontsize=11, fontweight='bold')
    ax.legend(loc='upper left', bbox_to_anchor=(1.02, 1), fontsize=8)

    ax.text(1.05, 0, '0°', fontsize=7, alpha=0.7)
    ax.text(0, 1.05, '90°', fontsize=7, alpha=0.7)

    plt.tight_layout()
    return fig

# ============================================================
# ESTADO DO STREAMLIT
# ============================================================
def inicializar_estado():
    defaults = {
        "texto_colado": "",
        "estrutura_nome": "",
        "altura": ALTURA_DEFAULT,
        "fator_cagaco": 0,
        "pressao_vento": 30.0,
        "considerar_vento": False,
        "topo_gaveta": "Topo para bissetriz",
        "lcs_marcados": {},
        "resultados": None,
        "lc_critico": None,
        "lc_grafico": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

def reset_geral():
    for k in list(st.session_state.keys()):
        del st.session_state[k]
    inicializar_estado()

def limpar_tabela():
    st.session_state["texto_colado"] = ""
    st.session_state["resultados"] = None
    st.session_state["lc_critico"] = None
    st.session_state["lc_grafico"] = None

# ============================================================
# INTERFACE
# ============================================================
inicializar_estado()
estruturas = carregar_estruturas()

st.subheader("🏗️ 1. Estrutura")

if not estruturas:
    st.error("Nenhuma estrutura carregada. Verifique o arquivo estruturas.xlsx.")
    st.stop()

nomes = [e["nome"] for e in estruturas]

estrutura_nome = st.selectbox(
    "Selecione a estrutura:",
    [""] + nomes,
    index=(nomes.index(st.session_state["estrutura_nome"]) + 1)
        if st.session_state["estrutura_nome"] in nomes else 0
)
st.session_state["estrutura_nome"] = estrutura_nome

estrutura = None
if estrutura_nome:
    estrutura = next(e for e in estruturas if e["nome"] == estrutura_nome)
    sets_req = SETS_REQUERIDOS[estrutura["config"]]
    st.info(
        f"**{estrutura['tipo_poste']}** — "
        f"Configuração: **{estrutura['config']}** — "
        f"Sets requeridos na tabela: **{sorted(sets_req)}**"
    )

st.markdown("---")
st.subheader("📋 2. Cole a tabela do PLS-CADD")

if estrutura is not None:
    st.caption(
        "Cole apenas os VALORES (sem a linha de cabeçalho). "
        "Ordem esperada: LC | Set | Phase | Attach | S_Vert | S_Trans | S_Long | "
        "B_Vert | B_Trans | B_Long | A_Vert | A_Trans | A_Long | Warnings"
    )

    texto_colado = st.text_area(
        "Cole aqui (Ctrl+V do Excel):",
        value=st.session_state["texto_colado"],
        height=200,
        key="textarea_tabela"
    )
    if texto_colado != st.session_state["texto_colado"]:
        st.session_state["texto_colado"] = texto_colado
        st.session_state["resultados"] = None

    df = None
    sets_presentes = set()
    lcs_disponiveis = []

    if st.session_state["texto_colado"].strip():
        df, erro = parsear_tabela(st.session_state["texto_colado"])
        if erro:
            st.error(f"❌ {erro}")
        else:
            sets_presentes = detectar_sets(df)
            sets_req = SETS_REQUERIDOS[estrutura["config"]]
            if sets_presentes != sets_req:
                faltam = sets_req - sets_presentes
                sobram = sets_presentes - sets_req
                msg = "❌ Incompatibilidade de sets."
                if faltam: msg += f"  Faltam: {sorted(faltam)}."
                if sobram: msg += f"  Sobram: {sorted(sobram)}."
                st.error(msg)
                df = None
            else:
                st.success(
                    f"✅ Tabela válida — {len(df)} linhas, "
                    f"sets {sorted(sets_presentes)}"
                )
                lcs_disponiveis = listar_lcs_disponiveis(sets_presentes)

    if df is not None and lcs_disponiveis:
        st.markdown("---")
        st.subheader("☑️ 3. Hipóteses a processar")
        st.caption("Todas desmarcadas por default. Marque as que deseja calcular.")

        cols = st.columns(4)
        for i, lc in enumerate(lcs_disponiveis):
            with cols[i % 4]:
                marcado = st.checkbox(
                    lc,
                    value=st.session_state["lcs_marcados"].get(lc, False),
                    key=f"chk_{lc}"
                )
                st.session_state["lcs_marcados"][lc] = marcado

        st.markdown("---")
        st.subheader("⚙️ 4. Configurações")

        c1, c2, c3 = st.columns(3)
        with c1:
            altura = st.selectbox(
                "Altura do poste (m)",
                ALTURAS_DISPONIVEIS,
                index=ALTURAS_DISPONIVEIS.index(st.session_state["altura"])
                    if st.session_state["altura"] in ALTURAS_DISPONIVEIS else 0,
                key="sel_altura"
            )
            st.session_state["altura"] = altura

        with c2:
            fator_cagaco = st.number_input(
                "Fator cagaço (%)",
                min_value=0, max_value=20,
                value=st.session_state["fator_cagaco"], step=1,
                key="inp_cagaco"
            )
            st.session_state["fator_cagaco"] = fator_cagaco

        with c3:
            pressao_vento = st.number_input(
                "Pressão de vento (kgf/m²)",
                min_value=0.0, max_value=200.0,
                value=st.session_state["pressao_vento"], step=1.0,
                key="inp_vento"
            )
            st.session_state["pressao_vento"] = pressao_vento

        considerar_vento = st.checkbox(
            "Considerar vento no fuste (apenas VENT MAX) — INATIVO por enquanto",
            value=st.session_state["considerar_vento"],
            key="chk_vento"
        )
        st.session_state["considerar_vento"] = considerar_vento

        theta_face = None
        if estrutura["tipo_poste"] == "Duplo T":
            topo_gaveta = st.selectbox(
                "Montagem Duplo T",
                ["Topo para bissetriz", "Gaveta para bissetriz"],
                index=0 if st.session_state["topo_gaveta"] == "Topo para bissetriz" else 1,
                key="sel_topo"
            )
            st.session_state["topo_gaveta"] = topo_gaveta
            theta_face = 90.0 if topo_gaveta == "Topo para bissetriz" else 0.0

        st.markdown("---")
        cbtn1, cbtn2, cbtn3 = st.columns(3)
        with cbtn1:
            botao_calcular = st.button(
                "🔍 Calcular",
                type="primary",
                use_container_width=True,
                disabled=(not any(st.session_state["lcs_marcados"].values()))
            )
            if botao_calcular:
                lcs_sel = [lc for lc, v in st.session_state["lcs_marcados"].items() if v]
                resultados = []
                for lc_nome in lcs_sel:
                    df_lc = df[df["LC"] == lc_nome]
                    if df_lc.empty:
                        continue
                    res = processar_lc(
                        lc_nome, df_lc, estrutura,
                        st.session_state["altura"],
                        st.session_state["fator_cagaco"],
                        theta_face,
                        st.session_state["considerar_vento"],
                        st.session_state["pressao_vento"]
                    )
                    if res is not None:
                        resultados.append(res)
                st.session_state["resultados"] = resultados
                if resultados:
                    crit = max(resultados, key=lambda x: x["R_final"])
                    st.session_state["lc_critico"] = crit["LC"]
                    st.session_state["lc_grafico"] = crit["LC"]

        with cbtn2:
            if st.button("🧹 Limpar tabela e resultados", use_container_width=True):
                limpar_tabela()
                st.rerun()

        with cbtn3:
            if st.button("🔄 Reset geral", use_container_width=True):
                reset_geral()
                st.rerun()

# ============================================================
# RESULTADOS
# ============================================================
if st.session_state["resultados"]:
    st.markdown("---")
    st.subheader("📊 Resultados")

    resultados = st.session_state["resultados"]

    linhas = []
    for r in resultados:
        linhas.append({
            "LC": r["LC"],
            "V_total": f"{r['V_total']:.1f}",
            "T_back": f"{r['T_back']:.1f}",
            "T_ahead": f"{r['T_ahead']:.1f}",
            "L_back": f"{r['L_back']:.1f}",
            "L_ahead": f"{r['L_ahead']:.1f}",
            "F_T (transf.)": f"{r['F_T']:.1f}",
            "F_L (transf.)": f"{r['F_L']:.1f}",
            "R (daN)": f"{r['R']:.1f}",
            "θ_R (°)": f"{r['theta_R']:.1f}",
            "θ_face (°)": f"{r['theta_face']:.1f}" if r['theta_face'] is not None else "—",
            "Fator DT": f"{r['fator_DT']:.3f}",
            "R_corr (daN)": f"{r['R_corr']:.1f}",
            "Divisor": f"{r['divisor']:.2f}",
            "R_final (daN)": f"{r['R_final']:.1f}",
            "Poste": r["poste"],
        })
    st.dataframe(pd.DataFrame(linhas), use_container_width=True, hide_index=True)

    poste_final = max(r["poste"] for r in resultados)
    crit = next(r for r in resultados if r["LC"] == st.session_state["lc_critico"])
    st.markdown("### 🎯 Resumo")
    cA, cB, cC = st.columns(3)
    with cA:
        st.metric("Poste recomendado", f"{poste_final} daN")
    with cB:
        st.metric("LC crítico", crit["LC"])
    with cC:
        st.metric("R_final do crítico", f"{crit['R_final']:.1f} daN")

    if any(r["estourou"] for r in resultados):
        st.warning("⚠️ Pelo menos um LC ultrapassou 4000 daN. Poste comercial pode ser insuficiente.")

    st.markdown("### 📈 Diagrama vetorial")
    opcoes = [r["LC"] for r in resultados]
    idx_graf = opcoes.index(st.session_state["lc_grafico"]) if st.session_state["lc_grafico"] in opcoes else 0
    lc_graf = st.selectbox("LC para visualizar:", opcoes, index=idx_graf)
    st.session_state["lc_grafico"] = lc_graf
    res_graf = next(r for r in resultados if r["LC"] == lc_graf)
    fig = plotar_diagrama(res_graf, estrutura)
    st.pyplot(fig)
    plt.close(fig)

    with st.expander("ℹ️ Detalhes do cálculo (LC crítico)"):
        st.write(f"**Altura:** {st.session_state['altura']} m")
        st.write(f"**Fator cagaço:** {st.session_state['fator_cagaco']}%")
        st.write(f"**Pressão de vento:** {st.session_state['pressao_vento']} kgf/m²")
        st.write(f"**Considerar vento no fuste:** {'Sim' if st.session_state['considerar_vento'] else 'Não (inativo)'}")
        st.write(f"**Montagem DT:** {st.session_state['topo_gaveta'] if estrutura['tipo_poste'] == 'Duplo T' else 'N/A (Circular)'}")
        st.write("**Coeficientes K por cabo:**")
        for chave, dados in crit["cabos"].items():
            st.text(f"  Set {chave[0]} Phase {chave[1]}: K = {dados['K']:.3f}, T = {dados['T']:.1f}, L = {dados['L']:.1f}")
