import streamlit as st
import pandas as pd
import numpy as np
import io
import math
import matplotlib.pyplot as plt
import datetime
import base64
import os

st.set_page_config(page_title="💻WENSLO + ALWAS Tool", layout="wide")

# --------------------------
# Inicializa estado de sessão
# --------------------------
if 'results_ready' not in st.session_state:
    st.session_state['results_ready'] = False
    st.session_state['wens'] = None
    st.session_state['alwas_res'] = None
    st.session_state['graduation_df'] = None
    st.session_state['classes_dict'] = None
    st.session_state['validation'] = None
    st.session_state['std_mat'] = None
    st.session_state['weights'] = None
    st.session_state['excel_bytes'] = None
    st.session_state['corr_matrix'] = None
    st.session_state['run_id'] = 0  # para gerar keys únicas por execução

# -----------------------
# Config / Logo path
# -----------------------
LOGO_PATH = "UFF_EN_brasao.png"

INSTITUTION_LINE = (
    "Universidade Federal Fluminense – Programa de Pós-Graduação em Engenharia de Produção<br/>"
     "Escola Naval"
)

# -----------------------
# Bilingual text — padrão centralizado
# -----------------------
TEXT = {
    # Geral / UI
    "app_title":            {"Português": "🤖 Ferramenta WENSLO + ALWAS — UFF/EN",
                             "English":   "🤖 WENSLO + ALWAS Tool — UFF/EN"},
    "upload_header":        {"Português": "Carregar matriz de decisão",
                             "English":   "Load decision matrix"},
    "download_template":    {"Português": "Baixar template Excel",
                             "English":   "Download Excel template"},
    "upload_prompt":        {"Português": "Carregue o arquivo (xlsx)",
                             "English":   "Upload your xlsx file"},
    "run":                  {"Português": "Calcular",
                             "English":   "Run"},
    "clear_results":        {"Português": "Limpar resultados",
                             "English":   "Clear results"},
    "use_g7":               {"Português": "Usar dados do G7 (validação)",
                             "English":   "Use G7 data (validation)"},
    "data_loaded":          {"Português": "Dados carregados (confira se estão corretos)",
                             "English":   "Loaded data (check correctness)"},
    "calculating":          {"Português": "Calculando...",
                             "English":   "Calculating..."},
    "calc_done":            {"Português": "Cálculo finalizado!",
                             "English":   "Calculation complete!"},
    "load_template_first":  {"Português": "Carregue o template preenchido ou use os dados G7 para teste.",
                             "English":   "Upload the filled template or use G7 data for testing."},
    "file_read_error":      {"Português": "Erro ao ler o arquivo: ",
                             "English":   "Error reading file: "},
    "second_row_error":     {"Português": "Segunda linha deve conter apenas 'MAX' ou 'MIN' para cada critério.",
                             "English":   "Second row must contain only 'MAX' or 'MIN' for each criterion."},

    # Resultados / subheaders
    "graduation_table":           {"Português": "Matriz de Decisão Normalizada",
                                   "English":   "Normalized Decision-Making Matrix"},
    "graduation_summary_header":  {"Português": "Resumo da graduação por critério",
                                   "English":   "Graduation summary by criterion"},
    "col_n_classes":              {"Português": "Nº de classes",
                                   "English":   "Number of classes"},
    "col_min_z":                  {"Português": "z (mín)",
                                   "English":   "z (min)"},
    "col_max_z":                  {"Português": "z (máx)",
                                   "English":   "z (max)"},
    "weights":                    {"Português": "Pesos (WENSLO)",
                                   "English":   "Weights (WENSLO)"},
    "norm_matrix_z":              {"Português": "Matriz normalizada Z",
                                   "English":   "Normalized matrix Z"},
    "delta_z_sturges":            {"Português": "Δz (Sturges) por critério",
                                   "English":   "Δz (Sturges) per criterion"},
    "envelope_tanphi":            {"Português": "Envelope E e tan_phi",
                                   "English":   "Envelope E and tan_phi"},
    "alwas_std_matrix":           {"Português": "ALWAS: matriz doméstica padronizada (ζ̂)",
                                   "English":   "ALWAS: standardized home matrix (ζ̂)"},
    "ranking":                    {"Português": "Ranking (ALWAS)",
                                   "English":   "Ranking (ALWAS)"},
    "validation":                 {"Português": "Validação da acumulação (MSE / Correlação)",
                                   "English":   "Accumulation validation (MSE / Correlation)"},
    "download_results":           {"Português": "Baixar todos os resultados (Excel)",
                                   "English":   "Download all results (Excel)"},

    # Heatmap / correlação
    "corr_method":          {"Português": "Método de correlação",
                             "English":   "Correlation method"},
    "show_corr":            {"Português": "Mostrar heatmap de correlação",
                             "English":   "Show correlation heatmap"},
    "heatmap_title":        {"Português": "Heatmap — Correlação entre critérios",
                             "English":   "Heatmap — Correlation between criteria"},
    "decimal_places":       {"Português": "Casas decimais",
                             "English":   "Decimal places"},
    "download_heatmap":     {"Português": "Baixar heatmap PNG",
                             "English":   "Download heatmap PNG"},
    "corr_error":           {"Português": "Erro ao calcular correlação: ",
                             "English":   "Error computing correlation: "},
    "corr_unavailable":     {"Português": "Não foi possível calcular a matriz de correlação — verifique os dados.",
                             "English":   "Could not compute correlation matrix — check your data."},
    "corr_no_results":      {"Português": "Resultados ainda não gerados. Clique em 'Calcular' para habilitar o heatmap de correlação.",
                             "English":   "Results not yet generated. Click 'Run' to enable the correlation heatmap."},

    # Acumulação
    "accum_title":          {"Português": "Acumulação: real vs artificial (por critério)",
                             "English":   "Accumulation: real vs artificial (by criterion)"},
    "select_criterion":     {"Português": "Selecione um critério para plotar acumulação",
                             "English":   "Select a criterion to plot accumulation"},
    "accum_xlabel":         {"Português": "Alternativas (na ordem do arquivo)",
                             "English":   "Alternatives (ordered as in input)"},
    "accum_ylabel":         {"Português": "Valor normalizado acumulado",
                             "English":   "Accumulated normalized value"},
    "accum_real":           {"Português": "Acumulação real",
                             "English":   "Real accumulation"},
    "accum_artificial":     {"Português": "Artificial (hipotenusa)",
                             "English":   "Artificial (hypotenuse)"},
    "accum_no_results":     {"Português": "Ainda não há resultados. Clique em 'Calcular' para gerar as tabelas e habilitar os gráficos interativos.",
                             "English":   "No results yet. Click 'Run' to generate tables and enable interactive charts."},

    # Sensibilidade
    "sensitivity":          {"Português": "Análise de sensibilidade (ξ, φ, θ)",
                             "English":   "Sensitivity analysis (ξ, φ, θ)"},

    # Sidebar
    "article_label":        {"Português": "Artigo",
                             "English":   "Article"},
    "developers_label":     {"Português": "Desenvolvedores",
                             "English":   "Developers"},
    "manual_download":      {"Português": "📘 Baixar manual",
                             "English":   "📘 Download manual"},
    "manual_not_found":     {"Português": "Manual PDF não encontrado.",
                             "English":   "PDF manual not found."},
    "citation_title":       {"Português": "📚 Como citar",
                             "English":   "📚 How to cite"},
    "openai_key_label":     {"Português": "Idioma / Language",
                             "English":   "Idioma / Language"},
}

# Função de tradução
def t(key, lang):
    return TEXT.get(key, {}).get(lang, key)

# -----------------------
# Citações — seguindo padrão app_streamlit5
# -----------------------
CITATIONS = {
    "ABNT": """
**Formato ABNT (NBR 6023):**

**Software:**
SANTOS, Marcos dos; GOMES, Carlos Francisco Simões. **Wenslo-Alwas Tool**. Titular: Anderson Gonçalves Portella. 2025. Programa de Computador. Registro INPI: BR512025005226-0. Disponível em: <https://wenslo-alwas-tool.streamlit.app/>. Acesso em: {date}.

**Artigo:**
SILVA, C. S.; SANTOS, M. R. Análise do nível de maturidade em gestão de riscos: um estudo de caso em uma empresa do setor elétrico. In: CONGRESSO NACIONAL DE EXCELÊNCIA EM GESTÃO, 19., 2025, Online. **Anais...** Rio de Janeiro: CNEG, 2025. DOI: 10.14488/cneg2025_cneg_pt_068_0567_23581. Acesso em: {date}.
""",
    "APA": """
**APA Format (7th Ed.):**

**Software:**
Santos, M. dos, & Gomes, C. F. S. (2025). *Wenslo-Alwas Tool* [Computer software]. Anderson Gonçalves Portella. https://wenslo-alwas-tool.streamlit.app/

**Article:**
Silva, C. S., & Santos, M. R. (2025). Análise do nível de maturidade em gestão de riscos: um estudo de caso em uma empresa do setor elétrico. In *Anais do XIX Congresso Nacional de Excelência em Gestão*. DOI: 10.14488/cneg2025_cneg_pt_068_0567_23581
"""
}

def get_citation(lang):
    citation_type = "ABNT" if lang == "Português" else "APA"
    today = datetime.datetime.now().strftime("%d/%m/%Y" if lang == "Português" else "%B %d, %Y")
    return CITATIONS[citation_type].format(date=today)

# -----------------------
# Sidebar
# -----------------------
with st.sidebar:
    lang = st.selectbox("Idioma / Language", options=["Português", "English"], index=0)
    is_pt = (lang == "Português")

    article_label = t("article_label", lang)
    dev_label = t("developers_label", lang)
    st.markdown(
        f"""
        **{article_label}**

        - [*A Novel WENSLO and ALWAS Multicriteria Methodology and Its Application to Green Growth Performance Evaluation*](https://doi.org/10.1109/TEM.2023.3321697)

        **{dev_label}**
        - [Me. Eng. Anderson Portella](https://www.linkedin.com/in/andersonportella/)
        - [Prof. Dr. Marcos dos Santos](https://www.linkedin.com/in/profmarcosdossantos/)
        - [Prof. Dr. Carlos Francisco Simões Gomes](https://www.linkedin.com/in/carlos-francisco-sim%C3%B5es-gomes-7284a3b/)
        """,
        unsafe_allow_html=True
    )

    # Download manual PDF
    try:
        with open("metodo.pdf", "rb") as f:
            pdf_bytes_manual = f.read()
        st.download_button(
            label=t("manual_download", lang),
            data=pdf_bytes_manual,
            file_name="WENSLO_ALWAS_manual.pdf",
            mime="application/pdf",
            key="download_method_pdf"
        )
    except FileNotFoundError:
        st.info(t("manual_not_found", lang))

    # Citação
    st.markdown("---")
    with st.expander(t("citation_title", lang)):
        st.markdown(get_citation(lang))
    st.markdown("---")

    # Limpar resultados
    if st.button(t("clear_results", lang)):
        st.session_state['results_ready'] = False
        st.session_state['wens'] = None
        st.session_state['alwas_res'] = None
        st.session_state['graduation_df'] = None
        st.session_state['classes_dict'] = None
        st.session_state['validation'] = None
        st.session_state['std_mat'] = None
        st.session_state['weights'] = None
        st.session_state['excel_bytes'] = None
        st.session_state['corr_matrix'] = None
        st.rerun()

    # Correlação
    corr_method = st.radio(t("corr_method", lang), options=["Pearson", "Spearman"], index=0)
    show_corr_checkbox = st.checkbox(t("show_corr", lang), value=True)

    # Sensibilidade
    st.subheader(t("sensitivity", lang))
    xi = st.slider("ξ (xi)", 1, 50, 1)
    phi = st.slider("φ (phi)", 0.0, 1.0, 0.5, step=0.01)
    theta = st.slider("θ (theta)", 1, 50, 1)

    # G7
    use_g7 = st.checkbox(t("use_g7", lang))

    # Contato
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: grey; font-size: 0.8em;'>"
        "for support/contact: andersonportella@yahoo.com.br"
        "</div>",
        unsafe_allow_html=True
    )

# ===========================================
# ÁREA PRINCIPAL — Logo + título + instituição
# ===========================================
st.markdown("<br>", unsafe_allow_html=True)

if os.path.exists(LOGO_PATH):
    try:
        with open(LOGO_PATH, "rb") as f:
            img_b64 = base64.b64encode(f.read()).decode()
        st.markdown(
            f"<div style='text-align:center;'><img src='data:image/png;base64,{img_b64}' width='160'></div>",
            unsafe_allow_html=True
        )
    except Exception:
        pass

st.markdown(f"<h1 style='text-align:center;'>{t('app_title', lang)}</h1>", unsafe_allow_html=True)
inst_html = INSTITUTION_LINE.replace("<br/>", "<br>")
st.markdown(f"<p style='text-align:center; font-weight:bold;'>{inst_html}</p>", unsafe_allow_html=True)
st.markdown("---")

st.header(t("upload_header", lang))

# ---------- G7 data (para validação) ----------
def get_g7_data():
    data = {
        'C11': [523.19, 258.23, 585.26, 280.37, 1024.07, 306.32, 4285.89],
        'C12': [5727.67, 1237.01, 14372.84, 15433.84, 12787.94, 17447.05, 9461.82],
        'C13': [7.57, 3.23, 3.33, 2.32, 3.18, 2.33, 6.17],
        'C14': [17.26, 11.81, 16.38, 19.42, 6.77, 13.91, 8.5],
        'C15': [67.9, 23.81, 43.56, 41.51, 19.04, 43.13, 19.74],
        'C21': [77.64, 52.32, 46.93, 38.89, 67.15, 63.11, 73.04],
        'C22': [5.78, 43.65, 45.54, 53.44, 24.72, 29.7, 20.37],
        'C23': [10.51, 0.74, 1.24, 1.01, 1.48, 1.42, 2.37],
        'C31': [16.78, 17.4, 28.3, 57.65, 24.06, 8.9, 38.86],
        'C32': [84, 79, 97, 96, 81, 98, 98],
        'C33': [85.7, 81, 97.13, 94, 79.7, 96, 75.4],
        'C41': [12.27, 12.54, 13.33, 9.63, 9.87, 11.33, 8.84],
        'C42': [0.44, 0.54, 0.57, 0.41, 0.43, 0.49, 0.38],
        'C43': [3.84, 1.84, 3.22, 2.94, 2.57, 1.78, 0.34],
        'C44': [11.26, 12.14, 19.91, 19.53, 13.73, 12.1, 10.91],
        'C45': [1.17, 2.38, 1.71, 3.09, 1.24, 2.07, 2.8],
        'C46': [8.8, 45.57, 45.35, 5.87, 8.99, 41.46, 19.11],
        'C51': [144.43, 126.03, 124.68, 101.03, 111.43, 134.46, 147.51],
        'C52': [66.12, 61.59, 64.36, 63.71, 59.15, 63.67, 65],
        'C53': [1.5, 1.85, 1.6, 1.31, 1.37, 1.75, 1.78]
    }
    criteria_types = {
        'C11': 'MIN', 'C12': 'MAX', 'C13': 'MAX', 'C14': 'MAX', 'C15': 'MAX',
        'C21': 'MAX', 'C22': 'MAX', 'C23': 'MAX', 'C31': 'MIN', 'C32': 'MAX',
        'C33': 'MAX', 'C41': 'MAX', 'C42': 'MAX', 'C43': 'MAX', 'C44': 'MAX',
        'C45': 'MAX', 'C46': 'MAX', 'C51': 'MAX', 'C52': 'MAX', 'C53': 'MAX'
    }
    alternatives = ['Canada', 'France', 'Germany', 'Italy', 'Japan', 'UK', 'USA']
    return pd.DataFrame(data, index=alternatives), criteria_types

# ---------- Template Excel ----------
@st.cache_data
def create_excel_template_bytes():
    df, types = get_g7_data()
    header = [''] + list(df.columns)
    type_row = [''] + [types[c] for c in df.columns]
    rows = [header, type_row] + [[idx] + list(df.loc[idx]) for idx in df.index]
    df_out = pd.DataFrame(rows)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
        df_out.to_excel(writer, sheet_name='Dados', header=False, index=False)
        workbook = writer.book
        ws = writer.sheets['Dados']
        header_format = workbook.add_format({'bold': True, 'fg_color': '#D7E4BC', 'border': 1})
        ws.set_row(0, None, header_format)
        ws.set_row(1, None, header_format)
        ws.set_column(0, 0, 16)
        for i in range(1, df.shape[1] + 1):
            ws.set_column(i, i, 14)
    return buf.getvalue()

# ---------- Graduation / Positioning (Tabela V) ----------
def compute_graduation(Z: pd.DataFrame, delta_z: pd.Series):
    records = []
    classes_dict = {}
    for c in Z.columns:
        z_col = Z[c].astype(float)
        z_min = float(z_col.min())
        z_max = float(z_col.max())
        dz = float(delta_z.get(c, 0.0))
        if dz <= 0 or math.isclose(dz, 0.0):
            intervals = [(z_min, z_max)]
            classes_dict[c] = intervals
            for alt, zval in zip(Z.index, z_col.values):
                records.append({
                    'Criterion': c, 'Alternative': alt, 'z_ij': float(zval),
                    'class_index_int': 0, 'class_index': round(0.0, 4),
                    'class_lower': float(z_min), 'class_upper': float(z_max), 'pos_in_class': np.nan
                })
            continue
        n_classes = max(1, int(math.ceil((z_max - z_min) / dz)))
        intervals = []
        for k in range(n_classes):
            lower = z_min + k * dz
            upper = lower + dz
            intervals.append((lower, upper))
        if intervals and intervals[-1][1] < z_max:
            intervals[-1] = (intervals[-1][0], z_max)
        classes_dict[c] = intervals
        for alt, zval in zip(Z.index, z_col.values):
            idx_int = int(math.floor((zval - z_min) / dz))
            if idx_int < 0: idx_int = 0
            if idx_int >= len(intervals): idx_int = len(intervals) - 1
            lower, upper = intervals[idx_int]
            pos = (zval - lower) / (upper - lower) if (upper - lower) > 0 else np.nan
            class_idx_cont = idx_int + (pos if not np.isnan(pos) else 0.0)
            records.append({
                'Criterion': c, 'Alternative': alt, 'z_ij': float(zval),
                'class_index_int': int(idx_int), 'class_index': round(float(class_idx_cont), 4),
                'class_lower': float(lower), 'class_upper': float(upper),
                'pos_in_class': float(pos) if not np.isnan(pos) else np.nan
            })
    graduation_df = pd.DataFrame.from_records(records)
    if not graduation_df.empty:
        graduation_df['class_index_int'] = graduation_df['class_index_int'].astype(int)
        graduation_df['class_index'] = graduation_df['class_index'].astype(float)
        graduation_df['class_lower'] = graduation_df['class_lower'].astype(float)
        graduation_df['class_upper'] = graduation_df['class_upper'].astype(float)
        graduation_df['pos_in_class'] = graduation_df['pos_in_class'].astype(float)
    graduation_df = graduation_df.sort_values(['Criterion', 'class_index', 'Alternative']).reset_index(drop=True)
    return graduation_df, classes_dict

# ---------- WENSLO ----------
def wenslo(decision_matrix: pd.DataFrame):
    m, n = decision_matrix.shape
    Z = decision_matrix.astype(float).copy()
    col_sums = Z.sum(axis=0).replace(0, np.nan)
    Z = Z / col_sums
    delta_z = {}
    for c in Z.columns:
        R = Z[c].max() - Z[c].min()
        delta_z[c] = R / (1.0 + 3.322 * math.log10(m)) if m > 1 else 0.0
    real_accum = Z.cumsum(axis=0)
    tan_phi = {}
    for c in Z.columns:
        dz = delta_z[c]
        if dz == 0 or (m - 1) == 0:
            tan_phi[c] = np.nan
        else:
            tan_phi[c] = Z[c].sum() / ((m - 1) * dz)
    artificial_accum = pd.DataFrame(index=Z.index, columns=Z.columns, dtype=float)
    for c in Z.columns:
        dz = delta_z[c]
        tp = tan_phi[c]
        if pd.isna(tp):
            artificial_accum[c] = np.nan
        else:
            for i, idx in enumerate(Z.index):
                artificial_accum.at[idx, c] = tp * (i * dz)
    E = {}
    for c in Z.columns:
        vals = Z[c].values
        dz = delta_z[c]
        total = 0.0
        for i in range(len(vals) - 1):
            diff = vals[i + 1] - vals[i]
            total += math.sqrt(diff * diff + dz * dz)
        E[c] = total
    tan_phi_series = pd.Series(tan_phi)
    q = {c: (E[c] / tan_phi[c]) if (tan_phi[c] != 0 and not pd.isna(tan_phi[c])) else np.nan for c in Z.columns}
    q_series = pd.Series(q)
    sum_q = q_series.sum(skipna=True)
    w_series = q_series / sum_q if sum_q and not np.isclose(sum_q, 0.0) else q_series * 0.0
    return {
        "Z": Z, "delta_z": pd.Series(delta_z), "real_accum": real_accum,
        "artificial_accum": artificial_accum, "E": pd.Series(E),
        "tan_phi": tan_phi_series, "q": q_series, "weights": w_series
    }

# ---------- ALWAS ----------
def alwas(decision_matrix: pd.DataFrame, weights: pd.Series, criteria_types: dict, xi=1.0, phi=0.5, theta=1.0):
    eps = 1e-12
    M = decision_matrix.astype(float).copy()
    std_mat = pd.DataFrame(index=M.index, columns=M.columns, dtype=float)
    col_max = M.max(axis=0)
    for c in M.columns:
        max_j = col_max[c]
        if max_j == 0:
            std_mat[c] = 0.0
            continue
        if criteria_types.get(c, "MAX").upper() == "MAX":
            std_mat[c] = M[c] / max_j
        else:
            scaled = M[c] / max_j
            std_mat[c] = -scaled + scaled.max() + scaled.min()
    std_mat = std_mat.clip(lower=eps, upper=1.0 - eps)
    R1 = pd.Series(index=M.index, dtype=float)
    R2 = pd.Series(index=M.index, dtype=float)
    w = weights.reindex(M.columns).fillna(0.0)
    for i in M.index:
        row = std_mat.loc[i]
        S_i = row.sum()
        if S_i <= 0:
            R1[i] = np.nan
            R2[i] = np.nan
            continue
        f_vals = (row / S_i).clip(eps, 1 - eps)
        inner1 = np.sum([w[c] * ((-math.log(1.0 - f_vals[c])) ** xi) for c in M.columns])
        inner2 = np.sum([w[c] * ((-math.log(f_vals[c])) ** xi) for c in M.columns])
        inner1 = max(inner1, 0.0)
        inner2 = max(inner2, 0.0)
        g1 = 1.0 - math.exp(-(inner1 ** (1.0 / xi))) if inner1 > 0 else 0.0
        g2 = math.exp(-(inner2 ** (1.0 / xi))) if inner2 > 0 else 1.0
        R1[i] = S_i * g1
        R2[i] = S_i * g2
    S_final = pd.Series(index=M.index, dtype=float)
    for i in M.index:
        r1 = R1[i]
        r2 = R2[i]
        if (pd.isna(r1) or pd.isna(r2)) or (r1 + r2) == 0:
            S_final[i] = np.nan
            continue
        f_R1 = r1 / (r1 + r2)
        f_R2 = r2 / (r1 + r2)
        f_R1 = min(max(f_R1, eps), 1 - eps)
        f_R2 = min(max(f_R2, eps), 1 - eps)
        term1 = phi * (((1.0 - f_R1) / f_R1) ** theta)
        term2 = (1.0 - phi) * (((1.0 - f_R2) / f_R2) ** theta)
        denom = 1.0 + ((term1 + term2) ** (1.0 / theta))
        S_final[i] = (r1 + r2) / denom
    results = pd.DataFrame({"R1": R1, "R2": R2, "S": S_final})
    results["Rank"] = results["S"].rank(ascending=False, method="min").astype(int)
    results = results.sort_values(["Rank", "S"])
    return results, std_mat

# ---------- Validação de acumulação ----------
def validate_accumulation(Z: pd.DataFrame, delta_z: pd.Series):
    m = Z.shape[0]
    real = Z.cumsum(axis=0)
    artificial = pd.DataFrame(index=Z.index, columns=Z.columns, dtype=float)
    mse = {}
    corr = {}
    for c in Z.columns:
        dz = delta_z[c]
        if dz == 0:
            artificial[c] = np.nan
            mse[c] = np.nan
            corr[c] = np.nan
            continue
        tan_phi_c = Z[c].sum() / ((m - 1) * dz)
        for i, idx in enumerate(Z.index):
            artificial.at[idx, c] = tan_phi_c * (i * dz)
        errors = real[c].values - artificial[c].values
        mse[c] = np.mean(errors ** 2)
        try:
            corr_val = np.corrcoef(real[c].values, artificial[c].values)[0, 1]
        except Exception:
            corr_val = np.nan
        corr[c] = corr_val
    return {"real": real, "artificial": artificial, "mse": pd.Series(mse), "corr": pd.Series(corr)}

# ---------- Escrever resultados em Excel ----------
def build_results_excel(buff_dict: dict):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
        for name, obj in buff_dict.items():
            try:
                if isinstance(obj, pd.DataFrame):
                    obj.to_excel(writer, sheet_name=name[:31])
                elif isinstance(obj, pd.Series):
                    obj.to_frame(name=name).to_excel(writer, sheet_name=name[:31])
                else:
                    pd.DataFrame(obj).to_excel(writer, sheet_name=name[:31])
            except Exception as e:
                pd.DataFrame({"info": [f"Could not write sheet {name}: {e}"]}).to_excel(writer, sheet_name=name[:31])
    return buf.getvalue()

# ===========================================
# Upload / Template
# ===========================================
excel_bytes = create_excel_template_bytes()
st.download_button(
    t("download_template", lang),
    excel_bytes,
    file_name="template_wenslo_alwas.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    key="download_template"
)

if use_g7:
    decision_df, criteria_types = get_g7_data()
    st.subheader(t("data_loaded", lang))
    st.dataframe(decision_df)
else:
    uploaded = st.file_uploader(t("upload_prompt", lang), type="xlsx")
    if uploaded is None:
        st.info(t("load_template_first", lang))
        st.stop()
    try:
        raw = pd.read_excel(uploaded, header=None)
        criteria_names = raw.iloc[0, 1:].astype(str).values
        criteria_types_raw = raw.iloc[1, 1:].astype(str).str.upper().values
        alternatives = raw.iloc[2:, 0].astype(str).values
        numeric = raw.iloc[2:, 1:].astype(float).values
        decision_df = pd.DataFrame(numeric, index=alternatives, columns=criteria_names)
        st.subheader(t("data_loaded", lang))
        st.dataframe(decision_df)
        if not all(tp in ["MAX", "MIN"] for tp in criteria_types_raw):
            st.error(t("second_row_error", lang))
            st.stop()
        criteria_types = {name: ('MAX' if typ == 'MAX' else 'MIN') for name, typ in zip(criteria_names, criteria_types_raw)}
    except Exception as e:
        st.error(t("file_read_error", lang) + str(e))
        st.stop()

# ===========================================
# Cálculo principal
# ===========================================
if st.button(t("run", lang)):
    with st.spinner(t("calculating", lang)):
        Z_temp = decision_df.astype(float).copy()
        col_sums = Z_temp.sum(axis=0).replace(0, np.nan)
        Z = Z_temp / col_sums
        m = Z.shape[0]
        delta_z = {}
        for c in Z.columns:
            R = Z[c].max() - Z[c].min()
            delta_z[c] = R / (1.0 + 3.322 * math.log10(m)) if m > 1 else 0.0
        delta_z = pd.Series(delta_z)

        graduation_df, classes_dict = compute_graduation(Z, delta_z)
        wens = wenslo(decision_df)
        weights = wens["weights"].rename("Weight")
        alwas_res, std_mat = alwas(decision_df, weights, criteria_types, xi=xi, phi=phi, theta=theta)
        validation = validate_accumulation(wens["Z"], wens["delta_z"])

        out_dict = {
            "decision_matrix": decision_df,
            "normalized_Z": wens["Z"],
            "delta_z": wens["delta_z"],
            "tan_phi": wens["tan_phi"],
            "E": wens["E"],
            "q": wens["q"],
            "weights": weights,
            "graduation_table": graduation_df,
            "std_matrix": std_mat,
            "ALWAS_results": alwas_res,
            "validation_mse": validation["mse"],
            "validation_corr": validation["corr"],
            "real_accumulation": validation["real"],
            "artificial_accumulation": validation["artificial"]
        }
        excel_out = build_results_excel(out_dict)

        st.session_state['results_ready'] = True
        st.session_state['wens'] = wens
        st.session_state['alwas_res'] = alwas_res
        st.session_state['graduation_df'] = graduation_df
        st.session_state['classes_dict'] = classes_dict
        st.session_state['validation'] = validation
        st.session_state['std_mat'] = std_mat
        st.session_state['weights'] = weights
        st.session_state['excel_bytes'] = excel_out
        st.session_state['run_id'] += 1
        run_id = st.session_state['run_id']

        # --- Resultados ---
        st.subheader(t("graduation_table", lang))
        st.dataframe(graduation_df.style.format({
            'z_ij': '{:.4f}', 'class_index_int': '{:.0f}', 'class_index': '{:.4f}',
            'class_lower': '{:.4f}', 'class_upper': '{:.4f}', 'pos_in_class': '{:.4f}'
        }))

        summary_loc = graduation_df.groupby('Criterion').agg(
            n_classes=('class_index_int', lambda s: s.nunique()),
            min_z=('z_ij', 'min'),
            max_z=('z_ij', 'max')
        )
        summary_loc = summary_loc.rename(columns={
            'n_classes': t('col_n_classes', lang),
            'min_z':     t('col_min_z', lang),
            'max_z':     t('col_max_z', lang)
        })
        st.subheader(t('graduation_summary_header', lang))
        st.dataframe(summary_loc.style.format({
            t('col_min_z', lang): '{:.4f}',
            t('col_max_z', lang): '{:.4f}'
        }))

        st.subheader(t("weights", lang))
        w_df = pd.DataFrame({
            "Weight": weights.round(6),
            "q": wens["q"].round(6),
            "E": wens["E"].round(6),
            "tan_phi": wens["tan_phi"].round(6)
        })
        w_df["Rank"] = w_df["Weight"].rank(ascending=False, method="min").astype(int)
        st.dataframe(w_df.style.format({"Weight": "{:.4f}", "q": "{:.4f}", "E": "{:.4f}", "tan_phi": "{:.4f}"}))

        st.subheader(t("norm_matrix_z", lang))
        st.dataframe(wens["Z"].style.format("{:.4f}"))

        st.subheader(t("delta_z_sturges", lang))
        st.dataframe(wens["delta_z"].to_frame("Delta_z").style.format("{:.4f}"))

        st.subheader(t("envelope_tanphi", lang))
        st.dataframe(pd.DataFrame({"E": wens["E"], "tan_phi": wens["tan_phi"]}).style.format("{:.4f}"))

        st.subheader(t("alwas_std_matrix", lang))
        st.dataframe(std_mat.style.format("{:.4f}"))

        st.subheader(t("ranking", lang))
        st.dataframe(alwas_res.style.format({"R1": "{:.4f}", "R2": "{:.4f}", "S": "{:.4f}", "Rank": "{:.0f}"}))

        st.subheader(t("validation", lang))
        val_df = pd.DataFrame({"MSE": validation["mse"], "Correlation": validation["corr"]})
        st.dataframe(val_df.style.format({"MSE": "{:.4f}", "Correlation": "{:.4f}"}))

        st.download_button(
            t("download_results", lang),
            st.session_state['excel_bytes'],
            file_name=f"wenslo_alwas_results_run{run_id}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"download_results_run_{run_id}"
        )

        st.success(t("calc_done", lang))

# ===========================================
# Heatmap de correlação
# ===========================================
if show_corr_checkbox:
    if st.session_state.get('results_ready', False):
        wens_saved = st.session_state['wens']
        Z_saved = wens_saved['Z']
        method = corr_method.lower()
        try:
            corr = Z_saved.corr(method=method)
        except Exception as e:
            st.error(t("corr_error", lang) + str(e))
            corr = None

        if corr is not None and not corr.empty:
            st.subheader(f"{t('heatmap_title', lang)} ({corr_method})")
            dec = st.sidebar.selectbox(t("decimal_places", lang), options=[2, 3], index=0)

            figsize_x = max(6, 0.4 * len(corr.columns))
            figsize_y = max(5, 0.4 * len(corr.columns))
            fig, ax = plt.subplots(figsize=(figsize_x, figsize_y))
            im = ax.imshow(corr.values, vmin=-1, vmax=1, aspect='auto', cmap='RdBu_r')
            ax.set_xticks(np.arange(len(corr.columns)))
            ax.set_yticks(np.arange(len(corr.index)))
            ax.set_xticklabels(corr.columns, rotation=45, ha='right')
            ax.set_yticklabels(corr.index)
            fmt_str = f"{{val:.{dec}f}}"
            for i in range(corr.shape[0]):
                for j in range(corr.shape[1]):
                    val = corr.values[i, j]
                    ax.text(j, i, fmt_str.format(val=val), ha='center', va='center', fontsize=8, color='black')
            cbar = fig.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
            cbar.ax.set_ylabel('Correlation', rotation=270, labelpad=12)
            plt.tight_layout()
            st.pyplot(fig)

            run_id_for_key = st.session_state['run_id']
            buf = io.BytesIO()
            fig.savefig(buf, format='png', dpi=300, bbox_inches='tight')
            buf.seek(0)
            st.download_button(
                label=f"{t('download_heatmap', lang)} ({corr_method})",
                data=buf,
                file_name=f"heatmap_correlation_{method}_run{run_id_for_key}.png",
                mime="image/png",
                key=f"download_heatmap_{method}_run{run_id_for_key}"
            )
            plt.close(fig)
        else:
            st.info(t("corr_unavailable", lang))
    else:
        st.info(t("corr_no_results", lang))

# ===========================================
# Plot de acumulação reativo
# ===========================================
st.subheader(t("accum_title", lang))

if st.session_state.get('results_ready', False):
    crit_choice = st.selectbox(t("select_criterion", lang), options=list(decision_df.columns))
    wens_saved = st.session_state['wens']

    fig, ax = plt.subplots(figsize=(6, 3))
    ax.plot(wens_saved["real_accum"].index, wens_saved["real_accum"][crit_choice].values,
            marker='o', label=t("accum_real", lang))
    ax.plot(wens_saved["artificial_accum"].index, wens_saved["artificial_accum"][crit_choice].values,
            marker='x', label=t("accum_artificial", lang))
    ax.set_xlabel(t("accum_xlabel", lang))
    ax.set_ylabel(t("accum_ylabel", lang))
    ax.set_title(f"{t('accum_real', lang)} vs {t('accum_artificial', lang)}: {crit_choice}")
    ax.legend()
    ax.grid(True)
    st.pyplot(fig)

    mse_val = st.session_state['validation']['mse'].get(crit_choice, np.nan)
    corr_val = st.session_state['validation']['corr'].get(crit_choice, np.nan)
    st.write(
        f"**MSE**: {np.nan if pd.isna(mse_val) else float(mse_val):.6f} — "
        f"**Correlation**: {np.nan if pd.isna(corr_val) else float(corr_val):.6f}"
    )
else:
    st.info(t("accum_no_results", lang))

# ===========================================
# Download persistente no final
# ===========================================
if st.session_state.get('excel_bytes', None) is not None:
    run_id_final = st.session_state['run_id']
    st.download_button(
        t("download_results", lang),
        st.session_state['excel_bytes'],
        file_name=f"wenslo_alwas_results_run{run_id_final}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"download_results_persistent_run{run_id_final}"
    )
