import streamlit as st
import pandas as pd
import numpy as np
import io
import math
import matplotlib.pyplot as plt
import datetime

st.set_page_config(page_title="WENSLO + ALWAS Tool", layout="wide")

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

# ---------- Traduções (PT / EN) ----------
translations = {
    "Português": {
        "title": "Ferramenta WENSLO + ALWAS",
        "upload_header": "1. Passos",
        "download_template": "Baixar template Excel",
        "upload_prompt": "Carregue o arquivo (xlsx)",
        "run": "Calcular",
        "weights": "Pesos (WENSLO)",
        "ranking": "Ranking (ALWAS)",
        "download_results": "Baixar todos os resultados (Excel)",
        "select_criterion": "Selecione um critério para plotar acumulação",
        "validation": "Validação da acumulação (MSE / Correlação)",
        "sensitivity": "Análise de sensibilidade (ξ, φ, θ)",
        "data_loaded": "Dados carregados (confira se estão corretos)",
        "graduation_table": "Matriz de Decisão Normalizada",
        "graduation_summary_header": "Resumo da graduação por critério",
        "col_n_classes": "Nº de classes",
        "col_min_z": "z (mín)",
        "col_max_z": "z (máx)",
        "citation_title": "📚 Como citar"
    },
    "English": {
        "title": "WENSLO + ALWAS Tool",
        "upload_header": "1. Steps",
        "download_template": "Download Excel template",
        "upload_prompt": "Upload your xlsx file",
        "run": "Run",
        "weights": "Weights (WENSLO)",
        "ranking": "Ranking (ALWAS)",
        "download_results": "Download all results (Excel)",
        "select_criterion": "Select a criterion to plot accumulation",
        "validation": "Accumulation validation (MSE / Correlation)",
        "sensitivity": "Sensitivity analysis (ξ, φ, θ)",
        "data_loaded": "Loaded data (check correctness)",
        "graduation_table": "Normalized Decision-Making Matrix",
        "graduation_summary_header": "Graduation summary by criterion",
        "col_n_classes": "Number of classes",
        "col_min_z": "z (min)",
        "col_max_z": "z (max)",
        "citation_title": "📚 How to cite"
    }
}

# ---------- CITAÇÕES ----------
CITATIONS = {
    "ABNT": """
**Formato ABNT:**

**Software:**
SANTOS, Marcos dos; GOMES, Carlos Francisco Simões. **Wenslo-Alwas Tool**. [S.l.]: Anderson Gonçalves Portella, 2025. Programa de Computador. Registro INPI: BR512025005226-0. Disponível em: <https://wenslo-alwas-tool.streamlit.app/>. Acesso em: {date}.

**Artigo:**
SILVA, C. S.; SANTOS, M. R. Análise do nível de maturidade em gestão de riscos: um estudo de caso em uma empresa do setor elétrico. In: CONGRESSO NACIONAL DE EXCELÊNCIA EM GESTÃO, 19., 2025, Online. **Anais...** Rio de Janeiro: CNEG, 2025. DOI: doi.org. Acesso em: {date}.
""",
    "APA": """
**APA Format:**

**Software:**
Santos, M. dos, & Gomes, C. F. S. (2025). *Wenslo-Alwas Tool* [Computer software]. Anderson Gonçalves Portella. https://wenslo-alwas-tool.streamlit.app/

**Article:**
Silva, C. S., & Santos, M. R. (2025). Análise do nível de maturidade em gestão de riscos: um estudo de caso em uma empresa do setor elétrico. Em *Anais do XIX Congresso Nacional de Excelência em Gestão*. DOI: 10.14488/cneg2025_cneg_pt_068_0567_23581
"""
}

def get_citation(lang):
    """Retorna a citação no formato apropriado com a data atual"""
    citation_type = "ABNT" if lang == "Português" else "APA"
    today = datetime.datetime.now().strftime("%d %b. %Y")
    return CITATIONS[citation_type].format(date=today)


# ---------- UI: idioma / strings ----------
lang = st.sidebar.selectbox("Idioma / Language", options=["Português", "English"], index=0)
T = translations[lang]

# Sidebar: links (artigo + autores)
st.sidebar.markdown(
    """
    **Artigo / Article**

    - [*A Novel WENSLO and ALWAS Multicriteria Methodology and Its Application to Green Growth Performance Evaluation*](https://doi.org/10.1109/TEM.2023.3321697)

    **Developers / Desenvolvedores**
    - [Anderson Portella](https://www.linkedin.com/in/andersonportella/)
    - [Prof. Dr. Marcos dos Santos](https://www.linkedin.com/in/profmarcosdossantos/)
    - [Prof. Dr. Carlos Francisco Simões Gomes](https://www.linkedin.com/in/carlos-francisco-sim%C3%B5es-gomes-7284a3b/)
    """,
    unsafe_allow_html=True
)
# botão de download para um PDF local
try:
    with open("metodo.pdf", "rb") as f:
        pdf_bytes = f.read()

    st.sidebar.download_button(
        label="📘 Baixar manual / Download manual",
        data=pdf_bytes,
        file_name="WENSLO_ALWAS_manual.pdf",
        mime="application/pdf",
        key="download_method_pdf"
    )
except FileNotFoundError:
    st.sidebar.info("PDF manual not found.")

# ---------- CITAÇÃO NA SIDEBAR ----------
st.sidebar.markdown("---")
with st.sidebar.expander(T["citation_title"]):
    st.markdown(get_citation(lang))
st.sidebar.markdown("---")

if st.sidebar.button("Limpar resultados / Clean results"):
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
    
# Escolha do tipo de correlação e controle de exibição
corr_method = st.sidebar.radio("Método de correlação / Correlation method", options=["Pearson", "Spearman"], index=0)
show_corr_checkbox = st.sidebar.checkbox("Mostrar heatmap de correlação", value=True)

# ---------- Sensitivity params (MOVIDO PARA CIMA) ----------
st.sidebar.subheader(T["sensitivity"])
xi = st.sidebar.slider("ξ (xi)", 1, 50, 1)
phi = st.sidebar.slider("φ (phi)", 0.0, 1.0, 0.5, step=0.01)
theta = st.sidebar.slider("θ (theta)", 1, 50, 1)

# --- CONTATO / SUPPORT (MOVIDO PARA CIMA) ---
st.sidebar.markdown("---")
st.sidebar.markdown("for support/contact: andersonportella@yahoo.com.br")

# ==========================================
# INÍCIO DA ÁREA PRINCIPAL
# ==========================================

st.title(T["title"])
st.header(T["upload_header"])

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
                    'Criterion': c,
                    'Alternative': alt,
                    'z_ij': float(zval),
                    'class_index_int': 0,
                    'class_index': round(0.0, 4),
                    'class_lower': float(z_min),
                    'class_upper': float(z_max),
                    'pos_in_class': np.nan
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
            if idx_int < 0:
                idx_int = 0
            if idx_int >= len(intervals):
                idx_int = len(intervals) - 1
            lower, upper = intervals[idx_int]
            pos = (zval - lower) / (upper - lower) if (upper - lower) > 0 else np.nan
            class_idx_cont = idx_int + (pos if not np.isnan(pos) else 0.0)
            records.append({
                'Criterion': c,
                'Alternative': alt,
                'z_ij': float(zval),
                'class_index_int': int(idx_int),
                'class_index': round(float(class_idx_cont), 4),
                'class_lower': float(lower),
                'class_upper': float(upper),
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
        "Z": Z,
        "delta_z": pd.Series(delta_z),
        "real_accum": real_accum,
        "artificial_accum": artificial_accum,
        "E": pd.Series(E),
        "tan_phi": tan_phi_series,
        "q": q_series,
        "weights": w_series
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
        g1 = 1.0 - math.exp(- (inner1 ** (1.0 / xi))) if inner1 > 0 else 0.0
        g2 = math.exp(- (inner2 ** (1.0 / xi))) if inner2 > 0 else 1.0
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

# ---------- Interface: upload / template ----------
excel_bytes = create_excel_template_bytes()
st.download_button(
    T["download_template"],
    excel_bytes,
    file_name="template_wenslo_alwas.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    key="download_template"
)

use_g7 = st.sidebar.checkbox("Usar dados do G7 (validação)")
if use_g7:
    decision_df, criteria_types = get_g7_data()
    st.subheader(T["data_loaded"])
    st.dataframe(decision_df)
else:
    uploaded = st.file_uploader(T["upload_prompt"], type="xlsx")
    if uploaded is None:
        st.info("Carregue o template preenchido ou use os dados G7 para teste.")
        st.stop()
    try:
        raw = pd.read_excel(uploaded, header=None)
        criteria_names = raw.iloc[0, 1:].astype(str).values
        criteria_types_raw = raw.iloc[1, 1:].astype(str).str.upper().values
        alternatives = raw.iloc[2:, 0].astype(str).values
        numeric = raw.iloc[2:, 1:].astype(float).values
        decision_df = pd.DataFrame(numeric, index=alternatives, columns=criteria_names)
        st.subheader(T["data_loaded"])
        st.dataframe(decision_df)
        if not all(t in ["MAX", "MIN"] for t in criteria_types_raw):
            st.error("Segunda linha deve conter apenas 'MAX' ou 'MIN' para cada critério.")
            st.stop()
        criteria_types = {name: ('MAX' if typ == 'MAX' else 'MIN') for name, typ in zip(criteria_names, criteria_types_raw)}
    except Exception as e:
        st.error("Erro ao ler o arquivo: " + str(e))
        st.stop()

# ---------- Run (calcular e salvar em session_state) ----------
if st.button(T["run"]):
    with st.spinner("Calculando..."):
        # Normalização Z e Δz
        Z_temp = decision_df.astype(float).copy()
        col_sums = Z_temp.sum(axis=0).replace(0, np.nan)
        Z = Z_temp / col_sums
        m = Z.shape[0]
        delta_z = {}
        for c in Z.columns:
            R = Z[c].max() - Z[c].min()
            delta_z[c] = R / (1.0 + 3.322 * math.log10(m)) if m > 1 else 0.0
        delta_z = pd.Series(delta_z)

        # Tabela V
        graduation_df, classes_dict = compute_graduation(Z, delta_z)

        # WENSLO
        wens = wenslo(decision_df)
        weights = wens["weights"].rename("Weight")

        # ALWAS
        alwas_res, std_mat = alwas(decision_df, weights, criteria_types, xi=xi, phi=phi, theta=theta)

        # Validação
        validation = validate_accumulation(wens["Z"], wens["delta_z"])

        # Preparar excel com todos os resultados
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

        # SALVA no session_state
        st.session_state['results_ready'] = True
        st.session_state['wens'] = wens
        st.session_state['alwas_res'] = alwas_res
        st.session_state['graduation_df'] = graduation_df
        st.session_state['classes_dict'] = classes_dict
        st.session_state['validation'] = validation
        st.session_state['std_mat'] = std_mat
        st.session_state['weights'] = weights
        st.session_state['excel_bytes'] = excel_out

        # atualiza run_id para keys únicas
        st.session_state['run_id'] += 1
        run_id = st.session_state['run_id']

        # Exibir resultados
        st.subheader(T["graduation_table"])
        st.dataframe(graduation_df.style.format({
            'z_ij': '{:.4f}',
            'class_index_int': '{:.0f}',
            'class_index': '{:.4f}',
            'class_lower': '{:.4f}',
            'class_upper': '{:.4f}',
            'pos_in_class': '{:.4f}'
        }))

        # Resumo da graduação
        summary_loc = graduation_df.groupby('Criterion').agg(
            n_classes = ('class_index_int', lambda s: s.nunique()),
            min_z = ('z_ij', 'min'),
            max_z = ('z_ij', 'max')
        )
        summary_loc = summary_loc.rename(columns={
            'n_classes': T['col_n_classes'],
            'min_z': T['col_min_z'],
            'max_z': T['col_max_z']
        })
        st.subheader(T['graduation_summary_header'])
        st.dataframe(summary_loc.style.format({
            T['col_min_z']: '{:.4f}',
            T['col_max_z']: '{:.4f}'
        }))

        # Pesos
        st.subheader(T["weights"])
        w_df = pd.DataFrame({
            "Weight": weights.round(6),
            "q": wens["q"].round(6),
            "E": wens["E"].round(6),
            "tan_phi": wens["tan_phi"].round(6)
        })
        w_df["Rank"] = w_df["Weight"].rank(ascending=False, method="min").astype(int)
        st.dataframe(w_df.style.format({"Weight": "{:.4f}", "q": "{:.4f}", "E": "{:.4f}", "tan_phi": "{:.4f}"}))

        st.subheader("Matriz normalizada Z / Normalized matrix Z")
        st.dataframe(wens["Z"].style.format("{:.4f}"))

        st.subheader("Δz (Sturges) per criterion")
        st.dataframe(wens["delta_z"].to_frame("Delta_z").style.format("{:.4f}"))

        st.subheader("Envelope E and tan_phi")
        st.dataframe(pd.DataFrame({"E": wens["E"], "tan_phi": wens["tan_phi"]}).style.format("{:.4f}"))

        st.subheader("ALWAS: standardized home matrix (ζ̂)")
        st.dataframe(std_mat.style.format("{:.4f}"))

        st.subheader(T["ranking"])
        st.dataframe(alwas_res.style.format({"R1": "{:.4f}", "R2": "{:.4f}", "S": "{:.4f}", "Rank": "{:.0f}"}))

        st.subheader(T["validation"])
        val_df = pd.DataFrame({"MSE": validation["mse"], "Correlation": validation["corr"]})
        st.dataframe(val_df.style.format({"MSE": "{:.4f}", "Correlation": "{:.4f}"}))

        # download persistente (unique key per run)
        st.download_button(
            T["download_results"],
            st.session_state['excel_bytes'],
            file_name=f"wenslo_alwas_results_run{run_id}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"download_results_run_{run_id}"
        )

        st.success("Cálculo finalizado!")

# ---------- Heatmap: correlação entre critérios (apresentação controlada por checkbox) ----------
if show_corr_checkbox:
    if st.session_state.get('results_ready', False):
        wens_saved = st.session_state['wens']
        Z_saved = wens_saved['Z']

        method = corr_method.lower()  # 'pearson' ou 'spearman'
        try:
            corr = Z_saved.corr(method=method)
        except Exception as e:
            st.error(f"Erro ao calcular correlação ({method}): {e}")
            corr = None

        if corr is not None and not corr.empty:
            st.subheader(f"Heatmap — Correlação entre critérios / Correlation between criteria ({corr_method})")
            dec = st.sidebar.selectbox("Casas decimais / Decimal places", options=[2, 3], index=0)

            figsize_x = max(6, 0.4 * len(corr.columns))
            figsize_y = max(5, 0.4 * len(corr.columns))
            fig, ax = plt.subplots(figsize=(figsize_x, figsize_y))
            im = ax.imshow(corr.values, vmin=-1, vmax=1, aspect='auto', cmap='RdBu_r')
            ax.set_xticks(np.arange(len(corr.columns)))
            ax.set_yticks(np.arange(len(corr.index)))
            ax.set_xticklabels(corr.columns, rotation=45, ha='right')
            ax.set_yticklabels(corr.index)
            fmt = f"{{val:.{dec}f}}"
            for i in range(corr.shape[0]):
                for j in range(corr.shape[1]):
                    val = corr.values[i, j]
                    ax.text(j, i, fmt.format(val=val), ha='center', va='center', fontsize=8, color='black')
            cbar = fig.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
            cbar.ax.set_ylabel('Correlation', rotation=270, labelpad=12)
            plt.tight_layout()
            st.pyplot(fig)

            # gerar PNG em memória e oferecer download (key inclui run_id e método)
            run_id_for_key = st.session_state['run_id']
            buf = io.BytesIO()
            fig.savefig(buf, format='png', dpi=300, bbox_inches='tight')
            buf.seek(0)
            file_name_png = f"heatmap_correlation_{method}_run{run_id_for_key}.png"
            st.download_button(
                label=f"Download heatmap PNG ({corr_method})",
                data=buf,
                file_name=file_name_png,
                mime="image/png",
                key=f"download_heatmap_{method}_run{run_id_for_key}"
            )
            plt.close(fig)
        else:
            st.info("Não foi possível calcular a matriz de correlação — verifique os dados.")
    else:
        st.info("Resultados ainda não gerados. Clique em 'Calcular' para habilitar o heatmap de correlação.")

# ---------- Visualização reativa do plot de acumulação (fora do botão) ----------
st.subheader("Acumulação: real vs artificial (por critério) / Accumulation: real vs artificial (by criteria)")

if st.session_state.get('results_ready', False):
    crit_choice = st.selectbox(T["select_criterion"], options=list(decision_df.columns))

    wens_saved = st.session_state['wens']

    fig, ax = plt.subplots(figsize=(6, 3))
    ax.plot(wens_saved["real_accum"].index, wens_saved["real_accum"][crit_choice].values, marker='o', label="Real accumulation")
    ax.plot(wens_saved["artificial_accum"].index, wens_saved["artificial_accum"][crit_choice].values, marker='x', label="Artificial (hypotenuse)")
    ax.set_xlabel("Alternatives (ordered as in input)")
    ax.set_ylabel("Accumulated normalized value")
    ax.set_title(f"Real vs Artificial accumulation: {crit_choice}")
    ax.legend()
    ax.grid(True)
    st.pyplot(fig)

    mse_val = st.session_state['validation']['mse'].get(crit_choice, np.nan)
    corr_val = st.session_state['validation']['corr'].get(crit_choice, np.nan)
    st.write(f"**MSE**: {np.nan if pd.isna(mse_val) else float(mse_val):.6f} — **Correlation**: {np.nan if pd.isna(corr_val) else float(corr_val):.6f}")
else:
    st.info("Ainda não há resultados. Clique em 'Calcular' para gerar as tabelas e habilitar os gráficos interativos.")

# ---------- Final: mostrar botão de download persistente se existir excel_bytes ----------
if st.session_state.get('excel_bytes', None) is not None:
    # botão persistente com key distinta (usa run_id para unicidade)
    run_id_final = st.session_state['run_id']
    st.download_button(
        T["download_results"],
        st.session_state['excel_bytes'],
        file_name=f"wenslo_alwas_results_run{run_id_final}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"download_results_persistent_run{run_id_final}"
    )