import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="Analisador de Dados Duplicados",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
    <style>
    .main {
        background-color: #0e1117;
    }
    .stMetric {
        background-color: #1e2130;
        padding: 15px;
        border-radius: 10px;
        border: 1px solid #4e5d6e;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
    }
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background-color: transparent;
        color: #555;
        text-align: center;
        padding: 10px;
        font-size: 12px;
        letter-spacing: 1px;
    }
    h1, h2, h3 {
        color: #00d4ff;
        font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
    }
    .stButton>button {
        width: 100%;
        background-color: #00d4ff;
        color: black;
        font-weight: bold;
        border-radius: 5px;
        border: none;
        transition: 0.3s;
    }
    .stButton>button:hover {
        background-color: #008fb3;
        color: white;
    }
    </style>
    """, unsafe_allow_html=True)

def clean_data(df):
    return df.astype(str).apply(lambda x: x.str.strip()).replace(['nan', 'None', 'NaN', ''], pd.NA)

def load_file(uploaded_file):
    try:
        if uploaded_file.name.endswith('.xlsx'):
            return pd.read_excel(uploaded_file, dtype=str)
        else:
            return pd.read_csv(uploaded_file, dtype=str)
    except Exception as e:
        st.error(f"Erro ao carregar: {e}")
        return None

def to_excel(df, sheet_name="Resultados"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output

st.title("🚀 Analisador de Dados Duplicados")
st.markdown("---")

with st.container():
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### 📂 Planilha Origem")
        file1 = st.file_uploader("Arraste ou selecione o arquivo 1", type=['xlsx', 'csv'], key="u1")
    with col2:
        st.markdown("### 📂 Planilha Destino")
        file2 = st.file_uploader("Arraste ou selecione o arquivo 2", type=['xlsx', 'csv'], key="u2")

if file1 and file2:
    df1 = load_file(file1)
    df2 = load_file(file2)
    if df1 is not None and df2 is not None:
        st.markdown("### ⚙️ Configurações de Análise")
        c_cfg1, c_cfg2, c_cfg3 = st.columns([1, 1, 2])
        with c_cfg1:
            col_choice1 = st.selectbox("Coluna na Planilha 1", df1.columns)
        with c_cfg2:
            col_choice2 = st.selectbox("Coluna na Planilha 2", df2.columns)
        with c_cfg3:
            search_value = st.text_input("🔍 Valor Específico (Opcional)", placeholder="Ex: 9202793806")
        if st.button("INICIAR PROCESSAMENTO"):
            with st.spinner('Analisando dados...'):
                df1_clean = clean_data(df1).dropna(subset=[col_choice1])
                df2_clean = clean_data(df2).dropna(subset=[col_choice2])
                if search_value.strip():
                    count1 = (df1_clean[col_choice1] == search_value).sum()
                    count2 = (df2_clean[col_choice2] == search_value).sum()
                    st.markdown(f"#### 🎯 Resultado para o valor: `{search_value}`")
                    m1, m2 = st.columns(2)
                    m1.metric("Ocorrências na P1", count1)
                    m2.metric("Ocorrências na P2", count2)
                    if count1 > 0 or count2 > 0:
                        tab1, tab2 = st.tabs(["Visualizar na P1", "Visualizar na P2"])
                        with tab1:
                            if count1 > 0:
                                df_res_p1 = df1_clean[df1_clean[col_choice1] == search_value]
                                st.dataframe(df_res_p1, use_container_width=True)
                                excel_p1 = to_excel(df_res_p1, "Resultado_P1")
                                st.download_button(
                                    label="📥 Baixar resultado P1 em XLSX",
                                    data=excel_p1,
                                    file_name=f"resultado_P1_{search_value}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key="down_p1"
                                )
                            else:
                                st.info("Não encontrado na Planilha 1")
                        with tab2:
                            if count2 > 0:
                                df_res_p2 = df2_clean[df2_clean[col_choice2] == search_value]
                                st.dataframe(df_res_p2, use_container_width=True)
                                excel_p2 = to_excel(df_res_p2, "Resultado_P2")
                                st.download_button(
                                    label="📥 Baixar resultado P2 em XLSX",
                                    data=excel_p2,
                                    file_name=f"resultado_P2_{search_value}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key="down_p2"
                                )
                            else:
                                st.info("Não encontrado na Planilha 2")
                    else:
                        st.warning("Valor não encontrado em nenhuma das bases.")
                else:
                    v1 = df1_clean[col_choice1]
                    v2 = df2_clean[col_choice2]
                    matches = set(v1).intersection(set(v2))
                    if matches:
                        st.success(f"✅ Foram identificados {len(matches)} valores duplicados entre as bases.")
                        m1, m2, m3 = st.columns(3)
                        m1.metric("Total Duplicados", len(matches))
                        m2.metric("Total Linhas P1", v1[v1.isin(matches)].count())
                        m3.metric("Total Linhas P2", v2[v2.isin(matches)].count())
                        df2_to_merge = df2_clean.rename(columns={col_choice2: col_choice1})
                        linhas_identicas = pd.merge(df1_clean, df2_to_merge, how='inner')
                        st.markdown("### 📋 Relatório de Linhas 100% Idênticas")
                        if not linhas_identicas.empty:
                            st.dataframe(linhas_identicas, use_container_width=True)
                            excel_id = to_excel(linhas_identicas, "Linhas_Identicas")
                            st.download_button(
                                label="📥 Baixar relatório em XLSX",
                                data=excel_id,
                                file_name="relatorio_linhas_identicas.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="down_identicas"
                            )
                        else:
                            st.warning("Os valores existem em ambas, mas as informações complementares das linhas divergem.")
                    else:
                        st.error("Nenhuma duplicidade encontrada entre as colunas selecionadas.")

st.markdown("""
    <div class="footer">
        Desenvolvido por Djalma A Barbosa. 2026. Direitos Reservados.
    </div>
    """, unsafe_allow_html=True)