import streamlit as st
import pandas as pd
from io import BytesIO

# Configuração da página
st.set_page_config(page_title="Processamento de Planilhas", layout="wide")

st.title("📊 Aplicativo de Processamento de Planilhas Excel")

# Upload do arquivo Excel
uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=["xlsx"])
if uploaded_file:
    # Lê todas as abas
    xls = pd.ExcelFile(uploaded_file)
    st.write("Abas encontradas:", xls.sheet_names)
    
    # Escolha da aba
    sheet_name = st.selectbox("Selecione a aba para processar", xls.sheet_names)
    
    if sheet_name:
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        st.write("Pré-visualização da planilha:")
        st.dataframe(df.head())
        
        # Escolha da avaliativa
        avaliativa = st.selectbox("Selecione a Avaliativa", [1, 2, 3, 4])
        
        # Filtrando estudantes com "--" na avaliativa selecionada
        col_name = f"Atividade avaliativa {avaliativa}"  # Ajuste para o seu padrão de colunas
        if col_name in df.columns:
            alunos_sem_resultado = df[df[col_name] == "--"][["DR", "Polo"]]
            
            st.subheader("Estudantes com resultado pendente")
            st.dataframe(alunos_sem_resultado)
            
            # Opção de download
            towrite = BytesIO()
            alunos_sem_resultado.to_excel(towrite, index=False)
            towrite.seek(0)
            st.download_button(
                label="⬇️ Baixar Excel",
                data=towrite,
                file_name=f"alunos_avaliativa_{avaliativa}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning(f"A coluna '{col_name}' não foi encontrada na planilha.")
