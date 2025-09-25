import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill

# Configuração da página
st.set_page_config(page_title="Processamento de Planilhas", layout="wide")
st.title("📊 App Online de Processamento de Planilhas de Estudantes")

# Upload da planilha
uploaded_file = st.file_uploader("Escolha um arquivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    # Lê todas as abas
    xls = pd.ExcelFile(uploaded_file)
    st.write("Abas encontradas:", xls.sheet_names)
    
    # Seleção da aba
    sheet_name = st.selectbox("Selecione a aba para processar", xls.sheet_names)
    
    if sheet_name:
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        st.write("Pré-visualização da planilha:")
        st.dataframe(df.head())
        
        # Seleção da avaliativa
        avaliativa = st.selectbox("Selecione a Avaliativa", [1, 2, 3, 4])
        
        # Nome da coluna que será filtrada
        col_name = f"atividade avaliativa {avaliativa}"
        
        if col_name in df.columns:
            # Filtra alunos com resultado pendente
            alunos_sem_resultado = df[df[col_name] == "--"][["DR", "Polo", "Nome"]]
            
            if not alunos_sem_resultado.empty:
                st.subheader("Estudantes com resultado pendente")
                
                # Aplicando cores alternadas
                def color_rows(row):
                    return ['background-color: #E0F7FA' if row.name % 2 == 0 else '']*len(row)
                
                st.dataframe(alunos_sem_resultado.style.apply(color_rows, axis=1))
                
                # Opção de download
                towrite = BytesIO()
                
                # Criando Excel com cores
                with pd.ExcelWriter(towrite, engine='openpyxl') as writer:
                    alunos_sem_resultado.to_excel(writer, index=False, sheet_name="Pendentes")
                    ws = writer.sheets["Pendentes"]
                    
                    # Aplicando cores alternadas no Excel
                    fill = PatternFill(start_color="E0F7FA", end_color="E0F7FA", fill_type="solid")
                    for idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=0):
                        if idx % 2 == 0:
                            for cell in row:
                                cell.fill = fill
                
                towrite.seek(0)
                
                st.download_button(
                    label="⬇️ Baixar Excel",
                    data=towrite,
                    file_name=f"alunos_avaliativa_{avaliativa}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("Nenhum aluno com resultado pendente encontrado.")
        else:
            st.warning(f"A coluna '{col_name}' não foi encontrada na planilha.")
