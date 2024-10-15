import pandas as pd
import plotly.express as px
import streamlit as st
import plotly.io as pio
from fpdf import FPDF
import os
import stat

uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=["xlsx", "xls"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    
    verificacao = ['instrutor', 'aprovado', 'nm_formacao', 'nm_unidade']
    
    if all(col in df.columns for col in verificacao):
        width, height = 700, 400

        aprovados_instrutor = df[df['aprovado'] == True].groupby('instrutor').size().reset_index(name='Quantidade')
        fig1 = px.bar(aprovados_instrutor, x='instrutor', y='Quantidade', title='Aprovados por Instrutor', 
                      color='instrutor', width=width, height=height, color_discrete_sequence=px.colors.qualitative.Set2)

        reprovados_instrutor = df[df['aprovado'] == False].groupby('instrutor').size().reset_index(name='Quantidade')
        fig2 = px.bar(reprovados_instrutor, x='instrutor', y='Quantidade', title='Reprovados por Instrutor', 
                      color='instrutor', width=width, height=height, color_discrete_sequence=px.colors.qualitative.Set2)

        aprovados_formacao = df[df['aprovado'] == True].groupby('nm_formacao').size().reset_index(name='Quantidade')
        fig3 = px.bar(aprovados_formacao, x='nm_formacao', y='Quantidade', title='Aprovados por Formação', 
                      color='nm_formacao', width=width, height=height, color_discrete_sequence=px.colors.qualitative.Set2)

        aprovados_unidade = df[df['aprovado'] == True].groupby('nm_unidade').size().reset_index(name='Quantidade')
        fig4 = px.bar(aprovados_unidade, x='nm_unidade', y='Quantidade', title='Aprovados por Unidade (Todos)', 
                      color='nm_unidade', width=width, height=height, color_discrete_sequence=px.colors.qualitative.Set2)


        aprovados_top10_unidade = aprovados_unidade.sort_values(by='Quantidade', ascending=False).head(10)
        fig5 = px.bar(aprovados_top10_unidade, x='nm_unidade', y='Quantidade', title='10 Unidades com mais Aprovados', 
                      color='nm_unidade', width=width, height=height, color_discrete_sequence=px.colors.qualitative.Set2)

        aprovados_bottom10_unidade = aprovados_unidade.sort_values(by='Quantidade', ascending=False).tail(10)
        fig6 = px.bar(aprovados_bottom10_unidade, x='nm_unidade', y='Quantidade', title='10 Unidades com menos Aprovados', 
                      color='nm_unidade', width=width, height=height, color_discrete_sequence=px.colors.qualitative.Set2)

        st.header('Aprovados por Instrutor')
        st.plotly_chart(fig1)

        st.header('Reprovados por Instrutor')
        st.plotly_chart(fig2)

        st.header('Aprovados por Formação')
        st.plotly_chart(fig3)

        st.header('Aprovados por Unidade (Todos)')
        st.plotly_chart(fig4)

        st.header('10 Unidades com mais Aprovados')
        st.plotly_chart(fig5)

        st.header('10 Unidades com menos Aprovados')
        st.plotly_chart(fig6)

        image_folder = 'images'
        garantir_permissoes_pasta(image_folder)

        pio.write_image(fig1, os.path.join(image_folder, 'aprovados_instrutor.png'))
        pio.write_image(fig2, os.path.join(image_folder, 'reprovados_instrutor.png'))
        pio.write_image(fig3, os.path.join(image_folder, 'aprovados_formacao.png'))
        pio.write_image(fig4, os.path.join(image_folder, 'aprovados_unidade_todos.png'))
        pio.write_image(fig5, os.path.join(image_folder, 'top10_unidades.png'))
        pio.write_image(fig6, os.path.join(image_folder, 'bottom10_unidades.png'))

        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)

        pdf.add_page()
        pdf.set_font('Arial', 'B', 16)
        pdf.cell(200, 10, txt="Aprovados por Instrutor", ln=True, align='C')
        pdf.image(os.path.join(image_folder, 'aprovados_instrutor.png'), x=10, y=20, w=pdf.w - 20)

        pdf.add_page()
        pdf.cell(200, 10, txt="Reprovados por Instrutor", ln=True, align='C')
        pdf.image(os.path.join(image_folder, 'reprovados_instrutor.png'), x=10, y=20, w=pdf.w - 20)

        pdf.add_page()
        pdf.cell(200, 10, txt="Aprovados por Formação", ln=True, align='C')
        pdf.image(os.path.join(image_folder, 'aprovados_formacao.png'), x=10, y=20, w=pdf.w - 20)

        pdf.add_page()
        pdf.cell(200, 10, txt="Aprovados por Unidade (Todos)", ln=True, align='C')
        pdf.image(os.path.join(image_folder, 'aprovados_unidade_todos.png'), x=10, y=20, w=pdf.w - 20)

        pdf.add_page()
        pdf.cell(200, 10, txt="10 Unidades com mais Aprovados", ln=True, align='C')
        pdf.image(os.path.join(image_folder, 'top10_unidades.png'), x=10, y=20, w=pdf.w - 20)

        pdf.add_page()
        pdf.cell(200, 10, txt="10 Unidades com menos Aprovados", ln=True, align='C')
        pdf.image(os.path.join(image_folder, 'bottom10_unidades.png'), x=10, y=20, w=pdf.w - 20)

        pdf.output('saida.pdf')

    else:
        st.error(f"Colunas faltando. O arquivo deve conter as colunas: {verificacao}")
else:
    st.info("Carregue um arquivo Excel.")
