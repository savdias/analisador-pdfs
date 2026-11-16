import os
import io
import re
import pandas as pd
import streamlit as st
import plotly.express as px
import requests
from analisepdf import analisar_pdf_para_planilha
from openpyxl.styles import Alignment, Font

st.set_page_config(
    page_title="Representatividade Feminina no PNLD - Física",
    page_icon="👩‍🔬",
    layout="wide"
)

st.title("👩‍🔬 Analisador de Representatividade Feminina no PNLD (Física)")

# ==========================================
# MENU LATERAL: DADOS DO MATERIAL (PNLD)
# ==========================================
st.sidebar.header("📚 1. Dados do Livro Didático")
st.sidebar.info("Preencha os metadados da obra do PNLD antes de exportar a planilha final.")

meta_livro = st.sidebar.text_input("Nome do Livro ou Coleção:")
meta_editora = st.sidebar.text_input("Editora:")
meta_volume = st.sidebar.text_input("Volume:")
meta_etapa = st.sidebar.text_input("Etapa de ensino:", value="Ensino Médio")
meta_ano = st.sidebar.text_input("Ano, série ou edição PNLD:")
meta_disciplina = st.sidebar.text_input("Disciplina:", value="Física")
meta_observacao = st.sidebar.text_area("Observações gerais (Opcional):")

metadados_livro = {
    "livro": meta_livro,
    "editora": meta_editora,
    "volume": meta_volume,
    "etapa": meta_etapa,
    "ano": meta_ano,
    "disciplina": meta_disciplina,
    "observacao": meta_observacao
}

# ==========================================
# FUNÇÃO: EXPORTAR ABAS FORMATADAS
# ==========================================
def converter_para_excel(df_resultado, metadados):
    buffer = io.BytesIO()
    
    col_pagina = "Página" if "Página" in df_resultado.columns else df_resultado.columns[0]
    col_palavra = (
        "Palavras-Chave" if "Palavras-Chave" in df_resultado.columns
        else "Categorias" if "Categorias" in df_resultado.columns
        else df_resultado.columns[1]
    )
    col_trecho = next((col for col in df_resultado.columns if "Trecho" in col), df_resultado.columns[2])
    
    col_classif_ia = "Classificação IA" if "Classificação IA" in df_resultado.columns else ""
    col_cat_ia = "Categoria Principal IA" if "Categoria Principal IA" in df_resultado.columns else ""
    col_genero_ia = (
        "Gênero Textual" if "Gênero Textual" in df_resultado.columns
        else "Gêneros Identificados na Página" if "Gêneros Identificados na Página" in df_resultado.columns
        else ""
    )
    col_just_ia = "Justificativa IA" if "Justificativa IA" in df_resultado.columns else ""

    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        
        # --- ABA 1: 00_COMECE_AQUI ---
        df_aba1 = pd.DataFrame({
            "Dado do livro": [
                "Nome do Livro ou Coleção", "Editora", "Volume",
                "Etapa de ensino", "Ano, série ou edição PNLD",
                "Disciplina", "Observações gerais"
            ],
            "Preenchimento": [
                metadados["livro"], metadados["editora"], metadados["volume"],
                metadados["etapa"], metadados["ano"], metadados["disciplina"],
                metadados["observacao"]
            ]
        })
        df_aba1.to_excel(writer, index=False, sheet_name='00_COMECE_AQUI')
        
        # --- ABA 2: 01_REGISTRO_UR ---
        df_aba2 = pd.DataFrame()
        df_aba2["Código da UR"] = ["UR" + str(i+1) for i in range(len(df_resultado))]
        df_aba2["Editora"] = metadados["editora"]
        df_aba2["Livro ou coleção"] = metadados["livro"]
        df_aba2["Volume"] = metadados["volume"]
        df_aba2["Etapa de ensino"] = metadados["etapa"]
        df_aba2["Ano ou série"] = metadados["ano"]
        df_aba2["Disciplina"] = metadados["disciplina"]
        df_aba2["Página PDF*"] = df_resultado[col_pagina]
        df_aba2["Página impressa"] = ""
        df_aba2["Capítulo / seção"] = ""
        df_aba2["Palavra-chave localizada"] = df_resultado[col_palavra]
        df_aba2["trecho retirado do livro"] = df_resultado[col_trecho]
        df_aba2["UC — contexto da ocorrência"] = df_resultado[col_trecho]
        df_aba2["Tipo de ocorrência*"] = ""
        
        df_aba2["Gênero textual / didático"] = df_resultado[col_genero_ia] if col_genero_ia else ""
        df_aba2["Classificação A–E*"] = df_resultado[col_classif_ia] if col_classif_ia else ""
        df_aba2["Categoria principal"] = df_resultado[col_cat_ia] if col_cat_ia else ""
        df_aba2["Justificativa curta"] = df_resultado[col_just_ia] if col_just_ia else ""
        
        df_aba2["Há elemento visual?*"] = ""
        df_aba2["Necessita revisão?*"] = ""
        df_aba2["Observações"] = ""
        
        df_aba2.to_excel(writer, index=False, sheet_name='01_REGISTRO_UR')
        
        # --- ABA 3: 04_QUANTIFICACAO ---
        quant_palavras = df_resultado[col_palavra].value_counts().reset_index()
        quant_palavras.columns = ["Palavra ou expressão", "Contagem lexical no PDF"]
        quant_palavras.to_excel(writer, index=False, sheet_name='04_QUANTIFICACAO')
        
        # --- ABA 4: 03_RESUMO ---
        if col_classif_ia:
            resumo_cat = df_resultado[col_classif_ia].value_counts().reset_index()
            resumo_cat.columns = ["Classificação / Categoria IA", "Quantidade de URs"]
        else:
            resumo_cat = pd.DataFrame(columns=["Classificação / Categoria IA", "Quantidade de URs"])
        resumo_cat.to_excel(writer, index=False, sheet_name='03_RESUMO')
        
        # --- ABA 5: 02_PALAVRAS_CHAVE ---
        df_aba5 = df_aba2.groupby("Palavra-chave localizada").agg(
            Quantidade=("Código da UR", "count"),
            Codigos_URs=("Código da UR", lambda x: ", ".join(x))
        ).reset_index()
        
        df_aba5.rename(columns={
            "Palavra-chave localizada": "Palavra ou expressão",
            "Codigos_URs": "Código(s) das URs"
        }, inplace=True)
        df_aba5.insert(2, "Gerou alguma UR?", "Sim")
        df_aba5["Observações"] = ""
        df_aba5.to_excel(writer, index=False, sheet_name='02_PALAVRAS_CHAVE')
        
        # Formatação
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for col in worksheet.columns:
                col_letter = col[0].column_letter
                worksheet.column_dimensions[col_letter].width = 35
                for cell in col:
                    cell.alignment = Alignment(wrap_text=True, vertical='top')
            for cell in worksheet[1]:
                cell.font = Font(bold=True)
                
    return buffer.getvalue()

# -------------------------------------------------------------------------
# SEÇÃO 1: EXTRAÇÃO DO PDF
# -------------------------------------------------------------------------
st.header("1. Extração de Trechos do Livro Didático (PDF)")

with st.container(border=True):
    col1, col2 = st.columns([1, 2])
    with col1:
        arquivo_enviado = st.file_uploader("Selecione o PDF do livro de Física", type=["pdf"])
        
        st.markdown("**Eixos e Palavras-chave de Busca:**")
        st.caption("Edite a tabela abaixo ou clique na última linha vazia para adicionar novos termos/cientistas.")
        
        palavras_chave_por_categoria = {
            "1. Cientistas Históricas (Internacionais)": [
                "Marie Curie", "Curie", "Irène Joliot-Curie", "Joliot-Curie",
                "Lise Meitner", "Meitner", "Emmy Noether", "Noether",
                "Chien-Shiung Wu", "Jocelyn Bell", "Bell Burnell",
                "Vera Rubin", "Henrietta Leavitt", "Henrietta Swan Leavitt",
                "Cecilia Payne", "Cecilia Payne-Gaposchkin",
                "Rosalind Franklin", "Mileva Marić", "Mileva Maric",
                "Hipátia", "Hipatia", "Émilie du Châtelet", "du Châtelet",
                "Maria Goeppert Mayer", "Goeppert Mayer", "Donna Strickland",
                "Andrea Ghez", "Anne L'Huillier", "Katherine Johnson",
                "Mary Jackson", "Dorothy Vaughan", "Valentina Tereshkova",
                "Sally Ride", "Mae Jemison", "Fabiola Gianotti",
                "Ada Lovelace", "Hedy Lamarr", "Annie Jump Cannon",
                "Williamina Fleming", "Antonia Maury", "Maria Mitchell",
                "Caroline Herschel", "Mary Somerville", "Sophie Germain",
                "Laura Bassi", "Marietta Blau", "Ida Noddack", "Marguerite Perey",
                "Hertha Ayrton", "Katharine Burr Blodgett", "Gerty Cori"
            ],
            "2. Cientistas Brasileiras e Contemporâneas": [
                "Sônia Guimarães", "Sonia Guimaraes", "Márcia Barbosa", "Marcia Barbosa",
                "Duília de Mello", "Duilia de Mello", "Elisa Frota-Pessôa", "Elisa Frota Pessoa",
                "Neusa Amato", "Thaisa Storchi Bergmann", "Zélia Ludwig", "Zelia Ludwig",
                "Katemari Rosa", "Jaqueline Mesquita", "Carolina Nemes",
                "Belita Koiller", "Yvonne Mascarenhas", "Débora Peres Menezes",
                "Debora Peres Menezes", "Sabrina Pasterski", "Katie Bouman"
            ],
            "3. Gênero, Equidade e Representatividade": [
                "mulher", "mulheres", "feminina", "feminino", "femininas", "femininos",
                "menina", "meninas", "garota", "garotas", "gênero", "questões de gênero",
                "igualdade de gênero", "equidade de gênero", "desigualdade de gênero",
                "machismo", "sexismo", "patriarcado", "preconceito", "discriminação",
                "empoderamento", "protagonismo feminino", "representatividade",
                "pioneira", "pioneiras", "mulheres na ciência", "mulheres na física",
                "cientistas mulheres", "prêmio nobel", "laureada", "laureadas"
            ],
            "4. Papéis Científicos e Profissionais (Feminino)": [
                "cientista", "cientistas", "pesquisadora", "pesquisadoras",
                "engenheira", "engenheiras", "astrônoma", "astrônomas",
                "astrofísica", "astrofísicas", "astronauta", "cosmonauta",
                "inventora", "inventoras", "professora", "professoras",
                "aluna", "alunas", "estudante", "estudantes", "doutora",
                "autora", "autoras", "coautora", "colaboradora", "diretora",
                "tecnóloga", "programadora", "pilota", "operadora", "técnica"
            ],
            "5. Personagens Femininas Comuns (Exercícios/Exemplos)": [
                "Maria", "Ana", "Joana", "Júlia", "Julia", "Mariana", "Beatriz",
                "Camila", "Fernanda", "Gabriela", "Helena", "Laura", "Letícia",
                "Luísa", "Luisa", "Patrícia", "Paula", "Sofia", "Cláudia",
                "Carla", "Renata", "Simone", "Sandra", "Marta", "Alice",
                "Bianca", "Bruna", "Carolina", "Daniela", "Eduarda", "Flávia",
                "Isabela", "Larissa", "Amanda", "Aline", "Raquel", "Teresa"
            ]
        }
        
        categorias_lista = list(palavras_chave_por_categoria.keys())
        palavras_lista = [", ".join(palavras) for palavras in palavras_chave_por_categoria.values()]

        dados_iniciais = pd.DataFrame({
            "Categoria": categorias_lista,
            "Palavras-chave": palavras_lista
        })
        
        tabela_editavel = st.data_editor(
            dados_iniciais,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True
        )
        
    with col2:
        btn_iniciar = st.button("🚀 Processar PDF e Gerar Planilha Bruta", use_container_width=True)

        if btn_iniciar:
            if arquivo_enviado is None:
                st.warning("⚠️ Selecione um
