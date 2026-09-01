# 1. Importação correta (SEM o .py no final)
from analisepdf import analisar_pdf_para_planilha

import streamlit as st
import os

# 2. Cole o dicionário de busca aqui no app.py para a interface enxergar
dicionario_de_busca = {
    "Agro": [
        "agronegócio", "negócio agrícola", "negócio agropecuário", "agropecuária",
        "setor agropecuário", "setor agrícola", "setor rural", "cadeia produtiva",
        "produção do campo", "setor primário", "campo e indústria", "fertilizantes"
    ],
    "Agro Novas": [
        "agrotech", "bioinsumos", "carbono zero", "agricultura regenerativa",
        "fazendas verticais", "conectividade rural", "rastreabilidade"
    ],
    "Tecnologia e Inovação": [
        "máquinas agrícolas", "tratores", "implementos", "irrigação", "drones",
        "agricultura de precisão", "sensores", "telemetria"
    ],
    "Economia e Crédito": [
        "crédito rural", "financiamento", "seguro rural", "assistência técnica",
        "extensão rural", "planejamento produtivo", "pesquisa agropecuária"
    ],
    "Política e Legislação": [
        "política agrícola", "legislação ambiental", "leis de incentivo"
    ], 
    
    "Mercado e Comércio": [
        "exportação agrícola", "importação agrícola", "comércio internacional"
    ], 
    "Desafios e Tendências": [
        "mudanças climáticas", "escassez de recursos", "inovação tecnológica", "tendências de mercado"
    ],
    "Diversidade e Inclusão": [
        "inclusão social", "diversidade no campo", "equidade de gênero", "inclusão de minorias"
    ] 
}

# =========================================================================
# INTERFACE WEB COM STREAMLIT
# =========================================================================
st.set_page_config(page_title="Analisador de PDFs", layout="centered")

st.title("📄 Analisador de PDFs")
st.write("Faça o upload do livro ou documento em PDF para gerar a análise em Excel.")

# 1. SELEÇÃO DE CATEGORIAS PRÉ-DEFINIDAS
categorias_disponiveis = list(dicionario_de_busca.keys())

categorias_selecionadas = st.multiselect(
    "Quais categorias de palavras-chave você quer analisar?",
    options=categorias_disponiveis,
    default=categorias_disponiveis
)

# 2. INICIALIZAÇÃO DO DICIONÁRIO (Deve vir antes do campo de texto!)
dicionario_filtrado = {categoria: dicionario_de_busca[categoria] for categoria in categorias_selecionadas}

# 3. CAMPO PARA PALAVRAS CUSTOMIZADAS
palavras_customizadas = st.text_input("Quer buscar outras palavras específicas? Digite-as aqui separadas por vírgula:")

if palavras_customizadas:
    lista_customizada = [palavra.strip() for palavra in palavras_customizadas.split(",") if palavra.strip()]
    if lista_customizada:
        dicionario_filtrado["Palavras Customizadas"] = lista_customizada

# 4. UPLOAD DO ARQUIVO PDF
arquivo_enviado = st.file_uploader("Selecione o arquivo PDF", type=["pdf"])

if arquivo_enviado is not None:
    # Trava a execução caso não haja nenhuma categoria ou palavra cadastrada
    if len(dicionario_filtrado) == 0:
        st.warning("⚠️ Por favor, selecione ao menos uma categoria ou digite palavras-chave customizadas.")
    else:
        # Salva o arquivo temporariamente
        caminho_pdf_temp = "temp_processamento.pdf"
        with open(caminho_pdf_temp, "wb") as f:
            f.write(arquivo_enviado.getbuffer())
        
        # Botão para disparar a execução
        if st.button("Iniciar Análise"):
            caminho_excel_saida = "resultado_analise.xlsx"
            
            with st.spinner("Analisando páginas e gêneros textuais... Isso pode levar alguns segundos."):
                analisar_pdf_para_planilha(caminho_pdf_temp, dicionario_filtrado, caminho_excel_saida)
            
            st.balloons()
            st.success("Análise concluída com sucesso!")
            
            # Botão para baixar a planilha final
            with open(caminho_excel_saida, "rb") as file:
                st.download_button(
                    label="📥 Baixar Planilha Excel Formatada",
                    data=file,
                    file_name=f"{os.path.splitext(arquivo_enviado.name)[0]}_analise.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            # Limpeza do arquivo temporário
            if os.path.exists(caminho_pdf_temp):
                os.remove(caminho_pdf_temp)