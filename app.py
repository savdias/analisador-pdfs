import os
import pandas as pd
import streamlit as st
import plotly.express as px
from analisepdf import analisar_pdf_para_planilha

# Configuração da página
st.set_page_config(page_title="Analisador de PDFs", page_icon="📄", layout="centered")

# 1. Dicionário de busca de palavras-chave
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

categorias_disponiveis = list(dicionario_de_busca.keys())

# =========================================================================
# INTERFACE WEB COM STREAMLIT
# =========================================================================
st.title("📄 Analisador de PDFs e Textos")

# CAIXA 1: CONFIGURAÇÕES DE BUSCA
with st.container(border=True):
    st.subheader("⚙️ Configurações de Busca")
    
    categorias_selecionadas = st.multiselect(
        "Quais categorias de palavras-chave você quer analisar?",
        options=categorias_disponiveis,
        default=categorias_disponiveis
    )
    st.caption("💡 **Observação:** Selecione as categorias desejadas para a análise.")
    
    palavras_customizadas = st.text_input("Quer buscar outras palavras específicas? Digite separadas por vírgula:")

dicionario_filtrado = {categoria: dicionario_de_busca[categoria] for categoria in categorias_selecionadas}

if palavras_customizadas:
    lista_customizada = [palavra.strip() for palavra in palavras_customizadas.split(",") if palavra.strip()]
    if lista_customizada:
        dicionario_filtrado["Palavras Customizadas"] = lista_customizada

# CAIXA 2: UPLOAD E BOTÃO DE DISPARO
with st.container(border=True):
    st.subheader("📂 1. Processamento de PDF")
    
    col1, col2 = st.columns(2)
    
    with col1:
        arquivo_enviado = st.file_uploader("Selecione o arquivo PDF", type=["pdf"])
        
    with col2:
        st.write(" ")
        st.write(" ")
        btn_iniciar = st.button("🚀 Iniciar Análise", use_container_width=True)

# LÓGICA DE PROCESSAMENTO (PDF)
if btn_iniciar:
    if arquivo_enviado is None:
        st.warning("⚠️ Por favor, selecione um arquivo PDF antes de iniciar a análise.")
    elif len(dicionario_filtrado) == 0:
        st.warning("⚠️ Por favor, selecione ao menos uma categoria ou digite palavras-chave.")
    else:
        caminho_pdf_temp = "temp_processamento.pdf"
        caminho_excel_saida = "resultado_analise.xlsx"
        
        with open(caminho_pdf_temp, "wb") as f:
            f.write(arquivo_enviado.getbuffer())
        
        with st.spinner("Analisando páginas e gêneros textuais... Isso pode levar alguns segundos."):
            analisar_pdf_para_planilha(caminho_pdf_temp, dicionario_filtrado, caminho_excel_saida)
        
        st.balloons()
        
        # CAIXA 3: RESULTADOS E GRÁFICOS DO PDF
        with st.container(border=True):
            st.header("📊 Resultados da Análise Automática")
            
            try:
                df = pd.read_excel(caminho_excel_saida)
                
                if not df.empty:
                    st.subheader("Prévia dos Dados Encontrados")
                    st.dataframe(df.head(10))
                    
                    # Gráfico de Pizza focado nos Gêneros
                    if "Gêneros Identificados na Página" in df.columns:
                        st.subheader("Distribuição por Gêneros Textuais")
                        # Conta as ocorrências
                        contagem = df["Gêneros Identificados na Página"].value_counts().reset_index()
                        contagem.columns = ["Gênero", "Quantidade"]
                        
                        # Gera o gráfico de pizza
                        fig = px.pie(contagem, values="Quantidade", names="Gênero", hole=0.3)
                        st.plotly_chart(fig, use_container_width=True)
            except Exception as e:
                st.error(f"Erro ao exibir estatísticas: {e}")
            
            if os.path.exists(caminho_excel_saida):
                with open(caminho_excel_saida, "rb") as file:
                    st.download_button(
                        label="📥 Baixar Planilha Excel Formatada",
                        data=file,
                        file_name=f"{os.path.splitext(arquivo_enviado.name)[0]}_analise.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
        
        if os.path.exists(caminho_pdf_temp):
            os.remove(caminho_pdf_temp)

st.markdown("---")

# CAIXA 4: ANÁLISE CUSTOMIZADA DE PLANILHAS (Fica sempre visível embaixo)
with st.container(border=True):
    st.header("📈 2. Análise de Planilha Customizada")
    st.write("Faça o upload da sua planilha editada para gerar gráficos com base nas colunas.")
    
    planilha_customizada = st.file_uploader("Selecione sua planilha", type=["xlsx", "xls"])
    
    if planilha_customizada is not None:
        try:
            df_custom = pd.read_excel(planilha_customizada)
            st.success("Planilha carregada com sucesso!")
            
            colunas_disponiveis = df_custom.columns.tolist()
            
            # Agora o usuário escolhe duas colunas: uma para os rótulos e outra para os valores
            col1, col2 = st.columns(2)
            with col1:
                coluna_nomes = st.selectbox("Escolha a coluna para as CATEGORIAS (ex: Categoria):", colunas_disponiveis)
            with col2:
                coluna_valores = st.selectbox("Escolha a coluna para os VALORES (ex: Percentual do corpus):", colunas_disponiveis)
            
            if coluna_nomes and coluna_valores:
                tipo_grafico = st.radio("Escolha o tipo de gráfico:", ["Gráfico de Pizza", "Gráfico de Barras"], horizontal=True)
                
                # Remove linhas vazias se houver
                df_plot = df_custom.dropna(subset=[coluna_nomes, coluna_valores])
                
                if tipo_grafico == "Gráfico de Pizza":
                    # Usa os valores diretamente da planilha
                    fig_custom = px.pie(df_plot, values=coluna_valores, names=coluna_nomes)
                    st.plotly_chart(fig_custom, use_container_width=True)
                else:
                    # Gráfico de barras usando Plotly para ficar mais bonito
                    fig_custom = px.bar(df_plot, x=coluna_nomes, y=coluna_valores, color=coluna_nomes)
                    st.plotly_chart(fig_custom, use_container_width=True)
                    
        except Exception as e:
            st.error(f"Erro ao ler a planilha: {e}")
