import os
import pandas as pd
import streamlit as st
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

# Define as categorias disponíveis antes de construir a interface
categorias_disponiveis = list(dicionario_de_busca.keys())

# =========================================================================
# INTERFACE WEB COM STREAMLIT
# =========================================================================
st.title("📄 Analisador de PDFs")

# CAIXA 1: CONFIGURAÇÕES DE BUSCA
with st.container(border=True):
    st.subheader("⚙️ Configurações de Busca")
    
    categorias_selecionadas = st.multiselect(
    "Quais categorias de palavras-chave você quer analisar?",
    options=categorias_disponiveis,
    default=categorias_disponiveis
)
st.caption("💡 **Observação:** Você pode selecionar mais de uma categoria ou remover as que não desejar. Recomendado para uma busca mais simplificada")

palavras_customizadas = st.text_input("Quer buscar outras palavras específicas? Digite separadas por vírgula:")

# Processamento das categorias e palavras do usuário
dicionario_filtrado = {categoria: dicionario_de_busca[categoria] for categoria in categorias_selecionadas}

if palavras_customizadas:
    lista_customizada = [palavra.strip() for palavra in palavras_customizadas.split(",") if palavra.strip()]
    if lista_customizada:
        dicionario_filtrado["Palavras Customizadas"] = lista_customizada

# CAIXA 2: UPLOAD E BOTÃO DE DISPARO
with st.container(border=True):
    st.subheader("📂 Processamento")
    
    col1, col2 = st.columns(2)
    
    with col1:
        arquivo_enviado = st.file_uploader("Selecione o arquivo PDF", type=["pdf"])
        
    with col2:
        st.write(" ")
        st.write(" ")
        btn_iniciar = st.button("🚀 Iniciar Análise", use_container_width=True)

# LÓGICA DE PROCESSAMENTO E EXIBIÇÃO DE RESULTADOS
if btn_iniciar:
    if arquivo_enviado is None:
        st.warning("⚠️ Por favor, selecione um arquivo PDF antes de iniciar a análise.")
    elif len(dicionario_filtrado) == 0:
        st.warning("⚠️ Por favor, selecione ao menos uma categoria ou digite palavras-chave.")
    else:
        caminho_pdf_temp = "temp_processamento.pdf"
        caminho_excel_saida = "resultado_analise.xlsx"
        
        # Salva arquivo temporário
        with open(caminho_pdf_temp, "wb") as f:
            f.write(arquivo_enviado.getbuffer())
        
        with st.spinner("Analisando páginas e gêneros textuais... Isso pode levar alguns segundos."):
            analisar_pdf_para_planilha(caminho_pdf_temp, dicionario_filtrado, caminho_excel_saida)
        
        st.balloons()
        st.success("Análise concluída com sucesso!")
        
        # CAIXA 3: RESULTADOS E GRÁFICOS
        with st.container(border=True):
            st.header("📊 Resultados Estatísticos")
            
            try:
                df = pd.read_excel(caminho_excel_saida)
                
                if not df.empty:
                    st.subheader("Prévia dos Dados Encontrados")
                    st.dataframe(df.head(10))
                    
                    colunas_da_planilha = df.columns.tolist()
                    
                    # Identifica a coluna correta para montar o gráfico de barras
                    coluna_alvo = None
                    for col in ["Categoria", "Gênero", "Palavra-chave", "Termo"]:
                        if col in colunas_da_planilha:
                            coluna_alvo = col
                            break
                            
                    if coluna_alvo:
                        st.subheader(f"Contagem por {coluna_alvo}")
                        contagem = df[coluna_alvo].value_counts()
                        st.bar_chart(contagem)
                    else:
                        st.info("Planilha gerada com sucesso!")
            except Exception as e:
                st.error(f"Erro ao exibir estatísticas: {e}")
            
            # Botão de Download dentro da caixa de resultados
            if os.path.exists(caminho_excel_saida):
                with open(caminho_excel_saida, "rb") as file:
                    st.download_button(
                        label="📥 Baixar Planilha Excel Formatada",
                        data=file,
                        file_name=f"{os.path.splitext(arquivo_enviado.name)[0]}_analise.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
        
        # Limpeza do arquivo temporário
        if os.path.exists(caminho_pdf_temp):
            os.remove(caminho_pdf_temp)
