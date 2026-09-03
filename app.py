import os
import io
import pandas as pd
import streamlit as st
import plotly.express as px
import requests
from analisepdf import analisar_pdf_para_planilha
from openpyxl.styles import Alignment, Font

st.set_page_config(page_title="Analisador de PDFs e Categorizador Local (Ollama)", page_icon="📄", layout="wide")

st.title("📄 Analisador de PDFs & Categorizador Inteligente com Ollama")

# ==========================================
# MENU LATERAL: DADOS DO MATERIAL
# ==========================================
st.sidebar.header("📚 1. Dados do Material")
st.sidebar.info("Preencha estes dados antes de exportar a planilha final.")

meta_livro = st.sidebar.text_input("Nome do Livro ou Coleção:")
meta_editora = st.sidebar.text_input("Editora:")
meta_volume = st.sidebar.text_input("Volume:")
meta_etapa = st.sidebar.text_input("Etapa de ensino:")
meta_ano = st.sidebar.text_input("Ano ou série:")
meta_disciplina = st.sidebar.text_input("Disciplina:")
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
    col_palavra = "Palavras-Chave" if "Palavras-Chave" in df_resultado.columns else "Categorias" if "Categorias" in df_resultado.columns else df_resultado.columns[1]
    col_trecho = next((col for col in df_resultado.columns if "Trecho" in col), df_resultado.columns[2])
    
    col_classif_ia = "Classificação IA" if "Classificação IA" in df_resultado.columns else ""
    col_cat_ia = "Categoria Principal IA" if "Categoria Principal IA" in df_resultado.columns else ""
    col_genero_ia = "Gênero Textual" if "Gênero Textual" in df_resultado.columns else ""
    col_just_ia = "Justificativa IA" if "Justificativa IA" in df_resultado.columns else ""

    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        
        # --- ABA 1: 00_COMECE_AQUI ---
        df_aba1 = pd.DataFrame({
            "Dado do livro": ["Nome do Livro ou Coleção", "Editora", "Volume", "Etapa de ensino", "Ano ou série", "Disciplina", "Observações gerais"],
            "Preenchimento": [metadados["livro"], metadados["editora"], metadados["volume"], metadados["etapa"], metadados["ano"], metadados["disciplina"], metadados["observacao"]]
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
        df_aba2["Observações"] = "" # Fica em branco para preenchimento manual no Excel
        
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
        
        df_aba5.rename(columns={"Palavra-chave localizada": "Palavra ou expressão", "Codigos_URs": "Código(s) das URs"}, inplace=True)
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
st.header("1. Extração de Trechos do PDF")

with st.container(border=True):
    col1, col2 = st.columns([1, 2])
    with col1:
        arquivo_enviado = st.file_uploader("Selecione o arquivo PDF", type=["pdf"])
        
        st.markdown("**Categorias e Palavras-chave de Busca:**")
        st.caption("Edite a tabela abaixo. Clique na última linha vazia para adicionar novas categorias.")
        
        # Cria os dados padrão divididos em colunas
        dados_iniciais = pd.DataFrame({
            "Categoria": ["Agro", "Tecnologia e Inovação", "Mineração e Extração", "Sustentabilidade"],
            "Palavras-chave": [
                "agronegócio, agropecuária, setor agrícola, campo",
                "tecnologia, inovação, drones, automação",
                "mineração, garimpo, extração, jazidas",
                "recuperação ambiental, sustentabilidade, reflorestamento"
            ]
        })
        
        # Renderiza a tabela editável na interface
        tabela_editavel = st.data_editor(
            dados_iniciais,
            num_rows="dynamic", # Isso é o que permite adicionar/deletar linhas
            use_container_width=True,
            hide_index=True
        )
        
    with col2:
        btn_iniciar = st.button("🚀 Processar PDF e Gerar Planilha Bruta", use_container_width=True)

        if btn_iniciar:
            if arquivo_enviado is None:
                st.warning("⚠️ Selecione um arquivo PDF.")
            else:
                # Transforma a tabela que o usuário editou de volta para o formato que a IA entende
                dicionario_filtrado = {}
                for index, row in tabela_editavel.iterrows():
                    cat = str(row["Categoria"]).strip()
                    palavras = str(row["Palavras-chave"]).strip()
                    
                    # Ignora linhas que foram deixadas em branco acidentalmente
                    if cat and cat.lower() != 'nan' and palavras and palavras.lower() != 'nan':
                        lista_palavras = [p.strip() for p in palavras.split(',') if p.strip()]
                        if lista_palavras:
                            dicionario_filtrado[cat] = lista_palavras

                if not dicionario_filtrado:
                    st.error("❌ Nenhuma categoria válida encontrada. Preencha a tabela corretamente.")
                else:
                    caminho_pdf_temp = "temp_processamento.pdf"
                    caminho_excel_saida = "resultado_analise.xlsx"
                    
                    with open(caminho_pdf_temp, "wb") as f:
                        f.write(arquivo_enviado.getbuffer())
                    
                    with st.spinner("Extraindo trechos do PDF..."):
                        analisar_pdf_para_planilha(caminho_pdf_temp, dicionario_filtrado, caminho_excel_saida)
                    
                    st.success("Extração concluída!")
                    if os.path.exists(caminho_excel_saida):
                        with open(caminho_excel_saida, "rb") as file:
                            st.download_button(
                                label="📥 Baixar Planilha de Trechos Extraídos (.xlsx)",
                                data=file,
                                file_name=f"{os.path.splitext(arquivo_enviado.name)[0]}_trechos.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
# -------------------------------------------------------------------------
# SEÇÃO 2: CATEGORIZAÇÃO VIA OLLAMA
# -------------------------------------------------------------------------
st.header("🦙 2. Categorização Inteligente Local (Ollama)")

with st.container(border=True):
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        modelo_ollama = st.text_input("Modelo do Ollama:", value="llama3.2")
    with col_c2:
        host_ollama = st.text_input("Endereço do Servidor Ollama:", value="http://localhost:11434")
    
    planilha_para_ia = st.file_uploader("Selecione a planilha com os trechos extraídos para categorizar", type=["xlsx", "xls"])
    
    if planilha_para_ia is not None:
        df_ia = pd.read_excel(planilha_para_ia)
        colunas = df_ia.columns.tolist()
        coluna_trechos = st.selectbox("Selecione a coluna que contém os TRECHOS de texto:", colunas)
        
        st.subheader("⚙️ Defina as Categorias e o Prompt")
        
        legenda_padrao = (
            "A - Tema Central do Agronegócio/Cadeia Agropecuária\n"
            "B - Cadeia Agropecuária Indispensável ao Sentido\n"
            "C - Proximidade Temática/Periférico\n"
            "D - Sem Relação Demonstrável\n"
            "E - Ilegível ou Incompleto"
        )
        
        categorias_ia = st.text_area("Legenda de Classificação (A-E):", legenda_padrao, height=130)
        
        instrucao_padrao = (
            "Análise o trecho fornecido. Você deve responder OBRIGATORIAMENTE no formato separado por '|':\n"
            "Classificação (A, B, C, D ou E) | Categoria Principal | Gênero Textual (ex: Texto Didático, Infográfico, Exercício, Mapa) | Breve Justificativa"
        )
        prompt_instrucao = st.text_area("Instruções para a IA:", instrucao_padrao, height=110)
        
        if st.button("🧠 Classificar Trechos com Ollama", type="primary"):
            if not host_ollama.strip() or "seu-link-aqui" in host_ollama:
                st.error("❌ Por favor, insira o endereço válido do Ollama.")
            else:
                progresso = st.progress(0)
                status_text = st.empty()
                
                classificacoes, categorias, generos, justificativas = [], [], [], []
                total_linhas = len(df_ia)
                
                url_base = host_ollama.strip().rstrip('/')
                endpoint = f"{url_base}/api/generate"
                
                headers = {
                    "ngrok-skip-browser-warning": "true",
                    "User-Agent": "Mozilla/5.0",
                    "Content-Type": "application/json"
                }
                
                for index, row in df_ia.iterrows():
                    trecho = str(row[coluna_trechos])
                    prompt_completo = f"{prompt_instrucao}\n\nLegenda:\n{categorias_ia}\n\nTrecho:\n\"{trecho}\""
                    payload = {"model": modelo_ollama, "prompt": prompt_completo, "stream": False}
                    
                    try:
                        response = requests.post(endpoint, json=payload, headers=headers, timeout=60)
                        response.raise_for_status()
                        resposta_texto = response.json().get("response", "").strip()
                        
                        partes = [p.strip() for p in resposta_texto.split("|")]
                        
                        if len(partes) >= 4:
                            classificacoes.append(partes[0][0].upper() if partes[0] else "E")
                            categorias.append(partes[1])
                            generos.append(partes[2])
                            justificativas.append(partes[3])
                        elif len(partes) == 2:
                            classificacoes.append(partes[0][0].upper() if partes[0] else "E")
                            categorias.append("Geral")
                            generos.append("Texto Didático")
                            justificativas.append(partes[1])
                        else:
                            classificacoes.append("E")
                            categorias.append("Geral")
                            generos.append("Texto Didático")
                            justificativas.append(resposta_texto)
                    except Exception as e:
                        classificacoes.append("Erro")
                        categorias.append("Erro")
                        generos.append("Erro")
                        justificativas.append(str(e))
                    
                    progresso.progress((index + 1) / total_linhas)
                    status_text.text(f"Analisando trecho {index + 1} de {total_linhas}...")
                
                df_ia["Classificação IA"] = classificacoes
                df_ia["Categoria Principal IA"] = categorias
                df_ia["Gênero Textual"] = generos
                df_ia["Justificativa IA"] = justificativas
                
                st.session_state["df_categorizado"] = df_ia
                st.success("✅ Categorização concluída!")

st.markdown("---")

# -------------------------------------------------------------------------
# SEÇÃO 3: GRÁFICOS
# -------------------------------------------------------------------------
st.header("📈 3. Visualização de Gráficos")

if "df_categorizado" in st.session_state:
    df_plot = st.session_state["df_categorizado"]
else:
    planilha_grafico = st.file_uploader("Ou faça upload de uma planilha já categorizada para ver os gráficos", type=["xlsx", "xls"])
    df_plot = pd.read_excel(planilha_grafico) if planilha_grafico is not None else None

if df_plot is not None:
    colunas_plot = df_plot.columns.tolist()
    col1, col2 = st.columns(2)
    with col1:
        idx_default = colunas_plot.index("Classificação IA") if "Classificação IA" in colunas_plot else 0
        coluna_cat = st.selectbox("Coluna das Categorias:", colunas_plot, index=idx_default)
    with col2:
        tipo_grafico = st.radio("Tipo de Gráfico:", ["Gráfico de Pizza", "Gráfico de Barras", "Diagrama (Funil)"], horizontal=True)
    
    contagem = df_plot[coluna_cat].value_counts().reset_index()
    contagem.columns = [coluna_cat, "Quantidade"]
    
    if tipo_grafico == "Gráfico de Pizza":
        fig = px.pie(contagem, values="Quantidade", names=coluna_cat, title="Distribuição por Classificação")
    elif tipo_grafico == "Gráfico de Barras":
        fig = px.bar(contagem, x=coluna_cat, y="Quantidade", color=coluna_cat, title="Quantidade por Classificação")
    else:
        fig = px.funnel(contagem, x="Quantidade", y=coluna_cat, color=coluna_cat, title="Diagrama de Funil de Classificações")
    
    st.plotly_chart(fig, use_container_width=True)

# -------------------------------------------------------------------------
# SEÇÃO 4: DOWNLOAD FINAL COM 5 ABAS
# -------------------------------------------------------------------------
if df_plot is not None:
    st.markdown("---")
    st.header("✅ 4. Exportar Planilha Final (5 Abas)")
    st.info("💡 Preencha os Dados do Material no menu lateral esquerdo antes de baixar!")
    
    excel_final = converter_para_excel(df_plot, metadados_livro)
    
    st.download_button(
        label="📥 BAIXAR PLANILHA FINAL FORMATADA (5 ABAS)",
        data=excel_final,
        file_name="Analise_Completa_Formatada.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
