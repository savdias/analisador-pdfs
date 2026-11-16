import os
import re
import fitz  # PyMuPDF
import pandas as pd
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.formatting.rule import ColorScaleRule
from concurrent.futures import ProcessPoolExecutor

# =========================================================================
# DICIONÁRIO DE BUSCA: REPRESENTATIVIDADE FEMININA NA FÍSICA (PNLD)
# =========================================================================
DICIONARIO_REPRESENTATIVIDADE_FEMININA = {
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
        "Belita Koiller", "Yvonne Mascarenhas", "Niedja", "Debora Peres Menezes",
        "Débora Peres Menezes", "Sabrina Pasterski", "Katie Bouman"
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

# =========================================================================
# FUNÇÃO DE DETECÇÃO DE GÊNERO TEXTUAL POR TRECHO / PARÁGRAFO
# =========================================================================
def identificar_genero_trecho(texto_trecho):
    """
    Analisa as marcas linguísticas, estruturais e lexicais de um trecho
    para classificar o gênero textual no contexto de livros didáticos de Física.
    """
    texto = texto_trecho.lower()

    pontuacao = {
        'Biografia / Boxe Histórico': 0,
        'Atividade / Exercício / Questão de Vestibular': 0,
        'Infográfico / Legenda de Imagem ou Gráfico': 0,
        'Notícia / Reportagem / Divulgação Científica': 0,
        'Artigo Científico / Acadêmico / Relatório': 0,
        'História em Quadrinhos / Charge / Tirinha': 0,
        'Poema / Poesia / Cordel / Letra de Música': 0,
        'Narrativa Literária (Conto / Crônica / Fábula)': 0,
        'Entrevista / Depoimento': 0,
        'Manual Técnico / Experimento / Roteiro Prático': 0,
        'Texto Legislativo / Documento Oficial': 0,
        'Texto Expositivo / Didático Geral': 0
    }

    # 1. Biografia / Boxe Histórico (Muito relevante para análise de representatividade)
    if re.search(r'\b(nasceu em|faleceu em|morreu em|sua infância|sua vida|trajetória|biografia|autobiografia|formou-se|doutorado|prêmio nobel|laureada|pioneira|contribuições|descobriu|homenagem|legado|século xix|século xx)\b', texto):
        pontuacao['Biografia / Boxe Histórico'] += 4

    # 2. Atividade / Exercício / Questão de Vestibular (Calibrado para Física)
    if re.search(r'\b(exercício|exercícios|atividade|questão|questões|responda|assinale|marque|calcule|determine|obtenha|sabendo que|considere|despreze|alternativa|enem|fuvest|unicamp|unesp|ita|ime|m/s|km/h|newton|joule|watt|volt)\b', texto):
        pontuacao['Atividade / Exercício / Questão de Vestibular'] += 4

    # 3. Infográfico / Legenda de Imagem ou Gráfico
    if re.search(r'\b(fonte:|gráfico \d|figura \d|fig\. \d|tabela \d|imagem \d|foto:|fotografia|ilustração|acervo|reprodução|legenda:)\b', texto):
        pontuacao['Infográfico / Legenda de Imagem ou Gráfico'] += 3

    # 4. Notícia / Reportagem / Divulgação Científica
    if re.search(r'\b(reportagem|notícia|jornal|manchete|jornalista|divulgado|revista|pesquisa recente|segundo especialistas|agência|superinteressante|galileu|ciência hoje|fapesp)\b', texto):
        pontuacao['Notícia / Reportagem / Divulgação Científica'] += 3

    # 5. Artigo Científico / Acadêmico / Relatório
    if re.search(r'\b(resumo|abstract|metodologia|referências bibliográficas|hipótese|amostra|conclui-se|et al\.|doi:|issn|universidade|instituto de física)\b', texto):
        pontuacao['Artigo Científico / Acadêmico / Relatório'] += 3

    # 6. História em Quadrinhos / Charge / Tirinha
    if re.search(r'\b(quadrinho|quadrinhos|hq|charge|cartum|tirinha|tirinhas|balão de fala|quadro \d+|tira \d+|mafalda|armandinho|calvin)\b', texto):
        pontuacao['História em Quadrinhos / Charge / Tirinha'] += 4

    # 7. Poema / Poesia / Cordel
    if re.search(r'\b(estrofe|verso|versos|rima|poema|poesia|cordel|soneto|lírico)\b', texto):
        pontuacao['Poema / Poesia / Cordel / Letra de Música'] += 3

    # 8. Narrativa Literária
    if re.search(r'\b(era uma vez|certo dia|de repente|personagem|narrador|disse|respondeu|perguntou|exclamou|fábula|mito|lenda)\b', texto) or '—' in texto_trecho:
        pontuacao['Narrativa Literária (Conto / Crônica / Fábula)'] += 2

    # 9. Entrevista / Depoimento
    if re.search(r'\b(entrevistado|entrevistada|entrevistador|pergunta:|resposta:|depoimento|relata que|afirma a pesquisadora)\b', texto):
        pontuacao['Entrevista / Depoimento'] += 4

    # 10. Manual Técnico / Experimento / Roteiro Prático
    if re.search(r'\b(passo \d|passo a passo|instruções|procedimento|materiais necessários|roteiro|experimento|atividade prática|mão na massa|montagem)\b', texto):
        pontuacao['Manual Técnico / Experimento / Roteiro Prático'] += 3

    # 11. Texto Legislativo / Documento Oficial
    if re.search(r'\b(artigo \d+|art\. \d+|decreto|lei nº|constituição|portaria|edital|resolução|bncc|pnld)\b', texto):
        pontuacao['Texto Legislativo / Documento Oficial'] += 4

    # 12. Texto Expositivo / Didático Geral
    if re.search(r'\b(conceito|definição|podemos definir|consiste em|caracteriza-se por|exemplo|é chamado de|fenômeno|lei de|teoria)\b', texto):
        pontuacao['Texto Expositivo / Didático Geral'] += 1

    genero_predominante = max(pontuacao, key=pontuacao.get)

    if pontuacao[genero_predominante] == 0:
        return "Texto Expositivo / Didático Geral"

    return genero_predominante

# =========================================================================
# FUNÇÃO DE LIMPEZA GERAL
# =========================================================================
def limpar_caracteres_invalidos(texto):
    if not isinstance(texto, str):
        return texto
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', texto)

# =========================================================================
# WORKER PARALELO OTIMIZADO
# =========================================================================
def _processar_lote_paginas_texto(argumentos):
    caminho_pdf, lote_paginas, mapa_palavras = argumentos
    doc = fitz.open(caminho_pdf)
    resultados_lote = []
    
    # Janela ampliada para 200 caracteres antes e depois para melhor análise qualitativa
    padroes_compilados = {}
    for palavra, categoria in mapa_palavras.items():
        padroes_compilados[palavra] = (
            categoria, 
            re.compile(r'(.{0,200}\b' + re.escape(palavra) + r'\b.{0,200})', flags=re.IGNORECASE)
        )
        
    for num_pagina in lote_paginas:
        pagina = doc[num_pagina]
        texto_pagina = pagina.get_text()
        
        if not texto_pagina: 
            continue
            
        texto_pagina_limpo = re.sub(r'\s+', ' ', texto_pagina).strip()
        texto_pagina_lower = texto_pagina_limpo.lower()
        
        for palavra, (categoria, padrao) in padroes_compilados.items():
            if palavra.lower() not in texto_pagina_lower: 
                continue
                
            for ocorrencia in padrao.finditer(texto_pagina_limpo):
                trecho = ocorrencia.group(0).strip()
                genero_trecho = identificar_genero_trecho(trecho)

                resultados_lote.append({
                    'Categoria': categoria,
                    'Palavra-Chave': palavra,
                    'Página': num_pagina + 1,
                    'Gênero Textual do Trecho': genero_trecho,
                    'Trecho Encontrado': f"... {trecho} ..."
                })
                
    doc.close()
    return resultados_lote

# =========================================================================
# FUNÇÃO PRINCIPAL DE EXTRAÇÃO E FORMATAÇÃO
# =========================================================================
def analisar_pdf_para_planilha(caminho_pdf, dicionario_palavras, caminho_saida_excel):
    print(f"Iniciando a leitura do livro didático: {caminho_pdf}")
    
    doc = fitz.open(caminho_pdf)
    total_paginas = len(doc)
    doc.close()

    mapa_palavras = {}
    vistos = set()
    for categoria, palavras in dicionario_palavras.items():
        for p in palavras:
            p_clean = p.strip()
            p_lower = p_clean.lower()
            if p_lower not in vistos:
                vistos.add(p_lower)
                mapa_palavras[p_clean] = categoria

    tamanho_lote = 30
    lotes_paginas = [list(range(i, min(i + tamanho_lote, total_paginas))) for i in range(0, total_paginas, tamanho_lote)]
    tarefas = [(caminho_pdf, lote, mapa_palavras) for lote in lotes_paginas]
    resultados = []
    
    with ProcessPoolExecutor() as executor:
        listas_de_resultados = executor.map(_processar_lote_paginas_texto, tarefas)
        for lista in listas_de_resultados:
            resultados.extend(lista)
            
    df_base = pd.DataFrame(resultados)

    if df_base.empty:
        print("Nenhuma das palavras-chave foi encontrada neste volume.")
        return

    df_base['Trecho Encontrado'] = df_base['Trecho Encontrado'].apply(limpar_caracteres_invalidos)
    df_base = df_base.drop_duplicates(subset=['Categoria', 'Palavra-Chave', 'Página', 'Trecho Encontrado']).reset_index(drop=True)
    
    # 1. ABA DE RESUMO POR CATEGORIA (Nova aba analítica para o PNLD)
    df_resumo_categorias = df_base.groupby('Categoria').agg(
        Total_de_Ocorrencias=('Palavra-Chave', 'count'),
        Termos_Distintos_Encontrados=('Palavra-Chave', 'nunique'),
        Paginas_Distintas=('Página', 'nunique'),
        Lista_de_Paginas=('Página', lambda x: ', '.join(map(str, sorted(set(x)))))
    ).reset_index().rename(columns={
        'Total_de_Ocorrencias': 'Total de Ocorrências',
        'Termos_Distintos_Encontrados': 'Termos Distintos Encontrados',
        'Paginas_Distintas': 'Qtd. de Páginas Diferentes',
        'Lista_de_Paginas': 'Páginas Onde Aparece'
    })

    # 2. ABA DE QUANTIFICAÇÃO DE PALAVRAS
    df_quantificacao = df_base.groupby(['Categoria', 'Palavra-Chave']).size().reset_index(name='Quantidade de Aparições')
    df_quantificacao = df_quantificacao.sort_values(by=['Categoria', 'Quantidade de Aparições'], ascending=[True, False])
    
    # 3. ABA DE GÊNEROS TEXTUAIS
    df_generos = df_base.groupby('Gênero Textual do Trecho').agg({
        'Gênero Textual do Trecho': 'count',
        'Página': lambda x: ', '.join(map(str, sorted(set(x))))
    }).rename(columns={
        'Gênero Textual do Trecho': 'Quantidade de Ocorrências', 
        'Página': 'Páginas Encontradas'
    }).reset_index().sort_values(by='Quantidade de Ocorrências', ascending=False)
    
    # 4. ABA DE PÁGINAS POR PALAVRA
    df_paginas = df_base.groupby(['Categoria', 'Palavra-Chave'])['Página'].apply(lambda x: ', '.join(map(str, sorted(set(x))))).reset_index(name='Páginas Encontradas')
    df_paginas = df_paginas.sort_values(by=['Categoria', 'Palavra-Chave'])
    
    # 5. ABA DE RESUMO GERAL DE PÁGINAS
    df_todas_paginas = pd.DataFrame({'Páginas com Ocorrências (Visão Geral)': sorted(df_base['Página'].unique())})

    # 6. ABA DE ANÁLISE DETALHADA POR PÁGINA
    df_agrupado_palavra = df_base.groupby(['Categoria', 'Palavra-Chave', 'Página', 'Gênero Textual do Trecho'])['Trecho Encontrado'].apply(lambda trechos: '\n\n'.join(trechos)).reset_index()
    df_agrupado_palavra['Bloco Formatado'] = df_agrupado_palavra.apply(
        lambda row: f"► {row['Palavra-Chave']} ({row['Categoria']} | Gênero: {row['Gênero Textual do Trecho']}):\n{row['Trecho Encontrado']}", axis=1
    )
    
    df_aba1 = df_agrupado_palavra.groupby('Página').agg({
        'Categoria': lambda x: ', '.join(sorted(set(x))),
        'Palavra-Chave': lambda x: ', '.join(sorted(set(x))),
        'Gênero Textual do Trecho': lambda x: ', '.join(sorted(set(x))),
        'Bloco Formatado': lambda x: '\n\n--------------------------------------------------\n\n'.join(x)
    }).reset_index()
    
    df_aba1.rename(columns={
        'Bloco Formatado': 'Trechos Encontrados', 
        'Categoria': 'Categorias na Página', 
        'Palavra-Chave': 'Palavras-Chave',
        'Gênero Textual do Trecho': 'Gêneros Identificados na Página'
    }, inplace=True)
    
    df_aba1 = df_aba1[['Página', 'Categorias na Página', 'Palavras-Chave', 'Gêneros Identificados na Página', 'Trechos Encontrados']]

    # Salvando no Excel e aplicando formatação
    with pd.ExcelWriter(caminho_saida_excel, engine='openpyxl') as writer:
        df_aba1.to_excel(writer, index=False, sheet_name='Análise Detalhada')
        df_resumo_categorias.to_excel(writer, index=False, sheet_name='Resumo por Categoria')
        df_quantificacao.to_excel(writer, index=False, sheet_name='Quantificação de Palavras')
        df_generos.to_excel(writer, index=False, sheet_name='Gêneros Textuais')
        df_paginas.to_excel(writer, index=False, sheet_name='Páginas por Palavra')
        df_todas_paginas.to_excel(writer, index=False, sheet_name='Resumo de Páginas')
        df_base.to_excel(writer, index=False, sheet_name='Base de Dados') 

        fonte_cabecalho = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
        fonte_dados = Font(name="Calibri", size=11)

        writer.sheets['Base de Dados'].sheet_state = 'hidden'

        # Estilo Aba 1: Análise Detalhada
        ws1 = writer.sheets['Análise Detalhada']
        cor_azul = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        for cell in ws1[1]:
            cell.fill = cor_azul
            cell.font = fonte_cabecalho
            cell.alignment = Alignment(horizontal="center", vertical="center")
        
        ws1.column_dimensions['A'].width = 12  
        ws1.column_dimensions['B'].width = 35  
        ws1.column_dimensions['C'].width = 30  
        ws1.column_dimensions['D'].width = 35  
        ws1.column_dimensions['E'].width = 110 
        
        for row in ws1.iter_rows(min_row=2, max_col=5, max_row=ws1.max_row):
            for cell in row:
                cell.font = fonte_dados
                cell.alignment = Alignment(
                    vertical="top", 
                    horizontal="center" if cell.column_letter == 'A' else "left",
                    wrap_text=(cell.column_letter in ['B', 'C', 'D', 'E'])
                )

        # Estilo Aba 2: Resumo por Categoria
        ws_cat = writer.sheets['Resumo por Categoria']
        cor_vinho = PatternFill(start_color="7030A0", end_color="7030A0", fill_type="solid")
        for cell in ws_cat[1]:
            cell.fill = cor_vinho
            cell.font = fonte_cabecalho
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
        ws_cat.column_dimensions['A'].width = 45
        ws_cat.column_dimensions['B'].width = 22
        ws_cat.column_dimensions['C'].width = 28
        ws_cat.column_dimensions['D'].width = 26
        ws_cat.column_dimensions['E'].width = 50
        
        for row in ws_cat.iter_rows(min_row=2, max_col=5, max_row=ws_cat.max_row):
            for cell in row:
                cell.font = fonte_dados
                cell.alignment = Alignment(
                    vertical="top", 
                    horizontal="center" if cell.column_letter in ['B', 'C', 'D'] else "left",
                    wrap_text=True
                )

        # Estilo Aba 3: Quantificação de Palavras
        ws2 = writer.sheets['Quantificação de Palavras']
        cor_verde = PatternFill(start_color="2E6B4E", end_color="2E6B4E", fill_type="solid")
        for cell in ws2[1]:
            cell.fill = cor_verde
            cell.font = fonte_cabecalho
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
        ws2.column_dimensions['A'].width = 42  
        ws2.column_dimensions['B'].width = 30  
        ws2.column_dimensions['C'].width = 25  
        
        for row in ws2.iter_rows(min_row=2, max_col=3, max_row=ws2.max_row):
            for cell in row:
                cell.font = fonte_dados
                cell.alignment = Alignment(vertical="center", horizontal="left" if cell.column_letter != 'C' else "center")

        # Estilo Aba 4: Gêneros Textuais
        ws_genero = writer.sheets['Gêneros Textuais']
        cor_azul_escuro = PatternFill(start_color="0F243E", end_color="0F243E", fill_type="solid")
        for cell in ws_genero[1]:
            cell.fill = cor_azul_escuro
            cell.font = fonte_cabecalho
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
        ws_genero.column_dimensions['A'].width = 45  
        ws_genero.column_dimensions['B'].width = 25  
        ws_genero.column_dimensions['C'].width = 50  
        
        for row in ws_genero.iter_rows(min_row=2, max_col=3, max_row=ws_genero.max_row):
            for cell in row:
                cell.font = fonte_dados
                cell.alignment = Alignment(vertical="top", horizontal="center" if cell.column_letter == 'B' else "left", wrap_text=True)

        if ws_genero.max_row >= 2:
            regra_escala_cores = ColorScaleRule(
                start_type='min', start_color='E0F2FE',
                mid_type='percentile', mid_value=50, mid_color='7DD3FC',
                end_type='max', end_color='0284C7'
            )
            ws_genero.conditional_formatting.add(f'B2:B{ws_genero.max_row}', regra_escala_cores)

        # Estilo Aba 5: Páginas por Palavra
        ws3 = writer.sheets['Páginas por Palavra']
        cor_laranja = PatternFill(start_color="D9534F", end_color="D9534F", fill_type="solid")
        for cell in ws3[1]:
            cell.fill = cor_laranja
            cell.font = fonte_cabecalho
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
        ws3.column_dimensions['A'].width = 42  
        ws3.column_dimensions['B'].width = 30  
        ws3.column_dimensions['C'].width = 50  
        
        for row in ws3.iter_rows(min_row=2, max_col=3, max_row=ws3.max_row):
            for cell in row:
                cell.font = fonte_dados
                cell.alignment = Alignment(vertical="center", horizontal="left" if cell.column_letter != 'C' else "center", wrap_text=True)
        
        # Estilo Aba 6: Resumo de Páginas
        ws4 = writer.sheets['Resumo de Páginas']
        cor_roxa = PatternFill(start_color="604A7B", end_color="604A7B", fill_type="solid")
        for cell in ws4[1]:
            cell.fill = cor_roxa
            cell.font = fonte_cabecalho
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
        ws4.column_dimensions['A'].width = 40  
        
        for row in ws4.iter_rows(min_row=2, max_col=1, max_row=ws4.max_row):
            for cell in row:
                cell.font = fonte_dados
                cell.alignment = Alignment(vertical="center", horizontal="center")
    
    print(f"Sucesso! Análise de representatividade salva em: {caminho_saida_excel}\n")

if __name__ == "__main__":
    diretorio_script = os.path.dirname(os.path.abspath(__file__))
    
    # Ajuste o intervalo range(1, 20) conforme a numeração dos volumes do PNLD na pasta "pdf"
    lista_de_livros = [os.path.join(diretorio_script, "pdf", f"livro{i}.pdf") for i in range(19, 20)]
    
    for arquivo_pdf in lista_de_livros:
        if os.path.exists(arquivo_pdf):
            nome_saida = os.path.splitext(os.path.basename(arquivo_pdf))[0] + "_representatividade_feminina.xlsx"
            caminho_saida = os.path.join(diretorio_script, nome_saida)
            analisar_pdf_para_planilha(arquivo_pdf, DICIONARIO_REPRESENTATIVIDADE_FEMININA, caminho_saida)
        else:
            print(f"Arquivo não encontrado: {arquivo_pdf}")
