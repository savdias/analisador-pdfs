        "Gêneros Textuais": generos,
        "Páginas por Palavra": paginas,
        "Resumo por Página": resumo_pagina,
        "Inventário de Livros": inventario,
        "Metodologia": metodologia,
    }


def ajustar_para_excel(df: pd.DataFrame) -> pd.DataFrame:
    """Excel admite no máximo 32.767 caracteres por célula."""
    df = df.copy()
    limite = 32000
    for coluna in df.select_dtypes(include=["object"]).columns:
        df[coluna] = df[coluna].map(
            lambda s: s[:limite] + " [CORTADO PARA EXCEL — CONFERIR PDF]"
            if isinstance(s, str) and len(s) > limite else s
        )
    return df


def exportar_excel(caminho: Path, df_oc: pd.DataFrame, df_qual: pd.DataFrame, inventario: pd.DataFrame) -> None:
    abas = preparar_abas(df_oc, df_qual, inventario)
    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        for nome, tabela in abas.items():
            ajustar_para_excel(tabela).to_excel(writer, sheet_name=nome, index=False)
            ws = writer.sheets[nome]
            ws.freeze_panes = "C2" if nome in ("Triagem Qualitativa", "Ocorrências") else "A2"
            ws.auto_filter.ref = ws.dimensions
            ws.sheet_view.showGridLines = False
            ws.row_dimensions[1].height = 28
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            for celula in ws[1]:
                celula.fill = PatternFill("solid", fgColor="243B60")
                celula.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                celula.alignment = Alignment(vertical="center", wrap_text=True)
            for colunas in ws.columns:
                cabecalho = str(colunas[0].value or "")
                letra = colunas[0].column_letter
                largura = 17
                if "ID do bloco" in cabecalho:
                    largura = 28
                elif any(t in cabecalho.lower() for t in ["trecho", "uc textual", "interpretação", "orientação"]):
                    largura = 82
                elif any(t in cabecalho.lower() for t in ["categoria", "gênero", "indício", "posição de sujeito", "modo de apresentação"]):
                    largura = 38
                elif any(t in cabecalho.lower() for t in ["palavra", "nomes de cientistas", "campo"]):
                    largura = 32
                elif any(t in cabecalho.lower() for t in ["página", "quantidade", "blocos", "correspondência"]):
                    largura = 19
                ws.column_dimensions[letra].width = largura
            for linha in ws.iter_rows(min_row=2):
                for celula in linha:
                    celula.font = Font(name="Calibri", size=10)
                    celula.alignment = Alignment(vertical="top", wrap_text=True)
                if linha[0].row <= 301:
                    ws.row_dimensions[linha[0].row].height = 46 if nome in ("Triagem Qualitativa", "Análise Detalhada") else 30

            if nome == "Categorias de Busca" and ws.max_row > 2:
                ws.conditional_formatting.add(
                    f"B2:B{ws.max_row}",
                    ColorScaleRule(start_type="min", start_color="E8EFF9",
                                   mid_type="percentile", mid_value=50, mid_color="91B8DE",
                                   end_type="max", end_color="2866A8"),
                )


def executar() -> None:
    pasta_script = Path(__file__).resolve().parent
    pasta_pdf = pasta_script / "pdf"
    pasta_saida = pasta_script / "resultados_pnld"
    pasta_saida.mkdir(exist_ok=True)

    arquivos_pdf = sorted(pasta_pdf.glob("*.pdf")) if pasta_pdf.is_dir() else []
    if not arquivos_pdf:
        print(f"Nenhum arquivo PDF em: {pasta_pdf}")
        print("Crie a pasta 'pdf' ao lado do script e coloque nela os livros de Física.")
        return

    ocorrencias_gerais = []
    triagens_gerais = []
    inventario_livros = []
    for arquivo in arquivos_pdf:
        print(f"Lendo {arquivo.name} ...")
        try:
            df_oc, df_qual, total_paginas = extrair_livro(arquivo)
        except Exception as exc:
            print(f"ERRO ao ler {arquivo.name}: {exc}")
            continue
        inventario_livros.append({
            "Livro": arquivo.stem,
            "Arquivo PDF": arquivo.name,
            "Páginas PDF (total)": total_paginas,
            "Blocos textuais candidatos": len(df_qual),
            "Correspondências textuais": len(df_oc),
            "Editora / coleção (preencher)": "",
            "Edição / ano (preencher)": "",
            "PNLD / ciclo (preencher)": "",
        })
        inventario_um = pd.DataFrame([inventario_livros[-1]])
        destino = pasta_saida / f"{arquivo.stem}_representatividade_feminina.xlsx"
        exportar_excel(destino, df_oc, df_qual, inventario_um)
        print(f"  {len(df_qual)} blocos candidatos; {len(df_oc)} correspondências; salvo: {destino.name}")
        ocorrencias_gerais.append(df_oc)
        triagens_gerais.append(df_qual)

    if not inventario_livros:
        print("Nenhum PDF pôde ser processado.")
        return
    df_oc_total = pd.concat(ocorrencias_gerais, ignore_index=True)
    df_qual_total = pd.concat(triagens_gerais, ignore_index=True)
    destino_geral = pasta_saida / "CONSOLIDADO_representatividade_feminina_PNLD.xlsx"
    exportar_excel(destino_geral, df_oc_total, df_qual_total, pd.DataFrame(inventario_livros))
    print(f"Consolidado salvo em: {destino_geral}")
    print("ATENÇÃO: fazer validação qualitativa e inspeção visual das páginas.")


if __name__ == "__main__":
    executar()
