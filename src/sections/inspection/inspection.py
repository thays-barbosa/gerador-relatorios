from docx import Document
from docx.shared import Pt, Cm 
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import adicionar_titulo_secao, adicionar_paragrafo_justificado, adicionar_titulo_quadro 
import pandas as pd


def gerar_secao_fiscalizacao(doc: Document, row, nao_conformidades_df):
    """
    Gera a seção '4. FISCALIZAÇÃO' e a tabela de Não Conformidades (Quadro 1), 
    seguindo o formato visual do relatório.
    """

    # 1. Título Principal
    adicionar_titulo_secao(doc, "4. FISCALIZAÇÃO")
    doc.add_paragraph() 
    
    # 2. Primeiro Parágrafo (Equipe e Datas) - Nomes em Negrito, Matrículas SEM Negrito
    par_equipe = doc.add_paragraph()
    par_equipe.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    par_equipe.add_run("As ações de fiscalização foram realizadas pela equipe formada pelos Analistas de Regulação ")
    
    # Nome 1 (Em negrito)
    par_equipe.add_run("Alcides Vieira de Azevedo Bezerra").bold = True
    # Matrícula 1 (SEM negrito)
    par_equipe.add_run(", matrícula 40672015/01") 
    par_equipe.add_run(" e ")
    
    # Nome 2 (Em negrito)
    par_equipe.add_run("Enildo Manoel da Silva Júnior").bold = True
    # Matrícula 2 (SEM negrito)
    par_equipe.add_run(", matrícula nº 1796500/02")
    
    par_equipe.add_run(", nos dias 22 de setembro, na cidade de Garanhuns; 24 de setembro, em Petrolina; 25 de setembro, em Caruaru; e 30 de setembro de 2025 em Recife (TIP).")


    # 3. Segundo Parágrafo (Introdução às Não Conformidades) - Menção ao Quadro 1 EM negrito
    par_nc_intro = doc.add_paragraph()
    par_nc_intro.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    par_nc_intro.add_run("As Não Conformidades constatadas estão relacionadas ao ")
    par_nc_intro.add_run("Programa de Manutenção dos Terminais Rodoviários").bold = True
    par_nc_intro.add_run(", Anexo V do Contrato de Concessão, conforme descritas no ")
    
    # ✅ CORREÇÃO: Quadro 1 EM negrito
    run_quadro1 = par_nc_intro.add_run("Quadro 1")
    run_quadro1.bold = True
    
    par_nc_intro.add_run(", a seguir, com indicação dos respectivos registros fotográficos no ")
    par_nc_intro.add_run("Apêndice 1").bold = True
    par_nc_intro.add_run(".")

    # Adiciona espaço para separar o parágrafo do quadro
    doc.add_paragraph() 

    # 4. Título do Quadro 
    adicionar_titulo_quadro(
        doc, 
        "Quadro 1 – Não Conformidades por Terminal Rodoviário de Passageiros", 
        negrito=True
    )

    # 5. Lógica da Tabela de Não Conformidades (Quadro 1)
    
    id_fisc = row["ID da Fiscalização"]

    nc_fisc = nao_conformidades_df[
        nao_conformidades_df["ID da Fiscalização"] == id_fisc
    ]

    if nc_fisc.empty:
        doc.add_paragraph("Nenhuma não conformidade registrada.")
        return

    if "Terminal" not in nc_fisc.columns or "Nº" not in nc_fisc.columns:
        doc.add_paragraph("⚠️ Colunas obrigatórias não encontradas na planilha.")
        return

    # Criar tabela
    tabela = doc.add_table(rows=1, cols=2)
    tabela.style = "Table Grid"
    tabela.alignment = WD_TABLE_ALIGNMENT.LEFT
    
    # Ajuste das larguras
    tabela.columns[0].width = Cm(4.5)  
    tabela.columns[1].width = Cm(12.5) 
    

    # Cabeçalhos
    cabecalho = tabela.rows[0].cells
    cabecalho[0].text = "TERMINAL"
    cabecalho[1].text = "NÃO CONFORMIDADE"

    # Estilo cabeçalhos (Manter em negrito e tamanho 11)
    for cell in cabecalho:
        for par in cell.paragraphs:
            run = par.runs[0]
            run.bold = True
            run.font.size = Pt(11)

    # Agrupar por Terminal
    for terminal, grupo in nc_fisc.groupby("Terminal"):
        grupo = grupo.sort_values(by="Nº")

        # Extrair sigla
        if "(" in terminal and ")" in terminal:
            sigla_terminal = terminal.split("(")[-1].replace(")", "").strip()
        else:
            sigla_terminal = terminal[:3].upper()

        # Remover "TERMINAL DE"/"TERMINAL DO" e deixar maiúsculo
        nome_terminal = terminal.upper()
        nome_terminal = nome_terminal.replace("TERMINAL DE ", "").replace("TERMINAL DO ", "").strip()

        num_nc = 1  # contador sequencial para NCs

        for idx, (_, linha) in enumerate(grupo.iterrows()):
            row_cells = tabela.add_row().cells

            # Primeira linha do grupo: título do terminal
            if idx == 0:
                run_terminal = row_cells[0].paragraphs[0].add_run(nome_terminal)
                run_terminal.bold = True
                run_terminal.font.size = Pt(11)
            else:
                row_cells[0].text = ""

            # Coluna da NC
            descricao = linha["Não Conformidade"].strip()

            paragrafo_nc = row_cells[1].paragraphs[0]
            run_titulo = paragrafo_nc.add_run(f"{sigla_terminal} {num_nc}")
            run_titulo.bold = True
            run_titulo.font.size = Pt(11)

            run_desc = paragrafo_nc.add_run(f"  {descricao}")
            run_desc.font.size = Pt(11)

            num_nc += 1

    # Adiciona um espaço para separar o Quadro 1 do próximo conteúdo
    doc.add_paragraph()