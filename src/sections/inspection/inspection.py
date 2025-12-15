# sections/inspection/inspection.py - MODIFICADO E CORRIGIDO

from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL # Importar WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor # Importação adicional útil para formatação de tabela
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml # Para aplicar o sombreamento (shading)
from utils import (
    adicionar_titulo_secao, 
    adicionar_paragrafo_justificado, 
    adicionar_titulo_quadro, 
    remover_espacamento_paragrafo,
    extrair_cidade,
    padronizar_processo # Assume que esta função converte para o formato do Excel (CTR 04/2025)
) 
import pandas as pd
import re 


# --- FUNÇÃO AUXILIAR PARA APLICAR SOMBREADO (SHADING) ---
def apply_shading(cell, color_hex="D9D9D9"): # Cor cinza claro (D9D9D9 é um tom comum)
    """Aplica cor de fundo (shading) a uma célula."""
    shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), color_hex))
    cell._tc.get_or_add_tcPr().append(shading_elm)

# --- FUNÇÃO PRINCIPAL ---

def gerar_secao_fiscalizacao(doc: Document, row, nao_conformidades_df):
    """
    Gera a seção '4. FISCALIZAÇÃO' e a tabela de Não Conformidades (Quadro 1), 
    com informações dinâmicas de equipe e locais.
    """

    # 1. Título Principal
    adicionar_titulo_secao(doc, "4. FISCALIZAÇÃO")

    # 2. Preparação dos dados dinâmicos (Equipe e Locais) - Sem alterações aqui

    # A. Dados da Equipe (row = aba Fiscalizações)
    nomes_responsaveis = [n.strip() for n in str(row["Pessoal Responsável"]).split(';') if n.strip()]
    matriculas_responsaveis = [m.strip() for m in str(row["Matrícula do Pessoal Responsável"]).split(';') if m.strip()]

    # Garante que o número de matrículas seja igual ao número de nomes (preenche com vazio se faltar)
    matriculas_ajustadas = matriculas_responsaveis + [''] * (len(nomes_responsaveis) - len(matriculas_responsaveis))
    responsaveis_info = list(zip(nomes_responsaveis, matriculas_ajustadas))

    # B. Dados dos Terminais e Cidades (row = aba Fiscalizações)
    terminais_originais = [t.strip() for t in str(row["Terminais"]).split(';') if t.strip()]

    locais_formatados = []

    for terminal_original in terminais_originais:
        cidade_limpa = extrair_cidade(terminal_original)

        if "recife (tip)" in terminal_original.lower():
            locais_formatados.append(f"em Recife (TIP)")
        elif cidade_limpa:
            locais_formatados.append(f"na cidade de {cidade_limpa}")

    # 3. Construção do parágrafo dinâmico (Equipe e Locais) - Sem alterações aqui

    par_equipe = doc.add_paragraph()
    remover_espacamento_paragrafo(par_equipe) 
    par_equipe.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    par_equipe.add_run("As ações de fiscalização foram realizadas pela equipe formada pelos Analistas de Regulação ")

    for i, (nome, matricula) in enumerate(responsaveis_info):
        par_equipe.add_run(nome).bold = True

        # Adiciona matrícula (se houver)
        if matricula:
            par_equipe.add_run(f" (matrícula nº {matricula})") 

        # Conectivos (vírgula, 'e')
        if i < len(responsaveis_info) - 2:
            par_equipe.add_run(", ")
        elif i == len(responsaveis_info) - 2:
            par_equipe.add_run(" e ")

    texto_locais = "; ".join(locais_formatados)
    par_equipe.add_run(f", {texto_locais}.") 

    # 4. Segundo Parágrafo (Introdução às Não Conformidades) - Sem alterações aqui
    doc.add_paragraph() # Adiciona espaço
    par_nc_intro = doc.add_paragraph()
    remover_espacamento_paragrafo(par_nc_intro)
    par_nc_intro.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    par_nc_intro.add_run("As Não Conformidades constatadas estão relacionadas ao ")
    par_nc_intro.add_run("Programa de Manutenção dos Terminais Rodoviários").bold = True
    par_nc_intro.add_run(", Anexo V do Contrato de Concessão, conforme descritas no ")
 
    run_quadro1 = par_nc_intro.add_run("Quadro 1")
    run_quadro1.bold = True

    par_nc_intro.add_run(", a seguir, com indicação dos respectivos registros fotográficos no ")
    par_nc_intro.add_run("Apêndice 1").bold = True
    par_nc_intro.add_run(".")

    # Adiciona espaço para separar o parágrafo do quadro
    doc.add_paragraph() 

    # 5. Título do Quadro 
    adicionar_titulo_quadro(
        doc, 
        "Quadro 1 – Não Conformidades por Terminal Rodoviário de Passageiros", 
        negrito=True
    )

    # 6. Lógica da Tabela de Não Conformidades (Quadro 1)

    # 🚨 MUDANÇA ESSENCIAL: USA A COLUNA PROCESSO, que é o seu ID principal
    processo_filtragem = str(row["PROCESSO"]).strip()

    # Verifica se a coluna 'PROCESSO' existe antes de filtrar
    if "PROCESSO" not in nao_conformidades_df.columns:
        adicionar_paragrafo_justificado(doc, "⚠️ Coluna 'PROCESSO' não encontrada na planilha de Não-conformidades. Não é possível filtrar os dados.")
        return

    # 🚨 CORREÇÃO DE FILTRAGEM: Limpar espaços em branco na coluna 'PROCESSO'
    # Esta coluna já foi limpa no report.py, mas repetimos para garantir robustez
    nao_conformidades_df['PROCESSO_LIMPO'] = nao_conformidades_df['PROCESSO'].astype(str).str.strip()

    # Filtra o DataFrame pelo PROCESSO LIMPO
    nc_fisc = nao_conformidades_df[
        nao_conformidades_df["PROCESSO_LIMPO"] == processo_filtragem
    ].copy()

    if nc_fisc.empty:
        # Inclui o valor buscado na mensagem de erro para debug
        adicionar_paragrafo_justificado(doc, f"Nenhuma não conformidade registrada para o processo buscado: **{processo_filtragem}**. Verifique a planilha.")
        return
        
    total_nc = len(nc_fisc)

    # 7. VALIDAÇÃO DAS COLUNAS NECESSÁRIAS
    colunas_necessarias = ["TERMINAL RODOVIÁRIO", "ID", "DESCRIÇÃO", "REGISTROS FOTOGRÁFICOS", "FUNDAMENTO DA INFRAÇÃO", "DETERMINAÇÃO"]

    for col in colunas_necessarias:
        # Apenas um aviso, o código tentará prosseguir
        if col not in nc_fisc.columns:
            print(f"⚠️ Coluna '{col}' não encontrada na planilha de Não-conformidades.")
            # É crucial garantir que as colunas existam para o .get() funcionar abaixo

    # 8. CRIAÇÃO DA TABELA NO WORD

    tabela = doc.add_table(rows=1, cols=6) 
    tabela.style = "Table Grid"
    tabela.alignment = WD_TABLE_ALIGNMENT.CENTER 

    # Ajuste das larguras (aproximadamente 17 cm total)
    tabela.columns[0].width = Cm(1.2) # TRP
    tabela.columns[1].width = Cm(2.5) # IDENTIFICAÇÃO
    tabela.columns[2].width = Cm(4.5) # DESCRIÇÃO
    tabela.columns[3].width = Cm(2.5) # REGISTRO FOTOGRÁFICO
    tabela.columns[4].width = Cm(3.8) # FUNDAMENTO DA INFRAÇÃO
    tabela.columns[5].width = Cm(2.5) # DETERMINAÇÃO


    # Cabeçalhos
    cabecalhos_nomes = ["TRP", "IDENTIFICAÇÃO", "DESCRIÇÃO", "REGISTRO FOTOGRÁFICO", "FUNDAMENTO DA INFRAÇÃO (ANEXO V CONTRATO DE CONCESSÃO)", "DETERMINAÇÃO"]

    cabecalho_cells = tabela.rows[0].cells

    for i, nome in enumerate(cabecalhos_nomes):
        cell = cabecalho_cells[i]
        # 🟢 APLICAR SOMBREADO CINZA AO CABEÇALHO
        apply_shading(cell) 

        cell.text = nome
        # Aplica estilo: centralizado, negrito, tamanho 9/10, compactado
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER # Alinhamento vertical centralizado

        for par in cell.paragraphs:
            remover_espacamento_paragrafo(par)
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in par.runs:
                run.bold = True
                run.font.size = Pt(9) 


    # 9. POPULANDO AS LINHAS (Agrupamento por TRP para Mesclar Células)

    # Ordena para garantir que os terminais fiquem agrupados corretamente
    nc_fisc = nc_fisc.sort_values(by=["TERMINAL RODOVIÁRIO", "ID"])

    # Dicionário para rastrear a primeira linha de cada TRP
    linhas_terminais = {} 

    # Criar todas as linhas primeiro
    for idx, linha in nc_fisc.iterrows():
        row_cells = tabela.add_row().cells

        # Colunas com dados
        data_map = {
            1: linha.get("ID", ""),
            2: linha.get("DESCRIÇÃO", ""),
            3: linha.get("REGISTROS FOTOGRÁFICOS", ""),
            4: linha.get("FUNDAMENTO DA INFRAÇÃO", ""),
            5: linha.get("DETERMINAÇÃO", ""),
         }

        # Coluna 0 (TRP) - Sigla
        trp_sigla = linha.get("TERMINAL RODOVIÁRIO", "")

        # Preenche as células de dados e aplica a formatação
        for i, (col_index, text) in enumerate(data_map.items()):
            par = row_cells[col_index].paragraphs[0]
            remover_espacamento_paragrafo(par)
            par.text = text
            # Alinhamento e tamanho
            par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if col_index == 2 else WD_ALIGN_PARAGRAPH.LEFT
            for run in par.runs:
                run.font.size = Pt(10)

            # 🚨 GARANTE ALINHAMENTO VERTICAL CENTRALIZADO PARA OS DADOS
            row_cells[col_index].vertical_alignment = WD_ALIGN_VERTICAL.CENTER


        # Armazena a posição da primeira célula do TRP para mesclagem
        if trp_sigla not in linhas_terminais:
            linhas_terminais[trp_sigla] = tabela.rows[-1].cells[0] # Usa a célula recém-criada

        # 🔴 CORREÇÃO 1: Preenche a sigla do TRP APENAS NA PRIMEIRA OCORRÊNCIA
        # As outras células na Coluna 0 (TRP) devem ficar vazias para que a mesclagem funcione corretamente
        par = row_cells[0].paragraphs[0]
        remover_espacamento_paragrafo(par)
        if trp_sigla == nc_fisc.iloc[nc_fisc.index.get_loc(idx) - 1].get("TERMINAL RODOVIÁRIO") and nc_fisc.index.get_loc(idx) != 0:
            par.text = "" # Repetição: limpa o texto
        else:
            par.text = trp_sigla # Primeira ocorrência: preenche o texto
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in par.runs:
                run.font.size = Pt(10)


    # 10. MESCLAGEM DAS CÉLULAS TRP

    # start_row começa na linha 1 (após o cabeçalho)
    start_row = 1 

    # Iteramos sobre o DataFrame agrupado para calcular as mesclagens
    for trp, group in nc_fisc.groupby("TERMINAL RODOVIÁRIO"):

        count_rows = len(group)
        end_row = start_row + count_rows - 1

        # Mescla as células se houver mais de uma linha para o TRP
        if count_rows > 1:
            primeira_celula = tabela.cell(start_row, 0)
            ultima_celula = tabela.cell(end_row, 0)

            merged_cell = primeira_celula.merge(ultima_celula)

            # 🟢 CORREÇÃO 2: APLICA ALINHAMENTO VERTICAL E HORIZONTAL FINAL NA CÉLULA MESCLADA
            merged_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            for par in merged_cell.paragraphs:
                par.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Atualiza o índice da próxima linha de início
        start_row = end_row + 1
        
    # 11. 🔴 ADIÇÃO DA LINHA DE TOTAL
    total_row_cells = tabela.add_row().cells
    
    # Mescla as colunas 0 a 4 para a célula TOTAL
    merged_total_cell = total_row_cells[0].merge(total_row_cells[4])
    
    # 🟢 Aplica sombreamento cinza na célula TOTAL mesclada
    apply_shading(merged_total_cell)
    
    # Texto "TOTAL"
    par_total = merged_total_cell.paragraphs[0]
    remover_espacamento_paragrafo(par_total)
    par_total.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_total = par_total.add_run("TOTAL")
    run_total.bold = True
    run_total.font.size = Pt(10)
    
    # Célula de contagem (coluna 5)
    cell_count = total_row_cells[5]
    # 🟢 Aplica sombreamento cinza na célula de contagem
    apply_shading(cell_count)
    cell_count.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    
    par_count = cell_count.paragraphs[0]
    remover_espacamento_paragrafo(par_count)
    par_count.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_count = par_count.add_run(str(total_nc))
    run_count.bold = True
    run_count.font.size = Pt(10)


     # Adiciona um espaço para separar o Quadro 1 do próximo conteúdo
    doc.add_paragraph()