# sections/inspection/inspection.py - MODIFICADO

from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor # Importação adicional útil para formatação de tabela
from docx.table import _Cell # Para alinhamento vertical
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


# --- FUNÇÃO AUXILIAR PARA ALINHAMENTO VERTICAL ---
# Referência: https://stackoverflow.com/questions/32339893/python-docx-set-vertical-alignment-of-cell
def set_vertical_cell_alignment(cell, align='center'):
    """Define o alinhamento vertical do conteúdo de uma célula (top, center, bottom)."""
    cell.vertical_alignment = align.upper() # Necessário importar WD_ALIGN_VERTICAL do docx.enum.table se usar os enums

# --- FUNÇÃO PRINCIPAL ---

def gerar_secao_fiscalizacao(doc: Document, row, nao_conformidades_df):
    """
    Gera a seção '4. FISCALIZAÇÃO' e a tabela de Não Conformidades (Quadro 1), 
    com informações dinâmicas de equipe e locais.
    """

    # 1. Título Principal
    adicionar_titulo_secao(doc, "4. FISCALIZAÇÃO")
    
    # 2. Preparação dos dados dinâmicos (Equipe e Locais)
    
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
        
    # 3. Construção do parágrafo dinâmico (Equipe e Locais)
    
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
    
    
    # 4. Segundo Parágrafo (Introdução às Não Conformidades)
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
    tabela.columns[1].width = Cm(2.5)  # IDENTIFICAÇÃO
    tabela.columns[2].width = Cm(4.5)  # DESCRIÇÃO
    tabela.columns[3].width = Cm(2.5) # REGISTRO FOTOGRÁFICO
    tabela.columns[4].width = Cm(3.8) # FUNDAMENTO DA INFRAÇÃO
    tabela.columns[5].width = Cm(2.5)  # DETERMINAÇÃO
    
    
    # Cabeçalhos
    cabecalhos_nomes = ["TRP", "IDENTIFICAÇÃO", "DESCRIÇÃO", "REGISTRO FOTOGRÁFICO", "FUNDAMENTO DA INFRAÇÃO (ANEXO V CONTRATO DE CONCESSÃO)", "DETERMINAÇÃO"]
    
    cabecalho_cells = tabela.rows[0].cells
    
    for i, nome in enumerate(cabecalhos_nomes):
        cell = cabecalho_cells[i]
        cell.text = nome
        # Aplica estilo: centralizado, negrito, tamanho 9/10, compactado
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
        
        # Armazena a posição da primeira célula do TRP para mesclagem
        if trp_sigla not in linhas_terminais:
            linhas_terminais[trp_sigla] = row_cells[0] 
        
        # Preenche a sigla do TRP na célula 0 (temporário)
        par = row_cells[0].paragraphs[0]
        remover_espacamento_paragrafo(par)
        par.text = trp_sigla
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in par.runs:
            run.font.size = Pt(10)


    # 10. MESCLAGEM DAS CÉLULAS TRP
    
    start_row = 1 # Começa após o cabeçalho (linha 0)
    
    for trp in linhas_terminais.keys():
        
        # Conta quantas linhas deste TRP existem no DataFrame filtrado e ordenado
        count_rows = len(nc_fisc[nc_fisc["TERMINAL RODOVIÁRIO"] == trp])
        end_row = start_row + count_rows - 1

        # Mescla as células se houver mais de uma linha para o TRP
        if count_rows > 1:
            primeira_celula = tabela.cell(start_row, 0)
            ultima_celula = tabela.cell(end_row, 0)
            
            merged_cell = primeira_celula.merge(ultima_celula)
            
            # 🚨 MELHORIA NO ALINHAMENTO VERTICAL DA CÉLULA MESCLADA
            # A função set_vertical_cell_alignment (nativa do docx, mas sem import) pode ser usada aqui.
            # Se a importação _Cell do docx.table fosse necessária, seria assim:
            # if isinstance(merged_cell, _Cell):
            #     set_vertical_cell_alignment(merged_cell, 'center') 
            
            # Centraliza horizontalmente o texto da TRP
            for par in merged_cell.paragraphs:
                 par.alignment = WD_ALIGN_PARAGRAPH.CENTER
                 
        # Atualiza o índice da próxima linha de início
        start_row = end_row + 1

    # Adiciona um espaço para separar o Quadro 1 do próximo conteúdo
    doc.add_paragraph()