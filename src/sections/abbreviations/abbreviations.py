from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

from utils import adicionar_titulo_secao, adicionar_tabela_abreviaturas,remover_espacamento_paragrafo 
import pandas as pd
import os
from docx.shared import  Pt 

def gerar_secao_abreviaturas(doc: Document, caminho_completo_planilha):
    """
    Gera a seção de Abreviaturas lendo diretamente a aba na planilha fornecida.
    
    Args:
        doc (Document): O objeto do documento Word.
        caminho_completo_planilha (str): O caminho absoluto para o arquivo Excel.
    """
    
    if len(doc.paragraphs) > 0 and doc.paragraphs[0].text == '':
        par_titulo = doc.paragraphs[0]
    else:
        par_titulo = doc.add_paragraph() 
        
    remover_espacamento_paragrafo(par_titulo) 
    par_titulo.paragraph_format.line_spacing_rule = None
    par_titulo.paragraph_format.line_spacing = None
    
    par_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par_titulo.text = ''
    
    run_titulo = par_titulo.add_run("RELATÓRIO DE FISCALIZAÇÃO")
    run_titulo.bold = True
    run_titulo.font.size = Pt(12) 

    doc.add_paragraph() 

    try:
   
        df_abreviaturas = pd.read_excel(
            caminho_completo_planilha, 
            sheet_name="Abreviaturas" 
        )
    except FileNotFoundError:
        doc.add_paragraph("ERRO: Planilha de fiscalização não encontrada (Verifique o caminho).")
        return
    except ValueError as e:

        print(f"\n⚠️ ERRO DE VALOR: Aba de Abreviaturas não encontrada ou nome incorreto: {e}")
        doc.add_paragraph("ERRO: Aba de Abreviaturas não encontrada ou nome incorreto. Verifique se o nome é 'Abreviaturas e Siglas'.")
        return

    df_abreviaturas.columns = df_abreviaturas.columns.str.strip() 


    df_abreviaturas = df_abreviaturas[['Sigla', 'Definição']].dropna(subset=['Sigla']).reset_index(drop=True)
    
    if df_abreviaturas.empty:
        doc.add_paragraph("AVISO: A aba de Abreviaturas está vazia ou as colunas não foram encontradas.")
        return

    adicionar_tabela_abreviaturas(doc, df_abreviaturas)
