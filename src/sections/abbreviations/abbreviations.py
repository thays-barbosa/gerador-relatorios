# --- CÓDIGO CORRIGIDO PARA abbreviations.py ---

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
# Importa a função de utils, mas sem WD_PARAGRAPH_ALIGNMENT que está em docx.enum.text
from utils import adicionar_titulo_secao, adicionar_tabela_abreviaturas 
import pandas as pd
import os

# Remove a definição de CAMINHO_PLANILHA aqui, pois ele virá como argumento

# A função AGORA espera o caminho COMPLETO da planilha, não o caminho base do diretório.
def gerar_secao_abreviaturas(doc: Document, caminho_completo_planilha):
    """
    Gera a seção de Abreviaturas lendo diretamente a aba na planilha fornecida.
    
    Args:
        doc (Document): O objeto do documento Word.
        caminho_completo_planilha (str): O caminho absoluto para o arquivo Excel.
    """
    
    # Adicionar Título (Usando a função de utils, se ela tiver sido definida)
    # Se você não definiu 'adicionar_titulo_secao' no utils, use esta linha:
    # doc.add_paragraph("LISTA DE ABREVIATURAS E SIGLAS", style='Heading 1').alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Se você usa a função utilitária para títulos numerados:
    adicionar_titulo_secao(doc, "LISTA DE ABREVIATURAS E SIGLAS") # Assumindo título não numerado na Etapa 2
    
    try:
        # Lê a aba 'Abreviaturas' - *NOTE BEM:* Se o nome da aba for "Abreviaturas e Siglas", 
        # mude o sheet_name abaixo!
        df_abreviaturas = pd.read_excel(
            caminho_completo_planilha, 
            sheet_name="Abreviaturas" # 🚨 VERIFIQUE SE O NOME DA SUA ABA É ESTE.
        )
    except FileNotFoundError:
        doc.add_paragraph("ERRO: Planilha de fiscalização não encontrada (Verifique o caminho).")
        return
    except ValueError as e:
        # Mostra o erro exato para diagnóstico (ex: nome da aba errado)

        print(f"\n⚠️ ERRO DE VALOR: Aba de Abreviaturas não encontrada ou nome incorreto: {e}")
        doc.add_paragraph("ERRO: Aba de Abreviaturas não encontrada ou nome incorreto. Verifique se o nome é 'Abreviaturas e Siglas'.")
        return

    # Limpa nomes de colunas por segurança
    df_abreviaturas.columns = df_abreviaturas.columns.str.strip() 

    # Seleciona as colunas 'Sigla' e 'Definição'
    # 🚨 NOTA: Se o nome da coluna no Excel é 'Definicao', você precisa ajustar aqui.
    df_abreviaturas = df_abreviaturas[['Sigla', 'Definição']].dropna(subset=['Sigla']).reset_index(drop=True)
    
    if df_abreviaturas.empty:
        doc.add_paragraph("AVISO: A aba de Abreviaturas está vazia ou as colunas não foram encontradas.")
        return

    # Adiciona a tabela ao documento
    adicionar_tabela_abreviaturas(doc, df_abreviaturas)


# Remova esta linha:
# CAMINHO_PLANILHA = "planilha_fiscalizacao.xlsx"