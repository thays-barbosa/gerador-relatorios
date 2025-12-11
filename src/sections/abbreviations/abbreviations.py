# --- CÓDIGO FINAL PARA abbreviations.py ---

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import adicionar_texto_centralizado, adicionar_tabela_abreviaturas # Importar a nova função
import pandas as pd
import os

# Defina o caminho da planilha (assumindo que está no mesmo BASE_DIR do report.py)
CAMINHO_PLANILHA = "planilha_fiscalizacao.xlsx" # Adapte este caminho se necessário

def gerar_secao_abreviaturas(doc: Document, caminho_base_dir):
    
    adicionar_texto_centralizado(doc, "LISTA DE ABREVIATURAS E SIGLAS")

    caminho_completo_planilha = os.path.join(caminho_base_dir, CAMINHO_PLANILHA)

    try:
        # Lê a aba 'Abreviaturas' (imagem_523323.png)
        df_abreviaturas = pd.read_excel(
            caminho_completo_planilha, 
            sheet_name="Abreviaturas"
        )
    except FileNotFoundError:
        doc.add_paragraph("ERRO: Planilha de fiscalização não encontrada.")
        return
    except ValueError:
        doc.add_paragraph("ERRO: Aba 'Abreviaturas' não encontrada na planilha. Verifique o nome da aba.")
        return

    # Seleciona apenas as colunas 'Sigla' e 'Definicao' (colunas B e C)
    # E remove linhas que possam ter 'NaN' ou estarem vazias
    df_abreviaturas = df_abreviaturas[['Sigla', 'Definição ']].dropna(subset=['Sigla', 'Definição '])
    
    if df_abreviaturas.empty:
        doc.add_paragraph("AVISO: A aba 'Abreviaturas' está vazia.")
        return

    # Adiciona a tabela ao documento
    adicionar_tabela_abreviaturas(doc, df_abreviaturas)