# report.py - COMPLETO E CORRIGIDO PARA FILTRO POR Nº PROCESSO

from docx import Document
from docx2pdf import convert
from docx.shared import Inches
import pandas as pd
from tqdm import tqdm
from docx.enum.section import WD_SECTION 
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import sys
import os
import re 
import warnings
from sections.introduction.introduction import gerar_secao_introducao
from sections.objective.objective import gerar_secao_objetivo
from sections.recommendations.recommendations import gerar_secao_recomendacoes
from sections.methodology.methodology import gerar_secao_metodologia
from sections.abbreviations.abbreviations import gerar_secao_abreviaturas
from sections.cover.cover import gerar_capa 
from sections.conclusions.conclusions import (
    gerar_secao_conclusoes,
)
from sections.inspection.inspection import (
    gerar_secao_fiscalizacao,
)
from sections.finalprovisions.finalprovisions import (
    gerar_secao_determinacoes_finais,
)
from utils import (
    adicionar_texto_centralizado,
    ajustar_largura_colunas,
    arquivo_em_uso,
)


def aguardar_e_encerrar(mensagem):
    """Exibe uma mensagem de erro e espera o usuário pressionar Enter para encerrar."""
    print(f"\n❌ ERRO: {mensagem}")
    input("Aperte Enter para encerrar...")
    sys.exit(1)


def gerar_relatorio():
    """
    Gera o relatório completo (docx + pdf) com base nos dados da fiscalização.
    """
    warnings.filterwarnings("ignore", category=pd.errors.SettingWithCopyWarning) 

    if getattr(sys, "frozen", False):
        BASE_DIR = os.path.dirname(sys.executable)
    else:
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))

    FOTOS_DIR = os.path.join(BASE_DIR, "assets")
    RELATORIOS_DIR = os.path.join(BASE_DIR, "reports")
    CAMINHO_PLANILHA = os.path.join(BASE_DIR, "planilha_fiscalizacao.xlsx")

    CAMINHO_PLANILHA_NCS = os.path.join(BASE_DIR, "levantamento de NCS - SOCICAM.xlsx")

    COLUNA_STATUS = "Relatório Gerado" 

    os.makedirs(RELATORIOS_DIR, exist_ok=True)
    os.makedirs(FOTOS_DIR, exist_ok=True)

    if arquivo_em_uso(CAMINHO_PLANILHA):
        print("⚠️ A planilha principal está em uso. Feche-a antes de executar o script.")
        exit(1)

    if arquivo_em_uso(CAMINHO_PLANILHA_NCS):
        print("⚠️ A planilha de Não-Conformidades está em uso. Feche-a antes de executar o script.")
        exit(1)


    DIRETORIO_PAI_FOTOS = FOTOS_DIR 

 # --- BLOCO DE ENTRADA E VALIDAÇÃO DE CAMINHO ---
    print("\n--- Configuração do Apêndice Fotográfico ---")

    pasta_principal_fotos = "" 
    subpasta_fotos = ""
    caminho_base_fotos = ""

    while True:
        pasta_principal_input = input("➡️ Digite o nome da PASTA (ex: CTR-02-2024): ").strip()
        pasta_principal_fotos = pasta_principal_input.upper() 
        caminho_pasta_principal = os.path.join(DIRETORIO_PAI_FOTOS, pasta_principal_fotos)

        if not os.path.isdir(caminho_pasta_principal):
            aguardar_e_encerrar("A Pasta procurada não existe, por favor, cheque seus documentos.")
        else:
            break

    while True:
        subpasta_input = input("➡️ Digite o nome da SUBPASTA (ex: F0): ").strip()
        subpasta_fotos = subpasta_input.upper()
        caminho_base_fotos = os.path.join(caminho_pasta_principal, subpasta_fotos)

        if not os.path.isdir(caminho_base_fotos):
            aguardar_e_encerrar("A subpasta procurada não existe, por favor, cheque seus documentos.")
        else:
            break

    print("------------------------------------------") 

    # --- NOVO BLOCO DE IDENTIFICAÇÃO DO PROCESSO ALVO ---

    # 🚨 ETAPA 1: DETERMINAÇÃO DO ANO DA ABA (Para a planilha de NCS)
    match_ano = re.search(r'(\d{4})$', pasta_principal_fotos)
    if not match_ano:
        aguardar_e_encerrar(f"Não foi possível extrair o ano do nome da pasta: {pasta_principal_fotos}")

    ano_aba_ncs = match_ano.group(1)

    # 🚨 ETAPA 2: DETERMINAÇÃO DO PROCESSO ALVO NO FORMATO DO EXCEL 
    match_processo = re.match(r'(CTR)-(\d+)-(\d{4})$', pasta_principal_fotos.strip())

    if match_processo:
        id_processo_alvo = f"{match_processo.group(1)} {match_processo.group(2)}/{match_processo.group(3)}"
        processo_numero_alvo = f"{match_processo.group(2)}/{match_processo.group(3)}"
    else:
        partes = pasta_principal_fotos.replace('-', ' ').split()
        if len(partes) == 3 and partes[0].upper() == 'CTR':
            id_processo_alvo = f"{partes[0].upper()} {partes[1]}/{partes[2]}"
            processo_numero_alvo = f"{partes[1]}/{partes[2]}"
        else:
            aguardar_e_encerrar(f"O nome da pasta ('{pasta_principal_fotos}') não pôde ser convertido para o formato PROCESSO ('CTR XX/YYYY').")


    # --- Lendo planilhas e limpando nomes de colunas por segurança ---
 
# 1. PLANILHA PRINCIPAL (Aba Fiscalizações)
    try:
        fiscalizacoes_df = pd.read_excel(CAMINHO_PLANILHA, sheet_name="Fiscalizações")
        fiscalizacoes_df.columns = fiscalizacoes_df.columns.str.strip() 

        if 'Nº Processo' not in fiscalizacoes_df.columns:
         aguardar_e_encerrar("Coluna 'Nº Processo' não encontrada na aba 'Fiscalizações' do arquivo principal. Verifique o cabeçalho.")

        fiscalizacoes_df['Nº Processo Limpo'] = fiscalizacoes_df['Nº Processo'].astype(str).str.strip()

    except FileNotFoundError:
        aguardar_e_encerrar(f"Planilha principal não encontrada: {CAMINHO_PLANILHA}")

    # 2. PLANILHA DE NÃO-CONFORMIDADES (Arquivo separado, aba dinâmica)
    try:
        nao_conformidades_df = pd.read_excel(
            CAMINHO_PLANILHA_NCS, sheet_name=ano_aba_ncs
        ).copy()
        nao_conformidades_df.columns = nao_conformidades_df.columns.str.strip()
 
        if 'PROCESSO' not in nao_conformidades_df.columns:
            aguardar_e_encerrar("Coluna 'PROCESSO' não encontrada na aba anual de Não-Conformidades. Verifique o cabeçalho.")

        nao_conformidades_df['PROCESSO_LIMPO'] = nao_conformidades_df['PROCESSO'].astype(str).str.strip()

        print(f"✅ Dados de Não-Conformidades carregados da aba '{ano_aba_ncs}'")
    except FileNotFoundError:
        aguardar_e_encerrar(f"Planilha de Não-Conformidades não encontrada: {CAMINHO_PLANILHA_NCS}")
    except ValueError:
        aguardar_e_encerrar(f"A aba '{ano_aba_ncs}' não foi encontrada na planilha de Não-Conformidades. Verifique se o nome da aba está correto.")


    # --- LÓGICA DE FILTRAGEM DO RELATÓRIO ALVO ---

    if COLUNA_STATUS not in fiscalizacoes_df.columns:
        fiscalizacoes_df[COLUNA_STATUS] = False

    # CORREÇÃO FINAL: Garante que a coluna seja estritamente booleana, evitando o FutureWarning.
    fiscalizacoes_df.loc[:, COLUNA_STATUS] = (
        fiscalizacoes_df[COLUNA_STATUS].fillna(False).astype(bool)
    )

    # BUSCA PELA COLUNA 'Nº Processo Limpo'
    relatorios_a_gerar = fiscalizacoes_df[
        (fiscalizacoes_df['Nº Processo Limpo'] == processo_numero_alvo) & 
        (~fiscalizacoes_df[COLUNA_STATUS])
    ].copy()

    if relatorios_a_gerar.empty:
        print(f"✅ Nenhum relatório pendente encontrado para o PROCESSO: {id_processo_alvo}.")

        if any(fiscalizacoes_df['Nº Processo Limpo'] == processo_numero_alvo):
            print("⚠️ Este relatório já foi marcado como 'Gerado'.")

        return

    # Iteração (deve ser apenas uma linha)
    for idx in tqdm(relatorios_a_gerar.index, desc="Gerando relatórios"):
        row = fiscalizacoes_df.loc[idx].copy()

        # INJEÇÃO CRÍTICA: Adiciona a chave 'PROCESSO' completa na linha do relatório
        row['PROCESSO'] = id_processo_alvo 

        #O ID usado para o nome do arquivo é o numérico
        id_fisc_numerico = row["ID da Fiscalização"] 
 
        doc = Document()

        # ... (Geração das seções) ...

        gerar_capa(doc, BASE_DIR, row)
        gerar_secao_abreviaturas(doc, CAMINHO_PLANILHA) 
        doc.add_section(WD_SECTION.NEW_PAGE) 
        gerar_secao_introducao(doc)
        gerar_secao_objetivo(doc,row)
        gerar_secao_metodologia(doc, row)

        # Passa a linha com a chave 'PROCESSO' e o df de NCs
        gerar_secao_fiscalizacao(doc, row, nao_conformidades_df) 

        gerar_secao_determinacoes_finais(doc, row)
        gerar_secao_recomendacoes(doc,row)

        gerar_secao_conclusoes(
            doc, 
            row, 
            caminho_planilha_legendas=CAMINHO_PLANILHA, 
            caminho_base_fotos=caminho_base_fotos 
        )

        # --- SALVAMENTO ---
        nome_arquivo = f"relatorio_{id_fisc_numerico}" 
        caminho_docx = os.path.join(RELATORIOS_DIR, f"{nome_arquivo}.docx")
        caminho_pdf = os.path.join(RELATORIOS_DIR, f"{nome_arquivo}.pdf")

        doc.save(caminho_docx)
        convert(caminho_docx, caminho_pdf)
        fiscalizacoes_df.loc[idx, COLUNA_STATUS] = True

    # 🔹 Garantir que a coluna Data seja salva no formato dd/mm/aaaa
    if "Data" in fiscalizacoes_df.columns:
        fiscalizacoes_df.loc[:, "Data"] = pd.to_datetime(
            fiscalizacoes_df["Data"], errors="coerce"
        ).dt.strftime("%d/%m/%Y")

    if not arquivo_em_uso(CAMINHO_PLANILHA):
        with pd.ExcelWriter(
            CAMINHO_PLANILHA, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            # Remove a coluna temporária 'Nº Processo Limpo' antes de salvar
            cols_to_drop = ['Nº Processo Limpo']
            fiscalizacoes_df.drop(columns=[col for col in cols_to_drop if col in fiscalizacoes_df.columns], errors='ignore').to_excel(writer, sheet_name="Fiscalizações", index=False) 

        ajustar_largura_colunas(CAMINHO_PLANILHA)

    print("🎉 Relatório gerado e planilha principal atualizada com sucesso.")

    return caminho_docx, caminho_pdf