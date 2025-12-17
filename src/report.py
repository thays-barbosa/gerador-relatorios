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
from sections.summary.summary import inserir_quebra_e_sumario 
from sections.conclusions.conclusions import gerar_secao_conclusoes
from sections.inspection.inspection import gerar_secao_fiscalizacao
from sections.finalprovisions.finalprovisions import gerar_secao_determinacoes_finais

from utils import (
    adicionar_texto_centralizado,
    ajustar_largura_colunas,
    arquivo_em_uso,
    forcar_atualizacao_campos 
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

    # Verificação de arquivos em uso
    if arquivo_em_uso(CAMINHO_PLANILHA):
        aguardar_e_encerrar("A planilha principal está em uso. Feche-a antes de executar.")
    if arquivo_em_uso(CAMINHO_PLANILHA_NCS):
        aguardar_e_encerrar("A planilha de Não-Conformidades está em uso. Feche-a antes de executar.")

    print("\n--- Configuração do Apêndice Fotográfico ---")
    while True:
        pasta_principal_input = input("➡️ Digite o nome da PASTA (ex: CTR-02-2024): ").strip()
        pasta_principal_fotos = pasta_principal_input.upper() 
        caminho_pasta_principal = os.path.join(FOTOS_DIR, pasta_principal_fotos)
        if not os.path.isdir(caminho_pasta_principal):
            print("❌ Pasta não encontrada. Tente novamente.")
        else:
            break

    while True:
        subpasta_input = input("➡️ Digite o nome da SUBPASTA (ex: F0): ").strip()
        subpasta_fotos = subpasta_input.upper()
        caminho_base_fotos = os.path.join(caminho_pasta_principal, subpasta_fotos)
        if not os.path.isdir(caminho_base_fotos):
            print("❌ Subpasta não encontrada. Tente novamente.")
        else:
            break

    match_ano = re.search(r'(\d{4})$', pasta_principal_fotos)
    if not match_ano:
        aguardar_e_encerrar(f"Não foi possível extrair o ano de: {pasta_principal_fotos}")
    ano_aba_ncs = match_ano.group(1)

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
            aguardar_e_encerrar(f"Formato de pasta inválido: {pasta_principal_fotos}")

    try:
        fiscalizacoes_df = pd.read_excel(CAMINHO_PLANILHA, sheet_name="Fiscalizações")
        fiscalizacoes_df.columns = fiscalizacoes_df.columns.str.strip() 
        fiscalizacoes_df['Nº Processo Limpo'] = fiscalizacoes_df['Nº Processo'].astype(str).str.strip()
    except Exception as e:
        aguardar_e_encerrar(f"Erro ao ler planilha principal: {e}")

    try:
        nao_conformidades_df = pd.read_excel(CAMINHO_PLANILHA_NCS, sheet_name=ano_aba_ncs).copy()
        nao_conformidades_df.columns = nao_conformidades_df.columns.str.strip()
        print(f"✅ Dados de Não-Conformidades carregados (Aba {ano_aba_ncs})")
    except Exception as e:
        aguardar_e_encerrar(f"Erro ao ler planilha de NCs: {e}")

    if COLUNA_STATUS not in fiscalizacoes_df.columns:
        fiscalizacoes_df[COLUNA_STATUS] = False
    fiscalizacoes_df[COLUNA_STATUS] = fiscalizacoes_df[COLUNA_STATUS].fillna(False).astype(bool)

    relatorios_a_gerar = fiscalizacoes_df[
        (fiscalizacoes_df['Nº Processo Limpo'] == processo_numero_alvo) & 
        (fiscalizacoes_df[COLUNA_STATUS] == False)
    ].copy()

    if relatorios_a_gerar.empty:
        print(f"✅ Nenhum relatório pendente para o PROCESSO: {id_processo_alvo}.")
        return

    for idx in tqdm(relatorios_a_gerar.index, desc="Gerando relatórios"):
        row = fiscalizacoes_df.loc[idx].copy()
        row['PROCESSO'] = id_processo_alvo 
        id_fisc_numerico = row["ID da Fiscalização"] 
 
        doc = Document()
    
        forcar_atualizacao_campos(doc)

        gerar_capa(doc, BASE_DIR, row)
        
        inserir_quebra_e_sumario(doc)

        gerar_secao_abreviaturas(doc, CAMINHO_PLANILHA) 
        doc.add_section(WD_SECTION.NEW_PAGE) 
        
        gerar_secao_introducao(doc)
        gerar_secao_objetivo(doc, row)
        gerar_secao_metodologia(doc, row)
        gerar_secao_fiscalizacao(doc, row, nao_conformidades_df) 
        gerar_secao_determinacoes_finais(doc, row)
        gerar_secao_recomendacoes(doc, row)
        gerar_secao_conclusoes(doc, row, CAMINHO_PLANILHA, caminho_base_fotos)

        nome_arquivo = f"relatorio_{id_fisc_numerico}" 
        caminho_docx = os.path.join(RELATORIOS_DIR, f"{nome_arquivo}.docx")
        caminho_pdf = os.path.join(RELATORIOS_DIR, f"{nome_arquivo}.pdf")

        doc.save(caminho_docx)
        
        try:
            convert(caminho_docx, caminho_pdf)
        except Exception as e:
            print(f"⚠️ Aviso: Falha ao converter para PDF (é necessário ter o Word instalado): {e}")
        
        fiscalizacoes_df.loc[idx, COLUNA_STATUS] = True

    if "Data" in fiscalizacoes_df.columns:
        fiscalizacoes_df["Data"] = pd.to_datetime(fiscalizacoes_df["Data"], errors="coerce").dt.strftime("%d/%m/%Y")

    if not arquivo_em_uso(CAMINHO_PLANILHA):
        with pd.ExcelWriter(CAMINHO_PLANILHA, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df_final = fiscalizacoes_df.drop(columns=['Nº Processo Limpo'], errors='ignore')
            df_final.to_excel(writer, sheet_name="Fiscalizações", index=False) 
        ajustar_largura_colunas(CAMINHO_PLANILHA)

    print("\n🎉 Processo concluído com sucesso!")
    print("👉 NOTA: Ao abrir o arquivo .docx, clique em 'Sim' quando o Word perguntar sobre atualização de campos.")
    return caminho_docx, caminho_pdf

if __name__ == "__main__":
    gerar_relatorio()