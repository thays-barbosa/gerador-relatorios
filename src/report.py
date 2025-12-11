from docx import Document
from docx2pdf import convert
from docx.shared import Inches
import pandas as pd
from tqdm import tqdm
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import sys
import os
from sections.introduction.introduction import gerar_secao_introducao
from sections.objective.objective import gerar_secao_objetivo
from sections.recommendations.recommendations import gerar_secao_recomendacoes
from sections.methodology.methodology import gerar_secao_metodologia
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


# --- NOVA FUNÇÃO AUXILIAR PARA SAÍDA CONTROLADA ---
def aguardar_e_encerrar(mensagem):
    """Exibe uma mensagem de erro e espera o usuário pressionar Enter para encerrar."""
    print(f"\n❌ ERRO: {mensagem}")
    input("Aperte Enter para encerrar...")
    sys.exit(1)
# ---------------------------------------------------


def gerar_relatorio():
    """
    Gera o relatório completo (docx + pdf) com base nos dados da fiscalização.
    """

    if getattr(sys, "frozen", False):
        BASE_DIR = os.path.dirname(sys.executable)
    else:
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))

    FOTOS_DIR = os.path.join(BASE_DIR, "assets")
    RELATORIOS_DIR = os.path.join(BASE_DIR, "reports")
    CAMINHO_PLANILHA = os.path.join(BASE_DIR, "planilha_fiscalizacao.xlsx")
    COLUNA_STATUS = "Relatório Gerado"

    os.makedirs(RELATORIOS_DIR, exist_ok=True)
    os.makedirs(FOTOS_DIR, exist_ok=True)

    if arquivo_em_uso(CAMINHO_PLANILHA):
        print("⚠️ A planilha está em uso. Feche-a antes de executar o script.")
        exit(1)

    DIRETORIO_PAI_FOTOS = FOTOS_DIR 
    
    
    # --- NOVO BLOCO DE ENTRADA E VALIDAÇÃO DE CAMINHO ---
    print("\n--- Configuração do Apêndice Fotográfico ---")
    
    # 1. VALIDAÇÃO DA PASTA PRINCIPAL (CTR-XX-XXXX)
    while True:
        pasta_principal_input = input("➡️ Digite o nome da PASTA (ex: CTR-02-2024): ").strip()
        
        # Converte para MAIÚSCULAS para ser case-insensitive na checagem
        pasta_principal_fotos = pasta_principal_input.upper() 
        
        caminho_pasta_principal = os.path.join(DIRETORIO_PAI_FOTOS, pasta_principal_fotos)
        
        if not os.path.isdir(caminho_pasta_principal):
            aguardar_e_encerrar("A Pasta procurada não existe, por favor, cheque seus documentos.")
        else:
            break
            
    # 2. VALIDAÇÃO DA SUBPASTA (F0)
    while True:
        subpasta_input = input("➡️ Digite o nome da SUBPASTA (ex: F0): ").strip()
        
        # Converte para MAIÚSCULAS para ser case-insensitive na checagem
        subpasta_fotos = subpasta_input.upper()
        
        caminho_base_fotos = os.path.join(caminho_pasta_principal, subpasta_fotos)
        
        if not os.path.isdir(caminho_base_fotos):
            aguardar_e_encerrar("A subpasta procurada não existe, por favor, cheque seus documentos.")
        else:
            break
            
    print("------------------------------------------") 
    # ----------------------------------------------------


    fiscalizacoes_df = pd.read_excel(CAMINHO_PLANILHA, sheet_name="Fiscalizações")
    nao_conformidades_df = pd.read_excel(
        CAMINHO_PLANILHA, sheet_name="Não-conformidades "
    )

    if COLUNA_STATUS not in fiscalizacoes_df.columns:
        fiscalizacoes_df[COLUNA_STATUS] = False
    fiscalizacoes_df[COLUNA_STATUS] = (
        fiscalizacoes_df[COLUNA_STATUS].fillna(False).astype(bool)
    )

    pendentes = fiscalizacoes_df[~fiscalizacoes_df[COLUNA_STATUS]]

    if pendentes.empty:
        print("✅ Nenhum relatório pendente.")
        return

    for idx in tqdm(pendentes.index, desc="Gerando relatórios"):
        row = fiscalizacoes_df.loc[idx]
        id_fisc = row["ID da Fiscalização"]
        doc = Document()

        adicionar_texto_centralizado(doc, "RELATÓRIO DE FISCALIZAÇÃO")
        doc.add_picture(os.path.join(BASE_DIR, "assets/logo_arpe.jpg"), width=Inches(6)) #mudei de 2 para 6
        logo_arpe = doc.paragraphs[-1]
        logo_arpe.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        adicionar_texto_centralizado(doc, "FISCALIZAÇÃO NOS TERMINAIS RODOVIÁRIOS INTERMUNICIPAIS DE PASSAGEIROSL")
        adicionar_texto_centralizado(doc, "PRESTADOR DE SERVIÇO: SOCICAM - ADMINISTRAÇÃO, PROJETOS E REPRESENTAÇÕES LTDA")
        adicionar_texto_centralizado(
            doc, "RELATÓRIO DE FISCALIZAÇÃO PROC ADM Nº xx/xxxx - CTR"
        )
        adicionar_texto_centralizado(
            doc, "SEI Nº xxxxxxxxxx.xxxxxx/xxxx-xx"
        )

        doc.add_section(WD_SECTION.NEW_PAGE)

        gerar_secao_introducao(doc)
        gerar_secao_objetivo(doc)
        gerar_secao_metodologia(doc, row)
        gerar_secao_fiscalizacao(doc, row, nao_conformidades_df)
        gerar_secao_determinacoes_finais(doc, row)
        gerar_secao_recomendacoes(doc,row)
        
        # --- CHAMADA MODIFICADA: Passa apenas o caminho final validado ---
        gerar_secao_conclusoes(
            doc, 
            row, 
            caminho_planilha_legendas=CAMINHO_PLANILHA, 
            caminho_base_fotos=caminho_base_fotos # NOVO ARGUMENTO ÚNICO
        )
        # -----------------------------------------------------------------

        nome_arquivo = f"relatorio_{id_fisc}"
        caminho_docx = os.path.join(RELATORIOS_DIR, f"{nome_arquivo}.docx")
        caminho_pdf = os.path.join(RELATORIOS_DIR, f"{nome_arquivo}.pdf")

        doc.save(caminho_docx)
        convert(caminho_docx, caminho_pdf)
        fiscalizacoes_df.at[idx, COLUNA_STATUS] = True

    # 🔹 Garantir que a coluna Data seja salva no formato dd/mm/aaaa
    if "Data" in fiscalizacoes_df.columns:
        fiscalizacoes_df["Data"] = pd.to_datetime(
            fiscalizacoes_df["Data"], errors="coerce"
        ).dt.strftime("%d/%m/%Y")

    if not arquivo_em_uso(CAMINHO_PLANILHA):
        with pd.ExcelWriter(
            CAMINHO_PLANILHA, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            fiscalizacoes_df.to_excel(writer, sheet_name="Fiscalizações", index=False)
            nao_conformidades_df.to_excel(
                writer, sheet_name="Não-conformidades ", index=False
            )

        ajustar_largura_colunas(CAMINHO_PLANILHA)

    print("🎉 Relatórios gerados e planilha atualizada com sucesso.")

    return caminho_docx, caminho_pdf