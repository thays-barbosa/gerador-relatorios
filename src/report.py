from typing import Optional

import os
import sys
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pandas as pd
from tqdm import tqdm
import win32com.client as win32

from sections.introduction.introduction import gerar_secao_introducao
from sections.objective.objective import gerar_secao_objetivo
from sections.anexo.anexo import gerar_secao_anexo_fotos
from sections.nonconformity.nonconformity import gerar_secao_nao_conformidades_constatadas
from sections.nonconformityresume.nonconformityresume import gerar_secao_resumo_nao_conformidades
from sections.finalconsiderations.finalconsiderations import gerar_secao_consideracoes_finais
from sections.summary.summary import inserir_quebra_e_sumario

from utils import (
    adicionar_texto_centralizado,
    adicionar_paragrafo_justificado,
    ajustar_largura_colunas,
    arquivo_em_uso,
    processar_imagem_para_relatorio,
    adicionar_imagem,
)


def atualizar_toc_e_converter_para_pdf(caminho_docx: str, caminho_pdf: str) -> bool:
    """
    Abre o Word via COM, atualiza campos (sumário) e salva como PDF.
    Retorna True em caso de sucesso, False em caso de erro.
    """
    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(caminho_docx)

        
        doc.Fields.Update()
        doc.Save()
       
        doc.SaveAs2(caminho_pdf, FileFormat=17)

        doc.Close(SaveChanges=False)
        word.Quit()
        return True

    except Exception as exc:
        print(f"❌ ERRO ao gerar PDF/Atualizar Sumário (verifique se o Word está instalado): {exc}")
        return False


def _obter_base_dir() -> str:
    """
    Retorna o diretório base do projeto.
    Quando empacotado (frozen), usa o executável; caso contrário, usa o __file__.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def gerar_relatorio() -> None:
    """
    Gera todos os relatórios DOCX e converte para PDF conforme as fiscalizações pendentes
    registradas na planilha 'planilha_fiscalizacao.xlsx'.
    """

    BASE_DIR = _obter_base_dir()

    # Diretórios e caminhos principais
    FOTOS_DIR = os.path.join(BASE_DIR, "assets")
    RELATORIOS_DIR = os.path.join(BASE_DIR, "reports")
    CAMINHO_PLANILHA = os.path.join(BASE_DIR, "planilha_fiscalizacao.xlsx")
    COLUNA_STATUS = "Relatório Gerado"

    # Garante diretórios existirem
    os.makedirs(RELATORIOS_DIR, exist_ok=True)
    os.makedirs(FOTOS_DIR, exist_ok=True)

    # Verifica se planilha está em uso
    if arquivo_em_uso(CAMINHO_PLANILHA):
        print("⚠️ A planilha está em uso. Feche-a antes de executar o script.")
        sys.exit(1)

   
    fiscalizacoes_df = pd.read_excel(CAMINHO_PLANILHA, sheet_name="Fiscalizações")
    nao_conformidades_df = pd.read_excel(CAMINHO_PLANILHA, sheet_name="Não-conformidades ")

   
    if COLUNA_STATUS not in fiscalizacoes_df.columns:
        fiscalizacoes_df[COLUNA_STATUS] = False
    fiscalizacoes_df[COLUNA_STATUS] = fiscalizacoes_df[COLUNA_STATUS].fillna(False).astype(bool)

    pendentes = fiscalizacoes_df[~fiscalizacoes_df[COLUNA_STATUS]]

    if pendentes.empty:
        print("\n✅ Todos os relatórios já foram gerados,não existe nenhum relatório pendente. Por favor, confira sua planilha.")
        input("Pressione Enter para sair...")
        return

    
    print("\n--- Configuração de Pastas de Fotos ---")

    # Validação da pasta CTR (contrato)
    while True:
        pasta_contrato_input = input("Digite a pasta de Fiscalização (Ex: CTR-01-2025 ou ctr-01-2025): ").strip()
        if pasta_contrato_input:
            pasta_contrato = pasta_contrato_input.upper()
            break
        print("⚠️ O nome da pasta CTR é obrigatório.")

    CAMINHO_CTR = os.path.join(FOTOS_DIR, pasta_contrato)

    if not os.path.isdir(CAMINHO_CTR):
        print(f"\n❌ A pasta de Contrato '{pasta_contrato}' não foi encontrada. Por favor, verifique seus arquivos.")
        print("Pressione Enter para sair...")
        input()
        sys.exit(0)

    if not os.listdir(CAMINHO_CTR):
        print(f"\n⚠️ A pasta de Contrato '{pasta_contrato}' está vazia.Por favor, verifique seus arquivos.")
        print("Não há pastas de Monitoramento. Pressione Enter para sair...")
        input()
        sys.exit(0)

    # Validação da pasta de Monitoramento
    while True:
        pasta_monitoramento_input = input("Digite a Pasta do Monitoramento (Ex: M0 ou m0): ").strip()
        if pasta_monitoramento_input:
            pasta_monitoramento = pasta_monitoramento_input.upper()
            break
        print("⚠️ O nome da pasta de monitoramento é obrigatório.")

    CAMINHO_RAIZ_FOTOS = os.path.join(CAMINHO_CTR, pasta_monitoramento)

    if not os.path.isdir(CAMINHO_RAIZ_FOTOS):
        print(f"\n❌ A subpasta de Monitoramento '{pasta_monitoramento}' não foi encontrada dentro de '{pasta_contrato}'. Por favor, verifique seus arquivos.")
        print("Pressione Enter para sair...")
        input()
        sys.exit(0)

    if not os.listdir(CAMINHO_RAIZ_FOTOS):
        print(f"\n⚠️ A subpasta de Monitoramento '{pasta_monitoramento}' está vazia.")
        print("O relatório será **gerado sem o Anexo de Fotos** para esta fiscalização.")

    
    NOME_LOGO = "capa_monitoramento_arpe.jpg"

    # Itera sobre fiscalizações pendentes
    for idx in tqdm(pendentes.index, desc="Gerando relatórios"):
        row = fiscalizacoes_df.loc[idx]
        id_fisc = row["ID da Fiscalização"]

        doc = Document()
       
        section = doc.sections[0]
        section.top_margin = Inches(0.25)

        
        adicionar_texto_centralizado(doc, "COORDENADORIA DE TRANSPORTES E RODOVIAS")
        primeiro_paragrafo = doc.paragraphs[-1]
        primeiro_paragrafo.paragraph_format.space_before = Pt(0)
        primeiro_paragrafo.paragraph_format.space_after = Pt(0)

        doc.add_paragraph()
        adicionar_texto_centralizado(doc, "RELATÓRIO DO Xº MONITORAMENTO DAS NÃO CONFORMIDADES DO PROCESSO CTR Nº XX/XXXX")
        doc.add_paragraph()

       
        caminho_logo = os.path.join(FOTOS_DIR, NOME_LOGO)
        if os.path.exists(caminho_logo):
            buffer_logo = processar_imagem_para_relatorio(caminho_logo, largura_max=500, qualidade=95)
            LARGURA_CAPA = Inches(6.5)
            ALTURA_CAPA = Inches(5.5)
            adicionar_imagem(doc, buffer_logo, largura_in=LARGURA_CAPA, altura_in=ALTURA_CAPA)
        else:
            print(f"⚠️ Imagem de capa '{NOME_LOGO}' não encontrada. Pular logo.")

        adicionar_texto_centralizado(
            doc,
            "PROCESSO DE FISCALIZAÇÃO TÉCNICO-OPERACIONAL DOS TERMINAIS RODOVIÁRIOS INTERMUNICIPAIS CONCEDIDOS À EMPRESA SOCICAM",
        )
        adicionar_texto_centralizado(doc, "CONTRATO DE CONCESSÃO DE SERVIÇO PÚBLICO Nº X.XXX.XXX/XX")
        adicionar_texto_centralizado(doc, "PROCESSO SEI Nº XXXXXXXXXX.XXXXXXXXX/XX")

        doc.add_paragraph()

        texto_data = "Recife, data de assinatura eletrônica"
        adicionar_paragrafo_justificado(doc, texto_data)
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

        
        inserir_quebra_e_sumario(doc)

        
        gerar_secao_introducao(doc, row)
        gerar_secao_objetivo(doc)

        gerar_secao_nao_conformidades_constatadas(doc, row, nao_conformidades_df, FOTOS_DIR)
        gerar_secao_resumo_nao_conformidades(doc, row, nao_conformidades_df)
        gerar_secao_consideracoes_finais(doc, row)

        
        if nao_conformidades_df is not None:
            try:
                gerar_secao_anexo_fotos(doc, row, nao_conformidades_df, CAMINHO_RAIZ_FOTOS)
            except Exception as exc:
                print(f"Erro ao gerar Anexo: {exc}")

       
        nome_arquivo = f"relatorio_{id_fisc}"
        caminho_docx = os.path.join(RELATORIOS_DIR, f"{nome_arquivo}.docx")
        caminho_pdf = os.path.join(RELATORIOS_DIR, f"{nome_arquivo}.pdf")

        doc.save(caminho_docx)
        sucesso_pdf = atualizar_toc_e_converter_para_pdf(caminho_docx, caminho_pdf)

        if os.path.exists(caminho_docx):
            fiscalizacoes_df.at[idx, COLUNA_STATUS] = True

    # Formata coluna "Data" (se existir) para dd/mm/YYYY
    if "Data" in fiscalizacoes_df.columns:
        fiscalizacoes_df["Data"] = pd.to_datetime(fiscalizacoes_df["Data"], errors="coerce").dt.strftime("%d/%m/%Y")

    # Atualiza a planilha (se não estiver em uso)
    if not arquivo_em_uso(CAMINHO_PLANILHA):
        with pd.ExcelWriter(CAMINHO_PLANILHA, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            fiscalizacoes_df.to_excel(writer, sheet_name="Fiscalizações", index=False)
            nao_conformidades_df.to_excel(writer, sheet_name="Não-conformidades ", index=False)

        ajustar_largura_colunas(CAMINHO_PLANILHA)

    print("🎉 Relatórios gerados e planilha atualizada com sucesso.")
    input("\nExecução concluída. Pressione Enter para sair...")
    return


if __name__ == "__main__":
    gerar_relatorio()
