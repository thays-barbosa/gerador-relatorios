from typing import Optional, Dict
import os
import sys
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pandas as pd
from tqdm import tqdm
import win32com.client as win32

# --- SEÇÕES ---
from sections.introduction.introduction import gerar_secao_introducao
from sections.objective.objective import gerar_secao_objetivo
from sections.anexo.anexo import gerar_secao_anexo_fotos
from sections.nonconformity.nonconformity import gerar_secao_nao_conformidades_constatadas
from sections.nonconformityresume.nonconformityresume import gerar_secao_resumo_nao_conformidades
from sections.finalconsiderations.finalconsiderations import gerar_secao_consideracoes_finais
from sections.summary.summary import inserir_quebra_e_sumario

# --- UTILS ---
from utils import (
    adicionar_texto_centralizado,
    adicionar_paragrafo_justificado,
    ajustar_largura_colunas,
    arquivo_em_uso,
    processar_imagem_para_relatorio,
    adicionar_imagem,
    carregar_base_nc
)

def _obter_base_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def atualizar_toc_e_converter_para_pdf(caminho_docx: str, caminho_pdf: str) -> bool:
    word = None
    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(caminho_docx)
        doc.Fields.Update()
        doc.Save()
        doc.SaveAs2(caminho_pdf, FileFormat=17)
        doc.Close(SaveChanges=False)
        word.Quit()
        return True
    except Exception as exc:
        if word is not None:
            word.DisplayAlerts = -1
            word.Quit()
        print(f"❌ ERRO ao gerar PDF: {exc}")
        return False

def gerar_relatorio() -> None:
    BASE_DIR = _obter_base_dir()
    FOTOS_DIR = os.path.join(BASE_DIR, "assets")
    RELATORIOS_DIR = os.path.join(BASE_DIR, "reports")
    CAMINHO_PLANILHA = os.path.join(BASE_DIR, "planilha_monitoramento.xlsx")
    CAMINHO_BASE_NC = os.path.join(BASE_DIR, "Levantamento de NC 2020-2025 - SOCICAM.xlsx")
    COLUNA_STATUS = "Relatório Gerado"

    os.makedirs(RELATORIOS_DIR, exist_ok=True)
    os.makedirs(FOTOS_DIR, exist_ok=True)

    if arquivo_em_uso(CAMINHO_PLANILHA):
        print("⚠️ A planilha de monitoramento está em uso. Feche-a antes de executar.")
        sys.exit(1)

    print("\n--- Carregando Dados ---")
    try:
        monitoramento_df = pd.read_excel(CAMINHO_PLANILHA, sheet_name="Monitoramento")
        nao_conformidades_df = pd.read_excel(CAMINHO_PLANILHA, sheet_name="Não-conformidades ")
        processos_df = pd.read_excel(CAMINHO_PLANILHA, sheet_name="Processos")
        
        monitoramento_df.columns = monitoramento_df.columns.str.strip()
        nao_conformidades_df.columns = nao_conformidades_df.columns.str.strip()
        processos_df.columns = processos_df.columns.str.strip()
    except Exception as e:
        print(f"❌ Erro ao ler planilha: {e}")
        sys.exit(1)

    if os.path.exists(CAMINHO_BASE_NC):
        df_base_nc = carregar_base_nc(CAMINHO_BASE_NC)
    else:
        print(f"⚠️ Planilha Base não encontrada em: {CAMINHO_BASE_NC}")
        df_base_nc = pd.DataFrame()

    # --- PERGUNTAS ---
    print("\n📋 CONFIGURAÇÃO DE FILTROS DA BASE DE DADOS:")
    ano_input = input(">> 1. Digite o ANO (ex: 2025): ").strip()
    proc_input = input(">> 2. Digite o PROCESSO (ex: CTR 04/2025): ").strip()
    monit_input = input(">> 3. Qual é o Nº do Monitoramento? (ex: 1): ").strip()
    
    print("\n--- Configuração de Pastas ---")
    while True:
        pasta_contrato_input = input(">> Digite a pasta de Fiscalização (Ex: CTR-01-2025): ").strip()
        if not pasta_contrato_input: continue
        CAMINHO_CTR = os.path.join(FOTOS_DIR, pasta_contrato_input.upper())
        if os.path.isdir(CAMINHO_CTR): break
        else: print(f"❌ Pasta não encontrada: {CAMINHO_CTR}")

    while True:
        pasta_monitoramento_input = input(">> Digite a Pasta do Monitoramento (Ex: M0): ").strip()
        if not pasta_monitoramento_input: continue
        CAMINHO_RAIZ_FOTOS = os.path.join(CAMINHO_CTR, pasta_monitoramento_input.upper())
        if os.path.isdir(CAMINHO_RAIZ_FOTOS): break
        else: print(f"❌ Pasta não encontrada: {CAMINHO_RAIZ_FOTOS}")

    # --- FILTRO DE PENDENTES (Lógica robusta) ---
    
    # 1. Se a coluna não existe, a inicializamos como string vazia (tipo 'object')
    if COLUNA_STATUS not in monitoramento_df.columns:
        monitoramento_df[COLUNA_STATUS] = ""
    
    # 2. Garante que a coluna é do tipo 'object' (string) para evitar o FutureWarning ao salvar "VERDADEIRO"
    # Este passo é crucial para o salvamento final.
    monitoramento_df[COLUNA_STATUS] = monitoramento_df[COLUNA_STATUS].astype(str)

    # 3. Cria a MÁSCARA booleana (True para concluído, False para pendente)
    mascara_concluido = (
        monitoramento_df[COLUNA_STATUS]
        .str.lower()
        .str.strip()
        .isin(['true', '1', 'sim', 'verdadeiro', 'ok'])
    )

    # 4. Filtra os pendentes (Onde a máscara NÃO é True)
    pendentes = monitoramento_df[~mascara_concluido].copy()


    if pendentes.empty:
        print("\n✅ Todos os relatórios já foram gerados.")
        input("Enter para sair...")
        return

    print(f"\n🚀 Gerando {len(pendentes)} relatórios pendentes...")

    # --- LOOP ---
    for idx in tqdm(pendentes.index, desc="Progresso"):
        row = monitoramento_df.loc[idx]
        id_fisc = row["ID da Fiscalização"]

        processo_info_filtered = processos_df[processos_df["ID da Fiscalização"] == id_fisc]
        processo_info = processo_info_filtered.iloc[0].to_dict() if not processo_info_filtered.empty else {}

        doc = Document()
        doc.sections[0].top_margin = Inches(0.25)

        # Capa
        adicionar_texto_centralizado(doc, "COORDENADORIA DE TRANSPORTES E RODOVIAS")
        doc.paragraphs[-1].paragraph_format.space_after = Pt(0)
        doc.add_paragraph()
        adicionar_texto_centralizado(doc, f"RELATÓRIO DO {id_fisc}º MONITORAMENTO DAS NÃO CONFORMIDADES DO PROCESSO CTR Nº {processo_info.get('Processo CTR Nº', 'XX/XXXX')}")
        doc.add_paragraph()

        caminho_logo = os.path.join(FOTOS_DIR, "capa_monitoramento_arpe.jpg")
        if os.path.exists(caminho_logo):
            buf = processar_imagem_para_relatorio(caminho_logo, largura_max=500)
            adicionar_imagem(doc, buf, largura_in=Inches(6.5), altura_in=Inches(5.5))
        
        adicionar_texto_centralizado(doc, "PROCESSO DE FISCALIZAÇÃO TÉCNICO-OPERACIONAL DOS TERMINAIS RODOVIÁRIOS INTERMUNICIPAIS CONCEDIDOS À EMPRESA SOCICAM")
        adicionar_texto_centralizado(doc, f"CONTRATO DE CONCESSÃO DE SERVIÇO PÚBLICO Nº {processo_info.get('Contrato de Concessão Nº', '')}")
        adicionar_texto_centralizado(doc, f"PROCESSO SEI Nº {processo_info.get('Processo SEI Nº', '')}")
        doc.add_paragraph()
        adicionar_texto_centralizado(doc, "Recife, data de assinatura eletrônica", negrito=False)

        # Seções
        inserir_quebra_e_sumario(doc)
        gerar_secao_introducao(doc, row)
        gerar_secao_objetivo(doc, row, processo_info, nao_conformidades_df)
        
        gerar_secao_nao_conformidades_constatadas(
            doc, row, nao_conformidades_df, FOTOS_DIR, processo_info, 
            df_base_nc=df_base_nc, ano_user=ano_input, proc_user=proc_input, monit_user=monit_input
        )
        
        gerar_secao_resumo_nao_conformidades(
            doc, row, nao_conformidades_df, processo_info, 
            df_base_nc=df_base_nc, ano_user=ano_input, proc_user=proc_input, monit_user=monit_input
        )
        
        gerar_secao_consideracoes_finais(doc, row, nao_conformidades_df, processo_info)

        if not nao_conformidades_df.empty:
            try:
                gerar_secao_anexo_fotos(
                    doc, row, nao_conformidades_df, CAMINHO_RAIZ_FOTOS, processo_info, 
                    df_base_nc=df_base_nc, ano_user=ano_input, proc_user=proc_input, monit_user=monit_input
                )
            except Exception as exc:
                tqdm.write(f"⚠️ Erro Anexo: {exc}")

        # Salvar e Converter
        nome_arquivo = f"relatorio_{id_fisc}"
        caminho_docx = os.path.join(RELATORIOS_DIR, f"{nome_arquivo}.docx")
        caminho_pdf = os.path.join(RELATORIOS_DIR, f"{nome_arquivo}.pdf")
        doc.save(caminho_docx)
        sucesso = atualizar_toc_e_converter_para_pdf(caminho_docx, caminho_pdf)
        
        if os.path.exists(caminho_docx) and sucesso:
            # Salvamento da string "VERDADEIRO"
            monitoramento_df.at[idx, COLUNA_STATUS] = "VERDADEIRO" 

    if not arquivo_em_uso(CAMINHO_PLANILHA):
        with pd.ExcelWriter(CAMINHO_PLANILHA, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            monitoramento_df.to_excel(writer, sheet_name="Monitoramento", index=False)
            nao_conformidades_df.to_excel(writer, sheet_name="Não-conformidades ", index=False)
            processos_df.to_excel(writer, sheet_name="Processos", index=False)
        ajustar_largura_colunas(CAMINHO_PLANILHA)
        print("\n🎉 Concluído!")
    else:
        print("\n❌ Planilha em uso. Status não salvo.")
    input("Enter para sair...")

if __name__ == "__main__":
    gerar_relatorio()