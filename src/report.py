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
import re  # IMPORT NECESSÁRIO PARA A VALIDAÇÃO DE PROCESSO

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
    """Retorna o diretório base do script, mesmo que esteja empacotado (frozen)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def atualizar_toc_e_converter_para_pdf(caminho_docx: str, caminho_pdf: str) -> bool:
    """Atualiza o Sumário (TOC) e converte o DOCX para PDF usando a aplicação Word (requer Windows/MS Word)."""
    word = None
    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(caminho_docx)
        doc.Fields.Update()
        doc.Save()
        doc.SaveAs2(caminho_pdf, FileFormat=17) # 17 é o código para PDF
        doc.Close(SaveChanges=False)
        word.Quit()
        return True
    except Exception as exc:
        if word is not None:
            word.DisplayAlerts = -1
            word.Quit()
        print(f"❌ ERRO ao gerar PDF (win32com): {exc}")
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

    try:
        # Carrega os DataFrames necessários para a lógica
        monitoramento_df = pd.read_excel(CAMINHO_PLANILHA, sheet_name="Monitoramento")
        nao_conformidades_df = pd.read_excel(CAMINHO_PLANILHA, sheet_name="Não-conformidades ")
        processos_df = pd.read_excel(CAMINHO_PLANILHA, sheet_name="Processos")
        
        monitoramento_df.columns = monitoramento_df.columns.str.strip()
        nao_conformidades_df.columns = nao_conformidades_df.columns.str.strip()
        processos_df.columns = processos_df.columns.str.strip()

    except Exception as e:
        print(f"❌ Erro ao ler planilha: {e}")
        sys.exit(1)

    # ============================================================
    # 🔍 ETAPA 1 — VERIFICAR PENDENTES
    # ============================================================

    # Garantir coluna STATUS
    if COLUNA_STATUS not in monitoramento_df.columns:
        monitoramento_df[COLUNA_STATUS] = ""

    monitoramento_df[COLUNA_STATUS] = monitoramento_df[COLUNA_STATUS].astype(str)

    mascara_concluido = (
        monitoramento_df[COLUNA_STATUS]
        .str.lower()
        .str.strip()
        .isin(["true", "1", "sim", "verdadeiro", "ok"])
    )

    pendentes = monitoramento_df[~mascara_concluido].copy()

    # Se NÃO existem pendentes → ENCERRAR
    if pendentes.empty:
        print("\n✅ Todos os relatórios já foram gerados anteriormente, por favor, verifique a planilha.")
        input("Pressione ENTER para sair...")
        return
        
    # --- Se há pendentes, as mensagens de carregamento de dados APENAS são exibidas aqui
    print("\n--- Carregando Dados da Planilha Principal ---")
    print(f"\n🚀 Existem {len(pendentes)} relatórios pendentes para gerar.\n")

    # --- CARREGAR BASE NC (SÓ SE HÁ PENDENTES) ---
    print("\n--- Carregando Base de Não Conformidades ---")
    if os.path.exists(CAMINHO_BASE_NC):
        # A chamada a carregar_base_nc JÁ IMPRIME A MENSAGEM de "Base carregada" (arquivo utils.py)
        df_base_nc = carregar_base_nc(CAMINHO_BASE_NC)
        # A linha duplicada que causava o problema foi removida daqui.
    else:
        print(f"⚠️ Planilha Base não encontrada em: {CAMINHO_BASE_NC}")
        df_base_nc = pd.DataFrame()


    # ======================================
    # 🔵 ETAPA 2 — COLETAR E VALIDAR DADOS INSTANTANEAMENTE
    # ======================================

    print("\n📋 CONFIGURAÇÃO DE FILTROS DA BASE DE DADOS:")

    # 1. COLETAR E VALIDAR ANO
    while True:
        ano_input = input(">> 1. Digite o ANO (ex: 2025): ").strip()
        if not ano_input: continue
        
        # VALIDAÇÃO DO ANO
        if not df_base_nc.empty and str(ano_input) not in df_base_nc["Ano"].astype(str).str.strip().unique():
            print(f"❌ O ANO '{ano_input}' não consta na Base de Não Conformidades. Pressione ENTER para sair...")
            input()
            sys.exit(1)
        break

    # 2. COLETAR E VALIDAR PROCESSO
    while True:
        proc_input = input(">> 2. Digite o PROCESSO (ex: CTR 01/2025 ou ctr 01/2025): ").strip()
        if not proc_input: continue
        
        # VALIDAÇÃO DO PROCESSO 
        if not df_base_nc.empty:
            # Lógica de limpeza para comparação
            proc_clean_user = re.sub(r'(CTR\s*Nº?|Nº?|:)\s*', ' ', proc_input.upper().strip())
            proc_clean_user = re.sub(r'\s+', ' ', proc_clean_user).strip()
            
            df_base_nc_copy = df_base_nc.copy()
            df_base_nc_copy['PROCESSO_LIMPO'] = df_base_nc_copy["PROCESSO"].astype(str).str.upper().str.strip()
            df_base_nc_copy['PROCESSO_LIMPO'] = df_base_nc_copy['PROCESSO_LIMPO'].apply(lambda t: re.sub(r'(CTR\s*Nº?|Nº?|:)\s*', ' ', t))
            df_base_nc_copy['PROCESSO_LIMPO'] = df_base_nc_copy['PROCESSO_LIMPO'].apply(lambda t: re.sub(r'\s+', ' ', t).strip())
            
            mask_proc = df_base_nc_copy['PROCESSO_LIMPO'].str.contains(proc_clean_user, na=False)

            if not mask_proc.any():
                print(f"❌ O PROCESSO '{proc_input}' não consta na Base de Não Conformidades. Pressione ENTER para sair...")
                input()
                sys.exit(1)
        break

    # 3. COLETAR E VALIDAR MONITORAMENTO
    while True:
        monit_input = input(">> 3. Qual é o Nº do Monitoramento? (ex: 1): ").strip()
        if not monit_input: continue
        
        # VALIDAÇÃO DO MONITORAMENTO
        if not df_base_nc.empty and "TIPO_DOC" in df_base_nc.columns:
            mask_monit_num = (df_base_nc["TIPO_DOC"].str.upper().str.contains("MONIT", na=False)) & \
                             (df_base_nc["TIPO_DOC"].astype(str).str.contains(str(monit_input), na=False))
            
            if not mask_monit_num.any():
                print(f"❌ O Nº do Monitoramento '{monit_input}' não consta na Base de Não Conformidades. Pressione ENTER para sair...")
                input()
                sys.exit(1)
        break
    
    # =====================================================
    # 📂 ETAPA 2.2 — VALIDAÇÃO DOS CAMINHOS DE PASTAS
    # =====================================================

    print("\n--- Configuração de Pastas ---")

    # VALIDAÇÃO PASTA DE FISCALIZAÇÃO (CAMINHO_CTR)
    CAMINHO_CTR = ""
    while True:
        pasta_contrato_input = input(">> Digite a pasta de Fiscalização (Ex: CTR-01-2025 ou ctr-01-2025): ").strip()
        if not pasta_contrato_input:
            continue
        
        CAMINHO_CTR = os.path.join(FOTOS_DIR, pasta_contrato_input.upper())
        
        if not os.path.isdir(CAMINHO_CTR):
            print(f"❌ Essa Pasta ('{pasta_contrato_input}') não existe, verifique seus arquivos. Pressione ENTER para sair...")
            input()
            sys.exit(1) 
            
        if not os.listdir(CAMINHO_CTR):
            print(f"❌ Essa Pasta ('{pasta_contrato_input}') está vazia, verifique seus arquivos. Pressione ENTER para sair...")
            input()
            sys.exit(1)

        break

    # VALIDAÇÃO PASTA DO MONITORAMENTO (CAMINHO_RAIZ_FOTOS)
    CAMINHO_RAIZ_FOTOS = ""
    while True:
        pasta_monitoramento_input = input(">> Digite a Pasta do Monitoramento (Ex: M0 ou m0): ").strip()
        if not pasta_monitoramento_input:
            continue
        
        CAMINHO_RAIZ_FOTOS = os.path.join(CAMINHO_CTR, pasta_monitoramento_input.upper())
        
        if not os.path.isdir(CAMINHO_RAIZ_FOTOS):
            print(f"❌ Pasta não encontrada: {CAMINHO_RAIZ_FOTOS}")
            continue
        
        if not os.listdir(CAMINHO_RAIZ_FOTOS):
            print(f"⚠️ Essa Subpasta ('{pasta_monitoramento_input}') está vazia, seu relatório de monitoramento não terá fotos.")
            input("Pressione ENTER para continuar (sem fotos) ou CTRL+C para sair...")
        
        break

    # =====================================================
    # 🔄 ETAPA 3 — LOOP PARA GERAR RELATÓRIOS PENDENTES
    # =====================================================

    print(f"\n🚀 Gerando {len(pendentes)} relatórios pendentes...\n")

    for idx in tqdm(pendentes.index, desc="Progresso"):
        row = monitoramento_df.loc[idx]
        id_fisc = row["ID da Fiscalização"]

        processo_info_filtered = processos_df[processos_df["ID da Fiscalização"] == id_fisc]
        processo_info = processo_info_filtered.iloc[0].to_dict() if not processo_info_filtered.empty else {}

        doc = Document()
        doc.sections[0].top_margin = Inches(0.25)

        # CAPA
        adicionar_texto_centralizado(doc, "COORDENADORIA DE TRANSPORTES E RODOVIAS")
        doc.paragraphs[-1].paragraph_format.space_after = Pt(0)
        doc.add_paragraph()

        adicionar_texto_centralizado(
            doc,
            f"RELATÓRIO DO {id_fisc}º MONITORAMENTO DAS NÃO CONFORMIDADES DO PROCESSO CTR Nº {processo_info.get('Processo CTR Nº', 'XX/XXXX')}"
        )
        doc.add_paragraph()

        caminho_logo = os.path.join(FOTOS_DIR, "capa_monitoramento_arpe.jpg")
        if os.path.exists(caminho_logo):
            buf = processar_imagem_para_relatorio(caminho_logo, largura_max=500)
            adicionar_imagem(doc, buf, largura_in=Inches(6.5), altura_in=Inches(5.5))

        adicionar_texto_centralizado(
            doc,
            "PROCESSO DE FISCALIZAÇÃO TÉCNICO-OPERACIONAL DOS TERMINAIS RODOVIÁRIOS INTERMUNICIPAIS CONCEDIDOS À EMPRESA SOCICAM"
        )
        adicionar_texto_centralizado(doc, f"CONTRATO Nº {processo_info.get('Contrato de Concessão Nº', '')}")
        adicionar_texto_centralizado(doc, f"PROCESSO SEI Nº {processo_info.get('Processo SEI Nº', '')}")
        doc.add_paragraph()
        adicionar_texto_centralizado(doc, "Recife, data de assinatura eletrônica", negrito=False)

        # SUMÁRIO + SEÇÕES
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
                tqdm.write(f"⚠️ Erro ao gerar anexos: {exc}")

        # SALVAR DOCX + PDF
        nome_arquivo = f"relatorio_{id_fisc}"
        caminho_docx = os.path.join(RELATORIOS_DIR, f"{nome_arquivo}.docx")
        caminho_pdf = os.path.join(RELATORIOS_DIR, f"{nome_arquivo}.pdf")

        doc.save(caminho_docx)
        sucesso = atualizar_toc_e_converter_para_pdf(caminho_docx, caminho_pdf)

        if os.path.exists(caminho_docx) and sucesso:
            monitoramento_df.at[idx, COLUNA_STATUS] = "VERDADEIRO"

    # =====================================================
    # 💾 ETAPA FINAL — SALVAR ALTERAÇÕES NA PLANILHA
    # =====================================================

    if not arquivo_em_uso(CAMINHO_PLANILHA):
        with pd.ExcelWriter(CAMINHO_PLANILHA, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            monitoramento_df.to_excel(writer, sheet_name="Monitoramento", index=False)
            nao_conformidades_df.to_excel(writer, sheet_name="Não-conformidades ", index=False)
            processos_df.to_excel(writer, sheet_name="Processos", index=False)

        ajustar_largura_colunas(CAMINHO_PLANILHA)
        print("\n🎉 Relatórios gerados com sucesso e planilha atualizada!")
    else:
        print("\n❌ A planilha de monitoramento está em uso — alterações de status (Relatório Gerado) não salvas.")

    input("Pressione ENTER para sair...")

if __name__ == "__main__":
    gerar_relatorio()