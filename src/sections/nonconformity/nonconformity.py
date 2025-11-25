from docx.shared import Pt, RGBColor, Inches
from docx.document import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from typing import Any, Dict, List
import pandas as pd
import re

from utils import (
    adicionar_paragrafo_justificado,
    adicionar_titulo_secao,
    adicionar_paragrafo_info_compacta,
    aplicar_estilo_corpo,
    aplicar_estilo_titulo,
    encontrar_dados_na_base,
    LARGURA_PADRAO_IN
)

def _limpar_inicio_texto(texto: str) -> str:
    if not texto: return ""
    s = str(texto).strip()
    padrao = r'^(?:[A-Z]{3}[\s\-]*)?\d+(?:[._]\d+)?\s*[-:–)]*\s*'
    texto_limpo = re.sub(padrao, '', s)
    if not texto_limpo and s:
        return s
    if texto_limpo and texto_limpo[0].islower():
        return texto_limpo[0].upper() + texto_limpo[1:]
    return texto_limpo

def _inserir_texto_nc(doc: Document, nc_titulo_identificador: str, descricao_bruta: str):
    if not descricao_bruta: 
        descricao_bruta = "Descrição indisponível."
    linhas = [l.strip() for l in descricao_bruta.split("\n") if l.strip()]
    if not linhas: linhas = ["Descrição indisponível."]

    paragrafo_nc = doc.add_paragraph()
    run_titulo = paragrafo_nc.add_run(f"Não Conformidade {nc_titulo_identificador}")
    aplicar_estilo_corpo(run_titulo, negrito=True)
    run_titulo.underline = True

    run_traco = paragrafo_nc.add_run(" - ") 
    aplicar_estilo_corpo(run_traco)

    texto_principal = _limpar_inicio_texto(linhas[0])
    run_desc = paragrafo_nc.add_run(texto_principal)
    aplicar_estilo_corpo(run_desc)

    paragrafo_nc.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY_LOW

    base_prefixo = ""
    match_base = re.match(r"([A-Z]{3}\s\d{4}_)", nc_titulo_identificador)
    if match_base:
        base_prefixo = match_base.group(1)

    if len(linhas) > 1:
        for sub_item_raw in linhas[1:]:
            p_sub = doc.add_paragraph()
            p_sub.paragraph_format.left_indent = Inches(0.5)
            p_sub.paragraph_format.space_after = Pt(2)
            p_sub.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY_LOW
            match_sub = re.match(r'^(?:[A-Z]{3}[\s\-]*)?(\d+[._]\d+)\s*[-:–]?\s*(.*)', sub_item_raw)
            if match_sub:
                numero_sub = match_sub.group(1).replace("_", ".") 
                texto_restante = match_sub.group(2)
                texto_final = _limpar_inicio_texto(texto_restante)
                id_completo = f"{base_prefixo}{numero_sub}" if base_prefixo else f"Item {numero_sub}"
                run_sub_id = p_sub.add_run(f"Não Conformidade {id_completo}")
                aplicar_estilo_corpo(run_sub_id, negrito=True)
                
                p_sub.add_run(" - ") 
                
                run_txt = p_sub.add_run(texto_final)
                aplicar_estilo_corpo(run_txt)
            else:
                texto_limpo = _limpar_inicio_texto(sub_item_raw)
                run_txt = p_sub.add_run(texto_limpo)
                aplicar_estilo_corpo(run_txt)

def _inserir_linhas_info(doc: Document, info_socicam: str, constatacao_monit: str, analise_arpe: str):
    for var_name, value in [
        ("info_socicam", info_socicam),
        ("constatacao_monit", constatacao_monit),
        ("analise_arpe", analise_arpe)
    ]:
        if str(value).strip().lower() in ["nan", "none", ""]:
            locals()[var_name] = "N/A"
    
    adicionar_paragrafo_info_compacta(doc, "Informação da SOCICAM: ", str(info_socicam))
    adicionar_paragrafo_info_compacta(doc, "Constatação: ", str(constatacao_monit))
    adicionar_paragrafo_info_compacta(doc, "Análise da ARPE: ", str(analise_arpe))

def _safe_split_blocos(texto: Any) -> List[str]:
    s = str(texto).strip()
    if not s or s.lower() == "nan": return []
    partes = re.split(r'[;\n]', s)
    return [x.strip() for x in partes if x.strip()]

def gerar_secao_nao_conformidades_constatadas(
    doc: Document,
    row: pd.Series,
    nao_conformidades_df: pd.DataFrame,
    FOTOS_DIR: str,
    processo_info: Dict[str, str],
    df_base_nc: pd.DataFrame,
    ano_user: str,
    proc_user: str,
    monit_user: str
):
    id_fiscalizacao = row["ID da Fiscalização"]
    processo_ctr = processo_info.get("Processo CTR Nº", "XX/XXXX")
  
    cartas_raw = str(processo_info.get("Carta SAP/PER/ARPE Nº", "")).strip()
    texto_cartas = ""
    
    if cartas_raw and cartas_raw.lower() != "nan":
        partes_cartas = [p.strip() for p in cartas_raw.split(";") if p.strip()]
        if partes_cartas:
            cartas_formatadas = [f"Carta SAP/PER/ARPE N° {p}" for p in partes_cartas]
            juncao_cartas = " e ".join(cartas_formatadas)
            texto_cartas = f"constante da {juncao_cartas}, "

    adicionar_titulo_secao(doc, "3. RESULTADO DAS VISTORIAS DAS NÃO CONFORMIDADES PENDENTES")

    adicionar_paragrafo_justificado(
        doc,
        (
            "Estão registrados para cada Terminal Rodoviário os resultados da verificação pela Arpe das ações "
            f"desenvolvidas pela SOCICAM, {texto_cartas}"
            "para solucionar as Não Conformidades ainda pendentes apresentadas no Relatório de Fiscalização Técnico-"
            f"Operacional ARPE/CTR nº {processo_ctr}."
        ),
    )

    grupo = nao_conformidades_df[nao_conformidades_df["ID da Fiscalização"] == id_fiscalizacao].copy()
    grupos_ordenados = sorted(grupo.groupby("Terminal"), key=lambda x: str(x[0]))
    
    num_terminal = 1
    for terminal, dados_terminal in grupos_ordenados:
        p = doc.add_paragraph()
        run = p.add_run(f"3.{num_terminal} - {str(terminal).upper()}")
        aplicar_estilo_titulo(run)
        p.paragraph_format.space_before = Pt(12)

        ncs_processadas = set()
        # Lista temporária para garantir a ORDENAÇÃO
        lista_ncs_para_imprimir = []

        for _, nc in dados_terminal.iterrows():
            raw_key_busca = str(nc.get("Legenda da Foto", "")).strip()
            if not raw_key_busca or raw_key_busca.lower() == "nan":
                raw_key_busca = str(nc.get("Constatação", "")).strip()

            raw_constatacao_coluna = str(nc.get("Constatação", "")).strip()
            if raw_constatacao_coluna.lower() == "nan": 
                blocos_exibicao = [] 
            else:
                blocos_exibicao = _safe_split_blocos(raw_constatacao_coluna)

            blocos_busca = _safe_split_blocos(raw_key_busca)
            infos = _safe_split_blocos(nc.get("Informação SOCICAM", ""))
            analises = _safe_split_blocos(nc.get("Análise da Arpe", ""))

            max_len = len(blocos_busca)
            infos += ["N/A"] * (max_len - len(infos))
            analises += ["N/A"] * (max_len - len(analises))

            for i in range(max_len):
                texto_busca = blocos_busca[i]
                if not texto_busca: continue

                texto_pre_limpo = texto_busca.split(':')[0].strip()
                texto_para_busca = re.sub(r'\b(?:em\s+)?\d{2}/\d{2}/\d{4}\b', '', texto_pre_limpo, flags=re.IGNORECASE).strip()
                texto_para_busca = re.sub(r'\s+', ' ', texto_para_busca).strip()

                dados = encontrar_dados_na_base(
                    texto_para_busca, df_base_nc, ano_user, proc_user, monit_user, terminal_user=str(terminal)
                )

                if dados["id"] != "ID_NAO_ENCONTRADO":
                    if dados["id"] in ncs_processadas: continue
                    ncs_processadas.add(dados["id"])

                # Lógica Estrita de Exibição
                if i < len(blocos_exibicao):
                    texto_final_relatorio = blocos_exibicao[i]
                    if not texto_final_relatorio:
                        texto_final_relatorio = "N/A"
                else:
                    texto_final_relatorio = "N/A"
                
                # EM VEZ DE IMPRIMIR, ADICIONA NA LISTA
                lista_ncs_para_imprimir.append({
                    "id": dados["id"],
                    "desc_base": dados["descricao"],
                    "info": infos[i],
                    "constatacao_real": texto_final_relatorio,
                    "analise": analises[i]
                })

        # ORDENAÇÃO: Ordena a lista pelo ID (ex: TIP 2025_01 antes de TIP 2025_03)
        lista_ncs_para_imprimir.sort(key=lambda x: x["id"])

        # AGORA SIM, IMPRIME NO WORD NA ORDEM CERTA
        for item in lista_ncs_para_imprimir:
            _inserir_texto_nc(doc, item["id"], item["desc_base"])
            _inserir_linhas_info(doc, item["info"], item["constatacao_real"], item["analise"])
            doc.add_paragraph().paragraph_format.space_after = Pt(12)

        num_terminal += 1