from docx.shared import Pt, RGBColor, Inches
from docx.document import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from typing import Any, Dict, List
import pandas as pd
import re

# Presume-se que 'utils' contém as funções auxiliares necessárias
# A função 'aplicar_estilo_corpo' é crucial para as correções.
from utils import (
    adicionar_paragrafo_justificado,
    adicionar_titulo_secao,
    adicionar_paragrafo_info_compacta,
    aplicar_estilo_corpo,
    aplicar_estilo_titulo,
    encontrar_dados_na_base,
    adicionar_imagem,
    processar_imagem_para_relatorio,
    LARGURA_PADRAO_IN,
    adicionar_duas_imagens_lado_a_lado 
)

# ============================================================
#              CORREÇÃO PRINCIPAL IMPLEMENTADA
# ============================================================

def _limpar_prefixo_id(id_prefix: str, texto_bruto: str) -> str:
    """
    Remove o ID da NC apenas no início da frase, sem apagar textos válidos.
    Elimina duplicações como:
        TIP 2025_04 – 04 – ...
        TIP 2025_06.1 – TIP 06.1 – ...
    """

    if not texto_bruto or str(texto_bruto).lower() == 'nan':
        return "Descrição não disponível."

    desc_str = str(texto_bruto).strip()
    id_prefix = str(id_prefix).strip()

    if not id_prefix:
        return desc_str

    # extrai terminal (TIP)
    terminal_sigla = id_prefix.split(" ")[0]  

    # extrai número base ex: 04, 06.1
    match = re.search(r"(\d+(?:\.\d+)?)$", id_prefix.replace("_", "."))
    base_num = match.group(1) if match else ""

    # prefixos possíveis
    candidatos = [
        re.escape(id_prefix),                    
        re.escape(id_prefix.replace("_", ".")),
        re.escape(terminal_sigla + " " + base_num),
        re.escape(base_num),
        re.escape(terminal_sigla)
    ]

    candidatos = sorted(list(set(candidatos)), key=len, reverse=True)

    # 🔥 CORREÇÃO AQUI
    # Remove apenas o prefixo e UM delimitador (–, -, :)
    pattern = re.compile(rf"^({('|'.join(candidatos))})\s*[-–:]?\s*", re.IGNORECASE)

    desc_str = pattern.sub("", desc_str).strip()

    # Corrige capitalização
    if desc_str and desc_str[0].islower():
        desc_str = desc_str[0].upper() + desc_str[1:]

    return desc_str


# ============================================================
#               INSERÇÃO DE TEXTO DA NÃO CONFORMIDADE
# ============================================================

def _inserir_texto_nc(doc: Document, nc_titulo_identificador: str, descricao_bruta: str):
    descricao_limpa = _limpar_prefixo_id(nc_titulo_identificador, descricao_bruta)

    linhas = [linha for linha in descricao_limpa.split("\n") if linha.strip()]
    if not linhas:
        return

    paragrafo_nc = doc.add_paragraph()

    # Título
    run_titulo = paragrafo_nc.add_run(f"Não Conformidade {nc_titulo_identificador}")
    aplicar_estilo_corpo(run_titulo, negrito=True)
    run_titulo.underline = True

    # Travessão único ✔️
    run_traco = paragrafo_nc.add_run(" – ")
    aplicar_estilo_corpo(run_traco)

    # Primeira linha
    texto_principal = _limpar_prefixo_id(nc_titulo_identificador, linhas[0])
    run_desc = paragrafo_nc.add_run(texto_principal)
    aplicar_estilo_corpo(run_desc)

    paragrafo_nc.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY_LOW

    # Subitens
    if len(linhas) > 1:
        for sub_item_raw in linhas[1:]:
            sub_item = _limpar_prefixo_id(nc_titulo_identificador, sub_item_raw)

            p_sub = doc.add_paragraph()
            p_sub.paragraph_format.left_indent = Inches(0.5)
            p_sub.paragraph_format.space_after = Pt(2)
            p_sub.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY_LOW

            partes = sub_item.split(" – ", 1)
            sub_id = partes[0].strip()
            sub_desc = partes[1].strip() if len(partes) > 1 else ""

            base = nc_titulo_identificador.rsplit("_", 1)[0]
            sub_id_completo = f"{base}_{sub_id}"

            run_prefixo = p_sub.add_run(f"Não Conformidade {sub_id_completo}")
            aplicar_estilo_corpo(run_prefixo, negrito=True)
            run_prefixo.underline = True

            run_sep = p_sub.add_run(" – ")
            aplicar_estilo_corpo(run_sep)

            run_desc = p_sub.add_run(sub_desc)
            aplicar_estilo_corpo(run_desc)


# ============================================================
#                   INFO DA NC
# ============================================================

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


# ============================================================
#                 SPLIT SEGURO DE STRINGS
# ============================================================

def _safe_split(texto: Any) -> List[str]:
    s = str(texto).strip()
    if not s or s.lower() == "nan":
        return []
    s = s.replace(":", ";")
    return [x.strip() for x in s.split(";") if x.strip()]


# ============================================================
#         GERAÇÃO COMPLETA DA SEÇÃO 3 DO RELATÓRIO
# ============================================================

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

    texto_cartas = f"constante da Carta SAP/PER/ARPE N° {cartas_raw}, " if cartas_raw else ""

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

    grupo = nao_conformidades_df[
        nao_conformidades_df["ID da Fiscalização"] == id_fiscalizacao
    ].copy()

    num_terminal = 1
    for terminal, dados_terminal in grupo.groupby("Terminal"):

        p = doc.add_paragraph()
        run = p.add_run(f"3.{num_terminal} - {terminal.upper()}")
        aplicar_estilo_titulo(run)
        p.paragraph_format.space_before = Pt(12)

        ncs_processadas = set()

        for _, nc in dados_terminal.iterrows():

            raw_key = str(nc.get("Legenda da Foto", "")).strip()
            if not raw_key or raw_key.lower() == "nan":
                raw_key = str(nc.get("Constatação", "")).strip()

            constatacoes = _safe_split(raw_key)
            infos = _safe_split(nc.get("Informação SOCICAM", ""))
            analises = _safe_split(nc.get("Análise da Arpe", ""))

            max_len = max(len(constatacoes), len(infos), len(analises))
            constatacoes += [""] * (max_len - len(constatacoes))
            infos += ["N/A"] * (max_len - len(infos))
            analises += ["N/A"] * (max_len - len(analises))

            for i in range(max_len):
                const = constatacoes[i]
                if not const:
                    continue

                dados = encontrar_dados_na_base(
                    const,
                    df_base_nc,
                    ano_user,
                    proc_user,
                    monit_user,
                    terminal_user=str(terminal)
                )

                if dados["id"] != "ID_NAO_ENCONTRADO":
                    if dados["id"] in ncs_processadas:
                        continue
                    ncs_processadas.add(dados["id"])

                _inserir_texto_nc(doc, dados["id"], dados["descricao"])
                _inserir_linhas_info(doc, infos[i], const, analises[i])
                doc.add_paragraph().paragraph_format.space_after = Pt(12)

        num_terminal += 1
