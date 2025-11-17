"""
anexo.py

Gera a seção 'ANEXO - MEMORIAL FOTOGRÁFICO' do relatório, buscando fotos no disco
com base no 'ID da não conformidade' presente na planilha.
"""

from typing import List, Optional
from itertools import zip_longest
import os
import re

from docx.document import Document
from pandas.core.series import Series as Row
from pandas import DataFrame
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

from utils import (
    adicionar_titulo_secao,
    adicionar_duas_imagens_lado_a_lado,
    aplicar_estilo_corpo,
    _adicionar_contexto_nc,
)

CHAVE_ID_ARPE = "ID da não conformidade"
CHAVE_NC_DESC = "Não Conformidade"


def _safe_split_anexo(content: str) -> List[str]:
    """
    Faz um split "seguro" do conteúdo vindo da planilha de anexo.
    - Se content não for str ou for vazio/NaN, retorna lista vazia.
    - Prioriza ';' para separar NCs principais; se não houver, usa ':'.
    """
    if not isinstance(content, str) or content.lower() == "nan" or not content.strip():
        return []

    # Prioriza ';' para separação principal (NCs)
    if ";" in content:
        return [d.strip() for d in content.split(";") if d.strip()]

    # Caso contrário, usa ':' (usado para separar legendas dentro de uma NC)
    return [d.strip() for d in content.split(":") if d.strip()]


def _adaptar_lista_anexo(lista: List[str], tamanho_alvo: int, default_val: str) -> List[str]:
    """
    Assegura que 'lista' tenha tamanho == tamanho_alvo.
    - Se maior, trunca.
    - Se menor, estende com default_val.
    """
    if len(lista) == tamanho_alvo:
        return lista

    if len(lista) > tamanho_alvo:
        return lista[:tamanho_alvo]

    nova = list(lista)
    faltando = tamanho_alvo - len(lista)
    nova.extend([default_val] * faltando)
    return nova


def _limpar_id_para_busca(nc_id: str) -> str:
    """
    Normaliza o ID da NC para servir como prefixo de busca de fotos.
    - Mantém IDs já no formato com underscores e números quando detectado.
    - Substitui espaços, hífens, pontos e barras por underscore e elimina duplos.
    """
    if not isinstance(nc_id, str):
        return ""

    parts = nc_id.split("_")
    if len(parts) >= 3 and parts[-2].isdigit() and parts[-1].isdigit():
        return nc_id.strip("_").upper()

    nc_id_norm = nc_id.replace(" ", "_").replace("-", "_").replace(".", "_").replace("/", "_")
    while "__" in nc_id_norm:
        nc_id_norm = nc_id_norm.replace("__", "_")

    return nc_id_norm.strip("_").upper()


def _formatar_periodos(periodo_raw: Optional[str]) -> Optional[str]:
    """
    Recebe a string bruta da coluna 'Periodo de Vistoria da ARPE' (ex: '01 a 22/09/2025;30/09/2025')
    e retorna uma string formatada pronta para inserir no título, ex:
    '01 A 22/09/2025 e 30/09/2025'
    """
    if not isinstance(periodo_raw, str) or not periodo_raw.strip() or periodo_raw.lower() == "nan":
        return None

    partes = [p.strip() for p in periodo_raw.split(";") if p.strip()]
    # Normaliza cada parte: maiúsculas e troca ' a ' por ' A ' (garantindo espaçamento)
    partes_norm = []
    for p in partes:
        # remove espaços duplicados
        p_clean = re.sub(r"\s+", " ", p)
        # substituir ' a ' ou ' A ' por ' A ' (maiúsculo)
        p_clean = re.sub(r"\s+[aA]\s+", " A ", p_clean)
        partes_norm.append(p_clean.upper())

    # Junta com ' e ' conforme solicitado (exemplo do usuário)
    if not partes_norm:
        return None
    if len(partes_norm) == 1:
        return partes_norm[0]
    return " e ".join(partes_norm)


def _extrair_cidades_desde_nc(nc_fisc_df: DataFrame) -> List[str]:
    """
    Extrai a lista de cidades/nomes a partir da coluna 'Terminal' no DataFrame de NCs.
    Remove 'Terminal de' (qualquer case) e remove a sigla entre parênteses.
    Retorna lista única e ordenada (por ordem de aparição).
    """
    cidades = []
    seen = set()
    if "Terminal" not in nc_fisc_df.columns:
        return []

    for raw in nc_fisc_df["Terminal"].fillna("").astype(str).tolist():
        s = raw.strip()
        if not s:
            continue

        # Remove prefixo 'Terminal de ' (case-insensitive)
        s = re.sub(r"(?i)^\s*terminal\s+de\s+", "", s)

        # Remove conteúdo entre parênteses no final e espaços sobrando
        s = re.sub(r"\s*\(.*?\)\s*$", "", s).strip()

        # Se após limpeza ainda houver algo, use capitalização conservadora
        if s and s not in seen:
            cidades.append(s)
            seen.add(s)

    return cidades


def gerar_secao_anexo_fotos(
    doc: Document,
    row: Row,
    nao_conformidades_df: DataFrame,
    fotos_dir: str,
    processo_info: dict,
) -> None:
    """
    Gera a seção 'ANEXO - MEMORIAL FOTOGRÁFICO' no documento `doc`.
    - row: linha da fiscalização (contém 'ID da Fiscalização' e info contextual).
    - nao_conformidades_df: DataFrame com as NCs.
    - fotos_dir: diretório onde as fotos do monitoramento estão armazenadas.
    - processo_info: dicionário extraído da aba 'Processos' para a fiscalização atual.
    """

    doc.add_page_break()

    caminho_fotos_monitoramento = fotos_dir

    # --- Preenche período de vistoria a partir de processo_info ---
    periodo_raw = processo_info.get("Periodo de Vistoria da ARPE", None)
    periodo_formatado = _formatar_periodos(periodo_raw)

    # --- Preenche o número do Processo CTR ---
    processo_ctr_num = processo_info.get("Processo CTR Nº", "XX/XXXX")

    # --- Extrai nomes das cidades (Terminal) a partir do DataFrame de não conformidades ---
    id_fisc = row["ID da Fiscalização"]
    nc_fisc: DataFrame = nao_conformidades_df[
        nao_conformidades_df["ID da Fiscalização"] == id_fisc
    ].copy()

    cidades = _extrair_cidades_desde_nc(nc_fisc)
    cidades_str = ", ".join(cidades) if cidades else "(NOMES DAS CIDADES)"

    # --- Monta título dinâmico ---
    titulo_base = "ANEXO - MEMORIAL FOTOGRÁFICO - VISTORIAS REALIZADAS"
    if periodo_formatado:
        titulo_completo = f"{titulo_base} EM {periodo_formatado}"
    else:
        titulo_completo = f"{titulo_base} EM XX a XX/XX/XXXX"

    adicionar_titulo_secao(
        doc,
        titulo_completo,
        aplicar_sombra=True,
    )

    if nc_fisc.empty:
        doc.add_page_break()
        doc.add_paragraph("Nenhuma não conformidade monitorada disponível.")
        return

    paragrafo_anexo = doc.add_paragraph()
    paragrafo_anexo.paragraph_format.space_after = Pt(12)

    # Monta o texto inicial substituindo o CTR e as cidades
    texto_inicial = (
        "Apresenta-se, a seguir, evidências fotográficas das Não Conformidades pendentes do "
        f"Relatório de Fiscalização Técnico-Operacional Arpe/CTR nº {processo_ctr_num} para os Terminais Rodoviários "
        f"de Passageiros concedidos à SOCICAM nas cidades de {cidades_str}."
    )

    run_paragrafo = paragrafo_anexo.add_run(texto_inicial)
    aplicar_estilo_corpo(run_paragrafo, negrito=False)
    paragrafo_anexo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY_LOW

    # Lista de arquivos na pasta (tratamento de erro se pasta ausente)
    try:
        arquivos_na_pasta = os.listdir(caminho_fotos_monitoramento)
    except FileNotFoundError:
        doc.add_paragraph(f"Não foi possível gerar o Memorial Fotográfico. Pasta ausente: {caminho_fotos_monitoramento}")
        return

    terminal_anterior = None

    for terminal, grupo_terminal in nc_fisc.groupby("Terminal"):

        nc_list_individual: List[dict] = []

        for _, linha in grupo_terminal.iterrows():
            ids_individuais = _safe_split_anexo(linha.get(CHAVE_ID_ARPE, ""))
            descricoes_individuais = _safe_split_anexo(linha.get(CHAVE_NC_DESC, ""))
            legenda_bruta_do_excel = str(linha.get("Legenda da Foto", "")).strip()

            num_descricoes = len(descricoes_individuais)

            ids_individuais = _adaptar_lista_anexo(ids_individuais, num_descricoes, "ID_FALTANDO")
            legendas_originais = _safe_split_anexo(legenda_bruta_do_excel)
            legendas_lista = _adaptar_lista_anexo(legendas_originais, num_descricoes, "")

            for idx in range(num_descricoes):
                nc_list_individual.append(
                    {
                        "terminal": terminal,
                        "id": ids_individuais[idx].strip(),
                        "desc": descricoes_individuais[idx].strip(),
                        "legenda_nc": legendas_lista[idx],
                    }
                )

        nc_list_individual.sort(key=lambda x: x["id"])

        for nc_data in nc_list_individual:
            nc_id_bruto = nc_data["id"]
            constatacao = nc_data["desc"]
            legenda_nc = nc_data["legenda_nc"]
            terminal_nome = nc_data["terminal"]

            if terminal_nome != terminal_anterior:
                # ao adicionar o heading para o terminal, não colocamos "Terminal de" nem a sigla
                terminal_limpo = re.sub(r"(?i)^\s*terminal\s+de\s+", "", str(terminal_nome).strip())
                terminal_limpo = re.sub(r"\s*\(.*?\)\s*$", "", terminal_limpo).strip()
                adicionar_titulo_secao(doc, f"TERMINAL DE {terminal_limpo.upper()}", nivel_heading=2)
                terminal_anterior = terminal_nome

            prefixo_busca = _limpar_id_para_busca(nc_id_bruto)

            if not prefixo_busca or prefixo_busca == "ID_FALTANDO":
                continue

            fotos_encontradas = [
                f
                for f in arquivos_na_pasta
                if f.upper().startswith(f"{prefixo_busca}_") and f.lower().endswith((".jpg", ".jpeg", ".png"))
            ]

            fotos_encontradas.sort()

            # Prepara legendas por foto
            legendas_por_foto_brutas = _safe_split_anexo(legenda_nc)
            num_fotos_encontradas = len(fotos_encontradas)
            legendas_alinhadas = _adaptar_lista_anexo(legendas_por_foto_brutas, num_fotos_encontradas, "")

            _adicionar_contexto_nc(doc, nc_id_bruto.strip(), constatacao.strip())

            if fotos_encontradas:

                fotos_em_pares = list(zip_longest(*[iter(fotos_encontradas)] * 2, fillvalue=None))

                for i, (foto1_nome, foto2_nome) in enumerate(fotos_em_pares):
                    idx_legenda1 = i * 2
                    idx_legenda2 = i * 2 + 1

                    legenda1 = legendas_alinhadas[idx_legenda1]
                    legenda2 = (
                        legendas_alinhadas[idx_legenda2] if (foto2_nome and idx_legenda2 < num_fotos_encontradas) else ""
                    )

                    adicionar_duas_imagens_lado_a_lado(
                        doc,
                        caminho_fotos_monitoramento,
                        foto1_nome,
                        legenda1,
                        foto2_nome,
                        legenda2,
                        contexto_nc_tupla=None,
                    )

    doc.add_page_break()

