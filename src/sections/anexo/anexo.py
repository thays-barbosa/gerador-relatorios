"""
anexo.py

Gera a seção 'ANEXO - MEMORIAL FOTOGRÁFICO' do relatório, buscando fotos no disco
com base no 'ID da não conformidade' presente na planilha.
"""

from typing import List
from itertools import zip_longest
import os

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


def gerar_secao_anexo_fotos(doc: Document, row: Row, nao_conformidades_df: DataFrame, fotos_dir: str) -> None:
    """
    Gera a seção 'ANEXO - MEMORIAL FOTOGRÁFICO' no documento `doc`.
    - row: linha da fiscalização (contém 'ID da Fiscalização' e info contextual).
    - nao_conformidades_df: DataFrame com as NCs.
    - fotos_dir: diretório onde as fotos do monitoramento estão armazenadas.
    """

    doc.add_page_break()

    caminho_fotos_monitoramento = fotos_dir

    adicionar_titulo_secao(
        doc,
        "ANEXO - MEMORIAL FOTOGRÁFICO - VISTORIAS REALIZADAS EM XX a XX/XX/XXXX",
        aplicar_sombra=True,
    )

    id_fisc = row["ID da Fiscalização"]

    nc_fisc: DataFrame = nao_conformidades_df[
        nao_conformidades_df["ID da Fiscalização"] == id_fisc
    ].copy()

    if nc_fisc.empty:
        doc.add_page_break()
        doc.add_paragraph("Nenhuma não conformidade monitorada disponível.")
        return

    paragrafo_anexo = doc.add_paragraph()
    paragrafo_anexo.paragraph_format.space_after = Pt(12)
    run_paragrafo = paragrafo_anexo.add_run(
        "Apresenta-se, a seguir, evidências fotográficas das Não Conformidades pendentes do Relatório de Fiscalização Técnico-Operacional "
        "Arpe/CTR nº XX/XXXX para os Terminais Rodoviários de Passageiros concedidos à SOCICAM nas cidades do (NOMES DAS CIDADES)."
    )
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
                adicionar_titulo_secao(doc, f"TERMINAL DE {terminal_nome.upper()}", nivel_heading=2)
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
                    legenda2 = legenda_alinhada = (
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
