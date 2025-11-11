from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import (
    adicionar_paragrafo_justificado,
    adicionar_titulo_secao,
    adicionar_paragrafo_info_compacta,
    aplicar_estilo_corpo,
    aplicar_estilo_titulo
)
from typing import Any
import pandas as pd

try:
    from tqdm import tqdm
    tqdm_write = tqdm.write
except ImportError:
    tqdm_write = print


def _inserir_texto_nc(doc, nc_titulo_identificador: str, descricao: str):
    """
    Insere o cabeçalho e a descrição de uma Não Conformidade no documento.
    """
    paragrafo_nc = doc.add_paragraph()

    run_titulo = paragrafo_nc.add_run(f"Não Conformidade {nc_titulo_identificador}")
    aplicar_estilo_corpo(run_titulo, negrito=True)
    run_titulo.underline = True
    run_titulo.font.color.rgb = RGBColor(0, 0, 0)

    run_traco = paragrafo_nc.add_run(" - ")
    aplicar_estilo_corpo(run_traco)
    run_traco.underline = False
    run_traco.font.color.rgb = RGBColor(0, 0, 0)

    run_desc = paragrafo_nc.add_run(descricao)
    aplicar_estilo_corpo(run_desc)
    run_desc.underline = False
    run_desc.font.color.rgb = RGBColor(0, 0, 0)

    paragrafo_nc.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY_LOW


def _inserir_linhas_info(doc, info_socicam: str, constatacao: str, analise_arpe: str):
    """
    Insere os blocos de 'Informação', 'Constatação' e 'Análise' de forma compacta.
    """
    adicionar_paragrafo_info_compacta(doc, "Informação da SOCICAM: ", info_socicam)
    adicionar_paragrafo_info_compacta(doc, "Constatação: ", constatacao)
    adicionar_paragrafo_info_compacta(doc, "Análise da ARPE: ", analise_arpe)


def _safe_split(nc_data: Any, col_name: str, default: str = "Texto não disponível") -> list:
    """
    Divide os textos de uma coluna (separados por ';') em uma lista limpa.
    Retorna uma lista contendo um valor padrão se o conteúdo for inválido.
    """
    if not isinstance(nc_data, (pd.Series, dict)):
        tqdm_write(f"⚠️ Erro grave: tipo inesperado em safe_split: {type(nc_data)}")
        return [default]

    content = str(nc_data.get(col_name, default)).strip()
    if not content or content.lower() == "nan":
        if col_name in ["ID da não conformidade", "Não Conformidade"]:
            return []
        return [default]

    return [item.strip() for item in content.split(";")]


def gerar_secao_nao_conformidades_constatadas(doc, row: dict, nao_conformidades_df: pd.DataFrame, FOTOS_DIR: str):
    """
    Gera a seção '3. RESULTADO DAS VISTORIAS DAS NÃO CONFORMIDADES PENDENTES'
    com base nas informações da planilha.
    """
    id_fiscalizacao = row["ID da Fiscalização"]

    nc_fiscalizacao = nao_conformidades_df[
        nao_conformidades_df["ID da Fiscalização"] == id_fiscalizacao
    ].copy()

    adicionar_titulo_secao(doc, "3. RESULTADO DAS VISTORIAS DAS NÃO CONFORMIDADES PENDENTES")

    adicionar_paragrafo_justificado(
        doc,
        (
            "Estão registrados para cada Terminal Rodoviário os resultados da verificação pela Arpe das ações "
            "desenvolvidas pela SOCICAM, constantes da Carta SAP/PER/ARPE N° XXXX/XXXX e Carta SAP/PER/ARPE N° XXX/XXXX, "
            "para solucionar as Não Conformidades ainda pendentes apresentadas no Relatório de Fiscalização Técnico-"
            "Operacional ARPE/CTR nº XX/XXXX. Monitoramento do Processo Arpe/CTR XX/XXXX (Item X)."
        ),
    )

    if "Terminal" not in nc_fiscalizacao.columns:
        adicionar_paragrafo_justificado(
            doc,
            "⚠️ Coluna 'Terminal' não encontrada na planilha de não conformidades."
        )
        return

    id_arpe_col = "ID da não conformidade"
    num_terminal = 1

    for terminal, grupo_terminal in nc_fiscalizacao.groupby("Terminal"):
        par_terminal = doc.add_paragraph()
        run_terminal = par_terminal.add_run(f"3.{num_terminal} - {terminal.upper()}")
        aplicar_estilo_titulo(run_terminal)

        par_terminal.paragraph_format.space_before = Pt(12)
        par_terminal.paragraph_format.space_after = Pt(6)

        for _, nc_data in grupo_terminal.iterrows():

            nc_identificadores = _safe_split(nc_data, id_arpe_col, default="")
            descricoes = _safe_split(nc_data, "Não Conformidade", default="")
            info_socicam = _safe_split(nc_data, "Informação SOCICAM")
            constatacao = _safe_split(nc_data, "Constatação")
            analise_arpe = _safe_split(nc_data, "Análise da Arpe")

            num_descricoes = len(descricoes)
            num_ids = len(nc_identificadores)
            log_id = f"Terminal: {terminal}"

            if num_ids != num_descricoes:
                if 0 < num_ids < num_descricoes:
                    nc_identificadores.extend([nc_identificadores[-1]] * (num_descricoes - num_ids))
                    tqdm_write(f"⚠️ Quantidade de IDs menores que descrições ({log_id}). Confira sua planilha.")
                elif num_ids > num_descricoes:
                    nc_identificadores = nc_identificadores[:num_descricoes]
                    tqdm_write(f"⚠️ Quantidade de IDs maiores que descrições ({log_id}). Confira sua planilha.")
                elif num_ids == 0 and num_descricoes > 0:
                    nc_identificadores = ["ID_FALTANDO"] * num_descricoes
                    tqdm_write(f"⚠️ Quantidade de IDs ausentes ({log_id}). Usado 'ID_FALTANDO' como substituto.")

            # Montagem da seção por item
            for i, descricao in enumerate(descricoes):
                alinhamento_fail = (
                    "ALINHAMENTO FALHOU! Verifique o uso de ';' em todas as 4 colunas."
                )

                info = info_socicam[i] if i < len(info_socicam) else "Texto não disponível"
                const = constatacao[i] if i < len(constatacao) else "Texto não disponível"
                analise = analise_arpe[i] if i < len(analise_arpe) else "Texto não disponível"

                max_len = max(
                    len(descricoes),
                    len(info_socicam),
                    len(constatacao),
                    len(analise_arpe)
                )

                if len(descricoes) != max_len:
                    info = const = analise = alinhamento_fail

                id_final = nc_identificadores[i]
                _inserir_texto_nc(doc, id_final, descricao)
                _inserir_linhas_info(doc, info, const, analise)

                doc.add_paragraph()

        num_terminal += 1
