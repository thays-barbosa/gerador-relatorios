from datetime import datetime
from docx.document import Document
from utils import (
    adicionar_titulo_secao,
    adicionar_paragrafo_justificado,
    adicionar_texto_centralizado,
    adicionar_texto_esquerda,
)


def _adicionar_assinatura_bloco(
    doc: Document,
    nome: str,
    cargo: str,
    matricula: str,
    negrito_nome: bool = True
) -> None:
    """
    Adiciona um bloco de assinatura formatado no centro do documento.

    Args:
        doc (Document): Documento Word em construção.
        nome (str): Nome da pessoa.
        cargo (str): Cargo da pessoa.
        matricula (str): Matrícula funcional.
        negrito_nome (bool, opcional): Define se o nome será negritado. Padrão é True.
    """
    if not nome or "xxxxxx" in nome.lower():
        return

    adicionar_texto_centralizado(doc, nome, negrito=negrito_nome)
    adicionar_texto_centralizado(doc, cargo, negrito=False)

    if matricula:
        # Garante prefixo “Matrícula” se ainda não estiver no texto
        texto_matricula = matricula
        if "matrícula" not in matricula.lower() and "nº" not in matricula.lower():
            texto_matricula = f"Matrícula: {matricula}"

        adicionar_texto_centralizado(doc, texto_matricula, negrito=False)

    doc.add_paragraph()


def gerar_secao_consideracoes_finais(doc: Document, row) -> None:
    """
    Gera a seção '5. CONCLUSÃO' do relatório.

    Esta seção apresenta o texto conclusivo, a data e os blocos de assinatura
    dos analistas e da coordenadora responsáveis pelo relatório.

    Args:
        doc (Document): Documento Word em construção.
        row (Series): Linha da planilha contendo os dados da fiscalização.
    """
    adicionar_titulo_secao(doc, "5. CONCLUSÃO")

    num_monitoramento = str(row.get("Num Monitoramento", "Xº"))
    ctr_original = str(row.get("CTR Original", "xx/xxxx"))
    periodo_vistoria = str(row.get("Periodo Vistoria Texto", "xx a xx de MÊS de ANO"))

    texto_conclusao = (
        f"Diante das constatações apontadas neste {num_monitoramento} Relatório de "
        f"Monitoramento do Relatório de Fiscalização Técnico-Operacional CTR {ctr_original}, "
        f"referente às vistorias técnicas realizadas no período de {periodo_vistoria}, "
        "solicitamos seu envio para a SOCICAM para que sejam informados das Não Conformidades."
    )

    adicionar_paragrafo_justificado(doc, texto_conclusao)

    # Adiciona data atual formatada
    data_atual = datetime.now().strftime("%d/%m/%Y")
    adicionar_texto_centralizado(doc, f"Recife, {data_atual}.", negrito=False, tamanho_fonte=11)

    doc.add_paragraph()

    assinantes = str(row.get("Assinatura", "")).split(";")
    matricula_analista_fixa = "Matrícula: nºxxxxxx/xx"

    for assinante in assinantes:
        nome = assinante.strip()
        _adicionar_assinatura_bloco(
            doc,
            nome,
            "Analista de Regulação",
            matricula_analista_fixa,
            negrito_nome=True
        )

    adicionar_texto_esquerda(doc, "Ciente.", negrito=False)

    coordenador = str(row.get("Coordenador", "")).strip()
    matricula_coordenador_fixa = "Matrícula: nº209640/01"

    _adicionar_assinatura_bloco(
        doc,
        coordenador,
        "Coordenadora de Transportes e Rodovias",
        matricula_coordenador_fixa,
        negrito_nome=True
    )
