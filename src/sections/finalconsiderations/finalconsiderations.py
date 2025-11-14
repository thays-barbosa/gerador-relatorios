# CÓDIGO COMPLETO — MESMA ESTRUTURA ORIGINAL + SUAS ALTERAÇÕES
# NADA FOI REORGANIZADO OU SIMPLIFICADO — APENAS ATUALIZADO CONFORME PEDIDO

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
    """
    if not nome or "xxxxxx" in nome.lower():
        return

    adicionar_texto_centralizado(doc, nome, negrito=negrito_nome)
    adicionar_texto_centralizado(doc, cargo, negrito=False)

    if matricula:
        # Garante prefixo “Matrícula” se ainda não estiver no texto
        texto_matricula = matricula
        if "matrícula" not in matricula.lower() and "nº" not in matricula.lower():
            texto_matricula = f"Matrícula: n°{matricula}"

        adicionar_texto_centralizado(doc, texto_matricula, negrito=False)

    doc.add_paragraph()



def gerar_secao_consideracoes_finais(doc: Document, row, nao_conformidades_df, processo_info) -> None:
    """
    Gera a seção '5. CONCLUSÃO' do relatório.
    """

    adicionar_titulo_secao(doc, "5. CONCLUSÃO")

    # --------- (1) ID DO MONITORAMENTO — ABA NÃO-CONFORMIDADES ----------
    id_fisc = row.get("ID da Fiscalização")
    dados_nc = nao_conformidades_df[nao_conformidades_df["ID da Fiscalização"] == id_fisc]

    if not dados_nc.empty:
        num_monitoramento = str(dados_nc.iloc[0].get("ID da Fiscalização", "X"))
    else:
        num_monitoramento = "X"

    # --------- (2) CTR ORIGINAL — ABA PROCESSO ----------
    ctr_original = str(processo_info.get("Processo CTR Nº", "xx/xxxx"))

    # --------- (3) PERÍODO DE VISTORIA — ABA PROCESSO ----------
    periodo_vistoria_raw = str(
        processo_info.get("Periodo de Vistoria da ARPE", "xx a xx de MÊS de ANO")
    )

    # Divide por ";" e junta com " e "
    periodos_list = [p.strip() for p in periodo_vistoria_raw.split(";") if p.strip()]
    periodo_vistoria = " e ".join(periodos_list) if periodos_list else "xx a xx de MÊS de ANO"

    # --------- TEXTO FINAL ----------
    texto_conclusao = (
        f"Diante das constatações apontadas neste {num_monitoramento}º Relatório de "
        f"Monitoramento do Relatório de Fiscalização Técnico-Operacional CTR {ctr_original}, "
        f"referente às vistorias técnicas realizadas no período de {periodo_vistoria}, "
        "solicitamos seu envio para a SOCICAM para que sejam informados das Não Conformidades."
    )

    adicionar_paragrafo_justificado(doc, texto_conclusao)

    # --------- DATA ----------
    data_atual = datetime.now().strftime("%d/%m/%Y")
    adicionar_texto_centralizado(doc, f"Recife, {data_atual}.", negrito=False, tamanho_fonte=11)

    doc.add_paragraph()

    # --------- (4) ASSINANTES — MATRÍCULAS DINÂMICAS ----------
    assinantes = str(row.get("Assinatura", "")).split(";")
    matriculas = str(row.get("Matriculas das Pessoas Responsáveis", "")).split(";")

    for i, nome in enumerate(assinantes):
        nome = nome.strip()
        matricula = matriculas[i].strip() if i < len(matriculas) else ""

        _adicionar_assinatura_bloco(
            doc,
            nome,
            "Analista de Regulação",
            matricula,
            negrito_nome=True
        )

    adicionar_texto_esquerda(doc, "Ciente.")

    # --------- (5) COORDENADOR ----------
    coordenador = str(row.get("Coordenador", "")).strip()
    matricula_coordenador = str(row.get("Matrícula Coordenador", "Matrícula: nº209640/01"))

    _adicionar_assinatura_bloco(
        doc,
        coordenador,
        "Coordenadora de Transportes e Rodovias",
        matricula_coordenador,
        negrito_nome=True
    )


