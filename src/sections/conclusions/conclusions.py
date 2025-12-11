from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
import pandas as pd
from utils import (
    adicionar_paragrafo_justificado,
    adicionar_titulo_secao,
    adicionar_texto_centralizado,
)
from datetime import datetime


def gerar_secao_conclusoes(
    doc, row
):
    """
    Gera a seção '3. NÃO CONFORMIDADES CONSTATADAS',
    incluindo Observações Importantes e Recomendações.
    """

    adicionar_titulo_secao(doc, "7. CONCLUSÕES")

    texto1 = "Tendo em vista as ações de fiscalização realizadas pela Arpe foram constatadas mais nove Não Conformidades distribuídas nos Terminais Rodoviários das cidades de Garanhuns (2), Petrolina (2), Caruaru (2) e do Recife-TIP (3), que devem ser solucionadas pela SOCICAM de acordo com as Determinações desta Agência de Regulação (v. Quadro 1). Cabe reforçar a recomendação de levantamento diagnóstico das cobertas dos Terminais Rodoviários,com o envio à Arpe dos respectivos laudos técnicos. "

    texto2 = "Por fim, solicita-se o encaminhamento deste Processo de Fiscalização para conhecimento e acompanhamento da EPTI, na qualidade de Poder Concedente do Contrato de Concessão e gestora do Sistema de Transporte Coletivo Intermunicipal de Passageiros (STCIP-PE). "

    adicionar_paragrafo_justificado(doc, texto1)
    adicionar_paragrafo_justificado(doc, texto2)

    adicionar_titulo_secao(doc, "APÊNDICE 1 - REGISTROS FOTOGRÁFICOS DAS NÃO CONFORMIDADES")



        # Pega a data atual no formato dd/mm/aaaa
    data_atual = datetime.now().strftime("%d/%m/%Y")
    adicionar_texto_centralizado(doc, f"\n\nRecife, {data_atual}.")
    adicionar_texto_centralizado(doc, "\n\n")

    # Assinaturas dos responsáveis (pode ser uma string separada por ";" ou ",")
    assinantes = str(row.get("Assinatura", "")).split(";")
    for assinante in assinantes:
        nome = assinante.strip()
        if nome:
                adicionar_texto_centralizado(doc, "_______________________")
                adicionar_texto_centralizado(doc, nome)
                adicionar_texto_centralizado(doc, "Analista de Regulação")
                adicionar_texto_centralizado(doc, "")  # Espaço em branco entre assinaturas

        adicionar_texto_centralizado(doc, "\nCiente e de acordo:\n")
        coordenador = str(row.get("Coordenador", "")).strip()
        if coordenador:
            adicionar_texto_centralizado(doc, "_______________________")
            adicionar_texto_centralizado(doc, coordenador)        