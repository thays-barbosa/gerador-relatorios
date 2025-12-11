from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
import pandas as pd
from utils import (
    adicionar_paragrafo_justificado,
    adicionar_titulo_secao,
    adicionar_texto_centralizado,
    adicionar_apendice_fotos,
)
from datetime import datetime


def gerar_secao_conclusoes(
    doc, 
    row,
    caminho_planilha_legendas,
    # Argumento UNIFICADO e VALIDADO, vindo do report.py:
    caminho_base_fotos 
):
    """
    Gera a seção '7. CONCLUSÕES' e o 'APÊNDICE 1'.
    """

    adicionar_titulo_secao(doc, "7. CONCLUSÕES")

    texto1 = "Tendo em vista as ações de fiscalização realizadas pela Arpe foram constatadas mais nove Não Conformidades distribuídas nos Terminais Rodoviários das cidades de Garanhuns (2), Petrolina (2), Caruaru (2) e do Recife-TIP (3), que devem ser solucionadas pela SOCICAM de acordo com as Determinações desta Agência de Regulação (v. Quadro 1). Cabe reforçar a recomendação de levantamento diagnóstico das cobertas dos Terminais Rodoviários,com o envio à Arpe dos respectivos laudos técnicos. "

    texto2 = "Por fim, solicita-se o encaminhamento deste Processo de Fiscalização para conhecimento e acompanhamento da EPTI, na qualidade de Poder Concedente do Contrato de Concessão e gestora do Sistema de Transporte Coletivo Intermunicipal de Passageiros (STCIP-PE). "

    adicionar_paragrafo_justificado(doc, texto1)
    adicionar_paragrafo_justificado(doc, texto2)

    adicionar_titulo_secao(doc, "APÊNDICE 1 - REGISTROS FOTOGRÁFICOS DAS NÃO CONFORMIDADES")

    # LINHAS REMOVIDAS: A montagem do caminho não é mais necessária aqui,
    # pois o caminho_base_fotos já vem validado e completo do report.py.
    # caminho_base_fotos = os.path.join(pasta_principal_fotos, subpasta_fotos)

    # 2. Pega o ID da fiscalização (Presumindo que seja uma coluna em 'row')
    id_fiscalizacao = row.get("ID da Fiscalização", 1) 
    
    # 3. Chama a nova função para gerar o apêndice
    adicionar_apendice_fotos(
        doc, 
        caminho_base_fotos, # Argumento ÚNICO e completo
        id_fiscalizacao,
        caminho_planilha_legendas
    )

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