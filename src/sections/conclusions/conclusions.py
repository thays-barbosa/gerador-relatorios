from docx.shared import  Pt 
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
import pandas as pd
from datetime import datetime
from utils import (
    adicionar_paragrafo_justificado,
    adicionar_titulo_secao,
   
    adicionar_apendice_fotos,
   
    extrair_cidade, 
    adicionar_assinaturas_formatadas,
)


def gerar_secao_conclusoes(
    doc, 
    row,
    caminho_planilha_legendas,
    caminho_base_fotos 
):
    """
    Gera a seção '7. CONCLUSÕES', o 'APÊNDICE 1' e a seção de ASSINATURAS (com dados FIXOS).
    """

    terminais_str = str(row.get("Terminais", "")).strip()

    cidades_fisc = [
        extrair_cidade(t.strip()) 
        for t in terminais_str.split(';') 
        if t.strip()
    ]

    cidades_unicas = sorted(list(set(c.capitalize().replace(' (Tip)', '-TIP') for c in cidades_fisc)))

    if not cidades_unicas:
        cidades_str = "nos Terminais Rodoviários de jurisdição da Agência"

    elif len(cidades_unicas) == 1:
        cidades_str = f"no Terminal Rodoviário da cidade de {cidades_unicas[0]}"
    else:
      
        parte_inicial = ', '.join(cidades_unicas[:-1])
        ultimo_terminal = cidades_unicas[-1]

        cidades_str = f"nos Terminais Rodoviários das cidades de {parte_inicial} e de {ultimo_terminal}"

    adicionar_titulo_secao(doc, "7. CONCLUSÕES")

    doc.add_paragraph()

    texto1 = (
        f"Tendo em vista as ações de fiscalização realizadas pela Arpe foram constatadas mais nove Não Conformidades "
        f"distribuídas {cidades_str}, que devem ser solucionadas pela SOCICAM de acordo com as "
        f"Determinações desta Agência de Regulação (v. Quadro 1). Cabe reforçar a recomendação de levantamento "
        f"diagnóstico das cobertas dos Terminais Rodoviários, com o envio à Arpe dos respectivos laudos técnicos."
    )

    texto2 = "Por fim, solicita-se o encaminhamento deste Processo de Fiscalização para conhecimento e acompanhamento da EPTI, na qualidade de Poder Concedente do Contrato de Concessão e gestora do Sistema de Transporte Coletivo Intermunicipal de Passageiros (STCIP-PE). "

    adicionar_paragrafo_justificado(doc, texto1)
    adicionar_paragrafo_justificado(doc, texto2)

    adicionar_titulo_secao(doc, "APÊNDICE 1 - REGISTROS FOTOGRÁFICOS DAS NÃO CONFORMIDADES")

    doc.add_paragraph() 
    
    id_fiscalizacao = str(row.get("ID da Fiscalização", 1)) 

    adicionar_apendice_fotos(
        doc, 
        caminho_base_fotos, 
        id_fiscalizacao,
        caminho_planilha_legendas
    )

    analistas_fixos = [
        ("Alcides Vieira de Azevedo Bezerra", "Analista de Regulação", "Matrícula nº 40672015/01"),
        ("Enildo Manoel da Silva Júnior", "Analista de Regulação", "Matrícula nº 1796500/02")
        ]
    coordenador_nome_fixo = "Maria Ângela Albuquerque de Freitas"

    adicionar_assinaturas_formatadas(
        doc,
        analistas_fixos,
        coordenador_nome_fixo,
        cidade_relatorio="Recife"
    )