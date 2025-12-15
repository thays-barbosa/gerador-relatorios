# sections/conclusions/conclusions.py (VERSÃO FINAL COM DADOS FIXOS)
from docx.shared import  Pt 
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
import pandas as pd
from datetime import datetime
from utils import (
    adicionar_paragrafo_justificado,
    adicionar_titulo_secao,
    # remover_espacamento_paragrafo, # <<< Garanta que essa função está no seu utils.py
    adicionar_apendice_fotos,
    # parear_e_formatar_assinaturas, # <<< Não precisamos mais desta para assinaturas fixas
    extrair_cidade, 
    adicionar_assinaturas_formatadas, # <<< FUNÇÃO DE LAYOUT NECESSÁRIA >>>
)
# Outras funções auxiliares (remover_espacamento_paragrafo, etc.)

def gerar_secao_conclusoes(
    doc, 
    row,
    caminho_planilha_legendas,
    caminho_base_fotos 
):
    """
    Gera a seção '7. CONCLUSÕES', o 'APÊNDICE 1' e a seção de ASSINATURAS (com dados FIXOS).
    """

    # --- LÓGICA DE EXTRAÇÃO DE CIDADES FISCALIZADAS (Mantida) ---

    terminais_str = str(row.get("Terminais", "")).strip()

    # 1. Divide a string de Terminais e extrai apenas a cidade
    cidades_fisc = [
        extrair_cidade(t.strip()) 
        for t in terminais_str.split(';') 
        if t.strip()
    ]

    # Remove duplicatas e capitaliza (a função extrair_cidade já capitaliza, mas garante)
    cidades_unicas = sorted(list(set(c.capitalize().replace(' (Tip)', '-TIP') for c in cidades_fisc)))

     # 2. Constrói a string de cidades formatada
    if not cidades_unicas:
        cidades_str = "nos Terminais Rodoviários de jurisdição da Agência"

    elif len(cidades_unicas) == 1:
        cidades_str = f"no Terminal Rodoviário da cidade de {cidades_unicas[0]}"
    else:
        # Junta todos os terminais, exceto o último, separados por vírgula
        parte_inicial = ', '.join(cidades_unicas[:-1])
        ultimo_terminal = cidades_unicas[-1]

        # Constrói a frase final: "...das cidades de X, Y e Z."
        cidades_str = f"nos Terminais Rodoviários das cidades de {parte_inicial} e de {ultimo_terminal}"

    # --- INÍCIO DA SEÇÃO 7 ---

    adicionar_titulo_secao(doc, "7. CONCLUSÕES")

    doc.add_paragraph()

    # Texto 1 Dinâmico 
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

    # --- SEÇÃO DE ASSINATURAS FORMATADA (DADOS FIXOS) ---

    # 1. DADOS FIXOS CONFORME SOLICITADO
    analistas_fixos = [
        ("Alcides Vieira de Azevedo Bezerra", "Analista de Regulação", "Matrícula nº 40672015/01"),
        ("Enildo Manoel da Silva Júnior", "Analista de Regulação", "Matrícula nº 1796500/02")
        ]
    coordenador_nome_fixo = "Maria Ângela Albuquerque de Freitas"

    # 2. CHAMA A FUNÇÃO QUE GERA O LAYOUT EM COLUNAS
    adicionar_assinaturas_formatadas(
        doc,
        analistas_fixos,
        coordenador_nome_fixo,
        cidade_relatorio="Recife"
    )