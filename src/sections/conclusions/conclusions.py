# sections/conclusions/conclusions.py (VERSÃO FINAL)

from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
import pandas as pd
from datetime import datetime
from utils import (
    adicionar_paragrafo_justificado,
    adicionar_titulo_secao,
    adicionar_texto_centralizado,
    adicionar_apendice_fotos,
    # <<< IMPORTAÇÃO DA NOVA FUNÇÃO AUXILIAR >>>
    parear_e_formatar_assinaturas, 
)


def gerar_secao_conclusoes(
    doc, 
    row,
    caminho_planilha_legendas,
    caminho_base_fotos 
):
    """
    Gera a seção '7. CONCLUSÕES', o 'APÊNDICE 1' e a seção de ASSINATURAS.
    """

    adicionar_titulo_secao(doc, "7. CONCLUSÕES")

    texto1 = "Tendo em vista as ações de fiscalização realizadas pela Arpe foram constatadas mais nove Não Conformidades distribuídas nos Terminais Rodoviários das cidades de Garanhuns (2), Petrolina (2), Caruaru (2) e do Recife-TIP (3), que devem ser solucionadas pela SOCICAM de acordo com as Determinações desta Agência de Regulação (v. Quadro 1). Cabe reforçar a recomendação de levantamento diagnóstico das cobertas dos Terminais Rodoviários,com o envio à Arpe dos respectivos laudos técnicos. "
    texto2 = "Por fim, solicita-se o encaminhamento deste Processo de Fiscalização para conhecimento e acompanhamento da EPTI, na qualidade de Poder Concedente do Contrato de Concessão e gestora do Sistema de Transporte Coletivo Intermunicipal de Passageiros (STCIP-PE). "

    adicionar_paragrafo_justificado(doc, texto1)
    adicionar_paragrafo_justificado(doc, texto2)

    adicionar_titulo_secao(doc, "APÊNDICE 1 - REGISTROS FOTOGRÁFICOS DAS NÃO CONFORMIDADES")

    id_fiscalizacao = str(row.get("ID da Fiscalização", 1)) 
    
    adicionar_apendice_fotos(
        doc, 
        caminho_base_fotos, 
        id_fiscalizacao,
        caminho_planilha_legendas
    )

    # --- SEÇÃO DE ASSINATURAS FORMATADA (Modelo da Imagem) ---
    
    adicionar_texto_centralizado(doc, f"\n\nRecife, data da assinatura eletrônica.")
    adicionar_texto_centralizado(doc, "\n\n") 

    # 1. Puxa dados dinâmicos
    nomes_responsaveis = row.get("Pessoal Responsável", "")
    matriculas_responsaveis = row.get("Matrícula do Pessoal Responsável", "")
    
    # 2. Pareia e formata os analistas usando a função do utils
    analistas = parear_e_formatar_assinaturas(
        nomes_responsaveis, 
        matriculas_responsaveis, 
        cargo_fixo="Analista de Regulação"
    )

    # 3. Adiciona as assinaturas dos Analistas
    for nome, cargo, matricula in analistas:
        if not nome: continue 
        
        adicionar_texto_centralizado(doc, "\n") # Espaçamento
        adicionar_texto_centralizado(doc, nome)
        adicionar_texto_centralizado(doc, cargo)
        
        if matricula:
            adicionar_texto_centralizado(doc, matricula)
            
        adicionar_texto_centralizado(doc, "\n") # Espaço após bloco (para replicar o layout da foto)

    # 4. Assinatura do Coordenador
    
    # Linha "Ciente e de acordo." (com espaçamento)
    adicionar_texto_centralizado(doc, "\n\nCiente e de acordo.\n\n")
    
    coordenador = str(row.get("Coordenador da Fiscalização", "")).strip()
    cargo_coordenador = str(row.get("Cargo do Coordenador", "Coordenador de Transportes e Rodovias")).strip()
    matricula_coordenador = str(row.get("Matrícula do Coordenador", "")).strip()
    
    if coordenador:
        adicionar_texto_centralizado(doc, "\n") # Espaçamento
        adicionar_texto_centralizado(doc, coordenador)
        adicionar_texto_centralizado(doc, cargo_coordenador)
        
        if matricula_coordenador:
             adicionar_texto_centralizado(doc, f"Matrícula nº {matricula_coordenador}")