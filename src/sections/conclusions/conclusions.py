# sections/conclusions/conclusions.py (VERSÃO FINAL SIMPLIFICADA SEM CONTAGEM)

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
    parear_e_formatar_assinaturas,
    extrair_cidade, # <<< NECESSÁRIO >>>
    
)


# ATENÇÃO: A assinatura da função VOLTOU à versão original
# sem 'nao_conformidades_df', pois a contagem foi removida.
def gerar_secao_conclusoes(
    doc, 
    row,
    caminho_planilha_legendas,
    caminho_base_fotos 
):
    """
    Gera a seção '7. CONCLUSÕES', o 'APÊNDICE 1' e a seção de ASSINATURAS.
    O texto de conclusões é gerado dinamicamente com a lista de Terminais fiscalizados.
    """

    # --- LÓGICA DE EXTRAÇÃO DE CIDADES FISCALIZADAS (SEM CONTAGEM) ---
    
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

    # Texto 1 Dinâmico (Revertido para a contagem fixa de 9, mas com lista de cidades dinâmica)
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