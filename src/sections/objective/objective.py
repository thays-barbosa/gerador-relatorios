from docx.document import Document
from typing import Dict, Any
import pandas as pd
import re
from utils import adicionar_titulo_secao, adicionar_paragrafo_justificado


def gerar_secao_objetivo(doc: Document, row: Dict[str, Any], processo_info: Dict[str, str], nao_conformidades_df: pd.DataFrame):
    """
    Gera a seção 2 - OBJETIVO do relatório, usando dados dinâmicos da fiscalização e do processo.
    """

    # --- 1. Extração de Variáveis Simples ---
    
    # {num_monitoramento}: ID da Fiscalização (formato Xº)
    id_fisc = str(row.get("ID da Fiscalização", "X"))
    num_monitoramento = f"{id_fisc}º" if id_fisc.isdigit() else id_fisc
    
    # {ctr_original}: Processo CTR Nº
    ctr_original = processo_info.get("Processo CTR Nº", "xx/xxxx")
    
    # {doc_sei}: Doc. SEI Nº
    # CORREÇÃO AQUI: Remove o ".0" que o Excel/Pandas adiciona automaticamente
    raw_doc_sei = processo_info.get("Doc. SEI Nº", "xxxxxxx")
    doc_sei = str(raw_doc_sei).strip()
    if doc_sei.endswith(".0"):
        doc_sei = doc_sei[:-2]

    # --- 2. Extração e Formatação de {cidades_str} ---
    
    # Filtra as não-conformidades apenas para o ID da fiscalização atual
    nao_conformidades_fisc = nao_conformidades_df[
        nao_conformidades_df["ID da Fiscalização"] == row["ID da Fiscalização"]
    ]

    # Extrai os nomes únicos dos Terminais
    terminais_unicos = nao_conformidades_fisc["Terminal"].dropna().unique()

    # Limpa e extrai apenas o nome do município
    cidades = []
    for terminal in terminais_unicos:
        # Remove a parte "Terminal de "
        nome_limpo = terminal.replace("Terminal de ", "").strip()
        
        # Remove a abreviação entre parênteses no final (Ex: "Caruaru (CAR)" -> "Caruaru")
        nome_limpo = re.sub(r'\s*\([^)]*\)$', '', nome_limpo).strip()
        
        if nome_limpo:
            cidades.append(nome_limpo)
    
    # Remove duplicatas e formata como string
    cidades = sorted(list(set(cidades)))
    
    if len(cidades) > 1:
        # Formata para 'Cidade A, Cidade B e Cidade C'
        cidades_str = ", ".join(cidades[:-1]) + f" e {cidades[-1]}"
        
    elif len(cidades) == 1:
        cidades_str = cidades[0]
        
    else:
        # Fallback se não encontrar cidades
        cidades_str = "diversos municípios" 
    
    
    # --- 3. Geração da Seção ---
    
    adicionar_titulo_secao(doc, "2. OBJETIVO")

    texto_objetivo = (
        f"Este Relatório do {num_monitoramento} Monitoramento objetiva apresentar os resultados "
        f"das vistorias acerca das Não Conformidades pendentes registradas no Relatório de "
        f"Fiscalização Técnico-Operacional nº {ctr_original} (Doc. SEI nº {doc_sei}), referentes "
        f"aos Terminais Rodoviários Intermunicipais dos municípios de {cidades_str}, "
        "conforme o cronograma de atividades encaminhado pela Concessionária SOCICAM."
    )

    adicionar_paragrafo_justificado(doc, texto_objetivo)