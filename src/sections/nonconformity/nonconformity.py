from docx.shared import Pt, RGBColor
from docx.document import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from typing import Any, Dict, List
import pandas as pd
import re

# Presume-se que 'utils' contém as funções auxiliares necessárias
from utils import (
    adicionar_paragrafo_justificado,
    adicionar_titulo_secao,
    adicionar_paragrafo_info_compacta,
    aplicar_estilo_corpo,
    aplicar_estilo_titulo,
    encontrar_dados_na_base
)

def _limpar_prefixo_id(id_prefix: str, texto_bruto: str) -> str:
    """Remove o ID da NC e seus delimitadores (-, :) do início da descrição para evitar duplicação."""
    desc_str = str(texto_bruto).strip()
    id_prefix = id_prefix.strip()

    # Regex que encontra o ID no início, seguido por 0 ou mais espaços,
    # um separador opcional (traço ou dois pontos), e 0 ou mais espaços.
    pattern = re.compile(re.escape(id_prefix) + r'\s*[-:]?\s*', re.IGNORECASE)
    match = pattern.match(desc_str)

    if match:
        # Remove a parte que corresponde ao ID e separador
        desc_str = desc_str[match.end():].strip()

    # Garante que a primeira letra da descrição limpa esteja em maiúsculo
    if desc_str and desc_str[0].islower():
        desc_str = desc_str[0].upper() + desc_str[1:]
        
    return desc_str


def _inserir_texto_nc(doc: Document, nc_titulo_identificador: str, descricao_bruta: str):
    """Insere o título e a descrição de uma Não Conformidade (NC) no documento, limpando o ID da descrição."""
    
    # Remove o ID do início da descrição para evitar a duplicação visual
    descricao_limpa = _limpar_prefixo_id(nc_titulo_identificador, descricao_bruta)

    paragrafo_nc = doc.add_paragraph()
    
    # Título da NC (negrito e sublinhado)
    run_titulo = paragrafo_nc.add_run(f"Não Conformidade {nc_titulo_identificador}")
    aplicar_estilo_corpo(run_titulo, negrito=True)
    run_titulo.underline = True
    
    # Separador: Alterado de " - " para um único espaço " "
    run_traco = paragrafo_nc.add_run(" ")
    aplicar_estilo_corpo(run_traco)
    
    # Descrição da NC (agora limpa)
    run_desc = paragrafo_nc.add_run(descricao_limpa)
    aplicar_estilo_corpo(run_desc)
    
    paragrafo_nc.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY_LOW

def _inserir_linhas_info(doc: Document, info_socicam: str, constatacao_monit: str, analise_arpe: str):
    """Insere as três linhas de informação compacta (SOCICAM, Constatação, ARPE)."""
    # Trata valores vazios/NaN/None para evitar erros e garantir texto legível
    if str(info_socicam).strip().lower() in ['nan', 'none', '']: info_socicam = "N/A"
    if str(constatacao_monit).strip().lower() in ['nan', 'none', '']: constatacao_monit = "N/A"
    if str(analise_arpe).strip().lower() in ['nan', 'none', '']: analise_arpe = "N/A"
    
    # Adiciona as informações usando a função auxiliar
    adicionar_paragrafo_info_compacta(doc, "Informação da SOCICAM: ", str(info_socicam))
    adicionar_paragrafo_info_compacta(doc, "Constatação: ", str(constatacao_monit)) 
    adicionar_paragrafo_info_compacta(doc, "Análise da ARPE: ", str(analise_arpe))

def _safe_split(texto: Any) -> List[str]:
    """
    Divide strings por ';' e agora também por ':' e filtra valores vazios ou 'nan'.
    """
    content = str(texto).strip()
    if not content or content.lower() == "nan": return []
    
    # Substitui todos os ":" por ";", e então divide por ";" para tratar ambos os delimitadores
    processed_content = content.replace(":", ";") 
    
    return [item.strip() for item in processed_content.split(";") if item.strip()]

def gerar_secao_nao_conformidades_constatadas(
    doc: Document, 
    row: pd.Series, 
    nao_conformidades_df: pd.DataFrame, 
    FOTOS_DIR: str, 
    processo_info: Dict[str, str],
    df_base_nc: pd.DataFrame,
    ano_user: str,
    proc_user: str,
    monit_user: str
):
    """
    Gera a seção 3 do relatório, detalhando os resultados das vistorias das NCs pendentes.
    """
    id_fiscalizacao = row["ID da Fiscalização"]
    processo_ctr = processo_info.get("Processo CTR Nº", "XX/XXXX")
    cartas_raw = str(processo_info.get("Carta SAP/PER/ARPE Nº", "")).strip()
    
    # Gera o texto da carta, se disponível
    texto_cartas = f"constante da Carta SAP/PER/ARPE N° {cartas_raw}, " if (cartas_raw and cartas_raw.lower() != 'nan') else ""

    # Filtra as NCs para a fiscalização atual
    nc_fiscalizacao = nao_conformidades_df[
        nao_conformidades_df["ID da Fiscalização"] == id_fiscalizacao
    ].copy()

    # Adiciona Título da Seção
    adicionar_titulo_secao(doc, "3. RESULTADO DAS VISTORIAS DAS NÃO CONFORMIDADES PENDENTES")
    
    # Adiciona Parágrafo Introdutório Justificado
    adicionar_paragrafo_justificado(
        doc,
        (
            "Estão registrados para cada Terminal Rodoviário os resultados da verificação pela Arpe das ações "
            f"desenvolvidas pela SOCICAM, {texto_cartas}"
            "para solucionar as Não Conformidades ainda pendentes apresentadas no Relatório de Fiscalização Técnico-"
            f"Operacional ARPE/CTR nº {processo_ctr}."
        ),
    )

    num_terminal = 1
    # Agrupa e itera por Terminal
    for terminal, grupo_terminal in nc_fiscalizacao.groupby("Terminal"):
        # Adiciona Título do Terminal (3.X - NOME DO TERMINAL)
        par_terminal = doc.add_paragraph()
        run_terminal = par_terminal.add_run(f"3.{num_terminal} - {terminal.upper()}")
        aplicar_estilo_titulo(run_terminal) 
        par_terminal.paragraph_format.space_before = Pt(12)

        ncs_processadas = set()

        # Itera sobre as linhas de NCs dentro do terminal
        for _, nc_data in grupo_terminal.iterrows():
            # 1. Extração da Chave de Busca (Constatação/Legenda da Foto)
            # Prioriza: Legenda -> Constatação para a busca
            raw_key = str(nc_data.get("Legenda da Foto", "")).strip()
            if not raw_key or raw_key.lower() == "nan":
                raw_key = str(nc_data.get("Constatação", "")).strip()
            
            # 2. Divide os campos por ';', permitindo múltiplos itens por linha (agora aceita ':' e ';')
            constatacoes_monit = _safe_split(raw_key)
            infos_socicam = _safe_split(nc_data.get("Informação SOCICAM", ""))
            analises_arpe = _safe_split(nc_data.get("Análise da Arpe", ""))

            # 3. Adapta listas para o tamanho máximo (preenchendo com vazios ou "N/A" se necessário)
            max_len = max(len(constatacoes_monit), len(infos_socicam), len(analises_arpe))
            
            if len(constatacoes_monit) < max_len: 
                constatacoes_monit.extend([""] * (max_len - len(constatacoes_monit)))
            
            if len(infos_socicam) < max_len: 
                infos_socicam.extend(["N/A"] * (max_len - len(infos_socicam)))
            if len(analises_arpe) < max_len: 
                analises_arpe.extend(["N/A"] * (max_len - len(analises_arpe)))

            # Itera sobre cada item individual (se houver split por ';')
            for i in range(max_len):
                texto_busca = constatacoes_monit[i]
                if not texto_busca: continue 

                # 4. Busca o ID oficial e a descrição na base
                dados_base = encontrar_dados_na_base(
                    texto_busca, 
                    df_base_nc, 
                    ano_user, 
                    proc_user, 
                    monit_user,
                    terminal_user=str(terminal)
                )
                
                # 5. Evita processar a mesma NC duas vezes 
                if dados_base["id"] != "ID_NAO_ENCONTRADO":
                    if dados_base["id"] in ncs_processadas:
                        continue 
                    ncs_processadas.add(dados_base["id"]) 

                # 6. Insere os dados no documento
                _inserir_texto_nc(doc, dados_base["id"], dados_base["descricao"])
                _inserir_linhas_info(doc, infos_socicam[i], texto_busca, analises_arpe[i])
                
                doc.add_paragraph() 
                
        num_terminal += 1