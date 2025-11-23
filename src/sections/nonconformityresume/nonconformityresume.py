from typing import List, Dict, Any
import pandas as pd
from docx.document import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml.shared import OxmlElement
import re

from utils import (
    adicionar_titulo_secao,
    aplicar_estilo_corpo,
    formatar_data_df, 
    aplicar_estilo_paragrafo_compacto,
    adicionar_paragrafo_justificado, 
    encontrar_dados_na_base
)

def _aplicar_cor_fundo_celula(cell, cor_hex: str = "D9D9D9"):
    """Aplica cor de fundo (shading) a uma célula da tabela."""
    tc = cell._element
    tcPr = tc.get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:fill"), cor_hex)
    tcPr.append(shading)

def _aplicar_estilo_resumo(run, negrito: bool = False):
    """Aplica o estilo de corpo para o resumo da tabela."""
    aplicar_estilo_corpo(run, negrito=negrito)

def _formatar_nome_terminal(nome_bruto: str) -> str:
    """Limpa e formata o nome do terminal para exibição na tabela."""
    return str(nome_bruto).upper().replace("TERMINAL DE ", "").replace("TERMINAL DO ", "").strip()

def _safe_split_resumo(texto: Any) -> List[str]:
    """
    Divide strings por ';' e agora também por ':' e filtra valores vazios ou 'nan'.
    """
    s = str(texto).strip()
    if not s or s.lower() == 'nan': return []
    
    # Substitui todos os ":" por ";", e então divide por ";" para tratar ambos os delimitadores
    processed_content = s.replace(":", ";")
    
    return [x.strip() for x in processed_content.split(';') if x.strip()]

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


def gerar_secao_resumo_nao_conformidades(
    doc: Document,
    row: pd.Series,
    nao_conformidades_df: pd.DataFrame,
    processo_info: Dict[str, str],
    df_base_nc: pd.DataFrame,
    ano_user: str,
    proc_user: str,
    monit_user: str
):
    """
    Gera a seção 4: Resumo da Situação das Não Conformidades Monitoradas em formato de tabela.
    """
    id_fisc = row["ID da Fiscalização"]
    processo_ctr = processo_info.get("Processo CTR Nº", "XX/XXXX")
    carta_bruta = str(processo_info.get("Carta SAP/PER/ARPE Nº", "XXX/XXXX"))
    carta_limpa = carta_bruta.replace(";", " ").strip()
    periodo_bruto = str(processo_info.get("Periodo de Vistoria da ARPE", "DATA INDEFINIDA"))
    periodo_vistoria = periodo_bruto.replace(";", " e ")

    # Filtra as NCs para o ID de Fiscalização atual
    nc_fisc = nao_conformidades_df[nao_conformidades_df["ID da Fiscalização"] == id_fisc].copy()
    if nc_fisc.empty:
        doc.add_paragraph("Nenhuma não conformidade registrada.")
        return

    # Adiciona título da seção 4
    doc.add_paragraph().paragraph_format.space_after = Pt(12)
    adicionar_titulo_secao(doc, "4. RESUMO DA SITUAÇÃO DAS NÃO CONFORMIDADES MONITORADAS")
    
    # Cria a tabela
    tabela = doc.add_table(rows=1, cols=5)
    tabela.style = "Table Grid"
    tabela.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Define os cabeçalhos da tabela
    headers = [
        "TERMINAL",
        f"NÃO CONFORMIDADE\nRELATÓRIO ARPE/CTR\n{processo_ctr}",
        f"INFORMAÇÃO SOCICAM\nCarta SAP/PER/ARPE\n{carta_limpa}",
        f"VISTORIA DA ARPE\n{periodo_vistoria}",
        "SITUAÇÃO"
    ]
    # Define as larguras das colunas
    col_widths = [Inches(1.0), Inches(2.2), Inches(1.9), Inches(1.5), Inches(0.8)]
    
    # Preenche o cabeçalho
    for i, titulo in enumerate(headers):
        cell = tabela.rows[0].cells[i]
        cell.text = titulo
        _aplicar_cor_fundo_celula(cell) # Aplica cor de fundo
        cell.width = col_widths[i]
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        # Aplica estilo (negrito) ao texto do cabeçalho
        if cell.paragraphs[0].runs:
            _aplicar_estilo_resumo(cell.paragraphs[0].runs[0], negrito=True)

    current_row_index = 1
    # Agrupa e itera por Terminal
    for terminal_bruto, grupo in nc_fisc.groupby("Terminal"):
        primeira_celula_terminal = None
        nome_terminal = _formatar_nome_terminal(terminal_bruto)
        lista_dados = []
        ids_adicionados = set()

        # Itera sobre as linhas de NCs dentro do terminal
        for _, linha in grupo.iterrows():
            # 1. Extração da Chave de Busca (Prioriza Legenda -> Constatação)
            raw_const = str(linha.get("Legenda da Foto", "")).strip()
            if not raw_const or raw_const.lower() == "nan":
                 raw_const = str(linha.get("Constatação", "")).strip()

            # Splita (agora aceita ':' e ';')
            constatacoes = _safe_split_resumo(raw_const)
            
            # 2. Extração da Informação SOCICAM (Prioriza campo 'carta' -> campo normal)
            raw_info = str(linha.get("Informação SOCICAM carta", "")).strip()
            if not raw_info or raw_info.lower() == 'nan':
                raw_info = str(linha.get("Informação SOCICAM", "")).strip()
            # Splita (agora aceita ':' e ';')
            infos = _safe_split_resumo(raw_info)
            
            # 3. Adapta listas para o mesmo tamanho, preenchendo com vazios/N/A
            max_len = max(len(constatacoes), len(infos))
            if len(constatacoes) < max_len: constatacoes.extend([""] * (max_len - len(constatacoes)))
            if len(infos) < max_len: infos.extend(["N/A"] * (max_len - len(infos)))

            # Itera sobre cada item individual
            for i in range(max_len):
                txt_busca = constatacoes[i]
                if not txt_busca: continue
                
                info_socicam_texto = infos[i]
                
                # Busca ID, descrição, data e situação na base
                dados_base = encontrar_dados_na_base(
                    txt_busca, 
                    df_base_nc, 
                    ano_user, 
                    proc_user, 
                    monit_user,
                    terminal_user=str(terminal_bruto)
                )
                
                # Evita duplicidade
                if dados_base["id"] in ids_adicionados and dados_base["id"] != "ID_NAO_ENCONTRADO":
                    continue
                
                if dados_base["id"] != "ID_NAO_ENCONTRADO":
                    ids_adicionados.add(dados_base["id"])
                
                # Limpa o prefixo ID da descrição ANTES de adicionar à lista
                descricao_limpa = _limpar_prefixo_id(dados_base["id"], dados_base["descricao"])

                # Adiciona os dados coletados à lista
                lista_dados.append({
                    "id": dados_base["id"],
                    "desc": descricao_limpa, # Usa a descrição limpa
                    "socicam": info_socicam_texto,
                    "data": dados_base["data_vistoria"],
                    "situacao": dados_base["situacao"]
                })

        # Ordena os dados pelo ID da NC antes de inserir na tabela
        lista_dados.sort(key=lambda x: x["id"])

        # Insere as linhas da tabela para cada NC
        for idx, item in enumerate(lista_dados):
            cells = tabela.add_row().cells
            
            # Coluna 1: TERMINAL (Merge vertical na primeira linha)
            if idx == 0:
                primeira_celula_terminal = cells[0]
                cells[0].text = nome_terminal
                cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                # Aplica estilo (negrito)
                if cells[0].paragraphs[0].runs:
                    _aplicar_estilo_resumo(cells[0].paragraphs[0].runs[0], negrito=True)
            
            # Coluna 2: NÃO CONFORMIDADE (ID e Descrição)
            p_nc = cells[1].paragraphs[0]
            aplicar_estilo_paragrafo_compacto(p_nc)
            r_id = p_nc.add_run(f"{item['id']}")
            _aplicar_estilo_resumo(r_id, negrito=True)
            
            # MUDANÇA AQUI: Alterado de f" - {item['desc']}" para f" {item['desc']}"
            r_desc = p_nc.add_run(f" {item['desc']}") 
            _aplicar_estilo_resumo(r_desc)
            
            # Coluna 3: INFORMAÇÃO SOCICAM
            cells[2].text = item["socicam"]
            aplicar_estilo_paragrafo_compacto(cells[2].paragraphs[0])
            if cells[2].paragraphs[0].runs: _aplicar_estilo_resumo(cells[2].paragraphs[0].runs[0])
            
            # Coluna 4: VISTORIA DA ARPE (Data)
            cells[3].text = item["data"]
            cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if cells[3].paragraphs[0].runs: _aplicar_estilo_resumo(cells[3].paragraphs[0].runs[0])
            
            # Coluna 5: SITUAÇÃO
            s = str(item["situacao"]).upper()
            cells[4].text = s
            cells[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            # Negrito se a situação for "PENDENTE", "NÃO CONFORME" ou "NAO CONFORME"
            if cells[4].paragraphs[0].runs:
                _aplicar_estilo_resumo(cells[4].paragraphs[0].runs[0], negrito=(s in ["PENDENTE", "NÃO CONFORME", "NAO CONFORME"]))

            current_row_index += 1

        # Mescla a célula do Terminal verticalmente
        if len(lista_dados) > 1 and primeira_celula_terminal:
            ultima_celula = tabela.cell(current_row_index - 1, 0)
            primeira_celula_terminal.merge(ultima_celula)

    doc.add_paragraph()
