from typing import List, Dict
import pandas as pd
from docx.document import Document
from docx.shared import Pt, Inches
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml.shared import OxmlElement
import re # Necessário para a limpeza do nome do Terminal

from utils import (
    adicionar_titulo_secao,
    aplicar_estilo_corpo,
    formatar_data_df,
    aplicar_estilo_paragrafo_compacto,
    adicionar_paragrafo_justificado
)

CHAVE_ID_ARPE = "ID da não conformidade"
CHAVE_NC_DESC = "Não Conformidade"
CHAVE_INF_SOCICAM = "Informação SOCICAM carta"
CHAVE_DATA_VISTORIA = "Vistoria da Arpe"
CHAVE_SITUACAO = "Situação"


def _aplicar_cor_fundo_celula(cell, cor_hex: str = "D9D9D9"):
    """Aplica cor de fundo (sombrado) a uma célula da tabela."""
    tc = cell._element
    tcPr = tc.get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:fill"), cor_hex)
    tcPr.append(shading)


def _aplicar_estilo_resumo(run, negrito: bool = False):
    """Aplica estilo de corpo padrão (Arial 11pt) para células do resumo."""
    aplicar_estilo_corpo(run, negrito=negrito)


def _formatar_nome_terminal(nome_bruto: str) -> str:
    """Formata o nome do terminal (remoção de prefixos e capitalização)."""
    if not isinstance(nome_bruto, str):
        return ""
    nome_maiusculo = nome_bruto.upper()
    nome_limpo = (
        nome_maiusculo
        .replace("TERMINAL DE ", "")
        .replace("TERMINAL DO ", "")
        .replace("DE ", "")
        .strip()
    )
    return nome_limpo


def _extrair_sigla(nome_bruto: str) -> str:
    """Extrai a sigla do terminal (texto entre parênteses ou as 3 primeiras letras)."""
    if not isinstance(nome_bruto, str):
        return ""

    if "(" in nome_bruto and ")" in nome_bruto:
        start, end = nome_bruto.find("(") + 1, nome_bruto.find(")")
        if start < end:
            return nome_bruto[start:end].strip()

    return _formatar_nome_terminal(nome_bruto)[:3].upper()


def _safe_split_resumo(linha: pd.Series, col_name: str, default: str) -> List[str]:
    """Divide de forma segura o conteúdo de uma célula do DataFrame."""
    content = str(linha.get(col_name, default)).strip()

    if not content or content.lower() == "nan":
        return [] if col_name == CHAVE_NC_DESC else [default]

    # Campo especial: IDs agrupados com ';' ou ':'
    if col_name == CHAVE_ID_ARPE:
        if ":" in content:
            return [d.strip() for d in content.replace(";", ":").split(":")]
        return [d.strip() for d in content.split(";")]

    return [d.strip() for d in content.split(";")]


def _adaptar_lista(
    lista: List[str],
    tamanho_alvo: int,
    default_val: str,
    repetir_valor_unico: bool
) -> List[str]:
    """
    Garante que a lista tenha o tamanho correto.
    """
    if len(lista) == tamanho_alvo:
        return lista
    if len(lista) > tamanho_alvo:
        return lista[:tamanho_alvo]

    if len(lista) == 1 and tamanho_alvo > 1 and repetir_valor_unico:
        return [lista[0]] * tamanho_alvo

    nova_lista = list(lista)
    nova_lista.extend([default_val] * (tamanho_alvo - len(lista)))
    return nova_lista


def _formatar_cidades_str(nc_fisc: pd.DataFrame) -> str:
    """Extrai e formata os nomes dos municípios a partir da coluna 'Terminal'."""
    terminais_unicos = nc_fisc["Terminal"].dropna().unique()
    cidades = []
    
    for terminal in terminais_unicos:
        # Remove a parte "Terminal de "
        nome_limpo = terminal.replace("Terminal de ", "").strip()
        
        # Remove a abreviação entre parênteses no final (Ex: "Caruaru (CAR)" -> "Caruaru")
        nome_limpo = re.sub(r'\s*\([^)]*\)$', '', nome_limpo).strip()
        
        if nome_limpo:
            cidades.append(nome_limpo)
    
    cidades = sorted(list(set(cidades)))
    
    if len(cidades) > 1:
        # Formata para 'Cidade A, Cidade B e Cidade C'
        return ", ".join(cidades[:-1]) + f" e {cidades[-1]}"
        
    elif len(cidades) == 1:
        return cidades[0]
        
    return "diversos municípios"

def _formatar_periodo_vistoria(processo_info: Dict[str, str]) -> str:
    """Extrai e formata os períodos de vistoria separados por ';'."""
    
    periodos_raw = str(processo_info.get("Periodo de Vistoria da ARPE", "")).strip()
    
    if not periodos_raw or periodos_raw.lower() == 'nan':
        return "DATA INDEFINIDA"
        
    # Divide os períodos pela vírgula
    periodos = [p.strip() for p in periodos_raw.split(';') if p.strip()]
    
    if len(periodos) == 1:
        return periodos[0]
    
    # Formata como 'Período A, Período B e Período C'
    partes = periodos[:-1]
    ultima = periodos[-1]
    
    return f"{', '.join(partes)} e {ultima}"

def _limpar_string_cabecalho(texto_bruto: str, default: str = "XXX/XXXX") -> str:
    """
    Simplifica strings com múltiplos valores (separados por ';', ',')
    para exibição limpa no cabeçalho da tabela.
    """
    texto = str(texto_bruto).strip()
    if not texto or texto.lower() == 'nan':
        return default
    
    # Substitui delimitadores comuns (;, /) por espaço, depois normaliza múltiplos espaços
    texto_limpo = re.sub(r'[;,]', ' ', texto)
    texto_limpo = re.sub(r'\s+', ' ', texto_limpo).strip()
    
    return texto_limpo

def gerar_secao_resumo_nao_conformidades(
    doc: Document,
    row: pd.Series,
    nao_conformidades_df: pd.DataFrame,
    processo_info: Dict[str, str] # Novo argumento
):
    """Gera a seção 4 - Resumo da Situação das Não Conformidades Monitoradas."""

    # --- Extração de dados dinâmicos ---
    id_fisc = row["ID da Fiscalização"]
    nc_fisc = nao_conformidades_df[nao_conformidades_df["ID da Fiscalização"] == id_fisc].copy()
    
    if nc_fisc.empty:
        doc.add_paragraph().paragraph_format.space_after = Pt(12)
        adicionar_titulo_secao(doc, "4. RESUMO DA SITUAÇÃO DAS NÃO CONFORMIDADES MONITORADAS")
        doc.add_paragraph("Nenhuma não conformidade registrada.")
        return

    # Dados da aba Processos (processo_info)
    processo_ctr = processo_info.get("Processo CTR Nº", "XX/XXXX")
    # MODIFICAÇÃO CHAVE AQUI: Limpar a string da Carta para o cabeçalho
    carta_sap_per_arpe_bruta = processo_info.get("Carta SAP/PER/ARPE Nº", "XXX/XXXX")
    carta_sap_per_arpe_limpa = _limpar_string_cabecalho(carta_sap_per_arpe_bruta)

    # 1. (TERMINAL) - Lista de municípios limpos
    cidades_str = _formatar_cidades_str(nc_fisc) 
    
    # 2. (DATA) - Períodos de vistoria formatados
    periodo_vistoria_str = _formatar_periodo_vistoria(processo_info) 
    
    # --- Geração da Seção ---
    
    doc.add_paragraph().paragraph_format.space_after = Pt(12)
    adicionar_titulo_secao(doc, "4. RESUMO DA SITUAÇÃO DAS NÃO CONFORMIDADES MONITORADAS")

    # Dinamizando o texto introdutório com Terminais e Períodos
    texto_introducao = (
        "O Quadro 1, a seguir, resume os resultados das vistorias da equipe da Arpe nos Terminais Rodoviários "
        "Intermunicipais concedidos à SOCICAM, as atividades aconteceram nos municípios de "
        f"{cidades_str} no período de {periodo_vistoria_str}."
    )
    adicionar_paragrafo_justificado(doc, texto_introducao)

    par_quadro = doc.add_paragraph()
    par_quadro.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par_quadro.paragraph_format.space_before = Pt(12)
    par_quadro.paragraph_format.space_after = Pt(6)

    run1 = par_quadro.add_run("Quadro 1 - ")
    aplicar_estilo_corpo(run1, negrito=True)

    # Dinamizando o título do quadro
    run2 = par_quadro.add_run(
        f"Resumo da Situação das Não Conformidades Pendentes - RELATÓRIO ARPE/CTR {processo_ctr}"
    )
    aplicar_estilo_corpo(run2, negrito=True)

    # Criação da tabela principal
    tabela = doc.add_table(rows=1, cols=5)
    tabela.style = "Table Grid"
    tabela.alignment = WD_TABLE_ALIGNMENT.CENTER

    col_widths = [Inches(1.0), Inches(2.2), Inches(1.9), Inches(1.5), Inches(0.8)]

    # MODIFICAÇÃO: Dinamizando o cabeçalho das colunas com os dados do Processo
    headers = [
        "TERMINAL",
        f"NÃO CONFORMIDADE\nRELATÓRIO ARPE/CTR\n{processo_ctr}", # Processo CTR Nº
        f"INFORMAÇÃO SOCICAM\nCarta SAP/PER/ARPE\n{carta_sap_per_arpe_limpa}", # Carta SAP/PER/ARPE Nº (LIMPA)
        f"VISTORIA DA ARPE\n{periodo_vistoria_str}", # Período de Vistoria da ARPE
        "SITUAÇÃO"
    ]

    cabecalho = tabela.rows[0].cells
    for i, titulo in enumerate(headers):
        cabecalho[i].text = titulo
        _aplicar_cor_fundo_celula(cabecalho[i])
        cabecalho[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cabecalho[i].width = col_widths[i]
        for par in cabecalho[i].paragraphs:
            if par.runs:
                _aplicar_estilo_resumo(par.runs[0], negrito=True)
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Preenchimento da tabela (Lógica mantida, pois está correta)
    current_row_index = 1

    for terminal_bruto, grupo in nc_fisc.groupby("Terminal"):
        start_row_index = current_row_index
        primeira_celula_terminal = None
        nome_terminal_limpo = _formatar_nome_terminal(terminal_bruto)
        nc_list_individual = []

        for _, linha in grupo.iterrows():
            descricoes = _safe_split_resumo(linha, CHAVE_NC_DESC, "")
            ids = _safe_split_resumo(linha, CHAVE_ID_ARPE, "")
            inf_socicam = _safe_split_resumo(linha, CHAVE_INF_SOCICAM, "N/A")
            situacoes = _safe_split_resumo(linha, CHAVE_SITUACAO, "N/A")
            datas_vistoria = _safe_split_resumo(linha, CHAVE_DATA_VISTORIA, "N/A")

            num_descricoes = len(descricoes)
            ids = _adaptar_lista(ids, num_descricoes, "ID_FALTANDO", False)
            inf_socicam = _adaptar_lista(inf_socicam, num_descricoes, "N/A", False)
            situacoes = _adaptar_lista(situacoes, num_descricoes, "N/A", False)
            datas_vistoria = _adaptar_lista(datas_vistoria, num_descricoes, "N/A", False)

            for i in range(num_descricoes):
                data_formatada = (
                    formatar_data_df(datas_vistoria[i].strip())
                    if datas_vistoria[i].strip() != "N/A"
                    else "N/A"
                )
                nc_list_individual.append({
                    "id": ids[i].strip(),
                    "desc": descricoes[i].strip(),
                    "socicam": inf_socicam[i].strip(),
                    "data": data_formatada,
                    "situacao": situacoes[i].strip()
                })

        # Ordena por ID antes de inserir
        nc_list_individual.sort(key=lambda x: x["id"])

        for idx, nc_data in enumerate(nc_list_individual):
            row_cells = tabela.add_row().cells
            for i, cell in enumerate(row_cells):
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                cell.width = col_widths[i]

            # Coluna 0: Terminal
            if idx == 0:
                primeira_celula_terminal = row_cells[0]
                par = primeira_celula_terminal.paragraphs[0]
                run_terminal = par.add_run(nome_terminal_limpo)
                _aplicar_estilo_resumo(run_terminal, negrito=True)
                par.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                row_cells[0].text = ""

            # Coluna 1: Não Conformidade
            par_nc = row_cells[1].paragraphs[0]
            aplicar_estilo_paragrafo_compacto(par_nc)
            par_nc.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run_id = par_nc.add_run(f"{nc_data['id']}")
            _aplicar_estilo_resumo(run_id, negrito=True)
            run_desc = par_nc.add_run(f" - {nc_data['desc']}")
            _aplicar_estilo_resumo(run_desc)

            # Coluna 2: Informação SOCICAM
            par_inf = row_cells[2].paragraphs[0]
            par_inf.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run_inf = par_inf.add_run(nc_data["socicam"])
            _aplicar_estilo_resumo(run_inf)

            # Coluna 3: Vistoria da Arpe
            par_data = row_cells[3].paragraphs[0]
            par_data.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run_data = par_data.add_run(nc_data["data"])
            _aplicar_estilo_resumo(run_data)

            # Coluna 4: Situação
            situacao = nc_data["situacao"].upper().strip()
            negrito_sit = situacao in ("PENDENTE", "NÃO CONFORME", "NAO CONFORME")
            par_sit = row_cells[4].paragraphs[0]
            par_sit.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_sit = par_sit.add_run(situacao)
            _aplicar_estilo_resumo(run_sit, negrito=negrito_sit)

            current_row_index += 1

        # Merge da célula do terminal
        if len(nc_list_individual) > 1 and primeira_celula_terminal:
            ultima_celula = tabela.cell(current_row_index - 1, 0)
            primeira_celula_terminal.merge(ultima_celula)

    espaco = doc.add_paragraph()
    espaco.paragraph_format.space_after = Pt(24)
