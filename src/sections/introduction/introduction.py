from datetime import datetime
import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import adicionar_titulo_secao


def _formatar_data(valor_data) -> str:
    """
    Formata uma data recebida em diferentes formatos para o padrão dd/mm/yyyy.
    """
    if isinstance(valor_data, datetime):
        return valor_data.strftime("%d/%m/%Y")
    
    if isinstance(valor_data, str):
        try:
            return datetime.strptime(valor_data, "%Y-%m-%d").strftime("%d/%m/%Y")
        except ValueError:
            return valor_data

    return str(valor_data)


def gerar_secao_introducao(doc: Document, row: dict):
    """
    Gera a seção '1. INTRODUÇÃO' do relatório, formatando texto e datas conforme os dados da planilha.
    """
    adicionar_titulo_secao(doc, "1. INTRODUÇÃO")

    paragrafo = doc.add_paragraph()
    paragrafo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    paragrafo.add_run(
        "A Coordenadoria de Transportes e Rodovias da Arpe realizou vistoria nos Terminais Rodoviários Intermunicipais "
        "concedidos com o objetivo de verificar as condições operacionais, de conservação, de manutenção e de segurança "
        "dos referidos terminais, conforme Contrato de Serviço Público "
    )
    paragrafo.add_run(str(row["Contrato"])).bold = True
    paragrafo.add_run(
        ", firmado entre o Governo do Estado, representado pela Secretaria de Transportes (SETRA) e a SOCICAM - "
        "Administração, Projetos e Representações Ltda. A ação foi no dia "
    )

    data_formatada = _formatar_data(row.get("Data"))
    paragrafo.add_run(data_formatada).bold = True
    paragrafo.add_run(", exclusivamente no ")
    paragrafo.add_run(str(row["Local"])).bold = True

   
    periodo = str(row.get("Período", "")).strip()
    if periodo and pd.notna(periodo):
        paragrafo.add_run(f" e nos dias {periodo}, nas cidades de Caruaru, Garanhuns, Arcoverde, Serra Talhada e Petrolina. ")
    else:
        paragrafo.add_run(" e nas cidades de Caruaru, Garanhuns, Arcoverde, Serra Talhada e Petrolina. ")

    paragrafo.add_run("As visitas técnicas foram realizadas pela equipe formada por ")
    paragrafo.add_run(str(row["Pessoal Responsável"])).bold = True
    paragrafo.add_run(
        ".\n\nNeste Relatório de Fiscalização foram observadas as condições de conservação, limpeza e higiene das áreas "
        "de embarque e desembarque, dos sanitários, as condições do pavimento das vias de circulação interna, a "
        "infraestrutura oferecida, os locais de estocagem de veículos, a segurança e o atendimento ao usuário, bem como "
        "toda estrutura para funcionamento dos terminais. A equipe da Arpe conversou com os responsáveis pelos seis "
        "terminais que forneceram informações complementares à fiscalização, principalmente sobre a implantação dos "
        "sistemas contra incêndio."
    )
