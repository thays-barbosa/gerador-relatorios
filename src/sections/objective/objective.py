from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pandas as pd
from typing import List, Tuple

from utils import adicionar_titulo_secao, adicionar_paragrafo_justificado, adicionar_tabela_informacoes



def gerar_secao_objetivo(doc: Document, row: pd.Series):
    """
    Gera as seções 2. OBJETIVO e 3. INFORMAÇÕES GERAIS,
    puxando o Responsável e o Período da linha de dados da fiscalização (row).
    
    Args:
        doc (Document): O objeto Documento do python-docx.
        row (pd.Series): A linha de dados (registro) da fiscalização atual no Pandas.
    """

    try:
        # Puxa os dados da linha de fiscalização atual
        # O 'str()' é usado para garantir que o Pandas.Series vire uma string.
        responsavel_fiscalizacao = str(row['Pessoal Responsável'])
        periodo_fiscalizacao = str(row['Período da Fiscalização'])
    except KeyError as e:
        # Tratamento de erro caso a coluna não seja encontrada
        print(f"ALERTA: Coluna '{e.args[0]}' não encontrada no DataFrame. Usando valor padrão.")
        responsavel_fiscalizacao = "ERRO: Coluna de responsável não encontrada"
        periodo_fiscalizacao = "ERRO: Coluna de período não encontrada"
    except Exception as e:
        print(f"ALERTA: Erro ao extrair dados da linha: {e}")
        responsavel_fiscalizacao = "ERRO INTERNO"
        periodo_fiscalizacao = "ERRO INTERNO"

    DADOS_INFORMACOES_GERAIS: List[Tuple[str, str]] = [
        ("3.1 DO TITULAR", ""),
        ("Titular:", "Empresa Pernambucana de Transportes Intermunicipal (EPTI)"),
        ("Endereço:", "Av. Caxangá, 2.200 Cordeiro Recife/PE CEP: 50.711-000"),
        ("Responsável:", "ANTÔNIO CARLOS REINAUX GOMES"),
        ("3.2 DO REGULADO", ""),
        ("Regulado:", "SOCICAM - Administração, Projetos e Representações Ltda"),
        ("Responsável:", "THIAGO DUARTE PIMENTEL"),
        ("Endereço:", "Avenida Prefeito Antônio Pereira, S/N Várzea Recife/PE CEP: 50.950-030"),
        ("Representantes para acompanhar:", "Monalisa da Silva Pereira (Recife/TIP)"),
        ("3.3 DO REGULADOR", ""),
        ("Regulador:", "Agência de Regulação de Pernambuco (Arpe)"),
        ("Diretor Presidente:", "CARLOS PORTO FILHO"),
        ("Endereço:", "Avenida Conselheiro Rosa e Silva, 975, Aflitos, Recife/PE, CEP: 52.050-020."),
        ("Estacionamento:", "Rua do Futuro, 150, Aflitos, Recife/PE."),
 
        ("Responsáveis pela fiscalização:", responsavel_fiscalizacao),
        ("Período da Fiscalização:", periodo_fiscalizacao),
   

        ("Tipo de Fiscalização:", "Direta e periódica."),
    ]

    
    adicionar_titulo_secao(doc, "2. OBJETIVO")

    doc.add_paragraph() 

    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    par.add_run(
        "A fiscalização direta e periódica dos Terminais Rodoviários Intermunicipais de Passageiros concedidos à SOCICAM, tem "
        "por objetivo verificar as condições de conservação, limpeza e higiene das áreas de embarque e desembarque, dos "
        "sanitários, as condições do pavimento das vias de circulação interna, a infraestrutura oferecida, a segurança e o "
        "atendimento ao usuário, bem como toda estrutura para funcionamento desses terminais. Dessa forma a ação de "
        "fiscalização da Arpe verifica o grau de conformidade dessas instalações com o Contrato de Concessão, bem como com "
        "a legislação e normas vigentes de modo a determinar e/ou recomendar medidas corretivas, com foco na qualidade "
        "dos serviços prestados. "
    )

    doc.add_paragraph()
   
    adicionar_tabela_informacoes(doc, DADOS_INFORMACOES_GERAIS)
    doc.add_paragraph()