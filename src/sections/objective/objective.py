from docx import Document
from utils import adicionar_titulo_secao, adicionar_paragrafo_justificado, adicionar_tabela_informacoes
from docx.enum.text import WD_ALIGN_PARAGRAPH

# 1. Definição dos Dados da Tabela (Pode ser transferido para um arquivo de configuração se houver muitos dados)
DADOS_INFORMACOES_GERAIS = [
    ("3.1 DO TITULAR", ""),
    ("Titular:", "Empresa Pernambucana de Transportes Intermunicipal (EPTI)"),
    ("Endereço:", "Av. Caxangá, 2.200  Cordeiro  Recife/PE  CEP: 50.711-000"),
    ("Responsável:", "ANTÔNIO CARLOS REINAUX GOMES"),
    ("3.2 DO REGULADO", ""),
    ("Regulado:", "SOCICAM - Administração, Projetos e Representações Ltda"),
    ("Responsável:", "THIAGO DUARTE PIMENTEL"),
    ("Endereço:", "Avenida Prefeito Antônio Pereira, S/N  Várzea  Recife/PE  CEP: 50.950-030"),
    ("Representantes para acompanhar:", "Monalisa da Silva Pereira (Recife/TIP)"),
    ("3.3 DO REGULADOR", ""),
    ("Regulador:", "Agência de Regulação de Pernambuco (Arpe)"),
    ("Diretor Presidente:", "CARLOS PORTO FILHO"),
    ("Endereço:", "Avenida Conselheiro Rosa e Silva, 975, Aflitos, Recife/PE, CEP: 52.050-020."),
    ("Estacionamento:", "Rua do Futuro, 150, Aflitos, Recife/PE."),
    ("Responsáveis pela fiscalização:", "Alcides Vieira de Azevedo Bezerra; Enildo Manoel da Silva Júnior"),
    ("Período da Fiscalização:", "22 a 30 de Setembro de 2025."),
    ("Tipo de Fiscalização:", "Direta e periódica."),
]

def gerar_secao_objetivo(doc: Document):
    adicionar_titulo_secao(doc, "2. OBJETIVO")

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