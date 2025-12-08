from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import adicionar_titulo_secao

def gerar_secao_introducao(doc: Document):
    adicionar_titulo_secao(doc, "1. INTRODUÇÃO")

    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    par.add_run(
        "A Coordenadoria de Transportes e Rodovias da Arpe realizou vistoria nos Terminais Rodoviários Intermunicipais de "
        "Passageiros concedidos com o objetivo de verificar as condições operacionais, de conservação, de manutenção e de "
        "segurança e da qualidade do serviço prestado nos referidos terminais, conforme Contrato de Concessão de Serviço "
        "Público No 1.041.080/08, firmado entre o Governo do Estado, atualmente representado pela Empresa Pernambucana "
        "de Transportes Intermunicipal (EPTI) e a SOCICAM - Administração, Projetos e Representações Ltda (SOCICAM) visando "
        "a operação, manutenção e administração de terminais rodoviários no Estado de Pernambuco, com execução de obras "
        "de reforma e construção, incluindo, ainda, a cessão de uso de espaços para a exploração comercial através de locação "
        "e publicidade. "
    )
    