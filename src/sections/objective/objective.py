from docx import Document
from utils import adicionar_titulo_secao, adicionar_paragrafo_justificado
from docx.enum.text import WD_ALIGN_PARAGRAPH

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