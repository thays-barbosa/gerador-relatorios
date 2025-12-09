from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import adicionar_titulo_secao


def gerar_secao_recomendacoes(doc: Document, row):
    adicionar_titulo_secao(doc, "6. RECOMENDAÇÕES")

    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    par.add_run(
       "Considerando as disposições do Contrato de Concessão, em especial, o Anexo III – Regulamento Interno dos Terminais" 
       "Rodoviários, aprovado pela Resolução Arpe no 46, de 07 de abril de 2008 (Antiga no 06/2008), bem como a legislação" 
       "aplicável, devem ser observadas pela SOCICAM as seguintes recomendações:" 
       "" 
       "Garantir condições de segurança, higiene, acessibilidade e conforto aos usuários dos Terminais Rodoviários," 
       "sejam passageiros, público em geral, comerciantes neles estabelecidos, empresas de transportes e de seus empregados." 
       "" 
       "Exigir a utilização de EPI adequados, inclusive por funcionários de empresas terceirizadas que prestem serviços nos Terminais Rodoviários." 
       "" 
       "Providenciar a correta manutenção (evitar o vencimento) de extintores de incêndios nos Terminais Rodoviários." 
       "" 
       "Instalar, sempre que necessário, aviso de sinalização de segurança, principalmente em pontos de risco de acidentes." 
       "" 
       "Levantar a necessidade de manutenção das cobertas de todos os Terminais Rodoviários, em especial, placas cimentícias das testeiras e vigas, considerando que quatro das nove NC registradas neste Relatório apontam" 
       "para essa questão. Assim, solicita-se laudos técnicos sobre a situação das cobertas, suas condições estruturais e funcionalidade de drenagem pluvial." 
       "" 
      
    )