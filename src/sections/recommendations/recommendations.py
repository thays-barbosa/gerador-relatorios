from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import adicionar_titulo_secao, adicionar_paragrafo_justificado # Usando o utilitário de justificação

def gerar_secao_recomendacoes(doc: Document, row):
    """
    Gera a seção '6. RECOMENDAÇÕES', formatando os itens como uma lista de bullet points.
    """
    
    # 1. Título Principal
    adicionar_titulo_secao(doc, "6. RECOMENDAÇÕES")
    doc.add_paragraph() # Espaço após o título

    # 2. Parágrafo Introdutório (Justificado)
    texto_intro = (
        "Considerando as disposições do Contrato de Concessão, em especial, o Anexo III – Regulamento Interno dos Terminais "
        "Rodoviários, aprovado pela Resolução Arpe nº 46, de 07 de abril de 2008 (Antiga nº 06/2008), bem como a legislação "
        "aplicável, devem ser observadas pela SOCICAM as seguintes recomendações:"
    )
    # Usando o utilitário para parágrafo justificado
    adicionar_paragrafo_justificado(doc, texto_intro)

    # -----------------------------------------------------------
    # 3. Lista de Recomendações (Bullet Points)
    # -----------------------------------------------------------
    
    recomendacoes = [
        # Item 1
        "Garantir condições de segurança, higiene, acessibilidade e conforto aos usuários dos Terminais Rodoviários, "
        "sejam passageiros, público em geral, comerciantes neles estabelecidos, empresas de transportes e de seus empregados.",
        
        # Item 2
        "Exigir a utilização de EPI adequados, inclusive por funcionários de empresas terceirizadas que prestem serviços nos Terminais Rodoviários.",
        
        # Item 3
        "Providenciar a correta manutenção (evitar o vencimento) de extintores de incêndios nos Terminais Rodoviários.",
        
        # Item 4
        "Instalar, sempre que necessário, aviso de sinalização de segurança, principalmente em pontos de risco de acidentes.",
        
        # Item 5 (Recomendação longa sobre cobertas)
        "Levantar a necessidade de manutenção das cobertas de todos os Terminais Rodoviários, em especial, placas cimentícias das "
        "testeiras e vigas, considerando que quatro das nove NC registradas neste Relatório apontam "
        "para essa questão. Assim, solicita-se laudos técnicos sobre a situação das cobertas, suas condições estruturais "
        "e funcionalidade de drenagem pluvial."
    ]

    for item in recomendacoes:
        par_item = doc.add_paragraph(item, style='List Bullet')
        # Opcional: Para garantir que o texto dentro do bullet point seja justificado, 
        # se o estilo padrão não fizer isso (comente se o estilo List Bullet já for suficiente)
        par_item.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        
    
    # Adiciona um espaço final
    doc.add_paragraph()