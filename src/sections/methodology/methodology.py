from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import adicionar_titulo_secao
from datetime import datetime
import pandas as pd

def gerar_secao_metodologia(doc: Document, row):
    adicionar_titulo_secao(doc, "3. METODOLOGIA")

    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    par.add_run(
        "A fiscalização direta e periódica realizada pela Coordenadoria de Transportes e Rodovias da Arpe está submetida a uma "
        "metodologia organizada em três etapas: Preparação e Planejamento, Execução da Fiscalização e Monitoramento e Avaliação. "
       
        "Preparação e Planejamento - compreende a organização e estruturação das atividades preliminares à execução da "
        "fiscalização, destacando-se a elaboração e o envio de avisos de fiscalização à Concessionária e demais atividades de "
        "suporte à fiscalização, bem como a análise de fiscalizações anteriores com a identificação de eventuais Não Conformidades pendentes. "

        "Execução da Fiscalização - a execução da fiscalização é pautada por um arcabouço de normas e diretrizes, " 
        "possibilitando que todas as etapas sejam desenvolvidas de maneira eficiente e em conformidade aos padrões "
        "estabelecidos, destacando-se: "

        "Lei no 13.254, de 21 de junho de 2007, alterada pela Lei no 15.200, de 17 de dezembro de 2013, e regulamentada pelo Decreto no 40.559, de 31 de março de 2014. "
        "Resoluções Arpe no 46, de 07 de abril de 2008 (Antiga no 06/2008), alterada pela Resolução ARPE no 53, de 26 de janeiro de 2009 (Antiga 003/2009); e no 083, de 30 de julho de 2013. " 
        "Contrato de Concessão de Serviço Público No 1.041.080/08, de 19 de setembro de 2008 e aditivos, em especial, o Segundo Termo Aditivo ao Contrato de Concessão, de 29 de setembro de 2017. " 
        "Normas Técnicas da ABNT." 


        "Monitoramento e Avaliação - Esta etapa é fundamental para garantir a eficácia das ações corretivas a serem  "
        "executadas pela Concessionária para a melhoria contínua dos serviços prestados. Os principais instrumentos do " 
        "Monitoramento e Avaliação são: Termo de Notificação e respectivo Relatório de Fiscalização, Plano de Ação da "
        "Concessionária e Relatórios de Monitoramento e Avaliação Final. " 

    )