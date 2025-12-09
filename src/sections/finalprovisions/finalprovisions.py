from utils import (
    adicionar_titulo_secao,
    adicionar_paragrafo_justificado,
   
)

def gerar_secao_determinacoes_finais(doc, row):
    """
    Gera a seção '5. CONSIDERAÇÕES FINAIS' do relatório.
    Parâmetros:
    - doc: objeto Document.
    - row: linha da fiscalização (Series do DataFrame fiscalizacoes_df).
    """

    adicionar_titulo_secao(doc, "5. DETERMINAÇÕES GERAIS")

    texto1 = "Considerando os dispositivos contratuais pertinentes e visando garantir a qualidade dos serviços prestados, determina-se que a SOCICAM tome as seguintes medidas através de um plano de ação: "

    texto2 = "Manutenção e Monitoramento: adotar medidas para assegurar a manutenção, o monitoramento contínuo e o cumprimento do Programa de Manutenção dos Terminais Rodoviários, constante da proposta da SOICICAM nos subitens 9.1.1  Manutenção Preventiva; 9.1.2  manutenção Corretiva e 9.13 tabela de classificação de níveis de falha (tabela de tempos máximos para os níveis de atendimento). "

    texto3 = "Medidas imediatas para resolutividade das 9 (nove) NC constatadas, nos prazos estabelecidos, conforme disposto no Quadro 1, na coluna denominada Determinações. "

    adicionar_paragrafo_justificado(doc, texto1)
    adicionar_paragrafo_justificado(doc, texto2)
    adicionar_paragrafo_justificado(doc, texto3)
   


