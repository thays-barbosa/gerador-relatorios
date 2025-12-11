import os
import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.shared import Inches, Pt 

# Importa as funções de utilidade (incluindo as de controle de espaçamento)
from utils import (
    parear_responsaveis_e_matriculas, 
    formatar_data_capa, 
    adicionar_texto_centralizado,
    remover_espacamento_paragrafo, 
    adicionar_quebra_linha_controlada 
)


def gerar_capa(doc: Document, caminho_base_dir, info_fiscalizacao):
    """
    Gera a capa do relatório de forma dinâmica, extraindo dados da planilha,
    com controle de espaçamento para manter o layout compacto.
    """
    
    # 1. Extração e Processamento dos Dados
    
    nomes_responsaveis = info_fiscalizacao.get('Pessoal Responsável', info_fiscalizacao.get('PESSOAL RESPONSÁVEL', ''))
    matriculas_responsaveis = info_fiscalizacao.get('Matrícula do Pessoal Responsável', '') 
    profissionais = parear_responsaveis_e_matriculas(nomes_responsaveis, matriculas_responsaveis)

    data_formatada = formatar_data_capa(str(info_fiscalizacao.get('Data', ''))) 
    num_processo = info_fiscalizacao.get('Nº Processo', 'N/A')
    num_sei = info_fiscalizacao.get('Nº SEI', 'N/A')
    
    relatorio_fixo = f"RELATÓRIO DE FISCALIZAÇÃO PROC ADM Nº {num_processo} - CTR"
    sei_fixo = f"SEI Nº {num_sei}"
    
    # 2. Geração dos Elementos no Documento (Baseado em layout compacto)
    
    # --- Espaçamento Inicial (Topo da Página) ---
    # Usa quebras de linha controladas para empurrar o conteúdo
    for _ in range(5): 
        adicionar_quebra_linha_controlada(doc, altura_pt=20) 
        
    # --- Título da Capa (Fixo) ---
    # Usa a função atualizada que aplica espaçamento zero.
    adicionar_texto_centralizado(doc, "RELATÓRIO DE FISCALIZAÇÃO")
    
    # Espaçamento de separação (entre título e logo)
    adicionar_quebra_linha_controlada(doc, altura_pt=20)

    # 3. Imagem da Logo
    try:
        caminho_logo = os.path.join(caminho_base_dir, "assets/logo_arpe.jpg")
        if not os.path.exists(caminho_logo):
            raise FileNotFoundError(f"Arquivo de logo não encontrado no caminho: {caminho_logo}")
            
        # Adiciona a imagem
        doc.add_picture(caminho_logo, width=Inches(6))
        logo_arpe = doc.paragraphs[-1]
        
        # Garante espaçamento zero no parágrafo da imagem
        remover_espacamento_paragrafo(logo_arpe)
        logo_arpe.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    except Exception as e:
        print(f"\n⚠️ ERRO GRAVE AO INSERIR LOGO NA CAPA: {e}") 
        doc.add_paragraph("--- FALHA AO CARREGAR LOGO ARPE (Verifique /assets/logo_arpe.jpg) ---").alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    # --- Espaçamento após Imagem ---
    for _ in range(5): 
        adicionar_quebra_linha_controlada(doc, altura_pt=20) 
        
    # --- Títulos Intermediários (Compactos) ---
    adicionar_texto_centralizado(doc, "FISCALIZAÇÃO NOS TERMINAIS RODOVIÁRIOS INTERMUNICIPAIS DE PASSAGEIROS")
    adicionar_texto_centralizado(doc, "PRESTADOR DE SERVIÇO: SOCICAM - ADMINISTRAÇÃO, PROJETOS E REPRESENTAÇÕES LTDA")
    
    # --- Espaçamento GRANDE para o centro vertical da página (Assinaturas) ---
    for _ in range(10): # Ajuste este loop para posicionar as assinaturas
        adicionar_quebra_linha_controlada(doc, altura_pt=20) 
        
    # --- Nomes e Matrículas (Centralizados e Compactos) ---
    for profissional in profissionais:
        # Nome (Negrito - com espaçamento zero)
        par_nome = doc.add_paragraph()
        remover_espacamento_paragrafo(par_nome) 
        par_nome.alignment = WD_ALIGN_PARAGRAPH.CENTER
        par_nome.add_run(profissional['nome']).bold = True
        
        # Cargo e Matrícula (Normal - com espaçamento zero)
        par_info = doc.add_paragraph()
        remover_espacamento_paragrafo(par_info) 
        par_info.alignment = WD_ALIGN_PARAGRAPH.CENTER
        par_info.add_run(profissional['info_completa'])
        
        # Espaço controlado entre profissionais
        adicionar_quebra_linha_controlada(doc, altura_pt=10) # Espaçamento menor

    # --- Espaçamento para a Data ---
    for _ in range(3): # Ajuste o espaçamento entre assinaturas e data
        adicionar_quebra_linha_controlada(doc, altura_pt=20) 
    
    # --- Data (Mês, Ano) ---
    par_data = doc.add_paragraph()
    remover_espacamento_paragrafo(par_data) 
    par_data.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par_data.add_run(data_formatada)
    
    # --- Espaçamento até o final da Página ---
    for _ in range(5): # Ajuste este loop para posicionar os números de processo
        adicionar_quebra_linha_controlada(doc, altura_pt=20) 
        
    # --- Números de Processo (Final da Página - Compactos) ---
    
    par_relatorio = doc.add_paragraph()
    remover_espacamento_paragrafo(par_relatorio) 
    par_relatorio.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par_relatorio.add_run(relatorio_fixo).bold = True 
    
    par_sei = doc.add_paragraph()
    remover_espacamento_paragrafo(par_sei) 
    par_sei.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par_sei.add_run(sei_fixo)
    
    # Adiciona quebra de página
    doc.add_section(WD_SECTION.NEW_PAGE)