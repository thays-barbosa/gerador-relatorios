from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import load_workbook
import os
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image
import io
from docx import Document
import pandas as pd  # Importado para pd.read_excel
from docx.enum.table import WD_ALIGN_VERTICAL  # Importado para WD_ALIGN_VERTICAL
from docx.table import _Cell # Importação útil

# --- Funções auxiliares de formatação: ---

def adicionar_paragrafo_justificado(doc, texto, tamanho_fonte=12):
    """Adiciona um parágrafo com texto justificado."""
    paragrafo = doc.add_paragraph(texto)
    paragrafo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    # Se precisar do JUSTIFY_LOW (justificação justificada em português):
    # paragrafo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY_LOW


def adicionar_texto_centralizado(doc, texto, tamanho_fonte=12):
    """Adiciona um parágrafo com texto centralizado."""
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run(texto)
    run.bold = True


def adicionar_titulo_secao(doc, texto):
    """Adiciona um título de seção formatado."""
    secao = doc.add_paragraph()
    secao.add_run(texto).bold = True

def adicionar_titulo_quadro(doc, texto, negrito=False, tamanho=11):
    """Adiciona um título de Quadro/Figura centralizado."""
    paragrafo = doc.add_paragraph()
    paragrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragrafo.add_run(texto)
    run.bold = negrito
    # run.font.size = Pt(tamanho) # Opcional: manter o tamanho padrão do corpo do texto ou forçar um específico


# Função para ajustar a largura das colunas (Excel)
def ajustar_largura_colunas(caminho_planilha):
    wb = load_workbook(caminho_planilha)
    # Itera sobre todas as planilhas do Excel para ajustar a largura
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        
        for coluna in ws.columns:
            max_length = 0
            coluna_letra = coluna[0].column_letter

            for celula in coluna:
                try:
                    if celula.value:
                        # Adiciona uma lógica para tratar números formatados como texto
                        if isinstance(celula.value, (int, float)):
                             # Se for número, trata como string formatada. Aqui simplifica para len(str)
                             length = len(str(celula.value))
                        else:
                             length = len(str(celula.value))
                             
                        max_length = max(max_length, length)
                except:
                    pass

            # Define largura da coluna com margem extra
            ajuste = max_length + 2
            # Garante que o ajuste mínimo seja feito para colunas vazias
            if ajuste > 3:
                ws.column_dimensions[coluna_letra].width = ajuste
            else:
                ws.column_dimensions[coluna_letra].width = 10 # Largura mínima padrão

    wb.save(caminho_planilha)


# Função para verificar se arquivo está em uso
def arquivo_em_uso(caminho):
    try:
        os.rename(caminho, caminho)
        return False
    except PermissionError:
        return True


def aplicar_estilo_texto(
    run, tamanho=12, negrito=False, fonte="Arial", cor_rgb=(0, 0, 0)
):
    run.font.name = fonte
    run._element.rPr.rFonts.set(qn("w:eastAsia"), fonte)
    run.font.size = Pt(tamanho)
    run.bold = negrito
    run.font.color.rgb = RGBColor(*cor_rgb)


def aplicar_borda_paragrafo(paragraph):
    p = paragraph._element
    pPr = p.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    for border_name in ("top", "left", "bottom", "right"):
        border = OxmlElement(f"w:{border_name}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "4")
        border.set(qn("w:space"), "2")
        border.set(qn("w:color"), "000000")
        borders.append(border)
    pPr.append(borders)


def adicionar_legenda_formatada(doc, texto):
    par = doc.add_paragraph()
    run = par.add_run(texto)
    aplicar_estilo_texto(run, tamanho=10, fonte="Arial", cor_rgb=(90, 90, 90))
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    aplicar_borda_paragrafo(par)


def processar_imagem_para_relatorio(caminho_imagem, largura_max=1024, qualidade=80):
    # Abre a imagem
    img = Image.open(caminho_imagem)
    # Converte para RGB se necessário (evita problemas com PNG/transparência)
    if img.mode != "RGB":
        img = img.convert("RGB")
    # Redimensiona mantendo proporção
    if img.width > largura_max:
        proporcao = largura_max / float(img.width)
        altura_nova = int(float(img.height) * proporcao)
        img = img.resize((largura_max, altura_nova), Image.LANCZOS)
    # Salva em memória, sem metadados, com compressão JPEG
    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=qualidade, optimize=True)
    buffer.seek(0)
    return buffer

# --- FUNÇÕES DE FORMATAÇÃO DE TABELA DE INFORMAÇÕES GERAIS ---

def aplicar_fundo_cinza(celula):
    """
    Aplica um sombreamento cinza (DDDDDD) em uma célula.
    """
    try:
        shading_elm = OxmlElement('w:shd')
        shading_elm.set(qn('w:val'), 'clear') # Tipo de preenchimento (solid)
        # Cor cinza (DDDDDD)
        shading_elm.set(qn('w:fill'), 'DDDDDD') 
        celula._element.tcPr.append(shading_elm)
    except Exception as e:
        print(f"Não foi possível aplicar fundo cinza à célula: {e}")


def adicionar_cabecalho_tabela(doc, texto, tamanho_fonte=12):
    """Adiciona um texto em caixa alta e negrito para servir como cabeçalho de tabela/seção (SEM FUNDO CINZA)."""
    
    # Cria um parágrafo normal
    paragrafo = doc.add_paragraph()
    run = paragrafo.add_run(texto.upper())
    run.bold = True
    
    paragrafo.alignment = WD_ALIGN_PARAGRAPH.LEFT 
    
    doc.add_paragraph() # Adiciona espaço após o título
    

def adicionar_tabela_informacoes(doc, dados_tabela):
    """Cria a tabela de Informações Gerais com formatação de fundo cinza, negrito e merge de células."""
    
    # 1. Título acima da tabela (sem fundo cinza)
    adicionar_cabecalho_tabela(doc, "INFORMAÇÕES GERAIS")
    
    # 2. Cria a tabela principal
    tabela = doc.add_table(rows=len(dados_tabela), cols=2)
    tabela.autofit = False
    
    # AJUSTE DA LARGURA DAS COLUNAS (6.0cm e 11.0cm)
    tabela.columns[0].width = Cm(6.0) 
    tabela.columns[1].width = Cm(11.0)
    
    # Define as linhas que devem ter fundo cinza e merge de células
    titulos_cinza = ["3.1 DO TITULAR", "3.2 DO REGULADO", "3.3 DO REGULADOR"]
    
    # Aplica o estilo de borda nativo do Word
    tabela.style = 'Table Grid'
    
    for i, (rotulo, valor) in enumerate(dados_tabela):
        celula_rotulo = tabela.rows[i].cells[0]
        celula_valor = tabela.rows[i].cells[1]
        
        # --- Lógica de Título Cinza (Merge de Células) ---
        if rotulo in titulos_cinza:
            # 1. Aplica fundo cinza em AMBAS as células
            aplicar_fundo_cinza(celula_rotulo)
            aplicar_fundo_cinza(celula_valor)
            
            # 2. Faz o merge das células
            celula_principal = celula_rotulo.merge(celula_valor)
            
            # 3. Adiciona e formata o texto
            par = celula_principal.paragraphs[0] if celula_principal.paragraphs else celula_principal.add_paragraph()
            par.text = ''
            run = par.add_run(rotulo)
            run.bold = True
            
            par.alignment = WD_ALIGN_PARAGRAPH.LEFT 
            
            continue # Pula para a próxima linha de dados

        # --- Lógica de Dados Normais ---
        
        # Rótulo (Primeira Coluna) - SEMPRE em negrito na imagem
        par_rotulo = celula_rotulo.paragraphs[0] if celula_rotulo.paragraphs else celula_rotulo.add_paragraph()
        par_rotulo.text = ''
        run_rotulo = par_rotulo.add_run(rotulo)
        run_rotulo.bold = True
        par_rotulo.alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        # Valor (Segunda Coluna) - Negrito Condicional
        par_valor = celula_valor.paragraphs[0] if celula_valor.paragraphs else celula_valor.add_paragraph()
        par_valor.text = ''
        run_valor = par_valor.add_run(valor)
        par_valor.alignment = WD_ALIGN_PARAGRAPH.LEFT

        # Lógica para aplicar negrito nos valores específicos (Responsável e Diretor Presidente)
        if rotulo in ["Responsável:", "Diretor Presidente:"]:
            run_valor.bold = True

# --- FUNÇÕES PARA APÊNDICE FOTOGRÁFICO ---

def adicionar_imagem_na_celula(celula, caminho_imagem, largura_max_cm=7.5):
    """
    Processa e insere uma imagem dentro de uma célula de tabela,
    garantindo que ela se ajuste à largura máxima.
    """
    # Adiciona um novo parágrafo na célula para a imagem
    paragrafo_img = celula.add_paragraph()
    paragrafo_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Processa a imagem para otimização (usando sua função existente)
    imagem_buffer = processar_imagem_para_relatorio(caminho_imagem)
    
    # Adiciona a imagem ao parágrafo
    run = paragrafo_img.add_run()
    # Adiciona a imagem usando o buffer e define a largura
    run.add_picture(imagem_buffer, width=Cm(largura_max_cm))
    
    # Define a alinhamento vertical da célula para topo (opcional)
    celula.vertical_alignment = WD_ALIGN_VERTICAL.TOP


def adicionar_apendice_fotos(doc, caminho_base_fotos, id_fiscalizacao, caminho_planilha_legendas):
    """
    Gera o Apêndice 1 com as fotos dinâmicas, seguindo o layout 2x2.
    
    Args:
        doc (Document): O objeto do documento Word.
        caminho_base_fotos (str): O caminho para a subpasta F0 (ex: 'CTR-02-2024/F0').
        id_fiscalizacao (str): O ID para buscar a linha de legendas.
        caminho_planilha_legendas (str): O caminho do arquivo Excel com as legendas.
    """
    # 1. Busca e prepara as legendas
    # Assumindo que o nome da aba foi corrigido para "Não-conformidades " (com espaço)
    df_legendas = pd.read_excel(caminho_planilha_legendas, sheet_name="Não-conformidades ")
    
    # Filtra a linha correta pelo ID
    linha_legenda = df_legendas[df_legendas['ID da Fiscalização'] == id_fiscalizacao]
    
    if linha_legenda.empty:
        adicionar_paragrafo_justificado(doc, f"AVISO: Legendas não encontradas na planilha para o ID: {id_fiscalizacao}. As fotos não serão legendadas.")
        # Retorna lista vazia de legendas para que as fotos sejam incluídas sem legenda
        legendas = [] 
        # return # Não queremos retornar, queremos incluir as fotos sem legenda
    else:
        # Extrai o texto da coluna 'Legenda da Foto' e separa
        # *CORRIGIDO*: Usando 'Legenda da Foto' (com F maiúsculo) para evitar o KeyError
        texto_legendas = str(linha_legenda['Legenda da Foto'].iloc[0]) 
        legendas = [l.strip() for l in texto_legendas.split(';') if l.strip()]

    # 2. Busca e prepara os caminhos das fotos
    # Obtém a lista de arquivos de imagem na pasta especificada
    try:
        arquivos_fotos = sorted([
            os.path.join(caminho_base_fotos, f)
            for f in os.listdir(caminho_base_fotos)
            if f.lower().endswith(('.png', '.jpg', '.jpeg'))
        ])
    except FileNotFoundError:
        # Este erro deve ser capturado antes no report.py, mas é bom ter uma segurança
        adicionar_paragrafo_justificado(doc, f"ERRO INTERNO: O caminho de fotos '{caminho_base_fotos}' não foi encontrado.")
        return

    if not arquivos_fotos:
        # AVISO solicitado pelo usuário para subpasta vazia
        adicionar_paragrafo_justificado(doc, f"AVISO: A subpasta procurada está vazia, Seu Apêndice não terá fotos.")
        return

    # 3. Cria a estrutura da tabela (layout de 2x2)
    
    # Número de fotos deve ser par para o layout 2x2.
    num_fotos = len(arquivos_fotos)
    # Calcula o número de linhas necessárias (duas células por linha, mais uma linha de legenda para cada linha de foto)
    num_linhas_tabela = ((num_fotos + 1) // 2) * 2 
    
    tabela = doc.add_table(rows=num_linhas_tabela, cols=2)
    tabela.autofit = False
    tabela.style = 'Table Grid'
    
    # Define a largura das colunas (metade da página, menos margem)
    largura_coluna_cm = 8.5 
    tabela.columns[0].width = Cm(largura_coluna_cm)
    tabela.columns[1].width = Cm(largura_coluna_cm)
    
    # 4. Popula a tabela
    for i in range(num_linhas_tabela // 2): # Itera sobre as linhas de fotos
        # Linha para Fotos
        linha_foto_idx = i * 2
        # Linha para Legendas
        linha_legenda_idx = i * 2 + 1 
        
        for j in range(2): # Coluna 0 e Coluna 1
            foto_idx = i * 2 + j
            
            if foto_idx < num_fotos:
                caminho_foto = arquivos_fotos[foto_idx]
                
                # --- A. Célula da Foto ---
                celula_foto = tabela.cell(linha_foto_idx, j)
                adicionar_imagem_na_celula(celula_foto, caminho_foto, largura_max_cm=7.5) 

                # --- B. Célula da Legenda ---
                celula_legenda = tabela.cell(linha_legenda_idx, j)
                
                # Alinhamento da Legenda
                celula_legenda.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                par_legenda = celula_legenda.paragraphs[0] if celula_legenda.paragraphs else celula_legenda.add_paragraph()
                par_legenda.alignment = WD_ALIGN_PARAGRAPH.CENTER

                # Pega a legenda correta (ou string de aviso)
                texto_legenda = legendas[foto_idx] if foto_idx < len(legendas) else f"Legenda {foto_idx + 1} não fornecida na planilha."
                
                # Adiciona e formata o texto da legenda
                run_legenda = par_legenda.add_run(texto_legenda)
                # Aplica estilo: tamanho=10, fonte="Arial", cor_rgb=(90, 90, 90)
                aplicar_estilo_texto(run_legenda, tamanho=10, fonte="Arial", cor_rgb=(90, 90, 90))
                
            else:
                # Se não houver mais fotos (para preencher as células que sobram)
                celula_foto_vazia = tabela.cell(linha_foto_idx, j)
                celula_legenda_vazia = tabela.cell(linha_legenda_idx, j)
                
                # Adiciona um espaço em branco para manter a célula formatada
                celula_foto_vazia.add_paragraph("").alignment = WD_ALIGN_PARAGRAPH.CENTER
                celula_legenda_vazia.add_paragraph("").alignment = WD_ALIGN_PARAGRAPH.CENTER

# --- SUBSTITUIR FUNÇÃO EM utils.py ---

# --- SUBSTITUIR FUNÇÃO EM utils.py ---

from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL 

def adicionar_tabela_abreviaturas(doc, df_abreviaturas):
    """
    Cria uma tabela simples de duas colunas (SIGLA e DEFINIÇÃO)
    com base nos dados do DataFrame, aplicando negrito na coluna Sigla e alinhamento vertical.
    """
    # Cria a tabela (número de linhas: cabeçalho + dados)
    tabela = doc.add_table(rows=len(df_abreviaturas) + 1, cols=2)
    tabela.autofit = False
    
    # Aplica o estilo de borda nativo do Word e centraliza a tabela
    tabela.style = 'Table Grid'
    tabela.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 🚨 AJUSTE DE LARGURA: Acentuando a diferença para que a coluna SIGLA fique visivelmente menor
    largura_sigla_cm = 2.5 # REDUZIDO para 2.5cm
    largura_definicao_cm = 14.5 # AUMENTADO para 14.5cm
    tabela.columns[0].width = Cm(largura_sigla_cm)
    tabela.columns[1].width = Cm(largura_definicao_cm)
    
    # 1. Cabeçalho
    header_cells = tabela.rows[0].cells
    
    # Configuração comum para células do Cabeçalho (Vertical alignment)
    for cell in header_cells:
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Célula SIGLA (Header)
    par_sigla = header_cells[0].paragraphs[0]
    par_sigla.text = ""
    run_sigla = par_sigla.add_run("SIGLA")
    run_sigla.bold = True
    par_sigla.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Célula DEFINIÇÃO (Header)
    par_def = header_cells[1].paragraphs[0]
    par_def.text = ""
    run_def = par_def.add_run("DEFINIÇÃO") # Texto no cabeçalho sem espaço
    run_def.bold = True
    par_def.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 2. Preenche as linhas de dados
    for i, row in df_abreviaturas.iterrows():
        cells = tabela.rows[i + 1].cells
        
        # Configuração comum para células de Dados (Vertical alignment)
        for cell in cells:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER 

        # Sigla (Centralizada e Negrito)
        par_sigla_data = cells[0].paragraphs[0]
        par_sigla_data.text = ""
        run_sigla_data = par_sigla_data.add_run(str(row['Sigla']))
        par_sigla_data.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Definição (Esquerda)
        par_def_data = cells[1].paragraphs[0]
        par_def_data.text = ""
        # Usando 'Definição ' (com espaço) para corresponder ao DataFrame do usuário
        par_def_data.add_run(str(row['Definição '])) 
        par_def_data.alignment = WD_ALIGN_PARAGRAPH.LEFT