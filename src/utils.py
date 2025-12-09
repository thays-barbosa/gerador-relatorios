from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import load_workbook
import os
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image
import io
from docx import Document


# Funções auxiliares de formatação:
def adicionar_paragrafo_justificado(doc, texto, tamanho_fonte=12):
    """Adiciona um parágrafo com texto justificado."""
    paragrafo = doc.add_paragraph(texto)
    paragrafo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY_LOW


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
    ws = wb.active

    for coluna in ws.columns:
        max_length = 0
        coluna_letra = coluna[0].column_letter

        for celula in coluna:
            try:
                if celula.value:
                    max_length = max(max_length, len(str(celula.value)))
            except:
                pass

        # Define largura da coluna com margem extra
        ajuste = max_length + 2
        ws.column_dimensions[coluna_letra].width = ajuste

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