from typing import Tuple, Optional
import io
import os
from datetime import datetime

import pandas as pd
from PIL import Image
from openpyxl import load_workbook
from docx.document import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Pt, RGBColor, Inches

# Constantes de layout e cores
LARGURA_PADRAO_IN = Inches(6)
LARGURA_IMAGEM_LADO_A_LADO = Inches(3.25)
ALTURA_IMAGEM_LADO_A_LADO = Inches(2.7)

_COR_CINZA_SOMBRA_HEX = "BFBFBF"
_COR_PRETO_RGB = (0, 0, 0)


# -------------------------
# Utilitários de formatação
# -------------------------
def _set_run_language(run, lang_code: str = "pt-BR") -> None:
    """Define o idioma do run para evitar marcações do corretor ortográfico."""
    rPr = run._element.get_or_add_rPr()
    lang = OxmlElement("w:lang")
    lang.set(qn("w:val"), lang_code)
    lang.set(qn("w:eastAsia"), lang_code)
    lang.set(qn("w:bidi"), lang_code)
    rPr.append(lang)


def desabilitar_correcao_paragrafo(paragrafo) -> None:
    """Desativa a verificação ortográfica para o parágrafo informado."""
    pPr = paragrafo._element.get_or_add_pPr()

    no_proof = OxmlElement("w:noProof")
    pPr.append(no_proof)

    lang = OxmlElement("w:lang")
    lang.set(qn("w:val"), "pt-BR")
    pPr.append(lang)


def aplicar_sombreamento_paragrafo(paragrafo, cor_hex: str) -> None:
    """Aplica sombreamento (shading) a um parágrafo usando cor hexadecimal."""
    pPr = paragrafo._element.get_or_add_pPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), cor_hex)
    pPr.append(shading)


# -------------------------
# Estilos de texto
# -------------------------
def aplicar_estilo_texto(
    run,
    tamanho: int = 12,
    negrito: bool = False,
    fonte: str = "Arial",
    cor_rgb: Tuple[int, int, int] = (0, 0, 0),
) -> None:
    """Aplica estilos básicos a um run (fonte, tamanho, peso, cor e idioma)."""
    run.font.name = fonte
    # garante compatibilidade com fontes em eastAsia (Word)
    run._element.rPr.rFonts.set(qn("w:eastAsia"), fonte)
    run.font.size = Pt(tamanho)
    run.bold = negrito
    run.font.color.rgb = RGBColor(*cor_rgb)
    _set_run_language(run, "pt-BR")


def aplicar_estilo_corpo(run, negrito: bool = False) -> None:
    """Estilo padrão do corpo: Arial 11pt."""
    aplicar_estilo_texto(run, tamanho=11, negrito=negrito, fonte="Arial")


def aplicar_estilo_titulo(run, cor_rgb: Tuple[int, int, int] = _COR_PRETO_RGB) -> None:
    """Estilo padrão para títulos de seção: Arial 12pt negrito."""
    aplicar_estilo_texto(run, tamanho=12, negrito=True, fonte="Arial", cor_rgb=cor_rgb)


# -------------------------
# Parágrafos / textos
# -------------------------
def adicionar_paragrafo_justificado(doc: Document, texto: str, negrito: bool = False):
    """Adiciona parágrafo justificado com estilo do corpo (retorna o parágrafo)."""
    paragrafo = doc.add_paragraph()
    paragrafo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY_LOW
    run = paragrafo.add_run(texto)
    aplicar_estilo_corpo(run, negrito=negrito)
    return paragrafo


def adicionar_texto_centralizado(
    doc: Document, texto: str, tamanho_fonte: int = 12, negrito: bool = True
):
    """Adiciona parágrafo centralizado (usado na capa e assinaturas)."""
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run(texto)

    if tamanho_fonte == 12 and negrito:
        aplicar_estilo_titulo(run)
    elif tamanho_fonte == 11:
        aplicar_estilo_corpo(run, negrito=negrito)
    else:
        aplicar_estilo_texto(run, tamanho_fonte, negrito, fonte="Arial")

    desabilitar_correcao_paragrafo(paragraph)
    return paragraph


# -------------------------
# Parágrafos compactos (info)
# -------------------------
def adicionar_paragrafo_info_compacta(doc: Document, titulo: str, conteudo: str):
    """Adiciona parágrafo no formato 'Título: Conteúdo' com espaçamento compacto."""
    par = doc.add_paragraph()
    par.paragraph_format.space_before = Pt(0)
    par.paragraph_format.space_after = Pt(0)
    par.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    par.paragraph_format.line_spacing = Pt(13)

    run_titulo = par.add_run(titulo)
    aplicar_estilo_corpo(run_titulo, negrito=True)

    run_conteudo = par.add_run(conteudo)
    aplicar_estilo_corpo(run_conteudo, negrito=False)

    par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY_LOW
    return par


def aplicar_estilo_paragrafo_compacto(paragrafo) -> None:
    """Aplica formatação compacta a um parágrafo (sem alterar conteúdo)."""
    paragrafo.paragraph_format.space_before = Pt(0)
    paragrafo.paragraph_format.space_after = Pt(0)
    paragrafo.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    paragrafo.paragraph_format.line_spacing = Pt(13)


# -------------------------
# Títulos/Seções
# -------------------------
def adicionar_titulo_secao(
    doc: Document, texto: str, nivel_heading: int = 1, cor_rgb: Tuple[int, int, int] = _COR_PRETO_RGB, aplicar_sombra: bool = False
):
    """
    Adiciona um título de seção usando doc.add_heading (para sumário).
    Mantém compatibilidade caso add_heading falhe.
    """
    try:
        par = doc.add_heading(level=nivel_heading)
    except Exception as exc:
        print(f"Aviso: erro aplicando heading {nivel_heading}. Usando parágrafo. Erro: {exc}")
        par = doc.add_paragraph()

    run = par.add_run(texto)
    aplicar_estilo_titulo(run, cor_rgb=cor_rgb)

    if aplicar_sombra:
        aplicar_sombreamento_paragrafo(par, _COR_CINZA_SOMBRA_HEX)
        par.alignment = WD_ALIGN_PARAGRAPH.LEFT
        par.paragraph_format.space_after = Pt(6)
    else:
        par.paragraph_format.space_after = Pt(6)

    run.font.all_caps = True
    desabilitar_correcao_paragrafo(par)
    return par


def adicionar_texto_contexto(doc: Document, contexto_tupla: Tuple[str, str]):
    """
    Adiciona parágrafo com o contexto: TERMINAL (título em caixa alta) + texto restante.
    contexto_tupla = (terminal, texto_restante)
    """
    terminal, texto_restante = contexto_tupla
    par = doc.add_paragraph()
    par.paragraph_format.space_before = Pt(12)
    par.paragraph_format.space_after = Pt(6)
    par.alignment = WD_ALIGN_PARAGRAPH.LEFT

    run_terminal = par.add_run(terminal)
    aplicar_estilo_titulo(run_terminal)
    run_terminal.font.all_caps = True

    desabilitar_correcao_paragrafo(par)

    run_restante = par.add_run(texto_restante)
    aplicar_estilo_corpo(run_restante, negrito=False)
    return par


def _adicionar_contexto_nc(doc: Document, nc_id: str, constatacao: str):
    """Adiciona o parágrafo de contexto para uma NC: 'ID - constatação'."""
    par = doc.add_paragraph()
    par.paragraph_format.space_before = Pt(12)
    par.paragraph_format.space_after = Pt(6)
    par.alignment = WD_ALIGN_PARAGRAPH.LEFT

    run_id = par.add_run(f"{nc_id.strip()}")
    aplicar_estilo_titulo(run_id)

    run_constatacao = par.add_run(f" - {constatacao.strip()}")
    aplicar_estilo_corpo(run_constatacao, negrito=False)

    desabilitar_correcao_paragrafo(par)
    return par


def adicionar_texto_esquerda(doc: Document, texto: str, tamanho_fonte: int = 11, negrito: bool = False, compacto: bool = False):
    """Adiciona parágrafo alinhado à esquerda (opção compacto para remover espaçamentos)."""
    paragraph = doc.add_paragraph()
    if compacto:
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)

    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = paragraph.add_run(texto)

    if tamanho_fonte == 11:
        aplicar_estilo_corpo(run, negrito=negrito)
    else:
        aplicar_estilo_texto(run, tamanho_fonte, negrito)

    return paragraph


# -------------------------
# Data / Planilha
# -------------------------
def formatar_data_df(data_value) -> str:
    """
    Formata um valor de data para 'dd/mm/YYYY'.
    Aceita datetime, strings no formato 'YYYY-MM-DD' ou outras strings.
    Retorna string vazia para NaNs.
    """
    if pd.isna(data_value):
        return ""

    if isinstance(data_value, datetime):
        return data_value.strftime("%d/%m/%Y")

    if isinstance(data_value, str):
        try:
            return datetime.strptime(data_value.split()[0], "%Y-%m-%d").strftime("%d/%m/%Y")
        except ValueError:
            # se não estiver no formato YYYY-MM-DD, retorna tal como veio
            return data_value

    return str(data_value)


def ajustar_largura_colunas(caminho_planilha: str) -> None:
    """Ajusta a largura das colunas de uma planilha com base no conteúdo."""
    wb = load_workbook(caminho_planilha)
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for coluna in ws.columns:
            max_length = 0
            coluna_letra = coluna[0].column_letter

            for celula in coluna:
                try:
                    if celula.value:
                        max_length = max(max_length, len(str(celula.value)))
                except Exception:
                    # ignora células problemáticas
                    pass

            ajuste = max_length + 2
            ws.column_dimensions[coluna_letra].width = ajuste

    wb.save(caminho_planilha)


def arquivo_em_uso(caminho: str) -> bool:
    """Retorna True se o arquivo estiver em uso/aberto (PermissionError ao renomear)."""
    try:
        os.rename(caminho, caminho)
        return False
    except PermissionError:
        return True


# -------------------------
# Bordas e legendas
# -------------------------
def aplicar_borda_paragrafo(paragraph) -> None:
    """Aplica borda simples em um parágrafo via XML."""
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


def adicionar_legenda_formatada(doc: Document, texto: str) -> None:
    """Adiciona legenda centralizada com borda (usado em anexos/imagens)."""
    par = doc.add_paragraph()
    run = par.add_run(texto)
    aplicar_estilo_texto(run, tamanho=10, fonte="Arial", cor_rgb=(90, 90, 90))
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    aplicar_borda_paragrafo(par)


# -------------------------
# Processamento de imagens
# -------------------------
def processar_imagem_para_relatorio(caminho_imagem: str, largura_max: int = 1024, qualidade: int = 80) -> Optional[io.BytesIO]:
    """
    Abre, converte (para RGB), redimensiona (se necessário) e retorna um buffer JPEG.
    Em caso de erro retorna None e imprime uma mensagem no console.
    """
    if not os.path.exists(caminho_imagem):
        print(f"ERRO DE FOTO (processar_imagem): Arquivo não encontrado: {caminho_imagem}")
        return None

    try:
        img = Image.open(caminho_imagem)
    except Exception as exc:
        print(f"ERRO DE FOTO (processar_imagem): Não foi possível abrir '{caminho_imagem}'. Erro: {exc}")
        return None

    # garante modo RGB para salvar como JPEG
    if img.mode != "RGB":
        img = img.convert("RGB")

    # redimensiona proporcionalmente se a largura for maior que largura_max
    if img.width > largura_max:
        proporcao = largura_max / float(img.width)
        altura_nova = int(float(img.height) * proporcao)
        img = img.resize((largura_max, altura_nova), Image.LANCZOS)

    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=qualidade, optimize=True)
    buffer.seek(0)
    return buffer


def adicionar_imagem(doc: Document, buffer_imagem: Optional[io.BytesIO], largura_in: Optional[Inches] = None, altura_in: Optional[Inches] = None) -> None:
    """Insere imagem a partir de buffer no documento e centraliza o parágrafo."""
    largura_final = largura_in if largura_in is not None else LARGURA_PADRAO_IN
    paragrafo = doc.add_paragraph()

    if buffer_imagem:
        try:
            if altura_in is not None:
                paragrafo.add_run().add_picture(buffer_imagem, width=largura_final, height=altura_in)
            else:
                paragrafo.add_run().add_picture(buffer_imagem, width=largura_final)
        except Exception as exc:
            paragrafo.add_run(f"Erro ao inserir imagem: {exc}")
    else:
        paragrafo.add_run("🚫 Imagem indisponível.")

    paragrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()


# -------------------------
# Bordas em células
# -------------------------
def set_cell_border(cell, **kwargs) -> None:
    """
    Define bordas de célula via XML.
    Uso: set_cell_border(cell, top={'val':'single','sz':'4'}, left={...}, ...)
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    tblBorders = tcPr.first_child_found_in("w:tcBorders")
    if tblBorders is None:
        tblBorders = OxmlElement("w:tcBorders")
        tcPr.append(tblBorders)

    for border_name, attrs in kwargs.items():
        if border_name in ("top", "left", "bottom", "right", "insideH", "insideV"):
            border_element = OxmlElement(f"w:{border_name}")
            for k, v in attrs.items():
                border_element.set(qn(f"w:{k}"), str(v))

            old_border = tblBorders.find(qn(f"w:{border_name}"))
            if old_border is not None:
                tblBorders.remove(old_border)

            tblBorders.append(border_element)


def adicionar_legenda_formatada_na_celula(cell, texto: str) -> None:
    """Adiciona legenda formatada dentro de uma célula de tabela (substitui conteúdo antigo)."""
    par = cell.paragraphs[0]

    # limpa conteúdo anterior (se houver)
    if len(par.runs) > 0 or par.text.strip() != "":
        try:
            par.clear()
        except Exception:
            # se não suportar clear, recria parágrafo (fallback)
            cell._tc.get_or_add_tcPr()  # noop para evitar lint
            par = cell.paragraphs[0]

    run = par.add_run(texto)
    aplicar_estilo_texto(run, tamanho=10, fonte="Arial", cor_rgb=(90, 90, 90))
    par.alignment = WD_ALIGN_PARAGRAPH.LEFT


# -------------------------
# Inserção de imagens lado a lado
# -------------------------
def adicionar_duas_imagens_lado_a_lado(
    doc: Document,
    fotos_dir: str,
    nome_foto1: str,
    legenda1: str,
    nome_foto2: Optional[str] = None,
    legenda2: Optional[str] = None,
    contexto_nc_tupla: Optional[Tuple[str, str, str]] = None,
) -> None:
    """
    Insere uma tabela 2x2 contendo uma ou duas imagens com legendas.
    Se nome_foto2 for None, a imagem ocupa as duas colunas (centralizada).
    Se contexto_nc_tupla for fornecido, insere o contexto antes (ID - constatação).
    """
    if contexto_nc_tupla:
        _, nc_id, constatacao = contexto_nc_tupla
        _adicionar_contexto_nc(doc, nc_id, constatacao)

    tabela = doc.add_table(rows=2, cols=2)
    tabela.alignment = WD_TABLE_ALIGNMENT.CENTER
    tabela.style = "Table Grid"

    col_width_lote = LARGURA_IMAGEM_LADO_A_LADO
    largura_imagem = Inches(3.0)

    if nome_foto2 is None:
        # imagem única — mescla a primeira linha
        celula_img_mesclada = tabela.cell(0, 0).merge(tabela.cell(0, 1))
        celula_img_mesclada.width = LARGURA_PADRAO_IN

        celula_legenda_mesclada = tabela.cell(1, 0).merge(tabela.cell(1, 1))
        celula_legenda_mesclada.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        celula_legenda_mesclada.width = LARGURA_PADRAO_IN

        largura_imagem_final = LARGURA_PADRAO_IN
        caminho_imagem1 = os.path.join(fotos_dir, nome_foto1)
        buffer_img1 = processar_imagem_para_relatorio(caminho_imagem1)

        par1 = celula_img_mesclada.paragraphs[0]
        par1.alignment = WD_ALIGN_PARAGRAPH.CENTER

        if buffer_img1 is not None:
            try:
                run1 = par1.add_run()
                run1.add_picture(buffer_img1, width=largura_imagem_final, height=None)
            except Exception as exc:
                par1.add_run(f"Erro ao inserir foto única: {exc}.")
        else:
            par1.add_run(f"🚫 Imagem indisponível: {nome_foto1}")

        adicionar_legenda_formatada_na_celula(celula_legenda_mesclada, legenda1)

    else:
        # duas imagens lado a lado
        for row_idx in range(2):
            for col_idx in range(2):
                cell = tabela.cell(row_idx, col_idx)
                cell.width = col_width_lote
            # Ajusta o alinhamento vertical do último cell do loop (comportamento herdado)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

        # primeira imagem
        caminho_imagem1 = os.path.join(fotos_dir, nome_foto1)
        buffer_img1 = processar_imagem_para_relatorio(caminho_imagem1)

        par1 = tabela.cell(0, 0).paragraphs[0]
        par1.alignment = WD_ALIGN_PARAGRAPH.CENTER

        if buffer_img1 is not None:
            try:
                run1 = par1.add_run()
                run1.add_picture(buffer_img1, width=largura_imagem, height=ALTURA_IMAGEM_LADO_A_LADO)
            except Exception as exc:
                par1.add_run(f"Erro ao inserir foto 1: {exc}.")
        else:
            par1.add_run(f"🚫 Imagem indisponível: {nome_foto1}")

        # segunda imagem
        caminho_imagem2 = os.path.join(fotos_dir, nome_foto2)
        buffer_img2 = processar_imagem_para_relatorio(caminho_imagem2)

        par2 = tabela.cell(0, 1).paragraphs[0]
        par2.alignment = WD_ALIGN_PARAGRAPH.CENTER

        if buffer_img2 is not None:
            try:
                run2 = par2.add_run()
                run2.add_picture(buffer_img2, width=largura_imagem, height=ALTURA_IMAGEM_LADO_A_LADO)
            except Exception as exc:
                par2.add_run(f"⚠️ Erro ao inserir foto 2: {exc}")
        else:
            par2.add_run(f"🚫 Imagem indisponível: {nome_foto2}")

        # legendas
        cell_legenda1 = tabela.cell(1, 0)
        cell_legenda1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        adicionar_legenda_formatada_na_celula(cell_legenda1, legenda1)

        cell_legenda2 = tabela.cell(1, 1)
        cell_legenda2.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        adicionar_legenda_formatada_na_celula(cell_legenda2, legenda2 or "")

    doc.add_paragraph().paragraph_format.space_after = Pt(12)


# -------------------------
# Helpers de busca/normalização
# -------------------------
def _limpar_terminal_para_busca(terminal_nome: str) -> str:
    """
    Limpa o nome do Terminal para ser usado como prefixo de busca.
    Ex.: extrai sigla entre parênteses ou transforma espaços/pontos/hífens em '_'.
    """
    if not isinstance(terminal_nome, str):
        return ""

    if "(" in terminal_nome and ")" in terminal_nome:
        start = terminal_nome.find("(") + 1
        end = terminal_nome.find(")")
        return terminal_nome[start:end].strip().upper()

    return terminal_nome.replace(" ", "_").replace("-", "_").upper()

