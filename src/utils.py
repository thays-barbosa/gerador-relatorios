from typing import Tuple, Optional, Dict, Any, List
import io
import os
from datetime import datetime
from difflib import SequenceMatcher
import re 
import unicodedata

import pandas as pd
from PIL import Image
from openpyxl import load_workbook
from docx.document import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Pt, RGBColor, Inches

LARGURA_PADRAO_IN = Inches(6)
LARGURA_IMAGEM_LADO_A_LADO = Inches(3.25)
ALTURA_IMAGEM_LADO_A_LADO = Inches(2.7)
_COR_CINZA_SOMBRA_HEX = "BFBFBF"
_COR_PRETO_RGB = (0, 0, 0)

# --- Mapeamento de Terminais para Busca Dinâmica ---
TERMINAIS_MAP = {
    # Chave (Input do Usuário) : Valor (Sigla na Base/ID)
    "RECIFE": "TIP",
    "TIP": "TIP",
    "CARUARU": "CAR",
    "CAR": "CAR",
    "GARANHUNS": "GAR",
    "GAR": "GAR",
    "ARCOVERDE": "ARC",
    "ARC": "ARC",
    "PETROLINA": "PET",
    "PET": "PET",
    "SERRA": "SER",
    "ST": "SER",
    "SERRA TALHADA": "SER", 
    # Adicione novos terminais aqui
}

# --- STOPWORDS (Palavras ignoradas na busca) ---
STOPWORDS = {
    "o", "a", "os", "as", "um", "uns", "uma", "umas",
    "de", "do", "da", "dos", "das", "em", "no", "na", "nos", "nas",
    "por", "para", "com", "sem", "que", "se", "e", "ou", "ao", "aos",
    "terminal", "rodoviario", "intermunicipal", "passageiros", "lugar", "local",
    "item", "nc", "nao", "conformidade", "ver", "foto", "fotos", "vide", "imagem",
    "situacao", "detalhe", "vista", "registro", # Adicionei palavras comuns de legenda
    # Ignora nomes de cidade na comparação de TEXTO
    "recife", "tip", "caruaru", "garanhuns", "arcoverde", "petrolina", "serra", "talhada",
    "cidade", "estado", "parte", "lado", "parede", "fachada" 
}

def carregar_base_nc(caminho_base: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(caminho_base, sheet_name="BASE", dtype=str)
        df.columns = df.columns.str.strip()
        df = df.dropna(how='all')
        
        cols_texto = ["TIPO_DOC", "PROCESSO", "Evidencia_Agregada", "Evidencia_Desagregada", "Localização/VIA", "Item", "Status-ARPE", "Terminal"]
        for col in cols_texto:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip()
        
        print(f"📂 Base carregada: {len(df)} linhas.")
        return df
    except Exception as e:
        print(f"❌ Erro Base: {e}")
        return pd.DataFrame()

def _remover_acentos(texto: str) -> str:
    try:
        return "".join([c for c in unicodedata.normalize('NFKD', str(texto)) if not unicodedata.combining(c)])
    except: return str(texto)

def _singularizar(palavra: str) -> str:
    """
    Remove apenas o 's' final para lidar com plurais simples.
    """
    p = palavra.lower()
    if p.endswith('s') and len(p) > 3: 
        return p[:-1]
    return p

def _extrair_palavras_chave(texto: str) -> set:
    """
    Limpa o texto e extrai tokens (palavras) relevantes singularizados.
    """
    t = _remover_acentos(str(texto)).lower().strip()
    
    # Remove datas
    t = re.sub(r'\b\d{2}/\d{2}/\d{4}\b', ' ', t)
    
    # Remove IDs (TIP 01, 2025_01)
    t = re.sub(r"(\s|^)([a-z]{3}\s)?(\d{2,4}\_)?(\d+(\.\d+)?)\s*[-:]?\s*", " ", t)
    
    # Remove pontuação
    t = re.sub(r"[^\w\s]", " ", t)
    
    palavras = t.split()
    
    tokens = set()
    for p in palavras:
        if len(p) > 2 and p not in STOPWORDS:
            tokens.add(_singularizar(p))
            
    return tokens

def _limpar_texto_simples(texto: str) -> str:
    t = _remover_acentos(str(texto)).lower().strip()
    t = re.sub(r'\b\d{2}/\d{2}/\d{4}\b', ' ', t)
    t = re.sub(r"[^\w\s]", " ", t)
    return " ".join(t.split())

def _limpar_terminal_para_busca(terminal_nome: str) -> str:
    return str(terminal_nome).upper().replace(" ", "_")

def _limpar_processo_para_match(processo: str) -> str:
    t = str(processo).upper().strip()
    t = re.sub(r'(CTR\s*Nº?|Nº?|:)\s*', ' ', t) 
    t = re.sub(r'\s+', ' ', t).strip()
    return t

def _ajustar_numero_nc(ano: str, processo: str, monit: str, item: str, prefixo_terminal: str) -> str:
    if not item:
        return "ID_NAO_ENCONTRADO"
    # Lógica para garantir que o prefixo do terminal (ex: CAR) e o ano estejam na frente
    item_limpo = str(item).strip().replace('.', '_', 1).replace('.', '').replace('_', '.', 1).strip()
    item_sem_prefixo = re.sub(r'([A-Z]{3}\s)?(\d{4}_)?', '', item_limpo).strip()
    return f"{prefixo_terminal} {ano}_{item_sem_prefixo}"

def encontrar_dados_na_base(
    texto_busca: str,
    df_base: pd.DataFrame,
    ano_user: str,
    proc_user: str,
    monit_user: str,
    terminal_user: str = None
) -> Dict[str, str]:
    
    # 1. Tokens para Match de Palavras-Chave
    tokens_busca = _extrair_palavras_chave(texto_busca)
    
    # 2. Texto simples para Match de Sequência
    texto_busca_seq = _limpar_texto_simples(texto_busca)

    resultado = {
        "id": "ID_NAO_ENCONTRADO",
        "descricao": texto_busca, 
        "situacao": "Não Identificado",
        "data_vistoria": ""
    }

    if df_base.empty or not texto_busca.strip(): return resultado

    df_filt = df_base.copy()

    # --- FILTROS DE CONTEXTO (Ano, Processo, Doc) ---
    if "Ano" in df_filt.columns and ano_user and str(ano_user).isdigit():
        df_filt = df_filt[df_filt["Ano"].astype(str).str.strip() == str(ano_user)]

    if "PROCESSO" in df_filt.columns and proc_user:
        p_clean_user = _limpar_processo_para_match(proc_user)
        df_filt['PROCESSO_LIMPO'] = df_filt["PROCESSO"].apply(_limpar_processo_para_match)
        df_filt = df_filt[df_filt["PROCESSO_LIMPO"].str.contains(p_clean_user, na=False)]
        df_filt = df_filt.drop(columns=['PROCESSO_LIMPO'])

    if "TIPO_DOC" in df_filt.columns and monit_user:
        mask_monit = df_filt["TIPO_DOC"].astype(str).str.upper().str.contains("MONIT", na=False)
        mask_num = df_filt["TIPO_DOC"].astype(str).str.contains(str(monit_user), na=False)
        df_filt = df_filt[mask_monit & mask_num]
        
    # --- FILTRO DE TERMINAL (REFORÇADO) ---
    prefixo_terminal = ""
    
    if terminal_user:
        t_user_upper = _limpar_texto_simples(terminal_user).upper()
        sigla = ""
        
        # 1. Identifica a sigla do terminal:
        for key, value in TERMINAIS_MAP.items():
            if key in t_user_upper:
                sigla = value
                prefixo_terminal = value
                break
        
        if sigla:
            # 2. **REFORÇO CRÍTICO**: Filtra a base pela SIGLA no campo "Item" OU "ID da Fiscalização"
            colunas_possiveis = [c for c in ["Item", "ID da Fiscalização"] if c in df_filt.columns]
            
            if colunas_possiveis:
                # Cria uma máscara que verifica se a sigla (ex: CAR) está no início
                # de qualquer uma das colunas identificadoras
                mask_sigla_item = df_filt[colunas_possiveis[0]].astype(str).str.upper().str.startswith(sigla, na=False)
                
                if len(colunas_possiveis) > 1:
                    for col_nome in colunas_possiveis[1:]:
                        mask_sigla_item = mask_sigla_item | df_filt[col_nome].astype(str).str.upper().str.startswith(sigla, na=False)
                
                df_filt = df_filt[mask_sigla_item]
                
            # 3. Adiciona o filtro de Localização/VIA (se disponível) como filtro secundário
            if "Localização/VIA" in df_filt.columns:
                term_col = df_filt["Localização/VIA"].astype(str).str.upper()
                nome_cidade_busca = ""
                for key, value in TERMINAIS_MAP.items():
                    if value == sigla and len(key) > 3: 
                         nome_cidade_busca = key
                         break
                
                mask_sigla_loc = term_col.str.contains(sigla, na=False)
                if nome_cidade_busca:
                    mask_nome_loc = term_col.str.contains(nome_cidade_busca, na=False)
                    df_filt = df_filt[mask_sigla_loc | mask_nome_loc]
                else:
                    df_filt = df_filt[mask_sigla_loc]


    if df_filt.empty: 
        print(f"Alerta: Base vazia após filtro de Terminal {terminal_user}.")
        return resultado

    melhor_score = 0
    melhor_row = None
    
    cols_busca = [c for c in ["Evidencia_Agregada", "Evidencia_Desagregada"] if c in df_filt.columns]

    for _, row in df_filt.iterrows():
        txt_base_raw = " ".join([str(row.get(c, "")) for c in cols_busca])
        
        # A. Keyword Match (Pontuação de Palavras-Chave)
        tokens_base = _extrair_palavras_chave(txt_base_raw)
        if len(tokens_busca) > 0:
            interseccao = tokens_busca.intersection(tokens_base)
            score_keywords = len(interseccao) / len(tokens_busca)
        else:
            score_keywords = 0
        
        # B. Sequence Match (Similaridade de Texto)
        txt_base_seq = _limpar_texto_simples(txt_base_raw)
        score_seq = SequenceMatcher(None, texto_busca_seq, txt_base_seq).ratio()
        
        # *** DEFINIÇÃO DE SCORE FINAL (O MAIOR ENTRE OS DOIS) ***
        # Usa o score mais alto entre a similaridade de sequência e a contagem de palavras-chave.
        score_final = max(score_keywords, score_seq) 
        # **********************************
            
        if score_final > melhor_score:
            melhor_score = score_final
            melhor_row = row

    # AJUSTE CHAVE FINAL: Threshold aumentado para 50% para evitar "falsos positivos"
    if melhor_row is not None and melhor_score >= 0.50:
        
        evidencia_chave = str(melhor_row.get("Evidencia_Agregada", "")).strip()
        item_raw = str(melhor_row.get("Item", "")).strip()
        base_id = item_raw.split(".")[0] if "." in item_raw else item_raw

        id_final_formatado = _ajustar_numero_nc(ano_user, proc_user, monit_user, base_id, prefixo_terminal)

        # Filtra o grupo de NCs para o relatório (ex: 7, 7.1, 7.2)
        df_grupo = df_filt[df_filt["Item"].astype(str).str.strip().str.startswith(base_id)].sort_values(by="Item")

        linhas_texto = []
        if evidencia_chave and evidencia_chave.lower() != 'nan':
             linhas_texto.append(evidencia_chave)
        
        for _, row_g in df_grupo.iterrows():
            item_n = str(row_g.get("Item", "")).strip()
            desag = str(row_g.get("Evidencia_Desagregada", "")).strip()
            # Adiciona o item desagregado, se diferente do agregado principal
            if desag and desag.lower() != 'nan' and desag != evidencia_chave:
                linhas_texto.append(f"{item_n} - {desag}")
        
        if not linhas_texto:
            desc_final = texto_busca
        else:
            linhas_texto = list(dict.fromkeys(linhas_texto))
            desc_final = "\n".join(linhas_texto)

        resultado = {
            "id": id_final_formatado, 
            "descricao": desc_final,
            "situacao": str(melhor_row.get("Status-ARPE", "N/A")).strip(),
            "data_vistoria": formatar_data_df(melhor_row.get("DATA FISC", ""))
        }

    return resultado

# --- FUNÇÕES DE FORMATAÇÃO E AUXILIARES (AS SEGUINTES FUNÇÕES NÃO FORAM ALTERADAS) ---

def _set_run_language(run, lang_code: str = "pt-BR") -> None:
    rPr = run._element.get_or_add_rPr()
    lang = OxmlElement("w:lang")
    lang.set(qn("w:val"), lang_code)
    lang.set(qn("w:eastAsia"), lang_code)
    lang.set(qn("w:bidi"), lang_code)
    rPr.append(lang)

def desabilitar_correcao_paragrafo(paragrafo) -> None:
    pPr = paragrafo._element.get_or_add_pPr()
    no_proof = OxmlElement("w:noProof")
    pPr.append(no_proof)
    lang = OxmlElement("w:lang")
    lang.set(qn("w:val"), "pt-BR")
    pPr.append(lang)

def aplicar_sombreamento_paragrafo(paragrafo, cor_hex: str) -> None:
    pPr = paragrafo._element.get_or_add_pPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), cor_hex)
    pPr.append(shading)

def aplicar_estilo_texto(run, tamanho: int = 12, negrito: bool = False, fonte: str = "Arial", cor_rgb: Tuple[int, int, int] = (0, 0, 0)) -> None:
    run.font.name = fonte
    run._element.rPr.rFonts.set(qn("w:eastAsia"), fonte)
    run.font.size = Pt(tamanho)
    run.bold = negrito
    run.font.color.rgb = RGBColor(*cor_rgb)
    _set_run_language(run, "pt-BR")

def aplicar_estilo_corpo(run, negrito: bool = False) -> None:
    aplicar_estilo_texto(run, tamanho=11, negrito=negrito, fonte="Arial")

def aplicar_estilo_titulo(run, cor_rgb: Tuple[int, int, int] = _COR_PRETO_RGB) -> None:
    aplicar_estilo_texto(run, tamanho=12, negrito=True, fonte="Arial", cor_rgb=cor_rgb)

def adicionar_paragrafo_justificado(doc: Document, texto: str, negrito: bool = False):
    par = doc.add_paragraph() 
    par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY_LOW
    run = par.add_run(texto)
    aplicar_estilo_corpo(run, negrito=negrito)
    return par 

def adicionar_texto_centralizado(doc: Document, texto: str, tamanho_fonte: int = 12, negrito: bool = True):
    par = doc.add_paragraph() 
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = par.add_run(texto)
    if tamanho_fonte == 12 and negrito:
        aplicar_estilo_titulo(run)
    elif tamanho_fonte == 11:
        aplicar_estilo_corpo(run, negrito=negrito)
    else:
        aplicar_estilo_texto(run, tamanho_fonte, negrito, fonte="Arial")
    desabilitar_correcao_paragrafo(par)
    return par 

def adicionar_paragrafo_info_compacta(doc: Document, titulo: str, conteudo: str):
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
    paragrafo.paragraph_format.space_before = Pt(0)
    paragrafo.paragraph_format.space_after = Pt(0)
    paragrafo.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    paragrafo.paragraph_format.line_spacing = Pt(13)

def adicionar_titulo_secao(doc: Document, texto: str, nivel_heading: int = 1, cor_rgb: Tuple[int, int, int] = _COR_PRETO_RGB, aplicar_sombra: bool = False):
    try:
        par = doc.add_heading(level=nivel_heading)
    except Exception:
        par = doc.add_paragraph()
    run = par.add_run(texto)
    aplicar_estilo_titulo(run, cor_rgb=cor_rgb)
    if aplicar_sombra:
        aplicar_sombreamento_paragrafo(par, _COR_CINZA_SOMBRA_HEX)
        par.alignment = WD_ALIGN_PARAGRAPH.LEFT
    par.paragraph_format.space_after = Pt(6)
    run.font.all_caps = True
    desabilitar_correcao_paragrafo(par)
    return par

def adicionar_texto_contexto(doc: Document, contexto_tupla: Tuple[str, str]):
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
    par = doc.add_paragraph()
    par.paragraph_format.space_before = Pt(12)
    par.paragraph_format.space_after = Pt(6)
    par.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_id = par.add_run(f"{nc_id.strip()}")
    aplicar_estilo_titulo(run_id)
    
    texto_limpo = constatacao.split('\n')[0].strip()
    run_constatacao = par.add_run(f" - {texto_limpo}")
    
    aplicar_estilo_corpo(run_constatacao, negrito=False)
    desabilitar_correcao_paragrafo(par)
    return par

def adicionar_texto_esquerda(doc: Document, texto: str, tamanho_fonte: int = 11, negrito: bool = False, compacto: bool = False):
    par = doc.add_paragraph()
    if compacto:
        par.paragraph_format.space_before = Pt(0)
        par.paragraph_format.space_after = Pt(0)
    par.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = par.add_run(texto)
    if tamanho_fonte == 11:
        aplicar_estilo_corpo(run, negrito=negrito)
    else:
        aplicar_estilo_texto(run, tamanho_fonte, negrito)
    return par

def formatar_data_df(data_value) -> str:
    if pd.isna(data_value) or str(data_value).lower() == 'nan':
        return ""
    if isinstance(data_value, datetime):
        return data_value.strftime("%d/%m/%Y")
    if isinstance(data_value, str):
        try:
            return datetime.strptime(data_value.split()[0], "%Y-%m-%d").strftime("%d/%m/%Y")
        except ValueError:
            return data_value
    return str(data_value)

def ajustar_largura_colunas(caminho_planilha: str) -> None:
    try:
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
                        pass
                ws.column_dimensions[coluna_letra].width = max_length + 2
        wb.save(caminho_planilha)
    except Exception:
        pass

def arquivo_em_uso(caminho: str) -> bool:
    if not os.path.exists(caminho):
        return False
    try:
        os.rename(caminho, caminho)
        return False
    except PermissionError:
        return True

def aplicar_borda_paragrafo(paragraph) -> None:
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
    par = doc.add_paragraph()
    run = par.add_run(texto)
    aplicar_estilo_texto(run, tamanho=10, fonte="Arial", cor_rgb=(90, 90, 90))
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    aplicar_borda_paragrafo(par)

def processar_imagem_para_relatorio(caminho_imagem: str, largura_max: int = 1024, qualidade: int = 80) -> Optional[io.BytesIO]:
    if not os.path.exists(caminho_imagem):
        return None
    try:
        img = Image.open(caminho_imagem)
    except Exception as exc:
        print(f"Erro abrindo imagem '{caminho_imagem}': {exc}")
        return None
    if img.mode != "RGB":
        img = img.convert("RGB")
    if img.width > largura_max:
        proporcao = largura_max / float(img.width)
        altura_nova = int(float(img.height) * proporcao)
        img = img.resize((largura_max, altura_nova), Image.LANCZOS)
    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=qualidade, optimize=True)
    buffer.seek(0)
    return buffer

def adicionar_imagem(doc: Document, buffer_imagem: Optional[io.BytesIO], largura_in: Optional[Inches] = None, altura_in: Optional[Inches] = None) -> None:
    largura_final = largura_in if largura_in is not None else LARGURA_PADRAO_IN
    par = doc.add_paragraph()
    if buffer_imagem:
        try:
            if altura_in is not None:
                par.add_run().add_picture(buffer_imagem, width=largura_final, height=altura_in)
            else:
                par.add_run().add_picture(buffer_imagem, width=largura_final)
        except Exception as exc:
            par.add_run(f"Erro ao inserir imagem: {exc}")
    else:
        par.add_run("🚫 Imagem indisponível.")
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

def set_cell_border(cell, **kwargs) -> None:
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
    if not cell.paragraphs:
        cell.add_paragraph()
    par = cell.paragraphs[0]
    if len(par.runs) > 0 or par.text.strip() != "":
        try:
            par.clear()
        except Exception:
            cell._tc.get_or_add_tcPr()
            par = cell.paragraphs[0]
    run = par.add_run(texto)
    aplicar_estilo_texto(run, tamanho=10, fonte="Arial", cor_rgb=(90, 90, 90))
    par.alignment = WD_ALIGN_PARAGRAPH.LEFT

def adicionar_duas_imagens_lado_a_lado(doc: Document, fotos_dir: str, nome_foto1: str, legenda1: str, nome_foto2: Optional[str] = None, legenda2: Optional[str] = None, contexto_nc_tupla: Optional[Tuple[str, str, str]] = None) -> None:
    if contexto_nc_tupla:
        _, nc_id, constatacao = contexto_nc_tupla
        _adicionar_contexto_nc(doc, nc_id, constatacao)
    tabela = doc.add_table(rows=2, cols=2)
    tabela.alignment = WD_TABLE_ALIGNMENT.CENTER
    tabela.style = "Table Grid"
    if nome_foto2 is None:
        c_img = tabela.cell(0, 0).merge(tabela.cell(0, 1))
        c_img.width = LARGURA_PADRAO_IN
        c_leg = tabela.cell(1, 0).merge(tabela.cell(1, 1))
        c_leg.width = LARGURA_PADRAO_IN
        p = c_img.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        buf = processar_imagem_para_relatorio(os.path.join(fotos_dir, nome_foto1))
        if buf: p.add_run().add_picture(buf, width=LARGURA_PADRAO_IN)
        else: p.add_run(f"🚫 {nome_foto1}")
        adicionar_legenda_formatada_na_celula(c_leg, legenda1)
    else:
        largura_celula_img = LARGURA_IMAGEM_LADO_A_LADO
        largura_foto_interna = Inches(3.0) 
        
        for r in range(2):
            for c in range(2):
                tabela.cell(r, c).width = largura_celula_img

        p1 = tabela.cell(0, 0).paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        b1 = processar_imagem_para_relatorio(os.path.join(fotos_dir, nome_foto1))
        if b1: p1.add_run().add_picture(b1, width=largura_foto_interna, height=ALTURA_IMAGEM_LADO_A_LADO)
        else: p1.add_run(f"🚫 {nome_foto1}")

        p2 = tabela.cell(0, 1).paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        b2 = processar_imagem_para_relatorio(os.path.join(fotos_dir, nome_foto2))
        if b2: p2.add_run().add_picture(b2, width=largura_foto_interna, height=ALTURA_IMAGEM_LADO_A_LADO)
        else: p2.add_run(f"🚫 {nome_foto2}")

        adicionar_legenda_formatada_na_celula(tabela.cell(1, 0), legenda1)
        adicionar_legenda_formatada_na_celula(tabela.cell(1, 1), legenda2 or "")
        
    doc.add_paragraph().paragraph_format.space_after = Pt(12)