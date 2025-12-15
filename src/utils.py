from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import load_workbook
import os
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image
import io
from docx import Document
import pandas as pd
from docx.enum.table import WD_ALIGN_VERTICAL
import re

# --- Funções auxiliares de formatação: ---

def remover_espacamento_paragrafo(paragrafo):
    """Remove o espaçamento antes e depois do parágrafo e define espaçamento zero."""
    # Garante que o espaçamento antes e depois seja 0pt (para compactar)
    paragrafo_format = paragrafo.paragraph_format
    paragrafo_format.space_before = Pt(0)
    paragrafo_format.space_after = Pt(0)
    
def adicionar_quebra_linha_controlada(doc, altura_pt=18):
    """Adiciona uma linha vazia com espaçamento controlado para forçar o layout na Capa."""
    paragrafo = doc.add_paragraph()
    # Usa a função auxiliar para remover o espaçamento padrão
    remover_espacamento_paragrafo(paragrafo)
    
    # Adiciona um "run" vazio e define o tamanho da fonte para controlar a altura da linha
    run = paragrafo.add_run("")
    run.font.size = Pt(altura_pt)


def adicionar_paragrafo_justificado(doc, texto, tamanho_fonte=12):
    """Adiciona um parágrafo com texto justificado."""
    paragrafo = doc.add_paragraph(texto)
    paragrafo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    # Se precisar de espaçamento zero (compacto):
    # remover_espacamento_paragrafo(paragrafo)


def adicionar_texto_centralizado(doc, texto, tamanho_fonte=12):
    """Adiciona um parágrafo com texto centralizado (com negrito por padrão) e ESPAÇAMENTO CONTROLADO."""
    
    paragraph = doc.add_paragraph()
    
    # 🚨 CORREÇÃO: Remove o espaçamento padrão para uso em blocos compactos (ex: Capa)
    remover_espacamento_paragrafo(paragraph)
    
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run(texto)
    run.bold = True
    # run.font.size = Pt(tamanho_fonte) # Se o tamanho for importante


def adicionar_titulo_secao(doc, texto):
    """Adiciona um título de seção formatado."""
    secao = doc.add_paragraph()
    secao.add_run(texto).bold = True
    # Remover espaçamento padrão para melhor controle entre títulos e texto
    remover_espacamento_paragrafo(secao)

def adicionar_titulo_quadro(doc, texto, negrito=False, tamanho=11):
    """Adiciona um título de Quadro/Figura centralizado."""
    paragrafo = doc.add_paragraph()
    paragrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # Remover espaçamento padrão para melhor controle com a figura/quadro
    remover_espacamento_paragrafo(paragrafo)
    run = paragrafo.add_run(texto)
    run.bold = negrito
    # run.font.size = Pt(tamanho) # Opcional: manter o tamanho padrão do corpo do texto ou forçar um específico


# Função para ajustar a largura das colunas (Excel)
def ajustar_largura_colunas(caminho_planilha):
    """Ajusta a largura das colunas de todas as abas no arquivo Excel."""
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
    """Verifica se o arquivo está aberto/em uso."""
    try:
        # Tenta renomear o arquivo para si mesmo. Falha se estiver em uso.
        os.rename(caminho, caminho)
        return False
    except PermissionError:
        return True


def aplicar_estilo_texto(
    run, tamanho=12, negrito=False, fonte="Arial", cor_rgb=(0, 0, 0)
):
    """Aplica estilo completo em um 'run' do Word (fonte, tamanho, cor, negrito)."""
    run.font.name = fonte
    run._element.rPr.rFonts.set(qn("w:eastAsia"), fonte)
    run.font.size = Pt(tamanho)
    run.bold = negrito
    run.font.color.rgb = RGBColor(*cor_rgb)


def aplicar_borda_paragrafo(paragraph):
    """Aplica borda completa ao redor de um parágrafo (usado em legendas)."""
    p = paragraph._element
    pPr = p.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    for border_name in ("top", "left", "bottom", "right"):
        border = OxmlElement(f"w:{border_name}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "4") # Espessura
        border.set(qn("w:space"), "2") # Espaçamento
        border.set(qn("w:color"), "000000")
        borders.append(border)
    pPr.append(borders)


def adicionar_legenda_formatada(doc, texto):
    """Adiciona uma legenda formatada com estilo cinza e borda."""
    par = doc.add_paragraph()
    remover_espacamento_paragrafo(par) # Compactação
    run = par.add_run(texto)
    aplicar_estilo_texto(run, tamanho=10, fonte="Arial", cor_rgb=(90, 90, 90))
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    aplicar_borda_paragrafo(par)


def processar_imagem_para_relatorio(caminho_imagem, largura_max=1024, qualidade=80):
    """
    Otimiza (redimensiona e comprime) uma imagem e a retorna em um buffer de memória.
    """
    # Abre a imagem
    img = Image.open(caminho_imagem)
    # Converte para RGB se necessário (evita problemas com PNG/transparência)
    if img.mode != "RGB":
        img = img.convert("RGB")
    # Redimensiona mantendo proporção
    if img.width > largura_max:
        proporcao = largura_max / float(img.width)
        altura_nova = int(float(img.height) * proporcao)
        # Usando Image.Resampling.LANCZOS no Pillow 9.0+
        img = img.resize((largura_max, altura_nova), Image.Resampling.LANCZOS)
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
    remover_espacamento_paragrafo(paragrafo) # Compactação
    run = paragrafo.add_run(texto.upper())
    run.bold = True
    
    paragrafo.alignment = WD_ALIGN_PARAGRAPH.LEFT 
    
    adicionar_quebra_linha_controlada(doc, altura_pt=12) # Adiciona espaço controlado após o título
    

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
            remover_espacamento_paragrafo(par) # Compactação
            par.text = ''
            run = par.add_run(rotulo)
            run.bold = True
            
            par.alignment = WD_ALIGN_PARAGRAPH.LEFT 
            
            continue # Pula para a próxima linha de dados

        # --- Lógica de Dados Normais ---
        
        # Rótulo (Primeira Coluna) - SEMPRE em negrito na imagem
        par_rotulo = celula_rotulo.paragraphs[0] if celula_rotulo.paragraphs else celula_rotulo.add_paragraph()
        remover_espacamento_paragrafo(par_rotulo) # Compactação
        par_rotulo.text = ''
        run_rotulo = par_rotulo.add_run(rotulo)
        run_rotulo.bold = True
        par_rotulo.alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        # Valor (Segunda Coluna) - Negrito Condicional
        par_valor = celula_valor.paragraphs[0] if celula_valor.paragraphs else celula_valor.add_paragraph()
        remover_espacamento_paragrafo(par_valor) # Compactação
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
    remover_espacamento_paragrafo(paragrafo_img) # Compactação
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
    """
    # 1. Busca e prepara as legendas
    
    # 🚨 PONTO CRÍTICO: Tenta carregar a aba. Se falhar no nome com espaço, tenta sem.
    aba_nc = "Não-conformidades " 
    try:
        df_legendas = pd.read_excel(caminho_planilha_legendas, sheet_name=aba_nc)
    except ValueError:
        try:
            df_legendas = pd.read_excel(caminho_planilha_legendas, sheet_name=aba_nc.strip())
        except Exception:
             adicionar_paragrafo_justificado(doc, f"ERRO: Não foi possível encontrar a aba de legendas ('{aba_nc}' ou '{aba_nc.strip()}').")
             return

    # Filtra a linha correta pelo ID
    # Limpa espaços nas colunas para evitar KeyError
    df_legendas.columns = df_legendas.columns.str.strip() 
    
    # 🚨 CORREÇÃO DE TIPAGEM/FORMATO: Garante que ambos são strings limpas para a comparação
    id_fisc_limpo = str(id_fiscalizacao).strip()
    
    # 2. Filtragem robusta: Converte a coluna para string limpa antes de comparar
    linha_legenda = df_legendas[
        df_legendas['ID da Fiscalização'].astype(str).str.strip() == id_fisc_limpo
    ]
     
    if linha_legenda.empty:
        adicionar_paragrafo_justificado(doc, f"AVISO: Legendas não encontradas na planilha para o ID: {id_fiscalizacao}. As fotos não serão legendadas.")
        legendas = [] 
    else:
        # Extrai o texto da coluna 'Legenda da Foto' e separa
        texto_legendas = str(linha_legenda['Legenda da Foto'].iloc[0]) 
        legendas = [l.strip() for l in texto_legendas.split(';') if l.strip()]

    # 3. Busca e prepara os caminhos das fotos
    # Obtém a lista de arquivos de imagem na pasta especificada
    try:
        arquivos_fotos = sorted([
            os.path.join(caminho_base_fotos, f)
            for f in os.listdir(caminho_base_fotos)
            if f.lower().endswith(('.png', '.jpg', '.jpeg'))
        ])
    except FileNotFoundError:
        adicionar_paragrafo_justificado(doc, f"ERRO INTERNO: O caminho de fotos '{caminho_base_fotos}' não foi encontrado.")
        return

    if not arquivos_fotos:
        adicionar_paragrafo_justificado(doc, f"AVISO: A subpasta procurada está vazia, Seu Apêndice não terá fotos.")
        return

    # 4. Cria a estrutura da tabela (layout de 2x2)

    num_fotos = len(arquivos_fotos)
    # Linhas: 1 linha de foto + 1 linha de legenda para cada 2 fotos
    num_linhas_tabela = ((num_fotos + 1) // 2) * 2 

    tabela = doc.add_table(rows=num_linhas_tabela, cols=2)
    tabela.autofit = False
    tabela.style = 'Table Grid'

    # Define a largura das colunas
    largura_coluna_cm = 8.5 
    tabela.columns[0].width = Cm(largura_coluna_cm)
    tabela.columns[1].width = Cm(largura_coluna_cm)

    # 5. Popula a tabela
    for i in range(num_linhas_tabela // 2): # Itera sobre os "pares" de linhas (Foto + Legenda)
        linha_foto_idx = i * 2
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

                celula_legenda.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                par_legenda = celula_legenda.paragraphs[0] if celula_legenda.paragraphs else celula_legenda.add_paragraph()
                remover_espacamento_paragrafo(par_legenda) # Compactação
                par_legenda.alignment = WD_ALIGN_PARAGRAPH.CENTER

                # Pega a legenda correta
                texto_legenda = legendas[foto_idx] if foto_idx < len(legendas) else f"Legenda {foto_idx + 1} não fornecida na planilha."

                # Adiciona e formata o texto da legenda
                run_legenda = par_legenda.add_run(texto_legenda)
                aplicar_estilo_texto(run_legenda, tamanho=10, fonte="Arial", cor_rgb=(90, 90, 90))

            else:
                # Preenche células vazias
                celula_foto_vazia = tabela.cell(linha_foto_idx, j)
                celula_legenda_vazia = tabela.cell(linha_legenda_idx, j)

                par_foto_vazia = celula_foto_vazia.add_paragraph("")
                remover_espacamento_paragrafo(par_foto_vazia)
                par_foto_vazia.alignment = WD_ALIGN_PARAGRAPH.CENTER

                par_legenda_vazia = celula_legenda_vazia.add_paragraph("")
                remover_espacamento_paragrafo(par_legenda_vazia)
                par_legenda_vazia.alignment = WD_ALIGN_PARAGRAPH.CENTER

# --- FUNÇÃO DE GERAÇÃO DA TABELA DE ABREVIATURAS (CORRIGIDA) ---

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

    # 🚨 AJUSTE DE LARGURA: Acentuando a diferença
    largura_sigla_cm = 2.5 # Ajustado
    largura_definicao_cm = 14.5 # Ajustado
    tabela.columns[0].width = Cm(largura_sigla_cm)
    tabela.columns[1].width = Cm(largura_definicao_cm)
    
    # 1. Cabeçalho (Negrito e Centralizado)
    header_cells = tabela.rows[0].cells
    
    for cell in header_cells:
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Célula SIGLA (Header)
    par_sigla = header_cells[0].paragraphs[0]
    remover_espacamento_paragrafo(par_sigla) # Compactação
    par_sigla.text = ""
    run_sigla = par_sigla.add_run("SIGLA")
    run_sigla.bold = True
    par_sigla.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Célula DEFINIÇÃO (Header)
    par_def = header_cells[1].paragraphs[0]
    remover_espacamento_paragrafo(par_def) # Compactação
    par_def.text = ""
    run_def = par_def.add_run("DEFINIÇÃO")
    run_def.bold = True
    par_def.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 2. Preenche as linhas de dados
    for i, row in df_abreviaturas.iterrows():
        cells = tabela.rows[i + 1].cells
        
        for cell in cells:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER 

        # Sigla (Centralizada e Negrito)
        par_sigla_data = cells[0].paragraphs[0]
        remover_espacamento_paragrafo(par_sigla_data) # Compactação
        par_sigla_data.text = ""
        run_sigla_data = par_sigla_data.add_run(str(row['Sigla']))
        run_sigla_data.bold = True 
        par_sigla_data.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Definição (Esquerda)
        par_def_data = cells[1].paragraphs[0]
        remover_espacamento_paragrafo(par_def_data) # Compactação
        par_def_data.text = ""
        par_def_data.add_run(str(row['Definição'])) 
        par_def_data.alignment = WD_ALIGN_PARAGRAPH.LEFT

# --- FUNÇÕES AUXILIARES PARA A CAPA ---

def parear_responsaveis_e_matriculas(nomes_str, matriculas_str, cargo_fixo="Analista de Regulação, matrícula nº"):
    """
    Divide as strings de nomes e matrículas pelo caractere ';' e as emparelha.
    Retorna uma lista de dicionários com 'nome' e 'info_completa'.
    """
    if not nomes_str or not matriculas_str:
        return []

    # Divide e remove espaços em branco (strip)
    nomes = [n.strip() for n in nomes_str.split(';') if n.strip()]
    matriculas = [m.strip() for m in matriculas_str.split(';') if m.strip()]
    
    profissionais = []
    
    # Emparelha nomes e matrículas
    for nome, matricula in zip(nomes, matriculas):
        info_completa = f"{cargo_fixo} {matricula}"
        profissionais.append({
            'nome': nome,
            'info_completa': info_completa
        })
        
    return profissionais

def formatar_data_capa(data_str):
    """
    Converte a data 'DD/MM/AAAA' para 'Mês, AAAA'.
    """
    if not data_str:
        return ""
    
    try:
        meses = {
            '01': 'Janeiro', '02': 'Fevereiro', '03': 'Março', '04': 'Abril',
            '05': 'Maio', '06': 'Junho', '07': 'Julho', '08': 'Agosto',
            '09': 'Setembro', '10': 'Outubro', '11': 'Novembro', '12': 'Dezembro'
        }
        
        # O str(data_str) garante que funciona com pd.Timestamp ou string
        partes = str(data_str).split('/')
        
        # Tentativa de converter se for um formato de data/hora completo do Pandas
        if len(partes) < 3 and '-' in str(data_str):
            # Tenta tratar como timestamp (ex: 2025-11-12 00:00:00)
            data_obj = pd.to_datetime(data_str, errors='coerce')
            if pd.isna(data_obj):
                    return data_str # Retorna original se falhar
            
            mes_numero = data_obj.strftime('%m')
            ano = data_obj.strftime('%Y')
        else:
            # Trata como string 'DD/MM/AAAA'
            if len(partes) < 3:
                return data_str 

            mes_numero = partes[1]
            ano = partes[2]
        
        mes_extenso = meses.get(mes_numero, mes_numero)
        
        return f"{mes_extenso}, {ano}"

    except Exception:
        return data_str
    
def parear_e_formatar_assinaturas(nomes_str, matriculas_str, cargo_fixo="Analista de Regulação"):
    """
    Divide as strings de nomes e matrículas pelo caractere ';' e as emparelha.
    Retorna uma lista de tuplas formatadas: (nome, cargo, matricula_formatada).
    """
    nomes = [n.strip() for n in str(nomes_str).split(';') if n.strip()]
    matriculas = [m.strip() for m in str(matriculas_str).split(';') if m.strip()]
    
    assinaturas_formatadas = []
    
    for i, nome in enumerate(nomes):
        matricula_raw = matriculas[i] if i < len(matriculas) else ""
        
        if matricula_raw:
            matricula_info = f"Matrícula nº {matricula_raw}"
        else:
            matricula_info = "" # Nenhuma matrícula
            
        assinaturas_formatadas.append((nome, cargo_fixo, matricula_info))
        
    return assinaturas_formatadas    

# --- FUNÇÕES AUXILIARES DE DATA E EXTRAÇÃO DE CIDADE ---

# --- FUNÇÃO AUXILIAR DE EXTRAÇÃO DE CIDADE ---

def extrair_cidade(terminal_str):
    """Extrai apenas o nome da cidade de uma string de terminal (ex: 'Terminal de Recife (TIP)')."""
    if not isinstance(terminal_str, str):
        return ""
    # Remove prefixos comuns e o que estiver entre parênteses
    cidade = re.sub(r'terminal d[eo]\s*', '', terminal_str, flags=re.IGNORECASE)
    # Mantém o (TIP) se for o caso, removendo apenas outros parênteses
    cidade = re.sub(r'\(TIP\)', ' (TIP)', cidade, flags=re.IGNORECASE).strip() 
    cidade = re.sub(r'\s*\(.*\)', '', cidade).strip()
    return cidade.capitalize() # Capitaliza para um formato mais limpo

# --- FUNÇÃO AUXILIAR DE FILTRAGEM ---

def padronizar_processo(id_fisc):
    """
    Converte o ID de Fiscalização (ex: 'CTR-04-2025') para o formato da planilha (ex: 'CTR 04/2025').
    O filtro é feito pela coluna 'PROCESSO' do Excel.
    """
    if not isinstance(id_fisc, str):
        return ""
        
    # Remove hífens e converte para maiúsculas
    processo_formatado = id_fisc.upper().replace('-', ' ')
    
    # Adiciona a barra (/) no formato XX XX/XXXX
    partes = processo_formatado.split()
    if len(partes) == 3:
        # Ex: ['CTR', '04', '2025'] -> 'CTR 04/2025'
        return f"{partes[0]} {partes[1]}/{partes[2]}"
        
    return processo_formatado