from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import (
    adicionar_paragrafo_justificado,
    adicionar_titulo_secao,
    adicionar_paragrafo_info_compacta,
    aplicar_estilo_corpo,
    aplicar_estilo_titulo,
    encontrar_dados_na_base
)
from typing import Any, Dict
import pandas as pd

def _inserir_texto_nc(doc, nc_titulo_identificador: str, descricao: str):
    paragrafo_nc = doc.add_paragraph()
    run_titulo = paragrafo_nc.add_run(f"Não Conformidade {nc_titulo_identificador}")
    aplicar_estilo_corpo(run_titulo, negrito=True)
    run_titulo.underline = True
    
    run_traco = paragrafo_nc.add_run(" - ")
    aplicar_estilo_corpo(run_traco)
    
    run_desc = paragrafo_nc.add_run(str(descricao))
    aplicar_estilo_corpo(run_desc)
    
    paragrafo_nc.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY_LOW

def _inserir_linhas_info(doc, info_socicam: str, constatacao_monit: str, analise_arpe: str):
    if str(info_socicam).lower() in ['nan', 'none', '']: info_socicam = " "
    if str(constatacao_monit).lower() in ['nan', 'none', '']: constatacao_monit = " "
    if str(analise_arpe).lower() in ['nan', 'none', '']: analise_arpe = " "
    
    adicionar_paragrafo_info_compacta(doc, "Informação da SOCICAM: ", str(info_socicam))
    adicionar_paragrafo_info_compacta(doc, "Constatação: ", str(constatacao_monit)) 
    adicionar_paragrafo_info_compacta(doc, "Análise da ARPE: ", str(analise_arpe))

def _safe_split(texto: Any) -> list:
    content = str(texto).strip()
    if not content or content.lower() == "nan": return []
    return [item.strip() for item in content.split(";") if item.strip()]

def gerar_secao_nao_conformidades_constatadas(
    doc, 
    row: dict, 
    nao_conformidades_df: pd.DataFrame, 
    FOTOS_DIR: str,
    processo_info: Dict[str, str],
    df_base_nc: pd.DataFrame,
    ano_user: str,
    proc_user: str,
    monit_user: str
):
    id_fiscalizacao = row["ID da Fiscalização"]
    processo_ctr = processo_info.get("Processo CTR Nº", "XX/XXXX")
    cartas_raw = str(processo_info.get("Carta SAP/PER/ARPE Nº", "")).strip()
    texto_cartas = f"constante da Carta SAP/PER/ARPE N° {cartas_raw}, " if (cartas_raw and cartas_raw.lower() != 'nan') else ""

    nc_fiscalizacao = nao_conformidades_df[
        nao_conformidades_df["ID da Fiscalização"] == id_fiscalizacao
    ].copy()

    adicionar_titulo_secao(doc, "3. RESULTADO DAS VISTORIAS DAS NÃO CONFORMIDADES PENDENTES")
    adicionar_paragrafo_justificado(
        doc,
        (
            "Estão registrados para cada Terminal Rodoviário os resultados da verificação pela Arpe das ações "
            f"desenvolvidas pela SOCICAM, {texto_cartas}"
            "para solucionar as Não Conformidades ainda pendentes apresentadas no Relatório de Fiscalização Técnico-"
            f"Operacional ARPE/CTR nº {processo_ctr}."
        ),
    )

    num_terminal = 1
    for terminal, grupo_terminal in nc_fiscalizacao.groupby("Terminal"):
        par_terminal = doc.add_paragraph()
        run_terminal = par_terminal.add_run(f"3.{num_terminal} - {terminal.upper()}")
        aplicar_estilo_titulo(run_terminal)
        par_terminal.paragraph_format.space_before = Pt(12)

        ncs_processadas = set()

        for _, nc_data in grupo_terminal.iterrows():
            # Prioriza: Legenda -> Constatação
            raw_key = str(nc_data.get("Legenda da Foto", "")).strip()
            if not raw_key or raw_key.lower() == "nan":
                raw_key = str(nc_data.get("Constatação", "")).strip()
            
            constatacoes_monit = _safe_split(raw_key)
            
            raw_info = str(nc_data.get("Informação SOCICAM", "")).strip()
            infos_socicam = _safe_split(raw_info)
            raw_analise = str(nc_data.get("Análise da Arpe", "")).strip()
            analises_arpe = _safe_split(raw_analise)

            max_len = max(len(constatacoes_monit), len(infos_socicam), len(analises_arpe))
            if len(constatacoes_monit) < max_len: constatacoes_monit.extend([""] * (max_len - len(constatacoes_monit)))
            if len(infos_socicam) < max_len: infos_socicam.extend(["N/A"] * (max_len - len(infos_socicam)))
            if len(analises_arpe) < max_len: analises_arpe.extend(["N/A"] * (max_len - len(analises_arpe)))

            for i in range(max_len):
                texto_busca = constatacoes_monit[i]
                if not texto_busca: continue

                dados_base = encontrar_dados_na_base(
                    texto_busca, 
                    df_base_nc, 
                    ano_user, 
                    proc_user, 
                    monit_user,
                    terminal_user=str(terminal)
                )
                
                if dados_base["id"] in ncs_processadas and dados_base["id"] != "ID_NAO_ENCONTRADO":
                    continue
                
                if dados_base["id"] != "ID_NAO_ENCONTRADO":
                    ncs_processadas.add(dados_base["id"])

                _inserir_texto_nc(doc, dados_base["id"], dados_base["descricao"])
                _inserir_linhas_info(doc, infos_socicam[i], texto_busca, analises_arpe[i])
                doc.add_paragraph()
        num_terminal += 1