## 📘 Gerador de Relatórios de Monitoramento (Automático)

Automação completa desenvolvida em Python para gerar relatórios padronizados em **DOCX e PDF**, cruzando informações de planilhas Excel com evidências fotográficas.

Relatórios profissionais produzidos em poucos segundos.

## 📄 Descrição

Este projeto automatiza a geração de relatórios de monitoramento, cruzando:

- **Planilha de Monitoramento (Diário de Campo)**

- **Planilha Oficial Base (Levantamento de NCs)**

O sistema reconcilia dados, organiza fotos, gera textos, estrutura um relatório completo e exporta tudo em Word e PDF.

## 🏷️ Badges
<p align="left"> <img src="https://img.shields.io/badge/Python-3.10%2B-blue" /> <img
src="https://img.shields.io/badge/Status-Em%20Desenvolvimento-yellow" /> <img
src="https://img.shields.io/badge/Automação-Relatórios-success" /> <img
src="https://img.shields.io/badge/Geração-DOCX%20%2F%20PDF-orange" /> </p>

## 🚀 Funcionalidades

- **Leitura Inteligente de Excel:** lê dados de múltiplas abas (Monitoramento, Não-conformidades, Processos, BASE).  
- **Cruzamento de Dados (Match):** conecta automaticamente o que foi visto em campo com o item oficial da base, mesmo que o texto não seja idêntico.  
- **Agrupamento de Famílias:** agrupa automaticamente itens relacionados (ex: 02.1, 02.2, 02.3).  
- **Geração de Documento Word:** cria um `.docx` completo com Capa, Sumário, Introdução, Tabelas e Assinaturas.  
- **Anexo Fotográfico Automático:** busca fotos na pasta, organiza em pares e legenda automaticamente.  
- **Conversão para PDF:** gera a versão final pronta para envio.

## 📝 Como Preencher as Planilhas
**1. Planilha de Monitoramento (planilha_monitoramento.xlsx)**

**Aba Monitoramento**

- **ID da Fiscalização:** número único do relatório

- **Relatório Gerado:** deixe FALSE ou vazio para gerar

**Aba Não-conformidades**

- **ID da Fiscalização:** mesmo número da aba anterior

- **Terminal: ex.:** "Terminal de Caruaru"

- **Legenda da Foto:** ⚠️ campo chave! usado para localizar o item oficial

- **Informação SOCICAM:** resposta da empresa (opcional)

## Preencher filtros (no terminal)

- Ano → ex.: 2025

- Processo → ex.: CTR 04/2025

- Nº Monitoramento → ex.: 1

- Pasta do Contrato → ex.: CTR-01-2025

- Ciclo → ex.: M0

## 💡 Dicas Importantes

- **Nome das fotos:** devem incluir o ID (ex.: CAR 2025_13.jpg)

- Funciona com variações (CAR 2025_13 (1).jpg)

- **Múltiplas legendas**: separe por :

- **Erro “ID NÃO ENCONTRADO”:**

- Ajuste palavras-chave da Legenda da Foto para se aproximar da BASE

 










