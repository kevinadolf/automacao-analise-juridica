# Extrator Inteligente de Processos do TCE

## 📖 Sobre o Projeto

Este projeto é um script em Python desenvolvido para automatizar a extração de informações chave de documentos processuais do Tribunal de Contas, como Acórdãos e Decisões Monocráticas. Ele analisa arquivos em formato `.pdf` e `.docx` para identificar e extrair o valor monetário principal do recurso fiscalizado, além de metadados importantes como Número do Processo, Número do Acórdão e Natureza.

## ✨ Funcionalidades Principais

- **Processamento em Lote:** Varre uma estrutura de pastas e processa múltiplos documentos de forma automática.
- **Suporte a Múltiplos Formatos:** Extrai texto de forma robusta de arquivos `.pdf` (usando PyMuPDF) e `.docx`.
- **Extração Inteligente de Valores:** Utiliza uma combinação de Regex e uma **hierarquia de contexto** para decidir qual é o valor monetário mais relevante, diferenciando o valor principal do objeto de sanções (multas) e outros valores secundários.
- **Extração de Metadados:** Identifica e extrai automaticamente o Nº do Processo, Nº do Acórdão e a Natureza do documento.
- **Otimização de Performance:** Possui um filtro que identifica processos arquivados por inadmissibilidade e pula a análise de valores, economizando tempo de processamento.
- **Exportação Estruturada:** Salva todos os resultados em uma única planilha Excel (`.xlsx`), com formatação condicional para destacar visualmente os processos arquivados.

## 🛠️ Tecnologias Utilizadas

- Python
- Pandas
- PyMuPDF (fitz)
- python-docx

## 🚀 Como Usar

1.  Clone este repositório.
2.  Instale as dependências: `pip install pandas PyMuPDF python-docx openpyxl`
3.  Crie uma pasta raiz (ex: `arquivos_para_teste`) e, dentro dela, crie subpastas para cada processo.
4.  Coloque os arquivos `.pdf` ou `.docx` dentro de suas respectivas subpastas.
5.  No script, ajuste a variável `PASTA_RAIZ_PROCESSOS` para o nome da sua pasta raiz.
6.  Execute o script: `python nome_do_seu_script.py`
7.  A planilha Excel com os resultados será gerada no diretório principal.
