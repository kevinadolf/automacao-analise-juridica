import os
import re
import sys
import time
import math
import fitz  # PyMuPDF
from docx import Document
import pandas as pd
from collections import defaultdict
import json

# --- Configurações Essenciais ---
PASTA_RAIZ_PROCESSOS = 'arquivos_para_teste'
MAX_PARAGRAPH_ETAPA_2 = 400

# --- CONFIGURAÇÃO DOS MODELOS LLM OPEN SOURCE PARA TESTE ---
# Instrução:
# 1. Baixe os modelos em formato GGUF do Hugging Face.
# 2. Crie uma pasta 'models' no seu projeto e coloque os arquivos .gguf nela.
# 3. Atualize os caminhos abaixo para corresponder aos nomes dos seus arquivos.
MODELOS_PARA_TESTAR = [
    {
        "nome": "Llama3-8B-Instruct",
        "path": "./models/Meta-Llama-3-8B-Instruct.Q4_K_M.gguf", # Exemplo de caminho
    },
    {
        "nome": "Mistral-7B-Instruct",
        "path": "./models/Mistral-7B-Instruct-v0.2.Q4_K_M.gguf", # Exemplo de caminho
    }
]

# --- Padrões de Extração e Palavras-Chave (Simplificados para focar no LLM) ---
PADROES_VALOR_REFINADOS = [
    r'R\$\s*(?P<value>\d{1,3}(?:[._]\d{3})*(?:,\d{2})?)',
    r'R\$\s*(?P<number>\d+(?:[.,]\d{1,2})?)\s*(?P<unit>milh[oõ]es|mil|bilh[oõ]es|bi|tri)(?:\s+de\s+reais)?',
]
RE_PROCESSO_PDF = re.compile(r"PROCESSO(?:.*?N[º°]?)?\s*[:\s]*([\w\d.-]+/\d{2,4})", re.IGNORECASE)
RE_NATUREZA_PDF = re.compile(r"NATUREZA:\s*(.+)", re.IGNORECASE)
RE_ACORDAO_PDF = re.compile(r"AC[OÓ]RD[AÃ]O Nº\s*([\w\d./-]+(?:-PLEN(?:V)?)?)", re.IGNORECASE)

# --- Funções ---

def print_progress_bar(iteration, total, prefix='', suffix='', length=50, fill='█'):
    percent = ("{0:.1f}").format(100 * (iteration / float(total)))
    filled_length = int(length * iteration // total)
    bar = fill * filled_length + '-' * (length - filled_length)
    sys.stdout.write(f'\r{prefix} |{bar}| {percent}% {suffix}')
    sys.stdout.flush()
    if iteration == total: sys.stdout.write('\n')

def converter_valor_para_numero_refinado(valor_str_original):
    if not isinstance(valor_str_original, str): return None
    valor_str = valor_str_original.lower()
    valor_str = re.sub(r"^(r\$\s*|valor\s+de\s*r\$\s*|montante\s+de\s*r\$\s*)", "", valor_str).strip()
    valor_str = re.sub(r"(\s*\((?:.*?)\)).*$", "", valor_str).strip()
    multiplicador = 1
    if 'tri' in valor_str: multiplicador = 1e12; valor_str = re.sub(r'tri(?:lh[oõ]es)?', '', valor_str).strip()
    elif 'bilh' in valor_str or ' bi' in valor_str: multiplicador = 1e9; valor_str = re.sub(r'bilh[oõ]es|bi', '', valor_str).strip()
    elif 'milh' in valor_str: multiplicador = 1e6; valor_str = re.sub(r'milh[oõ]es', '', valor_str).strip()
    elif 'mil' in valor_str: multiplicador = 1e3; valor_str = re.sub(r'mil', '', valor_str).strip()
    valor_str = re.sub(r'\s+', '', valor_str)
    if re.search(r',\d{1,2}$', valor_str): valor_str = valor_str.replace('.', '').replace(',', '.')
    else: valor_str = valor_str.replace(',', '')
    if '.' in valor_str:
        parts = valor_str.split('.'); valor_str = "".join(parts[:-1]) + "." + parts[-1] if len(parts[-1]) <= 2 and len(parts) > 1 else "".join(parts)
    valor_str = re.sub(r'[^\d\.]', '', valor_str)
    if valor_str.count('.') > 1: valor_str = valor_str.replace('.', '', valor_str.count('.') - 1)
    try: return float(valor_str) * multiplicador
    except (ValueError, TypeError): return None

def extrair_metadados_pdf(caminho_pdf):
    numero_processo_pdf, natureza, numero_acordao = "NÃO ENCONTRADO", "NÃO ESPECIFICADO", "NÃO ENCONTRADO"
    try:
        with fitz.open(caminho_pdf) as pdf_doc:
            if len(pdf_doc) > 0:
                texto_primeira_pagina = pdf_doc[0].get_text("text")
                match_acordao = RE_ACORDAO_PDF.search(texto_primeira_pagina)
                if match_acordao: numero_acordao = match_acordao.group(1).strip()
                match_processo = RE_PROCESSO_PDF.search(texto_primeira_pagina)
                if match_processo: numero_processo_pdf = match_processo.group(1).strip()
                match_natureza_direta = RE_NATUREZA_PDF.search(texto_primeira_pagina)
                if match_natureza_direta: natureza = re.split(r'\s+INTERESSADO:', match_natureza_direta.group(1).strip().upper(), 1)[0].strip()
    except Exception as e: print(f"  -> Erro ao extrair metadados: {e}")
    if numero_acordao != "NÃO ENCONTRADO" and natureza == "NÃO ESPECIFICADO": natureza = "ACÓRDÃO"
    return {"numero_processo_pdf": numero_processo_pdf, "natureza": natureza, "numero_acordao": numero_acordao}

def obter_texto_documento(caminho_arquivo):
    paragrafos = []
    try:
        if caminho_arquivo.lower().endswith('.docx'):
            with open(caminho_arquivo, "rb") as docx_file:
                doc = Document(docx_file); [paragrafos.append(p.text) for p in doc.paragraphs]
        elif caminho_arquivo.lower().endswith('.pdf'):
            with fitz.open(caminho_arquivo) as pdf_doc:
                for pagina in pdf_doc: paragrafos.extend(p[4].replace('\r\n', ' ').replace('\r', ' ') for p in sorted(pagina.get_text("blocks"), key=lambda b: (b[1], b[0])) if p[6] == 0)
    except Exception as e: print(f"  -> Erro ao ler documento {os.path.basename(caminho_arquivo)}: {e}"); return None
    return paragrafos

def verificar_admissibilidade_e_arquivamento(lista_de_paragrafos):
    if not lista_de_paragrafos: return "Indeterminado"
    texto_final = " ".join(lista_de_paragrafos[-20:]).upper()
    flag_nao_conhecimento = "NÃO CONHECIMENTO" in texto_final
    flag_admissibilidade = "ADMISSIBILIDADE" in texto_final
    flag_arquivamento = "ARQUIVAMENTO" in texto_final
    if flag_nao_conhecimento and flag_admissibilidade and flag_arquivamento: return "Sim"
    return "Não"

def encontrar_valores_candidatos(lista_de_paragrafos):
    candidatos = []
    if not lista_de_paragrafos: return candidatos
    for i, linha_texto in enumerate(lista_de_paragrafos):
        if i >= MAX_PARAGRAPH_ETAPA_2: break
        for padrao in PADROES_VALOR_REFINADOS:
            for match in re.finditer(padrao, linha_texto, re.IGNORECASE):
                valor_num = converter_valor_para_numero_refinado(match.group(0))
                if valor_num and valor_num > 0:
                    candidatos.append({
                        "valor_str": match.group(0),
                        "valor_num": valor_num,
                        "contexto": linha_texto.strip()
                    })
    return list({(v['valor_str'], v['contexto']): v for v in candidatos}.values())

def chamar_llm_local(prompt, modelo_config):
    """funcao para chamar um LLM local usando llama-cpp-python."""
    try:
        from llama_cpp import Llama
        
        if not os.path.exists(modelo_config["path"]):
            return f'{{"erro": "Arquivo do modelo LLM não encontrado em: {modelo_config["path"]}"}}'
            
        llm = Llama(
            model_path=modelo_config["path"],
            n_ctx=4096,
            n_threads=4, # ajuste conforme CPU
            n_gpu_layers=0, # forca o uso da CPU
            verbose=False
        )
        output = llm(prompt, max_tokens=256, temperature=0.1, stop=["}", "}\n", "\n\n"])
        # temperature = se refere ao parametro que controla a aleatoriedade do texto gerado, ou seja,
        # baixa temp signifca um output mais previsível/determinístico, high temp significa mais aleatoriedade/criatividade
        
        resposta_texto = output['choices'][0]['text']
        # garante que o JSON seja fechado corretamente, pq o "}" pode ser um stop token
        if not resposta_texto.strip().endswith('}'):
            resposta_texto += "}"
            
        return resposta_texto
    except Exception as e:
        return f'{{"erro": "Falha ao rodar modelo local {modelo_config["nome"]}: {e}"}}'

#MELHOR PROMPT EXISTENTE, NAO TEM JEITO
def criar_prompt_desambiguacao(candidatos):
    prompt = """
Você é um assistente especialista em análise de documentos do Tribunal de Contas (TCE). 
Sua especialidade é identificar a materialidade das irregularidades em processos, como o valor de um contrato, uma licitação, uma multa aplicada ou um dano ao erário consolidado.

Analise os seguintes 'Candidatos a Valor Principal', extraídos de um documento do TCE. Cada candidato inclui o valor e o parágrafo onde ele foi encontrado.

### CANDIDATOS PARA ANÁLISE:
"""
    # Limita o número de candidatos para não sobrecarregar o prompt
    for i, candidato in enumerate(candidatos[:5]): 
        contexto_limpo = candidato['contexto'].replace('\n', ' ').strip()
        prompt += f"\n{i+1}. Valor: \"{candidato['valor_str']}\"\n   Contexto: \"...{contexto_limpo}...\"\n"
    
    prompt += """
### TAREFA:
Com base nos dados acima, retorne APENAS o objeto JSON com sua análise. Não inclua nenhuma outra palavra ou explicação fora do JSON.

Formato de saída obrigatório:
{"valor_principal_escolhido": "escreva aqui o valor exato que você escolheu", "justificativa": "explique brevemente o motivo da sua escolha baseado nas regras e no contexto", "tipo_de_valor": "classifique o valor como 'Valor do Contrato', 'Multa Aplicada', 'Dano ao Erário', 'Valor de Devolução' ou 'Outro'"}
"""
    return prompt

# e vai tomando prompt, toma
def criar_prompt_resumo(paragrafos, metadados_documento):
    # Pega o começo e o fim do documento para ter o contexto completo
    texto_para_resumir = " ".join(paragrafos[:20] + paragrafos[-30:]) # Aumentado um pouco para mais contexto
    # Limita o tamanho do texto para não exceder limites e otimizar a chamada
    texto_para_resumir = texto_para_resumir[:4000] 
    
    # Adiciona metadados já extraídos para ajudar o LLM a focar
    contexto_extra = f"Número do Processo: {metadados_documento.get('numero_processo_pdf', 'N/A')}, Número do Acórdão: {metadados_documento.get('numero_acordao', 'N/A')}."

    prompt = f"""
Você é um assistente que resume documentos jurídicos do Tribunal de Contas de forma clara e objetiva.

### DADOS DO PROCESSO:
{contexto_extra}

### TRECHOS DO DOCUMENTO PARA ANÁLISE:
\"...{texto_para_resumir}...\"

### TAREFA:
Com base nos dados e no texto acima, gere um resumo conciso de, no máximo, duas frases.
O resumo deve OBRIGATORIAMENTe mencionar o objeto principal em análise e a decisão final, **incluindo o valor monetário principal associado (seja o valor do contrato, da multa, etc.)**.

Exemplo de um bom resumo:
"Análise de Representação sobre o Contrato nº 123/2023 para obras de saneamento, com decisão pela aplicação de multa no valor de R$ 50.000,00 por superfaturamento."

### RESUMO CONCISO:
"""
    return prompt


def processar_documentos(pasta_raiz):
    resultados_finais = {}
    subpastas = [d for d in os.listdir(pasta_raiz) if os.path.isdir(os.path.join(pasta_raiz, d))]
    
    print_progress_bar(0, len(subpastas), prefix='Progresso:', suffix='Completo', length=40)
    for i, nome_subpasta in enumerate(subpastas):
        caminho_subpasta = os.path.join(pasta_raiz, nome_subpasta)
        
        documento_encontrado_path, nome_arquivo_processado = None, "Nenhum Documento Encontrado"
        for ext in ['.pdf', '.docx', '.doc']:
            for arq in sorted(os.listdir(caminho_subpasta)):
                if arq.lower().endswith(ext) and not arq.startswith('~$'):
                    documento_encontrado_path, nome_arquivo_processado = os.path.join(caminho_subpasta, arq), arq
                    break
            if documento_encontrado_path: break
            
        metadados = {
            "nome_subpasta_original": nome_subpasta, "nome_arquivo_original": nome_arquivo_processado,
            "numero_processo_pdf": "NÃO ENCONTRADO", "natureza": "NÃO ESPECIFICADO", 
            "numero_acordao": "NÃO ENCONTRADO", "status_admissibilidade": "Indeterminado",
            "resumo_llm": "N/A", "valor_final_llm": "N/A", "justificativa_llm": "N/A"
        }
        metadados.update({f"resposta_{m['nome']}": "Não processado" for m in MODELOS_PARA_TESTAR})
        
        if not documento_encontrado_path:
            resultados_finais[nome_subpasta] = {"metadados": metadados}
            print_progress_bar(i + 1, len(subpastas), prefix='Progresso:', suffix=f'({nome_subpasta} - Sem Doc)', length=40)
            continue

        if documento_encontrado_path.lower().endswith('.pdf'): metadados.update(extrair_metadados_pdf(documento_encontrado_path))
        
        lista_de_paragrafos = obter_texto_documento(documento_encontrado_path)
        if lista_de_paragrafos is None:
            resultados_finais[nome_subpasta] = {"metadados": metadados, "criterio_usado": "erro_leitura_conteudo"}
            print_progress_bar(i + 1, len(subpastas), prefix='Progresso:', suffix=f'({nome_subpasta} - Erro Leitura)', length=40)
            continue

        status_admissibilidade = verificar_admissibilidade_e_arquivamento(lista_de_paragrafos)
        metadados["status_admissibilidade"] = status_admissibilidade
        
        #---------------------------- condicoes para chamada de LLM ----------------------------

        #1° condicao pra chamar LLM: checa se o proecsso foi arquivado
        if status_admissibilidade == "Sim":
            resultados_finais[nome_subpasta] = {"metadados": metadados}
            print_progress_bar(i + 1, len(subpastas), prefix='Progresso:', suffix=f'({nome_subpasta} - Arquivado)', length=40)
            continue #continue = means que foi arquivado, ou seja, pula a chamada do LLM

        #2° condicao pra chamar LLM: procura por valores candidatos
        candidatos = encontrar_valores_candidatos(lista_de_paragrafos)
        if not candidatos:
            #caso nao encontre candidato, ele define o resultado e pula pra proxima etapa
            resultados_finais[nome_subpasta] = {"metadados": metadados, "criterio_usado": "nenhum valor candidato encontrado"}
            print_progress_bar(i + 1, len(subpastas), prefix='Progresso:', suffix=f'({nome_subpasta} - Sem Valores)', length=40)
            continue #continue = pula a chamada do llm
        
        # --- Chamada aos LLMs ---
        # o codigo so chega aqui se as duas condicoes acima falharem
        prompt_desambiguacao = criar_prompt_desambiguacao(candidatos)
        for modelo in MODELOS_PARA_TESTAR:
            print(f"\n  -> Chamando LLM: {modelo['nome']} para {nome_subpasta}...")
            resposta_json_str = chamar_llm_local(prompt_desambiguacao, modelo)
            metadados[f"resposta_{modelo['nome']}"] = resposta_json_str
        
        try:
            resposta_principal = json.loads(metadados[f"resposta_{MODELOS_PARA_TESTAR[0]['nome']}"])
            metadados["valor_final_llm"] = resposta_principal.get("valor_principal_escolhido", "Erro no JSON")
            metadados["justificativa_llm"] = resposta_principal.get("justificativa", "Erro no JSON")
        except (json.JSONDecodeError, KeyError):
            metadados["valor_final_llm"] = "Erro ao decodificar JSON"
            metadados["justificativa_llm"] = metadados[f"resposta_{MODELOS_PARA_TESTAR[0]['nome']}"]

        # gerando resumo
        prompt_resumo = criar_prompt_resumo(lista_de_paragrafos)
        print(f"  ---> Gerando resumo com LLM: --{MODELOS_PARA_TESTAR[0]['nome']}--...")
        resumo = chamar_llm_local(prompt_resumo, MODELOS_PARA_TESTAR[0])
        metadados["resumo_llm"] = resumo.replace('"', '').strip().replace("RESUMO:", "").strip()

        resultados_finais[nome_subpasta] = {"metadados": metadados}
        print_progress_bar(i + 1, len(subpastas), prefix='Progresso:', suffix=f'({nome_subpasta})', length=40)
        
    return resultados_finais

def exportar_para_excel(resultados_completos, nome_arquivo_base_excel):
    if not resultados_completos: print("Nenhum resultado para exportar."); return
    linhas_para_df = []
    colunas_modelos = [f"Resposta {m['nome']}" for m in MODELOS_PARA_TESTAR]
    
    for nome_pasta_proc, dados_proc in resultados_completos.items():
        metadados = dados_proc.get("metadados", {})
        valor_principal_str = metadados.get("valor_final_llm")
        
        valor_final_num = 0.0
        if valor_principal_str and isinstance(valor_principal_str, str):
            valor_num = converter_valor_para_numero_refinado(valor_principal_str)
            if valor_num is not None: valor_final_num = valor_num
        
        linha = {
            "Nome Pasta Original": nome_pasta_proc,
            "Número Processo (PDF)": metadados.get("numero_processo_pdf", "N/A"),
            "Número Acórdão": metadados.get("numero_acordao", "N/A"),
            "Natureza": metadados.get("natureza", "N/A"),
            "Arquivamento por Admissibilidade": metadados.get("status_admissibilidade", "Indeterminado"),
            "Valor Principal (LLM)": valor_final_num,
            "Justificativa (LLM)": metadados.get("justificativa_llm", "N/A"),
            "Resumo (LLM)": metadados.get("resumo_llm", "N/A"),
            "Nome Arquivo Processado": metadados.get("nome_arquivo_original", "N/A"),
        }
        for modelo in MODELOS_PARA_TESTAR:
            linha[f"Resposta {modelo['nome']}"] = metadados.get(f"resposta_{modelo['nome']}", "Não processado")
        linhas_para_df.append(linha)

    df = pd.DataFrame(linhas_para_df)

    def aplicar_estilo_de_linha(row):
        style = 'background-color: #FFC7CE; color: #9C0006'
        if row["Arquivamento por Admissibilidade"] == "Sim": return [style] * len(row)
        return [''] * len(row)

    styled_df = df.style.apply(aplicar_estilo_de_linha, axis=1)

    caminho_excel = f"{nome_arquivo_base_excel}.xlsx"
    try:
        styled_df.to_excel(caminho_excel, index=False, engine='openpyxl')
        print(f"\nExcel '{caminho_excel}' salvo com sucesso!")
    except Exception as e: 
        print(f"\n!!!!!!!!!!!!! Erro ao salvar Excel estilizado: {e}.")

# --- Execução Principal ---
if __name__ == '__main__':
    if not os.path.exists(PASTA_RAIZ_PROCESSOS):
        print(f"Pasta Raiz '{PASTA_RAIZ_PROCESSOS}' não encontrada.")
        exit()

    inicio = time.time()
    resultados = processar_documentos(PASTA_RAIZ_PROCESSOS)
    fim = time.time()
    print(f"\nTempo total de execução: {fim - inicio:.2f} segundos")
    
    if resultados:
        nome_arquivo_excel_base = "analise_llm_" + os.path.basename(PASTA_RAIZ_PROCESSOS)
        exportar_para_excel(resultados, nome_arquivo_excel_base)
    else:
        print(f"Nenhuma subpasta válida foi processada em '{PASTA_RAIZ_PROCESSOS}'.")