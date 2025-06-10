import os
import re
import sys
import time
import math
import fitz  # PyMuPDF
from docx import Document
import pandas as pd
from collections import defaultdict

# --- Constantes e Configurações Essenciais ---
PASTA_RAIZ_PROCESSOS = 'proc_representacoes/representacoes_SGE'
MAX_PARAGRAPH_ETAPA_2 = 400

# --- Padrões de Extração (Regex) ---
PADROES_VALOR_REFINADOS = [
    r'R\$\s*(?P<value>\d{1,3}(?:[._]\d{3})*(?:,\d{2})?)',
    r'R\$\s*(?P<number>\d+(?:[.,]\d{1,2})?)\s*(?P<unit>milh[oõ]es|mil|bilh[oõ]es|bi|tri)(?:\s+de\s+reais)?',
]

PALAVRAS_CHAVE_PONDERADAS = {
    # Prioridade 1: Sanções e Decisões Diretas (multas, devoluções, etc.)
    r"multa\s+no\s+valor\s+de": (0.4, 'sancao_direta'),
    r"condeno\s+ao\s+pagamento\s+de": (0.4, 'sancao_direta'),
    r"devolução\s+da\s+quantia\s+de": (0.4, 'sancao_direta'),
    r"fixo\s+a\s+multa\s+em": (0.3, 'sancao_direta'),
    r"valor\s+da\s+decis[ãa]o": (0.98, 'sancao_direta'),
    
    # Prioridade 2: Objeto Principal do Processo (Contratos, Licitações)
    r"valor\s+global\s+estimado\s+de": (0.95, 'objeto_principal'),
    r"preço\s+global\s+estimado\s+de": (0.95, 'objeto_principal'),
    r"valor\s+total\s+do\s+contrato": (0.95, 'objeto_principal'),
    r"proposta\s+vencedora\s+no\s+valor\s+de": (0.95, 'objeto_principal'),
    r"valor\s+do\s+contrato\s*nº?[\s\w\d/-]+,?\s+no\s+valor\s+de": (0.95, 'objeto_principal'),
    r"valor\s+estimado\s+de": (0.90, 'objeto_principal'),
    
    # Prioridade 3: Consequências e Outros Valores Fortes
    r"dano\s+ao\s+erário\s*(?:de)?": (0.5, 'valor_consequencia'),
    r"prejuízo\s+aos?\s+cofres\s+públicos\s*(?:de)?": (0.5, 'valor_consequencia'),
    
    # Contexto Geral (menor prioridade)
    r"no\s+valor\s+de": (0.7, 'contexto_geral'),
    r"valor\s+total\s*de": (0.7, 'contexto_geral'),
    r"montante\s+de": (0.6, 'contexto_geral'),
}

PALAVRAS_CHAVE_NEGATIVAS = [
    r"prejuízo\s+alegado", r"economia\s+de", r"valor\s+da\s+causa", r"lote\s+\w*\s+no\s+valor\s+de",
    r"parcela\s+de\s*r\$", r"taxa\s+de", r"juros\s+de", r"honorários\s+em\s*r\$", r"custas\s+processuais",
    r"limite\s+de\s+gasto", r"salário-mínimo", r"orçamento\s+previa", r"contrato\s+anterior",
    r"empenhos.*foram\s+anulados", r"valor\s+anulado\s+de", r"cancelamento\s+do\s+valor",
]

SECOES_DECISAO_KEYWORDS = [r"DECIS\wO", r"VOTO", r"AC[OÓ]RD[AÃ]O", r"CONCLUS\wO", r"PELO\s+EXPOSTO"]

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
    if iteration == total:
        sys.stdout.write('\n')

def converter_valor_para_numero_refinado(valor_str_original):
    """Converte uma string de valor monetário para float. Retorna (valor, erro)."""
    if not isinstance(valor_str_original, str): 
        return None, "Entrada não é string"
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
    if not valor_str: return None, "String vazia após limpeza"
    try:
        return float(valor_str) * multiplicador, None # Retorna tupla (valor, None)
    except (ValueError, TypeError):
        return None, "Erro de conversão" # Retorna tupla (None, erro)

def calcular_score_valor(valor_numerico, texto_linha, is_in_decision_section):
    score, best_keyword_category = 0.0, 'contexto_geral'
    texto_linha_lower = texto_linha.lower()
    if valor_numerico > 0: score += math.log10(valor_numerico + 1) / 10
    
    for neg_kw_regex in PALAVRAS_CHAVE_NEGATIVAS:
        if re.search(neg_kw_regex, texto_linha_lower): return 0.0, 'negativo'
            
    max_keyword_weight = 0
    for kw_regex, (peso, categoria) in PALAVRAS_CHAVE_PONDERADAS.items():
        if re.search(kw_regex, texto_linha_lower):
            score += peso
            if peso > max_keyword_weight: max_keyword_weight = peso; best_keyword_category = categoria
            
    if is_in_decision_section and best_keyword_category == 'sancao_direta': score += 1.0
    
    return score, best_keyword_category

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
                for pagina in pdf_doc: paragrafos.extend(p[4] for p in sorted(pagina.get_text("blocks"), key=lambda b: (b[1], b[0])) if p[6] == 0)
    except Exception as e: print(f"  -> Erro ao ler documento {os.path.basename(caminho_arquivo)}: {e}"); return None
    return paragrafos

def verificar_admissibilidade_e_arquivamento(lista_de_paragrafos):
    """
    Verifica se o documento foi arquivado por inadmissibilidade nas últimas páginas.
    Lógica corrigida para ser mais robusta.
    """
    if not lista_de_paragrafos:
        return "Indeterminado"
    
    # Analisa o texto das últimas páginas (últimos 30 parágrafos/blocos)
    texto_final = " ".join(lista_de_paragrafos[-30:]).upper()

    # Verifica separadamente a presença das palavras-chave essenciais
    flag_nao_conhecimento = "NÃO CONHECIMENTO" in texto_final
    flag_admissibilidade = "ADMISSIBILIDADE" in texto_final
    flag_arquivamento = "ARQUIVAMENTO" in texto_final
    
    # A condição só é verdadeira se as três flags forem verdadeiras
    if flag_nao_conhecimento and flag_admissibilidade and flag_arquivamento:
        return "Sim"
    
    return "Não"

def analisar_conteudo_para_valores(lista_de_paragrafos):
    if not lista_de_paragrafos: return None, "lista de parágrafos vazia"
    candidatos = defaultdict(list)
    
    for i, linha_texto in enumerate(lista_de_paragrafos):
        if i >= MAX_PARAGRAPH_ETAPA_2: break
        linha_texto = linha_texto.strip()
        if not linha_texto: continue
        is_in_decision_section = any(re.search(kw, linha_texto, re.IGNORECASE) for kw in SECOES_DECISAO_KEYWORDS)

        for padrao_regex_str in PADROES_VALOR_REFINADOS:
            for match in re.finditer(padrao_regex_str, linha_texto, re.IGNORECASE):
                valor_num, _ = converter_valor_para_numero_refinado(match.group(0))
                if valor_num and valor_num > 0:
                    score, categoria = calcular_score_valor(valor_num, linha_texto, is_in_decision_section)
                    if score > 0 and categoria != 'negativo':
                        candidatos[categoria].append({"valor_str": match.group(0), "valor_num": valor_num, "score": score})

    for categoria_prioritaria in ['sancao_direta', 'objeto_principal', 'valor_consequencia', 'contexto_geral']:
        if categoria_prioritaria in candidatos:
            melhor_candidato = max(candidatos[categoria_prioritaria], key=lambda x: x['score'])
            return [melhor_candidato["valor_str"]], f"etapa 2 - hierarquia: {categoria_prioritaria}"
            
    return None, "nenhum valor relevante encontrado"

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
            "numero_acordao": "NÃO ENCONTRADO", "status_admissibilidade": "Indeterminado"
        }

        if not documento_encontrado_path:
            resultados_finais[nome_subpasta] = {"metadados": metadados, "valores_extraidos": None, "criterio_usado": "documento nao encontrado"}
            print_progress_bar(i + 1, len(subpastas), prefix='Progresso:', suffix=f'({nome_subpasta} - Sem Doc)', length=40)
            continue

        if documento_encontrado_path.lower().endswith('.pdf'): metadados.update(extrair_metadados_pdf(documento_encontrado_path))
        
        # Inferência de natureza pela pasta RAIZ (fallback)
        if metadados["natureza"] == "NÃO ESPECIFICADO":
            pasta_raiz_lower_norm = os.path.basename(pasta_raiz).lower().replace(" ", "_")
            if "denuncia" in pasta_raiz_lower_norm: metadados["natureza"] = "DENUNCIA"
            elif "representacoes_sge" in pasta_raiz_lower_norm: metadados["natureza"] = "REPRESENTAÇÃO DA SGE"
            elif "representacao" in pasta_raiz_lower_norm: metadados["natureza"] = "REPRESENTAÇÃO"
        
        lista_de_paragrafos = obter_texto_documento(documento_encontrado_path)
        if lista_de_paragrafos is None:
            resultados_finais[nome_subpasta] = {"metadados": metadados, "valores_extraidos": None, "criterio_usado": "erro_leitura_conteudo"}
            print_progress_bar(i + 1, len(subpastas), prefix='Progresso:', suffix=f'({nome_subpasta} - Erro Leitura)', length=40)
            continue

        status_admissibilidade = verificar_admissibilidade_e_arquivamento(lista_de_paragrafos)
        metadados["status_admissibilidade"] = status_admissibilidade
        
        valores_finais, criterio_usado = (None, status_admissibilidade) if status_admissibilidade == "Sim" else analisar_conteudo_para_valores(lista_de_paragrafos)
        resultados_finais[nome_subpasta] = {"metadados": metadados, "valores_extraidos": valores_finais, "criterio_usado": criterio_usado}
        
        print_progress_bar(i + 1, len(subpastas), prefix='Progresso:', suffix=f'({nome_subpasta})', length=40)
        
    return resultados_finais

def exportar_para_excel(resultados_completos, nome_arquivo_base_excel):
    if not resultados_completos:
        print("Nenhum resultado para exportar.")
        return

    linhas_para_df = []
    
    for nome_pasta_proc, dados_proc in resultados_completos.items():
        metadados = dados_proc.get("metadados", {})
        valor_principal_lista = dados_proc.get("valores_extraidos")
        
        valor_final_num = 0.0
        if valor_principal_lista and isinstance(valor_principal_lista, list) and valor_principal_lista[0] is not None:
            valor_num, _ = converter_valor_para_numero_refinado(valor_principal_lista[0])
            if valor_num is not None:
                valor_final_num = valor_num
        
        linhas_para_df.append({
            "Nome Pasta Original": nome_pasta_proc,
            "Número Processo (PDF)": metadados.get("numero_processo_pdf", "N/A"),
            "Número Acórdão": metadados.get("numero_acordao", "N/A"),
            "Natureza": metadados.get("natureza", "N/A"),
            "Arquivamento por Admissibilidade": metadados.get("status_admissibilidade", "Indeterminado"),
            "Valor Principal (R$)": valor_final_num,
            "Critério de Extração": dados_proc.get("criterio_usado", "N/A"),
            "Nome Arquivo Processado": metadados.get("nome_arquivo_original", "N/A"),
            "Informações OpenAI": ""
        })

    df = pd.DataFrame(linhas_para_df)

    # funcao para aplicar esquema de cores no excel
    def aplicar_estilo_de_linha(row):
        
        cor_vermelha = 'background-color: #E63946; color: #FFFFFF'
        cor_verde =    'background-color: #2A9D8F; color: #FFFFFF'
        cor_amarela =  'background-color: #f8fb74; color: #000000'       
        cor_laranja =  'background-color: #f79256; color: #000000'
        estilo_padrao = '' # Sem cor

        # Llogica de aplicacao
        if row["Arquivamento por Admissibilidade"] == "Sim":
            return [cor_vermelha] * len(row)

        criterio = row["Critério de Extração"]
        if "objeto_principal" in criterio:
            return [cor_verde] * len(row)
        if "contexto_geral" in criterio:
            return [cor_amarela] * len(row)
        if "nenhum valor" in criterio:
            return [cor_laranja] * len(row)
        
        return [estilo_padrao] * len(row)

    # aplica o estilo ao df
    styled_df = df.style.apply(aplicar_estilo_de_linha, axis=1)

    caminho_excel = f"{nome_arquivo_base_excel}.xlsx"
    try:
        # salva o df com cores no excel
        styled_df.to_excel(caminho_excel, index=False, engine='openpyxl')
        print(f"\nExcel '{caminho_excel}' salvo com sucesso!")
    except Exception as e: 
        print(f"\n!!!!!!!!!!!!! Erro ao salvar Excel com cores: {e}. Tentando salvar sem cor.")
        try:
            df.to_excel(caminho_excel, index=False, engine='openpyxl')
            print(f"Excel '{caminho_excel}' salvo com sucesso, mas SEM formatação de cor.")
        except Exception as e_simple:
            print(f"Falha total ao salvar Excel: {e_simple}")
            
# --- Execução Principal ---
if __name__ == '__main__':
    # 1. verifica se a pasta raiz existe para evitar erro
    if not os.path.exists(PASTA_RAIZ_PROCESSOS):
        print(f"Pasta Raiz '{PASTA_RAIZ_PROCESSOS}' não encontrada. Crie-a e adicione as subpastas dos processos.")
        exit()
    print("\n")
    # 2. mede o tempo de execucao
    inicio = time.time()
    # a funcao 'processar_documentos' agora retorna só os resultados, sem o "primeiro_id"
    resultados = processar_documentos(PASTA_RAIZ_PROCESSOS)
    fim = time.time()
    print(f"Tempo total de execução: {fim - inicio:.2f} segundos")
    print("\n------------------------------Escala de Confiança no valor classificado------------------------------\nVERDE---->Alta Confiança\nAMARELO-->Baixa Confiança\nLARANJA-->Nenhum Valor Encontrado\nVERMELHO->Arquivado por Admissibilidade\nBRANCO--->Default")
    print("-----------------------------------------------------------------------------------------------------\n\n")
    # 3. vai que ne man
    if not resultados:
        print(f"Nenhuma subpasta válida encontrada ou processada em '{PASTA_RAIZ_PROCESSOS}'.")
            
    # 4. define um nome unico e padrao para a planilha de saida
    nome_arquivo_excel_base = "extracao_final_colorida"
    
    # 5. chama a exportacao UMA UNICA VEZ com TODOS os resultados
    exportar_para_excel(resultados, nome_arquivo_excel_base)