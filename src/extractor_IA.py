import os
import re
import sys
import time
import pandas as pd
from collections import defaultdict
from docx import Document
from llama_cpp import Llama

from extractor_noAI import (
    analisar_conteudo_para_valores,
    obter_texto_documento, verificar_admissibilidade_e_arquivamento,
    extrair_metadados_pdf, print_progress_bar, converter_valor_para_numero_refinado
)

PASTA_RAIZ_PROCESSOS = 'arquivos_teste_llms'
CHUNK_SIZE = 500
MAX_TOKENS_RESUMO = 3500  # número máximo de tokens aproximado para resumo fallback

# --- Inicialização dos Modelos LLM Locais ---
LLM_MODELOS = [
    {
        "nome": "Llama3",
        "modelo": Llama(model_path="./models/Meta-Llama-3-8B-Instruct.Q5_K_M.gguf", n_ctx=4096)
    },
    {
        "nome": "Phi3",
        "modelo": Llama(model_path="./models/Phi-3-mini-4k-instruct-Q4_K_M.gguf", n_ctx=4096)
    }
]

def classificar_valor_com_llm(paragrafo, valor, modelo):
    prompt = (
        f"A seguir está um trecho de um documento fiscalizatório que menciona o valor monetário '{valor}':\n"
        f"\n{paragrafo}\n"
        "Este valor corresponde ao recurso fiscalizado principal deste processo, como um contrato, licitação ou sanção relevante?\n"
        "Responda apenas com 'SIM' ou 'NÃO'."
    )
    resposta = modelo(
        prompt=prompt,
        max_tokens=10,
        temperature=0.0
    )
    return resposta["choices"][0]["text"].strip().upper().startswith("SIM")

def selecionar_valor_via_llm(paragrafos, modelo):
    candidatos = []
    for paragrafo in paragrafos:
        matches = re.findall(r'R\$\s*[\d\.,]+', paragrafo)
        for match in matches:
            valor_num, erro = converter_valor_para_numero_refinado(match)
            if erro is None and valor_num > 0:
                candidatos.append((match, paragrafo.strip(), valor_num))

    melhores = []
    for valor_str, contexto, valor_num in candidatos:
        try:
            if classificar_valor_com_llm(contexto, valor_str, modelo):
                melhores.append((valor_str, valor_num, contexto))
        except Exception:
            continue

    if melhores:
        melhor_valor = max(melhores, key=lambda x: x[1])
        return melhor_valor[0], melhor_valor[2]
    return None, None

def fallback_resumo_llm(paragrafos, modelo):
    prompt = (
        "A seguir está o conteúdo parcial de um documento fiscalizatório.\n"
        "Com base nele, identifique o valor monetário principal relacionado ao recurso fiscalizado.\n"
        "Seja direto na resposta.\nTexto:\n"
    )
    palavras = []
    for p in paragrafos:
        palavras.extend(p.split())
        if len(palavras) > MAX_TOKENS_RESUMO:
            break
    texto_limitado = " ".join(palavras[:MAX_TOKENS_RESUMO])
    resposta = modelo(
        prompt=prompt + texto_limitado,
        max_tokens=200,
        temperature=0.3
    )
    return resposta["choices"][0]["text"].strip()

def executar_extracao_com_llm():
    resultados_finais = []
    subpastas = [d for d in os.listdir(PASTA_RAIZ_PROCESSOS) if os.path.isdir(os.path.join(PASTA_RAIZ_PROCESSOS, d))]

    print_progress_bar(0, len(subpastas), prefix='Progresso:', suffix='Completo', length=40)
    for i, nome_subpasta in enumerate(subpastas):
        caminho_subpasta = os.path.join(PASTA_RAIZ_PROCESSOS, nome_subpasta)
        documento_path = None
        for ext in ['.pdf', '.docx']:
            for arq in sorted(os.listdir(caminho_subpasta)):
                if arq.lower().endswith(ext) and not arq.startswith('~$'):
                    documento_path = os.path.join(caminho_subpasta, arq)
                    break
            if documento_path: break

        metadados = extrair_metadados_pdf(documento_path) if documento_path and documento_path.endswith('.pdf') else {}
        metadados.update({"nome_pasta": nome_subpasta, "nome_arquivo": os.path.basename(documento_path) if documento_path else "N/A"})

        if not documento_path:
            resultados_finais.append({"Nome Pasta Original": nome_subpasta, "Valor Fiscalizado Algoritmo (R$)": None})
            print_progress_bar(i + 1, len(subpastas), prefix='Progresso:', suffix='(Sem Doc)', length=40)
            continue

        paragrafos = obter_texto_documento(documento_path)
        if not paragrafos:
            resultados_finais.append({"Nome Pasta Original": nome_subpasta, "Valor Fiscalizado Algoritmo (R$)": None})
            print_progress_bar(i + 1, len(subpastas), prefix='Progresso:', suffix='(Erro Leitura)', length=40)
            continue

        admissibilidade = verificar_admissibilidade_e_arquivamento(paragrafos)
        if admissibilidade == "Sim":
            resultados_finais.append({"Nome Pasta Original": nome_subpasta, "Valor Fiscalizado Algoritmo (R$)": None})
            print_progress_bar(i + 1, len(subpastas), prefix='Progresso:', suffix='(Arquivado)', length=40)
            continue

        valor_extraido, _ = analisar_conteudo_para_valores(paragrafos)
        valor_algo = valor_extraido[0] if valor_extraido else None

        linha_resultado = {
            "Nome Pasta Original": nome_subpasta,
            "Valor Fiscalizado Algoritmo (R$)": valor_algo
        }

        for modelo in LLM_MODELOS:
            nome = modelo["nome"]
            modelo_llm = modelo["modelo"]
            valor_llm, contexto = selecionar_valor_via_llm(paragrafos, modelo_llm)
            if valor_llm is None:
                try:
                    valor_llm = fallback_resumo_llm(paragrafos, modelo_llm)
                    contexto = "Resumo automatizado"
                except Exception as e:
                    valor_llm = f"Erro: {e}"
                    contexto = "Erro no fallback"
            linha_resultado[f"Resposta Interpretativa {nome}"] = valor_llm
            linha_resultado[f"Resumo {nome}"] = contexto

        resultados_finais.append(linha_resultado)
        print_progress_bar(i + 1, len(subpastas), prefix='Progresso:', suffix=f'({nome_subpasta})', length=40)

    return resultados_finais

def salvar_excel_comparativo(dados, nome_saida="resultado_comparativo"):
    df = pd.DataFrame(dados)
    df.to_excel(f"{nome_saida}.xlsx", index=False)
    print(f"\nArquivo '{nome_saida}.xlsx' salvo com sucesso.")

if __name__ == '__main__':
    inicio = time.time()
    resultados = executar_extracao_com_llm()
    salvar_excel_comparativo(resultados)
    fim = time.time()
    print(f"\nTempo total: {fim - inicio:.2f} segundos")