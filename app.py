"""
Sistema de Análise Estratégica de Follow-ups de Vendas
Automação que lê planilhas Excel e gera estratégias personalizadas usando IA
"""

import os
import time
import random
import logging
import hashlib
import pandas as pd
from flask import Flask, render_template, request, flash, redirect, url_for, session, make_response
from google import genai
from google.genai import types as genai_types
from dotenv import load_dotenv
import tempfile
from werkzeug.utils import secure_filename
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from datetime import datetime, timedelta
import json
import pickle
from PyPDF2 import PdfReader
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed

# Cores da Vendamais
VM_GREEN = colors.HexColor('#006400')  # Verde escuro
VM_ORANGE = colors.HexColor('#FF8C00') # Laranja
import io

# Carrega variáveis de ambiente
load_dotenv()

# ==================== CONFIGURAÇÃO ====================
app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'diagnostic_vm_secret_key_2024_persistent')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max
app.config['SESSION_PERMANENT'] = True
app.config['SESSION_USE_SIGNER'] = True
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=24)

# Custom Jinja2 tests and filters
@app.template_test('contains')
def contains_test(value, search_term):
    """Check if value contains search term (case-insensitive)"""
    if value is None:
        return False
    return search_term.lower() in str(value).lower()

@app.template_filter('default_if_none')
def default_if_none_filter(value, default):
    """Return default value if value is None"""
    return default if value is None else value  

ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}

# Cache em arquivo para relatórios (solução para problema de sessão)
CACHE_DIR = 'cache'
if not os.path.exists(CACHE_DIR):
    os.makedirs(CACHE_DIR)

def salvar_relatorio_cache(relatorio_data, relatorio_id):
    """Salva relatório em cache de arquivo"""
    try:
        cache_file = os.path.join(CACHE_DIR, f'relatorio_{relatorio_id}.pkl')
        with open(cache_file, 'wb') as f:
            pickle.dump({
                'data': relatorio_data,
                'timestamp': datetime.now(),
                'id': relatorio_id
            }, f)
        logger.info(f"Relatório salvo em cache: {cache_file}")
        return True
    except Exception as e:
        logger.error(f"Erro ao salvar relatório em cache: {str(e)}")
        return False

def carregar_relatorio_cache(relatorio_id):
    """Carrega relatório do cache de arquivo"""
    try:
        cache_file = os.path.join(CACHE_DIR, f'relatorio_{relatorio_id}.pkl')
        if os.path.exists(cache_file):
            with open(cache_file, 'rb') as f:
                cache_data = pickle.load(f)
            
            # Verifica se o cache não é muito antigo (24 horas)
            if datetime.now() - cache_data['timestamp'] < timedelta(hours=24):
                logger.info(f"Relatório carregado do cache: {cache_file}")
                return cache_data['data']
            else:
                # Remove cache antigo
                os.unlink(cache_file)
                logger.info(f"Cache antigo removido: {cache_file}")
        return None
    except Exception as e:
        logger.error(f"Erro ao carregar relatório do cache: {str(e)}")
        return None

def limpar_cache_antigo():
    """Remove caches antigos (mais de 24 horas)"""
    try:
        agora = datetime.now()
        for filename in os.listdir(CACHE_DIR):
            if filename.startswith('relatorio_') and filename.endswith('.pkl'):
                cache_file = os.path.join(CACHE_DIR, filename)
                try:
                    with open(cache_file, 'rb') as f:
                        cache_data = pickle.load(f)
                    if agora - cache_data['timestamp'] > timedelta(hours=24):
                        os.unlink(cache_file)
                        logger.info(f"Cache antigo removido: {filename}")
                except:
                    # Se não conseguir ler, remove o arquivo
                    try:
                        os.unlink(cache_file)
                    except:
                        pass
    except Exception as e:
        logger.error(f"Erro ao limpar cache antigo: {str(e)}")

import sys

# Configuração de logging
_stream_handler = logging.StreamHandler(io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', line_buffering=True))
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('app.log', encoding='utf-8'),
        _stream_handler,
    ]
)
logger = logging.getLogger(__name__)

# Configurações da API Gemini
GEMINI_API_KEY = os.getenv('GEMINI_API_KEY')
GEMINI_MODEL = os.getenv('GEMINI_MODEL', 'gemini-2.5-flash')
GEMINI_FALLBACK_MODEL = os.getenv('GEMINI_FALLBACK_MODEL', 'gemini-2.5-flash-lite')
GEMINI_FALLBACK_MODEL_2 = os.getenv('GEMINI_FALLBACK_MODEL_2', 'gemini-2.0-flash-lite')
MAX_RETRIES = int(os.getenv('MAX_RETRIES', '5'))
RETRY_DELAY = int(os.getenv('RETRY_DELAY', '5'))
REQUEST_DELAY = float(os.getenv('REQUEST_DELAY', '0.5'))  # Reduzido para 0.5 segundos para maior velocidade
MAX_WORKERS = int(os.getenv('MAX_WORKERS', '4'))  # Número de threads paralelas

# Cache para evitar requisições duplicadas
cache_analises = {}

# Knowledge Base da empresa
KNOWLEDGE_BASE_DIR = 'knowledge_base'
knowledge_base_text = ""

def carregar_knowledge_base():
    """Carrega e processa todos os PDFs da pasta knowledge_base"""
    global knowledge_base_text
    knowledge_base_text = ""
    
    if not os.path.exists(KNOWLEDGE_BASE_DIR):
        logger.info(f"Pasta knowledge_base não encontrada. Criando pasta...")
        os.makedirs(KNOWLEDGE_BASE_DIR)
        return
    
    pdf_files = [f for f in os.listdir(KNOWLEDGE_BASE_DIR) if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        logger.info("Nenhum PDF encontrado na pasta knowledge_base")
        return
    
    logger.info(f"Carregando {len(pdf_files)} PDFs da knowledge base...")
    
    for pdf_file in pdf_files:
        try:
            pdf_path = os.path.join(KNOWLEDGE_BASE_DIR, pdf_file)
            with open(pdf_path, 'rb') as file:
                pdf_reader = PdfReader(file)
                text_content = ""
                
                for page in pdf_reader.pages:
                    text_content += page.extract_text() + "\n"
                
                knowledge_base_text += f"\n=== CONTEÚDO DO ARQUIVO: {pdf_file} ===\n"
                knowledge_base_text += text_content + "\n"
                
                logger.info(f"PDF carregado: {pdf_file} ({len(pdf_reader.pages)} páginas)")
                
        except Exception as e:
            logger.error(f"Erro ao ler PDF {pdf_file}: {str(e)}")
    
    if knowledge_base_text:
        logger.info(f"Knowledge base carregada com sucesso ({len(knowledge_base_text)} caracteres)")
    else:   
        logger.warning("Nenhum conteúdo pôde ser extraído dos PDFs")

if not GEMINI_API_KEY:
    logger.error("GEMINI_API_KEY não encontrada nas variáveis de ambiente")
    raise ValueError("GEMINI_API_KEY é obrigatória. Configure no arquivo .env")

# Inicializa o cliente Gemini (novo SDK google-genai 1.x)
try:
    gemini_client = genai.Client(api_key=GEMINI_API_KEY, http_options={'api_version': 'v1alpha'})
    logger.info(f"Cliente Gemini configurado com sucesso usando modelo: {GEMINI_MODEL}")
except Exception as e:
    logger.error(f"Erro ao configurar cliente Gemini: {str(e)}")
    raise ValueError("Não foi possível configurar o cliente Gemini. Verifique sua API key.")

# Carrega a knowledge base da empresa ao iniciar
carregar_knowledge_base()


# ==================== FUNÇÕES AUXILIARES ====================

def allowed_file(filename):
    """Verifica se o arquivo tem extensão permitida"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def gerar_hash_cache(dados_negocio):
    """Gera um hash único para os dados do negócio (evita requisições duplicadas)"""
    dados_str = f"{dados_negocio['negocio']}|{dados_negocio['empresa']}|{dados_negocio['fase']}"
    for i in range(1, 6):
        dados_str += f"|{dados_negocio['historico_descricoes'][f'D{i}']}"
        dados_str += f"|{dados_negocio['historico_temperaturas'][f'F{i}']}"
    return hashlib.md5(dados_str.encode()).hexdigest()


def processar_item_thread(item_data):
    """
    Processa um único item em thread separada para análise paralela
    """
    try:
        # Identifica follow-ups para exibição
        ultimo_follow, proximo_follow, temperatura_atual = identificar_ultimo_followup(item_data)
        item_data["ultimo_follow"] = ultimo_follow
        item_data["proximo_follow"] = proximo_follow
        item_data["temperatura_atual"] = temperatura_atual
        
        # Chama a IA para análise estratégica
        logger.info(f"Analisando negócio em paralelo: {item_data['negocio']}")
        item_data["analise_proximo_passo"] = pedir_estrategia_ia(item_data)
        
        return item_data, None  # Sucesso
        
    except Exception as e:
        logger.warning(f"Erro na IA para {item_data['negocio']}: {str(e)}")
        # Análise simples baseada nos dados disponíveis
        temp_atual = item_data.get('temperatura_atual', 'Não informada')
        fase = item_data.get('fase', 'Não informada')
        
        item_data["analise_proximo_passo"] = f"""1. **SITUAÇÃO:** Cliente aguardando retorno em {fase}.
2. **AÇÃO:** "Olá {item_data['empresa']}, passando para confirmar se recebeu minha proposta de {item_data['negocio']}."
3. **META:** Confirmar recebimento e agendar breve alinhamento."""
        
        return item_data, str(e)  # Erro mas com análise fallback


def identificar_ultimo_followup(dados_negocio):
    """
    REGRA DE OURO: Identifica onde a conversa parou.
    Procura do Follow-up 5 para o 1 para encontrar o último preenchido.
    Retorna: (numero_followup, proximo_followup, temperatura_atual)
    """
    ultimo_follow = 0
    temperatura_atual = "Não informada"
    
    # Procura do 5 para o 1 (do mais recente para o mais antigo)
    for i in range(5, 0, -1):
        descricao = dados_negocio['historico_descricoes'][f'D{i}'].strip()
        if descricao:  # Se encontrou descrição preenchida
            ultimo_follow = i
            temperatura_atual = dados_negocio['historico_temperaturas'][f'F{i}'].strip() or "Não informada"
            break
    
    # Se não encontrou nenhum, significa que está no início (Follow-up 1)
    if ultimo_follow == 0:
        proximo_follow = 1
    elif ultimo_follow < 5:
        proximo_follow = ultimo_follow + 1
    else:
        proximo_follow = 5  # Já está no último
    
    return ultimo_follow, proximo_follow, temperatura_atual


def pedir_estrategia_ia(dados_negocio):
    """
    Envia o contexto do negócio para a IA Gemini e recebe a estratégia de venda.
    A IA age como um Diretor Comercial experiente.
    """
    # Identifica o hash para evitar requisições duplicadas
    hash_cache = gerar_hash_cache(dados_negocio)
    
    # Busca cache desativada para garantir que as novas orientações curtas sejam aplicadas
    # if hash_cache in cache_analises:
    #     return cache_analises[hash_cache]

    # Identifica onde a conversa parou
    ultimo_follow, proximo_follow, temperatura_atual = identificar_ultimo_followup(dados_negocio)
    
    # Monta histórico relevante (apenas os follow-ups preenchidos)
    historico_texto = ""
    for i in range(1, ultimo_follow + 1):
        desc = dados_negocio['historico_descricoes'][f'D{i}'].strip()
        temp = dados_negocio['historico_temperaturas'][f'F{i}'].strip()
        if desc:
            historico_texto += f"Follow-up {i} (Temperatura: {temp or 'Não informada'}): {desc}\n"
    
    # Adiciona conhecimento da empresa se disponível
    conhecimento_empresa = ""
    if knowledge_base_text:
        conhecimento_empresa = f"""
CONHECIMENTO DA EMPRESA (VENDAMAIS):
{knowledge_base_text[:10000]}

Use TODO o conteúdo técnico acima para embasar sua análise. Cite produtos, serviços e metodologias específicas da Vendamais."""
    else:
        conhecimento_empresa = "NOTA: Nenhum documento da empresa disponível. Use as melhores práticas mundiais de vendas B2B de alto ticket."
    
    prompt = f"""Você é um Mentor Comercial experiente da Vendamais. Dê uma letra RÁPIDA e FLUIDA para o vendedor "{dados_negocio['responsavel']}". Imagine que você está dando um toque rápido para ele fechar o negócio.

{conhecimento_empresa}

CONTEXTO:
- Cliente: {dados_negocio['negocio']} ({dados_negocio['empresa']})
- Próximo Passo: Follow-up #{proximo_follow} (🌡️ {temperatura_atual})

HISTÓRICO: {historico_texto if historico_texto else 'Primeiro contato agora.'}

ESTRUTURA OBRIGATÓRIA (use EXATAMENTE esses cabeçalhos em negrito, nessa ordem, sem pular nenhum):
1. **SITUAÇÃO:** O que está rolando? Explique o cenário e a indecisão do cliente (visão JOLT) de forma bem natural, como uma conversa.
2. **MENSAGEM RECOMENDADA:** Um texto pronto que soe humano (WhatsApp/Email). Nada de "prezado" ou "venho por meio desta". Seja persuasivo, dê uma recomendação clara e tire o medo dele de decidir.
3. **PRÓXIMO PASSO:** Qual o jogo aqui? Define a meta pra avançar e matar a inércia.

REGRA: Papo reto, fluido e estratégico. Proibido introduções tipo "Muito bem...", "Com base no histórico..." ou "Prezado vendedor". Vai direto ao ponto com autoridade, mas sem formalismo excessivo. OBRIGATÓRIO incluir as 3 seções sempre."""

    logger.info(f"Gerando orientação direta e fluida para: {dados_negocio['negocio']} - #{proximo_follow}")

    # Tenta até o limite configurado com backoff exponencial para rate limits
    # Ao trocar de modelo não há delay — backoff só se aplica a retentativas no mesmo modelo
    fallback_level = 0  # 0=primary, 1=fallback, 2=fallback_2
    retry_no_modelo = 0  # quantas vezes tentou o modelo atual
    MODELOS = [GEMINI_MODEL, GEMINI_FALLBACK_MODEL, GEMINI_FALLBACK_MODEL_2]
    for tentativa in range(MAX_RETRIES):
        try:
            # Delay apenas quando retenta o MESMO modelo (não na troca)
            if retry_no_modelo > 0:
                delay = RETRY_DELAY * (2 ** (retry_no_modelo - 1))
                jitter = random.uniform(0, delay * 0.3)
                delay = delay + jitter
                logger.warning(f"Aguardando {delay:.1f}s antes da tentativa {tentativa + 1} (modelo: {MODELOS[fallback_level]})...")
                time.sleep(delay)

            retry_no_modelo += 1
            modelo_atual = MODELOS[fallback_level]

            response = gemini_client.models.generate_content(
                model=modelo_atual,
                contents=prompt,
                config=genai_types.GenerateContentConfig(
                    max_output_tokens=4096,  # Equilíbrio entre robustez e brevidade
                    temperature=0.8,         # Criatividade leve para melhores argumentos
                    top_p=0.95,
                    top_k=40
                )
            )

            resultado = response.text

            # Se a resposta vier vazia ou muito curta, força um erro para tentar de novo
            if not resultado or len(resultado) < 50:
                raise ValueError("Resposta da IA muito curta ou vazia.")

            # Salva no cache para uso futuro
            cache_analises[hash_cache] = resultado

            logger.info(f"Orientação gerada com sucesso para {dados_negocio['negocio']} (modelo: {modelo_atual})")
            return resultado

        except Exception as e:
            error_str = str(e)
            should_escalate = (
                '503' in error_str or
                'UNAVAILABLE' in error_str or
                'high demand' in error_str.lower() or
                '404' in error_str or
                'NOT_FOUND' in error_str or
                'no longer available' in error_str.lower()
            )
            logger.error(f"Erro na análise do negócio {dados_negocio['negocio']}: {error_str}")
            if tentativa == MAX_RETRIES - 1:
                return f"Erro na análise (IA indisponível): {error_str}"
            if should_escalate and fallback_level < len(MODELOS) - 1:
                fallback_level += 1
                retry_no_modelo = 0  # reseta contador — próxima tentativa no novo modelo sem delay
                logger.warning(f"503 detectado — alternando para modelo fallback nível {fallback_level}: {MODELOS[fallback_level]}")
            continue

    return "Não foi possível gerar a análise (limite de tentativas excedido)."


def filtrar_negocios_por_fase(relatorio):
    """
    Filtra negócios para mostrar apenas a partir da Fase Proposta
    Oculta: Oportunidade, Contato, Conectado e Reunião
    Mostra: Proposta, Follow up 1, Follow up 2, Follow up 3, Follow up 4, Follow up 5, Negociação, etc.
    """
    fases_ocultar = ['oportunidade', 'contato', 'conectado', 'reunião']
    
    relatorio_filtrado = []
    for item in relatorio:
        fase_atual = item.get('fase', '').lower().strip()
        
        # Verifica se a fase atual não está na lista de fases a ocultar
        # Se não tiver fase ou se for a partir de Proposta, inclui
        if not fase_atual or not any(fase_oculta in fase_atual for fase_oculta in fases_ocultar):
            relatorio_filtrado.append(item)
    
    # LOG DETALHADO para debug
    logger.info(f"=== DEBUG DO FILTRO ===")
    logger.info(f"Fases a ocultar: {fases_ocultar}")
    logger.info(f"Total original: {len(relatorio)} itens")
    
    # Conta por fase
    contagem_fases = {}
    for item in relatorio:
        fase = item.get('fase', 'Não informada')
        contagem_fases[fase] = contagem_fases.get(fase, 0) + 1
    
    logger.info(f"Contagem por fase: {contagem_fases}")
    logger.info(f"Total após filtro: {len(relatorio_filtrado)} itens")
    
    # Conta por responsável após filtro
    contagem_responsaveis = {}
    for item in relatorio_filtrado:
        resp = item.get('responsavel', 'Não informado')
        contagem_responsaveis[resp] = contagem_responsaveis.get(resp, 0) + 1
    
    logger.info(f"Responsáveis após filtro: {contagem_responsaveis}")
    logger.info(f"=========================")
    
    return relatorio_filtrado


def normalizar_nome_coluna(nome):
    """Normaliza nome de coluna removendo acentos, espaços extras, aspas e convertendo para minúsculas"""
    import unicodedata
    # Remove aspas primeiro
    nome = str(nome).replace('"', '').replace("'", "").strip()
    # Remove acentos
    nome = unicodedata.normalize('NFD', nome)
    nome = ''.join(char for char in nome if unicodedata.category(char) != 'Mn')
    # Remove espaços extras e converte para minúsculas
    nome = ' '.join(nome.split()).lower()
    return nome

def encontrar_coluna_similar(df, nome_procurado):
    """Encontra coluna similar no DataFrame (case-insensitive, sem acentos, ignora 'do')"""
    nome_normalizado = normalizar_nome_coluna(nome_procurado)
    
    # Remove palavras comuns que podem variar para comparação
    palavras_ignorar = {'do', 'da', 'de', 'o', 'a', 'e', 'up', 'follow', 'proposta', 'da', 'proposta'}
    
    def limpar_palavras(texto):
        palavras = texto.split()
        return set(p for p in palavras if p not in palavras_ignorar)
    
    palavras_procuradas = limpar_palavras(nome_normalizado)
    
    # Primeiro tenta match exato (sem palavras ignoradas)
    for col in df.columns:
        col_normalizada = normalizar_nome_coluna(str(col))
        if nome_normalizado == col_normalizada:
            logger.debug(f"Match exato encontrado: '{col}' -> '{nome_procurado}'")
            return col
    
    # Depois tenta match por palavras importantes
    melhor_match = None
    melhor_score = 0
    
    for col in df.columns:
        col_normalizada = normalizar_nome_coluna(str(col))
        palavras_coluna = limpar_palavras(col_normalizada)
        
        if palavras_procuradas and palavras_coluna:
            # Calcula quantas palavras importantes estão presentes
            palavras_comuns = palavras_procuradas.intersection(palavras_coluna)
            if palavras_procuradas:  # Evita divisão por zero
                score = len(palavras_comuns) / len(palavras_procuradas)
            else:
                score = 0
            
            # Se encontrou todas as palavras importantes ou pelo menos 60% (reduzido de 70% para ser mais flexível)
            if score > melhor_score and score >= 0.6:
                melhor_score = score
                melhor_match = col
                logger.debug(f"Match parcial encontrado (score {score:.2f}): '{col}' -> '{nome_procurado}'")
    
    return melhor_match

def normalizar_colunas_df(df):
    """Normaliza nomes das colunas do DataFrame para nomes padrão"""
    mapeamento = {}
    
    # LOG para debug das colunas encontradas
    logger.info(f"=== DEBUG COLUNAS ===")
    logger.info(f"Colunas originais: {list(df.columns)}")
    
    # Mapeamento de colunas esperadas para variações possíveis
    colunas_esperadas = {
        'Nome do negócio': ['nome do negocio', 'nome do negócio', 'negocio', 'negócio'],
        'Empresa': ['empresa'],
        'Fase': ['fase'],
        'Responsavel': ['responsavel', 'responsável', 'vendedor', 'usuario', 'usuário', 'usuario', 'usuário'],
        'Temperatura da Proposta Follow 1': ['temperatura da proposta follow 1', 'temperatura follow 1', 'temperatura 1'],
        'Descrição Follow up 1': ['descrição follow up 1', 'descrição do follow up 1', 'descricao follow up 1', 'descricao do follow up 1', 'descrição do follow up 1', 'descricao do follow up 1', 'follow up 1'],
        'Temperatura da Proposta Follow 2': ['temperatura da proposta follow 2', 'temperatura follow 2', 'temperatura 2'],
        'Descrição Follow up 2': ['descrição follow up 2', 'descrição do follow up 2', 'descricao follow up 2', 'descricao do follow up 2', 'follow up 2'],
        'Temperatura da Proposta Follow 3': ['temperatura da proposta follow 3', 'temperatura follow 3', 'temperatura 3'],
        'Descrição Follow up 3': ['descrição follow up 3', 'descrição do follow up 3', 'descricao follow up 3', 'descricao do follow up 3', 'follow up 3'],
        'Temperatura da Proposta Follow 4': ['temperatura da proposta follow 4', 'temperatura follow 4', 'temperatura 4'],
        'Descrição Follow up 4': ['descrição follow up 4', 'descrição do follow up 4', 'descricao follow up 4', 'descricao do follow up 4', 'follow up 4'],
        'Temperatura da Proposta Follow 5': ['temperatura da proposta follow 5', 'temperatura follow 5', 'temperatura 5'],
        'Descrição Follow up 5': ['descrição follow up 5', 'descrição do follow up 5', 'descricao follow up 5', 'descricao do follow up 5', 'follow up 5'],
    }
    
    # Para cada coluna esperada, tenta encontrar no DataFrame
    for coluna_esperada, variacoes in colunas_esperadas.items():
        coluna_encontrada = encontrar_coluna_similar(df, coluna_esperada)
        if coluna_encontrada:
            mapeamento[coluna_encontrada] = coluna_esperada
        else:
            # Tenta com variações
            for variacao in variacoes:
                coluna_encontrada = encontrar_coluna_similar(df, variacao)
                if coluna_encontrada:
                    mapeamento[coluna_encontrada] = coluna_esperada
                    break
    
    # Renomeia as colunas encontradas
    if mapeamento:
        df = df.rename(columns=mapeamento)
        logger.info(f"Colunas normalizadas ({len(mapeamento)} colunas): {list(mapeamento.items())}")
    else:
        logger.warning("Nenhuma coluna foi normalizada. Verifique se os nomes das colunas estão corretos.")
    
    # Se for CSV sem cabeçalho, mapeia por posição
    if df.columns.tolist() == list(range(df.shape[1])):
        logger.info("CSV sem cabeçalho detectado, mapeando por posição...")
        mapeamento_posicional = {
            0: 'Empresa',
            2: 'Fase', 
            3: 'Responsavel',
            7: 'Descrição Follow up 1',
            8: 'Descrição Follow up 2', 
            9: 'Descrição Follow up 3',
            10: 'Descrição Follow up 4',
            11: 'Descrição Follow up 5',
            12: 'Temperatura da Proposta Follow 1',
            13: 'Temperatura da Proposta Follow 2',
            14: 'Temperatura da Proposta Follow 3', 
            15: 'Temperatura da Proposta Follow 4',
            16: 'Temperatura da Proposta Follow 5',
        }
        
        for pos, nome in mapeamento_posicional.items():
            if pos < df.shape[1]:
                df.rename(columns={pos: nome}, inplace=True)
        
        logger.info(f"Mapeamento posicional aplicado: {mapeamento_posicional}")
    
    # Cria colunas faltantes com valores vazios (para garantir que o sistema funcione)
    colunas_esperadas = [
        'Nome do negócio', 'Empresa', 'Fase', 'Responsavel',
        'Temperatura da Proposta Follow 1', 'Descrição Follow up 1',
        'Temperatura da Proposta Follow 2', 'Descrição Follow up 2',
        'Temperatura da Proposta Follow 3', 'Descrição Follow up 3',
        'Temperatura da Proposta Follow 4', 'Descrição Follow up 4',
        'Temperatura da Proposta Follow 5', 'Descrição Follow up 5',
    ]
    
    colunas_criadas = []
    for coluna in colunas_esperadas:
        if coluna not in df.columns:
            df[coluna] = ''  # Cria coluna vazia
            colunas_criadas.append(coluna)
    
    if colunas_criadas:
        logger.info(f"Colunas criadas automaticamente (vazias): {', '.join(colunas_criadas)}")
    
    logger.info(f"Colunas finais: {list(df.columns)}")
    logger.info(f"=====================")
    
    return df

def validar_planilha(df):
    """
    Valida a planilha de forma flexível - apenas informa colunas faltantes, mas NUNCA bloqueia.
    Esta função sempre retorna True e nunca gera exceções.
    """
    try:
        colunas_desejadas = [
            'Nome do negócio', 'Empresa', 'Fase', 'Responsavel',
            'Temperatura da Proposta Follow 1', 'Descrição Follow up 1'
        ]

        colunas_faltantes = []
        colunas_encontradas = []
        
        for coluna in colunas_desejadas:
            if coluna in df.columns:
                colunas_encontradas.append(coluna)
            else:
                colunas_faltantes.append(coluna)

        if colunas_encontradas:
            logger.info(f"✅ Colunas encontradas: {', '.join(colunas_encontradas)}")
        
        if colunas_faltantes:
            logger.warning(f"⚠️ Colunas não encontradas (sistema continuará funcionando normalmente): {', '.join(colunas_faltantes)}")
            logger.info(f"📋 Todas as colunas disponíveis no arquivo: {', '.join(list(df.columns)[:20])}")

        # SEMPRE retorna True - nunca bloqueia
        return True
    except Exception as e:
        # Se der qualquer erro, apenas loga e continua
        logger.warning(f"Erro na validação (mas continuando): {str(e)}")
        return True  # Sempre retorna True para não bloquear


def ler_planilha_excel(file_path, filename):
    """
    Lê arquivo Excel/CSV com múltiplas estratégias de fallback.
    Suporta .xlsx, .xls, .csv e até arquivos HTML disfarçados de Excel.
    """
    file_ext = filename.lower().split('.')[-1] if '.' in filename else ''
    logger.info(f"Processando arquivo.{file_ext}: {filename}")
    
    df = None
    error_messages = []
    
    # PRIORIDADE 0: Se for CSV, lê diretamente (mais simples e confiável)
    if file_ext == 'csv':
        logger.info("Arquivo CSV detectado, lendo diretamente...")
        try:
            # Tenta diferentes separadores e encodings comuns
            separadores = [';', ',', '\t']
            encodings = ['utf-8-sig', 'utf-8', 'latin-1', 'iso-8859-1', 'cp1252']
            
            for encoding in encodings:
                for sep in separadores:
                    try:
                        # Primeiro tenta ler com cabeçalho (header='infer')
                        df = pd.read_csv(file_path, sep=sep, encoding=encoding, skipinitialspace=True)
                        if len(df.columns) > 1:
                            logger.info(f"✅ CSV lido com sucesso (separador='{sep}', encoding={encoding}): {len(df)} linhas, {len(df.columns)} colunas")
                            break
                    except Exception:
                        continue
                if df is not None and len(df.columns) > 1:
                    break
                # Se ainda não tem cabeçalho reconhecível, tenta ler sem cabeçalho
                try:
                    df = pd.read_csv(file_path, sep=sep, encoding=encoding, header=None, skipinitialspace=True)
                    logger.info(f"✅ CSV lido sem cabeçalho (separador='{sep}', encoding={encoding}): {len(df)} linhas, {len(df.columns)} colunas")
                except Exception:
                    df = None
            
            # Se ainda não conseguiu, tenta sem especificar separador (detecção automática)
            if df is None or len(df.columns) <= 1:
                for encoding in encodings:
                    try:
                        df = pd.read_csv(file_path, encoding=encoding, skipinitialspace=True)
                        if len(df.columns) > 1:
                            logger.info(f"✅ CSV lido com detecção automática (encoding={encoding}): {len(df)} linhas, {len(df.columns)} colunas")
                            break
                    except Exception as e:
                        continue
            
            if df is None or len(df.columns) <= 1:
                error_messages.append("Não foi possível ler o CSV com nenhum separador/encoding testado")
        except Exception as e:
            error_messages.append(f"Erro ao ler CSV: {str(e)}")
        
        if df is not None and not df.empty:
            # Se o CSV não tem cabeçalho, atribuímos nomes de colunas esperados com base na posição conhecida
            if df.columns.tolist() == list(range(df.shape[1])):
                # Mapeamento posicional (ajuste conforme seu CSV)
                colunas_pos = [
                    'Empresa',          # 0
                    'Tipo',            # 1 (ignorado)
                    'Fase',            # 2 (ignorado)
                    'Responsavel',     # 3
                    'Data',            # 4 (ignorado)
                    # ... campos intermediários ignorados ...
                    'Temperatura Atual',  # penúltimo antes do ID, ajuste conforme necessidade
                ]
                # Preencher até o número de colunas existentes
                for i, nome in enumerate(colunas_pos):
                    if i < df.shape[1]:
                        df.rename(columns={i: nome}, inplace=True)
                logger.info("Colunas do CSV sem cabeçalho foram renomeadas com base em posições conhecidas.")
            return df
    
    # PRIMEIRO: Verifica assinaturas de arquivo Excel válido
    is_valid_excel = False
    is_html = False
    
    try:
        with open(file_path, 'rb') as f:
            header = f.read(8)  # Lê apenas os primeiros 8 bytes para verificar assinatura
            
            # Assinaturas de arquivos Excel válidos
            excel_signatures = [
                b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1',  # .xls (OLE2 format)
                b'\x50\x4b\x03\x04',  # .xlsx (ZIP format - começa com PK)
                b'\x50\x4b\x05\x06',  # .xlsx (ZIP empty)
                b'\x50\x4b\x07\x08'   # .xlsx (ZIP spanned)
            ]
            
            # Verifica se é um Excel válido
            for sig in excel_signatures:
                if header.startswith(sig):
                    is_valid_excel = True
                    logger.info(f"Assinatura Excel válida detectada: {sig.hex()}")
                    break
            
            # Se não é Excel válido, verifica se é HTML (lê mais bytes)
            if not is_valid_excel:
                f.seek(0)
                header_full = f.read(500)
                
                # Detecta HTML de várias formas (incluindo BOM)
                # O caso mais comum: arquivo HTML salvo com extensão .xls
                is_html = (
                    header_full.startswith(b'\xef\xbb\xbf<meta') or  # BOM + <meta
                    header_full.startswith(b'<meta') or 
                    header_full.startswith(b'<!DOCTYPE') or 
                    header_full.startswith(b'<html') or
                    header_full.startswith(b'\xef\xbb\xbf<!DOCTYPE') or
                    header_full.startswith(b'\xef\xbb\xbf<html') or
                    (b'<table' in header_full and b'<tr>' in header_full and b'<td>' in header_full) or
                    (b'http-equiv' in header_full and b'Content-type' in header_full)  # Meta tag comum em HTML
                )
                
                if is_html:
                    logger.warning("HTML detectado no arquivo (arquivo HTML salvo com extensão .xls/.xlsx)")
    except Exception as e:
        logger.warning(f"Erro ao verificar header do arquivo: {str(e)}")
    
    # PRIORIDADE 1: Tenta ler como Excel primeiro (APENAS se tem assinatura válida E não é HTML)
    if is_valid_excel and not is_html:
        logger.info("Tentando ler como arquivo Excel válido...")
        
        if file_ext == 'xls':
            # Para .xls, tenta xlrd primeiro (mais compatível)
            try:
                df = pd.read_excel(file_path, engine='xlrd')
                logger.info("✅ Arquivo .xls lido com sucesso usando xlrd")
            except Exception as e1:
                logger.warning(f"xlrd falhou: {str(e1)}")
                error_messages.append(f"xlrd: {str(e1)}")
                
                # Tenta openpyxl como fallback
                try:
                    df = pd.read_excel(file_path, engine='openpyxl')
                    logger.info("✅ Arquivo .xls lido com sucesso usando openpyxl (fallback)")
                except Exception as e2:
                    logger.warning(f"openpyxl também falhou: {str(e2)}")
                    error_messages.append(f"openpyxl: {str(e2)}")
                
                # Tenta sem engine específica
                if df is None:
                    try:
                        df = pd.read_excel(file_path)
                        logger.info("✅ Arquivo .xls lido sem engine específica")
                    except Exception as e3:
                        error_messages.append(f"default: {str(e3)}")
        
        elif file_ext == 'xlsx':
            # Para .xlsx, tenta openpyxl primeiro
            try:
                df = pd.read_excel(file_path, engine='openpyxl')
                logger.info("✅ Arquivo .xlsx lido com sucesso usando openpyxl")
            except Exception as e1:
                logger.warning(f"openpyxl falhou: {str(e1)}")
                error_messages.append(f"openpyxl: {str(e1)}")
                
                # Tenta xlrd como fallback
                try:
                    df = pd.read_excel(file_path, engine='xlrd')
                    logger.info("✅ Arquivo .xlsx lido com sucesso usando xlrd (fallback)")
                except Exception as e2:
                    logger.warning(f"xlrd também falhou: {str(e2)}")
                    error_messages.append(f"xlrd: {str(e2)}")
                
                # Tenta sem engine específica
                if df is None:
                    try:
                        df = pd.read_excel(file_path)
                        logger.info("✅ Arquivo .xlsx lido sem engine específica")
                    except Exception as e3:
                        error_messages.append(f"default: {str(e3)}")
        
        # Se ainda não conseguiu e tem assinatura Excel, tenta tratamento especial
        if df is None and is_valid_excel:
            logger.warning("Arquivo tem assinatura Excel mas não foi possível ler. Tentando tratamento especial...")
            # Tenta remover BOM se existir e ler novamente
            try:
                with open(file_path, 'rb') as f:
                    content = f.read()
                
                # Remove BOM se existir no início
                if content.startswith(b'\xef\xbb\xbf'):
                    logger.info("Removendo BOM do arquivo...")
                    content = content[3:]
                    temp_path = file_path + '_no_bom.xls'
                    with open(temp_path, 'wb') as f:
                        f.write(content)
                    
                    try:
                        df = pd.read_excel(temp_path, engine='xlrd')
                        logger.info("✅ Arquivo lido após remover BOM")
                    except:
                        pass
                    finally:
                        try:
                            os.unlink(temp_path)
                        except:
                            pass
            except Exception as e:
                logger.warning(f"Tratamento especial falhou: {str(e)}")
    
    # PRIORIDADE 2: Se detectou HTML (mesmo que tenha extensão .xls/.xlsx), tenta converter HTML PRIMEIRO
    if is_html:
        logger.warning("Conteúdo HTML detectado, tentando converter HTML para DataFrame...")
        
        # Estratégia 1: Remove BOM primeiro e tenta pd.read_html
        try:
            with open(file_path, 'rb') as f:
                content_bytes = f.read()
            
            # Remove BOM se existir
            if content_bytes.startswith(b'\xef\xbb\xbf'):
                logger.info("Removendo BOM do arquivo HTML...")
                content_bytes = content_bytes[3:]
            
            # Salva temporariamente sem BOM
            temp_html_path = file_path + '_temp_clean.html'
            with open(temp_html_path, 'wb') as f:
                f.write(content_bytes)
            
            # Tenta ler HTML com diferentes encodings
            encodings_to_try = ['utf-8', 'utf-8-sig', 'latin-1', 'iso-8859-1', 'cp1252']
            for encoding in encodings_to_try:
                try:
                    df_html = pd.read_html(temp_html_path, encoding=encoding)
                    if df_html and len(df_html) > 0:
                        # Pega a primeira tabela com mais colunas (geralmente é a principal)
                        df = max(df_html, key=lambda x: len(x.columns) if not x.empty else 0)
                        if not df.empty:
                            logger.info(f"✅ HTML convertido com sucesso (encoding={encoding}): {len(df)} linhas, {len(df.columns)} colunas")
                            break
                except Exception as e1:
                    if encoding == encodings_to_try[0]:
                        logger.warning(f"pd.read_html com encoding {encoding} falhou: {str(e1)}")
                        error_messages.append(f"read_html({str(e1)})")
                    continue
            
            # Remove arquivo temporário
            try:
                os.unlink(temp_html_path)
            except:
                pass
                
        except Exception as e:
            logger.warning(f"Erro ao processar HTML: {str(e)}")
            error_messages.append(f"process_html: {str(e)}")
        
        # Estratégia 2: Se ainda não conseguiu, tenta direto no arquivo original
        if df is None:
            encodings_to_try = ['utf-8', 'utf-8-sig', 'latin-1', 'iso-8859-1', 'cp1252']
            for encoding in encodings_to_try:
                try:
                    df_html = pd.read_html(file_path, encoding=encoding)
                    if df_html and len(df_html) > 0:
                        df = max(df_html, key=lambda x: len(x.columns) if not x.empty else 0)
                        if not df.empty:
                            logger.info(f"✅ HTML convertido diretamente (encoding={encoding}): {len(df)} linhas, {len(df.columns)} colunas")
                            break
                except Exception as e1:
                    continue
        
        # Estratégia 2: Remove BOM manualmente e tenta novamente
        if df is None:
            try:
                with open(file_path, 'rb') as f:
                    content_bytes = f.read()
                
                # Remove BOM se existir
                if content_bytes.startswith(b'\xef\xbb\xbf'):
                    content_bytes = content_bytes[3:]
                
                # Salva temporariamente sem BOM
                temp_html_path = file_path + '_clean.html'
                with open(temp_html_path, 'wb') as f:
                    f.write(content_bytes)
                
                df_html = pd.read_html(temp_html_path, encoding='utf-8')
                if df_html and len(df_html) > 0:
                    df = max(df_html, key=lambda x: len(x.columns) if not x.empty else 0)
                    if not df.empty:
                        logger.info(f"HTML convertido após remover BOM: {len(df)} linhas, {len(df.columns)} colunas")
                
                # Remove arquivo temporário
                try:
                    os.unlink(temp_html_path)
                except:
                    pass
            except Exception as e2:
                logger.warning(f"Conversão HTML com BOM removido falhou: {str(e2)}")
                error_messages.append(f"read_html_bom({str(e2)})")
        
        # Estratégia 3: Tenta ler como CSV (às vezes HTML é salvo como CSV)
        if df is None:
            try:
                for sep in [';', ',', '\t']:
                    try:
                        df_test = pd.read_csv(file_path, sep=sep, encoding='utf-8-sig', skiprows=0)
                        if len(df_test.columns) > 1:  # Se encontrou múltiplas colunas
                            df = df_test
                            logger.info(f"HTML lido como CSV com separador '{sep}': {len(df)} linhas, {len(df.columns)} colunas")
                            break
                    except:
                        continue
            except Exception as e4:
                logger.warning(f"Leitura como CSV falhou: {str(e4)}")
    
    # PRIORIDADE 3: Se não é HTML e tem extensão .xls/.xlsx mas não tem assinatura válida, tenta ler como Excel
    if df is None and not is_html and file_ext in ['xls', 'xlsx']:
        logger.info("Tentando ler como Excel (extensão .xls/.xlsx mas sem assinatura detectada)...")
        if file_ext == 'xls':
            try:
                df = pd.read_excel(file_path, engine='xlrd')
                logger.info("✅ Arquivo .xls lido com sucesso usando xlrd")
            except Exception as e1:
                logger.warning(f"xlrd falhou: {str(e1)}")
                error_messages.append(f"xlrd: {str(e1)}")
                try:
                    df = pd.read_excel(file_path, engine='openpyxl')
                    logger.info("✅ Arquivo .xls lido com sucesso usando openpyxl")
                except Exception as e2:
                    error_messages.append(f"openpyxl: {str(e2)}")
        elif file_ext == 'xlsx':
            try:
                df = pd.read_excel(file_path, engine='openpyxl')
                logger.info("✅ Arquivo .xlsx lido com sucesso usando openpyxl")
            except Exception as e1:
                logger.warning(f"openpyxl falhou: {str(e1)}")
                error_messages.append(f"openpyxl: {str(e1)}")
                try:
                    df = pd.read_excel(file_path, engine='xlrd')
                    logger.info("✅ Arquivo .xlsx lido com sucesso usando xlrd")
                except Exception as e2:
                    error_messages.append(f"xlrd: {str(e2)}")
    
    # Se ainda não conseguiu ler, dá mensagem de erro clara
    if df is None:
        if is_html:
            raise ValueError(
                "O arquivo parece ser HTML (página web) e não um arquivo Excel válido. "
                "Por favor, abra o arquivo no Excel e salve como '.xlsx' ou '.xls' antes de enviar. "
                f"Erros de conversão: {', '.join(error_messages) if error_messages else 'Não foi possível converter HTML'}"
            )
        elif is_valid_excel:
            all_errors = ", ".join(error_messages) if error_messages else "Erro desconhecido"
            raise ValueError(
                f"O arquivo tem assinatura Excel válida mas não foi possível ler. "
                f"O arquivo pode estar corrompido. Erros: {all_errors}. "
                f"Tente abrir o arquivo no Excel e salvar novamente."
            )
        else:
            all_errors = ", ".join(error_messages) if error_messages else "Erro desconhecido"
            raise ValueError(
                f"Não foi possível ler o arquivo Excel. "
                f"Verifique se o arquivo não está corrompido. Erros: {all_errors}. "
                f"Se o problema persistir, tente abrir o arquivo no Excel e salvar novamente como .xlsx"
            )
    
    if df.empty:
        raise ValueError("O arquivo está vazio ou não contém dados válidos")
    
    return df


# ==================== ROTAS ====================

@app.route('/')
def index():
    """Página inicial com formulário de upload e relatórios recentes"""
    try:
        # Lista relatórios recentes do cache
        relatorios_recentes = []
        if os.path.exists(CACHE_DIR):
            cache_files = [f for f in os.listdir(CACHE_DIR) if f.startswith('relatorio_') and f.endswith('.pkl')]
            
            for cache_file in sorted(cache_files, reverse=True)[:10]:  # Últimos 10 relatórios
                try:
                    cache_path = os.path.join(CACHE_DIR, cache_file)
                    with open(cache_path, 'rb') as f:
                        cache_data = pickle.load(f)
                    
                    # Extrai ID do nome do arquivo
                    relatorio_id = cache_file.replace('relatorio_', '').replace('.pkl', '')
                    
                    relatorios_recentes.append({
                        'id': relatorio_id,
                        'timestamp': cache_data['timestamp'],
                        'data_formatada': cache_data['timestamp'].strftime('%d/%m/%Y %H:%M'),
                        'total_itens': len(cache_data['data']) if isinstance(cache_data['data'], list) else 0
                    })
                except Exception as e:
                    logger.warning(f"Erro ao ler cache {cache_file}: {str(e)}")
                    continue
        
        return render_template('index.html', relatorios_recentes=relatorios_recentes)
    except Exception as e:
        logger.error(f"Erro ao carregar página inicial: {str(e)}")
        return render_template('index.html', relatorios_recentes=[])


@app.route('/relatorios')
def listar_relatorios():
    """Lista todos os relatórios disponíveis no cache"""
    try:
        relatorios = []
        if os.path.exists(CACHE_DIR):
            cache_files = [f for f in os.listdir(CACHE_DIR) if f.startswith('relatorio_') and f.endswith('.pkl')]
            
            for cache_file in sorted(cache_files, reverse=True):
                try:
                    cache_path = os.path.join(CACHE_DIR, cache_file)
                    with open(cache_path, 'rb') as f:
                        cache_data = pickle.load(f)
                    
                    relatorio_id = cache_file.replace('relatorio_', '').replace('.pkl', '')
                    
                    relatorios.append({
                        'id': relatorio_id,
                        'timestamp': cache_data['timestamp'],
                        'data_formatada': cache_data['timestamp'].strftime('%d/%m/%Y %H:%M'),
                        'total_itens': len(cache_data['data']) if isinstance(cache_data['data'], list) else 0
                    })
                except Exception as e:
                    logger.warning(f"Erro ao ler cache {cache_file}: {str(e)}")
                    continue
        
        return render_template('relatorios.html', relatorios=relatorios)
    except Exception as e:
        logger.error(f"Erro ao listar relatórios: {str(e)}")
        flash('Erro ao carregar lista de relatórios', 'error')
        return redirect(url_for('index'))


@app.route('/relatorio/<relatorio_id>')
def ver_relatorio(relatorio_id):
    """Visualiza um relatório específico pelo ID"""
    try:
        logger.info(f"Carregando relatório específico: {relatorio_id}")
        
        # Carrega do cache usando o ID
        dados_cache = carregar_relatorio_cache(relatorio_id)
        
        if not dados_cache:
            flash('Relatório não encontrado ou expirado. Relatórios ficam disponíveis por 24 horas.', 'error')
            return redirect(url_for('listar_relatorios'))
        
        # Define o ID atual na sessão para outras funcionalidades
        session['relatorio_id_atual'] = relatorio_id
        
        # Processa os dados
        if isinstance(dados_cache, dict):
            relatorio_final = dados_cache.get('relatorio_final', [])
            relatorio_agrupado = dados_cache.get('relatorio_agrupado', {})
        else:
            # Formato antigo - converte
            relatorio_final = dados_cache
            relatorio_agrupado = {}
            for item in relatorio_final:
                responsavel_item = item.get('responsavel', 'Não informado')
                if responsavel_item not in relatorio_agrupado:
                    relatorio_agrupado[responsavel_item] = []
                relatorio_agrupado[responsavel_item].append(item)
        
        logger.info(f"Relatório {relatorio_id} carregado: {len(relatorio_final)} itens, {len(relatorio_agrupado)} responsáveis")
        
        return render_template('relatorio.html', 
                             relatorio_final=relatorio_final,
                             relatorio_agrupado=relatorio_agrupado,
                             total_negocios=len(relatorio_final))
        
    except Exception as e:
        logger.error(f"Erro ao carregar relatório {relatorio_id}: {str(e)}")
        flash('Erro ao carregar relatório', 'error')
        return redirect(url_for('listar_relatorios'))


@app.route('/gerar_pdf_relatorio/<relatorio_id>')
def gerar_pdf_relatorio(relatorio_id):
    """Gera PDF para um relatório específico pelo ID"""
    try:
        logger.info(f"Gerando PDF para relatório: {relatorio_id}")
        
        # Carrega do cache usando o ID
        dados_cache = carregar_relatorio_cache(relatorio_id)
        
        if not dados_cache:
            flash('Relatório não encontrado ou expirado', 'error')
            return redirect(url_for('listar_relatorios'))
        
        # Processa os dados
        if isinstance(dados_cache, dict):
            relatorio_final = dados_cache.get('relatorio_final', [])
            relatorio_agrupado = dados_cache.get('relatorio_agrupado', {})
        else:
            # Formato antigo - converte
            relatorio_final = dados_cache
            relatorio_agrupado = {}
            for item in relatorio_final:
                responsavel_item = item.get('responsavel', 'Não informado')
                if responsavel_item not in relatorio_agrupado:
                    relatorio_agrupado[responsavel_item] = []
                relatorio_agrupado[responsavel_item].append(item)
        
        # Gera o PDF usando a função existente
        from reportlab.lib.pagesizes import A4
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.lib.units import inch
        from reportlab.lib import colors
        from io import BytesIO
        
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)
        
        # Configura estilos
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle('CustomTitle', parent=styles['Heading1'], 
                                     textColor=VM_GREEN, spaceAfter=20)
        
        story = []
        
        # Título
        story.append(Paragraph("Relatório de Análise Estratégica de Follow-ups", title_style))
        story.append(Paragraph(f"Data: {datetime.now().strftime('%d/%m/%Y %H:%M')}", styles['Normal']))
        story.append(Spacer(1, 20))
        
        # Agrupa por responsável
        for responsavel, itens in sorted(relatorio_agrupado.items()):
            story.append(Paragraph(f"Responsável: {responsavel}", styles['Heading2']))
            
            for item in itens:
                story.append(Paragraph(f"<b>Negócio:</b> {item.get('negocio', 'N/A')}", styles['Normal']))
                story.append(Paragraph(f"<b>Empresa:</b> {item.get('empresa', 'N/A')}", styles['Normal']))
                story.append(Paragraph(f"<b>Fase:</b> {item.get('fase', 'N/A')}", styles['Normal']))
                story.append(Paragraph(f"<b>Próximo Passo:</b> Follow-up #{item.get('proximo_follow', 'N/A')}", styles['Normal']))
                
                # Análise da IA
                analise = item.get('analise_proximo_passo', 'Não disponível')
                story.append(Paragraph(f"<b>Análise Estratégica:</b>", styles['Normal']))
                
                # Formata o texto da análise
                linhas = analise.split('\n')
                for linha in linhas:
                    if linha.strip():
                        story.append(Paragraph(linha.strip(), styles['Normal']))
                
                story.append(Spacer(1, 12))
            
            story.append(PageBreak())
        
        doc.build(story)
        
        # Prepara resposta
        response = make_response(buffer.getvalue())
        response.headers['Content-Type'] = 'application/pdf'
        response.headers['Content-Disposition'] = f'inline; filename=relatorio_{relatorio_id}_{datetime.now().strftime("%Y%m%d_%H%M")}.pdf'
        
        logger.info(f"PDF gerado com sucesso para relatório {relatorio_id}")
        return response
        
    except Exception as e:
        logger.error(f"Erro ao gerar PDF para relatório {relatorio_id}: {str(e)}")
        flash(f'Erro ao gerar PDF: {str(e)}', 'error')
        return redirect(url_for('ver_relatorio', relatorio_id=relatorio_id))


@app.route('/todos')
def ver_todos():
    """Visualização de todos os negócios filtrados por fase (Proposta para frente)"""
    try:
        logger.info("=== ROTA /TOS INICIADA ===")
        
        # Carrega dados do cache
        relatorio_final = None
        relatorio_agrupado = None
        
        # Tenta 1: Cache de arquivo
        if 'relatorio_id_atual' in session:
            relatorio_id = session['relatorio_id_atual']
            logger.info(f"Found relatorio_id_atual in session: {relatorio_id}")
            dados_cache = carregar_relatorio_cache(relatorio_id)
            if dados_cache:
                if isinstance(dados_cache, dict):
                    relatorio_final = dados_cache.get('relatorio_final', [])
                    relatorio_agrupado = dados_cache.get('relatorio_agrupado', {})
                    logger.info(f"Cache loaded: {len(relatorio_final)} itens, {len(relatorio_agrupado)} responsáveis")
                else:
                    # Formato antigo - converte para novo
                    relatorio_final = dados_cache
                    relatorio_agrupado = {}
                    for item in relatorio_final:
                        responsavel_item = item.get('responsavel', 'Não informado')
                        if responsavel_item not in relatorio_agrupado:
                            relatorio_agrupado[responsavel_item] = []
                        relatorio_agrupado[responsavel_item].append(item)
                    logger.info(f"Old format converted: {len(relatorio_agrupado)} responsáveis")
            else:
                logger.warning(f"Failed to load cache for ID: {relatorio_id}")
        else:
            logger.warning("relatorio_id_atual not found in session")
        
        # Tenta 2: Sessão (backup)
        if not relatorio_agrupado and 'relatorio_data' in session:
            logger.info("Trying backup session data")
            dados_session = session['relatorio_data']
            if isinstance(dados_session, dict):
                relatorio_final = dados_session.get('relatorio_final', [])
                relatorio_agrupado = dados_session.get('relatorio_agrupado', {})
                logger.info(f"Session backup loaded: {len(relatorio_agrupado)} responsáveis")
            else:
                relatorio_final = dados_session
                relatorio_agrupado = {}
                for item in relatorio_final:
                    responsavel_item = item.get('responsavel', 'Não informado')
                    if responsavel_item not in relatorio_agrupado:
                        relatorio_agrupado[responsavel_item] = []
                    relatorio_agrupado[responsavel_item].append(item)
                logger.info(f"Session backup converted: {len(relatorio_agrupado)} responsáveis")
        
        # Tenta 3: Busca automática no cache
        if not relatorio_agrupado:
            logger.info("Trying auto-discovery in cache files")
            try:
                for filename in os.listdir(CACHE_DIR):
                    if filename.startswith('relatorio_') and filename.endswith('.pkl'):
                        cache_file = os.path.join(CACHE_DIR, filename)
                        try:
                            with open(cache_file, 'rb') as f:
                                cache_data = pickle.load(f)
                            
                            if datetime.now() - cache_data['timestamp'] < timedelta(hours=24):
                                dados_cache = cache_data['data']
                                if isinstance(dados_cache, dict) and 'relatorio_agrupado' in dados_cache:
                                    relatorio_agrupado = dados_cache['relatorio_agrupado']
                                    relatorio_final = dados_cache.get('relatorio_final', [])
                                    logger.info(f"Auto-discovered: {len(relatorio_agrupado)} responsáveis")
                                    break
                        except Exception as e:
                            logger.warning(f"Failed to read cache file {filename}: {e}")
                            continue
            except Exception as e:
                logger.error(f"Error in auto-discovery: {e}")
        
        # Debug final
        if relatorio_agrupado:
            logger.info(f"Responsáveis encontrados: {list(relatorio_agrupado.keys())}")
            for resp, itens in relatorio_agrupado.items():
                logger.info(f"  - {resp}: {len(itens)} itens")
        else:
            logger.error("Nenhum dado encontrado em nenhum lugar!")
        
        if not relatorio_final:
            flash('Dados não encontrados. Por favor, processe a planilha novamente.', 'error')
            return redirect(url_for('index'))
        
        logger.info(f"=== RENDERIZANDO TEMPLATES COM {len(relatorio_agrupado)} RESPONSÁVEIS ===")
        
        return render_template('relatorio.html', 
                             relatorio=relatorio_final, 
                             total=len(relatorio_final),
                             relatorio_agrupado=relatorio_agrupado, 
                             responsaveis=list(relatorio_agrupado.keys()))
        
    except Exception as e:
        logger.error(f"Erro ao visualizar todos os negócios: {str(e)}")
        flash(f'Erro ao carregar dados: {str(e)}', 'error')
        return redirect(url_for('index'))


@app.route('/processar', methods=['POST'])
def processar():
    """Processa a planilha enviada e gera análises estratégicas"""
    if 'file' not in request.files:
        logger.error("Nenhum arquivo enviado na requisição")
        flash('Nenhum arquivo selecionado', 'error')
        return redirect(url_for('index'))
    
    file = request.files['file']
    
    if file.filename == '':
        logger.error("Nome de arquivo vazio")
        flash('Nenhum arquivo selecionado', 'error')
        return redirect(url_for('index'))
    
    if not allowed_file(file.filename):
        logger.error(f"Formato de arquivo inválido: {file.filename}")
        flash('Formato de arquivo inválido. Envie arquivos Excel (.xlsx, .xls) ou CSV (.csv)', 'error')
        return redirect(url_for('index'))
    
    try:
        logger.info(f"Processando arquivo: {file.filename}")
        
        # Salva temporariamente
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            file.save(tmp_file.name)
            tmp_file_path = tmp_file.name
        
        try:
            # Lê a planilha
            df = ler_planilha_excel(tmp_file_path, file.filename)
            
            # Limpeza dos dados
            df = df.fillna("")
            # Remove aspas e espaços extras dos nomes das colunas
            df.columns = df.columns.str.strip().str.replace('"', '').str.replace("'", "")
            
            logger.info(f"Arquivo lido: {len(df)} linhas, {len(df.columns)} colunas")
            logger.info(f"Colunas encontradas (após limpeza): {list(df.columns)[:15]}")
            
            # Normaliza nomes das colunas (aceita variações como acentos, "do", etc)
            df = normalizar_colunas_df(df)
            
            logger.info(f"Colunas após normalização: {list(df.columns)[:15]}")
            
            # Valida estrutura (apenas informa, não bloqueia - NUNCA bloqueia)
            try:
                validar_planilha(df)
            except Exception as e:
                # Se por algum motivo der erro na validação, apenas loga e continua
                logger.warning(f"Validação retornou erro (mas continuando): {str(e)}")
            
            # Verifica se tem pelo menos algumas colunas básicas
            colunas_basicas = ['Nome do negócio', 'Empresa', 'Fase', 'Responsavel']
            tem_colunas_basicas = any(col in df.columns for col in colunas_basicas)
            
            if not tem_colunas_basicas:
                logger.warning("Nenhuma coluna básica encontrada, mas continuando processamento...")
                flash('Aviso: Algumas colunas esperadas não foram encontradas. O sistema continuará processando com os dados disponíveis.', 'warning')
            
            # Agrupa dados por Responsável
            relatorio_agrupado = {}
            relatorio_final = []
            linhas_processadas = 0
            linhas_com_erro = 0

            # Prepara todos os itens para processamento
            itens_para_processar = []
            
            for index, linha in df.iterrows():
                try:
                    # Monta o dicionário de dados da linha (usa valores padrão se coluna não existir)
                    # Busca colunas de forma flexível
                    def buscar_coluna(coluna_principal, alternativas=None):
                        """Busca coluna no DataFrame, tentando variações e, se necessário, posições conhecidas."""
                        # Tenta coluna principal
                        if coluna_principal in df.columns:
                            valor = linha.get(coluna_principal, '')
                            if pd.notna(valor):
                                return str(valor).strip()
                        # Tenta alternativas
                        if alternativas:
                            for alt in alternativas:
                                if alt in df.columns:
                                    valor = linha.get(alt, '')
                                    if pd.notna(valor):
                                        return str(valor).strip()
                        # Fallback: tenta usar posição baseada em nomes conhecidos
                        pos_map = {
                            'Empresa': 0,
                            'Responsavel': 3,
                            'Temperatura da Proposta Follow 1': -2,  # penúltimo campo antes do ID (ajuste conforme CSV)
                        }
                        if coluna_principal in pos_map:
                            idx = pos_map[coluna_principal]
                            if isinstance(idx, int) and abs(idx) < len(linha):
                                valor = linha.iloc[idx] if hasattr(linha, 'iloc') else linha[idx]
                                if pd.notna(valor):
                                    return str(valor).strip()
                        return ''
                    
                    item = {
                        "negocio": buscar_coluna('Nome do negócio', ['Nome do negocio', 'Negócio', 'Negocio']) or f'Negócio {index + 1}',
                        "fase": buscar_coluna('Fase') or 'Não informada',
                        "responsavel": buscar_coluna('Responsavel', ['Responsável', 'Vendedor', 'Usuario', 'Usuário']) or 'Não informado',
                        "empresa": buscar_coluna('Empresa') or 'Não informada',
                        "historico_temperaturas": {
                            "F1": buscar_coluna('Temperatura da Proposta Follow 1', ['Temperatura Follow 1', 'Temperatura 1']),
                            "F2": buscar_coluna('Temperatura da Proposta Follow 2', ['Temperatura Follow 2', 'Temperatura 2']),
                            "F3": buscar_coluna('Temperatura da Proposta Follow 3', ['Temperatura Follow 3', 'Temperatura 3']),
                            "F4": buscar_coluna('Temperatura da Proposta Follow 4', ['Temperatura Follow 4', 'Temperatura 4']),
                            "F5": buscar_coluna('Temperatura da Proposta Follow 5', ['Temperatura Follow 5', 'Temperatura 5']),
                        },
                        "historico_descricoes": {
                            "D1": buscar_coluna('Descrição Follow up 1', ['Descrição do Follow up 1', 'Descricao Follow up 1', 'Follow up 1']),
                            "D2": buscar_coluna('Descrição Follow up 2', ['Descrição do Follow up 2', 'Descricao Follow up 2', 'Follow up 2']),
                            "D3": buscar_coluna('Descrição Follow up 3', ['Descrição do Follow up 3', 'Descricao Follow up 3', 'Follow up 3']),
                            "D4": buscar_coluna('Descrição Follow up 4', ['Descrição do Follow up 4', 'Descricao Follow up 4', 'Follow up 4']),
                            "D5": buscar_coluna('Descrição Follow up 5', ['Descrição do Follow up 5', 'Descricao Follow up 5', 'Follow up 5']),
                        }
                    }
                    
                    # Pula linhas completamente vazias (apenas se não tiver nome do negócio E empresa)
                    if (not item['negocio'] or item['negocio'] == f'Negócio {index + 1}') and \
                       (not item['empresa'] or item['empresa'] == 'Não informada'):
                        logger.info(f"Pulando linha {index + 1} - sem dados básicos (negócio/empresa)")
                        continue
                    
                    itens_para_processar.append(item)
                        
                except Exception as e:
                    logger.error(f"Erro ao preparar linha {index + 1}: {str(e)}")
                    linhas_com_erro += 1
                    continue

            logger.info(f"Iniciando processamento sequencial de {len(itens_para_processar)} itens (Qualidade > Velocidade)")
            
            # Processamento sequencial (um por um) para maior robustez
            for index, item in enumerate(itens_para_processar):
                try:
                    logger.info(f"Processando item {index + 1}/{len(itens_para_processar)}: {item['negocio']}")
                    
                    # Delay entre requisições para evitar burst e garantir estabilidade
                    if index > 0:
                        time.sleep(REQUEST_DELAY)
                    
                    item_processado, erro = processar_item_thread(item)
                    
                    if erro:
                        logger.warning(f"Item {index + 1} processado com erro: {erro}")
                        linhas_com_erro += 1
                    else:
                        linhas_processadas += 1
                        logger.info(f"✅ Item {index + 1} concluído com sucesso")
                    
                    relatorio_final.append(item_processado)
                    
                    # Agrupamento por Responsável
                    responsavel = item_processado.get('responsavel') or 'Não informado'
                    if responsavel not in relatorio_agrupado:
                        relatorio_agrupado[responsavel] = []
                    relatorio_agrupado[responsavel].append(item_processado)
                    
                except Exception as e:
                    logger.error(f"Erro ao processar item {index + 1}: {str(e)}")
                    linhas_com_erro += 1

            logger.info(f"Processamento concluído: {linhas_processadas} sucessos, {linhas_com_erro} erros")
            logger.info(f"Responsáveis identificados: {list(relatorio_agrupado.keys())}")
            
            if linhas_processadas == 0:
                flash('Nenhuma linha válida encontrada na planilha', 'warning')
                return redirect(url_for('index'))
            
            # Aplica filtro de fase - mostra apenas a partir da Fase Proposta
            relatorio_final_filtrado = filtrar_negocios_por_fase(relatorio_final)
            
            # Reagrupa os dados filtrados por responsável
            relatorio_agrupado_filtrado = {}
            for item in relatorio_final_filtrado:
                responsavel = item.get('responsavel', 'Não informado')
                if responsavel not in relatorio_agrupado_filtrado:
                    relatorio_agrupado_filtrado[responsavel] = []
                relatorio_agrupado_filtrado[responsavel].append(item)
            
            # Armazena na sessão e no cache
            import uuid
            relatorio_id = str(uuid.uuid4())[:8]
            
            # Salva no cache de arquivo (agora com dados filtrados)
            dados_cache = {
                'relatorio_final': relatorio_final_filtrado,
                'relatorio_agrupado': relatorio_agrupado_filtrado,
                'responsaveis': list(relatorio_agrupado_filtrado.keys())
            }
            salvar_relatorio_cache(dados_cache, relatorio_id)
            
            # ATENÇÃO: Salvamos APENAS o ID na sessão para não estourar o limite de cookie
            session['relatorio_id_atual'] = relatorio_id
            session.permanent = True 
            
            # Removemos dados pesados da sessão que causam erro 'cookie too large'
            session.pop('relatorio_data', None)
            
            logger.info(f"Relatório armazenado com ID: {relatorio_id}")
            
            # Limpa caches antigos
            limpar_cache_antigo()
            
            return render_template('relatorio.html', relatorio=relatorio_final_filtrado, total=len(relatorio_final_filtrado), 
                           relatorio_agrupado=relatorio_agrupado_filtrado, responsaveis=list(relatorio_agrupado_filtrado.keys()))
            
        finally:
            # Remove arquivo temporário
            try:
                os.unlink(tmp_file_path)
            except:
                pass

    except ValueError as e:
        # Só bloqueia se for erro crítico (não relacionado a validação de colunas)
        error_msg = str(e)
        if "Colunas obrigatórias" in error_msg or "colunas faltando" in error_msg.lower():
            # Se for erro de colunas, apenas avisa mas continua
            logger.warning(f"Aviso de validação (continuando processamento): {error_msg}")
            flash(f'Aviso: {error_msg}. O sistema continuará processando com os dados disponíveis.', 'warning')
            # NÃO retorna redirect - continua processamento
        else:
            # Outros erros ValueError são críticos
            logger.error(f"Erro crítico: {error_msg}")
            flash(f'Erro ao processar arquivo: {error_msg}', 'error')
            return redirect(url_for('index'))
    except Exception as e:
        logger.error(f"Erro crítico ao processar a planilha: {str(e)}")
        flash(f'Erro ao processar arquivo: {str(e)}', 'error')
        return redirect(url_for('index'))


@app.route('/responsavel/<responsavel>')
def ver_responsavel(responsavel):
    """Visualização individual de análises por Responsável"""
    try:
        # Debug: Log session state
        logger.info(f"Session keys at start: {list(session.keys())}")
        logger.info(f"Requested responsável: '{responsavel}'")
        
        # Carrega dados do cache
        relatorio_final = None
        relatorio_agrupado = None
        
        # Tenta 1: Cache de arquivo
        if 'relatorio_id_atual' in session:
            relatorio_id = session['relatorio_id_atual']
            logger.info(f"Found relatorio_id_atual in session: {relatorio_id}")
            dados_cache = carregar_relatorio_cache(relatorio_id)
            if dados_cache:
                logger.info(f"Cache loaded successfully, type: {type(dados_cache)}")
                # Verifica se é o formato novo (com dicionário) ou antigo (lista direta)
                if isinstance(dados_cache, dict):
                    relatorio_final = dados_cache.get('relatorio_final', [])
                    relatorio_agrupado = dados_cache.get('relatorio_agrupado', {})
                    logger.info(f"Using new cache format - found {len(relatorio_agrupado)} responsáveis")
                else:
                    # Formato antigo - converte para novo
                    relatorio_final = dados_cache
                    relatorio_agrupado = {}
                    # Agrupa por responsável
                    for item in relatorio_final:
                        responsavel_item = item.get('responsavel', 'Não informado')
                        if responsavel_item not in relatorio_agrupado:
                            relatorio_agrupado[responsavel_item] = []
                        relatorio_agrupado[responsavel_item].append(item)
                    logger.info(f"Converted old cache format - found {len(relatorio_agrupado)} responsáveis")
            else:
                logger.warning(f"Failed to load cache for ID: {relatorio_id}")
        else:
            logger.warning("relatorio_id_atual not found in session")
        
        # Tenta 2: Sessão (backup)
        if not relatorio_agrupado and 'relatorio_data' in session:
            logger.info("Trying backup session data")
            dados_session = session['relatorio_data']
            if isinstance(dados_session, dict):
                relatorio_final = dados_session.get('relatorio_final', [])
                relatorio_agrupado = dados_session.get('relatorio_agrupado', {})
                logger.info(f"Using session backup - found {len(relatorio_agrupado)} responsáveis")
            else:
                # Formato antigo na sessão
                relatorio_final = dados_session
                relatorio_agrupado = {}
                for item in relatorio_final:
                    responsavel_item = item.get('responsavel', 'Não informado')
                    if responsavel_item not in relatorio_agrupado:
                        relatorio_agrupado[responsavel_item] = []
                    relatorio_agrupado[responsavel_item].append(item)
                logger.info(f"Converted session backup - found {len(relatorio_agrupado)} responsáveis")
        
        # Tenta 3: Busca automática no cache se não encontrou nada
        if not relatorio_agrupado:
            logger.info("Trying auto-discovery in cache files")
            try:
                for filename in os.listdir(CACHE_DIR):
                    if filename.startswith('relatorio_') and filename.endswith('.pkl'):
                        cache_file = os.path.join(CACHE_DIR, filename)
                        try:
                            with open(cache_file, 'rb') as f:
                                cache_data = pickle.load(f)
                            
                            # Verifica se o cache não é muito antigo (24 horas)
                            if datetime.now() - cache_data['timestamp'] < timedelta(hours=24):
                                dados_cache = cache_data['data']
                                if isinstance(dados_cache, dict) and 'relatorio_agrupado' in dados_cache:
                                    relatorio_agrupado = dados_cache['relatorio_agrupado']
                                    relatorio_final = dados_cache.get('relatorio_final', [])
                                    logger.info(f"Auto-discovered cache with {len(relatorio_agrupado)} responsáveis")
                                    break
                        except Exception as e:
                            logger.warning(f"Failed to read cache file {filename}: {e}")
                            continue
            except Exception as e:
                logger.error(f"Error in auto-discovery: {e}")
        
        # Debug: Log do que foi encontrado
        logger.info(f"Final result - Cache: {bool(relatorio_agrupado)}, Session: {bool('relatorio_data' in session)}")
        if relatorio_agrupado:
            logger.info(f"Responsáveis disponíveis: {list(relatorio_agrupado.keys())}")
            logger.info(f"Responsável solicitado: '{responsavel}'")
        
        if not relatorio_agrupado:
            logger.error("Dados agrupados não encontrados em nenhum lugar")
            flash('Dados não encontrados. Por favor, processe a planilha novamente.', 'error')
            return redirect(url_for('index'))
        
        # Busca dados do responsável específico
        dados_responsavel = relatorio_agrupado.get(responsavel, [])
        
        if not dados_responsavel:
            logger.warning(f"Responsável '{responsavel}' não encontrado nos dados")
            logger.info(f"Available responsáveis: {list(relatorio_agrupado.keys())}")
            flash(f'Responsável "{responsavel}" não encontrado nos dados.', 'warning')
            return redirect(url_for('index'))
        
        logger.info(f"Exibindo {len(dados_responsavel)} itens para o responsável: {responsavel}")
        
        return render_template('responsavel.html', 
                             relatorio=dados_responsavel, 
                             total=len(dados_responsavel),
                             responsavel=responsavel,
                             todos_responsaveis=list(relatorio_agrupado.keys()))
        
    except Exception as e:
        logger.error(f"Erro ao visualizar responsável: {str(e)}")
        flash(f'Erro ao carregar dados do responsável: {str(e)}', 'error')
        return redirect(url_for('index'))


@app.route('/gerar_pdf_responsavel/<responsavel>')
def gerar_pdf_responsavel(responsavel):
    """Gera PDF individual para um Responsável específico"""
    try:
        logger.info(f"Generating PDF for responsável: '{responsavel}'")
        
        # Carrega dados do cache (usando a mesma lógica de auto-discovery)
        relatorio_agrupado = None
        relatorio_final = None
        
        # Tenta 1: Cache de arquivo
        if 'relatorio_id_atual' in session:
            relatorio_id = session['relatorio_id_atual']
            logger.info(f"Found relatorio_id_atual in session: {relatorio_id}")
            dados_cache = carregar_relatorio_cache(relatorio_id)
            if dados_cache:
                logger.info(f"Cache loaded successfully for PDF")
                # Verifica se é o formato novo (com dicionário) ou antigo (lista direta)
                if isinstance(dados_cache, dict):
                    relatorio_agrupado = dados_cache.get('relatorio_agrupado', {})
                    relatorio_final = dados_cache.get('relatorio_final', [])
                else:
                    # Formato antigo - agrupa por responsável
                    relatorio_final = dados_cache
                    relatorio_agrupado = {}
                    for item in relatorio_final:
                        responsavel_item = item.get('responsavel', 'Não informado')
                        if responsavel_item not in relatorio_agrupado:
                            relatorio_agrupado[responsavel_item] = []
                        relatorio_agrupado[responsavel_item].append(item)
            else:
                logger.warning(f"Failed to load cache for PDF with ID: {relatorio_id}")
        else:
            logger.warning("relatorio_id_atual not found in session for PDF")
        
        # Tenta 2: Sessão (backup)
        if not relatorio_agrupado and 'relatorio_data' in session:
            logger.info("Trying backup session data for PDF")
            dados_session = session['relatorio_data']
            if isinstance(dados_session, dict):
                relatorio_agrupado = dados_session.get('relatorio_agrupado', {})
                relatorio_final = dados_session.get('relatorio_final', [])
            else:
                # Formato antigo na sessão
                relatorio_final = dados_session
                relatorio_agrupado = {}
                for item in relatorio_final:
                    responsavel_item = item.get('responsavel', 'Não informado')
                    if responsavel_item not in relatorio_agrupado:
                        relatorio_agrupado[responsavel_item] = []
                    relatorio_agrupado[responsavel_item].append(item)
        
        # Tenta 3: Busca automática no cache se não encontrou nada
        if not relatorio_agrupado:
            logger.info("Trying auto-discovery in cache files for PDF")
            try:
                for filename in os.listdir(CACHE_DIR):
                    if filename.startswith('relatorio_') and filename.endswith('.pkl'):
                        cache_file = os.path.join(CACHE_DIR, filename)
                        try:
                            with open(cache_file, 'rb') as f:
                                cache_data = pickle.load(f)
                            
                            # Verifica se o cache não é muito antigo (24 horas)
                            if datetime.now() - cache_data['timestamp'] < timedelta(hours=24):
                                dados_cache = cache_data['data']
                                if isinstance(dados_cache, dict) and 'relatorio_agrupado' in dados_cache:
                                    relatorio_agrupado = dados_cache['relatorio_agrupado']
                                    relatorio_final = dados_cache.get('relatorio_final', [])
                                    logger.info(f"Auto-discovered cache for PDF with {len(relatorio_agrupado)} responsáveis")
                                    break
                        except Exception as e:
                            logger.warning(f"Failed to read cache file {filename} for PDF: {e}")
                            continue
            except Exception as e:
                logger.error(f"Error in auto-discovery for PDF: {e}")
        
        if not relatorio_agrupado:
            logger.error("Dados agrupados não encontrados para PDF")
            flash('Dados não encontrados. Por favor, processe a planilha novamente.', 'error')
            return redirect(url_for('index'))
        
        # Busca dados do responsável específico
        dados_responsavel = relatorio_agrupado.get(responsavel, [])
        
        if not dados_responsavel:
            logger.warning(f"Responsável '{responsavel}' não encontrado para PDF")
            logger.info(f"Available responsáveis for PDF: {list(relatorio_agrupado.keys())}")
            flash(f'Responsável "{responsavel}" não encontrado nos dados.', 'warning')
            return redirect(url_for('processar'))
        
        total = len(dados_responsavel)
        logger.info(f"Gerando PDF individual para {responsavel}: {total} itens")
        
        # Cria buffer em memória
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=50, leftMargin=50, topMargin=50, bottomMargin=50)
        
        # Estilos
        styles = getSampleStyleSheet()
        
        # Estilo do Título Principal
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=24,
            spaceAfter=30,
            alignment=1, # Center
            textColor=VM_GREEN,
            fontName='Helvetica-Bold'
        )
        
        # Estilo para Nome do Negócio - Visual limpo e direto
        business_style = ParagraphStyle(
            'BusinessTitle',
            parent=styles['Heading2'],
            fontSize=16,
            spaceAfter=12,
            spaceBefore=20,
            textColor=VM_ORANGE,
            fontName='Helvetica-Bold',
            borderWidth=0,
            leading=20
        )
        
        # Estilo para Subtítulos (Diagnóstico, Estratégia, etc)
        section_header_style = ParagraphStyle(
            'SectionHeader',
            parent=styles['Heading3'],
            fontSize=12,
            spaceAfter=6,
            spaceBefore=12,
            textColor=VM_GREEN,
            fontName='Helvetica-Bold'
        )
        
        # Estilo Normal
        normal_style = ParagraphStyle(
            'CustomNormal',
            parent=styles['Normal'],
            fontSize=10,
            spaceAfter=6,
            leading=14,
            textColor=colors.HexColor('#2c3e50')
        )
        
        # Estilo para Labels (Empresa, Responsável, etc)
        label_style = ParagraphStyle(
            'LabelStyle',
            parent=styles['Normal'],
            fontSize=10,
            textColor=colors.HexColor('#7f8c8d'),
            leading=14
        )

        # Conteúdo do PDF
        story = []
        
        # Cabeçalho Limpo - Sem fundo verde pesado (como solicitado)
        header_data = [
            [Paragraph(f"<font color='{VM_GREEN}'>RELATÓRIO ESTRATÉGICO</font>", 
                       ParagraphStyle('HeaderTitle', parent=title_style, fontSize=22, alignment=0, spaceAfter=0)),
             Paragraph(f"<b>Vendedor:</b> {responsavel}<br/><b>Data:</b> {datetime.now().strftime('%d/%m/%Y')}", 
                       ParagraphStyle('HeaderInfo', parent=normal_style, textColor=colors.gray, alignment=2, leading=14))]
        ]
        header_table = Table(header_data, colWidths=[4.2*inch, 2.8*inch])
        header_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (0, 0), 'LEFT'),
            ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
            ('VALIGN', (0, 0), (-1, -1), 'BOTTOM'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
            ('LINEBELOW', (0, 0), (-1, -1), 1, VM_GREEN), # Linha sutil verde apenas embaixo
        ]))
        story.append(header_table)
        story.append(Spacer(1, 25))
        story.append(Spacer(1, 30))
        
        # Análises detalhadas
        for i, item in enumerate(dados_responsavel, 1):
            # Container para manter o bloco junto se possível
            elements = []
            
            # Cabeçalho do Cliente
            elements.append(Paragraph(f"{i}. {item['negocio']}", business_style))
            
            # Dados principais em tabela para organização
            data = [
                [Paragraph(f"<b>Empresa:</b> {item['empresa']}", normal_style),
                 Paragraph(f"<b>Fase:</b> {item['fase']}", normal_style)],
                [Paragraph(f"<b>Temperatura:</b> {item.get('temperatura_atual', 'Não informada')}", normal_style),
                 Paragraph(f"<b>Último Follow-up:</b> #{item.get('ultimo_follow', 0)}", normal_style)],
                [Paragraph(f"<b>Próximo Passo:</b> #{item.get('proximo_follow', 1)}", normal_style),
                 Paragraph("", normal_style)]
            ]
            
            t = Table(data, colWidths=[3.5*inch, 3.5*inch])
            t.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('TOPPADDING', (0,0), (-1,-1), 2),
                ('BOTTOMPADDING', (0,0), (-1,-1), 2),
            ]))
            elements.append(t)
            elements.append(Spacer(1, 10))
            
            # Processamento da Análise da IA
            analise_text = item.get('analise_proximo_passo', '')
            
            # Divide o texto em linhas para processar
            lines = analise_text.split('\n')
            sections_added = set() # Controle de duplicidade
            
            for line in lines:
                line_plain = line.replace('*', '').strip()
                if not line_plain: continue

                line_upper = line_plain.upper()

                if 'SITUAÇÃO' in line_upper and 'SIT' not in sections_added:
                    elements.append(Paragraph("🔍 SITUAÇÃO", section_header_style))
                    sections_added.add('SIT')
                    inline = line_plain.split(':', 1)[1].strip() if ':' in line_plain else ''
                    if inline:
                        elements.append(Paragraph(inline, normal_style))
                elif 'MENSAGEM' in line_upper and 'MSG' not in sections_added:
                    elements.append(Paragraph("💬 MENSAGEM RECOMENDADA", section_header_style))
                    sections_added.add('MSG')
                    inline = line_plain.split(':', 1)[1].strip() if ':' in line_plain else ''
                    if inline:
                        elements.append(Paragraph(inline, normal_style))
                elif ('PRÓXIMO PASSO' in line_upper or 'PRÓXIMOS PASSOS' in line_upper or 'PRÓXIMOS' in line_upper or 'META' in line_upper) and 'PROX' not in sections_added:
                    elements.append(Paragraph("🎯 PRÓXIMO PASSO & META", section_header_style))
                    sections_added.add('PROX')
                    inline = line_plain.split(':', 1)[1].strip() if ':' in line_plain else ''
                    if inline:
                        elements.append(Paragraph(inline, normal_style))
                else:
                    clean_line = line.replace('**', '').strip()
                    if clean_line.startswith('-') or clean_line.startswith('•'):
                        texto_limpo = clean_line[1:].strip()
                        if texto_limpo:
                            elements.append(Paragraph(f"• {texto_limpo}", normal_style))
                    elif clean_line:
                        elements.append(Paragraph(clean_line, normal_style))

            elements.append(Spacer(1, 15))

            # Adiciona ao story (tenta manter junto)
            story.append(KeepTogether(elements))

            # Linha divisória sutil
            if i < total:
                story.append(Spacer(1, 10))
                story.append(Table([[Spacer(1, 1)]], colWidths=[7*inch],
                                  style=[('LINEABOVE', (0,0), (-1,-1), 0.5, colors.HexColor('#e0e0e0'))]))
                story.append(Spacer(1, 20))

        # Rodapé
        story.append(Spacer(1, 30))
        story.append(Paragraph(f"Relatório individual gerado para: <b>{responsavel}</b>",
                             ParagraphStyle('Footer', parent=normal_style, alignment=1, fontSize=8, textColor=colors.gray)))
        story.append(Paragraph("Este relatório utiliza inteligência artificial para sugerir as melhores práticas comerciais da Vendamais.",
                             ParagraphStyle('Footer', parent=normal_style, alignment=1, fontSize=8, textColor=colors.gray)))
        
        # Gera o PDF
        doc.build(story)
        buffer.seek(0)
        
        # Prepara resposta
        response = make_response(buffer.getvalue())
        response.headers['Content-Type'] = 'application/pdf'
        response.headers['Content-Disposition'] = f'inline; filename=relatorio_{responsavel.replace(" ", "_")}_{datetime.now().strftime("%Y%m%d_%H%M")}.pdf'
        
        logger.info(f"PDF individual gerado com sucesso para {responsavel}: {total} itens")
        return response
        
    except Exception as e:
        logger.error(f"Erro ao gerar PDF individual: {str(e)}")
        flash(f'Erro ao gerar PDF individual: {str(e)}', 'error')
        return redirect(url_for('index'))


@app.route('/gerar_pdf')
def gerar_pdf():
    """Gera PDF profissional do relatório de análises"""
    try:
        relatorio_final = None
        
        # Debug: Log session contents
        logger.info(f"Session keys: {list(session.keys())}")
        
        # Tenta carregar do cache de arquivo primeiro (prioridade máxima)
        if 'relatorio_id_atual' in session:
            relatorio_id = session['relatorio_id_atual']
            relatorio_final = carregar_relatorio_cache(relatorio_id)
            if relatorio_final:
                logger.info(f"Relatório carregado do cache de arquivo: {len(relatorio_final)} itens")
        
        # Se não encontrou pelo ID atual, tenta outros métodos
        if not relatorio_final and 'relatorios' in session and session['relatorios']:
            # Tenta o último ID da sessão (fallback)
            ultimo_id = list(session['relatorios'].keys())[-1]
            if isinstance(session['relatorios'][ultimo_id], str):
                # Se for string, é um ID de cache
                relatorio_final = carregar_relatorio_cache(session['relatorios'][ultimo_id])
            else:
                # Se for lista diretamente (modo antigo)
                relatorio_final = session['relatorios'][ultimo_id]
            
            if relatorio_final:
                logger.info(f"Usando relatorios da sessão: {len(relatorio_final)} itens")
        
        # Se ainda não encontrou, tenta obter da sessão (apenas como último recurso)
        if not relatorio_final and 'relatorio_data' in session and session['relatorio_data']:
            relatorio_final = session['relatorio_data']
            logger.warning(f"Usando relatorio_data limitado da sessão: {len(relatorio_final)} itens (backup apenas)")
        
        # Se ainda não encontrou, tenta encontrar o cache mais recente
        if not relatorio_final:
            try:
                cache_files = [f for f in os.listdir(CACHE_DIR) if f.startswith('relatorio_') and f.endswith('.pkl')]
                if cache_files:
                    # Pega o arquivo mais recente
                    cache_files.sort(key=lambda x: os.path.getmtime(os.path.join(CACHE_DIR, x)), reverse=True)
                    latest_cache = cache_files[0]
                    relatorio_id = latest_cache.replace('relatorio_', '').replace('.pkl', '')
                    relatorio_final = carregar_relatorio_cache(relatorio_id)
                    if relatorio_final:
                        logger.info(f"Usando cache mais recente encontrado: {relatorio_id}")
            except Exception as e:
                logger.warning(f"Erro ao buscar cache mais recente: {str(e)}")
        
        if not relatorio_final:
            logger.error("Dados do relatório não encontrados na sessão nem no cache")
            logger.error(f"Session data: {dict(session)}")
            flash('Dados do relatório não encontrados. Por favor, processe a planilha novamente.', 'error')
            return redirect(url_for('index'))
        
        total = len(relatorio_final)
        logger.info(f"Gerando PDF para {total} itens")
        
        # Cria buffer em memória
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=50, leftMargin=50, topMargin=50, bottomMargin=50)
        
        # Estilos
        styles = getSampleStyleSheet()
        
        # Estilo do Título Principal
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=24,
            spaceAfter=30,
            alignment=1, # Center
            textColor=VM_GREEN,
            fontName='Helvetica-Bold'
        )
        
        # Estilo para Nome do Negócio
        business_style = ParagraphStyle(
            'BusinessTitle',
            parent=styles['Heading2'],
            fontSize=16,
            spaceAfter=10,
            spaceBefore=20,
            textColor=VM_ORANGE,
            fontName='Helvetica-Bold',
            borderPadding=5,
            borderColor=VM_GREEN,
            borderWidth=0,
            backColor=colors.Color(0.95, 0.95, 0.95) # Fundo cinza claro
        )
        
        # Estilo para Subtítulos (Diagnóstico, Estratégia, etc)
        section_header_style = ParagraphStyle(
            'SectionHeader',
            parent=styles['Heading3'],
            fontSize=12,
            spaceAfter=6,
            spaceBefore=12,
            textColor=VM_GREEN,
            fontName='Helvetica-Bold'
        )
        
        # Estilo Normal
        normal_style = ParagraphStyle(
            'CustomNormal',
            parent=styles['Normal'],
            fontSize=10,
            spaceAfter=6,
            leading=14,
            textColor=colors.HexColor('#2c3e50')
        )
        
        # Estilo para Labels (Empresa, Responsável, etc)
        label_style = ParagraphStyle(
            'LabelStyle',
            parent=styles['Normal'],
            fontSize=10,
            textColor=colors.HexColor('#7f8c8d'),
            leading=14
        )

        # Conteúdo do PDF
        story = []
        
        # Cabeçalho Limpo - Sem fundo verde pesado
        header_data = [
            [Paragraph(f"<font color='{VM_GREEN}'>RELATÓRIO ESTRATÉGICO GERAL</font>", 
                       ParagraphStyle('HeaderTitle', parent=title_style, fontSize=20, alignment=0, spaceAfter=0)),
             Paragraph(f"<b>Data:</b> {datetime.now().strftime('%d/%m/%Y')}<br/><b>Total:</b> {total} análises", 
                       ParagraphStyle('HeaderInfo', parent=normal_style, textColor=colors.gray, alignment=2, leading=14))]
        ]
        header_table = Table(header_data, colWidths=[4.2*inch, 2.8*inch])
        header_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (0, 0), 'LEFT'),
            ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
            ('VALIGN', (0, 0), (-1, -1), 'BOTTOM'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
            ('LINEBELOW', (0, 0), (-1, -1), 1, VM_GREEN), # Linha sutil verde apenas embaixo
        ]))
        story.append(header_table)
        story.append(Spacer(1, 25))
        story.append(Spacer(1, 30))
        
        # Análises detalhadas
        for i, item in enumerate(relatorio_final, 1):
            # Container para manter o bloco junto se possível
            elements = []
            
            # Cabeçalho do Cliente
            elements.append(Paragraph(f"{i}. {item['negocio']}", business_style))
            
            # Dados principais em tabela para organização
            data = [
                [Paragraph(f"<b>Empresa:</b> {item['empresa']}", normal_style),
                 Paragraph(f"<b>Responsável:</b> {item['responsavel']}", normal_style)],
                [Paragraph(f"<b>Fase:</b> {item['fase']}", normal_style),
                 Paragraph(f"<b>Temperatura:</b> {item.get('temperatura_atual', 'Não informada')}", normal_style)],
                [Paragraph(f"<b>Último Follow-up:</b> #{item.get('ultimo_follow', 0)}", normal_style),
                 Paragraph(f"<b>Próximo Passo:</b> #{item.get('proximo_follow', 1)}", normal_style)]
            ]
            
            t = Table(data, colWidths=[3.5*inch, 3.5*inch])
            t.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('TOPPADDING', (0,0), (-1,-1), 2),
                ('BOTTOMPADDING', (0,0), (-1,-1), 2),
            ]))
            elements.append(t)
            elements.append(Spacer(1, 10))
            
            # Processamento da Análise da IA
            analise_text = item.get('analise_proximo_passo', '')
            
            # Divide o texto em linhas para processar
            lines = analise_text.split('\n')
            sections_added = set() # Controle de duplicidade
            
            for line in lines:
                line_plain = line.replace('*', '').strip()
                if not line_plain: continue

                line_upper = line_plain.upper()

                if 'SITUAÇÃO' in line_upper and 'SIT' not in sections_added:
                    elements.append(Paragraph("🔍 SITUAÇÃO", section_header_style))
                    sections_added.add('SIT')
                    inline = line_plain.split(':', 1)[1].strip() if ':' in line_plain else ''
                    if inline:
                        elements.append(Paragraph(inline, normal_style))
                elif 'MENSAGEM' in line_upper and 'MSG' not in sections_added:
                    elements.append(Paragraph("💬 MENSAGEM RECOMENDADA", section_header_style))
                    sections_added.add('MSG')
                    inline = line_plain.split(':', 1)[1].strip() if ':' in line_plain else ''
                    if inline:
                        elements.append(Paragraph(inline, normal_style))
                elif ('PRÓXIMO PASSO' in line_upper or 'PRÓXIMOS PASSOS' in line_upper or 'PRÓXIMOS' in line_upper or 'META' in line_upper) and 'PROX' not in sections_added:
                    elements.append(Paragraph("🎯 PRÓXIMO PASSO & META", section_header_style))
                    sections_added.add('PROX')
                    inline = line_plain.split(':', 1)[1].strip() if ':' in line_plain else ''
                    if inline:
                        elements.append(Paragraph(inline, normal_style))
                else:
                    clean_line = line.replace('**', '').strip()
                    if clean_line.startswith('-') or clean_line.startswith('•'):
                        texto_limpo = clean_line[1:].strip()
                        if texto_limpo:
                            elements.append(Paragraph(f"• {texto_limpo}", normal_style))
                    elif clean_line:
                        elements.append(Paragraph(clean_line, normal_style))

            elements.append(Spacer(1, 15))

            # Adiciona ao story (tenta manter junto)
            story.append(KeepTogether(elements))

            # Linha divisória sutil
            if i < total:
                story.append(Spacer(1, 10))
                story.append(Table([[Spacer(1, 1)]], colWidths=[7*inch],
                                  style=[('LINEABOVE', (0,0), (-1,-1), 0.5, colors.HexColor('#e0e0e0'))]))
                story.append(Spacer(1, 20))

        # Rodapé
        story.append(Spacer(1, 30))
        story.append(Paragraph("Este relatório utiliza inteligência artificial para sugerir as melhores práticas comerciais da Vendamais.",
                             ParagraphStyle('Footer', parent=normal_style, alignment=1, fontSize=8, textColor=colors.gray)))
        
        # Gera o PDF
        doc.build(story)
        buffer.seek(0)
        
        # Prepara resposta
        response = make_response(buffer.getvalue())
        response.headers['Content-Type'] = 'application/pdf'
        response.headers['Content-Disposition'] = f'inline; filename=relatorio_estrategico_{datetime.now().strftime("%Y%m%d_%H%M")}.pdf'
        
        logger.info(f"PDF gerado com sucesso: {total} itens")
        return response
        
    except Exception as e:
        logger.error(f"Erro ao gerar PDF: {str(e)}")
        flash(f'Erro ao gerar PDF: {str(e)}', 'error')
        return redirect(url_for('index'))


@app.route('/debug_session')
def debug_session():
    """Debug route to check session status"""
    session_info = {
        'keys': list(session.keys()),
        'relatorio_data_exists': 'relatorio_data' in session,
        'relatorios_exists': 'relatorios' in session,
        'relatorio_data_length': len(session.get('relatorio_data', [])) if 'relatorio_data' in session else 0,
        'relatorios_count': len(session.get('relatorios', {})) if 'relatorios' in session else 0,
    }
    return session_info


@app.route('/limpar_sessao')
def limpar_sessao():
    """Limpa dados da sessão para permitir novos processamentos"""
    try:
        keys_to_clear = ['relatorio_data', 'relatorios', 'relatorio_id_atual']
        for key in keys_to_clear:
            if key in session:
                session.pop(key, None)
        
        logger.info("Sessão de relatórios limpa com sucesso")
        flash('Sessão limpa. Você pode processar novas planilhas agora.', 'success')
        
    except Exception as e:
        logger.error(f"Erro ao limpar sessão: {str(e)}")
        flash(f'Erro ao limpar sessão: {str(e)}', 'error')
    
    return redirect(url_for('index'))


# ==================== INICIALIZAÇÃO ====================

if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    debug = os.getenv('FLASK_DEBUG', 'True').lower() == 'true'
    
    logger.info(f"Iniciando servidor Flask na porta {port} (debug={debug})")
    logger.info(f"Usando API Gemini com modelo: {GEMINI_MODEL}")
    app.run(debug=debug, port=port)
