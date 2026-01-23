"""
Sistema de Análise Estratégica de Follow-ups de Vendas
Automação que lê planilhas Excel e gera estratégias personalizadas usando IA
"""

import os
import time
import logging
import hashlib
import pandas as pd
from flask import Flask, render_template, request, flash, redirect, url_for, session, make_response
from groq import Groq
from dotenv import load_dotenv
import tempfile
from werkzeug.utils import secure_filename
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from datetime import datetime
import io

# Carrega variáveis de ambiente
load_dotenv()

# ==================== CONFIGURAÇÃO ====================
app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', os.urandom(24).hex())
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max

ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('app.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Configurações da API Groq (GRATUITA - Recomendada)
GROQ_API_KEY = os.getenv('GROQ_API_KEY')
GROQ_MODEL = os.getenv('GROQ_MODEL', 'llama-3.3-70b-versatile')
MAX_RETRIES = int(os.getenv('MAX_RETRIES', '3'))
RETRY_DELAY = int(os.getenv('RETRY_DELAY', '5'))
REQUEST_DELAY = float(os.getenv('REQUEST_DELAY', '10'))

# Cache para evitar requisições duplicadas
cache_analises = {}

if not GROQ_API_KEY:
    logger.error("GROQ_API_KEY não encontrada nas variáveis de ambiente")
    raise ValueError("GROQ_API_KEY é obrigatória. Configure no arquivo .env")

# Inicializa o cliente Groq
try:
    client = Groq(api_key=GROQ_API_KEY)
    logger.info(f"Cliente Groq configurado com sucesso usando modelo: {GROQ_MODEL}")
except Exception as e:
    logger.error(f"Erro ao configurar cliente Groq: {str(e)}")
    raise ValueError("Não foi possível configurar o cliente Groq. Verifique sua API key.")


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
    Envia o contexto do negócio para a IA Groq e recebe a estratégia de venda.
    A IA age como um Diretor Comercial experiente.
    """
    # Verifica cache primeiro
    hash_cache = gerar_hash_cache(dados_negocio)
    if hash_cache in cache_analises:
        logger.info(f"Retornando análise em cache para {dados_negocio['negocio']}")
        return cache_analises[hash_cache]

    # Identifica onde a conversa parou
    ultimo_follow, proximo_follow, temperatura_atual = identificar_ultimo_followup(dados_negocio)
    
    # Monta histórico relevante (apenas os follow-ups preenchidos)
    historico_texto = ""
    for i in range(1, ultimo_follow + 1):
        desc = dados_negocio['historico_descricoes'][f'D{i}'].strip()
        temp = dados_negocio['historico_temperaturas'][f'F{i}'].strip()
        if desc:
            historico_texto += f"Follow-up {i} (Temperatura: {temp or 'Não informada'}): {desc}\n"
    
    prompt = f"""Você é um Diretor Comercial experiente com anos de experiência em fechamento de vendas.

ANÁLISE DO NEGÓCIO:
- Nome do Negócio: {dados_negocio['negocio']}
- Empresa Cliente: {dados_negocio['empresa']}
- Responsável: {dados_negocio['responsavel']}
- Fase Atual: {dados_negocio['fase']}
- Último Follow-up Realizado: #{ultimo_follow}
- Próximo Follow-up a Realizar: #{proximo_follow}
- Temperatura Atual: {temperatura_atual}

HISTÓRICO DE CONVERSAS:
{historico_texto if historico_texto else 'Nenhum follow-up realizado ainda.'}

SUA MISSÃO:
Analise a situação e forneça uma orientação estratégica PRÁTICA e DIRETA para o Follow-up #{proximo_follow}.

A resposta DEVE conter exatamente estas 3 seções:

1. **DIAGNÓSTICO DA SITUAÇÃO:**
   - Identifique claramente a temperatura atual (QUENTE/MORNO/FRIO)
   - Analise o que aconteceu até agora
   - Identifique objeções, pontos de atenção ou oportunidades

2. **ESTRATÉGIA PARA O PRÓXIMO PASSO:**
   - O que dizer exatamente no próximo contato (mensagem direta)
   - Argumentos de fechamento específicos para esta situação
   - Gatilhos mentais ou técnicas de persuasão adequadas

3. **AÇÃO RECOMENDADA:**
   - Pergunta de fechamento específica
   - Próximo passo concreto para avançar na venda
   - Prazo sugerido para o follow-up

Seja DIRETO, PRÁTICO e FOQUE EM FECHAR A VENDA. Não seja genérico."""

    logger.info(f"Processando negócio: {dados_negocio['negocio']} - Empresa: {dados_negocio['empresa']} - Próximo Follow-up: #{proximo_follow}")

    # Lista de modelos válidos (em ordem de preferência)
    modelos_validos = [
        'llama-3.3-70b-versatile',  # Modelo atual recomendado
        'llama-3.1-8b-instruct',   # Fallback rápido
        'mixtral-8x7b-32768',       # Alternativa Mixtral
        'gemma2-9b-it'              # Alternativa Gemma
    ]
    
    modelo_usar = GROQ_MODEL if GROQ_MODEL in modelos_validos else modelos_validos[0]
    
    # Tenta até o limite configurado caso a API esteja ocupada
    for tentativa in range(MAX_RETRIES):
        try:
            response = client.chat.completions.create(
                model=modelo_usar,
                messages=[{"role": "user", "content": prompt}],
                max_tokens=800,
                temperature=0.7
            )
            
            resultado = response.choices[0].message.content
            
            # Salva no cache
            cache_analises[hash_cache] = resultado
            
            logger.info(f"Análise gerada com sucesso para {dados_negocio['negocio']} usando modelo {modelo_usar} (tentativa {tentativa + 1})")
            return resultado
            
        except Exception as e:
            error_msg = str(e).lower()
            
            # Se o modelo foi descontinuado, tenta outro modelo
            if "decommissioned" in error_msg or "no longer supported" in error_msg or "model_decommissioned" in error_msg:
                logger.warning(f"Modelo {modelo_usar} foi descontinuado. Tentando modelo alternativo...")
                # Tenta próximo modelo da lista
                idx_atual = modelos_validos.index(modelo_usar) if modelo_usar in modelos_validos else 0
                if idx_atual < len(modelos_validos) - 1:
                    modelo_usar = modelos_validos[idx_atual + 1]
                    logger.info(f"Tentando com modelo alternativo: {modelo_usar}")
                    continue
                else:
                    logger.error(f"Todos os modelos testados foram descontinuados")
                    return "Erro: Modelo de IA descontinuado. Por favor, atualize o GROQ_MODEL no arquivo .env para 'llama-3.3-70b-versatile'"
            
            if "rate" in error_msg or "limit" in error_msg or "too many" in error_msg:
                logger.warning(f"Limite de cota Groq atingido. Tentativa {tentativa + 1}/{MAX_RETRIES}")
                if tentativa < MAX_RETRIES - 1:
                    time.sleep(RETRY_DELAY)
                continue
            else:
                logger.error(f"Erro na análise do negócio {dados_negocio['negocio']}: {str(e)}")
                return f"Erro na análise desta linha: {str(e)}"

    logger.error(f"Não foi possível gerar análise para {dados_negocio['negocio']} (limite de tentativas excedido)")
    return "Não foi possível gerar a análise para este item (limite de tentativas excedido)."


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
        logger.info(f"Colunas normalizadas ({len(mapeamento)} colunas): {list(mapeamento.items())[:5]}")
    else:
        logger.warning("Nenhuma coluna foi normalizada. Verifique se os nomes das colunas estão corretos.")
    
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
    """Página inicial com formulário de upload"""
    return render_template('index.html')


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
            
            # Processa cada linha
            relatorio_final = []
            linhas_processadas = 0
            linhas_com_erro = 0

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
                    
                    # Pula linhas completamente vazias (mas é mais flexível agora)
                    if (not item['negocio'] or item['negocio'] == f'Negócio {index + 1}') and \
                       (not item['empresa'] or item['empresa'] == 'Não informada') and \
                       not any(item['historico_descricoes'].values()):
                        logger.info(f"Pulando linha {index + 1} - dados completamente vazios")
                        continue
                    
                    # Identifica follow-ups para exibição
                    ultimo_follow, proximo_follow, temperatura_atual = identificar_ultimo_followup(item)
                    item["ultimo_follow"] = ultimo_follow
                    item["proximo_follow"] = proximo_follow
                    item["temperatura_atual"] = temperatura_atual
                    
                    # Chama a IA para análise estratégica
                    item["analise_proximo_passo"] = pedir_estrategia_ia(item)
                    
                    # Pausa para não sobrecarregar a API
                    time.sleep(REQUEST_DELAY)
                    
                    relatorio_final.append(item)
                    linhas_processadas += 1
                    
                    # Progress log
                    if (index + 1) % 10 == 0:
                        logger.info(f"Progresso: {index + 1}/{len(df)} linhas processadas")
                        
                except Exception as e:
                    logger.error(f"Erro ao processar linha {index + 1}: {str(e)}")
                    linhas_com_erro += 1
                    continue

            logger.info(f"Processamento concluído: {linhas_processadas} sucessos, {linhas_com_erro} erros")
            
            if linhas_processadas == 0:
                flash('Nenhuma linha válida encontrada na planilha', 'warning')
                return redirect(url_for('index'))
            
            # Armazena na sessão
            import uuid
            relatorio_id = str(uuid.uuid4())[:8]
            
            if 'relatorios' not in session:
                session['relatorios'] = {}
            session['relatorios'][relatorio_id] = relatorio_final
            session['relatorio_id_atual'] = relatorio_id
            session['relatorio_data'] = relatorio_final
            
            logger.info(f"Relatório armazenado com ID: {relatorio_id}")
            
            return render_template('relatorio.html', relatorio=relatorio_final, total=len(relatorio_final))
            
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


@app.route('/gerar_pdf')
def gerar_pdf():
    """Gera PDF profissional do relatório de análises"""
    try:
        relatorio_final = None
        
        if 'relatorio_data' in session and session['relatorio_data']:
            relatorio_final = session['relatorio_data']
        elif 'relatorios' in session and session['relatorios']:
            ultimo_id = list(session['relatorios'].keys())[-1]
            relatorio_final = session['relatorios'][ultimo_id]
        
        if not relatorio_final:
            logger.error("Dados do relatório não encontrados na sessão")
            flash('Dados do relatório não encontrados. Por favor, processe a planilha novamente.', 'error')
            return redirect(url_for('index'))
        
        total = len(relatorio_final)
        logger.info(f"Gerando PDF para {total} itens")
        
        # Cria buffer em memória
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=72, leftMargin=72, topMargin=72, bottomMargin=18)
        
        # Estilos
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=20,
            spaceAfter=30,
            alignment=1,
            textColor=colors.HexColor('#2c3e50')
        )
        
        heading_style = ParagraphStyle(
            'CustomHeading',
            parent=styles['Heading2'],
            fontSize=14,
            spaceAfter=12,
            textColor=colors.HexColor('#34495e')
        )
        
        normal_style = ParagraphStyle(
            'CustomNormal',
            parent=styles['Normal'],
            fontSize=10,
            spaceAfter=6,
            leading=14
        )
        
        # Conteúdo do PDF
        story = []
        
        # Título
        story.append(Paragraph("Relatório de Análise Estratégica de CRM", title_style))
        story.append(Spacer(1, 20))
        
        # Data e resumo
        data_atual = datetime.now().strftime("%d/%m/%Y %H:%M")
        story.append(Paragraph(f"<b>Data:</b> {data_atual}", normal_style))
        story.append(Paragraph(f"<b>Total de Negócios Analisados:</b> {total}", normal_style))
        story.append(Spacer(1, 20))
        
        # Análises detalhadas
        for i, item in enumerate(relatorio_final, 1):
            # Cabeçalho do Cliente
            story.append(Paragraph(f"<b>{i}. {item['negocio']}</b>", heading_style))
            story.append(Paragraph(f"<b>Empresa:</b> {item['empresa']}", normal_style))
            story.append(Paragraph(f"<b>Responsável:</b> {item['responsavel']}", normal_style))
            
            # Status Atual
            story.append(Paragraph(f"<b>Fase:</b> {item['fase']}", normal_style))
            story.append(Paragraph(f"<b>Temperatura Atual:</b> {item.get('temperatura_atual', 'Não informada')}", normal_style))
            
            # Follow-up
            ultimo = item.get('ultimo_follow', 0)
            proximo = item.get('proximo_follow', 1)
            if ultimo > 0:
                story.append(Paragraph(f"<b>Último Follow-up Realizado:</b> #{ultimo}", normal_style))
            story.append(Paragraph(f"<b>Próximo Follow-up:</b> #{proximo}", normal_style))
            story.append(Spacer(1, 10))
            
            # Plano de Ação (IA)
            story.append(Paragraph("<b>Plano de Ação Estratégico (IA):</b>", normal_style))
            analise_text = item.get('analise_proximo_passo', 'Análise não disponível')
            # Limita tamanho para não quebrar o PDF
            if len(analise_text) > 1500:
                analise_text = analise_text[:1500] + '...'
            story.append(Paragraph(analise_text, normal_style))
            story.append(Spacer(1, 20))
            
            # Quebra de página entre empresas (exceto na última)
            if i < total:
                story.append(PageBreak())
        
        # Rodapé
        story.append(Spacer(1, 20))
        story.append(Paragraph("<b>Relatório gerado por:</b> Sistema de Automação de Vendas com IA", normal_style))
        story.append(Paragraph(f"<b>Emissão:</b> {data_atual}", normal_style))
        
        # Gera o PDF
        doc.build(story)
        buffer.seek(0)
        
        # Prepara resposta
        response = make_response(buffer.getvalue())
        response.headers['Content-Type'] = 'application/pdf'
        response.headers['Content-Disposition'] = f'inline; filename=relatorio_analise_{datetime.now().strftime("%Y%m%d_%H%M")}.pdf'
        
        logger.info(f"PDF gerado com sucesso: {total} itens")
        return response
        
    except Exception as e:
        logger.error(f"Erro ao gerar PDF: {str(e)}")
        flash(f'Erro ao gerar PDF: {str(e)}', 'error')
        return redirect(url_for('index'))


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
    logger.info(f"Usando API Groq com modelo: {GROQ_MODEL}")
    app.run(debug=debug, port=port)
