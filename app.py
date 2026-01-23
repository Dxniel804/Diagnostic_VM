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

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

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
REQUEST_DELAY = float(os.getenv('REQUEST_DELAY', '2'))

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

    # Tenta até o limite configurado caso a API esteja ocupada
    for tentativa in range(MAX_RETRIES):
        try:
            response = client.chat.completions.create(
                model=GROQ_MODEL,
                messages=[{"role": "user", "content": prompt}],
                max_tokens=800,
                temperature=0.7
            )
            
            resultado = response.choices[0].message.content
            
            # Salva no cache
            cache_analises[hash_cache] = resultado
            
            logger.info(f"Análise gerada com sucesso para {dados_negocio['negocio']} (tentativa {tentativa + 1})")
            return resultado
            
        except Exception as e:
            error_msg = str(e).lower()
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


def validar_planilha(df):
    """Valida se a planilha possui as colunas obrigatórias"""
    colunas_obrigatorias = [
        'Nome do negócio', 'Empresa', 'Fase', 'Responsavel',
        'Temperatura da Proposta Follow 1', 'Descrição Follow up 1'
    ]

    colunas_faltantes = []
    for coluna in colunas_obrigatorias:
        if coluna not in df.columns:
            colunas_faltantes.append(coluna)

    if colunas_faltantes:
        raise ValueError(f"Colunas obrigatórias faltando: {', '.join(colunas_faltantes)}")

    return True


def ler_planilha_excel(file_path, filename):
    """
    Lê arquivo Excel com múltiplas estratégias de fallback.
    Suporta .xlsx, .xls e até arquivos HTML disfarçados de Excel.
    """
    file_ext = filename.lower().split('.')[-1] if '.' in filename else ''
    logger.info(f"Processando arquivo.{file_ext}: {filename}")
    
    df = None
    error_messages = []
    
    # Verifica se é HTML disfarçado de Excel
    with open(file_path, 'rb') as f:
        header = f.read(50)
        is_html = (
            header.startswith(b'\xef\xbb\xbf<meta') or 
            header.startswith(b'<meta') or 
            header.startswith(b'<!DOCTYPE') or 
            header.startswith(b'<html') or
            b'<table' in header
        )
    
    if is_html:
        logger.warning("Conteúdo HTML detectado, tentando converter...")
        try:
            df_html = pd.read_html(file_path)
            if df_html and len(df_html) > 0:
                df = df_html[0]
                logger.info(f"HTML convertido: {len(df)} linhas, {len(df.columns)} colunas")
        except Exception as e:
            logger.warning(f"Conversão HTML falhou: {str(e)}")
    
    # Se não é HTML ou conversão falhou, tenta leitura normal
    if df is None:
        if file_ext == 'xls':
            try:
                df = pd.read_excel(file_path, engine='xlrd')
                logger.info("Arquivo .xls lido com xlrd")
            except Exception as e1:
                logger.warning(f"xlrd falhou: {str(e1)}")
                try:
                    df = pd.read_excel(file_path, engine='openpyxl')
                    logger.info("Arquivo .xls lido com openpyxl (fallback)")
                except Exception as e2:
                    error_messages.append(f"xlrd({str(e1)}), openpyxl({str(e2)})")
        elif file_ext == 'xlsx':
            try:
                df = pd.read_excel(file_path, engine='openpyxl')
                logger.info("Arquivo .xlsx lido com openpyxl")
            except Exception as e1:
                logger.warning(f"openpyxl falhou: {str(e1)}")
                try:
                    df = pd.read_excel(file_path, engine='xlrd')
                    logger.info("Arquivo .xlsx lido com xlrd (fallback)")
                except Exception as e2:
                    error_messages.append(f"openpyxl({str(e1)}), xlrd({str(e2)})")
        
        # Última tentativa sem engine específica
        if df is None:
            try:
                df = pd.read_excel(file_path)
                logger.info("Arquivo lido sem engine específica")
            except Exception as e3:
                error_messages.append(f"default({str(e3)})")
    
    if df is None:
        all_errors = ", ".join(error_messages) if error_messages else "Erro desconhecido"
        raise ValueError(f"Não foi possível ler o arquivo Excel. Erros: {all_errors}")
    
    if df.empty:
        raise ValueError("O arquivo está vazio")
    
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
        flash('Formato de arquivo inválido. Envie apenas arquivos Excel (.xlsx, .xls)', 'error')
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
            df.columns = df.columns.str.strip()
            
            logger.info(f"Arquivo lido: {len(df)} linhas, {len(df.columns)} colunas")
            logger.info(f"Colunas encontradas: {list(df.columns)}")
            
            # Valida estrutura
            validar_planilha(df)
            
            # Processa cada linha
            relatorio_final = []
            linhas_processadas = 0
            linhas_com_erro = 0

            for index, linha in df.iterrows():
                try:
                    # Monta o dicionário de dados da linha
                    item = {
                        "negocio": str(linha.get('Nome do negócio', 'N/A')).strip(),
                        "fase": str(linha.get('Fase', 'N/A')).strip(),
                        "responsavel": str(linha.get('Responsavel', 'N/A')).strip(),
                        "empresa": str(linha.get('Empresa', 'N/A')).strip(),
                        "historico_temperaturas": {
                            "F1": str(linha.get('Temperatura da Proposta Follow 1', '')).strip(),
                            "F2": str(linha.get('Temperatura da Proposta Follow 2', '')).strip(),
                            "F3": str(linha.get('Temperatura da Proposta Follow 3', '')).strip(),
                            "F4": str(linha.get('Temperatura da Proposta Follow 4', '')).strip(),
                            "F5": str(linha.get('Temperatura da Proposta Follow 5', '')).strip(),
                        },
                        "historico_descricoes": {
                            "D1": str(linha.get('Descrição Follow up 1', '')).strip(),
                            "D2": str(linha.get('Descrição Follow up 2', '')).strip(),
                            "D3": str(linha.get('Descrição Follow up 3', '')).strip(),
                            "D4": str(linha.get('Descrição Follow up 4', '')).strip(),
                            "D5": str(linha.get('Descrição Follow up 5', '')).strip(),
                        }
                    }
                    
                    # Pula linhas completamente vazias
                    if (item['negocio'] == 'N/A' and item['empresa'] == 'N/A' and 
                        item['historico_descricoes']['D1'] == ''):
                        logger.info(f"Pulando linha {index + 1} - dados vazios")
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
        logger.error(f"Erro de validação: {str(e)}")
        flash(f'Erro na validação da planilha: {str(e)}', 'error')
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
