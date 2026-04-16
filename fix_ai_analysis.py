
# Processamento da Análise da IA - VERSÃO MELHORADA
analise_text = item.get('analise_proximo_passo', '')

# Divide o texto em linhas para processar
lines = analise_text.split('\n')
sections_added = set() # Controle de duplicidade
content_buffer = []

# Função para adicionar conteúdo buffer ao story
def flush_content_buffer():
    global content_buffer
    if content_buffer:
        for content_line in content_buffer:
            clean_line = content_line.replace('**', '').strip()
            if clean_line.startswith('-') or clean_line.startswith(''):
                texto_limpo = clean_line[1:].strip()
                if texto_limpo:
                    elements.append(Paragraph(f" {texto_limpo}", normal_style))
            else:
                elements.append(Paragraph(clean_line, normal_style))
        content_buffer = []

# Função para adicionar seção se não existir
def ensure_section_exists(section_key, header_text):
    if section_key not in sections_added:
        elements.append(Paragraph(header_text, section_header_style))
        sections_added.add(section_key)
        return True
    return False

# Análise melhorada para detectar seções mesmo que incompletas
for line in lines:
    line_plain = line.replace('*', '').strip()
    if not line_plain: 
        continue
    
    line_upper = line_plain.upper()
    
    # Detecta cabeçalhos de seção com padrões mais flexíveis
    is_situation = any(kw in line_upper for kw in [
        'SITUAÇÃO', 'SITUACAO', 'SITUAÇÃO:', '1.', '**SITUAÇÃO**', 'SITUACAO:'
    ])
    is_message = any(kw in line_upper for kw in [
        'MENSAGEM', 'MENSAGEM RECOMENDADA', 'MENSAGEM RECOMENDADA:', '2.', 
        '**MENSAGEM**', '**MENSAGEM RECOMENDADA**', 'MENSAGEM:'
    ])
    is_next_step = any(kw in line_upper for kw in [
        'PRÓXIMO PASSO', 'PROXIMO PASSO', 'META', 'PRÓXIMO PASSO:', '3.', 
        '**PRÓXIMO PASSO**', '**META**', 'PROXIMO PASSO:', 'META:'
    ])
    
    if is_situation:
        flush_content_buffer()
        if ensure_section_exists('SIT', "SITUAÇÃO"):
            current_section = 'SIT'
    elif is_message:
        flush_content_buffer()
        if ensure_section_exists('MSG', "MENSAGEM RECOMENDADA"):
            current_section = 'MSG'
    elif is_next_step:
        flush_content_buffer()
        if ensure_section_exists('PROX', "PRÓXIMO PASSO & META"):
            current_section = 'PROX'
    elif any(kw in line_upper for kw in ['SITUAÇÃO', 'MENSAGEM', 'PRÓXIMO PASSO', 'META']):
        # Se já adicionou o cabeçalho, ignora a linha que repete o nome da seção
        continue
    else:
        # Conteúdo normal - adiciona ao buffer
        content_buffer.append(line)

# Flush final do conteúdo
flush_content_buffer()

# Se não encontrou nenhuma seção estruturada, trata como conteúdo único
if not sections_added and analise_text.strip():
    elements.append(Paragraph("ANÁLISE ESTRATÉGICA", section_header_style))
    for line in lines:
        clean_line = line.replace('**', '').strip()
        if clean_line:
            if clean_line.startswith('-') or clean_line.startswith(''):
                texto_limpo = clean_line[1:].strip()
                if texto_limpo:
                    elements.append(Paragraph(f" {texto_limpo}", normal_style))
            else:
                elements.append(Paragraph(clean_line, normal_style))

print("""
INSTRUÇÕES PARA APLICAR O FIX:

1. Abra o arquivo app.py
2. Localize as linhas 1641-1675 (seção "Processamento da Análise da IA")
3. Substitua todo esse bloco pelo código acima
4. Salve o arquivo

O que este fix resolve:
- Detecta seções mesmo que a IA não siga o formato exato
- Aceita variações como "SITUAÇÃO:", "1.", "**SITUAÇÃO**", etc.
- Garante que todas as seções apareçam no PDF
- Trata respostas sem estrutura como conteúdo único
- Melhora a consistência da análise no PDF

Este problema acontecia porque a IA às vezes não seguia exatamente o formato solicitado,
e o parser original só detectava os cabeçalhos exatos.
""")
