import re

# Read the original file
with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Define the old section to replace - with exact emojis
old_section = '''            # Processamento da Análise da IA
            analise_text = item.get('analise_proximo_passo', '')
            
            # Divide o texto em linhas para processar
            lines = analise_text.split('\n')
            sections_added = set() # Controle de duplicidade
            
            for line in lines:
                line_plain = line.replace('*', '').strip()
                if not line_plain: continue
                
                line_upper = line_plain.upper()
                
                # Verifica se é um cabeçalho de seção - Lógica mais rigorosa para evitar duplicados
                if 'SITUAÇÃO' in line_upper and 'SIT' not in sections_added:
                    elements.append(Paragraph("SITUAÇÃO", section_header_style))
                    sections_added.add('SIT')
                elif 'MENSAGEM' in line_upper and 'MSG' not in sections_added:
                    elements.append(Paragraph("MENSAGEM RECOMENDADA", section_header_style))
                    sections_added.add('MSG')
                elif ('PRÓXIMO PASSO' in line_upper or 'META' in line_upper) and 'PROX' not in sections_added:
                    elements.append(Paragraph("PRÓXIMO PASSO & META", section_header_style))
                    sections_added.add('PROX')
                elif any(kw in line_upper for kw in ['SITUAÇÃO', 'MENSAGEM', 'PRÓXIMO PASSO', 'META']):
                    # Se já adicionou o cabeçalho, ignora a linha que repete o nome da seção
                    continue
                else:
                    # Conteúdo normal
                    clean_line = line.replace('**', '').strip()
                    if clean_line.startswith('-') or clean_line.startswith(''):
                        texto_limpo = clean_line[1:].strip()
                        if texto_limpo:
                            elements.append(Paragraph(f" {texto_limpo}", normal_style))
                    else:
                        elements.append(Paragraph(clean_line, normal_style))'''

# Define the new improved section
new_section = '''            # Processamento da Análise da IA - VERSÃO MELHORADA
            analise_text = item.get('analise_proximo_passo', '')
            
            # Divide o texto em linhas para processar
            lines = analise_text.split('\n')
            sections_added = set() # Controle de duplicidade
            content_buffer = []
            
            # Função para adicionar conteúdo buffer ao story
            def flush_content_buffer():
                nonlocal content_buffer
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
                
                if is_situation and 'SIT' not in sections_added:
                    flush_content_buffer()
                    elements.append(Paragraph("SITUAÇÃO", section_header_style))
                    sections_added.add('SIT')
                elif is_message and 'MSG' not in sections_added:
                    flush_content_buffer()
                    elements.append(Paragraph("MENSAGEM RECOMENDADA", section_header_style))
                    sections_added.add('MSG')
                elif is_next_step and 'PROX' not in sections_added:
                    flush_content_buffer()
                    elements.append(Paragraph("PRÓXIMO PASSO & META", section_header_style))
                    sections_added.add('PROX')
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
                            elements.append(Paragraph(clean_line, normal_style))'''

# Replace the section
if old_section in content:
    content = content.replace(old_section, new_section)
    
    # Write back to file
    with open('app.py', 'w', encoding='utf-8') as f:
        f.write(content)
    
    print("SUCCESS: AI analysis section has been updated with improved parsing logic!")
    print("The fix will now handle incomplete AI responses better.")
else:
    print("ERROR: Could not find the exact section to replace.")
    print("The file structure might have changed.")
