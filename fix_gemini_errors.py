"""
Fix for Gemini API 503 UNAVAILABLE Error
This script improves error handling for high demand situations.
"""

import re

# Read the original file
with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Update the retry logic section
old_retry_section = '''    # Tenta até o limite configurado com backoff exponencial para rate limits
    for tentativa in range(MAX_RETRIES):
        try:
            # Delay entre requisições para estabilidade
            if tentativa > 0:
                delay = RETRY_DELAY * (2 ** (tentativa - 1))
                logger.warning(f"Aguardando {delay}s antes da tentativa {tentativa + 1}...")
                time.sleep(delay)

            response = gemini_client.models.generate_content(
                model=GEMINI_MODEL,
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
            
            logger.info(f"Orientação gerada com sucesso para {dados_negocio['negocio']}")
            return resultado
            
        except Exception as e:
            error_msg = str(e).lower()
            logger.error(f"Erro na análise do negócio {dados_negocio['negocio']}: {str(e)}")
            if tentativa == MAX_RETRIES - 1:
                return f"Erro na análise (IA indisponível): {str(e)}"
            continue

    return "Não foi possível gerar a análise (limite de tentativas excedido)."'''

new_retry_section = '''    # Tenta até o limite configurado com backoff exponencial melhorado
    for tentativa in range(MAX_RETRIES):
        try:
            # Delay progressivo entre requisições para estabilidade
            if tentativa > 0:
                # Backoff exponencial com jitter para evitar picos simultâneos
                base_delay = RETRY_DELAY * (2 ** (tentativa - 1))
                jitter = base_delay * 0.1 * (hash(hash_cache) % 10) / 10  # Jitter de 0-10%
                delay = base_delay + jitter
                logger.warning(f"Aguardando {delay:.1f}s antes da tentativa {tentativa + 1}/{MAX_RETRIES}...")
                time.sleep(delay)

            response = gemini_client.models.generate_content(
                model=GEMINI_MODEL,
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
            
            logger.info(f"Orientação gerada com sucesso para {dados_negocio['negocio']} (tentativa {tentativa + 1})")
            return resultado
            
        except Exception as e:
            error_msg = str(e).lower()
            logger.error(f"Erro na análise do negócio {dados_negocio['negocio']} (tentativa {tentativa + 1}): {str(e)}")
            
            # Tratamento específico para erros de alta demanda
            if '503' in error_msg or 'unavailable' in error_msg or 'high demand' in error_msg:
                if tentativa < MAX_RETRIES - 1:
                    # Para erros 503, usa delay maior e mais progressivo
                    delay_503 = 5 + (tentativa * 3)  # 5s, 8s, 11s...
                    logger.warning(f"Erro 503 detectado - alta demanda. Aguardando {delay_503}s...")
                    time.sleep(delay_503)
                    continue
                else:
                    return f"Serviço da IA está sob alta demanda. Tente novamente em alguns minutos. (Erro: {str(e)})"
            
            # Tratamento para rate limits
            elif 'rate limit' in error_msg or 'too many requests' in error_msg:
                if tentativa < MAX_RETRIES - 1:
                    delay_rate = 10 + (tentativa * 5)  # 10s, 15s, 20s...
                    logger.warning(f"Rate limit detectado. Aguardando {delay_rate}s...")
                    time.sleep(delay_rate)
                    continue
                else:
                    return f"Limite de requisições atingido. Aguarde alguns minutos antes de tentar novamente. (Erro: {str(e)})"
            
            # Para outros erros, continua com o retry normal
            elif tentativa == MAX_RETRIES - 1:
                return f"Erro na análise (IA indisponível): {str(e)}"
            
            continue

    return "Não foi possível gerar a análise (limite de tentativas excedido)."'''

# Replace the retry section
if old_retry_section in content:
    content = content.replace(old_retry_section, new_retry_section)
    
    # Also update the MAX_RETRIES to be higher for better resilience
    content = content.replace(
        "MAX_RETRIES = int(os.getenv('MAX_RETRIES', '3'))",
        "MAX_RETRIES = int(os.getenv('MAX_RETRIES', '5'))"  # Increased from 3 to 5
    )
    
    # Write back to file
    with open('app.py', 'w', encoding='utf-8') as f:
        f.write(content)
    
    print("SUCCESS: Gemini error handling has been improved!")
    print("- Increased MAX_RETRIES from 3 to 5")
    print("- Added specific handling for 503 UNAVAILABLE errors")
    print("- Added exponential backoff with jitter")
    print("- Added specific delays for rate limits")
    print("- Better error messages for users")
else:
    print("ERROR: Could not find the retry section to replace.")
