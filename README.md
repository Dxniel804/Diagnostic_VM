# 🤖 Sistema de Análise Estratégica de Follow-ups de Vendas

Sistema inteligente que lê planilhas Excel ou CSV de CRM e gera estratégias personalizadas para o próximo follow-up usando Inteligência Artificial.

## 🎯 Objetivo

Resolver o problema do vendedor que tem centenas de leads e se perde no meio de tantos follow-ups. A ferramenta:
- Lê o histórico de conversas até agora
- Identifica automaticamente onde a conversa parou
- Usa IA para gerar orientações estratégicas precisas para o próximo passo
- Gera relatórios em PDF profissionais

## 🏗️ Arquitetura

- **Backend**: Python com Flask
- **Análise de Dados**: Pandas (leitura e processamento de planilhas Excel/CSV)
- **IA**: Groq API (gratuita e muito rápida)
- **Frontend**: HTML5 + CSS3 (interface moderna e responsiva)
- **Exportação**: PDF profissional com ReportLab

## 📋 Estrutura da Planilha

A planilha (Excel ou CSV) deve conter as seguintes colunas:

### Colunas Obrigatórias:
- `Nome do negócio` - Nome do negócio/proposta
- `Empresa` - Nome da empresa cliente
- `Fase` - Fase atual do negócio (ex: Proposta, Negociação, etc)
- `Responsavel` - Vendedor responsável
- `Temperatura da Proposta Follow 1` - Temperatura do 1º follow-up
- `Descrição Follow up 1` - Descrição do 1º follow-up

### Colunas Opcionais (até 5 follow-ups):
- `Temperatura da Proposta Follow 2-5`
- `Descrição Follow up 2-5`

## 🚀 Instalação

1. **Clone o repositório:**
```bash
git clone <repositorio>
cd Analise_propostas
```

2. **Crie ambiente virtual:**
```bash
python -m venv venv
source venv/bin/activate  # Linux/Mac
# ou
venv\Scripts\activate  # Windows
```

3. **Instale dependências:**
```bash
pip install -r requirements.txt
```

4. **Configure variáveis de ambiente:**
```bash
cp .env.example .env
# Edite .env com sua GROQ_API_KEY
```

## ⚙️ Configuração da API de IA

### 🎯 Groq API (Recomendada - GRATUITA)

**Por que Groq?**
- ✅ **Gratuita** - Tier gratuito generoso
- ✅ **Muito Rápida** - Respostas em milissegundos
- ✅ **Sem cartão de crédito** - Para começar
- ✅ **Modelos poderosos** - Llama 3.3 70B, Mixtral, etc

**Como obter a chave:**
1. Acesse: https://console.groq.com/
2. Crie uma conta (gratuita)
3. Vá em "API Keys" e gere uma nova chave
4. Cole no arquivo `.env`:

```env
GROQ_API_KEY=gsk_sua_chave_aqui
GROQ_MODEL=llama-3.3-70b-versatile
```

**Limites do Tier Gratuito:**
- Rate limit: ~30 requisições por minuto (varia por modelo)
- Tokens por minuto: Generoso para uso pessoal
- Sem limite de requisições totais (dentro do rate limit)

**Configuração no .env:**
```env
GROQ_API_KEY=sua_chave_aqui
GROQ_MODEL=llama-3.3-70b-versatile
MAX_RETRIES=3
RETRY_DELAY=5
REQUEST_DELAY=2
```

### 🔄 Outras APIs (Alternativas)

Se preferir usar outras APIs, você pode modificar o código em `app.py`:

- **Hugging Face Inference API**: Gratuita com limite razoável
- **OpenRouter**: Agrega vários modelos gratuitos
- **Ollama**: Local (não precisa de API key, mas precisa instalar)

## 🏃‍♂️ Execução

Inicie a aplicação:
```bash
python app.py
```

Acesse `http://localhost:5000` no navegador.

## 📊 Funcionalidades

### ✅ Implementado:

1. **Upload de Planilhas**
   - Suporta `.xlsx`, `.xls` e `.csv` (CSV é recomendado - mais confiável!)
   - Detecta e converte arquivos HTML disfarçados de Excel
   - Validação automática de estrutura
   - **Dica:** Se tiver problemas com arquivos Excel, converta para CSV - é mais simples e confiável!

2. **Processamento Inteligente**
   - Identifica automaticamente onde a conversa parou
   - Foca no próximo follow-up a ser realizado
   - Processa linha por linha com análise personalizada

3. **Análise Estratégica com IA**
   - A IA age como um Diretor Comercial experiente
   - Gera 3 seções: Diagnóstico, Estratégia e Ação Recomendada
   - Foca em fechamento de vendas

4. **Dashboard Interativo**
   - Visualização clara de cada negócio
   - Histórico de follow-ups
   - Análise estratégica destacada

5. **Relatório PDF Profissional**
   - Cabeçalho do Cliente (Nome da Empresa e Responsável)
   - Status Atual (Fase e Temperatura)
   - Follow-up atual e próximo
   - Plano de Ação completo da IA
   - Pronto para impressão ou reunião

6. **Sistema de Cache**
   - Evita requisições duplicadas à API
   - Acelera processamento de planilhas grandes

7. **Tratamento de Erros**
   - Retry automático em caso de rate limit
   - Logs detalhados para debugging
   - Mensagens de erro amigáveis

## 🔍 Como Funciona

### Fluxo do Sistema:

1. **Upload** → Usuário faz upload da planilha Excel
2. **Processamento** → Sistema lê e valida a planilha
3. **Identificação** → Para cada linha, identifica onde a conversa parou
4. **Análise IA** → Gera estratégia para o próximo follow-up
5. **Visualização** → Exibe resultados no dashboard
6. **Exportação** → Gera PDF profissional

### Regra de Ouro:

O sistema identifica automaticamente onde a conversa parou:
- Se o vendedor preencheu até o "Follow-up 2", a IA foca no "Follow-up 3"
- Se não escreveu nada, foca no "Follow-up 1"
- A IA lê apenas os follow-ups preenchidos para contexto

## 📝 Logs

A aplicação gera logs em:
- **Console** (tempo real)
- **Arquivo `app.log`** (persistente)

Níveis de log:
- `INFO`: Operações normais e progresso
- `WARNING`: Limites de API atingidos
- `ERROR`: Erros de processamento

## 🛡️ Segurança

- ✅ Variáveis de ambiente para dados sensíveis
- ✅ Validação de formato de arquivo
- ✅ Tratamento de erros robusto
- ✅ Sanitização de nomes de colunas
- ✅ Arquivos temporários são removidos automaticamente

## 📈 Performance

- **Tempo por linha**: ~2-3 segundos (incluindo delay da API)
- **Cache**: Análises repetidas são instantâneas
- **Rate Limiting**: Configurado para respeitar limites da API

## 🆘 Solução de Problemas

### Erro: "GROQ_API_KEY não encontrada"
- Verifique se o arquivo `.env` existe
- Confirme que a chave está correta no `.env`
- Reinicie o servidor após alterar o `.env`

### Erro: "Colunas obrigatórias faltando"
- Verifique se a planilha tem todas as colunas necessárias
- Confirme os nomes exatos das colunas (case-sensitive)

### Erro: "Não foi possível ler o arquivo Excel"
- Tente converter o arquivo para `.xlsx` primeiro
- Verifique se o arquivo não está corrompido
- Arquivos HTML disfarçados de Excel são convertidos automaticamente

### Rate Limit da API
- O sistema faz retry automático
- Aumente `REQUEST_DELAY` no `.env` se necessário
- Verifique seus limites em: https://console.groq.com/settings/limits

## 🤝 Contribuição

1. Fork o projeto
2. Crie branch para feature (`git checkout -b feature/nova-funcionalidade`)
3. Commit mudanças (`git commit -am 'Adiciona nova funcionalidade'`)
4. Push para branch (`git push origin feature/nova-funcionalidade`)
5. Abra Pull Request

## 📄 Licença

MIT License - ver arquivo LICENSE para detalhes.

## 🆘 Suporte

Para problemas:
1. Verifique os logs em `app.log`
2. Confirme estrutura da planilha
3. Valide configuração da API Key
4. Verifique limites de cota da Groq API em: https://console.groq.com/settings/limits

---

**Desenvolvido com ❤️ para vendedores que querem fechar mais vendas!**
