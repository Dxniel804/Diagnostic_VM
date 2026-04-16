import re

# Read the original app.py file
with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Read the corrected AI analysis code
with open('fix_ai_analysis.py', 'r', encoding='utf-8') as f:
    fix_content = f.read()

# Extract just the code part (excluding comments and print statements)
fix_code_start = fix_content.find('# Processamento da Análise da IA - VERSÃO MELHORADA')
fix_code_end = fix_content.find('\n\nprint("""')
ai_fix_code = fix_content[fix_code_start:fix_code_end].strip()

# Find and replace the old AI analysis section in app.py
old_pattern = r'# Processamento da Análise da IA.*?elements\.append\(Paragraph\(clean_line, normal_style\)\)'
new_pattern = ai_fix_code

# Use regex to replace the entire section
content = re.sub(old_pattern, new_pattern, content, flags=re.DOTALL)

# Write the updated content back to app.py
with open('app.py', 'w', encoding='utf-8') as f:
    f.write(content)

print("SUCCESS: AI analysis code has been copied and applied to app.py!")
print("The improved parsing logic is now active in your main application.")
