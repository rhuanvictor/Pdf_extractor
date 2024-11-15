import re
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import pandas as pd

def read_excel_to_dataframe(file_path):
    try:
        # Lê o arquivo Excel em um DataFrame
        dataframe = pd.read_excel(file_path)
        return dataframe
    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {e}")
        return None

def save_dataframe_to_excel(file_path, dataframe):
    try:
        # Salva o DataFrame em um arquivo Excel
        dataframe.to_excel(file_path, index=False)
    except Exception as e:
        print(f"Erro ao salvar o arquivo Excel: {e}")


def read_excel_to_dataframe(file_path):
    try:
        # Lê o arquivo Excel em um DataFrame
        dataframe = pd.read_excel(file_path)
        return dataframe
    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {e}")
        return None


def extract_field_especializacao(text):
    # Lista de palavras a serem ignoradas
    palavras_ignoradas = [
        "Ensino Fundamental - ", "Ensino Médio - ", "Ensino Médio Completo ", "Segundo Grau Completo - ", 
        "Técnico em ", "Graduação em ", "Pós-Graduação em ", "MBA em ", 
        "Doutorado em ", "Mestrado em ", "Pós-Doutorado  em "
    ]
    
    # Criar uma expressão regular para remover as palavras ignoradas
    for palavra in palavras_ignoradas:
        text = re.sub(rf"\b{re.escape(palavra)}\b", "", text, flags=re.IGNORECASE)
    
    # Buscar o texto após "Formação" e retornar o resultado
    match = re.search(r"Formação\s*[:\-]?\s*(.*?)(?:\n|$)", text)
    if match:
        return match.group(1).strip()
    return ""


def extract_email(text):
    email = re.findall(r"[\w\.-]+@[\w\.-]+\.[a-zA-Z]{2,}", text)
    if email:
        return email[0]
    return ""


def extract_phones(text):
    
    phones = re.findall(r"\(?\d{2}\)?\s?\d{4,5}-\d{4}", text)
    
    if phones:
        # Remover os caracteres indesejados (parênteses, espaço e hífen)
        phone = re.sub(r"[^\d]", "", phones[0])
        return phone
    return ""


def extract_nome(text):
    # Remove o "soft hyphen" e espaços ao redor, se houver
    text = re.sub(r'\s*\xad\s*', '', text)
    text = re.sub(r'\s+', ' ', text).strip()
    
    # Expressão regular ajustada para permitir iniciais com ponto no nome (ex: "W.")
    match = re.search(r"Nome\s*[:\-]?\s*([A-Za-zÀ-ÿ'\s\-\.]+?)(?=\s*(Sexo|$|[\d]+))", text)
    if match:
        nome = match.group(1).strip()
        # Formata a primeira letra de cada palavra, incluindo iniciais
        nome_formatado = ' '.join([palavra.capitalize() for palavra in nome.split()])
        return nome_formatado
    return ""



def extract_gender(text):
    match = re.search(r"Sexo\s*[:\-]?\s*(Masculino|Feminino|Outros)", text, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return ""



def extract_formacao(text):
    # Ordem dos níveis de formação do mais baixo para o mais alto
    formacao_ordem = [
        "Ensino Fundamental", "Ensino Fundamental Completo", "Segundo Grau Completo", 
        "Ensino Médio", "Ensino Médio Completo", "Técnico", "Graduação", 
        "Pós-Graduação", "MBA", "Mestrado", "Doutorado", "Pós-Doutorado"
    ]
    
    # Níveis de formação considerados "superiores ao ensino médio"
    formacao_superior = formacao_ordem[5:]  # Técnico, Graduação, Pós-Graduação, MBA, Mestrado, Doutorado, Pós-Doutorado
    
    # Regex para identificar os níveis de formação
    pattern = r"(Ensino Fundamental|Ensino Fundamental Completo|Segundo Grau Completo|Ensino Médio Completo|Ensino Médio|Técnico|Graduação|Pós-Graduação|MBA|Mestrado|Doutorado|Pós-Doutorado)"
    
    # Posições onde a busca deve ser interrompida
    stop_sections = ["Cursos e especializações", "Idiomas", "Dados Pessoais"]

    # Encontra a posição da palavra "Formação" para iniciar a busca a partir desse ponto
    start_pos = text.find("Formação")
    if start_pos == -1:
        return ""  # Retorna vazio se "Formação" não for encontrada

    # Limita o texto a partir da palavra "Formação"
    text = text[start_pos:]
    
    # Limita a busca até a primeira seção de parada encontrada
    for stop_section in stop_sections:
        if stop_section in text:
            text = text.split(stop_section)[0]
            break

    # Busca todas as formações encontradas no texto
    matches = re.findall(pattern, text)
    
    # Filtra as formações para separar níveis superiores e inferiores ao ensino médio
    formacoes_superiores = [match for match in matches if match in formacao_superior]
    formacoes_inferiores = [match for match in matches if match not in formacao_superior]
    
    # Se houver formações superiores ao ensino médio, retornamos todas em ordem do mais baixo para o mais alto
    if formacoes_superiores:
        # Ordena as formações encontradas de acordo com a hierarquia em `formacao_ordem`
        formacoes_superiores = sorted(formacoes_superiores, key=formacao_ordem.index)
        return ", ".join(formacoes_superiores)
    
    # Se houver apenas níveis abaixo do ensino médio, retorna o mais alto desses
    elif formacoes_inferiores:
        return max(formacoes_inferiores, key=formacao_ordem.index)
    
    # Retorna vazio se nenhuma formação for encontrada
    return ""



def extract_city_uf(text):
    city_uf_match = re.search(r"Cidade\s*[:\-]?\s*([\w\s]+)\s*-\s*([A-Za-z]{2})", text)
    if city_uf_match:
        return city_uf_match.group(1).strip(), city_uf_match.group(2).strip()
    return "", ""


def extract_idiomas_niveis(text):
    idiomas = []
    niveis = []
    
    # Encontra a posição de "Dados Pessoais" no texto
    dados_pessoais_pos = text.lower().find("dados pessoais")
    
    if dados_pessoais_pos != -1:
        # Pega o texto até "Dados Pessoais"
        relevant_text = text[:dados_pessoais_pos]
        
        # Encontra a última ocorrência de "Idiomas" antes de "Dados Pessoais"
        idiomas_pos = relevant_text.lower().rfind("idiomas")
        
        if idiomas_pos != -1:
            # Pega o texto entre "Idiomas" e "Dados Pessoais"
            idiomas_section = relevant_text[idiomas_pos + len("idiomas"):]  # Pula a palavra "Idiomas"
            
            # Limpa quebras de linha e espaços extras
            idiomas_section = idiomas_section.replace("\n", " ").strip()

            # Remove URLs, datas, frações (ex: 9/48, 19/42) e palavras irrelevantes (ex: Catho)
            idiomas_section = re.sub(r"https?://[^\s]+", "", idiomas_section)  # Remove URLs
            idiomas_section = re.sub(r"\d{2}/\d{2}/\d{4}(?:, \d{2}:\d{2})?", "", idiomas_section)  # Remove datas
            idiomas_section = re.sub(r"\d+/\d+", "", idiomas_section)  # Remove frações como 9/48, 19/42
            idiomas_section = re.sub(r"\bCatho\b", "", idiomas_section, flags=re.IGNORECASE)  # Remove a palavra 'Catho'

            # Divide o texto em partes, alternando entre idiomas e níveis
            parts = re.split(r'(?<=\w)\s+', idiomas_section)
            for i in range(0, len(parts), 2):  # Pega duas palavras por vez
                idioma = parts[i].strip()
                nivel = parts[i + 1].strip() if i + 1 < len(parts) else ''
                
                if idioma and nivel:
                    idiomas.append(idioma.capitalize())
                    niveis.append(nivel.capitalize())
    
    return idiomas, niveis
    

def extract_pretensao_salarial(text):
    match = re.search(r"Pretensão\s*Salarial\s*[:\-]?\s*(.*?)(?=Atualizado)", text, re.IGNORECASE | re.DOTALL)
    if match:
        return match.group(1).strip()
    return ""

def extract_cnh(text):
    match = re.search(r"CNH\s*(?:-?\s*Categoria\s*)?([A-Za-z])", text, re.IGNORECASE)
    if match:
        return match.group(0).strip()
    return ""

def extract_cargo_interesse(text):
    match = re.search(r"Cargo\s*de\s*interesse\s*[:\-]?\s*(.*)", text, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return ""

def clean_text(text):
    # Limpa o texto: remove quebras de linha e múltiplos espaços
    text = text.replace('\n', ' ')
    text = re.sub(r'\s+', ' ', text)
    text = text.strip()
    return text

def extract_experiencia_profissional(text):
    # Limpa o texto de entrada
    text = clean_text(text)
    
    # Listas para armazenar as experiências e os tempos
    experiencias = []
    tempos = []
    
    # Encontra o bloco de "Experiência Profissional"
    experiencia_block = re.search(r"Experiência Profissional(.*?)(Formação|Cursos e especializações|Informações Adicionais|Dados Pessoais|$)", text, re.DOTALL)
    if experiencia_block:
        experiencia_text = experiencia_block.group(1)
        
        # Lista temporária para empilhar os cargos até encontrar "Último salário"
        cargos_encontrados = []
        
        # Divide o bloco de experiência em seções entre "Cargo:" e "Último salário"
        tokens = re.split(r"(Cargo:|Último salário)", experiencia_text)
        
        i = 0
        while i < len(tokens):
            token = tokens[i].strip()
            
            # Se o token é "Cargo:", o próximo item é o cargo que deve ser empilhado
            if token == "Cargo:" and i + 1 < len(tokens):
                cargo_text = tokens[i + 1].strip()
                cargos_encontrados.append(cargo_text)
                i += 2  # Pula para o próximo token após o cargo
            
            # Se o token é "Último salário", pega o último cargo empilhado e limpa a lista
            elif token == "Último salário" and cargos_encontrados:
                # Pega apenas o último cargo armazenado, descartando os anteriores
                ultimo_cargo = cargos_encontrados[-1]
                
                # Procura o tempo de experiência (ex: "1 ano e 4 meses")
                tempo_experiencia = re.search(r"(\d{1,2} anos? e \d{1,2} meses?|1 mês|1 ano|\d{1,2} meses?)", ultimo_cargo)
                tempo = tempo_experiencia.group(0) if tempo_experiencia else ""
                
                # Adiciona à lista de experiências e tempos
                experiencias.append(ultimo_cargo)
                tempos.append(tempo)
                
                # Limpa a lista de cargos encontrados após usar o último
                cargos_encontrados.clear()
                i += 1  # Continua para o próximo token
            
            else:
                i += 1
    
    # Garantir que as listas de experiências e tempos tenham 3 itens
    while len(experiencias) < 3:
        experiencias.append("")
    while len(tempos) < 3:
        tempos.append("")
    
    return experiencias[:3], tempos[:3]  # Retornar apenas os 3 primeiros


def limpar_vagas(texto):
    # Expressão regular para identificar e remover o texto a partir de uma data ou um link
    padrao = r" - Último cargo.*| - \d{2}/\d{4}.*| \d{2}/\d{2}/\d{4}.*|https?://.*"
    
    # Substitui a parte identificada pelo padrão por uma string vazia
    texto_limpo = re.sub(padrao, "", texto)
    
    # Remove espaços extras antes e depois do texto
    return texto_limpo.strip()