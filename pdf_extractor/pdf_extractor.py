import fitz  # PyMuPDF
import re
import pandas as pd
from utils import limpar_vagas, extract_experiencia_profissional, extract_field_especializacao, extract_email, extract_phones, extract_nome, extract_gender, extract_formacao, extract_city_uf, extract_idiomas_niveis, extract_pretensao_salarial, extract_cnh, extract_cargo_interesse
import os  # Para manipular o nome do arquivo
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
import xlwings as xw

import time
import win32com.client as win32

class PDFExtractor:
    # Variável de classe para manter o último ID atribuído
    last_id = 0

    def __init__(self, pdf_path):
        self.pdf_path = pdf_path
        self.extracted_data = []
        
        # Verifica o nome do arquivo para determinar a fonte
        self.fonte = self.identify_source()

    def identify_source(self):
        # Extrai o nome do arquivo sem a extensão
        return os.path.splitext(os.path.basename(self.pdf_path))[0]

    def extract_data(self):
        # Extrai dados do PDF
        full_text = ""
        with fitz.open(self.pdf_path) as pdf:
            for page_num in range(len(pdf)):
                full_text += pdf[page_num].get_text("text")

        profiles = full_text.split("País")
        for profile in profiles[:-1]:
            # Extrai os dados do perfil
            experiencia_profissional, tempo_experiencia = extract_experiencia_profissional(profile)
            
            # Incrementa o ID globalmente
            PDFExtractor.last_id += 1
            
            # Construindo o dicionário de dados e aplicando limpar_vagas apenas nas chaves de Experiência Profissional
            data = {
                "ID": PDFExtractor.last_id,  # Usa o ID global incrementado
                "Fonte": self.fonte,  # Fonte identificada pelo nome do arquivo PDF
                "Nome Completo": extract_nome(profile),
                "Telefone": extract_phones(profile),
                "E-mail": extract_email(profile),
                "Gênero": extract_gender(profile),
                "Cidade": extract_city_uf(profile)[0],
                "UF": extract_city_uf(profile)[1],
                "Formação": extract_formacao(profile),
                "Especialização": extract_field_especializacao(profile),
                "Pretensão Salarial (R$)": extract_pretensao_salarial(profile),
                "Cargo de Interesse": extract_cargo_interesse(profile),
                "Idiomas": ", ".join(extract_idiomas_niveis(profile)[0]),
                "Nível do Idioma": ", ".join(extract_idiomas_niveis(profile)[1]),
                "CNH": extract_cnh(profile),
                # Adicionando as experiências e tempos extraídos e limpos
                "Experiência Profissional 1": limpar_vagas(experiencia_profissional[0]),
                "Tempo de Experiência 1": tempo_experiencia[0],
                "Experiência Profissional 2": limpar_vagas(experiencia_profissional[1]),
                "Tempo de Experiência 2": tempo_experiencia[1],
                "Experiência Profissional 3": limpar_vagas(experiencia_profissional[2]),
                "Tempo de Experiência 3": tempo_experiencia[2],
            }
            self.extracted_data.append(data)
        
        return self.extracted_data



    def save_to_excel(self, file_path, data):
        # Salva os dados extraídos em um arquivo Excel
        df = pd.DataFrame(data)
        
        # Salva a planilha no Excel com pandas e openpyxl
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Currículos')
            worksheet = writer.sheets['Currículos']

            # Ajusta a largura das colunas
            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                worksheet.column_dimensions[column].width = adjusted_width

        # Inicializa o Excel com xlwings e executa o código VBA
        app = xw.App(visible=False)  # Abre o Excel em modo invisível para processamento em segundo plano
        try:
            wb = app.books.open(file_path)  # Abre o arquivo Excel criado

            # Código VBA a ser inserido
            vba_code = '''
            Sub CongelarPainel()
                Dim ws As Worksheet
                ' Definir a planilha chamada "Currículos"
                Set ws = ThisWorkbook.Sheets("Currículos")
                
                ' Congelar a primeira linha e as primeiras três colunas
                ws.Activate
                ws.Range("D2").Select ' Definir a célula que ficará no ponto de congelamento
                
                ' Congelar a linha 1 (quando rolar verticalmente) e as colunas 1, 2, 3 (quando rolar horizontalmente)
                ActiveWindow.FreezePanes = True
            End Sub
            '''

            # Acessa o módulo VBA e insere o código
            vb_module = wb.api.VBProject.VBComponents.Add(1)  # Adiciona um módulo
            vb_module.CodeModule.AddFromString(vba_code)  # Adiciona o código VBA ao módulo

            # Executa o código VBA inserido
            wb.macro('CongelarPainel')()

            # Espera um tempo para que o código seja executado
            time.sleep(2)  # Tempo de espera para garantir que o código VBA finalize

            # Deleta o código VBA
            wb.api.VBProject.VBComponents.Remove(vb_module)  # Remove o módulo VBA

            # Salva o arquivo
            wb.save()
        finally:
            # Fecha o workbook e o aplicativo
            wb.close()
            app.quit()
