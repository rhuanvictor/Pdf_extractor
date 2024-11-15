from PyQt6.QtWidgets import QMainWindow, QFileDialog, QMessageBox, QPushButton, QLabel, QTableWidget, QTableWidgetItem, QDialog, QVBoxLayout, QLineEdit, QComboBox, QLineEdit
from PyQt6.QtCore import Qt
from pdf_extractor import PDFExtractor
import fitz  # PyMuPDF para manipulação de PDFs
import os

class PDFDataExtractor(QMainWindow):
    def __init__(self):
        super().__init__()

        # Configuração da janela principal
        self.setWindowTitle("Extrator de Dados de Currículo em PDF")
        self.setGeometry(200, 200, 400, 400)
        
        # Botões e Labels
        self.btn_select_folder = QPushButton("Selecionar Pasta", self)
        self.btn_select_folder.setGeometry(50, 50, 300, 40)
        self.btn_select_folder.clicked.connect(self.select_folder)
        
        self.btn_save_excel = QPushButton("Salvar como Excel", self)
        self.btn_save_excel.setGeometry(50, 100, 300, 40)
        self.btn_save_excel.clicked.connect(self.save_excel)
        self.btn_save_excel.setEnabled(False)

        # Botão para ler Excel
        self.btn_read_excel = QPushButton("Ler Excel", self)
        self.btn_read_excel.setGeometry(50, 150, 300, 40)
        self.btn_read_excel.clicked.connect(self.read_excel)
        
        # Botão para localizar dados no PDF
        self.btn_find_data = QPushButton("Localizar Dados no PDF", self)
        self.btn_find_data.setGeometry(50, 200, 300, 40)
        self.btn_find_data.clicked.connect(self.find_data_in_pdf)
        self.btn_find_data.setEnabled(False)
        
        self.label_folder = QLabel("Pasta: Não selecionada", self)
        self.label_folder.setGeometry(50, 20, 300, 20)
        self.label_folder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.folder_path = ""
        self.extracted_data = []
        self.dataframe = None

    def select_folder(self):
        # Abre o diálogo para selecionar a pasta
        self.folder_path = QFileDialog.getExistingDirectory(self, "Selecionar pasta com PDFs", "")
        
        if self.folder_path:
            self.label_folder.setText(f"Pasta: {self.folder_path}")
            self.btn_save_excel.setEnabled(True)
            self.btn_find_data.setEnabled(True)
            self.extract_data_from_pdfs()

    def extract_data_from_pdfs(self):
        # Extrai dados de todos os PDFs na pasta selecionada
        self.extracted_data = []  # Limpa dados extraídos antes de começar

        for filename in os.listdir(self.folder_path):
            if filename.endswith(".pdf"):
                pdf_path = os.path.join(self.folder_path, filename)
                extractor = PDFExtractor(pdf_path)
                self.extracted_data.extend(extractor.extract_data())  # Adiciona os dados extraídos

    def find_data_in_pdf(self):
        # Abre uma nova janela para inserir a palavra-chave de busca
        dialog = QDialog(self)
        dialog.setWindowTitle("Buscar Dados em PDFs")
        layout = QVBoxLayout(dialog)

        # Campo para digitar a palavra-chave
        search_label = QLabel("Digite a palavra-chave para buscar:", dialog)
        layout.addWidget(search_label)
        
        search_input = QLineEdit(dialog)
        layout.addWidget(search_input)

        # Botão de busca
        btn_search = QPushButton("Buscar", dialog)
        layout.addWidget(btn_search)
        btn_search.clicked.connect(lambda: self.perform_search(search_input.text(), dialog))

        dialog.setLayout(layout)
        dialog.resize(300, 150)
        dialog.exec()

    def perform_search(self, keyword, dialog):
        # Função que realiza a busca nos PDFs
        if not keyword:
            QMessageBox.warning(self, "Atenção", "Por favor, insira uma palavra-chave.")
            return

        results = []  # Armazena resultados da busca

        for filename in os.listdir(self.folder_path):
            if filename.endswith(".pdf"):
                pdf_path = os.path.join(self.folder_path, filename)
                with fitz.open(pdf_path) as pdf:
                    for page_num in range(len(pdf)):
                        page = pdf[page_num]
                        text = page.get_text("text")
                        if keyword.lower() in text.lower():
                            results.append((filename, page_num + 1))  # Armazena nome do arquivo e página encontrada

        # Exibe resultados em uma caixa de diálogo
        if results:
            result_text = "\n".join([f"Arquivo: {res[0]}, Página: {res[1]}" for res in results])
            QMessageBox.information(self, "Resultados da Busca", f"Encontrado:\n\n{result_text}")
        else:
            QMessageBox.information(self, "Resultados da Busca", "Nenhum resultado encontrado para a palavra-chave.")
        
        dialog.close()

    def save_excel(self):
        # Salva os dados extraídos em um arquivo Excel
        file_path, _ = QFileDialog.getSaveFileName(self, "Salvar Arquivo", "", "Excel Files (*.xlsx);;All Files (*)")
        if file_path:
            extractor = PDFExtractor(self.folder_path)
            extractor.save_to_excel(file_path, self.extracted_data)
            QMessageBox.information(self, "Sucesso", "Dados extraídos e salvos com sucesso no Excel.")

    def read_excel(self):
        # Abre o diálogo para selecionar o arquivo Excel
        file_path, _ = QFileDialog.getOpenFileName(self, "Selecionar Arquivo Excel", "", "Excel Files (*.xlsx);;All Files (*)")
        if file_path:
            self.dataframe = read_excel_to_dataframe(file_path)
            if self.dataframe is not None:
                self.show_data_in_table(self.dataframe)
            else:
                QMessageBox.warning(self, "Erro", "Falha ao ler o arquivo Excel.")
