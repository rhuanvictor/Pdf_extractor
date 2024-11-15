# Extrator de Dados de Currículos em PDF

Este projeto tem como objetivo extrair informações importantes de currículos em formato PDF, como nome, telefone, e-mail, gênero, cidade, estado (UF), e outros campos relacionados à formação, experiência e pretensão salarial. Ele utiliza o PyQt6 para criar a interface gráfica, o PyMuPDF (fitz) para ler os arquivos PDF, e o pandas para gerar uma planilha Excel com os dados extraídos.

## Funcionalidades

- **Seleção de PDF**: O usuário pode selecionar varios arquivos em PDF contendo varios currículos.
- **Extração de Dados**: O sistema extrai informações como:
  - Nome completo
  - Telefone
  - E-mail
  - Gênero
  - Cidade
  - Estado (UF)
  - Formação
  - Especialização
  - Pretensão salarial
  - Idiomas e nível de idioma
- **Geração de Planilha Excel**: Os dados extraídos podem ser salvos em um arquivo Excel.

## Tecnologias Utilizadas

- **PyQt6**: Para a interface gráfica do usuário (GUI).
- **PyMuPDF (fitz)**: Para a leitura e extração de texto de arquivos PDF.
- **pandas**: Para manipulação de dados e criação da planilha Excel.
- **re (expressões regulares)**: Para a extração de informações específicas do currículo.

## Como Usar

1. **Instalar Dependências**:

   Primeiro, instale as dependências do projeto. Você pode usar o `pip` para isso:

   ```bash
   pip install PyQt6 PyMuPDF pandas
