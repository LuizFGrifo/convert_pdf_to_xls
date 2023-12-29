import aspose.pdf as ap
import os

# Diretório onde os arquivos PDF estão localizados
diretorio_pdf = 'documentation'

# Lista todos os arquivos no diretório
arquivo_pdf = [os.path.join(diretorio_pdf, arquivo) for arquivo in os.listdir(diretorio_pdf) if arquivo.lower().endswith('.pdf')]

# Se houver pelo menos um arquivo PDF no diretório
if arquivo_pdf:
    # Encontra o arquivo mais recente
    arquivo_mais_recente = max(arquivo_pdf, key=os.path.getmtime)

    nome_sem_extensao = os.path.splitext(os.path.basename(arquivo_mais_recente))[0]

    input_pdf = arquivo_mais_recente
    output_pdf = os.path.join(diretorio_pdf, f'{nome_sem_extensao}.xls') 

    # Abrir o arquivo PDF
    document = ap.Document(input_pdf)

    save_option = ap.ExcelSaveOptions()
    save_option.format = ap.ExcelSaveOptions.ExcelFormat.XML_SPREAD_SHEET2003

    # Salva o arquivo
    document.save(output_pdf, save_option)