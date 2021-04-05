import openpyxl
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.utils.cell import coordinate_from_string
from openpyxl.workbook.protection import WorkbookProtection
import sys

arquivo = sys.argv[1]

# Here you need to set the file and images that you want to add to your workbook.
# Aqui voce seta o arquivo que as imagens que quer adicionar na sua planilha.
filename = arquivo # File
logo_header =  # First Image - File 
logo_footer =  # Second Image - File

# Here you turn the images on openpyxl object
# Aqui voce torna suas imagens em um objeto do openpyxl
img_header = openpyxl.drawing.image.Image(logo_header)
img_footer = openpyxl.drawing.image.Image(logo_footer)

# Here you load your existent workbook
# Aqui voce carrega a planilha existente
wb = load_workbook(filename)

# Here you can set protection to your workbook (optional)
# Aqui voce pode setar uma protecao em sua planilha (opcional)
# wb.security = WorkbookProtection(workbookPassword = 'eeccb553354788931b24a0f081cd8fcd', lockStructure = True)

# Here you set the worksheet that you want to del/add images
# Aqui voce seta a folha da planilha onde quer remover/adicionar imagens
ws = wb.active

# Deleting images 
# Deletando imagens existentes
for image in range(0, len(ws._images)):
   del ws._images[0]

# Setting protections at Worksheet
# Setando protecoes na folha
# ws.protection.set_password('eeccb553354788931b24a0f081cd8fcd')
# ws.protection.sheet = True
# ws.protection.enable()

# Adding Images
# Adicionando as imagens
# ws.add_image(img_header, 'A1')
# ws.add_image(img_footer, 'C46')

# Saving your file.
# Salvando seu arquivo.
wb.save(filename)