import pandas as pd
from openpyxl.reader.excel import load_workbook
import pickle
from datetime import datetime
from operator import itemgetter
from PyPDF2 import PdfFileWriter, PdfFileReader
import os
import glob
from reportlab.pdfgen import canvas
import time
import concurrent.futures


lista = []
# lista = [['DV-004-2021', '29/09/2021'], ['IN-006-2021', '30/09/2021']]
pasta1 = os.listdir('G:\GECOT\Análise Contábil_Tributária_Licitações\\2021\\1Pendentes\\')
pasta = []
for item in pasta1:
    if item.endswith('.pdf') is True and item != 'watermark.pdf':
        pasta.append(item)

caminho = 'G:\GECOT\Análise Contábil_Tributária_Licitações\\2021\\1Pendentes\\'
salvos = []
for n, arquivo in enumerate(pasta):

    dir_atual = 'G:\GECOT\Análise Contábil_Tributária_Licitações\\2021\\1Pendentes\\'
    os.chdir(dir_atual)

    start_time = time.time()
    if n == 0:
        # Create the watermark from an image
        c = canvas.Canvas('watermark.pdf')
        # Draw the image at x, y. I positioned the x,y to be where i like here
        c.drawImage('Paulo.png', 440, 30, 100, 60, mask='auto')
        c.save()
        # Get the watermark file you just created
        watermark = PdfFileReader(open("watermark.pdf", "rb"))
        # Get our files ready
        output_file = PdfFileWriter()
    with open(caminho + arquivo, "rb") as f:
        input_file = PdfFileReader(f, "rb")
        # Number of pages in input document
        page_count = input_file.getNumPages()

        # Go through all the input file pages to add a watermark to them
        for page_number in range(page_count):
            input_page = input_file.getPage(page_number)
            if page_number == page_count - 1:
                input_page.mergePage(watermark.getPage(0))
            output_file.addPage(input_page)
        print("--- %s seconds ---" % (time.time() - start_time))
        # dir = os.getcwd()
        path = 'G:\GECOT\Análise Contábil_Tributária_Licitações\\2021'
        os.chdir(path)
        file = glob.glob(str(arquivo[21:32]) + '*')
        file = ''.join(file)
        try:
            os.chdir(file)
        except:
            os.chdir(path)

        # finally, write "output" to document-output.pdf
        with open('Análise Tributária - ' + str(arquivo[21:]) + '.pdf', "wb") as outputStream:
            output_file.write(outputStream)

    os.chdir(caminho)
    # os.remove(arquivo)
    # salvos.append(n)



# troca = 0
# for i in salvos:
#     pasta.pop(i-troca)
#
#     troca += 1


# for n, arquivo in enumerate(pasta):
#     start_time = time.time()
#     print(arquivo)
#     print("--- %s seconds ---" % (time.time() - start_time))