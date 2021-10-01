from fpdf import FPDF
from tkinter import *
from tkinter import ttk
from tkinter.tix import *
from csv import reader
from tkscrolledframe import ScrolledFrame
import pandas as pd
from datetime import date
from tkinter import messagebox
import os
import datetime
import glob
from reportlab.pdfgen import canvas
from PyPDF2 import PdfFileWriter, PdfFileReader
from time import sleep



class NotasServicos:

    def __init__(self, janela):
        self.janela = janela
        titulo = ' '
        self.janela.title(100 * titulo + 'Consultas')
        self.janela.geometry('800x500+100+20')
        self.janela.resizable(width=False, height=False)

        estilo = ttk.Style()
        estilo.theme_use('default')
        estilo.configure('Treeview', background='#D3D3D3', foreground='black', rowheight=25,
                         fieldbackground='#D3D3D3')
        estilo.map('Treeview', background=[('selected', '#347083')])

        # Treeview frame
        tree_frame = Frame(janela)
        tree_frame.pack(pady=100)
        # Barra rolagem
        tree_scroll = Scrollbar(tree_frame)
        tree_scroll.pack(side=RIGHT, fill=Y)
        # Criar Treeview
        nf_tree = ttk.Treeview(tree_frame, yscrollcommand=tree_scroll.set, selectmode='extended')
        nf_tree.pack(side=LEFT)
        # Configurar Barra Rolagem
        tree_scroll.config(command=nf_tree.yview)
        # Definir colunas
        colunas2 = ['Análise', 'Data']
        nf_tree['columns'] = colunas2
        # formatar colunas
        nf_tree.column('Análise', width=350)
        nf_tree.column('Data', width=100)

        # formatar títulos
        nf_tree.heading('Análise', text='Análise', anchor=W)
        nf_tree.heading('Data', text='Data', anchor=W)

        nf_tree['show'] = 'headings'

        self.lista = []
        # lista = [['DV-004-2021', '29/09/2021'], ['IN-006-2021', '30/09/2021']]
        self.pasta = os.listdir('C:\\Users\loliveira\PycharmProjects\Python Pdf\Relatorio-pdf\pendentes')
        print(self.pasta)
        for i, n in enumerate(self.pasta):
            mod = os.path.getctime('C:\\Users\loliveira\PycharmProjects\Python Pdf\Relatorio-pdf\pendentes')
            mod = datetime.datetime.fromtimestamp(mod)
            data = date.strftime(mod, '%d/%m/%Y')
            self.lista.append([])
            self.lista[i].append(n)
            self.lista[i].append(data)


        # print(self.lista)
        # inserir dados do banco no treeview
        def inserir_tree(lista):
            nf_tree.delete(*nf_tree.get_children())
            contagem = 0
            for row in lista:  # loop para inserir cores diferentes nas linhas
                if contagem % 2 == 0:
                    nf_tree.insert(parent='', index='end', text='', iid=contagem,
                                   values=(row[0], row[1]), tags=('evenrow',))
                else:
                    nf_tree.insert(parent='', index='end', text='', iid=contagem,
                                   values=(row[0], row[1]), tags=('oddrow',))
                contagem += 1

        def assinatura():
            caminho = 'C:\\Users\loliveira\PycharmProjects\Python Pdf\Relatorio-pdf\pendentes\\'
            for n, arquivo in enumerate(self.pasta):
                if self.list_check[n].get() == 1:
                    dir_atual = 'C:\\Users\loliveira\PycharmProjects\Python Pdf\\'
                    os.chdir(dir_atual)

                    # Create the watermark from an image
                    c = canvas.Canvas('watermark.pdf')
                    # Draw the image at x, y. I positioned the x,y to be where i like here
                    c.drawImage('Mari.png', 440, 30, 100, 60, mask='auto')
                    c.save()
                    # Get the watermark file you just created
                    watermark = PdfFileReader(open("watermark.pdf", "rb"))
                    # Get our files ready
                    output_file = PdfFileWriter()
                    input_file = PdfFileReader(open(caminho + arquivo, "rb"))
                    # Number of pages in input document
                    page_count = input_file.getNumPages()

                    # Go through all the input file pages to add a watermark to them
                    for page_number in range(page_count):
                        input_page = input_file.getPage(page_number)
                        if page_number == page_count - 1:
                            input_page.mergePage(watermark.getPage(0))
                        output_file.addPage(input_page)

                    dir = os.getcwd()
                    path = 'G:\GECOT\Análise Contábil_Tributária_Licitações\\2021'
                    os.chdir(path)
                    file = glob.glob(str(arquivo[:11]) + '*')
                    file = ''.join(file)
                    os.chdir(file)
                    input_file.close()

                    # finally, write "output" to document-output.pdf
                    with open('Análise Tributária - ' + str(arquivo[:11]) + '.pdf', "wb") as outputStream:
                        output_file.write(outputStream)
                        outputStream.close()

                    os.chdir(caminho)
                    os.remove(arquivo)


        py = 115
        self.list_check = []
        for lin in self.lista:
            self.check1 = IntVar()
            lbl1 = Checkbutton(janela, var=self.check1, onvalue=1, offvalue=0).place(y=py, x=138)
            py += 26
            self.list_check.append(self.check1)
        assinar = Button(janela, text='Assinar', font=('arial', 14, 'bold'), width=10, command=assinatura).place(x=350, y=400)


        def NotasInfo2(ev):
            # fn_id.delete(0, END)
            verinfo2 = nf_tree.focus()
            dados2 = nf_tree.item(verinfo2)
            row = dados2['values']
            print(row)
            # entr_atual.delete(0, END)
            # entr_atual.insert(0, row[3])
            # fn_id.insert(0, row[0])
            os.startfile('C:\\Users\loliveira\PycharmProjects\Python Pdf\Relatorio-pdf\pendentes' + '\\' + row[0])


        # adicionar a tela
        nf_tree.tag_configure('oddrow', background='white')
        nf_tree.tag_configure('evenrow', background='lightblue')
        inserir_tree(self.lista)
        nf_tree.bind('<Double-Button>', NotasInfo2)


if __name__=='__main__':
    janela = Tk()
    aplicacao = NotasServicos(janela)
    janela.mainloop()