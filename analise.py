from fpdf import FPDF
from tkinter import *
from tkinter import ttk

class Analise:

    def __init__(self, janela):
        self.janela = janela
        titulo = ' '
        self.janela.title(160 * titulo + 'Análise Tributária')
        self.janela.geometry('1200x680+100+20')
        self.janela.resizable(width=False, height=False)

        self.mainframe = Frame(self.janela, width=1200, height=680, relief=RIDGE, bg='sky blue')
        self.mainframe.place(x=0, y=0)

        info_frame = Frame(self.mainframe, width=1200, height=480, relief=RIDGE, bd=5)
        info_frame.place(x=0, y=0)
        self.fonte = ('arial', 12, 'bold')

        self.rotulos = ['Gerência Contratante: ', 'Nº Processo: ', 'Requisição de Compras: ', 'Código Material/Serviço:',
                   'Consta no Orçamento?', 'Sim', 'Não', '1.	Classificação Contábil:',
                   '1.	Objeto de Custos:', 'Serviço', 'Material ', 'Serviço com Fornecimento de Material']

        Label(self.mainframe, text=self.rotulos[0], font=self.fonte, bd=0).place(x=20, y=50)
        Label(self.mainframe, text=self.rotulos[1], font=self.fonte, bd=0).place(x=20, y=90)
        Label(self.mainframe, text=self.rotulos[2], font=self.fonte, bd=0).place(x=20, y=130)
        Label(self.mainframe, text=self.rotulos[4], font=self.fonte, bd=0).place(x=20, y=170)
        Label(self.mainframe, text=self.rotulos[3], font=self.fonte, bd=0).place(x=20, y=210)
        Label(self.mainframe, text=self.rotulos[7][3:], font=self.fonte, bd=0).place(x=20, y=250)
        Label(self.mainframe, text=self.rotulos[8][3:], font=self.fonte, bd=0).place(x=20, y=290)

        self.gere = Entry(info_frame, width=20, bd=4)
        self.gere.place(x=230, y=42)
        self.gere.insert(0, 'GECOT')
        self.proc = Entry(info_frame, width=20, bd=4)
        self.proc.place(x=230, y=82)
        self.proc.insert(0, 'DV-004/2021')
        self.req = Entry(info_frame, width=20, bd=4)
        self.req.place(x=230, y=122)
        self.req.insert(0, '10189080')
        self.orcam = ttk.Combobox(info_frame, font=self.fonte, width=12)
        self.orcam['values'] = ('Não', 'Sim')
        self.orcam.current(1)
        self.orcam.place(x=230, y=162)
        self.clascont = Entry(info_frame, width=20, bd=4)
        self.clascont.place(x=230, y=202)
        self.clascont.insert(0, '1011500')
        self.codmat = Entry(info_frame, width=20, bd=4)
        self.codmat.place(x=230, y=242)
        self.codmat.insert(0, '4405')
        self.objcust = Entry(info_frame, width=20, bd=4)
        self.objcust.place(x=230, y=282)
        self.objcust.insert(0, '11350')



        self.btngerar = Button(info_frame, font=self.fonte, text='Gerar PDF', bd=4, command=self.salvar).place(x=20, y=350)

    def salvar(self):

        self.pdf = FPDF(orientation='P', unit='mm', format='A4')
        self.pdf.add_page()

        self.pdf_w = 210
        self.pdf_h = 297


        self.pdf.rect(5.0, 5.0, 200.0,20.0)

        self.pdf.image('logo.jpg', x=7.0, y=7.0, h=15.0, w=50.0)
        self.pdf.line(70.0, 5.0, 70.0, 25.0)


        self.pdf.set_font('Arial', 'B', 10)
        self.pdf.set_xy(75.0, 9.0)
        self.pdf.multi_cell(w=125, h=5, txt='Análise Contábil e Tributária para Processos de Licitação e ou Contratação Direta')

        self.pdf.rect(5.0, 30.0, 200.0, 30.0)
        self.pdf.rect(5.0, 40.0, 200.0, 10.0)
        self.pdf.line(80.0, 30.0, 80.0, 60.0)
        self.pdf.line(140.0, 30.0, 140.0, 40.0)
        self.pdf.set_xy(10.0, 25.0)
        self.pdf.cell(w=40, h=20, txt=self.rotulos[0])
        self.pdf.set_xy(50.0, 25.0)
        self.pdf.cell(w=40, h=20, txt=self.gere.get())
        self.pdf.set_xy(10.0, 35.0)
        self.pdf.cell(w=40, h=20, txt=self.rotulos[1])
        self.pdf.set_xy(82.0, 22.5)
        self.pdf.cell(w=40, h=20, txt=self.rotulos[2])
        self.pdf.set_xy(82.0, 26.5)
        self.pdf.cell(w=40, h=20, txt=self.proc.get())
        self.pdf.set_xy(142.0, 22.5)
        self.pdf.cell(w=40, h=20, txt=self.rotulos[3])
        self.pdf.set_xy(142.0, 26.5)
        self.pdf.cell(w=40, h=20, txt=self.req.get())

        self.pdf.set_xy(82.0, 35.0)
        self.pdf.cell(w=40, h=20, txt='Consta no Orçamento? ')
        self.pdf.set_xy(10.0, 45.0)
        self.pdf.cell(w=40, h=20, txt='1.   Classificação Contábil:  ')
        self.pdf.set_xy(82.0, 45.0)
        self.pdf.cell(w=40, h=20, txt='1.   Objeto de Custos:  ')
        # self.pdf.set_xy(0.0, 0.0)
        self.pdf.rect(5.0, 65.0, 200.0, 225.0)
        self.pdf.line(5.0, 71.0, 205.0, 71.0)
        self.pdf.set_xy(100.0, 58.0)
        self.pdf.cell(w=40, h=20, txt='ANÁLISE')
        self.pdf.line(5.0, 80.0, 205.0, 80.0)
        self.pdf.set_xy(10.0, 65.0)

        print(self.pdf.get_y())


        self.pdf.output('teste.pdf', 'F')


if __name__=='__main__':
    janela = Tk()
    aplicacao = Analise(janela)
    janela.mainloop()