from fpdf import FPDF
from tkinter import *
from tkinter import ttk
from operator import itemgetter

class Analise:

    def __init__(self, janela):
        self.janela = janela
        titulo = ' '
        self.janela.title(160 * titulo + 'Análise Tributária')
        self.janela.geometry('1200x680+100+20')
        self.janela.resizable(width=False, height=False)

        self.mainframe = Frame(self.janela, width=1200, height=680, relief=RIDGE, bg='sky blue')
        self.mainframe.place(x=0, y=0)

        info_frame = Frame(self.mainframe, width=1200, height=260, relief=RIDGE, bd=7)
        info_frame.place(x=0, y=30)
        info_frame2 = Frame(self.mainframe, width=1200, height=260, relief=RIDGE, bd=7)
        info_frame2.place(x=0, y=300)
        self.fonte = ('arial', 12, 'bold')

        self.rotulos = ['Gerência Contratante: ', 'Nº Processo GECBS: ', 'Requisição de Compras: ', 'Código Material/Serviço:',
                   'Consta no Orçamento?', 'Sim', 'Não', '1.	Classificação Contábil:',
                   '1.	Objeto de Custos:', 'Serviço', 'Material ', 'Serviço com Fornecimento de Material']

        Label(info_frame, text='Gerência Contratante: ', font=self.fonte, bd=0).place(x=20, y=50)
        Label(info_frame, text='Nº Processo GECBS: ', font=self.fonte, bd=0).place(x=500, y=50)
        Label(info_frame, text='Requisição de Compras: ', font=self.fonte, bd=0).place(x=20, y=90)
        Label(info_frame, text='Consta no Orçamento?', font=self.fonte, bd=0).place(x=500, y=90)
        Label(info_frame, text='Objeto de Custos:', font=self.fonte, bd=0).place(x=20, y=130)
        Label(info_frame, text='Tipo de Análise: ', font=self.fonte, bd=0).place(x=500, y=130)
        Label(info_frame2, text='Código Material/Serviço:', font=self.fonte, bd=0).place(x=20, y=20)
        Label(info_frame2, text='Classificação Contábil ', font=self.fonte, bd=0).place(x=500, y=20)
        # Label(info_frame, text='Valor Estimado: ', font=self.fonte, bd=0).place(x=520, y=380)
        # Label(info_frame, text='Código IVA: ', font=self.fonte, bd=0).place(x=520, y=420)

        self.gere = Entry(info_frame, width=20, bd=4)
        self.gere.place(x=230, y=48)
        self.gere.insert(0, 'GECOT')
        self.proc = Entry(info_frame, width=20, bd=4)
        self.proc.place(x=700, y=50)
        self.proc.insert(0, 'DV-004/2021')
        self.req = Entry(info_frame, width=20, bd=4)
        self.req.place(x=230, y=88)
        self.req.insert(0, '10189080')
        self.orcam = ttk.Combobox(info_frame, font=self.fonte, width=12)
        self.orcam['values'] = ('Não', 'Sim')
        self.orcam.current(1)
        self.orcam.place(x=700, y=88)
        self.objcust = Entry(info_frame, width=20, bd=4)
        self.objcust.place(x=230, y=128)
        self.objcust.insert(0, 'PEP RSG.01.001.001.01')
        self.tipo = ttk.Combobox(info_frame, font=self.fonte, width=25)
        self.tipo['values'] = ('Serviço', 'Material', 'Serviço com Fornec de Mat')
        self.tipo.current(0)
        self.tipo.place(x=700, y=128)
        self.codmat = Entry(info_frame2, width=20, bd=4)
        self.codmat.place(x=230, y=20)
        self.codmat.insert(0, '4405')
        self.clascont = Entry(info_frame2, width=20, bd=4)
        self.clascont.place(x=700, y=20)
        self.clascont.insert(0, '1011500')
        self.var1 = IntVar()
        self.multiserv = Checkbutton(info_frame2, text='Múltiplos Serviços', variable=self.var1, onvalue=1,
                                     offvalue=0, font=('arial', 12))
        self.multiserv.place(x=100, y=55)


        self.objeto = Text(info_frame, width=50, height=5, bd=4, font='arial')
        self.objeto.place(x=660, y=350)
        self.objeto.insert('end', 'Serviços de reparo da infraestrutura de cabeamento de dados e telefonia do prédio '
                              'administrativo da GasBrasiliano.')
        self.valor = Entry(info_frame, width=20, bd=4)
        self.valor.place(x=660, y=380)
        self.valor.insert('end', 'R$ 44.072,64')
        self.iva = Entry(info_frame, width=20, bd=4)
        self.iva.place(x=660, y=420)
        self.iva.insert(0, 'ZJ')





        self.btn_tela_serv = Button(info_frame2, font=self.fonte, text='Avançar', bd=4,
                                    command=self.tela_servicos).place(x=1000, y=160)

    def tela_servicos(self):

        servicos = Toplevel()
        titulo = ' '
        servicos.title(160 * titulo + 'Serviços')
        servicos.geometry('1200x680+100+20')

        self.serv_frame = Frame(servicos, width=1200, height=680, relief=RIDGE, bd=7, bg='floral white')
        self.serv_frame.place(x=0, y=0)

        # frame_1 = Frame(serv_frame, height=500, width=1190, bd=5).place(x=0, y=0)
        Label(self.serv_frame, text='Codigo Serviço', font=self.fonte, bd=0).place(x=20, y=50)

        self.serv = Text(self.serv_frame, width=70, height=7, bd=4, font='arial')
        self.serv.place(x=150, y=50)
        self.serv.insert('end', '17.01 Assessoria ou consultoria de qualquer natureza, não contida em outros itens '
                                'desta lista; análise, exame, pesquisa, coleta, compilação e fornecimento de dados e '
                                'informações de qualquer natureza, inclusive cadastro e similares.\n'
                                '\n'
                                'Sobre os serviços haverá as seguintes retenções tributárias: \nIR e PIS/COFINS/CSLL'
                                'IRRF: serviços constam no artigo 714 RIR/2018.\n'
                                'PCC:  serviços constam na IN SRF nº 459/2004, artigo 1º, § 2º.')

        lista = [[], [], []]
        self.data = [['DESCRIÇÃO', 'CÓDIGO', 'C.C']]
        # múltiplos serviços
        px = 30
        py = 240
        for i in range(10):
            for c in range(3):
                serv = Entry(self.serv_frame, width=15, bd=4, font='arial')
                serv.place(x=px, y=py)
                lista[c].append(serv)
                px += 150
            py += 35
            px = 30



        def colar(ev):
            rows = servicos.clipboard_get().split('\n')
            for r, row in enumerate(rows):
                values = row.split('\t')
                ordem = [1, 0, 2]
                values = [values[i] for i in ordem]
                self.data.append(values)
                for b, value in enumerate(values):
                    lista[b][r].delete(0, END)
                    lista[b][r].insert(0, value)



        def limpar():
            self.data.clear()
            servicos.clipboard_clear()


        self.btnlimpar = Button(self.serv_frame, font=self.fonte, text='Limpar campos', bd=4,
                               command=limpar).place(x=500, y=500)

        servicos.bind_all("<<Paste>>", colar)

        # epw = pdf.w - 2 * pdf.l_margin
        # col_width = epw / 3
        # data2 = ['CÓDIGO', 'DESCRIÇÃO', 'C.C']








        self.btngerar = Button(self.serv_frame, font=self.fonte, text='Gerar PDF', bd=4,
                               command=self.salvar).place(x=100, y=600)


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

        self.pdf.rect(5.0, 30.0, 200.0, 25.0)
        self.pdf.line(5.0, 40.0, 205.0, 40.0)
        self.pdf.line(88.0, 30.0, 88.0, 55.0)
        # self.pdf.line(140.0, 30.0, 140.0, 40.0)
        self.pdf.set_xy(10.0, 25.0)
        self.pdf.cell(w=40, h=20, txt='Gerência Contratante:')
        self.pdf.set_xy(50.0, 25.0)
        self.pdf.cell(w=40, h=20, txt=self.gere.get())
        self.pdf.set_xy(90.0, 22.5)
        self.pdf.cell(w=40, h=20, txt='N° do Processo GECBS:')
        self.pdf.set_xy(90.0, 26.5)
        self.pdf.cell(w=40, h=20, txt=self.proc.get())
        self.pdf.set_xy(142.0, 22.5)
        self.pdf.cell(w=40, h=20, txt='Requisição de Compras: ')
        self.pdf.set_xy(142.0, 26.5)
        self.pdf.cell(w=40, h=20, txt=self.req.get())
        self.pdf.set_xy(10.0, 35.0)
        self.pdf.cell(w=40, h=20, txt='Objeto de Custos:')
        self.pdf.set_xy(43.0, 42.5)
        self.pdf.multi_cell(w=45, h=5, align='L', txt=self.objcust.get())
        self.pdf.set_xy(90.0, 35.0)
        self.pdf.cell(w=40, h=20, txt='Consta no Orçamento? ')
        self.pdf.set_xy(148.0, 35.0)
        self.pdf.cell(w=40, h=20, txt='Sim')
        self.pdf.set_xy(175.0, 35.0)
        self.pdf.cell(w=40, h=20, txt='Não')
        if self.orcam.get() == 'Sim':
            self.pdf.set_xy(138.0, 35.0)
            self.pdf.cell(w=40, h=20, txt='X')
        else:
            self.pdf.set_xy(169.0, 35.0)
            self.pdf.cell(w=40, h=20, txt='X')
        self.pdf.rect(5.0, 60.0, 200.0, 225.0)
        self.pdf.line(5.0, 66.0, 205.0, 66.0)
        self.pdf.set_xy(100.0, 53.0)
        self.pdf.cell(w=40, h=20, txt='ANÁLISE')
        self.pdf.line(5.0, 75.0, 205.0, 75.0)
        self.pdf.set_xy(30.0, 60.5)
        self.pdf.cell(w=40, h=20, txt='Serviço')
        self.pdf.set_xy(90.0, 60.5)
        self.pdf.cell(w=40, h=20, txt='Material')
        self.pdf.set_xy(135.0, 60.5)
        self.pdf.cell(w=40, h=20, txt='Serviço com Fornecimento de Material')
        if self.tipo.current() == 0:
            self.pdf.set_xy(25.0, 60.5)
            self.pdf.cell(w=40, h=20, txt='X')
        elif self.tipo.current() == 1:
            self.pdf.set_xy(85.0, 60.5)
            self.pdf.cell(w=40, h=20, txt='X')
        else:
            self.pdf.set_xy(130.0, 65.0)
            self.pdf.cell(w=40, h=20, txt='X')
        self.pdf.set_xy(10.0, 77.0)
        self.pdf.cell(w=40, h=5, txt='Objeto: ')
        self.pdf.set_xy(30.0, 77.0)
        self.pdf.multi_cell(w=160, h=5, txt=self.objeto.get(1.0, 'end'))
        self.pdf.set_xy(10.0, self.pdf.get_y()+5)
        self.pdf.cell(w=40, h=5, txt='Valor estimado: ', border=1)
        self.pdf.set_xy(40.0, self.pdf.get_y())
        self.pdf.cell(w=40, h=5, txt=self.valor.get())
        self.pdf.set_xy(15.0, self.pdf.get_y()+10)
        self.pdf.multi_cell(w=180, h=5, txt=self.serv.get(1.0, 'end'))
        self.pdf.set_xy(15.0, self.pdf.get_y() + 10)

        self.pdf.set_auto_page_break(True, 23.0)

        print(self.pdf.get_y())

        def cria_tabela(data):
            self.pdf.set_xy(10, self.pdf.get_y())
            cont_lista = 0
            cont = 3
            px = 10
            py = self.pdf.get_y()
            # while self.pdf.get_y() < 274:
            for row in data:
                for datum in row:
                    if cont % 3 == 0:
                        self.pdf.set_xy(px + 20, py)
                        self.pdf.multi_cell(w=150, h=5, txt=datum, border=1)
                    elif cont % 4 == 0:
                        atual = self.pdf.get_y() - py
                        self.pdf.set_xy(px, py)
                        self.pdf.multi_cell(w=20, h=atual, txt=datum, border=1)
                    else:
                        atual = self.pdf.get_y() - py
                        self.pdf.set_xy(px + 170, py)
                        self.pdf.multi_cell(w=20, h=atual, txt=datum, border=1)
                    cont += 1
                px = 10
                py = self.pdf.get_y()
                cont = 3
                cont_lista += 1
                print(py)
                if py > 270:
                    self.dados_faltantes = data[cont_lista:]
                    self.pdf.add_page()
                    break

            cria_tabela(self.data)

            if self.dados_faltantes:
                self.cria_tabela(self.dados_faltantes)
            else:
                pass



        print(self.dados_faltantes)

            # self.pdf2 = self.pdf.add_page()
            # self.pdf2.cell(w=40, h=5, txt='Objeto: ')
            # self.pdf.set_xy(40.0, 50.0)



        self.pdf.output('teste.pdf', 'F')



if __name__=='__main__':
    janela = Tk()
    aplicacao = Analise(janela)
    janela.mainloop()

