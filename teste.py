from fpdf import FPDF
from tkinter import *
from tkinter import ttk
from tkinter.tix import *
from csv import reader
from tkscrolledframe import ScrolledFrame

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
        self.tipo.current(1)
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

        self.pdf = FPDF(orientation='P', unit='mm', format='A4')
        self.pdf.add_page()

        self.pdf_w = 210
        self.pdf_h = 297

        self.pdf.rect(5.0, 5.0, 200.0, 20.0)

        self.pdf.image('logo.jpg', x=7.0, y=7.0, h=15.0, w=50.0)
        self.pdf.line(70.0, 5.0, 70.0, 25.0)

        self.pdf.set_font('Arial', 'B', 10)
        self.pdf.set_xy(75.0, 9.0)
        self.pdf.multi_cell(w=125, h=5,
                            txt='Análise Contábil e Tributária para Processos de Licitação e ou Contratação Direta')

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
        self.pdf.set_xy(10.0, self.pdf.get_y() + 5)
        self.pdf.cell(w=40, h=5, txt='Valor estimado: ')
        self.pdf.set_xy(40.0, self.pdf.get_y())
        self.pdf.cell(w=40, h=5, txt=self.valor.get())
        self.pdf.set_xy(15.0, self.pdf.get_y() + 10)

        self.pdf.set_auto_page_break(True, 20.0)

        def abre_tela():
            if self.tipo.get() == 'Serviço':
                self.tela_servicos()
            else:
                self.tela_materiais()

        self.btn_tela_serv = Button(info_frame2, font=self.fonte, text='Avançar', bd=4,
        command=abre_tela).place(x=1000, y=160)

    def tela_servicos(self):

        self.servicos = Toplevel()
        titulo = ' '
        self.servicos.title(160 * titulo + 'Serviços')
        self.servicos.geometry('1200x680+100+20')

        self.serv_frame = Frame(self.servicos, width=1200, height=680, relief=RIDGE, bd=7, bg='floral white')
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
        self.entradas = []
        self.data = [['DESCRIÇÃO', 'CÓDIGO', 'C.C']]
        # múltiplos serviços
        px = 150
        py = 240
        for i in range(10):
            for c in range(3):
                if c == 1:
                    largura = 30
                    espaco = 285
                else:
                    largura = 15
                    espaco = 150
                serv = Entry(self.serv_frame, width=largura, bd=4, font='arial')
                serv.place(x=px, y=py)
                self.entradas.append(serv)
                lista[c].append(serv)
                px += espaco
            py += 35
            px = 150

        self.pdf.multi_cell(w=180, h=5, txt=self.serv.get(1.0, 'end'))
        self.pdf.set_xy(15.0, self.pdf.get_y() + 10)


        def colar(ev):
            rows = self.servicos.clipboard_get().split('\n')
            rows.pop()
            for r, row in enumerate(rows):
                values = row.split('\t')

                ordem = [1, 0, 2]
                values = [values[i] for i in ordem]
                # self.data.append(values)
                for b, value in enumerate(values):
                    lista[b][r].delete(0, END)
                    lista[b][r].insert(0, value)


        def adicionar():
            cont2 = 0
            temp_list = []
            for lin in self.entradas:
                if lin.get() != '':
                    temp_list.append(lin.get())

                    cont2 += 1
                    if cont2 == 3:
                        lista_nova = temp_list.copy()
                        self.data.append(lista_nova)
                        temp_list.clear()
                        cont2 = 0
            for i in self.data[1:]:
                i.sort(key=len, reverse=True)
            for lin in self.entradas:
                lin.delete(0, END)



            self.pdf.set_auto_page_break(True, 20.0)

            print(self.data)
            self.dados_faltantes = []
            # def cria_tabela(data):
            self.pdf.set_xy(10, self.pdf.get_y())
            cont_lista = 0
            cont = 3
            px = 10
            py = self.pdf.get_y()
            # while self.pdf.get_y() < 274:
            for row in self.data:
                for datum in row:
                    if cont % 3 == 0:
                        self.pdf.set_font('') if cont_lista != 0 else self.pdf.set_font('Arial', 'B', 10)
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
                if py > 270:
                    self.dados_faltantes = self.data[cont_lista:]
                    self.pdf.add_page()
                    self.pdf.rect(5.0, 5.0, 200.0, 267.0)
                    break

            if self.dados_faltantes:
                self.dados_faltantes.insert(0, ['DESCRIÇÃO', 'CÓDIGO', 'C.C'])
                cont = 3
                px = 10
                py = self.pdf.get_y()
                # while self.pdf.get_y() < 274:
                for row in self.dados_faltantes:
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

            else:
                pass


        def limpar():
            self.data.clear()
            for lin in self.entradas:
                lin.delete(0, END)


        self.servicos.bind_all("<<Paste>>", colar)

        # epw = pdf.w - 2 * pdf.l_margin
        # col_width = epw / 3
        # data2 = ['CÓDIGO', 'DESCRIÇÃO', 'C.C']

        self.btngerar = Button(self.serv_frame, font=self.fonte, text='Gerar PDF', bd=4,
                               command=self.salvar).place(x=1000, y=600)

        self.btnadicionar = Button(self.serv_frame, font=self.fonte, text='Adicionar', bd=4,
                                   command=adicionar).place(x=200, y=600)

        self.btnlimpar = Button(self.serv_frame, font=self.fonte, text='Limpar campos', bd=4,
                                command=limpar).place(x=400, y=600)

        self.avancar = Button(self.serv_frame, font=self.fonte, text='Avançar', bd=4,
                               command=self.contratos).place(x=1100, y=600)

    def tela_materiais(self):
        self.materiais = Toplevel()
        titulo = ' '
        self.materiais.title(160 * titulo + 'Materiais')
        self.materiais.geometry('1200x680+100+20')

        self.mat_frame = Frame(self.materiais, width=1200, height=680, relief=RIDGE, bd=7, bg='floral white')
        self.mat_frame.place(x=0, y=0)

        # frame_1 = Frame(serv_frame, height=500, width=1190, bd=5).place(x=0, y=0)
        Label(self.mat_frame, text='Informações Tributárias', font=self.fonte, bd=0).place(x=20, y=50)



        lista_mat = [[], [], [], [], [], [], [], []]
        self.entradas_mat = []
        self.data_mat = [['CÓDIGO', 'DESCRIÇÃO', 'IVA', 'NCM', 'ICMS', 'IPI', 'PIS', 'COFINS']]
        self.mat_check = []

        px = 550
        for f, zebra in enumerate(self.data_mat[0][2:]):
            print(zebra)
            self.mat_check.append(IntVar())
            check2 = Checkbutton(self.mat_frame, variable=self.mat_check[f], onvalue=1,
                                 offvalue=0, bd=0, font=('arial', 14))
            check2.place(x=px, y=170)
            px = px + 75 if f == 1 else + px + 65


        # múltiplos materiais
        px = 150
        py = 200
        for i in range(10):
            for c in range(8):
                if c == 1:
                    largura = 30
                    espaco = 280
                elif c == 0:
                    largura = 10
                    espaco = 100
                elif c == 3:
                    largura = 10
                    espaco = 100
                else:
                    largura = 5
                    espaco = 55
                mater = Entry(self.mat_frame, width=largura, bd=4, font='arial')
                mater.place(x=px, y=py)
                self.entradas_mat.append(mater)
                lista_mat[c].append(mater)
                px += espaco
            py += 30
            px = 150




        def colar(ev):
            rows = self.materiais.clipboard_get().split('\n')
            rows.pop()
            for r, row in enumerate(rows):
                values = row.split('\t')
                if len(values) == 1:
                    for b, value in enumerate(values):
                        if lista_mat[0] == '':
                            b += 1
                            lista_mat[b][r].delete(0, END)
                            lista_mat[b][r].insert(0, value)
                    else:
                        for b, value in enumerate(values):
                            lista_mat[b][r].delete(0, END)
                            lista_mat[b][r].insert(0, value)
        def adicionar():
            cont2 = 0
            mat_list = []
            for lin in self.entradas_mat:
                if lin.get() != '':
                    mat_list.append(lin.get())

                    cont2 += 1
                    if cont2 == 8:
                        lista_nova_mat = mat_list.copy()
                        self.data_mat.append(lista_nova_mat)
                        mat_list.clear()
                        cont2 = 0
            for lin in self.entradas_mat:
                lin.delete(0, END)

            self.pdf.set_auto_page_break(True, 20.0)

            print(self.data_mat)
            self.dados_faltantes = []
            # def cria_tabela(data):
            self.pdf.set_xy(10, self.pdf.get_y())
            cont_lista = 0
            cont = 1
            px = 10
            py = self.pdf.get_y()
            # while self.pdf.get_y() < 274:
            for row in self.data_mat:
                for datum in row:
                    if cont == 1:
                        self.pdf.set_font('') if cont_lista != 0 else self.pdf.set_font('Arial', 'B', 10)
                        self.pdf.set_xy(px, py)
                        self.pdf.multi_cell(w=20, h=5, txt=datum, border=1)
                    elif cont == 2:
                        self.pdf.set_xy(px + 20, py)
                        self.pdf.multi_cell(w=75, h=5, txt=datum, border=1)
                    elif cont == 3:
                        self.pdf.set_xy(px + 95, py)
                        self.pdf.multi_cell(w=10, h=5, txt=datum, border=1)
                    elif cont == 4:
                        self.pdf.set_xy(px + 105, py)
                        self.pdf.multi_cell(w=20, h=5, txt=datum, border=1)
                    else:
                        self.pdf.set_xy(px + 125, py)
                        px += 16
                        self.pdf.multi_cell(w=16, h=5, txt=datum, border=1)
                    cont += 1
                px = 10
                py = self.pdf.get_y()
                cont = 1
                cont_lista += 1
                if py > 270:
                    self.dados_faltantes = self.data_mat[cont_lista:]
                    self.pdf.add_page()
                    self.pdf.rect(5.0, 5.0, 200.0, 267.0)
                    break

            if self.dados_faltantes:
                self.dados_faltantes.insert(0, ['DESCRIÇÃO', 'CÓDIGO', 'C.C'])
                cont = 3
                px = 10
                py = self.pdf.get_y()
                # while self.pdf.get_y() < 274:
                for row in self.dados_faltantes:
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

            else:
                pass

        def limpar():
            self.data.clear()
            for lin in self.entradas:
                lin.delete(0, END)

        self.materiais.bind_all("<<Paste>>", colar)

        # epw = pdf.w - 2 * pdf.l_margin
        # col_width = epw / 3
        # data2 = ['CÓDIGO', 'DESCRIÇÃO', 'C.C']

        self.btngerar = Button(self.mat_frame, font=self.fonte, text='Gerar PDF', bd=4,
                               command=self.salvar).place(x=1000, y=600)

        self.btnadicionar = Button(self.mat_frame, font=self.fonte, text='Adicionar', bd=4,
                                   command=adicionar).place(x=200, y=600)

        self.btnlimpar = Button(self.mat_frame, font=self.fonte, text='Limpar campos', bd=4,
                                command=limpar).place(x=400, y=600)

        self.avancar = Button(self.mat_frame, font=self.fonte, text='Avançar', bd=4,
                              command=self.contratos).place(x=1100, y=600)

    def contratos(self):
        try:
            self.servicos.destroy()
        except:
            self.materiais.destroy()

        self.contratos = Toplevel()
        titulo = ' '
        self.contratos.title(160 * titulo + 'Serviços')
        self.contratos.geometry('1200x680+100+20')

        # self.cont_frame = Frame(self.contratos, width=1200, height=680, relief=RIDGE, bd=7, bg='floral white')
        # self.cont_frame.pack(fill=BOTH, expand=1)

        sf = ScrolledFrame(self.contratos, width=940, height=480)
        sf.pack(side="top", expand=1, fill="both")

        # Bind the arrow keys and scroll wheel
        sf.bind_arrow_keys(self.contratos)
        sf.bind_scroll_wheel(self.contratos)

        # Create a frame within the ScrolledFrame
        inner_frame = sf.display_widget(Frame)

        # Add a bunch of widgets to fill some space
        self.nomes = ['2.3.7.2.', '2.3.7', '2.3.7.1', '2.3.7.1.', '15.1', '2.3.7.2', '2.3.7.3', '2.3.7.4', '6.8.2']
        self.var_check = []
        linha = 3
        for check in range(9):
            self.var_check.append(IntVar())
            check1 = Checkbutton(inner_frame, text=self.nomes[check], variable=self.var_check[check], onvalue=1, offvalue=0, bd=0, font=('arial', 14))
            check1.grid(row=linha, column=1, padx=70)
            linha += 2

        self.info_lbl = Label(inner_frame, bd=0, bg='floral white', font=('arial', 14),
                              text='Informações  contratuais: ').grid(row=2, column=2, pady=30)

        with open('texto.txt', 'r', encoding='latin-1') as read_obj:
            csv_reader = read_obj.readlines()
            self.infos = []
            linha = 2
            for row in csv_reader:
                Label(inner_frame).grid(row=linha, column=2)
                self.campo = Text(inner_frame, height=5, width=100)
                self.campo.grid(row=linha + 1, column=2)
                self.campo.insert('end', row)
                self.infos.append(self.campo)
                linha += 2




        self.btngerar = Button(inner_frame, font=self.fonte, text='Gerar PDF', bd=4,
                               command=self.salvar).grid(row=27, column=2, pady=30)



    def salvar(self):
        print(self.pdf.get_y())
        for i, item in enumerate(self.var_check):
            if item.get() == 1:
                self.pdf.set_xy(15.0, self.pdf.get_y()+5)
                self.pdf.multi_cell(w=180, h=5, align='L', txt=self.infos[i].get(1.0, 'end'))
                if self.pdf.get_y() > 270:
                    self.pdf.add_page()
                    self.pdf.rect(5.0, 5.0, 200.0, 267.0)

        self.pdf.output('teste.pdf', 'F')



if __name__=='__main__':
    janela = Tk()
    aplicacao = Analise(janela)
    janela.mainloop()

