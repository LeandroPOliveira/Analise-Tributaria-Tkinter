# Import FPDF class
from fpdf import FPDF

# Create instance of FPDF class
# Letter size paper, use inches as unit of measure
pdf = FPDF()

# Add new page. Without this you cannot create the document.
pdf.add_page()

# Remember to always put one of these at least once.
pdf.set_font('Times', '', 10.0)

# Effective page width, or just epw
epw = pdf.w - 2 * pdf.l_margin

# Set column width to 1/4 of effective page width to distribute content
# evenly across table and page
col_width = epw / 3

# Since we do not need to draw lines anymore, there is no need to separate
# headers from data matrix.

data = [['DESCRIÇÃO', 'CÓDIGO', 'C.C'],
        ['REMOÇÃO DE TELHAS DA EDIFICAÇÃO EXISTENTE', '3002140', '4508'],
        ['AMPLIAÇÃO CIVIL DA EDIFICAÇÃO EXISTENTE - FUNDAÇÃO', '3002141', '4508'],
        ['AMPLIAÇÃO CIVIL DA EDIFICAÇÃO EXISTENTE - PISO CONCRETO', '3002141', '4508'],
        ['ACABAMENTO IMPERMEABILIZAÇÃO E PINTURA: ESQUADRIAS, PISO, PAREDE -INTERNA E EXTERNA E CALÇADAS', '3002146',
        '4508']]



# pdf.set_font('Times', 'B', 14.0)
# pdf.cell(epw, 0.0, 'With more padding', align='C')
# pdf.set_font('Times', '', 10.0)
# pdf.ln(0.5)
data2 = ['CÓDIGO', 'DESCRIÇÃO', 'C.C']

pdf.set_xy(10, 20)
cont = 3
px = 10
py = 20
for row in data:
    for datum in row:
        if cont % 3 == 0:
            pdf.set_xy(px + 20, py)
            pdf.multi_cell(w=150, h=5, txt=datum, border=1)
        elif cont % 4 == 0:
            atual = pdf.get_y() - py
            print(atual)
            pdf.set_xy(px, py)
            pdf.multi_cell(w=20, h=atual, txt=datum, border=1)
        else:
            atual = pdf.get_y() - py
            pdf.set_xy(px+170, py)
            pdf.multi_cell(w=20, h=atual, txt=datum, border=1)
        cont += 1
    px = 10
    py += 5
    cont = 3

pdf.output('table-using-cell-borders.pdf', 'F')

