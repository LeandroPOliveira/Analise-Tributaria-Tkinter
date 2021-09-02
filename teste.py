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

data = [['CÓDIGO', 'DESCRIÇÃO', 'C.C'],
        ['3002140', 'REMOÇÃO DE TELHAS DA EDIFICAÇÃO EXISTENTE', '4508'],
        ['3002141', 'AMPLIAÇÃO CIVIL DA EDIFICAÇÃO EXISTENTE - FUNDAÇÃO', '4508'],
        ['3002141', 'AMPLIAÇÃO CIVIL DA EDIFICAÇÃO EXISTENTE - PISO CONCRETO', '4508'],
        ['3002146', 'ACABAMENTO IMPERMEABILIZAÇÃO E PINTURA: ESQUADRIAS, PISO, PAREDE -INTERNA E EXTERNA E CALÇADAS',
        '4508']]



# pdf.set_font('Times', 'B', 14.0)
# pdf.cell(epw, 0.0, 'With more padding', align='C')
# pdf.set_font('Times', '', 10.0)
# pdf.ln(0.5)


px = pdf.set_xy(10, 20)
for row in data:
    cont = 3
    px = pdf.set_xy(10, 20)
    for datum in row:
        # Enter data in colums
        if cont % 3 == 0 or cont % 5 == 0:

            pdf.multi_cell(w=20, h=5, txt=datum, border=1)
        else:
            px = pdf.set_xy(pdf.get_x(), pdf.get_y())
            pdf.multi_cell(w=120, h=5, txt=datum, border=1)
        pdf.set_x(pdf.get_x()+20 if cont % 3 == 0 or cont % 5 == 0 else pdf.set_x(pdf.get_x()+150))
        cont += 1







pdf.output('table-using-cell-borders.pdf', 'F')

