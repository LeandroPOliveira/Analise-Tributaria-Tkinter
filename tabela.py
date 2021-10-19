import pandas as pd
from openpyxl.reader.excel import load_workbook
import pickle


# lista = ['01/01/2021', 'GETIN', 'DV-015/2021', '10199999', 'Sim', '11330', '1', 'Serviço de BUNDA.\n', 'R$ 50.000,00', '\n', [['CÓDIGO', 'DESCRIÇÃO', 'IVA', 'NCM', 'ICMS', 'IPI', 'PIS', 'COFINS'], ['2450', 'BUNDA', 'Y0', '7400.00', '18', '5', '7,6', '1,65']], '0', '', 'Y0', [['DESCRIÇÃO', 'CÓDIGO', 'C.C'], ['bunda', '1000', '4598']], '0', '\n', 'Obs 1: Caso o BUNDA possua alguma especificidade que implique tratamento tributário diverso do exposto acima, ou seja do regime tributário "SIMPLES NACIONAL" deverá  apresentar documentação hábil que comprove sua condição peculiar, a qual será alvo de análise prévia pela GECOT.\n', 'Obs 2: Essa BUNDA não é exaustiva, podendo sofrer alterações no decorrer do processo de contratação em relação ao produto/serviço.\n']
# lista2 = ['10/08/2021', 'GEOPE', 'DV-124/2021', '10193535', 'Sim', '11440', '1', 'Serviço de Segurança Eletrônica 24 horas e Pronta Intervenção na EO e ECP São Carlos, com o fornecimento e instalação de equipamentos.\n', 'R$ 49.915,00', '\n', [['CÓDIGO', 'DESCRIÇÃO', 'IVA', 'NCM', 'ICMS', 'IPI', 'PIS', 'COFINS']], '0', '', 'ZJ', [['DESCRIÇÃO', 'CÓDIGO', 'C.C']], '0', '\n', 'Obs 1: Caso o fornecedor possua alguma especificidade que implique tratamento tributário diverso do exposto acima, ou seja do regime tributário "SIMPLES NACIONAL" deverá  apresentar documentação hábil que comprove sua condição peculiar, a qual será alvo de análise prévia pela GECOT.\n', 'Obs 2: Essa Análise não é exaustiva, podendo sofrer alterações no decorrer do processo de contratação em relação ao produto/serviço.\n']
# # book = load_workbook('G:\GECOT\Análise Contábil_Tributária_Licitações\\2021\\1Pendentes\\Base.xlsx')
# # # writer = pd.ExcelWriter('G:\GECOT\Análise Contábil_Tributária_Licitações\\2021\\1Pendentes\\Base.xlsx',
# # #                         engine='openpyxl')
# # # writer.book = book
# # ws = book.worksheets[0]
# # tabela = pd.DataFrame(lista).transpose()
#
# with open("G:\GECOT\Análise Contábil_Tributária_Licitações\\2021\\1Pendentes\\Base.txt", "wb") as fp:   #Pickling
#     pickle.dump(lista, fp)
#     pickle.dump(lista2, fp)



with open("G:\GECOT\Análise Contábil_Tributária_Licitações\\2021\\1Pendentes\\Base.txt", "rb") as fp:   # Unpickling
    b = pickle.load(fp)

for i in b:
    print(i)

