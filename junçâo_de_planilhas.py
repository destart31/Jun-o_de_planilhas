import openpyxl
import openpyxl.styles
import openpyxl.workbook
import pandas as pd

# filename = 'data/sales_one.xlsx'

# workbook = openpyxl.load_workbook(filename)
# # print(workbook.worksheets)
# worksheet = workbook['Sales1']
# # Existe algumas formas de ler a informação da tabela

# # 1 forma trasendo a posição
# # worksheet['D2'].value # 11.95

# # 2 forma criando uma lista e trazendo a linha e coluna
# # worksheet_mat = list(worksheet.columns)
# # worksheet_mat[3][1].value

# # 3 forma usando o pandas
# df = pd.read_excel(filename, index_col= 0)
# # df.loc(0, 'Price Each') # retorna o 11.95 da posição 0 da coluna Price Each
# print(df)

filename_one = 'data/Sales_one.xlsx'
filename_two = 'data/Sales_two.xlsx'

df_one = pd.read_excel(filename_one, index_col=0)
# print(df_one)

df_two = pd.read_excel(filename_two, index_col=0)
# print(df_two)

# Junta as duas planilhas e retorna as informações REPETIDAS em ambas as 
# duas planilhas
# Neste caso são 150 informações que existem em ambas planilhas
# por padrão o how ='inner' que faz pegar as informações repetidas
# df_merge = df_one.merge(df_two)
# print(df_merge)

# test para ver se essa informação realmete existe em ambas as tabelas
# retorno da infromação da linha 850 que é a mesma que a linha 0 da df_two
# print(df_one[df_one['Order Date'] == '04/10/19 18:32'])

# how='outer' faz retornar as informações unicas ou seja de 2000 linhas 
# retorna 1850 retirando as 150 repetidas
df_merge = df_one.merge(df_two, how ='outer')
print(df_merge)

# Salvar um novo arquivo com a junção das duas planilhas
filename = 'data/sales_merge.xlsx'
df_merge.to_excel(filename, sheet_name='Sales merged')

# Escrevendo na planilha excel

workbook = openpyxl.load_workbook(filename)
worksheet = workbook['Sales merged']

# pegar o tamanho da planilha
# para escrever, ou seja, o que for escrito vai para a linha 1852
last_index = str(len(df_merge) + 2)

# escrendo e aplicando font em negrito
worksheet['A' + last_index] = 'Total of sales'
worksheet['A' + last_index].font = openpyxl.styles.Font(bold=True)

# Soma do produto
worksheet['D' + last_index] = '=SUMPRODUCT(C2:C{}, D2:D{})'.format(str(int(last_index) - 1),
                                                                   str(int(last_index) - 1))
workbook.save(filename)