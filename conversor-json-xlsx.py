import json
import xlsxwriter


# Inserção titulo de colunas
def insert_title_columns(worksheet):
    insert_in_table(worksheet, 0, 0, 'Código')
    insert_in_table(worksheet, 0, 1, 'Funcionário')
    insert_in_table(worksheet, 0, 2, 'Dependente')
    insert_in_table(worksheet, 0, 3, 'Parentesco')
    insert_in_table(worksheet, 0, 4, 'Presente')


# Obtendo acesso ao arqivo JSON
def read_file_json(file):
    with open(file, encoding="utf8") as json_file:
        return json.load(json_file)


# controle dos dados na tabela 
def manipulation_data(worksheet):
    column = 0
    row = 1

    file_json = read_file_json("teste-146d5-export.json")

    for data in file_json['employee'].items():
        employee = data[1]
        dependents = employee['dependents']

        insert_in_table(worksheet, row, column, employee['id'])
        insert_in_table(worksheet, row, column + 1, employee['name'])

        for num, dependent in enumerate(dependents):

            if num != 0:
                insert_in_table(worksheet, row, column, employee['id'])
                insert_in_table(worksheet, row, column + 1, employee['name'])

            insert_in_table(worksheet, row, column + 2, dependent['name'])
            insert_in_table(worksheet, row, column + 3, dependent['kinship'])
            insert_in_table(worksheet, row, column + 4, dependent['present'])

            row += 1


# insere os dados na tabela
def insert_in_table(worksheet, row, column, value):
    worksheet.write(row, column, value)


# inicia e cria os arquivos
def create_xls_file(name_file):
    workbook = xlsxwriter.Workbook(name_file)
    worksheet = workbook.add_worksheet('Dados')

    insert_title_columns(worksheet)
    manipulation_data(worksheet)

    workbook.close()


# principal
def main():
    create_xls_file('Dados Dia da Familia.xlsx')


main()
