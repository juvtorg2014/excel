import os
import pandas as pd
import pip


def find_row(df) -> int:
    """ Поиск строки с названиями нужных колонок"""
    number = 0
    while True:
        if 'Серия номенклатуры' in df.iloc[number].values:
            break
        else:
            number += 1
    return number

def find_nun(df) -> list:
    """ Поиск пустых колонок для удаления"""
    deL_list = []
    for item in range(len(df.columns.values) - 1):
        if type(df.columns[item])  is not str :
            deL_list.append(item)
    return deL_list


def comparison(file_source, file_receiver):
    pip.main(['install', 'openpyxl'])
    
    source = pd.read_excel(file_source)
    row = find_row(source)
    del_list = [x for x in range(row)]
    source = source.drop(index=del_list)
    source.columns = source.iloc[0]
    source = source[1:]
    num_col = find_nun(source)
    source.drop(source.columns[num_col], axis=1, inplace=True)
    source = source[:-1]
    source = source.rename(columns={'Серия номенклатуры': 'Серия', 'Сумма НДС вознаграждения': 'НДС',
                                    'Сумма комиссионного вознаграждения ': 'Комиссия', })
    new_source = source[['Серия', 'Количество', 'Выручка', 'Комиссия', 'НДС']].copy(deep=False)
    new_source['Серия'] = new_source['Серия'].astype(str)
    new_source['Серия'] = new_source['Серия'].str.strip()
    new_source = new_source.sort_values(by=['Серия', 'Количество', 'Выручка'])
    
    receiver = pd.read_excel(file_receiver)
    receiver = receiver.rename(columns={'Серия номенклатуры': 'Серия', 'Сумма НДС вознаграждения(RUB)': 'НДС',
                                        'Сумма (RUB)': 'Выручка', 'Вознаграждение(RUB)': 'Комиссия'})
    new_reciever = receiver[['Серия', 'Количество', 'Выручка', 'Комиссия', 'НДС']].copy(deep=False)
    new_reciever['Серия'] = new_reciever['Серия'].astype(str)
    new_reciever['Серия'] = new_reciever['Серия'].str.strip()
    new_reciever = new_reciever.sort_values(by=['Серия', 'Количество', 'Выручка'])
    
    new_source.to_csv(file_source.replace('.xlsx', '.csv'), index=False, header=True, encoding='cp1251', sep=';')
    new_reciever.to_csv(file_receiver.replace('.xlsx', '.csv'), index=False, header=True, encoding='cp1251', sep=';')

    dif_col = new_source.set_index('Серия').subtract(new_reciever.set_index('Серия'))
    dif_table = dif_col[(dif_col['Выручка'] != float(0)) | (dif_col['Комиссия'] != float(0))
                                                     | (dif_col['Количество'] != float(0))]

    dif_table.to_csv(os.getcwd() + '\\' + 'Товары.csv', index=True, header=True, encoding='cp1251', sep=';')
    print(dif_table)


if __name__ == '__main__':
    file_1 = input("Введите имя файла-отчета контрагента без расширения\n")
    file_1 = os.path.abspath(file_1 + '.xlsx')
    file_2 = input("Введите имя второго файла-документа без расширения\n")
    file_2 = os.path.abspath(file_2 + '.xlsx')
    comparison(file_1, file_2)
