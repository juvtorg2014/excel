import os
import pandas as pd
import openpyxl

NAME_1 = 'Количество'
NAME_2 = 'Выручка'
NAME_3 = 'Сумма комиссионного вознаграждения'
NUMBER_STROK = 5
HEAD = ['Поставщик, Количество, Выручка, Комиссия']


def comparison(file_source, file_receiver):
    source = pd.read_excel(file_source)
    source = source.drop(index=[0, 1, 2, 3, 4])
    source = source.rename(columns={'Unnamed: 0': 'Поставщик', 'Unnamed: 3': 'Количество',
                                    'Unnamed: 5': 'Выручка', 'Unnamed: 7': 'Комиссия'})
    new_source = source[['Поставщик', 'Количество', 'Выручка', 'Комиссия']].copy(deep=False)
    new_source = new_source[:-1]
    new_source['Поставщик'] = new_source['Поставщик'].str.strip()
    receiver = pd.read_excel(file_receiver)
    col_1 = receiver.columns[2]
    col_2 = receiver.columns[4]
    col_3 = receiver.columns[6]
    col_4 = receiver.columns[5]
    receiver = receiver.rename(columns={col_1: 'Поставщик',col_2: 'Количество',col_3: 'Выручка',col_4: 'Комиссия'})
    new_reciever = receiver[['Поставщик', 'Количество', 'Выручка', 'Комиссия']].copy(deep=False)
    new_reciever['Поставщик'] = new_reciever['Поставщик'].str.strip()
    # new_source = new_source.reindex(range(len_dataframe), method='ffill')
    new_source.to_csv('Отчет.csv', index=False, header=True, encoding='cp1251', sep=';')
    new_reciever.to_csv('Документы.csv', index=False, header=True, encoding='cp1251', sep=';')

    dif_col = new_source.set_index('Поставщик').subtract(new_reciever.set_index('Поставщик'))
    dif_table = dif_col[(dif_col['Количество'] != float(0)) | (dif_col['Выручка'] != float(0))
                                             | (dif_col['Комиссия'] != float(0))]
    # new = pd.concat([new_source,new_reciever]).drop_duplicates(keep=False)
    # diff = new_source.compare(new_reciever, align_axis=1)
    dif_table.to_csv('Разница.csv', index=True, header=True, encoding='cp1251', sep=';', float_format='{15.2f}')
    print(dif_table)




if __name__ == '__main__':
    file_1 = input("Введите имя файла отчета без расширения:\n")
    file_1 = os.path.abspath(file_1 + '.xlsx')
    file_2 = input("Введите имя второго файла без расширения:\n")
    file_2 = os.path.abspath(file_2 + '.xlsx')
    comparison(file_1, file_2)


