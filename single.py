import os
import pandas as pd
import pip



def comparison(file_source, file_receiver):
    pip.main(['install', 'openpyxl'])
    #pd.options.display.float_format = '{.2f}'.format
    source = pd.read_excel(file_source)
    source = source.drop(index=[0, 1, 2, 3, 4])
    source = source.rename(columns={'Unnamed: 4': 'Seria', 'Unnamed: 5': 'Quant',
                                    'Unnamed: 8': 'Sum', 'Unnamed: 9': 'Commiss'})
    new_source = source[['Seria', 'Quant', 'Sum', 'Commiss']].copy(deep=False)
    new_source = new_source[:-1]
    # new_source['Seria'].astype(int)
    # new_source['Seria'].astype(int)
    # new_source['Seria'].astype(float)
    # new_source['Seria'].astype(float)
    #new_source['Seria'] = new_source['Seria'].str.strip()
    new_source = new_source.sort_values(by=['Seria'])

    receiver = pd.read_excel(file_receiver)
    col_1 = receiver.columns[5]
    col_2 = receiver.columns[8]
    col_3 = receiver.columns[14]
    col_4 = receiver.columns[15]
    #col_5 = receiver.columns[19]
    receiver = receiver.rename(columns={'Серия номенклатуры': 'Seria', 'Количество': 'Quant',
                                        'Сумма (RUB)': 'Sum', 'Вознаграждение(RUB)': 'Commiss'})
    new_reciever = receiver[['Seria', 'Quant', 'Sum', 'Commiss']].copy(deep=False)
    new_reciever = new_reciever['Seria'].astype(str)
    #new_reciever['Seria'] = new_reciever['Seria'].str.strip()
    new_reciever = new_reciever.sort_values(by=['Seria'])
    #new_reciever = new_reciever.round({'Sum': 2, 'Commiss': 2})

    # new_source = new_source.reindex(range(len_dataframe), method='ffill')
    new_source.to_csv(file_source.replace('.xlsx', '.csv'), index=False, header=True, encoding='cp1251', sep=';')
    new_reciever.to_csv(file_receiver.replace('.xlsx', '.csv'), index=False, header=True, encoding='cp1251', sep=';')

    dif_col = new_source.set_index('Seria').subtract(new_reciever.set_index('Seria'))
    dif_table = dif_col[(dif_col['Sum'] != float(0)) | (dif_col['Commiss'] != float(0))]

    dif_table.to_csv('Товары.csv', index=True, header=True, encoding='cp1251', sep=';')
    print(dif_table)




if __name__ == '__main__':
    #file_1 = input("Введите имя файла контрагента без расширения\n")
    file_1 = 'Кавелина-отч'
    file_1 = os.path.abspath(file_1 + '.xlsx')
    #file_2 = input("Введите имя второго файла без расширения\n")
    file_2 = 'Кавелина-док'
    file_2 = os.path.abspath(file_2 + '.xlsx')
    comparison(file_1, file_2)


