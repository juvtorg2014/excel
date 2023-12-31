import os
import pandas as pd


def find_row(df) -> int:
	""" Поиск строки с названиями нужных колонок"""
	number = 0
	while True:
		if df.iloc[number][0] == 'Серия номенклатуры.Поставщик':
			break
		else:
			number += 1
	return number


def find_nun(df) -> list:
	""" Поиск пустых колонок для удаления"""
	del_list = []
	for item in range(len(df.columns) - 1):
		if type(df.columns[item]) is not str:
			del_list.append(item)
	return del_list


def comparison(file_source, file_receiver):
	"""Сравение двух Еxcel-файлов на разницу в суммах"""
	source = pd.read_excel(file_source)
	row = find_row(source)
	del_list = [x for x in range(row)]
	source = source.drop(index=del_list)
	source.columns = source.iloc[0]
	source = source[1:]
	
	num_col = find_nun(source)
	source.drop(source.columns[num_col], axis=1, inplace=True)
	source = source[:-1]
	source = source.rename(columns={'Серия номенклатуры.Поставщик': 'Поставщик',
									'Сумма комиссионного вознаграждения ': 'Комиссия'})
	new_source = source[['Поставщик', 'Количество', 'Выручка', 'Комиссия']].copy(deep=False)
	new_source = new_source['Поставщик'].str.strip()
	new_source = new_source.sort_values(by=['Поставщик'])
	new_source.to_csv(file_source.replace('.xlsx', '.csv'), index=False, header=True, encoding='cp1251', sep=';')
	
	receiver = pd.read_excel(file_receiver)
	receiver = receiver.rename(columns={'Контрагент': 'Поставщик', 'Количество документа': 'Количество',
										'Сумма документа': 'Выручка', 'Сумма вознаграждения ': 'Комиссия'})
	receiver = receiver.sort_values(by=['Поставщик'])
	new_reciever = receiver[['Поставщик', 'Количество', 'Выручка', 'Комиссия']].copy(deep=False)
	new_reciever = new_reciever['Поставщик'].str.strip()
	new_reciever = new_reciever.sort_values(by=['Поставщик'])
	new_reciever.to_csv(file_receiver.replace('.xlsx', '.csv'), index=False, header=True, encoding='cp1251', sep=';')
	
	dif_col = new_source.set_index('Поставщик').subtract(new_reciever.set_index('Поставщик'))
	dif_table = dif_col[(dif_col['Количество'] != float(0)) | (dif_col['Выручка'] != float(0))
						| (dif_col['Комиссия'] != float(0))]
	dif_table.to_csv(os.getcwd() + '\\' + 'Разница.csv', index=True, header=True, encoding='cp1251', sep=';',
					                                                                  float_format='{15.2f}')
	print(dif_table)


if __name__ == '__main__':
	file_1 = input("Введите имя файла отчета без расширения:\n")
	file_1 = os.path.abspath(file_1 + '.xlsx')
	file_2 = input("Введите имя файла документов без расширения:\n")
	file_2 = os.path.abspath(file_2 + '.xlsx')
	comparison(file_1, file_2)
