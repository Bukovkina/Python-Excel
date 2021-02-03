import requests
import datetime as dt

import apimoex
import pandas as pd


def get_data(currencies=[]):
	# set arguments for request
	date_till = dt.datetime.now(dt.timezone(dt.timedelta(hours=3))).date()
	date_from = date_till - dt.timedelta(days=30)

	request_urls = [('https://iss.moex.com/iss/statistics/engines/futures/'
				     'markets/indicativerates/securities/{}.json'
				     '?till={}&from={}'.format(currency, str(date_till), str(date_from))) 
										   for currency in currencies]
	dfs = [] # list of currency dataframes 
	with requests.Session() as session:
		for req in request_urls:
			iss = apimoex.ISSClient(session, req)
			data = iss.get()
			dfs.append(pd.DataFrame(data['securities']))
	return dfs

	
def create_excel(dfs=[]):
	output_df = pd.DataFrame()
	for df in dfs:
		cur_name = df.secid[0]
		output_df[f'Дата {cur_name}'] = df.tradedate
		output_df[f'Курс {cur_name}'] = df.rate
		output_df[f'Изменение {cur_name}'] = output_df[f'Курс {cur_name}'].diff(periods=-1)
	output_df['Отношение EUR к USD'] = output_df['Курс EUR/RUB'] / output_df['Курс USD/RUB']
	
	file_name = f'Currencies_{str(date_from)}_{str(date_till)}.xlsx'
	output_df.to_excel(file_name, index=False, engine='openpyxl') 

	
if __name__ == '__main__':
	currencies = ['USD/RUB', 'EUR/RUB']
	data = get_data(currencies)
	create_excel(data)
