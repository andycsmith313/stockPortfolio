#!/usr/bin/python
# -*- coding: utf-8 -*-

import requests, json, pprint, datetime, openpyxl, os, csv
import matplotlib.pyplot as plt
import pandas as pd
from PyQt4 import QtGui

def historicalDataGET(stockTicker,startDate,endDate):
	'''
	This function returns a JSON object for the Yahoo YQL API,
	the user inputs the stock they want data for the start date of the data an the end date
	'''
	url = r'https://query.yahooapis.com/v1/public/yql?q=select%20*%20from%20yahoo.finance.historicaldata%20where%20symbol%20%3D%20"' + stockTicker + r'"%20and%20startDate%20%3D%20"' + startDate + r'"%20and%20endDate%20%3D%20"' + endDate + r'"&format=json&diagnostics=true&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys&callback='
	response = requests.get(url)
	response.raise_for_status()
	historicalData = json.loads(response.text)
	keyStockData = historicalData['query']['results']['quote']
	return keyStockData

def datePriceDF():
	'''
	Added more detail
	'''
	for stock in stockTicker:
		historicalData = historicalDataGET(stock,startDate,endDate)
		datePrice = {}
		for i, key in enumerate(historicalData):
			if 'Date' in historicalData[i].keys():
				datePrice[historicalData[i]['Date']] = (float(historicalData[i]['Close'])/100)
		stockTicker[stock] = datePrice
	datePriceDF = pd.DataFrame(stockTicker)
	datePriceDF.index = pd.to_datetime(datePriceDF.index)
	return datePriceDF

def portfolioValueDF(dataFrame):
	'''
	asd
	'''
	for stock in stockTicker:
		dataFrame[stock] = stockQuantity[stock] * dataFrame[stock]
	dataFrame['totalValue'] = dataFrame[list(dataFrame)].sum(axis=1)
	return dataFrame
	
os.chdir(r"C:\Users\Andrew\Google Drive\Python\FilesToRead") # Set current directory
wb = openpyxl.load_workbook('StockPortfolio.xlsx', data_only=True) # Open Stock Portfolio Workbook
# Table layout: | Stock_Name | Ticker | Quantity | Avg_Price | Book_Cost |
Portfolio = wb.get_sheet_by_name('Portfolio')

stockTicker = {} #Create empty dictionary
stockBookCost = {} #Create empty dictionary
stockAvgPrice = {} #Create empty dictionary
stockQuantity = {} #Create empty dictionary
for r in range(2, Portfolio.max_row + 1): #Start loop at row 2 and finish at the last row with data in
	stockTicker[Portfolio.cell(row=r, column=2).value] = None
	stockQuantity[Portfolio.cell(row=r, column=2).value] = Portfolio.cell(row=r, column=3).value
	stockAvgPrice[Portfolio.cell(row=r, column=2).value] = Portfolio.cell(row=r, column=4).value
	stockBookCost[Portfolio.cell(row=r, column=2).value] = Portfolio.cell(row=r, column=5).value

startDate = '2016-04-01'
endDate = '2016-09-23'

df = datePriceDF()

print(portfolioValueDF(df))
#### Plot Graph
plt.plot(df['totalValue'])
plt.show()