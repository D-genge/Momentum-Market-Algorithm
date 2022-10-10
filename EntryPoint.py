import numpy as np
import pandas as pd
import requests
import xlsxwriter
import math
import time
import datetime
import matplotlib.pyplot as plt
from urllib.error import HTTPError

sym = pd.read_csv("C:/Users/dunca/OneDrive/Desktop/MarketAlgorithm/symbols.csv")
myColumns = ['Ticker', 'Slope', 'Error', 'Avg Close', 'Close']
dfEntryArray = [0, 0, 0]
finalDF = pd.DataFrame(columns = myColumns)

from sklearn.model_selection import train_test_split
from sklearn import linear_model
from sklearn.metrics import mean_squared_error, r2_score
from sklearn import metrics

def LRegression(dfIn):
    X = dfIn[['Indexes']]
    Y = dfIn[['Close']]
    X_train, X_test, Y_train, Y_test = train_test_split(X, Y, test_size = 0.2)

    model = linear_model.LinearRegression()
    model.fit(X_train, Y_train)

    predictions = model.predict(X_test)
    coeffecients = model.coef_
    RMSE = metrics.r2_score(Y_test, predictions)
    avgClose = pd.Series.mean(dfIn['Close'])
    dfEntryArray[0] = model.coef_
    dfEntryArray[1] = RMSE
    dfEntryArray[2] = avgClose
    return
         


def riskAnalysis(slope, error, avgClose):
    if(error > 0.65 or error < -0.65): # Checks for steady growth based on error
        print("Error Good")
        # if(slope/avgClose > 0.015 or slope/avgClose < -0.015): # Checks regression has some movement and price is not stagnant
        return True
    return False

for i in sym['symbols'][:]:
    try:
        ticker = f'{i}'
        P1 = int(time.mktime(datetime.datetime(2017, 9, 25, 23, 59).timetuple()))
        P2 = int(time.mktime(datetime.datetime(2022, 9, 25, 23, 59).timetuple()))
        interval = '1d'
        queryString = f'https://query1.finance.yahoo.com/v7/finance/download/{ticker}?period1={P1}&period2={P2}&interval={interval}&events=history&includeAdjustedClose=true'

        InDf = pd.read_csv(queryString) # Reads the csv file returned from Yahoo
        df = InDf.drop(labels = ['Volume', 'High', 'Open', 'Low', 'Adj Close'], axis = 1)
        df['Indexes'] = range(1, len(df)+1)
        LRegression(df.tail(115)) # Does Linear Regression of last 6 months (115 business days)
        indicatorST =  riskAnalysis(dfEntryArray[0], dfEntryArray[1], dfEntryArray[2])
    except HTTPError:
        continue
    except TimeoutError:
        break
    except ConnectionResetError:
        break
    except ValueError:
        continue

    if(indicatorST == True):
        finalDF = finalDF.append(pd.Series([ticker, float(dfEntryArray[0]), dfEntryArray[1], dfEntryArray[2], df['Close'].iat[-1]], 
                                 index = myColumns), ignore_index = True)
    time.sleep(2.5) # Yahoo Finance has a 2000 requests per hour limit. This keeps number of requests below the limit

finalDF = finalDF.sort_values(by=['Error'], ascending = False)
print(finalDF)

Bank = 1000000 # Amount of money available
MoneyIntervals = Bank/len(finalDF)
workbook = xlsxwriter.Workbook("Stock-Info.xlsx")
worksheet = workbook.add_worksheet("Purchase-Info")

worksheet.write(0, 0, "#")
worksheet.write(0, 1, "Ticker")
worksheet.write(0, 2, "Long/Short")
worksheet.write(0, 3, "Price at Purchse (Close)")
worksheet.write(0, 4, "Number of Stocks Purchased")


for i, j in finalDF.iterrows(): # Gives equal weight to each stock 
    NumOfStocks = math.floor(MoneyIntervals/j['Close'])
    worksheet.write(i+1, 0, i)
    worksheet.write(i+1, 1, j['Ticker'])
    if j['Slope'] > 0:
        worksheet.write(i+1, 2, "Long")
    else:
        worksheet.write(i+1, 2, "Short")
        
    worksheet.write(i+1, 3, j['Close'])
    worksheet.write(i+1, 4, NumOfStocks)
    
workbook.close()

