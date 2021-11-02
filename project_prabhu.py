import nsepy
import pandas as pd
import yfinance as yf
from datetime import date
from openpyxl import Workbook


def shorter_the_name(company_name):
    if len(company_name)>=31:
        company_name = company_name[:30]
    else:
        company_name = company_name
    return company_name


df = pd.read_csv("https://www1.nseindia.com/content/indices/ind_nifty50list.csv")
sheet_names = df["Company Name"].tolist()
sheet_names = list(map(shorter_the_name,sheet_names))
tickerSymbol = df.Symbol.tolist()
# print(tickerSymbol)
temp_dic = dict()
temp_dic["Company Name"]=[]
temp_dic["symbol"] = []
temp_dic["Avgprice"] = []
temp_dic["mean"] = []
temp_dic["maxprice"] = []
temp_dic["minprice"] = []
temp_dic["growth"] =[]

writer = pd.ExcelWriter("project_prabhu.xlsx", engine='xlsxwriter')
for i in range(len(df.Symbol)):
    tickerDf = nsepy.get_history(symbol=tickerSymbol[i], start=date(2021, 1, 1), end=date(2021, 8, 31))
  #  print(tickerDf.columns)
    temp_dic["Company Name"].append(sheet_names[i])
    temp_dic["symbol"].append(tickerSymbol[i])
    temp_dic["Avgprice"].append(tickerDf["VWAP"].mean())
    temp_dic["maxprice"].append(max(tickerDf["Close"]))
    temp_dic["mean"].append(tickerDf["Close"].mean())
    temp_dic["minprice"].append(min(tickerDf["Close"]))
    growt= (((tickerDf.tail(1)["VWAP"].values[0])-(tickerDf.head(1)["VWAP"].values[0]))/(tickerDf.head(1)["VWAP"].values[0]))*100
    temp_dic["growth"].append(growt)
    #print(tickerDf)
    tickerDf.reset_index(inplace=True)
    #print(tickerDf.head())
    #print(tickerDf.columns)
    #tickerDf.to_excel(writer, index=False, sheet_name=sheet_names[i])

CompanyDf = pd.DataFrame(data=temp_dic)
CompanyDf.to_excel(writer,index=False,sheet_name="nifty50")



writer.save()
writer.close()