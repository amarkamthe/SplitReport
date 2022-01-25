import pandas as pd
from openpyxl import load_workbook

def write_to_excel(df, name):
    book = load_workbook('stock.xlsx')
    with pd.ExcelWriter('stock.xlsx') as writer:
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        df.to_excel(writer, sheet_name=name)
        writer.save()

xls = pd.ExcelFile('stock.xlsx')
xls_split = pd.ExcelFile('stock_split.xlsx')
df1 = pd.read_excel(xls, 'Stock')
df2 = pd.read_excel(xls_split, 'Split')


#Sym, Quanitiy Available, Avg price, Invested Amt, Current Amt, Abs return, % of Invested, % of current, % of change
df1['Invested Amt'] = df1['Quantity Available'] * df1['Average Price']
df1['Current Amt'] = df1['Quantity Available'] * df1['Previous Closing Price']
df1['Abs return'] = round(((df1['Current Amt'] - df1['Invested Amt']) / df1['Invested Amt'])*100, 2)
#df1['Perentage allocation'] =

total_invested_price = df1['Invested Amt'].sum()
total_current_price = df1['Current Amt'].sum()

stock = {};
c_name =''
for (column_name, column_data) in df2.iteritems():
    if c_name != "" and c_name.lower() + ' percentage' == column_name.lower():
        stock[c_name] = dict(zip(stock[c_name], column_data.dropna().to_list()))
    else:
        stock[column_name] = column_data.dropna().to_list()
        c_name = column_name

data=[]
for (group, values) in stock.items():
    dfs = df1[df1['Symbol'].isin(values.keys())].copy()
    for key,val in values.items():
        if val < 100:
            d = dfs[ dfs['Symbol'] == key]
            d['Quantity Available'] = d['Quantity Available'] * (val/100)
            dfs[dfs['Symbol'] == key] = d

    dfs['Invested Amt'] = dfs['Quantity Available'] * dfs['Average Price']
    dfs['Current Amt'] = dfs['Quantity Available'] * dfs['Previous Closing Price']
    dfs['Abs return'] = round(((dfs['Current Amt'] - dfs['Invested Amt']) / dfs['Invested Amt']) * 100, 2)
    total_invested_price_dfs = dfs['Invested Amt'].sum()
    total_current_price_dfs = dfs['Current Amt'].sum()
    dfs['% of invested'] = round((dfs['Invested Amt']/total_invested_price_dfs)*100, 2)
    dfs['% of current'] = round((dfs['Current Amt']/total_current_price_dfs)*100, 2)
    data.append([group, total_invested_price_dfs, total_current_price_dfs])
    df = dfs[['Symbol', 'Quantity Available', 'Average Price', 'Invested Amt', 'Current Amt', 'Abs return', '% of invested', '% of current' ]]
    write_to_excel(df, group)

df_tmp = pd.DataFrame(data, columns = ['Name', 'Invested Amt', 'Current Amt'])

df_tmp['% of invested Amt'] = round((df_tmp['Invested Amt']/total_invested_price)*100, 2)
df_tmp['% of current Amt'] = round((df_tmp['Current Amt']/total_current_price)*100, 2)
write_to_excel(df_tmp, 'Output')

