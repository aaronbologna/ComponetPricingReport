# -*- coding: utf-8 -*-
"""
Created on Tue Apr 14 10:05:12 2020

@author: Aaron
"""


import numpy as np
import pandas as pd



wu = pd.read_excel('All BOM WU with Inv.xlsx') #where used information from Access DB, need to import list of PNs to Access to export WU
wu = wu[['PN','Description 1','Parent','SumOfQty']]
wu.columns = ['PN','Description','65M','QtyPer']
mrp = pd.read_excel('65_MRP.xlsx') #MRP controller of all 65Ms from SAP ZMAT_STATUS
crestron_price = pd.read_excel('Crestron_Price.xlsx') #ZCOSTPRICEINFO for all PNs 2M/4M/6M, includes planned price 1 as well
crestron_price = crestron_price[['Material','Standard Price','PlndPrice1']]
crestron_price['Standard Price'] = crestron_price['Standard Price'].apply(lambda x: x/1000)
crestron_price['PlndPrice1'] = crestron_price['PlndPrice1'].apply(lambda x: x/1000)
crestron_price.columns = ['PN','Crestron Price', 'Crestron Planned Price']
jabil_price = pd.read_excel('Jabil_Price.xlsx',sheet_name='BOM Detail') #entire reprice file
jabil_price = jabil_price[['Component','New price']]
jabil_price['Component'] = jabil_price['Component'].apply(lambda x: x[2:9])
jabil_price['Component'] = jabil_price['Component'].apply(lambda x: int(x) if str(x).isdigit() else None)
jabil_price.columns = ['PN','Jabil Price']
jabil_price = jabil_price.drop_duplicates(subset='PN')
neo_price = pd.read_excel('Neo_Price.xlsx', sheet_name= 'CBOM') #cost of all 2M from Neo reprice, NEED TO CHANGE TO XLXS
neo_price = neo_price[['CustomerPartNum','StdCost']]
neo_price.columns = ['PN','Neo Price']
neo_price = neo_price.drop_duplicates(subset='PN')
ma = pd.read_excel('MA.xlsx') #3 month average from monthly email report

mrp_join = pd.merge(wu,mrp, on ='65M', how = 'left')
crestron_join = pd.merge(mrp_join, crestron_price, on ='PN', how = 'left')
jabil_join = crestron_join.merge(jabil_price, on ='PN', how = 'left')
neo_join = pd.merge(jabil_join, neo_price, on ='PN', how = 'left')
df = pd.merge(neo_join, ma, on ='65M', how = 'left')
df = df.dropna(subset=['MA'])

pd.set_option('max_columns',None)
pd.set_option('max_rows',None)




#new columns to calculate the annual usage based on the MRP of the top level
df['Jabil Usage'] = df.apply(lambda row: row['QtyPer'] * 12 * row['MA'] if row['MRP'] in (33, 41) else 0, axis = 1)
df['Neo Usage'] = df.apply(lambda row: row['QtyPer'] * 12 * row['MA'] if row['MRP'] == 23 else 0, axis = 1)
df['Crestron Usage'] = df.apply(lambda row: row['QtyPer'] * 12 * row['MA'] if row['MRP'] in (12, 13, 14, 15, 16, 17, 18, 3, 7) else 0, axis = 1)

#new columns to calculate the potential savings if the demand in one plant was purchased at the price of another
#ensures that no savings are generated based off of switching to a price of zero
df['CtoJ_Savings'] = df.apply(lambda row: (row['Crestron Price'] - row['Jabil Price']) * row['Crestron Usage'] if row['Jabil Price'] >0 else 0, axis = 1)
df['CtoN_Savings'] = df.apply(lambda row: (row['Crestron Price'] - row['Neo Price']) * row['Crestron Usage'] if row['Neo Price'] >0  else 0, axis = 1)
df['NtoJ_Savings'] = df.apply(lambda row: (row['Neo Price'] - row['Jabil Price']) * row['Neo Usage'] if row['Jabil Price'] >0 else 0, axis = 1)
df['JtoN_Savings'] = df.apply(lambda row: (row['Jabil Price'] - row['Neo Price']) * row['Jabil Usage'] if row['Neo Price'] >0  else 0, axis = 1)

#removing negative savings
df['CtoJ_Savings'] = df['CtoJ_Savings'].apply(lambda x: x if x>0 else 0)
df['CtoN_Savings'] = df['CtoN_Savings'].apply(lambda x: x if x>0 else 0)
df['NtoJ_Savings'] = df['NtoJ_Savings'].apply(lambda x: x if x>0 else 0)
df['JtoN_Savings'] = df['JtoN_Savings'].apply(lambda x: x if x>0 else 0)

#new dfs to summarize the potential cost savings per PN per plant
CtoJ = df.groupby('PN')['CtoJ_Savings'].sum().sort_values(ascending=False)
CtoN = df.groupby('PN')['CtoN_Savings'].sum().sort_values(ascending=False)
NtoJ = df.groupby('PN')['NtoJ_Savings'].sum().sort_values(ascending=False)
JtoN = df.groupby('PN')['JtoN_Savings'].sum().sort_values(ascending=False)

#new df to summarize the usage per PN per plant
C_Usage = df.groupby('PN')['Crestron Usage'].sum()
J_Usage = df.groupby('PN')['Jabil Usage'].sum()
N_Usage = df.groupby('PN')['Neo Usage'].sum()
usage = pd.merge(C_Usage, J_Usage, on= 'PN', how = 'left')
usage = pd.merge(usage, N_Usage, on = 'PN', how = 'left')

#compiling the summary PN information to add to savings and usage
desc = df[['PN','Description', 'Crestron Price','Crestron Planned Price', 'Jabil Price', 'Neo Price']].drop_duplicates(subset='PN')
desc.reset_index(drop=True,inplace=True)
#add usage to summary
desc = pd.merge(desc, usage, on = 'PN', how = 'left')

#compile potential savings into one df and add in summary information per pn
Potential_Savings = pd.merge(CtoJ, CtoN, on = 'PN', how = 'left')
Potential_Savings = pd.merge(Potential_Savings, NtoJ, on = 'PN', how = 'left')
Potential_Savings = pd.merge(Potential_Savings, JtoN, on = 'PN', how = 'left')
Potential_Savings = pd.merge(Potential_Savings, desc, on = 'PN', how = 'left')

Potential_Savings.to_excel('Comp_Price_Report.xlsx',sheet_name='Summary_Component')

#65M evaluation
CtoJ65 = df.groupby('65M')['CtoJ_Savings'].sum().sort_values(ascending=False)
CtoN65 = df.groupby('65M')['CtoN_Savings'].sum().sort_values(ascending=False)
NtoJ65 = df.groupby('65M')['NtoJ_Savings'].sum().sort_values(ascending=False)
JtoN65 = df.groupby('65M')['JtoN_Savings'].sum().sort_values(ascending=False)
print()

Summary_65M = pd.merge(CtoJ65, CtoN65, on = '65M', how = 'left')
Summary_65M = pd.merge(Summary_65M, NtoJ65, on = '65M', how = 'left')
Summary_65M = pd.merge(Summary_65M, JtoN65, on = '65M', how = 'left')
Summary_65M['Total Potential'] = Summary_65M.apply(lambda row: (row['CtoJ_Savings'] + row['CtoN_Savings']+ row['NtoJ_Savings']+ row['JtoN_Savings']) , axis = 1)
Summary_65M = pd.merge(Summary_65M, mrp, on = '65M', how = 'left')

from openpyxl import load_workbook
book = load_workbook('Comp_Price_Report.xlsx')
writer = pd.ExcelWriter('Comp_Price_Report.xlsx',engine='openpyxl')
writer.book = book

#new sheets to summarize savings for the top 65Ms
#finding the top 65Ms
Summary_65M.sort_values('Total Potential',inplace=True,ascending=False)
Summary_65M.reset_index(drop=True,inplace=True)
Top_65M = Summary_65M.iloc[0:10]
list_65M = Top_65M['65M'].to_list()

Summary_65M.to_excel(writer,sheet_name='Summary_FG')

#creating summary sheets for each top 65M, returns all the rows in the main df with the 65M and saving them in a new sheet in the summary file
for i in range(len(list_65M)):
    mask = df['65M'] == list_65M[i]
    savings = df[mask]
    savings.to_excel(writer,sheet_name=str(list_65M[i]))
writer.save()



