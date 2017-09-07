import pandas as pd

import numpy as np

import  xlwings as xw


from __future__ import division

#open the excel and copy it into Python. Takes a minute please wait
SCF_book = xw.Book(r'C:\Economics PhD\Third Year\Advanced Macro\Homework 1\SCFP2007.xlsx')

SCF = SCF_book.sheets[0].range('A1').expand().value


#Turn the disorganized input into a Pandas Dataframe

labels = SCF[0]

del SCF[0]   #Remove the list now since we don't want it to appear again in the dataframe besides being a label



SCF_1 = pd.DataFrame.from_records(SCF, columns=labels)

print(SCF_1['INCOME'])


#Creating the variables


inf = 1.1235  #Inflation rate to be adjusted
#Earnings
SCF_1['EARNINGS'] = (SCF_1['WAGEINC']+ (.863)*SCF_1['BUSSEFARMINC'])/inf


#Income

SCF_1['INCOME1']=(SCF_1['WAGEINC']+SCF_1['TRANSFOTHINC']+SCF_1['SSRETINC']+SCF_1['KGINC']+SCF_1['INTDIVINC']+SCF_1['BUSSEFARMINC'])/inf


#Wealth

SCF_1['NETWORTH'] =  SCF_1['NETWORTH']/inf


#Create the quantiles

q = [0, .01, .05,.1,.2,.4,.6,.8,.9,.95,.99, 1.]


#Earnings Quantiles

EAR_qunt = []

for i in q:
    x = (SCF_1['EARNINGS']/(1000)).quantile(i)
    EAR_qunt.append(x)

print(EAR_qunt)

#Wealth

WEALTH_qunt = []

for i in q:
    x = (SCF_1['NETWORTH']/(1000)).quantile(i)
    WEALTH_qunt.append(x)

print(WEALTH_qunt)

#Income Quantiles

Inc_qunt = []

for i in q:
    x = SCF_1['INCOME1'].quantile(i)
    Inc_qunt.append(x)

print(Inc_qunt)



#Coefficient of Variation = standard Deviation/Mean

ear_cof = np.std(SCF_1['EARNINGS'])/np.mean(SCF_1['EARNINGS'])

inc_cof = np.std(SCF_1['INCOME1'])/np.mean(SCF_1['INCOME1'])

wealth_cof = np.std(SCF_1['NETWORTH'])/np.mean(SCF_1['NETWORTH'])

cof_var = [ear_cof,inc_cof,wealth_cof]


#Variance of the logs




ear_logvar = np.nanvar(np.log(SCF_1['EARNINGS'][SCF_1['EARNINGS']>0]))

inc_logvar = np.nanvar(np.log(SCF_1['INCOME1'][SCF_1['INCOME1']>0]))

wealth_logvar = np.nanvar(np.log(SCF_1['NETWORTH'][SCF_1['NETWORTH']>0]))


log_var = [ear_logvar,inc_logvar,wealth_logvar]



#Gini Index

####################Borrowed Code#################
def gini(list_of_values):
    sorted_list = sorted(list_of_values)
    height, area = 0, 0
    for value in sorted_list:
        height += value
        area += height - value / 2.
    fair_area = height * len(list_of_values) / 2.
    return (fair_area - area) / fair_area
########################


GINI = [gini(SCF_1['EARNINGS']),gini(SCF_1['INCOME1']) , gini(SCF_1['NETWORTH'])]


#Top 1% over lowest 40%

ear_to = 'Place Holder'

inc_to = 'Place Holder'

wealth_to =  'Place Holder'


topoverlowest = [ear_to,inc_to,wealth_to]

#Location of Mean %

ear_locm = len(SCF_1['EARNINGS'][SCF_1['EARNINGS']<np.mean(SCF_1['EARNINGS'])])/len(SCF_1['EARNINGS'])

inc_locm = len(SCF_1['INCOME1'][SCF_1['INCOME1']<np.mean(SCF_1['INCOME1'])])/len(SCF_1['INCOME1'])

wealth_locm = len(SCF_1['NETWORTH'][SCF_1['NETWORTH']<np.mean(SCF_1['NETWORTH'])])/len(SCF_1['NETWORTH'])

loc_mean = [ear_locm,inc_locm, wealth_locm]

#Mean/Median

ear_MM = np.mean(SCF_1['EARNINGS'])/np.median(SCF_1['EARNINGS'])

Inc_MM = np.mean(SCF_1['INCOME1'])/np.median(SCF_1['INCOME1'])

Wealth_MM = np.mean(SCF_1['NETWORTH'])/np.median(SCF_1['NETWORTH'])

meanovermedian = [ear_MM,Inc_MM, Wealth_MM]




#Table 2

Table2 = {'Cof. of Variation' : cof_var , 'Variance of Logs' : log_var , 'Gini Index' : GINI , 'Top 1% over 40%': topoverlowest , 'location of mean (%)': loc_mean , 'Mean/Median': meanovermedian}

Table2 = pd.DataFrame(Table2, index = ['Earnings','Income','Wealth'])

SCF_book.sheets.add(name = 'Table2')

SCF_book.sheets('Table2').range('A1').value = Table2.transpose()



#Table 1 as Pandas Dataframe

Table1 = {'Earnings' : EAR_qunt, 'Income' : Inc_qunt, 'Wealth' : WEALTH_qunt}

Table1 = pd.DataFrame(Table1, index = q)

SCF_book.sheets.add(name = 'Table1')

SCF_book.sheets('Table1').range('A1').value = Table1.transpose()



