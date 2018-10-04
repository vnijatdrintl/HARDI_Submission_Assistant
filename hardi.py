import pandas as pd
from math import ceil
import sys
import datetime
import os

def createDF(df,orgName):

    #get number of rows
    row=df.shape[0]
    #create result data frame
    ans=pd.DataFrame(pd.np.empty((row, 13)) * pd.np.nan,columns=['Organization','Branch_Num','To_Zip','From_Zip','Model_Num','Product_Type','Invoice_Num','Total_Unit_Sales','Sales_Price','Year','Quarter','Month','Day'])

    #write organization name to the data frame
    orgNameCol=[orgName]*row
    orgNameCol=pd.Series(orgNameCol)
    ans['Organization']=orgNameCol.values

    return row,ans


def cleanCDJ(df):

    if df.shape[1]<8:
        return None

    df.columns=['Product','Description','Recs','Quantity','Net Sale','Disc ','Margin','Margin']

    row=df.shape[0]
    branchCol=[pd.np.nan]*row
    branchCol=pd.Series(branchCol)
    df['Branch']=branchCol.values
    df=df.fillna('')

    branch=''
    for index,row in df.iterrows():
        if 'Warehouse' in str(row['Product']):
            branch=str(row['Product'])[-4:].upper()
        row['Branch']=branch

    df=df[df['Branch']!='']
    df=df[df['Recs']!='']
    df=df[df['Product']!='']
    df=df[df['Product']!='Product']
    df=df[df['Description']!='']
    df=df.reset_index()
    df=df.drop(['index'],axis=1)

    return df


def cleanNine(df,missCol):
    for i in range(0,df.shape[1]):
        if df[df.columns[i]].isnull().sum()>0.95*df.shape[0]:
            missCol=i
            break
    if missCol==1:
        df[df.columns[0]]=df[df.columns[0]].map(str)+' '+df[df.columns[1]].map(str)
    elif missCol==2:
        df[df.columns[1]]=df[df.columns[1]].map(str)+' '+df[df.columns[2]].map(str)
    elif missCol==4:
        for index,row in df.iterrows():
            if row.isnull().sum()==0:
                temp=str(df.iat[index,1])+' '+str(df.iat[index,2])
                df.iat[index,1]=temp
                recs=df.iat[index,3]
                df.iat[index,2]=recs
                qty=df.iat[index,4]
                df.iat[index,3]=qty
    df=df.drop(df.columns[missCol],axis=1)

    return df

def twoJsupply(file_path,orgName,year,month):

    #load the file to a data frame
    df=pd.read_csv(file_path,header=None)
    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df[0]
    ans['Total_Unit_Sales']=df[2]
    ans['Branch_Num']=df[3]
    ans['Invoice_Num']=df[4]

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df[1]=pd.to_datetime(df[1])
    ans['Year']=df[1].dt.year
    ans['Month']=df[1].dt.month
    ans['Day']=df[1].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def ABRWholeSalers(file_path,orgName,year,month):

    #load the file to a data frame
    xls=pd.ExcelFile(file_path)

    if len(xls.sheet_names)>1:
        df=pd.read_excel(xls,"Sales Fields",skiprows=4)
    else:
        df=pd.read_excel(xls,skiprows=4)

    df=df.fillna('')
    df=df[df['Model Number']!='']
    df=df.reset_index()
    df=df.drop(['index'],axis=1)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['Model Number']
    ans['Total_Unit_Sales']=df['Number of Units']
    ans['Branch_Num']=df['Location Number']
    ans['Invoice_Num']=df['Invoice Number']
    ans['To_Zip']=df['Delivery ZIP Code']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Sale Date']=pd.to_datetime(df['Sale Date'])
    ans['Year']=df['Sale Date'].dt.year
    ans['Month']=df['Sale Date'].dt.month
    ans['Day']=df['Sale Date'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def acpro(file_path,orgName,year,month):

	xls=pd.ExcelFile(file_path)
	df=pd.read_excel(xls)

	row,ans=createDF(df,orgName)

	#Copy other columns
	ans['Model_Num']=df.iloc[:,3]
	ans['Total_Unit_Sales']=pd.to_numeric(df.NumberOfUnits)
	ans['Branch_Num']=df.iloc[:,8]
	ans['Sales_Price']=pd.to_numeric(df.SalePrice)
	ans['To_Zip']=df.iloc[:,9]
	ans['Invoice_Num']=df.iloc[:,0]

	priceCol=[pd.np.nan]*row
	priceCol=pd.Series(priceCol)

	for i in range(0,row):
		if ans['Total_Unit_Sales'].iloc[i]==0:
			priceCol.iloc[i]=0
		else:
			priceCol.iloc[i]=ans['Sales_Price'].iloc[i]/ans['Total_Unit_Sales'].iloc[i]

	ans['Sales_Price']=priceCol.values

	#quarter
	quarter=ceil(month/3)
	quarterCol=[quarter]*row
	quarterCol=pd.Series(quarterCol)
	ans['Quarter']=quarterCol.values

	df['Date']=pd.to_datetime(df.Date)
	ans['Year']=df['Date'].dt.year
	ans['Month']=df['Date'].dt.month
	ans['Day']=df['Date'].dt.day

	ans=ans.fillna('')

	return ans,row


def acrsupply(file_path,orgName,year,month):

    xls = pd.ExcelFile(file_path)
    df = pd.read_excel(xls)
    row,ans=createDF(df,orgName)

    for i in range(0,row):
        if pd.isna(df['Item Num'].iloc[i]):
            df['Item Num'].iloc[i]=df['Item Num'].iloc[i-1]

    for i in range(0,row):
        if pd.isna(df['Branch'].iloc[i]):
            df['Branch'].iloc[i]=df['Branch'].iloc[i-1]

    row,ans=createDF(df,orgName)

    # Copy other columns
    ans['Model_Num'] = df.iloc[:, 2]
    ans['Branch_Num'] = df.iloc[:, 3]
    ans['Total_Unit_Sales'] = df.iloc[:, 5]
    ans['Sales_Price'] = df.iloc[:, 6]

    priceCol = [pd.np.nan] * row
    priceCol = pd.Series(priceCol)

    for i in range(0, row):
        if ans['Total_Unit_Sales'].iloc[i] == 0:
            priceCol.iloc[i] = 0
        else:
            priceCol.iloc[i] = ans['Sales_Price'].iloc[i] / ans['Total_Unit_Sales'].iloc[i]

    ans['Sales_Price'] = priceCol.values

    # quarter
    quarter = ceil(month / 3)
    quarterCol = [quarter] * row
    quarterCol = pd.Series(quarterCol)
    ans['Quarter'] = quarterCol.values

    df['Inv Date'] = pd.to_datetime(df['Inv Date'])
    ans['Year'] = df['Inv Date'].dt.year
    ans['Month'] = df['Inv Date'].dt.month
    ans['Day'] = df['Inv Date'].dt.day

    ans = ans.fillna('')
    ans = ans[:-1]

    return ans,row


def aireco(file_path,orgName,year,month):

    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df.iloc[:,2]
    ans['Total_Unit_Sales']=df.iloc[:,7]
    ans['Branch_Num']=df.iloc[:,0]
    ans['Invoice_Num']=df.iloc[:,5]
    ans['Product_Type']=df.iloc[:,3]

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    df['Invoicedate']=pd.to_datetime(df['Invoicedate'])
    ans['Year']=df['Invoicedate'].dt.year
    ans['Month']=df['Invoicedate'].dt.month
    ans['Day']=df['Invoicedate'].dt.day

    ans=ans.fillna('')

    return ans,row


def AmericanAirDistributing(file_path,orgName,year,month):

    #load the file to a data frame
    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls,"Sales Fields",skiprows=3)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['Model Number']
    ans['Total_Unit_Sales']=df['Number of Units']
    ans['Branch_Num']=df['Location Number']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Sale Date']=pd.to_datetime(df['Sale Date'])
    ans['Year']=df['Sale Date'].dt.year
    ans['Month']=df['Sale Date'].dt.month
    ans['Day']=df['Sale Date'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def APIofNH(file_path,orgName,year,month):

    #load the file to a data frame
    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls,skiprows=1)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['DESCRIPTION']
    ans['Total_Unit_Sales']=df['QTY']

    if 'BRANCH' in df.columns.values:
        ans['Branch_Num']=df['BRANCH']
    else:
        ans['Branch_Num']=df['BR']

    ans['Invoice_Num']=df['INVOICE #']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['INV DATE']=pd.to_datetime(df['INV DATE'])
    ans['Year']=df['INV DATE'].dt.year
    ans['Month']=df['INV DATE'].dt.month
    ans['Day']=df['INV DATE'].dt.day

    ans=ans.fillna('')

    return ans,row


def airefco(file_path,orgName,year,month):

    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df.iloc[:,4]
    ans['Total_Unit_Sales']=df.iloc[:,7]
    ans['Branch_Num']=df['Shipping WHSE']
    ans['Invoice_Num']=df.iloc[:,2]

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Date']=pd.to_datetime(df['Date'])
    ans['Year']=df['Date'].dt.year
    ans['Month']=df['Date'].dt.month
    ans['Day']=df['Date'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def associatedequipment(file_path,orgName,year,month):

    dict={1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'May',6:'Jun',
         7:'Jul',8:'Aug',9:'Sep',10:'Oct',11:'Nov',12:'Dec'}
    month_abb=dict[month]
    sheet_name=''

    xls=pd.ExcelFile(file_path)
    for tab_name in xls.sheet_names:
        if month_abb in tab_name:
            sheet_name=tab_name
            break
    df=pd.read_excel(xls,sheet_name)

    row,ans=createDF(df,orgName)

    # Copy other columns
    ans['Model_Num'] = df.iloc[:, 3]
    ans['Branch_Num'] = df.iloc[:, 7]
    ans['Total_Unit_Sales'] = df.iloc[:, 5]
    ans['Sales_Price'] = df.iloc[:, 6]
    ans['To_Zip']=df.iloc[:,8]

    priceCol = [pd.np.nan] * row
    priceCol = pd.Series(priceCol)

    for i in range(0, row):
        if ans['Total_Unit_Sales'].iloc[i] == 0:
            priceCol.iloc[i] = 0
        else:
            priceCol.iloc[i] = ans['Sales_Price'].iloc[i] / ans['Total_Unit_Sales'].iloc[i]

    ans['Sales_Price'] = priceCol.values

    # quarter
    quarter = ceil(month / 3)
    quarterCol = [quarter] * row
    quarterCol = pd.Series(quarterCol)
    ans['Quarter'] = quarterCol.values

    df['Inv Date'] = pd.to_datetime(df['Inv Date'],format='%Y%m%d')
    ans['Year'] = df['Inv Date'].dt.year
    ans['Month'] = df['Inv Date'].dt.month
    ans['Day'] = df['Inv Date'].dt.day

    ans = ans.fillna('')

    return ans,row


def AuerSteel(file_path,orgName,year,month):

    #load the file to a data frame
    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['Model Number']
    ans['Total_Unit_Sales']=df['Number of Units']
    ans['Branch_Num']=df['Location Number']
    ans['To_Zip']=df['Zip Code']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Sale Date']=pd.to_datetime(df['Sale Date'])
    ans['Year']=df['Sale Date'].dt.year
    ans['Month']=df['Sale Date'].dt.month
    ans['Day']=df['Sale Date'].dt.day

    ans=ans.fillna('')

    return ans,row


def behleryoung(file_path,orgName,year,month):

    df=pd.read_csv(file_path,encoding = "ISO-8859-1")

    row,ans=createDF(df,orgName)

    # Copy other columns
    ans['Model_Num'] = df.iloc[:, 0]
    ans['Branch_Num'] = df.iloc[:, 3]
    ans['Total_Unit_Sales'] = df.iloc[:, 2]
    ans['To_Zip']=df.iloc[:,6]
    ans['Invoice_Num']=df.iloc[:,4]

    # quarter
    quarter = ceil(month / 3)
    quarterCol = [quarter] * row
    quarterCol = pd.Series(quarterCol)
    ans['Quarter'] = quarterCol.values

    df['INVDATE']=pd.to_datetime(df['INVDATE'],format='%Y%m%d')
    ans['Year']=df['INVDATE'].dt.year
    ans['Month']=df['INVDATE'].dt.month
    ans['Day']=df['INVDATE'].dt.day

    ans = ans.fillna('')

    return ans,row


def BestChoice(file_path,orgName,year,month):

    #load the file to a data frame
    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['Model Number']
    ans['Total_Unit_Sales']=df['Number of Units']
    ans['Branch_Num']=df['Location Number']
    ans['To_Zip']=df['Delivery Zip Code']
    ans['Invoice_Num']=df['Invoice Number']
    ans['Sales_Price']=df['Sale Price']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Sale Date']=pd.to_datetime(df['Sale Date'])
    ans['Year']=df['Sale Date'].dt.year
    ans['Month']=df['Sale Date'].dt.month
    ans['Day']=df['Sale Date'].dt.day

    ans=ans.fillna('')

    return ans,row


def capitol(file_path,orgName,year,month):

    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df.iloc[:,2]
    ans['Total_Unit_Sales']=df.iloc[:,3]
    ans['Branch_Num']=df.iloc[:,0]

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['ShipDate']=pd.to_datetime(df['ShipDate'])
    ans['Year']=df['ShipDate'].dt.year
    ans['Month']=df['ShipDate'].dt.month
    ans['Day']=df['ShipDate'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def capco(file_path,orgName,year,month):

    #load the file to a data frame
    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls)

    df.columns.values[3]='Model Number'
    df.columns.values[5]='Desc_2'
    df.columns.values[6]='Qty'
    df.columns.values[8]='Price'
    df.columns.values[9]='Invoice Date'


    df=df.dropna(subset=['Order'])
    #df=df[df['Order']!='']
    df=df.fillna('')
    df=df.reset_index()
    df=df.drop(['index'],axis=1)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ModelCol=[pd.np.nan]*row
    ModelCol=pd.Series(ModelCol)

    for i in range(0,row):

        if df['Model Number'].iloc[i]=='':
            ModelCol.iloc[i]=df['DESC'].iloc[i]
        else:
            ModelCol.iloc[i]=df['Model Number'].iloc[i]
    ans['Model_Num']=ModelCol.values

    branchCol=[1]*row
    branchCol=pd.Series(branchCol)
    ans['Branch_Num']=branchCol.values

    ans['Total_Unit_Sales']=df['Qty']
    ans['Total_Unit_Sales']=pd.to_numeric(ans['Total_Unit_Sales'])
    ans['Product_Type']=df['Desc_2']
    ans['Invoice_Num']=df['Order']
    ans['Sales_Price']=df['Price']
    ans['Sales_Price']=pd.to_numeric(ans['Sales_Price'])

    priceCol=[pd.np.nan]*row
    priceCol=pd.Series(priceCol)

    for i in range(0,row):

        if ans['Total_Unit_Sales'].iloc[i]==0:
            priceCol.iloc[i]=0
        else:
            priceCol.iloc[i]=ans['Sales_Price'].iloc[i]/ans['Total_Unit_Sales'].iloc[i]

    ans['Sales_Price']=priceCol.values

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Invoice Date']=pd.to_datetime(df['Invoice Date'])
    ans['Year']=df['Invoice Date'].dt.year
    ans['Month']=df['Invoice Date'].dt.month
    ans['Day']=df['Invoice Date'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def CarrSupply(file_path,orgName,year,month):

    #load the file to a data frame
    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls,"Sales Fields",skiprows=3)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['Model Number']
    ans['Total_Unit_Sales']=df['Number of Units']
    ans['Branch_Num']=df['Location Number']
    ans['Invoice_Num']=df['Invoice Number']
    ans['Sales_Price']=df['Sale Price']
    ans['To_Zip']=df['Delivery ZIP Code']

    if orgName=='Weathertech Distributing Co':
        ans['Sales_Price']=ans['Sales_Price']/ans['Total_Unit_Sales']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Sale Date']=pd.to_datetime(df['Sale Date'])
    ans['Year']=df['Sale Date'].dt.year
    ans['Month']=df['Sale Date'].dt.month
    ans['Day']=df['Sale Date'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def CenturyACSupply(file_path,orgName,year,month):

    #load the file to a data frame
    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls,skiprows=2)

    row,ans=createDF(df,orgName)

    #Copy other columns
    #df.columns.values[5]='Item Desc'
    ans['Model_Num']=df['Item Desc']
    #df.columns.values[6]='Qty Ordered'
    ans['Total_Unit_Sales']=df['Qty Ordered']
    #df.columns.values[8]='Sales Location ID'
    ans['Branch_Num']=df['Sales Location ID']
    #df.columns.values[0]='Invoice No'
    ans['Invoice_Num']=df['Invoice No']
    #df.columns.values[7]='Sales'
    ans['Sales_Price']=df['Sales']

    ans['Sales_Price']=ans['Sales_Price']/ans['Total_Unit_Sales']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Invoice Date']=pd.to_datetime(df['Invoice Date'])
    ans['Year']=df['Invoice Date'].dt.year
    ans['Month']=df['Invoice Date'].dt.month
    ans['Day']=df['Invoice Date'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def CAD(file_path,orgName,year,month):

    #load the file to a data frame
    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls)

    df=df[df['invoice_no']!=0]
    df=df.reset_index()
    df=df.drop(['index'],axis=1)

    df=df.dropna(subset=['invoice_no'])
    df.fillna('')
    df=df.reset_index()
    df=df.drop(['index'],axis=1)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['item_id']
    ans['Total_Unit_Sales']=df['units']
    ans['Branch_Num']=df['branch_id']
    ans['Invoice_Num']=df['invoice_no']
    ans['Product_Type']=df['item_desc']
    ans['Year']=df['year']
    ans['Month']=df['month']
    ans['Sales_Price']=df['sales']
    ans['To_Zip']=df['mail_postal_code']

    ans['Sales_Price']=ans['Sales_Price']/ans['Total_Unit_Sales']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #day
    day=1
    DayCol=[day]*row
    DayCol=pd.Series(DayCol)
    ans['Day']=DayCol.values

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def cdj(file_names,orgName,year,month):

    n=0
    global df
    missCol=-1

    for file in file_names:
        xls=pd.ExcelFile(file)
        for tab in xls.sheet_names:
            subDF=pd.read_excel(xls,sheet_name=tab)
            if subDF.shape[1]<8:
                continue
            elif subDF.shape[1]==9:
                subDF=cleanNine(subDF,missCol)
            subDF=cleanCDJ(subDF)
            if n==0:
                df=subDF
            else:
                df=df.append(subDF,ignore_index=True)
            n+=1

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['Product']
    ans['Branch_Num']=df['Branch']
    ans['Product_Type']=df['Description']

    qtyCol=df['Quantity']
    saleCol=df['Net Sale']
    priceCol=[pd.np.nan]*row
    priceCol=pd.Series(priceCol)

    for i in range(0,row):
        if '-' in str(qtyCol.iloc[i]):
            qtyCol.iloc[i]=qtyCol.iloc[i].replace('-','')
            qtyCol.iloc[i]=int(float(qtyCol.iloc[i]))*(-1)
        qtyCol.iloc[i]=int(float(qtyCol.iloc[i]))

        if '-' in str(saleCol.iloc[i]):
            saleCol.iloc[i]=saleCol.iloc[i].replace('-','')
            saleCol.iloc[i]=int(float(saleCol.iloc[i]))*(-1)
        saleCol.iloc[i]=int(float(saleCol.iloc[i]))

        if qtyCol.iloc[i]==0:
            priceCol.iloc[i]=0
        else:
            priceCol.iloc[i]=saleCol.iloc[i]/qtyCol.iloc[i]

    ans['Sales_Price']=priceCol.values
    ans['Total_Unit_Sales']=qtyCol.values

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    day=1
    dayCol=[day]*row
    dayCol=pd.Series(dayCol)
    ans['Day']=dayCol.values

    monthCol=[month]*row
    monthCol=pd.Series(monthCol)
    ans['Month']=monthCol.values

    yearCol=[year]*row
    yearCol=pd.Series(yearCol)
    ans['Year']=yearCol.values

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def CSPC(file_path,orgName,year,month):

    #load the file to a data frame
    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['ModelNo']
    ans['Total_Unit_Sales']=df['Qty']
    ans['Branch_Num']=df['Whseid']
    ans['Invoice_Num']=df['InvNo']
    ans['Sales_Price']=df['Price Ea.']
    ans['To_Zip']=df['Zip Code']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['InvDate']=pd.to_datetime(df['InvDate'])
    ans['Year']=df['InvDate'].dt.year
    ans['Month']=df['InvDate'].dt.month
    ans['Day']=df['InvDate'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def comfortsupply(file_path,orgName,year,month):

	df=pd.read_csv(file_path,skiprows=2)

	df=df.fillna('')
	df=df[df['Invoice No']!='']
	df=df.reset_index()
	df=df.drop(['index'],axis=1)

	row,ans=createDF(df,orgName)

	#Copy other columns
	ans['Model_Num']=df['Item ID']
	ans['Product_Type']=df['Item Desc']
	ans['Total_Unit_Sales']=df['Qty Shipped']
	ans['Branch_Num']=df['Branch ID (Inv)']
	ans['Invoice_Num']=df['Invoice No']
	ans['Sales_Price'] = df['Sales']
	ans['Sales_Price']=ans['Sales_Price'].replace('[\(,]', '', regex=True)
	ans['Sales_Price']=ans['Sales_Price'].replace('[\),]', '', regex=True)
	ans['Sales_Price']=ans['Sales_Price'].replace('[\$,]', '', regex=True)
	ans['Sales_Price']=ans['Sales_Price'].astype(float)
	ans['To_Zip']=df['Ship2 Zip']

	priceCol = [pd.np.nan] * row
	priceCol = pd.Series(priceCol)

	for i in range(0, row):
		if ans['Total_Unit_Sales'].iloc[i] == 0:
			priceCol.iloc[i] = 0
		else:
			priceCol.iloc[i] = ans['Sales_Price'].iloc[i] / ans['Total_Unit_Sales'].iloc[i]

	ans['Sales_Price'] = priceCol.values

	#quarter
	quarter=ceil(month/3)
	quarterCol=[quarter]*row
	quarterCol=pd.Series(quarterCol)
	ans['Quarter']=quarterCol.values

	#dates
	df['Invoice Date']=pd.to_datetime(df['Invoice Date'])
	ans['Year']=df['Invoice Date'].dt.year
	ans['Month']=df['Invoice Date'].dt.month
	ans['Day']=df['Invoice Date'].dt.day

	ans = ans.fillna('')

	return ans, row


def crescent(file_path,orgName,year,month):

    df=pd.read_csv(file_path)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['Model']
    ans['Total_Unit_Sales']=df['QTY OR LBS']
    ans['Total_Unit_Sales']=ans['Total_Unit_Sales'].replace('[\,,]', '', regex=True)
    ans['Total_Unit_Sales']=pd.to_numeric(ans['Total_Unit_Sales'])

    ans['Branch_Num']=df['BranchNumber']
    ans.loc[ans.Branch_Num=='ELEC','Branch_Num']='INDP'

    ans['Sales_Price'] = df['UnitAmount']
    ans['Sales_Price']=ans['Sales_Price'].replace('[\,,]', '', regex=True)
    ans['To_Zip']=df['ZIP']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['InvoiceDate']=pd.to_datetime(df['InvoiceDate'])
    ans['Year']=df['InvoiceDate'].dt.year
    ans['Month']=df['InvoiceDate'].dt.month
    ans['Day']=df['InvoiceDate'].dt.day

    ans = ans.fillna('')

    return ans, row


def benoist(file_path,orgName,year,month):

    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls,skiprows=1)

    row,ans=createDF(df,orgName)

    # Copy other columns
    ans['Model_Num'] = df.iloc[:, 2]
    ans['Branch_Num'] = df.iloc[:, 7]
    ans['Total_Unit_Sales'] = df.iloc[:, 6]
    ans['To_Zip']=df.iloc[:,8]
    ans['Invoice_Num']=df.iloc[:,9]
    ans['Product_Type']=df.iloc[:,3]

    # quarter
    quarter = ceil(month / 3)
    quarterCol = [quarter] * row
    quarterCol = pd.Series(quarterCol)
    ans['Quarter'] = quarterCol.values

    df['InvoiceDate']=pd.to_datetime(df['InvoiceDate'])
    ans['Year']=df['InvoiceDate'].dt.year
    ans['Month']=df['InvoiceDate'].dt.month
    ans['Day']=df['InvoiceDate'].dt.day

    ans = ans.fillna('')

    return ans, row


def cfmdis(file_path,orgName,year,month):

    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls)

    row,ans=createDF(df,orgName)

    # Copy other columns
    ans['Model_Num'] = df.iloc[:, 0]
    ans['Branch_Num'] = df.iloc[:, 1]
    ans['Total_Unit_Sales'] = df.iloc[:, 2]

    # quarter
    quarter = ceil(month / 3)
    quarterCol = [quarter] * row
    quarterCol = pd.Series(quarterCol)
    ans['Quarter'] = quarterCol.values
    #dates
    day=1
    dayCol=[day]*row
    dayCol=pd.Series(dayCol)
    ans['Day']=dayCol.values

    monthCol=[month]*row
    monthCol=pd.Series(monthCol)
    ans['Month']=monthCol.values

    yearCol=[year]*row
    yearCol=pd.Series(yearCol)
    ans['Year']=yearCol.values

    #get rid of nans
    ans=ans.fillna('')

    return ans, row


def designAir(file_path,orgName,year,month):

    #load the file to a data frame
    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls)
    row,ans=createDF(df,orgName)

    #Copy other columns
    ModelCol=[pd.np.nan]*row
    ModelCol=pd.Series(ModelCol)

    for i in range(0,row):

        if df['supplier_part_no'].iloc[i]=='':
            ModelCol.iloc[i]=df['item_desc'].iloc[i]
        else:
            ModelCol.iloc[i]=df['supplier_part_no'].iloc[i]
    ans['Model_Num']=ModelCol.values

    ans['Branch_Num']=df['location id']
    ans['Total_Unit_Sales']=df['qty_shipped']
    ans['Product_Type']=df['item_desc']
    ans['Sales_Price']=df['unit_price']
    ans['To_Zip']=df['zip_code']
    ans['Invoice_Num']=df['invoice_no']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['InvoiceDate']=pd.to_datetime(df['InvoiceDate'])
    ans['Year']=df['InvoiceDate'].dt.year
    ans['Month']=df['InvoiceDate'].dt.month
    ans['Day']=df['InvoiceDate'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def dcne(file_path,orgName,year,month):

    #load the file to a data frame
    # xls=pd.ExcelFile(file_path)
    # df=pd.read_excel(xls,sheet_name=1)
    dict={1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'May',6:'Jun',
         7:'Jul',8:'Aug',9:'Sep',10:'Oct',11:'Nov',12:'Dec'}
    month_abb=dict[month]
    sheet_name=''

    xls=pd.ExcelFile(file_path)
    for tab_name in xls.sheet_names:
        if month_abb in tab_name:
            sheet_name=tab_name
            break
    df=pd.read_excel(xls,sheet_name)

    row,ans=createDF(df,orgName)

    ans['Model_Num']=df['Prod Id']
    ans['Branch_Num']=df['Branch']
    ans['Total_Unit_Sales']=df['Number of Units']
    ans['Sales_Price']=df['Unit Price']
    ans['To_Zip']=df['Delivery Zip']
    ans['Invoice_Num']=df['Order Number']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Sale Date']=pd.to_datetime(df['Sale Date'])
    ans['Year']=df['Sale Date'].dt.year
    ans['Month']=df['Sale Date'].dt.month
    ans['Day']=df['Sale Date'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def dsc(file_names,orgName,year,month):

    # print('please re-select their ducted and ductless data')
    #
    # file_path = filedialog.askopenfilename()
    # xls=pd.ExcelFile(file_path)
    # df1=pd.read_excel(xls,skiprows=1)
    #
    # file_path = filedialog.askopenfilename()
    #
    # xls=pd.ExcelFile(file_path)
    # df2=pd.read_excel(xls,skiprows=1)
    # df=df1.append(df2,ignore_index=True)
    file_names=list(file_names)

    global df_bo

    for i in range(0,len(file_names)):
        file_path=file_names[i]
        file=os.path.split(file_path)[1]
        file_name=file.split('.')[0]

        if 'Bleed Over' in file_name:
            xls=pd.ExcelFile(file_path)
            df_bo=pd.read_excel(xls,header=None)
            df_bo.columns=['Organization','Branch_Num','To_Zip','From_Zip','Model_Num','Product_Type','Invoice_Num','Total_Unit_Sales','Sales_Price','Year','Quarter','Month','Day']
            del file_names[i]
            break

    xls1=pd.ExcelFile(file_names[0])
    xls2=pd.ExcelFile(file_names[1])
    df1=pd.read_excel(xls1,skiprows=1)
    df2=pd.read_excel(xls2,skiprows=1)
    df=df1.append(df2,ignore_index=True)

    row,ans=createDF(df,orgName)

    ans['Model_Num']=df['Model Number']
    ans['Branch_Num']=df['Location Number']
    ans['Total_Unit_Sales']=pd.to_numeric(df['Number of Units'])
    ans['Sales_Price']=df['Sale Price']
    ans['Invoice_Num']=df['Invoice Number']
    ans['To_Zip']=df['Code']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Sale Date']=pd.to_datetime(df['Sale Date'])
    ans['Year']=df['Sale Date'].dt.year
    ans['Month']=df['Sale Date'].dt.month
    ans['Day']=df['Sale Date'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def edSupply(file_path,orgName,year,month):

    #load the file to a data frame
    xls=pd.ExcelFile(file_path)
    df1=pd.read_excel(xls,sheet_name=1)
    df2=pd.read_excel(xls,sheet_name=2)

    df=pd.concat([df1,df2])

    df[['Model Number','Extra']]=df['Model Number'].str.split(' RHEEM ',expand=True)
    df=df.drop(['Extra'],axis=1)
    df.columns.values[1]='Sale Date'

    df=df.reset_index()
    df=df.drop(['index'],axis=1)

    row,ans=createDF(df,orgName)

    ans['Model_Num']=df['Model Number']
    ans['Branch_Num']=df['Branch']
    ans['Total_Unit_Sales']=df['Quantity']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Sale Date']=pd.to_datetime(df['Sale Date'])
    ans['Year']=df['Sale Date'].dt.year
    ans['Month']=df['Sale Date'].dt.month
    ans['Day']=df['Sale Date'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def ferguson(file_path,orgName,year,month):

    # xls=pd.ExcelFile(file_path)
    # df=pd.read_excel(xls)
    df=pd.read_csv(file_path)

    #df=df[(df['DET1']=='EQUIPMENT')|(df['DET1']=='HYDRONICS')|(df['DET1']=='OTHER')]
    df=df[(df['DET1']=='EQUIPMENT')|(df['DET1']=='HYDRONICS')]
    df=df.dropna(subset=['VENDOR_CODE'])
    df=df.reset_index()
    df=df.drop(['index'],axis=1)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['VENDOR_CODE']

    df['SUM(D.SHIPPED_QTY)']=pd.to_numeric(df['SUM(D.SHIPPED_QTY)'])
    ans['Total_Unit_Sales']=df['SUM(D.SHIPPED_QTY)']

    ans['Branch_Num']=df['SELL_WAREHOUSE_NUMBER_NK']
    ans['To_Zip']=df['SHIP_TO_ZIP']
    ans['Invoice_Num']=df['INVOICE_NUMBER_NK']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['PROCESS_DATE']=pd.to_datetime(df['PROCESS_DATE'])
    ans['Year']=df['PROCESS_DATE'].dt.year
    ans['Month']=df['PROCESS_DATE'].dt.month
    ans['Day']=df['PROCESS_DATE'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def gwberkheimer(file_path,orgName,year,month):

    df=pd.read_csv(file_path)
    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df.iloc[:,0]
    ans['Total_Unit_Sales']=df.iloc[:,2]
    ans['Total_Unit_Sales']=pd.to_numeric(ans['Total_Unit_Sales'])
    ans['Branch_Num']=df.iloc[:,3]
    ans['Invoice_Num']=df.iloc[:,4]
    ans['Sales_Price'] = df.iloc[:,5]

    zipCol=[pd.np.nan]*row
    zipCol=pd.Series(zipCol)

    for i in range(0,row):
        zipCol.iloc[i]=str(df['ZIP CODE'].iloc[i])[:5]

    ans['To_Zip']=zipCol.values

    priceCol = [pd.np.nan] * row
    priceCol = pd.Series(priceCol)

    for i in range(0, row):
        if ans['Total_Unit_Sales'].iloc[i] == 0:
            priceCol.iloc[i] = 0
        else:
            priceCol.iloc[i] = ans['Sales_Price'].iloc[i] / ans['Total_Unit_Sales'].iloc[i]

    ans['Sales_Price'] = priceCol.values

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['SALES DATE']=pd.to_datetime(df['SALES DATE'])
    ans['Year']=df['SALES DATE'].dt.year
    ans['Month']=df['SALES DATE'].dt.month
    ans['Day']=df['SALES DATE'].dt.day

    ans = ans.fillna('')

    return ans, row


def gearypacific(file_path,orgName,year,month):

    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls,skiprows=4)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df.iloc[:,5]
    ans['Product_Type']=df.iloc[:,6]
    ans['Total_Unit_Sales']=df.iloc[:,7]
    ans['Branch_Num']=df.iloc[:,8]
    ans['Invoice_Num']=df.iloc[:,9]
    ans['Sales_Price']=df.iloc[:,11]
    ans['To_Zip']=df.iloc[:,12]

    priceCol=[pd.np.nan]*row
    priceCol=pd.Series(priceCol)

    for i in range(0,row):
        if ans['Total_Unit_Sales'].iloc[i]==0:
            priceCol.iloc[i]=0
        else:
            priceCol.iloc[i]=ans['Sales_Price'].iloc[i]/ans['Total_Unit_Sales'].iloc[i]

    ans['Sales_Price']=priceCol.values

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    dateCol=[pd.np.nan]*row
    dateCol=pd.Series(dateCol)

    for i in range (0,row):
        if len(str(df['Inv Date'].iloc[i]))>5:
            dateCol.iloc[i]=df['Inv Date'].iloc[i]
        elif len(str(df['Inv Date'].iloc[i]))==5:
            dateCol.iloc[i]=datetime.timedelta(df['Inv Date'].iloc[i])+datetime.datetime(1899,12,30)

    #dates
    df['Inv Date'] = dateCol.values
    ans['Year']=df['Inv Date'].dt.year
    ans['Month']=df['Inv Date'].dt.month
    ans['Day']=df['Inv Date'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans, row


def gustave(file_path,orgName,year,month):

    xls=pd.ExcelFile(file_path)
    if len(xls.sheet_names)>1:
        df=pd.read_excel(xls,"Sales Fields",skiprows=3)
    else:
        df=pd.read_excel(xls,skiprows=3)

    df=df.fillna('')
    df=df[df['Invoice Number']!='']
    df=df.reset_index()
    df=df.drop(['index'],axis=1)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['Model Number']
    ans['Total_Unit_Sales']=df['Number of Units']
    ans['Branch_Num']=df['Location Number']
    ans['Invoice_Num']=df['Invoice Number']
    ans['To_Zip']=df['Delivery ZIP Code']
    ans['Sales_Price']=df['Gross Sales $']

    priceCol=[pd.np.nan]*row
    priceCol=pd.Series(priceCol)

    for i in range(0,row):
        if ans['Total_Unit_Sales'].iloc[i]==0:
            priceCol.iloc[i]=0
        else:
            priceCol.iloc[i]=ans['Sales_Price'].iloc[i]/ans['Total_Unit_Sales'].iloc[i]

    ans['Sales_Price']=priceCol.values

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Sale Date']=pd.to_datetime(df['Sale Date'])
    ans['Year']=df['Sale Date'].dt.year
    ans['Month']=df['Sale Date'].dt.month
    ans['Day']=df['Sale Date'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans, row


def hcs(file_path,orgName,year,month):

    xls=pd.ExcelFile(file_path)
    df1=pd.read_excel(xls,sheet_name=0)
    df2=pd.read_excel(xls,sheet_name=1)
    df=pd.concat([df1,df2])

    df=df.reset_index()
    df=df.drop(['index'],axis=1)

    row,ans=createDF(df,orgName)

    ans['Model_Num']=df['Model Number']
    ans['Branch_Num']=df['Location Number']
    ans['Total_Unit_Sales']=df['Number of Units']
    ans['Sales_Price']=df['Sale Price']
    ans['Invoice_Num']=df['Invoice Number']
    ans['To_Zip']=df['Delivery Zip Code']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Sale Date']=pd.to_datetime(df['Sale Date'])
    ans['Year']=df['Sale Date'].dt.year
    ans['Month']=df['Sale Date'].dt.month
    ans['Day']=df['Sale Date'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def HerculesIndustries(file_path,orgName,year,month):

    #load the file to a data frame
    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['Code (Product)']
    ans['Total_Unit_Sales']=df['Quantity']
    ans['Branch_Num']=df['WH']
    ans['Invoice_Num']=df['Inv No']
    ans['To_Zip']=df['Zip (Customer)']
    ans['Product_Type']=df['Name (Product)']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Date']=pd.to_datetime(df['Date'])
    ans['Year']=df['Date'].dt.year
    ans['Month']=df['Date'].dt.month
    ans['Day']=df['Date'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def HVACDistributors(file_path,orgName,year,month):

    #load the file to a data frame
    df=pd.read_csv(file_path,skiprows=8)
    df=df.fillna('')
    df=df[df['INVOICE NUMBER']!='']
    df=df.reset_index()
    df=df.drop(['index'],axis=1)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['MODEL NUMBER']
    ans['Total_Unit_Sales']=df['NUMBER']
    ans['Total_Unit_Sales']=pd.to_numeric(ans['Total_Unit_Sales'])
    ans['Branch_Num']=df['BRANCH']
    ans['Invoice_Num']=df['INVOICE NUMBER']
    ans['Sales_Price']=df['SALE PRICE']
    ans['To_Zip']=df['ZIP']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['SALE DATE']=pd.to_datetime(df['SALE DATE'])
    ans['Year']=df['SALE DATE'].dt.year
    ans['Month']=df['SALE DATE'].dt.month
    ans['Day']=df['SALE DATE'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def HSSC(file_path,orgName,year,month):

    #load the file to a data frame
    df=pd.read_csv(file_path)
    df['Product Number']=df['Product Number'].str.replace('*','')

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['Product Number']
    ans['Total_Unit_Sales']=df['Product Quantity Shipped']
    ans['Sales_Price']=df['Product Unit Price']
    ans['Branch_Num']=df['Invoice Warehouse']
    ans['Invoice_Num']=df['Invoice Number']
    ans['To_Zip']=df['Invoice ShipTo Zip']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Invoice Date']=pd.to_datetime(df['Invoice Date'])
    ans['Year']=df['Invoice Date'].dt.year
    ans['Month']=df['Invoice Date'].dt.month
    ans['Day']=df['Invoice Date'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def interline(file_path,orgName,year,month):

    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls,skiprows=1)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['manufacturer_part_number']
    ans['Total_Unit_Sales']=df['Qty_Shipped']
    ans['Branch_Num']=df['warehouse_num']
    ans['Invoice_Num']=df['invoice_num']
    ans['To_Zip']=df['ship_to_postal_code']
    ans['Sales_Price']=df['price']

    priceCol=[pd.np.nan]*row
    priceCol=pd.Series(priceCol)

    for i in range(0,row):
        if ans['Total_Unit_Sales'].iloc[i]==0:
            priceCol.iloc[i]=0
        else:
            priceCol.iloc[i]=ans['Sales_Price'].iloc[i]/ans['Total_Unit_Sales'].iloc[i]

    ans['Sales_Price']=priceCol.values

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['invoice_date']=pd.to_datetime(df['invoice_date'].astype(str))
    ans['Year']=df['invoice_date'].dt.year
    ans['Month']=df['invoice_date'].dt.month
    ans['Day']=df['invoice_date'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans, row


def johnsonSupply(file_path,orgName,year,month):

    #load the file to a data frame
    # xls=pd.ExcelFile(file_path)
    # df=pd.read_excel(xls,skiprows=2)
    df=pd.read_csv(file_path,skiprows=2)
    df=df.fillna('')
    #df=df[df['Invoice#']!='']
    df=df.dropna(subset=['Invoice#'])
    df=df.reset_index()
    df=df.drop(['index'],axis=1)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['Description']
    ans['Total_Unit_Sales']=df['Shp-Qty']
    ans['Branch_Num']=df['Cust_Terr']
    ans['Invoice_Num']=df['Invoice#']
    ans['To_Zip']=df['Zip']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Inv-Date']=pd.to_datetime(df['Inv-Date'])
    ans['Year']=df['Inv-Date'].dt.year
    ans['Month']=df['Inv-Date'].dt.month
    ans['Day']=df['Inv-Date'].dt.day

    ans=ans[ans['Month']==month]

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def jsPopma(file_path,orgName,year,month):

    #load the file to a data frame
    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls,skiprows=1)
    df.columns.values[2]='Units'
    df.columns.values[3]='WH'
    df.columns.values[4]='Invoice'

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['Model Number']
    ans['Total_Unit_Sales']=df['Units']
    ans['Branch_Num']=df['WH']
    ans['Invoice_Num']=df['Invoice']
    ans['To_Zip']=df['ZIP Code']
    ans['Sales_Price']=df['Price']

    ans['Sales_Price']=ans['Sales_Price']/ans['Total_Unit_Sales']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Sale Date']=pd.to_datetime(df['Sale Date'])
    ans['Year']=df['Sale Date'].dt.year
    ans['Month']=df['Sale Date'].dt.month
    ans['Day']=df['Sale Date'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def kochAir(file_path,orgName,year,month):

    xls=pd.ExcelFile(file_path)
    #df=pd.read_excel(xls,sheet_name=1)
    df=pd.read_excel(xls)

    s=pd.Series(['BE','CE','HO','PE','CP','BP','BV','BF','CH','EL','FC','FD','FI','FS','HA','IN','LS','MA','MF','MR','PG','PV','RP','SH','TA','VT','WH','WN'])

    row,ans=createDF(df,orgName)

    #Copy other columns

    ModelCol=[pd.np.nan]*row
    ModelCol=pd.Series(ModelCol)

    for i in range(0,len(df['Item Number'].values)):
        model=str(df['Item Number'].iloc[i])
        first2=model[:2]
        if first2 in s.values:
            model=model[2:]
            ModelCol.iloc[i]=model
        else:
            ModelCol.iloc[i]=model
    ans['Model_Num']=ModelCol.values

    ans['Total_Unit_Sales']=df['MOVEMENT']
    ans['Branch_Num']=df['Location']
    ans['To_Zip']=df['ZIP']
    ans['Invoice_Num']=df['INVNO']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['INVDT']=pd.to_datetime(df['INVDT'])
    ans['Year']=df['INVDT'].dt.year
    ans['Month']=df['INVDT'].dt.month
    ans['Day']=df['INVDT'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def lockeSupply(file_path,orgName,year,month):

    #load the file to a data frame
    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['VendorPrt#']
    ans['Total_Unit_Sales']=df['Qty']
    ans['Branch_Num']=df['whse']
    ans['Invoice_Num']=df['InvoiceNumber']
    ans['To_Zip']=df['zipcd']
    ans['Product_Type']=df['Description']
    ans['Sales_Price']=df['SALEPRICE']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['DateOfSale']=pd.to_datetime(df['DateOfSale'])
    ans['Year']=df['DateOfSale'].dt.year
    ans['Month']=df['DateOfSale'].dt.month
    ans['Day']=df['DateOfSale'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def lohmiller(file_path,orgName,year,month):

    #load the file to a data frame
    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['item_id']
    ans['Total_Unit_Sales']=df['qty_shipped']
    ans['Branch_Num']=df['branch_id']
    ans['Invoice_Num']=df['invoice_no']
    ans['To_Zip']=df['ship2_postal_code']
    ans['Product_Type']=df['item_desc']
    ans['Sales_Price']=df['unit_price']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['invoice_date']=pd.to_datetime(df['invoice_date'])
    ans['Year']=df['invoice_date'].dt.year
    ans['Month']=df['invoice_date'].dt.month
    ans['Day']=df['invoice_date'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def luce(file_path,orgName,year,month):

    xls=pd.ExcelFile(file_path)

    global df
    for i in range(0,len(xls.sheet_names)):
        if i==0:
            df=pd.read_excel(xls,sheet_name=i)
        else:
            temp=pd.read_excel(xls,sheet_name=i)
            df=pd.concat([df,temp])

    df=df.fillna('')
    df=df[df['Invoice Number']!='']
    df=df.reset_index()
    df=df.drop(['index'],axis=1)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['Model Number']
    ans['Total_Unit_Sales']=df['Number of Units']
    ans['Branch_Num']=df['Location Number']
    ans['Sales_Price'] = df['Sale Price']
    ans['To_Zip']=df['Delivery ZIP Code']
    ans['Invoice_Num']=df['Invoice Number']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Sale Date']=pd.to_datetime(df['Sale Date'])
    ans['Year']=df['Sale Date'].dt.year
    ans['Month']=df['Sale Date'].dt.month
    ans['Day']=df['Sale Date'].dt.day

    ans = ans.fillna('')

    return ans, row


def m_and_a(file_path,orgName,year,month):

    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['supplier_part_no']
    ans['Product_Type']=df['item_desc']
    ans['Total_Unit_Sales']=df['qty_shipped']
    ans['Branch_Num']=df['sales_location_id']
    ans['Sales_Price'] = df['unit_price']
    ans['To_Zip']=df['ship2_postal_code']
    ans['Invoice_Num']=df['invoice_no']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['invoice_date']=pd.to_datetime(df['invoice_date'])
    ans['Year']=df['invoice_date'].dt.year
    ans['Month']=df['invoice_date'].dt.month
    ans['Day']=df['invoice_date'].dt.day

    ans = ans.fillna('')

    return ans, row


def mccall(file_path,orgName,year,month):
    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df.iloc[:,4]
    ans['Total_Unit_Sales']=df.iloc[:,6]
    ans['Branch_Num']=df.iloc[:,0]
    ans['To_Zip']=df.iloc[:,3]

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    day=1
    dayCol=[day]*row
    dayCol=pd.Series(dayCol)
    ans['Day']=dayCol.values

    monthCol=[month]*row
    monthCol=pd.Series(monthCol)
    ans['Month']=monthCol.values

    yearCol=[year]*row
    yearCol=pd.Series(yearCol)
    ans['Year']=yearCol.values

    ans = ans.fillna('')

    return ans, row


def minnesota(file_path,orgName,year,month):

    #load the file to a data frame
    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['Prodid']
    ans['Total_Unit_Sales']=df['Quantity']
    ans['Branch_Num']=df['Whseid']
    ans['Invoice_Num']=df['Ordernumber']
    ans['Sales_Price']=df['Price']
    ans['To_Zip']=df['Custzip']
    ans['Product_Type']=df['Proddesc']

    ans['Sales_Price']=ans['Sales_Price']/ans['Total_Unit_Sales']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates

    df['Invoicedate']=pd.to_datetime(df['Invoicedate'])
    ans['Year']=df['Invoicedate'].dt.year
    ans['Month']=df['Invoicedate'].dt.month
    ans['Day']=df['Invoicedate'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def mingledorff(file_path,orgName,year,month):

    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df.iloc[:,3]
    ans['Total_Unit_Sales']=df.iloc[:,5]
    ans['Branch_Num']=df.iloc[:,0]
    ans['Product_Type']=df.iloc[:,4]
    ans['Sales_Price']=df.iloc[:,6]
    ans['To_Zip']=df.iloc[:,8]

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Inv date']=pd.to_datetime(df['Inv date'])
    ans['Year']=df['Inv date'].dt.year
    ans['Month']=df['Inv date'].dt.month
    ans['Day']=df['Inv date'].dt.day

    ans = ans.fillna('')

    return ans, row


def meier(file_path,orgName,year,month):

    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df.iloc[:,0]
    ans['Total_Unit_Sales']=df.iloc[:,1]
    ans['Branch_Num']=df.iloc[:,3]

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['invoice_date']=pd.to_datetime(df['invoice_date'])
    ans['Year']=df['invoice_date'].dt.year
    ans['Month']=df['invoice_date'].dt.month
    ans['Day']=df['invoice_date'].dt.day

    ans = ans.fillna('')

    return ans, row


def munch(file_path,orgName,year,month):

    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['Model']
    ans['Total_Unit_Sales']=df['Ship Qty']
    ans['Branch_Num']=df['Shipping Branch']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Ship/Rec. Date']=pd.to_datetime(df['Ship/Rec. Date'])
    ans['Year']=df['Ship/Rec. Date'].dt.year
    ans['Month']=df['Ship/Rec. Date'].dt.month
    ans['Day']=df['Ship/Rec. Date'].dt.day

    ans=ans.fillna('')

    return ans, row


def morsco(file_path,orgName,year,month):

    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls,"Sales Fields",skiprows=3)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['Model Number']
    ans['Total_Unit_Sales']=df['Number of Units']
    ans['Branch_Num']=df['Location Number']
    ans['Invoice_Num']=df['Invoice Number']
    ans['Sales_Price']=df['Sale Price']
    ans['Sales_Price']=ans['Sales_Price']/ans['Total_Unit_Sales']
    ans['To_Zip']=df['Delivery ZIP Code']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Sale Date']=pd.to_datetime(df['Sale Date'])
    ans['Year']=df['Sale Date'].dt.year
    ans['Month']=df['Sale Date'].dt.month
    ans['Day']=df['Sale Date'].dt.day

    ans=ans.fillna('')

    return ans, row


def nbhandy(file_path,orgName,year,month):

    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['Code (Product)']
    ans['Total_Unit_Sales']=df['Quantity']
    ans['Branch_Num']=df['Name (Assigned Branch)']
    ans['Invoice_Num']=df['Inv No']
    ans['Sales_Price']=df['Sales']/df['Quantity']

    zipCol=[pd.np.nan]*row
    zipCol=pd.Series(zipCol)

    for i in range(0,row):
        zipCol.iloc[i]=str(df['Zip (Customer)'].iloc[i])[:5]

    ans['To_Zip']=zipCol.values

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Date']=pd.to_datetime(df['Date'])
    ans['Year']=df['Date'].dt.year
    ans['Month']=df['Date'].dt.month
    ans['Day']=df['Date'].dt.day

    ans=ans.fillna('')

    return ans, row


def oconnor(file_path,orgName,year,month):

    df=pd.read_csv(file_path,skiprows=2)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['Item ID']
    ans['Total_Unit_Sales']=df['Qty Shipped']
    ans['Branch_Num']=df['Sales Location ID']
    ans['Invoice_Num']=df['Invoice No']
    ans['Sales_Price']=df['Unit Price']
    ans['To_Zip']=df['Ship2 Postal Code']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Invoice Date']=pd.to_datetime(df['Invoice Date'])
    ans['Year']=df['Invoice Date'].dt.year
    ans['Month']=df['Invoice Date'].dt.month
    ans['Day']=df['Invoice Date'].dt.day

    ans=ans.fillna('')

    return ans, row


def peirce(file_path,orgName,year,month):

    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls,skiprows=2)
    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['Item ID']
    ans['Total_Unit_Sales']=df['Qty Shipped']
    ans['Branch_Num']=df['Ship Location Id']
    ans['Invoice_Num']=df['Invoice No']
    ans['To_Zip']=df['Ship2 Postal Code']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Invoice Date']=pd.to_datetime(df['Invoice Date'])
    ans['Year']=df['Invoice Date'].dt.year
    ans['Month']=df['Invoice Date'].dt.month
    ans['Day']=df['Invoice Date'].dt.day

    ans=ans.fillna('')

    return ans, row


def remichel(file_path,orgName,year,month):

    df=pd.read_table(file_path,header=None)
    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df.iloc[:,5]
    ans['Total_Unit_Sales']=df.iloc[:,7]
    ans['Branch_Num']=df.iloc[:,0]
    ans['Invoice_Num']=df.iloc[:,1]
    ans['To_Zip']=df.iloc[:,4]
    ans['Product_Type']=df.iloc[:,6]

    zipCol=[pd.np.nan]*row
    zipCol=pd.Series(zipCol)

    for i in range(0,row):
        zipCol.iloc[i]=str(ans['To_Zip'].iloc[i])[:5]

    ans['To_Zip']=zipCol.values

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Invoice Date']=pd.to_datetime(df.iloc[:,3])
    ans['Year']=df['Invoice Date'].dt.year
    ans['Month']=df['Invoice Date'].dt.month
    ans['Day']=df['Invoice Date'].dt.day

    ans=ans.fillna('')

    return ans, row


def refrigeration(file_path,orgName,year,month):
    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls,header=None)

    df=df.drop(df.index[[0]])
    df=df.reset_index()
    df=df.drop(['index'],axis=1)
    df.columns=['Type','Model_Num','Date','Invoice_Num','Branch_Num','Qty','Sales_Price','To_Zip']

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['Model_Num']
    ans['Total_Unit_Sales']=df['Qty']
    ans['Branch_Num']=df['Branch_Num']
    ans['Invoice_Num']=df['Invoice_Num']
    ans['To_Zip']=df['To_Zip']
    ans['Product_Type']=df['Type']
    ans['Sales_Price']=df['Sales_Price']

    priceCol=[pd.np.nan]*row
    priceCol=pd.Series(priceCol)

    for i in range(0,row):
        if ans['Total_Unit_Sales'].iloc[i]==0:
            priceCol.iloc[i]=0
        else:
            priceCol.iloc[i]=ans['Sales_Price'].iloc[i]/ans['Total_Unit_Sales'].iloc[i]

    ans['Sales_Price']=priceCol.values

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    ans['Year']=df['Date'].dt.year
    ans['Month']=df['Date'].dt.month
    ans['Day']=df['Date'].dt.day

    ans=ans.fillna('')

    return ans, row


def robertmadden(file_path,orgName,year,month):

    xls=pd.ExcelFile(file_path)
    global df
    for i in range(0,len(xls.sheet_names)):
        if i==0:
            df=pd.read_excel(xls,sheet_name=i,skiprows=2)
        else:
            temp=pd.read_excel(xls,sheet_name=i,skiprows=2)
            df=pd.concat([df,temp])

    df=df.fillna('')
    df=df[df['Item ID']!='']
    df=df.reset_index()
    df=df.drop(['index'],axis=1)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['Item ID']
    ans['Total_Unit_Sales']=df['Qty Shpd']
    ans['Branch_Num']=df['Sales Location ID']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Invoice Date']=pd.to_datetime(df['Invoice Date'])
    ans['Year']=df['Invoice Date'].dt.year
    ans['Month']=df['Invoice Date'].dt.month
    ans['Day']=df['Invoice Date'].dt.day

    ans=ans.fillna('')

    return ans, row


def robertson(file_path,orgName,year,month):
    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls,skiprows=2)
    row,ans=createDF(df,orgName)

    ans['Branch_Num']=df['Branch #']
    ans['Model_Num']=df['Model #']
    ans['Total_Unit_Sales']=df['Quantity Sold']
    ans['To_Zip']=df['Delivery Zip Code']
    ans['Sales_Price']=df['Sale Price']
    ans['Invoice_Num']=df['Invoice #']

    #dates
    df['Sale Date']=pd.to_datetime(df['Sale Date'])
    ans['Year']=df['Sale Date'].dt.year
    ans['Month']=df['Sale Date'].dt.month
    ans['Day']=df['Sale Date'].dt.day

    ans=ans.fillna('')

    return ans, row


def shearer(file_path,orgName,year,month):

    #what if submitted in csv format?

    #load the file to a data frame
    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['PROD']
    ans['Total_Unit_Sales']=df['QTY.BILLED']
    ans['Branch_Num']=df['WHSE']
    ans['Invoice_Num']=df['Invoice#']
    ans['Sales_Price']=df['Unit Cost']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Inv-Date']=pd.to_datetime(df['Inv-Date'])
    ans['Year']=df['Inv-Date'].dt.year
    ans['Month']=df['Inv-Date'].dt.month
    ans['Day']=df['Inv-Date'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def sidharvey(file_path,orgName,year,month):

    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls)

    row, ans = createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['Sid Item']
    ans['Total_Unit_Sales']=df['Qty']
    ans['Branch_Num']=df['Loc.']
    ans['Product_Type']=df['MFG Item']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    day=1
    dayCol=[day]*row
    dayCol=pd.Series(dayCol)
    ans['Day']=dayCol.values

    monthCol=[month]*row
    monthCol=pd.Series(monthCol)
    ans['Month']=monthCol.values

    yearCol=[year]*row
    yearCol=pd.Series(yearCol)
    ans['Year']=yearCol.values

    #get rid of nans
    ans=ans.fillna('')

    return ans, row


def sigler(file_path,orgName,year,month):

    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls)
    df=df.drop(df.columns[[8,9,10,11,12,13]],axis=1)
    df.columns=['Invoice_Num','Date','Unknown1','Model_Num','Type','Qty','Branch_Num','Zip']

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['Model_Num']
    ans['Total_Unit_Sales']=df['Qty']
    ans['Branch_Num']=df['Branch_Num']
    ans['Invoice_Num']=df['Invoice_Num']

    zipCol=[pd.np.nan]*row
    zipCol=pd.Series(zipCol)

    for i in range(0,row):
        zipCol.iloc[i]=str(df['Zip'].iloc[i])[:5]

    ans['To_Zip']=zipCol.values


    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    ans['Year']=df['Date'].dt.year
    ans['Month']=df['Date'].dt.month
    ans['Day']=df['Date'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans, row


def standard(file_path,orgName,year,month):

    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls,sheet_name="Sales Fields",skiprows=4)

    df1=df.iloc[:,0:9]

    df2=df.iloc[:,10:19]
    df2.columns=['Invoice Number','Sale Date','Brand/ Manufacturer','Model Number','AHRI Number','Number of Units','Sale Price','Location Number','Delivery ZIP Code']

    df2=df2.fillna('')
    df2=df2[df2['Invoice Number']!='']
    df2=df2.reset_index()
    df2=df2.drop(['index'],axis=1)

    df3=df.iloc[:,20:29]
    df3.columns=['Invoice Number','Sale Date','Brand/ Manufacturer','Model Number','AHRI Number','Number of Units','Sale Price','Location Number','Delivery ZIP Code']

    df3=df3.fillna('')
    df3=df3[df3['Invoice Number']!='']
    df3=df3.reset_index()
    df3=df3.drop(['index'],axis=1)

    df=df1.append(df2,ignore_index=True,sort=False)
    df=df.append(df3,ignore_index=True,sort=False)


    row,ans=createDF(df,orgName)

    ans['Model_Num']=df['Model Number']
    ans['Branch_Num']=df['Location Number']
    ans['Total_Unit_Sales']=df['Number of Units']
    ans['Sales_Price']=df['Sale Price']
    ans['Invoice_Num']=df['Invoice Number']
    ans['To_Zip']=df['Delivery ZIP Code']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Sale Date']=pd.to_datetime(df['Sale Date'])
    ans['Year']=df['Sale Date'].dt.year
    ans['Month']=df['Sale Date'].dt.month
    ans['Day']=df['Sale Date'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def starSupply(file_path,orgName,year,month):

    #load the file to a data frame
    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['Model Number']
    ans['Total_Unit_Sales']=df['Number of Units']
    ans['Branch_Num']=df['Location Number']
    ans['Invoice_Num']=df['Invoice Number']
    ans['To_Zip']=df['Delivery ZIP Code']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Sale Date']=pd.to_datetime(df['Sale Date'])
    ans['Year']=df['Sale Date'].dt.year
    ans['Month']=df['Sale Date'].dt.month
    ans['Day']=df['Sale Date'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def southcentral(file_path,orgName,year,month):

    df=pd.read_csv(file_path)
    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['Model Number']
    ans['Total_Unit_Sales']=df['Number of Units']
    ans['Branch_Num']=df['Location Number']
    ans['Invoice_Num']=df['Invoice Number']
    ans['To_Zip']=df['Delivery ZIP Code']
    ans['Sales_Price']=df['Sale Price']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Sale Date']=pd.to_datetime(df['Sale Date'])
    ans['Year']=df['Sale Date'].dt.year
    ans['Month']=df['Sale Date'].dt.month
    ans['Day']=df['Sale Date'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans, row


def teamair(file_path,orgName,year,month):

    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls)

    row, ans = createDF(df, orgName)

    #Copy other columns
    ans['Model_Num']=df['Item Num']
    ans['Total_Unit_Sales']=df['Qty Ship']
    ans['Branch_Num']=df['Ship Whse-LN']
    ans['Invoice_Num']=df['Invoice']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    ans['Year']=df['Invoice Dt'].dt.year
    ans['Month']=df['Invoice Dt'].dt.month
    ans['Day']=df['Invoice Dt'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans, row


def temperature(file_path,orgName,year,month):

    #load the file to a data frame
    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls)

    df.columns.values[0]='Whse'
    df.columns.values[1]='Invoice Date'
    df.columns.values[2]='Product'
    df.columns.values[3]='Ship Qty'
    df.columns.values[4]='Zip'
    df.columns.values[5]='Sell Price'
    df.columns.values[6]='Order'

    df=df.fillna('')
    df=df[df['Order']!='']
    df=df.reset_index()
    df=df.drop(['index'],axis=1)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['Product']
    ans['Total_Unit_Sales']=df['Ship Qty']
    ans['Branch_Num']=df['Whse']
    ans['Invoice_Num']=df['Order']
    ans['Sales_Price']=df['Sell Price']
    ans['To_Zip']=df['Zip']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Invoice Date']=pd.to_datetime(df['Invoice Date'])
    ans['Year']=df['Invoice Date'].dt.year
    ans['Month']=df['Invoice Date'].dt.month
    ans['Day']=df['Invoice Date'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def totalAir(file_path,orgName,year,month):
    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls)
    df=df.fillna('')

    row,ans=createDF(df,orgName)

    ans['Model_Num']=df['Model']
    ans['Invoice_Num']=df['Invoice Number']

    branchCol=[1]*row
    branchCol=pd.Series(branchCol)
    ans['Branch_Num']=branchCol.values

    df['Number of Units'].fillna(1)
    ans['Total_Unit_Sales']=df['Number of Units']
    ans['To_Zip']=df['Delivery ZIP Code']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    df['Sale Date']=pd.to_datetime(df['Sale Date'],format='%Y%m%d')
    ans['Year']=df['Sale Date'].dt.year
    ans['Month']=df['Sale Date'].dt.month
    ans['Day']=df['Sale Date'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def granite(file_path,orgName,year,month):

    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls,sheet_name='Data')

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['Code (Product)']
    ans['Total_Unit_Sales']=df['Quantity']
    ans['Branch_Num']=df['Code (Warehouse)']
    ans['Invoice_Num']=df['Order Number']
    ans['Product_Type']=df['Description (Product)']
    ans['To_Zip']=df['Zip (Customer Shipto)']
    ans['Sales_Price']=df['Value']/df['Quantity']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    ans['Year']=df['Date'].dt.year
    ans['Month']=df['Date'].dt.month
    ans['Day']=df['Date'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans, row


def thossomerville(file_path,orgName,year,month):

    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls)
    row,ans=createDF(df,orgName)
    df=df.fillna('')

    #Copy other columns
    ans['Total_Unit_Sales']=df['Qty']
    ans['Branch_Num']=df['LocNum']
    ans['Invoice_Num']=df['InvoiceNum']
    ans['Product_Type']=df['Description']
    ans['To_Zip']=df['DelivZip']

    #Copy other columns
    ModelCol=[pd.np.nan]*row
    ModelCol=pd.Series(ModelCol)

    for i in range(0,row):

        if df['Model'].iloc[i]=='':
            ModelCol.iloc[i]=df['shipprod'].iloc[i]
        else:
            ModelCol.iloc[i]=df['Model'].iloc[i]
    ans['Model_Num']=ModelCol.values

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    ans['Year']=df['SaleDt'].dt.year
    ans['Month']=df['SaleDt'].dt.month
    ans['Day']=df['SaleDt'].dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans, row


def williams(file_path,orgName,year,month):

    ##load the file to a data frame
    xls=pd.ExcelFile(file_path)
    df=pd.read_excel(xls,skiprows=4)
    df.columns.values[6]='Type'

    df=df.fillna('')
    df=df[df['Item']!='']
    df=df.reset_index()
    df=df.drop(['index'],axis=1)

    row,ans=createDF(df,orgName)

    #Copy other columns
    ans['Model_Num']=df['Item']
    ans['Total_Unit_Sales']=df['UNITS']
    ans['Branch_Num']=df['STATE']
    ans['To_Zip']=df['ZIP']
    ans['Product_Type']=df['Type']

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    day=1
    dayCol=[day]*row
    dayCol=pd.Series(dayCol)
    ans['Day']=dayCol.values

    monthCol=[month]*row
    monthCol=pd.Series(monthCol)
    ans['Month']=monthCol.values

    yearCol=[year]*row
    yearCol=pd.Series(yearCol)
    ans['Year']=yearCol.values

    #get rid of nans
    ans=ans.fillna('')

    return ans,row


def winsupply(file_path,orgName,year,month):

    df=pd.read_csv(file_path,quotechar='"')

    row,ans=createDF(df,orgName)

    temp=df.iloc[:,0]
    temp=temp.str.replace('\t','')
    ans['Model_Num']=temp
    ans['Model_Num']=ans['Model_Num'].str[3:]

    temp=df.iloc[:,4]
    temp=temp.str.replace('\t','')
    ans['From_Zip']=temp

    temp=df.iloc[:,6]
    temp=temp.str.replace('\t','')
    ans['To_Zip']=temp
    ans['To_Zip']=ans['To_Zip'].str[:5]

    ans['Total_Unit_Sales']=df.iloc[:,2]
    ans['Branch_Num']=df.iloc[:,3]
    ans['Invoice_Num']=df.iloc[:,7]
    ans['Sales_Price']=df.iloc[:,10]

    #quarter
    quarter=ceil(month/3)
    quarterCol=[quarter]*row
    quarterCol=pd.Series(quarterCol)
    ans['Quarter']=quarterCol.values

    #dates
    dateCol=df.iloc[:,1]
    dateCol=dateCol.str.replace('\t','')
    dateCol=pd.to_datetime(dateCol)

    ans['Year']=dateCol.dt.year
    ans['Month']=dateCol.dt.month
    ans['Day']=dateCol.dt.day

    #get rid of nans
    ans=ans.fillna('')

    return ans,row
