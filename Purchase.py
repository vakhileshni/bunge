import pandas as pd
import os
parent_dir=os.getcwd()
com=pd.read_excel(parent_dir+"\\Total_Rollout_Status.xlsx",sheet_name='Complete Only')
com=com[['Distributor SAP Code']]
Sap=pd.read_excel(parent_dir+"\\SAPT.xlsx")
Com=pd.merge(Sap,com,left_on='Sold To Party',right_on='Distributor SAP Code',how='left')
Com['Distributor SAP Code']=Com['Distributor SAP Code'].fillna('OK')
Com=Com[Com['Distributor SAP Code']!='OK']
Com['Billing Date']=pd.to_datetime(Com['Billing Date'])
Com['Month-Year']=Com['Billing Date'].dt.to_period('M')
Pur=pd.read_excel(parent_dir+"\\PURCHASE.xlsx")
Pur=pd.merge(Pur,com,left_on='Distributor Code',right_on='Distributor SAP Code',how='left')
Pur['Distributor SAP Code']=Pur['Distributor SAP Code'].fillna('OK')
Pur=Pur[Pur['Distributor SAP Code']!='OK']
Com=Com[['Month-Year','Billing Date','Sold To Party','Sold To Party Name','Material','Material Desc','Billing Quantity']]
Com['SAP/DMS']="SAP"
Pur['Billing Date']=pd.to_datetime(Pur['Voucher Date'])
Pur['Month-Year']=Pur['Billing Date'].dt.to_period('M')
Pur=Pur[['Month-Year','Billing Date','Distributor Code','Distributor Name','Item Code','Item Description','Purchase Qty']]
Com=Com.rename(columns={'Sold To Party':'Distributor Code','Sold To Party Name':'Distributor Name'})
Pur=Pur.rename(columns={'Item Code':'Material','Item Description':'Material Desc','Purchase Qty':'Billing Quantity'})
Pur['SAP/DMS']='DMS purchage'
Comm=Com.append(Pur)
#Sale=pd.read_excel(parent_dir+"\\Sale.xlsx")
#Sale=pd.merge(Sale,com,left_on='Distributor Code',right_on='Distributor SAP Code',how='left')
#Sale['Distributor SAP Code']=Sale['Distributor SAP Code'].fillna('OK')
#Sale=Sale[Sale['Distributor SAP Code']!='OK']
#Sale['Billing Date']=pd.to_datetime(Sale['Voucher Date'])
#Sale['Month-Year']=Sale['Billing Date'].dt.to_period('M')
#Sale=Sale[['Month-Year','Billing Date','Distributor Code','Distributor Name','Item Code','Item Description','Sale Qty']]
#Sale=Sale.rename(columns={'Item Code':'Material','Item Description':'Material Desc','Sale Qty':'Billing Quantity'})
#Sale['SAP/DMS']='DMS Sale'
#Comm=Comm.append(Sale)
city=pd.read_excel(r"parent_dir+"\\City Master.xlsx")
Comm=pd.merge(Comm,city,left_on='Distributor Code',right_on='Ship To Party',how='left')
ASM=Comm[['ASM']]
ASM=ASM.drop_duplicates(subset='SM')
for ind in ASM.index:
    a=ASM['ASM'][ind]
    Ak=Comm[Comm['ASM']==a]
    path=parent_dir+"\\files\\"+a
    os.mkdir(path)
    Ak.to_excel(path+"\\"+a+"Dashboard.xlsx",index=False)

