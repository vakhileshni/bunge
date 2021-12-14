import pandas as pd
#df=pd.read_excel('final report.xlsx')
df=pd.read_excel(r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\EXPORT.XLSX")
df=df[df['Billing Doc']!='']
df1=df[df['Billing Doc Type'].isin(['Invoice','IND Credit fr Return'])]
df1=df1[['Billing Doc','Billing Doc Type','Cancelled Billing Doc','Billing Date','Sold To Party','Sold To Party Name','Sold To Party City','Sold To Party State','Billing Quantity','Billing Quantity (MT)','Billing Qty Unit','Material','Material Desc','Price (INR)','Material Grp5 Desc','Ship To Party']]
df1['Material Grp5 Desc']=df1['Material Grp5 Desc'].fillna(0)
df2=df1[df1['Material Grp5 Desc']!=0]
df3=df1[df1['Material Grp5 Desc']==0]
df3['Material Grp5 Desc']=df3['Material Desc'].replace({'GAGAN FRIDGE BOTTLE':'Promo Item','OIL COUPON':'Promo Item','VANASPATI COUPON':'Promo Item'})
df3=df3[df3['Material Grp5 Desc']=='Promo Item']
dff=pd.concat([df2,df3],axis=0)
dff.to_excel('final.xlsx',index=False)
