import pandas as pd
Rollout1=pd.read_excel(r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\DBMS mapping darta\Daily Rollout Working files\Rollout Status.xlsx")
Rollout=Rollout1[Rollout1['Status']=='Done']
Sap=pd.read_excel(r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\DBMS mapping darta\Daily Rollout Working files\SAP.xlsx")
Map=pd.read_excel(r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\DBMS mapping darta\Daily Rollout Working files\Mapped.xlsx")
Sync=pd.read_excel(r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\DBMS mapping darta\Daily Rollout Working files\Sync.xlsx")
Tally=pd.read_excel(r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\DBMS mapping darta\Daily Rollout Working files\Tally.xlsx")
Item=pd.read_excel(r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\DBMS mapping darta\Daily Rollout Working files\Item Master.xlsx")
Tally=Tally[['Distributors Code']]
Tally=Tally.drop_duplicates(subset=['Distributors Code'],keep='first')
Sap1=Sap[['Sold To Party','Sold To Party Name']]
Sap1=Sap1.drop_duplicates(subset=['Sold To Party'],keep='first')
Rollout=pd.merge(Rollout,Sap1,left_on='Code',right_on='Sold To Party',how='left')
Sync1=Sync[['Distributor Code','Distributor Name']]
Sync1=Sync1.drop_duplicates(subset=['Distributor Code'],keep='first')
Rollout=pd.merge(Rollout,Sync1,left_on='Code',right_on='Distributor Code',how='left')
Rollout=Rollout.rename(columns={'Distributor Name':'Sync Distributor Name'})
Map1=Map[['Distributor Code','Distributor Name']]
Map1=Map1.drop_duplicates(subset=['Distributor Code'],keep='first')
Rollout=pd.merge(Rollout,Map1,left_on='Code',right_on='Distributor Code',how='left')
Rollout=Rollout.rename(columns={'Sold To Party Name':'Sap Distributor Name','Distributor Name':'DMS Distributor Name'})
Rollout=Rollout[['Code','Division',"Distributor's Name",'Status','Sap Distributor Name','Sync Distributor Name','DMS Distributor Name']]
Sap2=pd.merge(Sap,Rollout,left_on='Sold To Party',right_on='Code',how='left')
Sap2=Sap2.drop(columns=['Division',"Distributor's Name",'Status','Sap Distributor Name','Sync Distributor Name','DMS Distributor Name'])
Sap2['Code']=Sap2['Code'].fillna(0)
Sap2=Sap2[Sap2['Code']!=0]
Sap2=Sap2.drop(columns=['Code'])
Map2=pd.merge(Map,Rollout,left_on='Distributor Code',right_on='Code',how='left')
Map2=Map2.drop(columns=['Division',"Distributor's Name",'Status','Sap Distributor Name','Sync Distributor Name','DMS Distributor Name'])
Map2['Code']=Map2['Code'].fillna(0)
Map2=Map2[Map2['Code']!=0]
Map2=Map2.drop(columns=['Code'])
Sap2['Code-Mat']=Sap2['Sold To Party'].astype(str)+Sap2['Material'].astype(str)
Map2['Code-Mat1']=Map2['Distributor Code'].astype(str)+Map2['DMS Item Code'].astype(str)
Sap3=Sap2.drop_duplicates(subset=['Code-Mat'],keep='first')
Map3=Map2[['Code-Mat1']]
Sap3=pd.merge(Sap3,Map3,left_on='Code-Mat',right_on='Code-Mat1',how='left')
Sap3=Sap3.drop_duplicates(subset=['Code-Mat'],keep='first')
Sap4=Sap3
Sap4['Code-Mat1']=Sap4['Code-Mat1'].fillna(0)
Sap4=Sap4[Sap4['Code-Mat1']==0]
Sap4=Sap4.drop(columns=['Code-Mat1'])
Sap4=pd.merge(Sap4,Tally,left_on='Sold To Party',right_on='Distributors Code',how='left')
Sap4=Sap4.rename(columns={'Distributors Code':'Tally'})
Sap4.loc[Sap4['Tally']>0,'Tally']='Y'
Sap4['Tally']=Sap4['Tally'].fillna('N')
Item=Item[['Item Code']]
Sap5=pd.merge(Sap4,Item,how='left',left_on='Material',right_on='Item Code')
Sap5.loc[Sap5['Item Code']>0,'Item Code']='Y'
Sap5['Item Code']=Sap5['Item Code'].fillna('N')
Sap5=Sap5.rename(columns={'Item Code':'Availability In DMS'})
Sap5=Sap5.pivot_table(index=['Tally','Sold To Party','Sold To Party Name','Material','Material Desc','Availability In DMS'],values=['Billing Quantity'])
Sap5=Sap5.drop(columns=['Billing Quantity'],axis=1)
Rollout=pd.merge(Rollout,Tally,left_on='Code',right_on='Distributors Code',how='left')
Rollout=Rollout.rename(columns={'Distributors Code':'Tally'})
Rollout.loc[Rollout['Tally']>0,'Tally']='Y'
Rollout['Tally']=Rollout['Tally'].fillna('N')
#Rollout['Tally']=Rollout[Rollout['Tally']!='N'=='Y']
dell=Sap4[['Sold To Party','Material']]
dell=dell.groupby(['Sold To Party']).count()
dell=dell.reset_index()
dell=dell.rename(columns={'Sold To Party':'Code'})
Rollout=pd.merge(Rollout,dell,left_on='Code',right_on='Code',how='left')
Rollout=Rollout.rename(columns={'Material':'Sap sale Sku code check with DMS mapped items'})
Rollout['Sap sale Sku code check with DMS mapped items']=Rollout['Sap sale Sku code check with DMS mapped items'].fillna('OK')
Rollout=Rollout.drop_duplicates(subset=['Code'],keep='first')
count=Sap4[['Sold To Party','Code-Mat']]
count1=count.groupby(['Sold To Party']).count()
count2=Sap3[['Sold To Party','Sold To Party Name','Material Desc']]
count3=count2.groupby(['Sold To Party','Sold To Party Name']).count()
count3=count3.reset_index()
count1=count1.reset_index()
count4=pd.merge(count3,count1,how='left',left_on='Sold To Party',right_on='Sold To Party')
count4['Code-Mat']=count4['Code-Mat'].fillna(0)
count4['Code-Mat']=count4['Code-Mat'].astype(int)
count4['Mapped SKUs']=count4['Material Desc']-count4['Code-Mat']
count4=count4.rename(columns={'Material Desc':'Total SKUs','Code-Mat':'Unmapped SKUs'})
count4=count4[['Sold To Party','Sold To Party Name','Total SKUs','Mapped SKUs','Unmapped SKUs']]
count4['Work done %']=count4['Mapped SKUs']/count4['Total SKUs']
count4['Work done %']=count4['Work done %']*[100]
count4['Work done %']=count4['Work done %'].astype(int)
count4['Work done %']=count4['Work done %'].astype(str)+'%'
write=pd.ExcelWriter('Rollout_Status.xlsx',engine='xlsxwriter')
Rollout1.to_excel(write,sheet_name='Rollout data',index=False)
Rollout.to_excel(write,sheet_name='Done Only',index=False)
#Sap2.to_excel(write,sheet_name='Sales data',index=False)
Sap3.to_excel(write,sheet_name='Sales Unique data',index=False)
#Map2.to_excel(write,sheet_name='DMS data',index=False)
Sap4.to_excel(write,sheet_name='Pending Item code',index=False)
Sap5.to_excel(write,sheet_name='Table')
count4.to_excel(write,sheet_name='count',index=False)
write.save()
