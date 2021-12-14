import pandas as pd
import numpy as np
DRollout=pd.read_excel(r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\DBMS mapping darta\Daily Rollout Working files\Rollout_Status.xlsx",sheet_name='Rollout data')
DRollout['Date'] = pd.to_datetime(DRollout['Date'])
DRollout['Date']=DRollout['Date'].dt.strftime('%m/%d/%y')
Con_dms=pd.read_excel(r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\DBMS mapping darta\Daily Rollout Working files\Total_Rollout_Status.xlsx",sheet_name='Rollout data')
Con_dms1=pd.merge(DRollout,Con_dms,left_on='Code',right_on='Distributor SAP Code',how='left')
Con_dms1['Distributor SAP Code']=Con_dms1['Distributor SAP Code'].fillna(0)
Con_dms1=Con_dms1[Con_dms1['Distributor SAP Code']==0]
Con_dms1=Con_dms1[['Code','Division',"Distributor's Name",'Status','Remarks','Date']]
Con_dms1['Date1']=Con_dms1['Date']
Con_dms1=Con_dms1.rename(columns={'Code':'Distributor SAP Code','Division':'Biz (B2B or B2C)',"Distributor's Name":'Distributors Name','Status':'Current status','Date':'Installation Start','Remarks':'Daily Status','Date1':'Date Daily Status'})
Con_dms2=pd.concat([Con_dms,Con_dms1],axis=0)
ak1=DRollout[['Code','Status']]
ak1=ak1.rename(columns={'Code':'Distributor SAP Code','Status':'Current status'})
ak3=pd.merge(Con_dms2,ak1,left_on='Distributor SAP Code',right_on='Distributor SAP Code',how='left')
ak3['Current status_y']=ak3['Current status_y'].fillna(0)
ak31=ak3[ak3['Current status_y']==0]
ak32=ak3[ak3['Current status_y']!=0]
ak32=ak32.rename(columns={'Current status_y':'Current status'})
ak31=ak31.rename(columns={'Current status_x':'Current status'})
ak31=ak31[['Biz (B2B or B2C)','Distributor SAP Code','Distributors Name','Installation Start','Installation End','Current status','Date Daily Status','Daily Status']]
ak32=ak32[['Biz (B2B or B2C)','Distributor SAP Code','Distributors Name','Installation Start','Installation End','Current status','Date Daily Status','Daily Status']]
ak3=pd.concat([ak32,ak31],axis=0)
ak3['Current status']=ak3['Current status'].replace({'Done':'Complete'})
ak4=DRollout[['Code','Date']]
ak4=ak4.rename(columns={'Code':'Distributor SAP Code','Date':'Installation End'})
ak5=ak3[ak3['Current status']=='Complete']
ak6=pd.merge(ak5,ak4,left_on='Distributor SAP Code',right_on='Distributor SAP Code',how='left')
ak6['Installation End_y']=ak6['Installation End_y'].fillna(0)
ak61=ak6[ak6['Installation End_y']==0]
ak62=ak6[ak6['Installation End_y']!=0]
ak62=ak62.rename(columns={'Installation End_y':'Installation End'})
ak61=ak61.rename(columns={'Installation End_x':'Installation End'})
ak61=ak61[['Biz (B2B or B2C)','Distributor SAP Code','Distributors Name','Installation Start','Installation End','Current status']]
ak62=ak62[['Biz (B2B or B2C)','Distributor SAP Code','Distributors Name','Installation Start','Installation End','Current status']]
ak6=pd.concat([ak62,ak61],axis=0)
ak6['Date Daily Status']=""
ak6['Daily Status']=""
ak7=ak3[ak3['Current status']!='Complete']
ak8=DRollout[['Code','Date','Remarks']]
ak8=ak8.rename(columns={'Code':'Distributor SAP Code','Date':'Date Daily Status','Remarks':'Daily Status'})
ak9=pd.merge(ak7,ak8,left_on='Distributor SAP Code',right_on='Distributor SAP Code',how='left')
ak9['Date Daily Status_y']=ak9['Date Daily Status_y'].fillna(0)
ak91=ak9[ak9['Date Daily Status_y']==0]
ak92=ak9[ak9['Date Daily Status_y']!=0]
ak92=ak92.rename(columns={'Date Daily Status_y':'Date Daily Status','Daily Status_y':'Daily Status'})
ak91=ak91.rename(columns={'Date Daily Status_x':'Date Daily Status','Daily Status_x':'Daily Status'})
ak91=ak91[['Biz (B2B or B2C)','Distributor SAP Code','Distributors Name','Installation Start','Installation End','Current status','Date Daily Status','Daily Status']]
ak92=ak92[['Biz (B2B or B2C)','Distributor SAP Code','Distributors Name','Installation Start','Installation End','Current status','Date Daily Status','Daily Status']]
ak9=pd.concat([ak92,ak91],axis=0)
ak10=pd.concat([ak6,ak9],axis=0)
ak10['Installation Start'] =pd.to_datetime(ak10['Installation Start'])
ak10=ak10.rename(columns={'Distributor Name':'Distributors Name'})
ak10.sort_values(by=['Distributor SAP Code'])
Rollout=ak10[ak10['Current status']=='Complete']
Sap=pd.read_excel(r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\DBMS mapping darta\Daily Rollout Working files\SAPT.xlsx")
Map=pd.read_excel(r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\DBMS mapping darta\Daily Rollout Working files\Mapped.xlsx")
Sync=pd.read_excel(r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\DBMS mapping darta\Daily Rollout Working files\Sync.xlsx")
Tally=pd.read_excel(r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\DBMS mapping darta\Daily Rollout Working files\Tally.xlsx")
Tally=Tally[['Distributors Code']]
Tally=Tally.drop_duplicates(subset=['Distributors Code'],keep='first')
Sap1=Sap[['Sold To Party','Sold To Party Name']]
Sap1=Sap1.drop_duplicates(subset=['Sold To Party'],keep='first')
Rollout=pd.merge(Rollout,Sap1,left_on='Distributor SAP Code',right_on='Sold To Party',how='left')
Sync1=Sync[['Distributor Code','Distributor Name']]
Sync1=Sync1.drop_duplicates(subset=['Distributor Code'],keep='first')
Rollout=pd.merge(Rollout,Sync1,left_on='Distributor SAP Code',right_on='Distributor Code',how='left')
Rollout=Rollout.rename(columns={'Distributor Name':'Sync Distributor Name'})
Map1=Map[['Distributor Code','Distributor Name']]
Map1=Map1.drop_duplicates(subset=['Distributor Code'],keep='first')
Rollout=pd.merge(Rollout,Map1,left_on='Distributor SAP Code',right_on='Distributor Code',how='left')
Rollout=Rollout.rename(columns={'Sold To Party Name':'Sap Distributor Name','Distributor Name':'DMS Distributor Name'})
Rollout=Rollout[['Distributor SAP Code','Biz (B2B or B2C)',"Distributors Name",'Current status','Installation Start','Installation End','Sap Distributor Name','Sync Distributor Name','DMS Distributor Name']]
Sap2=pd.merge(Sap,Rollout,left_on='Sold To Party',right_on='Distributor SAP Code',how='left')
Sap2=Sap2.drop(columns=['Biz (B2B or B2C)',"Distributors Name",'Current status','Sap Distributor Name','Sync Distributor Name','DMS Distributor Name'])
Sap2['Distributor SAP Code']=Sap2['Distributor SAP Code'].fillna(0)
Sap2=Sap2[Sap2['Distributor SAP Code']!=0]
Sap2=Sap2.drop(columns=['Distributor SAP Code'])
Map2=pd.merge(Map,Rollout,left_on='Distributor Code',right_on='Distributor SAP Code',how='left')
Map2=Map2.drop(columns=['Biz (B2B or B2C)','Distributors Name','Current status','Sap Distributor Name','Sync Distributor Name','DMS Distributor Name'])
Map2['Distributor SAP Code']=Map2['Distributor SAP Code'].fillna(0)
Map2=Map2[Map2['Distributor SAP Code']!=0]
Map2=Map2.drop(columns=['Distributor SAP Code'])
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
Sap5=Sap4.pivot_table(index=['Tally','Sold To Party','Sold To Party Name','Material','Material Desc'],values=['Billing Quantity'])
Rollout['Distributor SAP code verify is correct per SAP']=Rollout['Sap Distributor Name']==Rollout['Sync Distributor Name']
Rollout=pd.merge(Rollout,Tally,left_on='Distributor SAP Code',right_on='Distributors Code',how='left')
Rollout=Rollout.rename(columns={'Distributors Code':'Tally'})
Rollout.loc[Rollout['Tally']>0,'Tally']='Y'
Rollout['Tally']=Rollout['Tally'].fillna('N')
dell=Sap4[['Sold To Party','Material']]
dell=dell.groupby(['Sold To Party']).count()
dell=dell.reset_index()
dell=dell.rename(columns={'Sold To Party':'Distributor SAP Code'})
Rollout=pd.merge(Rollout,dell,left_on='Distributor SAP Code',right_on='Distributor SAP Code',how='left')
Rollout=Rollout.rename(columns={'Material':'Sap sale Sku code check with DMS mapped items'})
Rollout['Sap sale Sku code check with DMS mapped items']=Rollout['Sap sale Sku code check with DMS mapped items'].fillna('OK')
Rollout=Rollout.drop_duplicates(subset=['Distributor SAP Code'],keep='first')
Sap3=Sap3.drop(columns=['Installation Start','Installation End'])
Sap4=Sap4.drop(columns=['Installation Start','Installation End'])
#table= pd.pivot_table(Rollout, values='Distributor SAP Code', index=['Current status'],columns=['Biz (B2B or B2C)'], aggfunc=np.counts())
write=pd.ExcelWriter('Total_Rollout_Status.xlsx',engine='xlsxwriter')
ak10.to_excel(write,sheet_name='Rollout data',index=False)
Rollout.to_excel(write,sheet_name='Complete Only',index=False)
#Sap2.to_excel(write,sheet_name='Sales data',index=False)
Sap3.to_excel(write,sheet_name='Sales Unique data',index=False)
#Map2.to_excel(write,sheet_name='DMS data',index=False)
Sap4.to_excel(write,sheet_name='Pending Item code',index=False)
#dell.to_excel(write,sheet_name='count')
Sap5.to_excel(write,sheet_name='Table')
#table.to_excel(write,sheet_name='summary')
write.save()
