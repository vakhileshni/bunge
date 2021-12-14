import pandas as pd
Sync=pd.read_excel(r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\DBMS mapping darta\Daily Rollout Working files\Sync.xlsx")
Sync=Sync.drop_duplicates(subset=['Distributor Code'])
Sync['Check']=Sync['Distributor Code'].astype(str)
Sync=Sync[Sync['Check'].str.len()>7]
Sync=Sync[['Distributor Name','Distributor Code','Software','Sync Due Days','Mode','Status']]
Sync['Ageing']=''
Sync.loc[Sync['Sync Due Days'] > 7, 'Ageing'] = '>7'
Sync.loc[(Sync['Sync Due Days']<=7) & (Sync['Sync Due Days']>=5), 'Ageing'] = '5-7'
Sync.loc[(Sync['Sync Due Days']<=4) & (Sync['Sync Due Days']>=3), 'Ageing'] = '3-5'
Sync.loc[(Sync['Sync Due Days']<=2) & (Sync['Sync Due Days']>=0), 'Ageing'] = '0-2'
Tally=pd.read_excel(r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\DBMS mapping darta\Daily Rollout Working files\Tally.xlsx")
Tally=Tally.drop_duplicates(subset=['Distributors Code'])
Tally=Tally[['Distributors Code']]
Tally=Tally.rename(columns={'Distributors Code':'New Tally'})
Sync=pd.merge(Sync,Tally,left_on='Distributor Code',right_on='New Tally',how='left')
Sync.loc[Sync['New Tally']>0,'New Tally']='Y'
Sync['New Tally']=Sync['New Tally'].fillna('N')
City=pd.read_excel(r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\DBMS mapping darta\Daily Rollout Working files\City Master.xlsx")
Sync=pd.merge(Sync,City,how='left',left_on='Distributor Code',right_on='Ship To Party')
Sync=Sync[['Distributor Name','Distributor Code','Software','Sync Due Days','Mode','Status','Ageing','New Tally','Biz','Zone','SM','RBM']]
Sync.to_excel('Sync.xlsx',index=False)
