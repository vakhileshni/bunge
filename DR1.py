import pandas as pd
import os, shutil
parent_dir=os.getcwd()
directory = "\\BR\\DR_PUR"
try:    
    path1= parent_dir+directory 
    shutil.rmtree(path1)
    os.mkdir(path1)
except:
    path1=parent_dir+directory
    os.mkdir(path1)
parent_dir1=parent_dir+directory
df=pd.read_excel(parent_dir+"\\Plant details.xlsx")
DR=pd.read_excel(parent_dir+"\\DR Report.xlsx")
DR2=pd.read_excel(parent_dir+"\\Purchase Report.xlsx")
#FR=pd.read_excel(parent_dir+"\\FREIGHT.xlsx")
STN=pd.read_excel(parent_dir+"\\STN.xlsx")
for ind in df.index:
    a=df['Plant Code'][ind]
    b=a
    b=b.astype(str)
    DR1=DR[DR['PLANT NAME']==a]
    DR3=DR2[DR2['Plant']==a]
    STN1=STN[STN['Plant']==a]
    path = os.path.join(parent_dir1,b)
    os.mkdir(path)
    C=pd.Series([DR1['PLANT NAME']]).sum()
    D=C.sum()
    E=pd.Series([DR3['Plant']]).sum()
    F=E.sum()
    G=pd.Series([STN1['Plant']]).sum()
    H=G.sum()
    if D > 0:
        DR1.to_excel(path+"\\"+b+" Sale.xlsx",index=False)
        #FR.to_excel(path+"\\"+b+" Freight.xlsx",index=False)
        shutil.copy(parent_dir+"\\Freight.xlsb", path+"\\"+b+" Freight.xlsb" )
    if F > 0:
        DR3.to_excel(path+"\\"+b+" Purchase.xlsx",index=False)
    if H > 0:
        STN1.to_excel(path+"\\"+b+" STN.xlsx",index=False)        
path3=parent_dir+directory
for ind in df.index:
    a=df['Plant Code'][ind]
    b=a
    b=b.astype(str)
    if len(os.listdir(path3+"\\"+b))==0:
        shutil.rmtree(path3+"\\"+b)
files = os.listdir(path3)
list=[]
for name in files:
    list.append(name)
df1= pd.DataFrame(list,columns=['PLANT NO'])
df1.to_excel(parent_dir+"\\Plant_Name.xlsx",index=False)
data=pd.read_excel("Record.xlsx")
df1=pd.read_excel("Plant_Name.xlsx")
df1=pd.merge(df1,data,on='PLANT NO',how='left')
df1.to_excel(parent_dir+"\\Plant_Name.xlsx",index=False)
