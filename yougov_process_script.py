#path of yougov resaved file
path = 'Downloads/yougov_eclass_6222.xlsx'

#Load in brand metric columns
yougov = pd.DataFrame()
yougov['date'] = pd.read_excel(path, 
                   sheet_name=1, 
                   skiprows=[0,1,2,3,4,5,6],
                   skipfooter = 2,
                   header=0)['Unnamed: 0']

for i in range(1,10):
    yougov['col'+str(i)] = pd.read_excel(path, 
                   sheet_name=i, 
                   skiprows=[0,1,2,3,4,5,6],
                   skipfooter = 2,
                   header=0).Score


yougov['col10'] = pd.read_excel(path, 
               sheet_name=10, 
               skiprows=[0,1,2,3,4,5,6],
               skipfooter = 2,
               header=0).Attention

for i in range(11,17):
    yougov['col'+str(i)] = pd.read_excel(path, 
                   sheet_name=i, 
                   skiprows=[0,1,2,3,4,5,6],
                   skipfooter = 2,
                   header=0).Score

#Rename columns

yougov.rename(columns={'col1':'index',
                    'col2':'buzz',
                    'col3':'impression',
                    'col4':'quality',
                    'col5':'value',
                    'col6':'reputation',
                    'col7':'satisfaction',
                    'col8':'recommend',
                    'col9':'awareness',
                    'col10':'attention',
                    'col11':'adawareness',
                    'col12':'wom',
                    'col13':'consideration',
                    'col14':'purchaseintent',
                    'col15':'currentcustomer',
                    'col16':'formercustomer'}, inplace=True)

#save it as csv, should change the name
yougov.to_csv('Downloads/yougov_eclass_6.csv',header=True,index=False)