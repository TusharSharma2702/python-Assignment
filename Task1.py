import pandas as pd

#Creating the first dataframe of sheet=2
df_1=pd.read_excel("C:/Users/91992/Downloads/Python Assignment.xlsx",sheet_name=1,usecols="D:G")
df_1=df_1.iloc[9:31]          #only taking the required coloumns and rows 
df_1.index=range(len(df_1))    #redefining the indexing
df_1.columns=df_1.iloc[0]     #setting the first row as a header
df_1=df_1.drop(index=0)      


#Creating the first dataframe of sheet=3
df_2=pd.read_excel("C:/Users/91992/Downloads/Python Assignment.xlsx",sheet_name=2,usecols="C:G")
df_2=df_2.iloc[6:30]                  #only taking the required coloumns and rows 
df_2.index=range(len(df_2))           #redefining the indexing
df_2.columns=df_2.iloc[0]             #setting the first row as a header
df_2=df_2.drop(index=0)
df_2.columns=["S No","Name","User ID","total_statements","total_reasons"]


#Merging the two data frames into one
merge_df=pd.merge(left=df_1,right=df_2,on=["S No","Name","User ID"])


#Taking part of data frame for team ranking
pf_1=merge_df[["S No","Team Name","total_statements","total_reasons"]]
pf_2=pf_1["Team Name"]
pf_2=pf_2.drop_duplicates(keep="first")
t_statements=[]
t_reasons=[]
for i in pf_2:
    t_statements.append(round(pf_1[pf_1.eq(i).any(axis=1)]["total_statements"].mean(),ndigits=2))
    t_reasons.append(round(pf_1[pf_1.eq(i).any(axis=1)]["total_reasons"].mean(),ndigits=2))
pf_2.index=range(len(pf_2))
pf_3=pd.DataFrame({"Team_Name":pf_2,"Average_Statements":t_statements,"Average_Reasons":t_reasons},index=range(len(pf_2)))
pf_3.sort_values(by=["Average_Statements","Average_Reasons"],ascending=False,inplace=True)
pf_3.index=range(1,10)
pf_3.index.name="Rank"

#Taking part of data frame for indivisual ranking
merge_df.sort_values(by=["total_statements","total_reasons"],ascending=False,inplace=True)
merge_df.index=range(1,22)
merge_df.index.name="Rank"
rank_df=merge_df[["Name","User ID","total_statements","total_reasons"]]


#Using both data frames two create output sheets in a single excel file
with pd.ExcelWriter('C:/Users/91992/Downloads/output.xlsx') as writer:
    pf_3.to_excel(writer, sheet_name='Leaderboard Teamwise')
    rank_df.to_excel(writer, sheet_name='Leaderboard Individual')

