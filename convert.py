import pandas as pd


# Store the current file into a dataframe
df_old = pd.read_excel('JiraExample.xlsx')

#Initialzie a new dataframe to store the qualified row
df_new = pd.DataFrame()

for index, row in df_old.iterrows():
    print(index)
"""
# iterate through the old dataframe
for index, row in df_old.iterrows():
    print(index)
    if(row["Ticket Type"] == "Deploy"):
        print("yes")
        df_new.append(row)

# write the new dataframe into a excel file
df_new.to_excel("readyToImport.xlsx")
"""
    
