import pandas as pd
from tkinter import *
from tkinter import filedialog

root = Tk()
root.withdraw()

print('Open service file')
df1 = pd.read_excel(filedialog.askopenfilename())
grouped = df1.groupby('ClientID')['BeginDate'].last()

print('Open Prioritization List')
plist = (filedialog.askopenfilename())

root.quit()

df2 = pd.read_excel(plist)

final = pd.merge(grouped.to_frame(), df2, left_on='ClientID', right_on='Custom_VW_PrioritizationList_ClientID',
                 how='right')
final.rename(columns={final.columns[0]: "Last Service Date"}, inplace=True)
final['Last Service Date'].fillna('No service in last 90 days', inplace=True)

writer = pd.ExcelWriter('Prioritizaion last service in 90 days.xlsx')
final.to_excel(writer, 'List with Last Service Date')
writer.save()
