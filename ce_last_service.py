import pandas as pd
from tkinter import *
from tkinter import filedialog

#  I don't understand this part
root = Tk()
root.withdraw()

print('Open service file')
df1 = pd.read_excel(filedialog.askopenfilename())
grouped = df1.groupby('ClientID')['BeginDate'].last()

print('Open Prioritization List')
plist = (filedialog.askopenfilename())

#  closes tkinter window
root.quit()

df2 = pd.read_excel(plist)

final = pd.merge(grouped.to_frame(), df2, left_on='ClientID', right_on='Custom_VW_PrioritizationList_ClientID', how='right')
final.rename(columns={final.columns[0]: "Last ServiceDate"}, inplace=True)

writer = pd.ExcelWriter('Prioritizaion last service in 90 days.xlsx')
final.to_excel(writer, 'List with Last Service Date')
writer.save()
