#uses pandas to merge 2 csv files as 1 xlsx file. 
import sys
import pandas as pd
from styleframe import StyleFrame, Styler, utils

file1 = sys.argv[1]
file2 = sys.argv[2]

df1 = pd.read_csv(file1)
df2 = pd.read_csv(file2)
                                        
xlwriter = pd.ExcelWriter('output.xlsx')

df1.to_excel(xlwriter, sheet_name = 'Sheet 1', index = False)
df2.to_excel(xlwriter, sheet_name = 'Sheet 2', index = False)

xlwriter.close()
print(".csv -> .xlsx completed")