# with open('invest.csv', 'r') as f:
#     lines = f.readlines()
#     lines = lines[:-1]
#
# with open('invest.csv', 'w') as f:
#     f.writelines(lines)
#
import pandas as pd

data = pd.read_excel('rez_file_Y_v2.xlsx', header=0)[:-40]
data.to_excel('rez_file_Y_v2.xlsx', index=False)
