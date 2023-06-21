"""Script for filtering rows with invalid emails."""

import pandas as pd
from re import search

data = pd.read_excel("./data/Ckodon Bio Submission Form (Responses).xlsx")
output = pd.DataFrame(columns=data.columns)

# add rows with invalid emails to output dataframe
for row in data.index:
    if not search("\S+@\S+", data.loc[row]["Email"]):
        output.loc[len(output)] = data.loc[row]
        data.drop(row, inplace=True)

# export invalid emails to Excel
output.to_excel("./data/Ckodon Bio Submission Form - NOEMAIL.xlsx", index=False)
