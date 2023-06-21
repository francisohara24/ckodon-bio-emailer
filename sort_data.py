# Script for filtering rows with invalid emails.

import pandas as pd
from re import search

data = pd.read_excel("./data/Ckodon Bio Submission Form (Responses).xlsx")

for row in data.index:
    # check for empty email values
    if data.loc[row].isna()["Email"] == True:
        print(data.loc[row])

    # check for invalid email values
    elif not(search("\S+@\S+", data.loc[row]["Email"])):
        print(data.loc[row])
