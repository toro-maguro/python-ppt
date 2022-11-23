import pandas as pd

def make_dataframe():
    list_data = [["ABC001", "A", 5000, 9000, 3000],
    ["ABC001", "B", 3000, 9000, 3000],
    ["ABC001", "C", 1000, 9000, 3000],
    ["XYZ123", "A", 7000, 7000, 2000],
    ["XXX001", "C", 1000, 1000, 1000]]
    column_name = ["ProjectID", "Type", "Price", "Cost", "Cost_lastmonth"]
    df = pd.DataFrame(data=list_data, columns=column_name)
    return(df)

