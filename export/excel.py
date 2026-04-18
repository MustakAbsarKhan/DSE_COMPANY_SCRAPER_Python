import pandas as pd


def save_to_excel(data):
    df = pd.DataFrame(data)
    df.to_excel("DSE_Company_Details.xlsx", index=False)