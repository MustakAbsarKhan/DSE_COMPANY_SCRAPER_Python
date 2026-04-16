import pandas as pd


def save_to_excel(data, filename="DSE_Company_Details.xlsx"):
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)