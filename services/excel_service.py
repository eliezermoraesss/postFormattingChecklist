import pandas as pd
import os

EXCEL_PATH = "/mnt/data/e56b7408-0ce9-49c5-bcf4-87e2a0f943e3.xlsx"

def load_data():
    df = pd.read_excel(EXCEL_PATH)
    df.columns = ["item", "checked", "programa", "descricao"]
    df["checked"] = df["checked"].apply(lambda x: True if str(x).strip() not in ["", "", "False"] else False)
    return df

def save_data(df):
    df.to_excel(EXCEL_PATH, index=False)

def add_row(descricao, programa):
    df = load_data()
    new_row = {
        "item": len(df) + 1,
        "checked": False,
        "programa": programa,
        "descricao": descricao
    }
    df = df.append(new_row, ignore_index=True)
    save_data(df)

def update_row(index, descricao, programa, checked):
    df = load_data()
    df.at[index, "descricao"] = descricao
    df.at[index, "programa"] = programa
    df.at[index, "checked"] = checked
    save_data(df)

def delete_row(index):
    df = load_data()
    df = df.drop(index)
    df.reset_index(drop=True, inplace=True)
    save_data(df)
