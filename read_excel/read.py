import polars as pl
import pandas as pd


def readExcel(archivo: str):
    df_pandas = pd.read_excel(archivo)
    df_polars = pl.from_pandas(df_pandas)
    return df_polars




