import pandas as pd

def create_df(data: list) -> pd.DataFrame:
    if not data:
        return pd.DataFrame({})
    
    return pd.DataFrame(data[1:], columns = data[0])