import pandas as pd
from sqlalchemy import create_engine

#reads excel file and uses its data to update an sql database
df = pd.read_excel(r"C:\Users\madis\Downloads\abcd_excel_file.xlsx")
df = df.set_index("ID")
df = df.rename(columns={"State Name":"state_name", "Key Words":"key_words"})
engine = create_engine("mysql://root@localhost/abcd_db")
df.to_sql("dresses", con = engine, if_exists="replace")
print("Database has been updated.")