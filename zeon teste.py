import pandas as pd
from sqlalchemy import create_engine
import oracledb
conn = oracledb.connect(
    user="",
    password="",
    dsn=""
)
query= """
    SELECT CODPROD, DESCRICAO, CUSTOREP, DTULTALTCUSTOREP
    FROM PCPRODUCT
"""
df = pd.read_sql(query,conn)
df.to_excel("output.xlsx", index=False)
conn.close()