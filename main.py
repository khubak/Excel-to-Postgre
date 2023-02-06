import pandas as pd
import numpy as np
import plotly.express as px
import os
import sqlalchemy
import psycopg2
from psycopg2.extensions import ISOLATION_LEVEL_AUTOCOMMIT

#todo: parser then db creation

excel = pd.read_excel(r'SpaceNK_2.0.xlsx', sheet_name=0)
print(excel)
row_list = excel.loc[4, :].values.flatten().tolist()

excel.columns = row_list
excel.drop(excel.columns[[0, 1, 4]], axis=1, inplace=True)
excel.drop(excel.index[[0, 1, 2, 3, 4]], axis=0, inplace=True)
excel.reset_index(drop=True, inplace=True)

print(excel.to_string())

# connect to db
pgconn = psycopg2.connect(
    host='localhost',
    user='postgres',
    password='admin',
    database='ExcelToPostgre')

# cursor
pgcursor = pgconn.cursor()
pgconn.set_isolation_level(ISOLATION_LEVEL_AUTOCOMMIT)

# drop table
pgcursor.execute('DROP TABLE IF EXISTS crypto_table')

# create table
pgcursor.execute("""

CREATE TABLE IF NOT EXISTS crypto_table
(
    [Store No]               VAR_CHAR
  , Store             VARCHAR(50) NOT NULL
  , [TY Units]           VARCHAR(20) NOT NULL
  , [LY Units]        FLOAT
  , [TW Sales]   BIGINT
  , [LW Sales]   BIGINT
  , [LW Var %]   BIGINT
  , [LY Sales]   BIGINT
  , [LY Var %]   BIGINT
  , [YTD Sales]   BIGINT
  , [LYTD Sales]   BIGINT
  , [LYTD Var %]   BIGINT
);

""")

# commit
pgconn.commit()

# close
pgconn.close()

engine = sqlalchemy.create_engine('postgresql+psycopg2://postgres:admin@localhost/')
excel.to_sql('tablename', engine, if_exists='replace', index=False)
#close the connection