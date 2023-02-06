import pandas as pd
import sqlalchemy

#todo: parser then table creation

excel = pd.read_excel(r'SpaceNK_2.0.xlsx', sheet_name=0)
print(excel)
row_list = excel.loc[4, :].values.flatten().tolist()

excel.columns = row_list
excel.drop(excel.columns[[0, 1, 4]], axis=1, inplace=True)
excel.drop(excel.index[[0, 1, 2, 3, 4, 36]], axis=0, inplace=True)
excel.reset_index(drop=True, inplace=True)

print(excel.to_string())

engine = sqlalchemy.create_engine('postgresql+psycopg2://postgres:admin@localhost/ExcelToPostgre')
print(engine)
print(sqlalchemy.inspect(engine).has_table("Stores"))
excel.to_sql('Stores', engine, if_exists='replace', index=False)