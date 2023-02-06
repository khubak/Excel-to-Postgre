import pandas as pd
from sqlalchemy.orm import declarative_base
from sqlalchemy import Table, Column, String, MetaData, VARCHAR, DECIMAL, BIGINT, FLOAT, create_engine


class ExcelToPostgre:
    df = None
    engine = None
    metadata = None
    base = None

    def __init__(self, filename, tablename, engineURL):
        self.filename = filename
        self.tablename = tablename
        self.engineURL = engineURL

    def __str__(self):
        return f"{self.filename} ({self.engineURL})"

    # main class function
    def extractAndLoad(self):
        self.__extractAndCleanup()
        self.engine = create_engine(self.engineURL)
        self.__drop_table()

        # checks if table 'Stores' exists
        # print(sqlalchemy.inspect(self.engine).has_table("Stores"))

        self.__createTable()
        self.__insertDataIntoDB()

    # deletes a table with the same name if it exists
    def __drop_table(self):
        self.base = declarative_base()
        self.metadata = MetaData()
        self.metadata.reflect(bind=self.engine)
        table = self.metadata.tables.get(self.tablename)
        if table is not None:
            self.base.metadata.drop_all(self.engine, [table], checkfirst=True)

    # extracts data from .xlsx file into a Pandas DataFrame and deletes excess rows and columns
    def __extractAndCleanup(self):
        # read .xlsx
        self.df = pd.read_excel(r'SpaceNK_2.0.xlsx', sheet_name=0)
        print(f"\n\nOriginal DataFrame:\n{self.df}")

        # cleanup DataFrame
        row_list = self.df.loc[4, :].values.flatten().tolist()
        self.df.columns = row_list
        self.df.drop(self.df.columns[[0, 1, 4]], axis=1, inplace=True)
        self.df.drop(self.df.index[[0, 1, 2, 3, 4, len(self.df.index) - 1]], axis=0, inplace=True)
        self.df.reset_index(drop=True, inplace=True)
        print(f"\n\nCleaned up DataFrame\n{self.df.to_string()}")

    # creates a table of appropriate structure for 'Stores' sheet
    def __createTable(self):
        self.metadata.clear()

        Stores = Table(
            'Stores', self.metadata,
            Column('Store No', VARCHAR, primary_key=True),
            Column('Store', String),
            Column('TY Units', BIGINT),
            Column('LY Units', BIGINT),
            Column('TW Sales', FLOAT),
            Column('LW Sales', FLOAT),
            Column('LW Var %', DECIMAL(5, 2)),
            Column('LY Sales', FLOAT),
            Column('LY Var %', DECIMAL(5, 2)),
            Column('YTD Sales', FLOAT),
            Column('LYTD Sales', FLOAT),
            Column('LYTD Var %', DECIMAL(5, 2)),
        )

        self.metadata.create_all(self.engine)

    # inserts data from DataFrame into PostgreSQL DB
    def __insertDataIntoDB(self):
        self.df.to_sql('Stores', self.engine, if_exists='append', index=False)
        print("\nData has been successfully inserted into DB!")