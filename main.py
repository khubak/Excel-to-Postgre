from ExcelToPostgre import ExcelToPostgre

excelToPostgre = ExcelToPostgre('SpaceNK_2.0.xlsx',
                                'Stores',
                                'postgresql+psycopg2://postgres:admin@localhost/ExcelToPostgre')
excelToPostgre.extractAndLoad()