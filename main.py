from ExcelToPostgre import ExcelToPostgre

def main():
    excelToPostgre = ExcelToPostgre('SpaceNK_2.0.xlsx',
                                    'Stores',
                                    'postgresql+psycopg2://postgres:admin@localhost/ExcelToPostgre')
    excelToPostgre.extractAndLoad()

if __name__ == "__main__":
    main()
