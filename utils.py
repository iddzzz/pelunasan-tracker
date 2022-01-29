import pandas as pd
import xlrd as excel
import mysql.connector as sql
import math


NULL = 'NULL'


# baca excel, return n x 3 array
# excel: kolom 0 = ID, kolom 1 = nama, kolom 7 = bdn
def read_tajnid(path=r"data.xlsx"):

    input_workbook = excel.open_workbook(path)
    input_worksheet = input_workbook.sheet_by_index(0)

    data = []
    badan_dict = {'K': 'Khm', 'L': 'Ljh',
                  'A': 'Ahr', 'T': 'Ahf',
                  'N': 'Nhr', 'G': 'Bnh',
                  'B': 'Aba'}

    for i in range(input_worksheet.nrows - 1):
        if input_worksheet.cell_value(i, 0) == '':
            continue
        ID = str(round(input_worksheet.cell_value(i, 0), None))
        nama = input_worksheet.cell_value(i, 1)
        badan = input_worksheet.cell_value(i, 7)
        badan = badan_dict[badan]
        data2 = (ID, nama, badan)
        data.append(data2)

    return data


# input data ke database, argument: nx3 array
def insert_tajnid(data):

    mydb = sql.connect(
        host="localhost",
        user="root"
    )

    mycursor = mydb.cursor()
    mycursor.execute("use yogyakarta")

    sql_formula = "INSERT INTO data VALUES (%s, %s, %s, 1, NULL, NULL)"

    for item in data:
        try:
            mycursor.execute(sql_formula, item)
        except Exception as why:
            print(why)
            continue

    mydb.commit()


# input data ke database 2
def insert_ch(data):

    mydb = sql.connect(
        host="localhost",
        user="root"
    )

    mycursor = mydb.cursor()
    mycursor.execute("use yogyakarta")

    sql_formula = "INSERT INTO data VALUES (%s, %s, %s, %s, %s, %s, %s)"

    for item in data:
        try:
            mycursor.execute(sql_formula, item)
        except Exception as why:
            print(why, "Error di", item)
            return

    mydb.commit()


# baca excel, return n x 7 array
# 1 kuitansi, 2 ID, 9 tgl transaksi, 15 mjd, 80 jh, 81 Std, 82 RSID
# in excel format tgl transaksi 0000 00 00
def read_ch(path=r"data.xlsx"):

    input_workbook = excel.open_workbook(path)
    input_worksheet = input_workbook.sheet_by_index(0)

    data = []

    for i in range(input_worksheet.nrows):
        if i == 0:
            continue
        kuitansi = str(round(input_worksheet.cell_value(i, 1), None))
        ID = str(round(input_worksheet.cell_value(i, 2), None))
        tgl = input_worksheet.cell_value(i, 9)

        tgl = tgl.replace(' ', '-')

        masjid = input_worksheet.cell_value(i, 15)
        jamiah = input_worksheet.cell_value(i, 80)
        mta = input_worksheet.cell_value(i, 81)
        rs = input_worksheet.cell_value(i, 82)
        data2 = (kuitansi, ID, tgl, masjid, jamiah, mta, rs)
        data.append(data2)

    return data


def read_column(path):

    book = excel.open_workbook(path)
    sheet = book.sheet_by_index(0)

    for i in range(sheet.ncols):
        print(i, sheet.cell_value(0, i))


def read_mysql(table="tjd"):

    mydb = sql.connect(
        host="localhost",
        user="root"
    )

    mycursor = mydb.cursor()
    mycursor.execute("use yogyakarta")

    sql_formula = "SELECT * FROM " + table

    mycursor.execute(sql_formula)

    return mycursor.fetchall()


# membaca janji excel
# excel: 1 ID, 7 periode, 8 perjanjian, 9 nominal
def read_perjanjian(path=r"data.xlsx"):

    book = excel.open_workbook(path)
    sheet = book.sheet_by_index(0)

    nary = {'KomplekMjd': 2, 'Jmh': 3, 'Studio': 4, 'RSID': 5}

    cek = []
    data = []

    for i in range(sheet.nrows):
        if i == 0:
            continue
        if round(sheet.cell_value(i, 1), None) in cek:
            for item in data:
                if str(round(sheet.cell_value(i, 1), None)) in item:
                    if sheet.cell_value(i, 8) in nary.keys():
                        item[nary[sheet.cell_value(i, 8)]+1] += sheet.cell_value(i, 9)
        else:
            cek.append(round(sheet.cell_value(i, 1), None))
            ID = str(round(sheet.cell_value(i, 1), None))
            period = sheet.cell_value(i, 7)
            period = period[:4]
            no = ID+'-'+period[2:4]
            datum = [no, ID, period, 0, 0, 0, 0]
            if sheet.cell_value(i, 8) in nary.keys():
                datum[nary[sheet.cell_value(i, 8)]+1] += sheet.cell_value(i, 9)
            data.append(datum)

    hasil = []

    for item in data:
        hasil.append(tuple(item))

    print("Cek", len(cek), "sama dengan", len(data), "sehingga", len(cek) == len(data))

    return hasil


def insert_perjanjian(data):

    mydb = sql.connect(
        host="localhost",
        user="root"
    )

    mycursor = mydb.cursor()
    mycursor.execute("use yogyakarta")

    sql_formula = "INSERT INTO proyek VALUES (%s, %s, %s, %s, %s, %s, %s)"

    for item in data:
        try:
            mycursor.execute(sql_formula, item)
        except Exception as why:
            print(why, "Error di", item)
            return

    mydb.commit()

    print("Inserting to Database Succeed")


# Path of file excel, str filename -> str path
def path_cd(excel_name):
    return r"D:\project\{}".format(excel_name)


# Path of file excel perjanjian lainnya, str filename -> str path
def path_perj_pl(periode='1920'):
    return r"D:\project\{}.xlsx".format(periode)


# return database, cursor
def connect_to_mysql(database_name="yogyakarta"):
    db = sql.connect(
        host="localhost",
        user="root",
        database=database_name
    )

    my_cursor = db.cursor()
    # mycursor.execute("")

    return db, my_cursor


# Deleting space, /, -, (, ), ., '
def deleting_decoration(the_list):
    new_list = []
    for item in the_list:
        var = item
        var = var.replace('/', '')
        var = var.replace('-', '')
        var = var.replace('(', '')
        var = var.replace(')', '')
        var = var.replace('.', '')
        var = var.replace('\'', '')
        var = ''.join(var.split())
        # print(var)
        new_list.append(var)
    return new_list


# dictionary data type
def data_type_dict():
    tipe_tipe = {
        'M1': 'VARCHAR(30)',
        'ID': 'VARCHAR(10)',
        'Nama': 'VARCHAR(50)',
        'NamaC': 'VARCHAR(50)',
        'Bn': 'VARCHAR(20)',
        'NoTelepon': 'VARCHAR(15)',
        'BulanBayar': 'VARCHAR(15)',
        'Keterangan': 'TEXT',
        'Alh': 'TEXT',
        'TglTransaksi': 'DATE',
        'default': 'INT',

        'id': 'VARCHAR(15)',
        'Periode': 'VARCHAR(10)'
    }
    return tipe_tipe


# Creating table in mysql
def create_table_sql(column_list, table_name, cursor, dict_type):

    inside_query = query_creating_table(column_list, dict_type)

    query = "CREATE TABLE {} ({});".format(table_name, inside_query)
    print(query)
    cursor.execute(query)


# What is my type? return string value
def my_type(name, dict_type):
    if name in dict_type.keys():
        return dict_type[name]
    return dict_type['default']


# sql query formula for creating table, return string
def query_creating_table(list_column, dict_type):
    the_string = ''
    for item in list_column:
        the_string += item + ' '
        the_string += my_type(item, dict_type) + ', '
        if item == list_column[-1]:
            the_string += 'PRIMARY KEY ({})'.format(list_column[0])
            break

    return the_string


# Creating table in database mysql
def creating_table(dataframe, table_name='data'):
    dictionary = data_type_dict()
    # print(dictionary.keys())
    db, mycursor = connect_to_mysql()

    # dataframe = pd.read_excel(r"D:\project\data.xlsx")
    list_of_columns = deleting_decoration(list(dataframe.columns))

    create_table_sql(list_of_columns, table_name, mycursor, dictionary)


# to string all, general list -> string list
def tostringall(the_list):
    final_list = []
    for item in the_list:
        if item == NULL:
            final_list.append(NULL)
            continue
        final_list.append('\''+str(item)+'\'')
    return final_list


# replace nan with 'NULL', list->list
def nan_to_null(the_list):
    final_list = []
    for item in the_list:
        try:
            if math.isnan(item):
                final_list.append(NULL)
            else:
                final_list.append(item)
        except Exception as why:
            final_list.append(item)
    return final_list


# Ganti format tgl transaksi, cth 2021 01 01 -> 2021-01-01, pd_df->pd_df
def ganti_format_tgl(pd_df, kolom='Tgl Transaksi'):
    the_list = list(pd_df[kolom].values)
    for i, item in enumerate(the_list):
        the_list[i] = '-'.join(str(item).split())
    pd_df[kolom] = the_list
    return pd_df


# Ganti tipe data kolom ke str
def change_column_type(pd_df, kolom):
    the_list = list(pd_df[kolom].values)
    for i, item in enumerate(the_list):
        the_list[i] = str(item)
    pd_df[kolom] = the_list
    return pd_df


# Check sameness two lists | list, list -> print(list, list)
def sameness(list_one, list_two):
    final_one = []
    final_two = []
    for item in list_one:
        if item not in list_two:
            final_one.append(item)
    for item in list_two:
        if item not in list_one:
            final_two.append(item)
    print('There are {} problems'.format(len(final_one) + len(final_two)))
    if any(final_one) or any(final_two):
        print('List one:')
        for item in final_one:
            print(item)
        print('\n')
        print('List two:')
        for item in final_two:
            print(item)
    return not (any(final_one) or any(final_two))


# database to Pandas DataFrame
def sql_to_pd(database, table_name='tajnid'):
    return pd.read_sql("select * from {}".format(table_name), database)


# Return unique items in list, list->list
def unique_list(the_list):
    new_list = []
    for item in the_list:
        if item in new_list:
            continue
        new_list.append(item)
    return new_list


# Creating 2D list
def list_2d(rows, columns):
    the_list = []
    for i in range(rows):
        the_list_child = []
        for j in range(columns):
            the_list_child.append(math.nan)
        the_list.append(the_list_child)
    return the_list


# Clear decoration of items in a column, pd df -> pd df
def deleting_decoration_df(dataframe, nama_kolom):
    the_list = list(dataframe[nama_kolom].values)
    the_list = deleting_decoration(the_list)
    dataframe[nama_kolom] = the_list
    return dataframe


# Khusus untuk bagian perjanjian
# Transport values, pd_df pd_df -> pd_df
def transport_perjanjian(df_base, target_df, perj_type='JhIdn'):
    nominals = []
    for nomor_ID in target_df['ID'].values:
        filtered = df_base[[value == nomor_ID for value in df_base['ID']]]
        filtered = filtered[[value == perj_type for value in filtered['PERJANJIAN']]]
        filtered = filtered['NOMINAL PERJANJIAN (Rp)'].values
        if any(filtered):
            nominals.append(filtered[0])
        else:
            nominals.append(0)
    target_df[perj_type] = nominals
    return target_df


# Khusus untuk bagian perjanjian
# Ngisi id, gabungan ID dan PERIODE, pd_df pd_df -> pd_df
def ngisi_id(df_base, df_target):
    # insert id
    ids = []
    for nomor_ID in df_target['ID'].values:
        filtered = df_base[[value == nomor_ID for value in df_base['ID']]]
        the_period = filtered['PERIODE'].values[0]

        ids.append(str(nomor_ID) + str(the_period))

    df_target['id'] = ids
    return df_target


# Insert row in a table mysql, input pd df
def insert_row(dataframe, table_name='data'):
    extract = list(dataframe.values)
    list_of_columns = deleting_decoration(list(dataframe.columns))
    list_of_columns = ', '.join(list_of_columns)

    mydb, mycursor = connect_to_mysql()

    countfull = dataframe.shape[0]
    for i in range(countfull):
        try:
            values = extract[i]
            values = nan_to_null(values)
            values = ', '.join(tostringall(values))
            query_insert = "INSERT INTO {}({}) VALUES ({});".format(table_name, list_of_columns, values)
            mycursor.execute(query_insert)
            print('{} / {}'.format(i + 1, countfull))
        except Exception as why:
            print(why)

    mydb.commit()


# inserting items excel into db
def inserting_row(dataframe, table_name='data'):
    try:
        dataframe = ganti_format_tgl(dataframe)
    except Exception as why:
        print(why)

    insert_row(dataframe, table_name)


# Complete input data excel to database
def input_chid(the_excel='Yogyakarta 20210101-0131.xlsx', nama_table='data'):
    the_database, mycursor = connect_to_mysql()
    dataframe = pd.read_sql('DESCRIBE {};'.format(nama_table), the_database)
    kolom_one = list(dataframe['Field'].values)

    df_2 = pd.read_excel(path_chid(the_excel))
    kolom_two = deleting_decoration(list(df_2.columns))

    if sameness(kolom_one, kolom_two):
        inserting_row(df_2, nama_table)


# Converting excel perjanjian ke pd df ready upload, return pd_df
def converting_perjanjian(the_period='2021'):
    df = pd.read_excel(path_perj_pl(the_period))
    # print(df.head())
    df = deleting_decoration_df(df, 'PERJANJIAN')

    jenis_perjanjian = list(df['PERJANJIAN'].values)
    jenis_perjanjian = unique_list(jenis_perjanjian)

    kolom_kolom = ['id', 'ID', 'Nama', 'Periode'] + jenis_perjanjian

    nama_pejanji = unique_list(list(df['NAMA'].values))
    kumpulan_ID = unique_list(list(df['ID'].values))
    tahun_periode = unique_list(list(df['PERIODE']))

    new_df = pd.DataFrame(list_2d(len(nama_pejanji), len(kolom_kolom)), columns=kolom_kolom)

    new_df['ID'] = kumpulan_ID
    new_df['Nama'] = nama_pejanji
    new_df['Periode'] = [tahun_periode[0] for item in new_df['Periode']]

    # insert id
    new_df = ngisi_id(df, new_df)

    # insert perjanjian-perjanjian
    for jenis in jenis_perjanjian:
        new_df = transport_perjanjian(df, new_df, jenis)

    new_df = change_column_type(new_df, 'ID')

    return new_df


# input data perjanjian pd df to perjanjianlainnya database
def input_perjanjianlainnya(the_period='2021', nama_table='chid'):
    the_database, mycursor = connect_to_mysql()
    dataframe = pd.read_sql('DESCRIBE {};'.format(nama_table), the_database)
    kolom_one = list(dataframe['Field'].values)

    df_2 = converting_perjanjian(the_period)
    kolom_two = deleting_decoration(list(df_2.columns))

    if sameness(kolom_one, kolom_two):
        inserting_row(df_2, nama_table)


if __name__ == '__main__':

    pass