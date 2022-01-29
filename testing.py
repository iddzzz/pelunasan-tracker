import mysql.connector
import utils
import pandas as pd


if __name__ == '__main__':
    db, mycursor = utils.connect_to_mysql()

    tabel = 'cid'
    nama = 'Said Ahmad'
    cutoff = '2020-06-30'
    querymysql = "SELECT \
                 SQad, \
                 DanaM, \
                 PembMd, \
                 DanaPend, \
                 DanaThn, \
                 PembStudio, \
                 PembJI, \
                 PembRSID \
                 FROM `{}` WHERE Nama = \"{}\" AND TglTransaksi > \"{}\";"\
        .format(tabel, nama, cutoff)

    df = pd.read_sql(querymysql, db)
    print(df)