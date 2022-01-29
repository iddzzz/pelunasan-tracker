import utils
import mysql.connector as sql
import xlrd as excel

path = input('The path: ')

data = utils.read_perjanjian()

for item in data:
    print(item)

utils.insert_perjanjian(data)