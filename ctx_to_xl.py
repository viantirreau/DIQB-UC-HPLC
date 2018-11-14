import xlsxwriter
import pandas

FILEPATH = "ctx/B_car2.ctx"
df = pandas.read_csv(FILEPATH, sep=";", decimal=",")
print(df)