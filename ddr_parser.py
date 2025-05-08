import csv
import datetime
import os
import inspect

today = datetime.date.today().strftime('%m-%d-%Y')
directory = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
dictionary = {}
fields = ['Date', 'Desc', 'Interest']
file = f'\\ddr-out {today}.csv'
filename = directory + file
final = []
ddr = open(directory + '\\ddr_input.txt')
count = 0

for line in ddr:
    count = count + 1
    data = line.split()
    date = data[0]
    desc = data[1]
    interest = data[-2]
    dictionary["Date"] = date
    dictionary ["Desc"] = desc
    dictionary["Interest"] = interest
    final.append(dictionary.copy())

with open(filename, 'w', newline = '') as csvfile:
    writer = csv.DictWriter(csvfile, fieldnames = fields)
    writer.writeheader()
    writer.writerows(final)