
import random

"""
Python script for choosing three random winners from the answers of the KPI
"""

def read_and_return(file):
    f = open(file, 'r')
    l = []
    while f.readline():
        l.append(f.readline().strip())
    f.close()
    result = []
    while len(result) < 3:
        r = random.randint(0,len(l) - 1)
        if l[r] not in result and l[r] != '':
            result.append(l[r])
    return result

print(read_and_return('kpi_mails.txt'))