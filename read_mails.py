
import pandas as pd

# Read file from given textfile
def read_file(file):
    f = open(file, 'r')
    mails = []
    while f.readline():
        line = f.readline().strip()
        if line not in mails:
            mails.append(line)
    return mails

# Read from report from hoopla
def read_hoopla(file):
    # Get the cols needed in the DataFrame, [name, email]
    df = pd.read_excel(file, usecols=[22, 30])

    mail_dict = {}
    # Iterate over the DataFrame and make a dictionary containing name and email
    for idx, row in df.iterrows():
        if row[1] not in mail_dict:
            mail_dict[row[1]] = [row[0]]
        else:
            mail_dict[row[1]].append(row[0])

    # Return the mails in a dictionary containing names and emails
    return mail_dict

def read_starters(file):
    # Get the cols needed in the DataFrame, [name, email]
    df = pd.read_excel(file, usecols=[7, 19])

    mail_dict = {}
    # Iterate over the DataFrame and make a dictionary containing name and email
    for idx, row in df.iterrows():
        if row[1] not in mail_dict:
            mail_dict[row[1]] = [row[0]]
        else:
            mail_dict[row[1]].append(row[0])

    # Return the mails in a dictionary containing names and emails
    return mail_dict

start_mails = read_starters('hoopla_start.xls')
print(start_mails)
