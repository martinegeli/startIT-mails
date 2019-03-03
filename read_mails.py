
import pandas as pd

def read_file(file):
    f = open(file, 'r')
    mails = []
    while f.readline():
        line = f.readline().strip()
        if line not in mails:
            mails.append(line)
    return mails

def read_hoopla(file):
    # Get the cols needed in the DataFrame
    df = pd.read_excel(file, usecols=[22, 30])

    mail_dict = {}
    # Iterate over the DataFrame and make a dictionary containing name and email
    for idx, row in df.iterrows():
        if row[1] not in mail_dict:
            mail_dict[row[1]] = [row[0]]
        else:
            mail_dict[row[1]].append(row[0])

    return mail_dict


mail_dict = read_hoopla('hoopla.xls')
for key, value in mail_dict.items():
    if len(value) == 1:
        names = value[0]
        plural = "din billett"
    else:
        names = ", ".join(value[0:len(value) - 1]) + "og " + value[-1]
        plural = "deres billetter"
print(mail_dict)


