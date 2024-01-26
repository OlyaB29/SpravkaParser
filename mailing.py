import collections

import pandas as pd
from pysendpulse.pysendpulse import PySendPulse
from config import ID, SECRET, TOKEN_STORAGE, MEMCACHED_HOST


# Сбор email по разным листам
def assemble_emails():
    emails = []
    sheet_names = pd.ExcelFile("./spravka.xlsx").sheet_names
    for sheet in sheet_names:
        data = pd.read_excel("./spravka.xlsx", sheet_name=sheet)
        email_col = data.iloc[:,4]
        condition1 = email_col != "Электронная почта"
        condition2 = email_col.notnull()
        sheet_emails = email_col.loc[condition1 & condition2].tolist()
        emails += sheet_emails

    unique_emails = list(collections.OrderedDict.fromkeys(emails))
    print(len(unique_emails))

    # Второй способ
    # unique_emails2 = [e for i, e in enumerate(emails) if e not in emails[:i]]
    # print(len(unique_emails2))

    # Удаление пустых строк
    # data = data.iloc[:,4].dropna(axis=0).tolist()

    # Запись списка email на отдельном листе Excel
    df = pd.Series(unique_emails)
    with pd.ExcelWriter("./spravka.xlsx", mode="a", engine="openpyxl", if_sheet_exists="overlay") as writer:
        df.to_excel(writer, sheet_name="Emails", index=False, header=False)


def get_emails():
    data = pd.read_excel("./spravka.xlsx", sheet_name="Emails", header=None)
    email_list = data[0].to_list()
    return email_list


def make_mailing():
    email_list = get_emails()
    to = list(map(lambda el: {'email': el}, email_list))
    print(to)

    SPApiProxy = PySendPulse(ID, SECRET, TOKEN_STORAGE, memcached_host=MEMCACHED_HOST)
    email = {
        'subject': 'This is the test task from REST API',
        'html': '<h1>Hello, John!</h1><p>This is the test task from https://sendpulse.com/api REST API!</p>',
        'text': 'Hello, John!\nThis is the test task from https://sendpulse.com/api REST API!',
        'from': {'name': 'MyService', 'email': 'myservice@myservice.by'},
        'to': to
    }
    response = SPApiProxy.smtp_send_mail(email)
    print(response)


make_mailing()