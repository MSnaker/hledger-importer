from openpyxl import load_workbook
from datetime import datetime
import yaml
import re

def len_str(*args) -> int:
    totlen = 0
    for arg in args:
        if type(arg) != str:
            raise TypeError
        foo = len(arg)
        totlen += foo

    return totlen

def conditional_get_col(category:dict) -> int | str:
    try:
        if category['col'] is None: raise KeyError('Missing Column value in config file')
        column = ord((category['col']).upper()) - 65
        return column
    except KeyError:
        value = category['val']
        return value

def convert_to_date(val) -> datetime.date:
    
    if datetime == type(val): 
        return val.strftime('%Y-%m-%d')
    elif str == type(val): 
        return datetime.fromisoformat(val).strftime('%Y-%m-d')

def get_transaction(row) -> tuple:

    date = convert_to_date(row[transaction_date].value) if int == type(transaction_date) else convert_to_date(transaction_date)
    destination = row[transaction_dest].value if int == type(transaction_dest) else transaction_dest
    source = row[transaction_source].value if int == type(transaction_source) else transaction_source
    description = row[transaction_description].value if int == type(transaction_description) else transaction_description
    value = row[transaction_value].value
    ticket = row[transaction_ticket].value if int == type(transaction_description) else None


    return (date, destination, source, description, value, ticket)


####################################################
#               MAIN SECTION                       #
####################################################

with open('conf.yml', 'r', encoding='utf-8') as fconf:
    conf = yaml.safe_load(fconf)

journalfile = conf['journalfile']
importfile = conf['filename']
wb = load_workbook(filename=importfile)
worksheet = wb[conf['sheetname']]

transaction_date = conditional_get_col(conf['transaction date'])
transaction_description = conditional_get_col(conf['transaction description'])
transaction_source = conditional_get_col(conf['transaction source'])
transaction_dest = conditional_get_col(conf['transaction dest'])
transaction_value = conditional_get_col(conf['transaction value'])
transaction_ticket = conditional_get_col(conf['transaction ticket'])
try:
    data_range = range(conf['data']['start row'] - 1, conf['data']['end row'])
except KeyError:
    data_range = None

with open(journalfile, "a", encoding="utf-8") as fw:
    # rows_to_analyze = worksheet.rows[data_range[0] : data_range[1]] if data_range is not None else worksheet.rows

    for i, row in enumerate(worksheet.rows):
        if i in data_range:
            date, destination, source, description, value, ticket = get_transaction(row)

            if (0 == ticket) or (ticket is None):
                ticket_desc = ""
                ticket_value = 0
                value_source = value
            else:
                ticket_desc = 'Marco:assets:Ticket Restaurant'
                ticket_value = 8 * ticket 
                value_source = float(value) - ticket_value
            
            value_source_stringified = f"{value_source:.2f}"
            value_dest_stringified = f"{value:.2f}"
            line1 = f"\n{date} {description}\n"

            len2 = len_str(destination, value_dest_stringified) + 8
            len3 = len_str(source, value_source_stringified) + 9
            len4 = len_str(ticket_desc, str(ticket_value)) + 9 if ticket is not None else 0
            padding = max(len2, len3, len4)
            
            line2 = f"\t{destination:<{padding}}{value_dest_stringified}\n"
            line3 = f"\t{source:<{padding-1}}-{value_source_stringified}\n"

            line4 = f"\t{ticket_desc:<{padding-1}}-{ticket_value}" if (ticket is not None and 0 != ticket) else ""
            string_to_append = line1 + line2 + line3 + line4
            print(string_to_append)
            # fw.write(string_to_append)