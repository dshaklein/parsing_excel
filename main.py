# -*- coding: utf-8 -*-
import re
from openpyxl import load_workbook, Workbook
from collections import defaultdict

rows_to_pass = set()


def fix_card_numbers(sheet, key):
    for i, col in enumerate(sheet[key]):
        if i == 0: continue
        num = str(col.value)
        if not re.match('^\d+$', num):
            rows_to_pass.add(i)


def fix_phones(sheet, key):
    phones = defaultdict(int)
    for i, col in enumerate(sheet[key]):
        if i == 0 or i in rows_to_pass: continue
        phone = str(col.value)
        phone = re.sub('[^\d]', '', phone)
        if phone is None or phone == '':
            col.value = 'NULL'
            continue
        if re.match('^[^78]\d{9}$', phone):
            phone = '7{}'.format(phone)
            print('10 digit number + 7: {}'.format(phone))
        if re.match('^8\d{10}$', phone):
            print(phone)
            phone = '7' + phone[1:]
            print('replace 8 with 7: {}'.format(phone))
        if len(phone) != 11:
            print('not 11 digit number: {}'.format(phone))
            col.value = 'NULL'
            continue
        phones[phone] += 1
        col.value = int(phone)
    return phones


def fix_emails(sheet, key):
    emails = defaultdict(int)
    for i, col in enumerate(sheet[key]):
        if i == 0 or i in rows_to_pass: continue
        email = col.value
        if email is None or type(email) == int or '@' not in email:
            col.value = 'NULL'
            continue
        awoid_domains = ['bj-gold.ru', 'bronnitsy.com', 'noemail.ru', 'no@mail.ru']
        if any(domain in email for domain in awoid_domains):
            rows_to_pass.add(i)
        re.sub('\s+', '', email)
        emails[email] += 1
        col.value = email
    return emails


def find_equal(sheet, key, items):
    repeated_items = {k: value for k, value in items.items() if value > 1}
    repeated_rows = defaultdict(list)
    for i, col in enumerate(sheet[key]):
        if i == 0 or i in rows_to_pass: continue
        item = str(col.value)
        if item in repeated_items.keys():
            repeated_rows[item].append(i)
    for k in repeated_rows:
        repeated_rows[k].sort(key=lambda index: sheet[index][0].value)
    return repeated_rows


def merge_data(sheet, repeated_data):
    """repeated_data: dict with key of repeated instance(email or phone)
    and list of rows where the key is repeated"""
    # заменяем более современными значениями
    # ячейки в расположены по возрастанию, поэтому будет идти от самой последней к самой первой
    # и заполнять пропуски, если они имеются
    for key, rows in repeated_data.items():
        data = {}
        for row in rows[::-1]:
            for cell in sheet[row]:
                if cell.column not in data or data[cell.column] == 'NULL':
                    data[cell.column] = cell.value
                cell.value = None
        for cell in sheet[rows[-1]]:
            cell.value = data[cell.column]
        for row in rows[:-1]:
            rows_to_pass.add(row)
        print(data)


if __name__ == '__main__':

    original = load_workbook('main.xlsx')

    fix_card_numbers(original['cols-2-upload'], 'A')
    phones_2 = fix_phones(original['cols-2-upload'], 'F')
    phones_1 = fix_phones(original['cols-2-upload'], 'E')
    emails = fix_emails(original['cols-2-upload'], 'D')

    phone_2_cells = find_equal(original['cols-2-upload'], 'F', phones_2)
    phone_1_cells = find_equal(original['cols-2-upload'], 'E', phones_1)
    email_cells = find_equal(original['cols-2-upload'], 'D', emails)

    merge_data(original['cols-2-upload'], phone_1_cells)
    merge_data(original['cols-2-upload'], phone_2_cells)
    merge_data(original['cols-2-upload'], email_cells)

    new_wb = Workbook()
    ws = new_wb.active
    ws.title = 'cols-2-upload'
    rs = [r for i, r in enumerate(original['cols-2-upload'].rows) if i not in rows_to_pass]
    for i, r in enumerate(rs, start=1):
        for j, c in enumerate(r, start=1):
            ws.cell(row=i, column=j, value=c.value)
    new_wb.save('updated.xlsx')


# import threading
# from functools import partial
# from concurrent.futures import ThreadPoolExecutor
# def process(tupl, sheet):
#     key, rows = tupl
#     data = {}
#     for row in rows[::-1]:
#         for cell in sheet[row]:
#             if cell.column not in data or data[cell.column] == 'NULL':
#                 data[cell.column] = cell.value
#     with :
#         for cell in sheet[rows[-1]]:
#             cell.value = data[cell.column]
#         for row in rows[:-1]:
#             rows_to_pass.add(row)
#     print(data)
#
#
# def merge_data(sheet, repeated_data):
#     """repeated_data: dict with key of repeated instance(email or phone)
#     and list of rows where the key is repeated"""
#     # заменяем более современными значениями
#     # ячейки в расположены по возрастанию, поэтому будет идти от самой последней к самой первой
#     # и заполнять пропуски, если они имеются
#     extended_map = partial(process, sheet=sheet)
#     args = list(repeated_data.items())
#     with ThreadPoolExecutor(max_workers=5) as tpe:
#         for _ in tpe.map(extended_map, args):
#             pass