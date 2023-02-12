# This is a sample Python script.

#import pandas
#Revisions 1.0 Todat 2/11/: playing hith github

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.
#if __name__ == '__main__':
#    print_hi('PyCharm')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

import os
import openpyxl
import json
import pandas as pd
import os.path
from openpyxl import load_workbook
import numbers

citi_regex = "placeholder"
first_tech_regex = "placeholder"
us_bank_regex = "nop"
regex_list = "yuo"
bank_dictionary = {}

def intro():
    print("1 for Capital One")
    print("2 for First Tech")
    print ("3 for US Bank")
    print ("4 for a sum of costs")
    print ("0 for exit")
    selection = input("select one\n")
    return selection
# line 82 for code that uses this menu
category_items = "1: boat\n 2: food\n 3:dining\n 4: house\n 5:travel\n " \
                 "6: utilities\n 7: deposit\n 8: medical\n 9: charitable\n 20:print bank_data\n 21: print sum\n 99: save data to desktop"
category_items1 = {1:"boat", 2:"food" ,3:"dining", 4:"house", 5:"travel",
                  6:"utilities", 7:"deposit", 8:"medical", 9:"charitable", 20:"print bank_data", 21: "print sum", 99: "save data to desktop"}
class Bank:
    def __init__(self, bank_name:str, bank_file_name:str, month:float,
    download_directory:str, excel_us_bank_headings):
        self.bank_name = bank_name
        self.bank_file_name = bank_file_name
        self.month = month
        self.download_directory = download_directory
        self.excel_us_bank_headings = excel_us_bank_headings

    def get_wb_and_ws (self):
        list_of_bank_files = []
        print (f" bank file name is: {self.bank_file_name} bank name is: {self.bank_name}")
        x = excel_us_bank_headings["transaction_date"]
        for file_name in os.listdir(self.download_directory):
            if file_name.endswith(".xlsx") and (self.bank_file_name in file_name) == True :
                list_of_bank_files.append(file_name)
                print(f' file to open:{self.download_directory + "/" + file_name}')
                wb = openpyxl.load_workbook(self.download_directory + "/" + file_name )
                ws = wb.active
                if self.bank_name == "us_bank":
                    cell_date = ws['A2'].value
                else:
                    cell_date = ws['B3'].value
                if self.month == cell_date.month:
                    return (self.download_directory + "/" + file_name, wb, ws)
                else:
                    return None
    def get_statement_length():
        return (ws.min_row,  ws.max_row, ws.min_column,  ws.max_column)


class Add_categories:
    def __init__(self, statement_headings, bank_file_name, wb, ws, bank_name):
        self.statement_headings = statement_headings
        self.bank_file = bank_file_name
        self.wb = wb
        self.ws = ws
        self.bank_name = bank_name
#menu at line 41
    def get_category(self):
        self.print_menu()
        c = self.fetch_category_number()
        if int(c) == 20:
            self.print_json()
        if int(c) == 21:
            s = Summaries()
            s.add_money()
        if int(c) == 99:
            save_file = ("c:\python-write-data\saved_bank_data.xlsx")
            s = Save_and_load(save_file)
            s.save() #line 145
            print("SAVED")
            print ("CLOSED")
            exit()
        return c
        #excel_us_bank_headings = {"transaction_date":0, "type":1, "description" : 2,
        # "download": 3, "debit" : 4 } line:

    def get_size_of_bank_statement(self):
        pass
#        bd_rows = ws.max_row
#        for i in ws.iter_rows(min_row=2, max_row=bd_rows, min_col=0, max_col=7,
#                        values_only=True):
#            x = self.statement_headings["description"]
#            y = self.statement_headings["debit"]
#            z = self.statement_headings["credit"]
#===========================================|
#   Class: Add_categories                   |
#       Fetch One Line of Bank Statement    |
#===========================================|


class Summaries():
    def add_money(self):
        sum= 0
        one_key = 0
        global bank_sum
        for keys in bank_data:
            dd  = len (bank_data[keys][0])
            x = bank_data[keys][0]
            for q in range(0,dd):
                one_key = x[q]

            bank_sum = bank_sum + one_key
            print(f"amount is: {sum}")

#called from line 93
class Save_and_load():
    def __init__(self, save_file_name_and_location):
        self.save_file_name_and_location = save_file_name_and_location

    def save_bank_data_to_excel(self):
        pass


    def close(self):
        pass
    def save (self):  #
#        wb = openpyxl.load_workbook(self.save_file_name_and_location)
        wb1 = openpyxl.Workbook()

        ws1 = wb1.active
        ws1.title = "new sheet name"
        ws1['B2'] = 1
        wb1.save ("c:\python-write-data/test.xlsx")
#=====

        df = pd.DataFrame(data=bank_data)
        df_t = df.T
        df_t.to_excel("c:\python-write-data/bank_data.xlsx")
        return (wb1)


class Populate_cc(Add_categories):
    def __init__(self, bank_name):
        self.bank_name = bank_name
    pass
class Populate_us_bank(Add_categories):
    pass
class Populate_first(Add_categories):
    pass

class Populate_category_file():
    def __init__(self, category_file, bank_data):
        self.category_file = category_file
        self.bank_data = bank_data



    def open_category_file(self):
        def __inti__ (self, category_file):
            self.category_file = category_file

        wb = openpyxl.load_workbook(self.category_file)
        ws = wb.active

    def add_category(self):
        def __init__(self):
            print ("add a category for")


class New_statement_row:
    def __init__(self, row_tupple, column_headings):
        self.row_tupple = row_tupple
        self.column_headings = column_headings
    pass
    def change_sign_if_credit(self):
        pass
        y = row_tupple
        if isinstance(self[5],numbers.Number) == True:
            debit = self[5]
        else:
            debit = int(0)
        if isinstance(self[4],numbers.Number) == True:
            credit = self.row_tupple[z]
        else:
            credit = int(0)
        return (debit, credit)

    def update_amounts(self):
        if self.row_tupple[2] in bank_data:
            print (f"{i[x]}  ->already exists in bank_data<-")
            bank_data[i[x]][0] +=  int(debit)
            bank_data[i[x]][4] +=  int(credit)
        return
    def update_category(self):
            print (f"{self.row_tupple[1]} -> does not exist in bank_data <-")
#            cat_num = self.get_category()
#            type_purchase = category_items1[int(cat_num)]
#            bank_data[i[x]]= [int(debit), int(cat_num), type_purchase,
#                            bank_name, int(credit)]
            pass


#


    def fetch_category_number(self):
        i = input ("pick a number")
        return(i)

    def print_json(self):
        print("=========================")
        print(json.dumps(bank_data, indent=4))
        print("=========================")

    def print_menu(self):
        print ("pick category\n")
        print (category_items)


bank_data = {}
bank_sum = 1
uts = {"capital_one":"transaction_download" ,"us_bank":"Checking" ,"first_tech": "ExportedTransactions"}
excel_credit_card_headings = {"transaction_date": 1, "card_number": 2, "description":3, "category": 4, "debit": 5, "credit":6}
excel_first_tech_headings = {"transaction_date": 2, "card_number": "n/a", "check_number": 5, "type":9, "description":7, "category": 8, "debit": 4, "credit":"n/a, "}
excel_us_bank_headings = {"transaction_date":0, "type":1, "description" : 2, "download": 3, "debit" : 4, "credit" : 5}
banks = ["capital_one", "first_tech"," US_bank"]
month_wanted = 1
bank_choice = "9"
bank_name = "none picked yet"
bank_choice = intro()
statement_directory = "f:Libraries-System-Win10/Downloads"

if os.path.isfile("c:\python-write-data/bank_data.xlsx"):
#    bank_data = pd.read_excel("c:\python-write-data/bank_data.xlsx")
    wb = load_workbook("c:\python-write-data/bank_data.xlsx")
    ws = wb.active
    for row in list(ws.rows)[1:]:
        bank_data[row[0].value] = [c.value for c in row[1:]]
    pass
    pass
else:  #create one
    wb = openpyxl.Workbook()
    ws = wb.active

 #   exit("no file in excel directory line 233")

while bank_choice != str(0):
    if bank_choice == str(1):
        bank_name = "Capital One"
        bank_file = Bank("capital_one", uts["capital_one"],month_wanted,
                    statement_directory, excel_credit_card_headings)
        y = bank_file.get_wb_and_ws()  # tupple of file_name, wb and ws
        if y == None:
            print ("NO MONTH FOUND")
        ws = y[2]
        wb = y[1]
        bank_file_name = y[0]
        pop_credit_card = Add_categories(excel_credit_card_headings, bank_file_name,
                          wb, ws, bank_name)
        size = Bank.get_statement_length()
        for new_row in ws.iter_rows(min_row=2, max_row=size[3], min_col=0, max_col=7,
                              values_only=True):
            new_statement_row = New_statement_row(new_row, excel_credit_card_headings)
            new_statement_row.change_sign_if_credit
            new_statement_row.update_amounts()
            new_statement_row.update_category()
            request_new_category()  <----------------------------need to write these two
        bank_name = "finished Capital One"





    elif bank_choice == str(2):
        bank_name = "First Tech"
        bank_file = Bank("first_tech", uts["first_tech"],month_wanted,
                         statement_directory, excel_first_tech_headings)
        y = bank_file.get_wb_and_ws()  # tupple of file_name, wb and ws
        ws = y[2]
        wb = y[1]
        bank_file_name = y[0]
        pop = Add_categories
        pop_first_tech = pop(excel_first_tech_headings, bank_file_name, wb, ws, bank_name)
        pop_first_tech.read_excel_and_populate_dictionary()
        bank_name = "finished First Tech"

    elif bank_choice == str(3):
        bank_name = "US Bank"
        bank_file = Bank("us_bank", uts["us_bank"], month_wanted,
                         statement_directory, excel_us_bank_headings)
        y = bank_file.get_wb_and_ws()  # tupple of file_name, wb and ws
        if y == None:
            exit ("NO FILE FOUND for REQUESTED MONTH")
        ws = y[2]
        wb = y[1]
        bank_file_name = y[0]
        pop = Add_categories
        pop_us_bank = pop(excel_us_bank_headings, bank_file_name, wb, ws, bank_name)
        pop_us_bank.read_excel_and_populate_dictionary()
    elif bank_choice == str(0):
        pass









#pop_cc = Populate(excel_credit_card_headings, y, wb, ws)
#pop_cc.read_excel_and_populate_dictionary()
#category_file = "f:\Libraries-System-Win10\Desktop\category_file.xlsx"

#bank_file.read_excel_and_populate_dictionary()


