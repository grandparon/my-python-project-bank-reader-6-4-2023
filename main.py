
#import pandas
#Revisions 1.0 Todat 2/11/: playing hith github
# revision 2.0 3/21 Added SBF (Standard Banking Format) and removed redundent sections of main for each bank


# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
#*****************************************************************************
# **************************** GLOBAL VARIABLE DECLARATIONS  *****************
#*****************************************************************************
#bank_data = {"Beginning Key":[0,0,0,0,0,0,0,0,0]}
bank_data = {}
#category_sums = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]

category_sums_dict = {'0':[0],'1':[0],'2':[0], '3':[0],'4':[0],'5':[0],'6':[0],'7':[0],'8':[0],'9':[0],'10':[0],'11':[0],'12':[0],'13':[0],'14':[0],'15':[0],'16':[0],'17':[0]}
sums_info_dict = {"date":"str_date"}
pass
save_file = ("c:\python-write-data\saved_bank_data.xlsx")
statement_directory = "f:Libraries-System-Win10/Downloads"
sbf_dict = {}
# order below is --->  "NV",   capital one , 1st Tech, US Bank
statement_headings_dict = {"STORE": ["NV",3,7,2], "AMOUNT": ["NV",5,4,4], "DATE": ["NV",1,2,0], "CATEGORY": ["NV",4, 8, 1],"PAYMENT": ["NV",6, "NV", "NV"]}
bank_file_name_dict = {"CAPITAL_ONE":"transaction_download." ,"US_BANK":"Checking" ,"FIRST_TECH": "ExportedTransactions"}
sbf_headings_dict = {"BANK NAME": 0, "DATE": 1, "AMOUNT": 2, "CATEGORY":4}
month_wanted = 1
temp_cat_sums_from_file = {}
# *****************************************************************************


import os
import openpyxl
import csv
import json
import pandas as pd
import os.path
from openpyxl import load_workbook
import numbers
from enum import Enum
import datetime as dt
from datetime import date


us_bank_regex = "nop"
regex_list = "yuo"
bank_dictionary = {}
bank_name_dict = { 0: "EXIT", 1: "CAPITAL_ONE", 2: "FIRST_TECH", 3: "US_BANK"}
category_dict = {1:"Boat", 2: "Food", 3:"Dining", 4: "House", 5:"Travel", 6:"Utilities", 7:"Auto", 8:"Health",
                 9:"Deductable", 10: "Recreation", 11: "Unknown", 12: "Pay Credit Card", 13: "Deposits to USB",
                 14:"Money Transfers", 15: "Interest", 16:"Trivials",17: "Taxes", 20: "Print bank_data", 21: "Print Sums",
                 22: "Quit but don't Save", 23: "Save and Quit"}

def month_intro():
    mw = input ("use month number")
    if 1 <= int(mw) <= 12:
        return (int(mw))
    else:
        return ("NOT A VALID MONTH")
def intro():
    selection = "NONE"
    print("1 for Capital One")
    print("2 for First Tech")
    print ("3 for US Bank")
    print ("4 for a sum of costs")
    print ("0 for exit")
    selection = input("select one\n")
    if (int(selection) >= 0 and int(selection) < 4 ):
        return (int(selection), bank_name_dict[int(selection)])
    else:
        return ("NONE" "NONE")



def  get_key_by_value(val):
    for key, value in category_dict.items():
        if val == value:
            return key


def get_category_name_from_number(category_number):
    return (category_dict[category_number])



class Sbf_enum (Enum):
    COMPANY             = "IS KEY"
    BANK_NAME           = 0
    TRANSACTION_DATE    = 1
    AMOUNT              = 2
    CATEGORY            = 3
    STATEMENT_MONTH     = 4
    STR_CATEGORY        = 5
    PAYMENT             = 6

class First_Tech_format (Enum):
    DATE            = 2
    DEBIT_OR_CREDIT = 3
    AMOUNT          = 4
    COMPANY         = 8
    CATEGORY        = 9


class Bank_Names (Enum):
    NONE                = 0
    CAPITAL_ONE         = 1
    FIRST_TECHNOLOGY    = 2
    US_BANK             = 3

class Num_To_Month (Enum):
    JANUARY         = 1

class SBF:
    def __init__(self):
        pass

class Pivot:
    pass
 #   def __init__(self, wb, ws, bank_data):
class Bank:
    def __init__(self, bank_name:str, bank_num:int, bank_file_name:str, month:float,
    download_directory:str):
        self.bank_name = bank_name
        self.bank_name = bank_num
        self.bank_file_name = bank_file_name
        self.month = month
        self.download_directory = download_directory
    pass

    def csv_to_xl(self):
        # (self.bank_file_name in file_name)
        #    file_is = "2023-01-30_transaction_download.csv"

        for file_name in os.listdir(self.download_directory):
            if file_name.endswith(".csv") and (self.bank_file_name in file_name) == True :
#                list_of_bank_files.append(file_name)
                a_csv_file = statement_directory + '/' + file_name
                wb = openpyxl.Workbook()
                ws = wb.active
                with open(a_csv_file) as f:
                    reader = csv.reader(f, delimiter=",")
                    for row in reader:
                        ws.append(row)
                file_name_with_dot_xlsx = file_name.replace("csv", "xlsx")
                wb.save(statement_directory + "/" + file_name_with_dot_xlsx )
                os.remove(statement_directory + "/" + file_name )

    def get_wb_and_ws (self):
        list_of_bank_files = []
        pass
#        print (f" bank file name is: {self.bank_file_name} bank name is: {self.bank_name}")
        for file_name in os.listdir(self.download_directory):
            if file_name.endswith(".xlsx") and (self.bank_file_name in file_name) == True :
                list_of_bank_files.append(file_name)
#                print(f' file to open:{self.download_directory + "/" + file_name}')
                wb = openpyxl.load_workbook(self.download_directory + "/" + file_name )
                ws = wb.active
                if bank_name == "FIRST_TECH":
                    cell_date = ws['B3'].value
#                    cell_month = int(cell_date.strftime("%m"))
                    cell_split_date = cell_date.split('/')
                    cell_month = int(cell_split_date[0])
                else:
                    cell_date = ws['A3'].value
                    cell_split_date = cell_date.split('-')
                    cell_month = int(cell_split_date[1])
                if self.month == cell_month:
                    pass
                    return (self.download_directory + "/" + file_name, wb, ws)

    def get_sheet_length(a_bank_file_name, sheet_named):
#test to see sheet exists
        df_dict = pd.read_excel(a_bank_file_name, sheet_name=None)
        if sheet_named in df_dict:
            df1 = pd.read_excel(a_bank_file_name, sheet_name=sheet_named)
            num_rows = len(df1)
            return num_rows
        else:
            return


    def get_statement_length(a_bank_file_name):
            pass
            if os.path.isfile(a_bank_file_name):
                dataframe1 = pd.read_excel(a_bank_file_name)
                pandas_num_rows = dataframe1.shape[0]
        #       print (f"pandas num rows:{num_rows}")
                return (ws.min_row,  ws.max_row, ws.min_column,  ws.max_column, pandas_num_rows)
            else:
                print(f" could not find {a_bank_file_name}")
                exit








class Add_categories:
    def get_category(self):
        sc = 0
        nc2 = 0

        c = input ("pick a number")
        if 0 <= int(c) < 20:
            return int(c), int(0)
        if (20 < int(c) <=30):
            print (f" ok special but lets finish this statement row first")
            nc2 = input("pick a number for category")
            return int(nc2), int(c)
        else:
            print ("NOT A VALID CATEGORY NUMBER")
            return int(-1), int(-1)


    def print_category_menu(self):
#            k = sbf_dict[new_cat_key][2]
#            print(f"PICK CATEGORY for {new_cat_key} cost is:${k}")
            for key, value in category_dict.items():
                print(f'{key}\t{value}')

    def parse_cat_nums_and_non_standard_exit_programs(self):
        if 0 < cat_num < 20:
            return ( False, False)
        if cat_num == 20:
            print ("PRINTING BANK_DATA then continue")
            self.print_json()
            return (20, True)
        elif cat_num == 21:
            print ("PRINTING CATEGORY SUMS then continue")
            ms = Master_sum(0,0,0)
            ms.print_cat_sums()
            return (21, True)
        elif cat_num == 22:
            print ("EXITING without saving session")
            exit()
        elif cat_num == 23:
            print ("SAVE and EXIT after this entry")
            return (23, True)


    def add_new_cat_to_sbf_dict(self, cat_num, new_sbf_dict, cat_num_text_str):
        x = list (new_sbf_dict)[-1]
        pass
        new_sbf_dict[x][Sbf_enum.CATEGORY.value] = cat_num
        new_sbf_dict[x][Sbf_enum.STR_CATEGORY.value] = cat_num_text_str
        pass
    def print_json(self):
        print("=========================")
        print(json.dumps(bank_data, indent=4))
        print("=========================")
class Category_to_xl_column:
    pass
    def read_bank_data(self):
        pass
class Master_sum:
    def __init__(self, new_amount, cat_num, sbf_dict):

        self.new_amount = new_amount
        self.cat_num = cat_num
        self.sbf_dict = sbf_dict

    pass
    def fix_new_amount_signs (self, a_new_sbf_dict_key):
        pass
        amount_to_add = sbf_dict[a_new_sbf_dict_key][sbf_headings_dict["AMOUNT"]]
        if amount_to_add >= 0:
            total_deposits = amount_to_add;
            return ( 0, total_deposits,self.cat_num)  #return (deposits, abs(debits) )
        elif amount_to_add < 0:
            #amount = abs(amount_to_add)
            amount = round(float(amount_to_add),2)
            pass
            return ( amount, 0, self.cat_num)




#int(round(float(amount_to_add))))




    def print_cat_sums (self):
        i=0
        pass
        for i in range (1, 13):
            print (f"sum for {category_dict[i]} is ${category_sums[i]}")

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
 #           print(f"amount is: {sum}")

#called from line 93
class Merge_and_save:
    def __init__(self, save_file, sbf_dict):
        pass
        self.save_file = save_file
        self.sbf_dict = sbf_dict


    def save_bank_data_to_excel(self):
        pass

    def merge_data(self):
        pass
        bank_data.update(sbf_dict)
        pass

    def merge_sums_dicts(self):

        category_sums_dict.update(sums_info_dict)

    def save (self):
        with pd.ExcelFile(save_file) as xls:
            if "category sums" in xls.sheet_names:
                df = pd.read_excel(xls, "category sums")
                temp_cat_sums_from_file = df.to_dict('list')
                pass
            else:
                temp_cat_sums_from_file = category_sums_dict
        sums_info_dict["date"] = date.today()
        sums_info_dict["Bank"] = bank_name
        pass
        #--------
        xl_lengths  = Bank.get_sheet_length(save_file, "category sums")
        category_sums_dict.update(sums_info_dict)
        df_from_bank_data = pd.DataFrame.from_dict (bank_data)
        df_from_cat_sums = pd.DataFrame.from_dict (category_sums_dict)
        df_from_saved_file = pd.DataFrame.from_dict (temp_cat_sums_from_file)
        excel_pandas_file = pd.ExcelFile(save_file)
        with pd.ExcelWriter(save_file) as writer:
 #           df_from_bank_data.T.to_excel(writer, sheet_name="Sheet1")
            if "category sums" in excel_pandas_file.sheet_names:
                df_from_bank_data.T.to_excel(writer)
                df_from_cat_sums.to_excel(writer, sheet_name="category sums", startrow=0, index=False)
                df_from_saved_file.to_excel(writer, sheet_name="category sums", startrow=2, index=False)
            else:
#                with pd.ExcelWriter(save_file) as writer:
                df_from_bank_data.T.to_excel(writer, sheet_name="Sheet1")
                df_from_cat_sums.to_excel(writer, sheet_name="category sums")







    def add_category(self):
        def __init__(self):
            print ("add a category for")
#===========================================
# Class to for all new statement rows
#===========================================
class New_statement_row:
    def __init__(self, new_row, bank_name, bank_num):
        self.new_row = new_row
        self.bank_name = bank_name
        self.bank_num = bank_num


#----------------------------------------------------------------------------------------------------
#    SBF definition : Bank_Instition   Transaction_date   Statement_month   Store   Amount  Category
# --------------------------
# Copies the new statement row to RAM (so CATEGORY can be added)
#----------------------------------------------------------------------------------------------------

#SBF is:
    def new_sbf_dict(self):
        pur_date = self.new_row[statement_headings_dict["DATE"][bank_num]]
#        purchase_date = pur_date.strftime ('%m/%d/%y')
        pass
        sbf_dict [new_row [statement_headings_dict["STORE"][bank_num]]] = [ bank_name, pur_date,
                                                                            int(round(float(new_row[statement_headings_dict["AMOUNT"][bank_num]]),2)),
                                                                            "RESERVED FOR ASSIGNED CATEGORY NUMBER",
                                                                            month_wanted,
                                                                            new_row[statement_headings_dict["PAYMENT"][bank_num]],
#                                                                            new_row[statement_headings_dict["CREDIT"][bank_num]] ,
                                                                            "RESERVED", "RESERVED", "RESERVED"]
 #       if bank_name == "CAPITAL ONE":
 #           modified_credit = sbf_dict[statement_headings_dict["CREDIT"][bank_name]]
 #           modified_debit = sbf_dic[statement_headings_dict["CREDIT"][bank_name]]

 #   def copy_sbf_row_to_sbf_dict(self, sbf_row):
 #       print (f"THIS is sbf_row {sbf_row}  <- in copy_sbf_row...")
#        sbf_dict[sbf_row[Sbf.COMPANY.value]] = [  "KEY", sbf_row[Sbf.BANK_NAME.value], sbf_row[Sbf.TRANSACTION_DATE.value], sbf_row[Sbf.AMOUNT.value],
#                                                  new_cat,    sbf_row[Sbf.STATEMENT_MONTH.value], "EXTRA"]
#        pass
        return (new_row [statement_headings_dict["STORE"][bank_num]])

    def is_statement_row_in_bd_and_not_in_saved_xl(self, new_statement_key):
        temp_saved_file = {}
        if os.path.isfile(save_file):
            wb1 = load_workbook(save_file)
            ws1 = wb1.active
            pass
            for row in list(ws1.rows)[1:]:
                pass
                temp_saved_file[row[0].value] = [c.value for c in row[1:]]
            temp_saved_file.pop("null", "not_found")
            pass
        else:
            print ("NO SAVED FILE FOUND")
            exit("NO")
            pass
        if new_statement_key in bank_data:
            print (f"found {new_statement_key} in bank_data")
            if new_statement_key in temp_saved_file:
                found_one = True
            else:
                found_one = False

            if (found_one == False):
                print (" found {key} IN bank_data NOT in saved_file")
                pass
            pass
    def is_it_new_to_bank_data (self, sbf_dict, new_row):
        new_statement_key = new_row[statement_headings_dict["STORE"][bank_num]]

        new_amount = round(float(new_row[ statement_headings_dict["AMOUNT"][bank_num]]),2)
        pass
        new_statement_date = new_row[statement_headings_dict["DATE"][bank_num]]
        new_statement_amount = round(float (new_amount),2)
        in_xl_not_saved = self.is_statement_row_in_bd_and_not_in_saved_xl (new_statement_key)
        if new_statement_key in bank_data.keys():
            category_num_from_bank_data = bank_data[new_statement_key][sbf_headings_dict["CATEGORY"]]
            if(bank_data[new_statement_key][sbf_headings_dict["DATE"]] == new_statement_date and bank_data[new_statement_key][sbf_headings_dict["AMOUNT"]] == new_amount):
                return "BANK_DATA EXACTLY SAME AS NEW STATEMENT",new_statement_key, new_statement_amount, new_statement_date,category_num_from_bank_data
            else:
                return "EXISTS IN BANK_DATA", new_statement_key,new_statement_amount, new_statement_date, category_num_from_bank_data
        if  sbf_dict:  # has got stuff in it if true
            list_of_keys = list (sbf_dict.keys())
            my_list = "EXISTING ENTRY", list_of_keys[-1], new_statement_amount
            if new_statement_key in sbf_dict.keys():
                pass
                print ("exist")
            else:
                pass
        return "NEW ENTRY", new_statement_key, new_statement_amount, new_statement_date, "NO CATEGORY UNTIL PICKED"
#            return "NEW ENTRY"

    def  update_amounts(self, latest_sbf_dict, new_amount_to_check):
        pass
        new_amount_to_check = int(round(float(new_amount_to_check),2))
        print(f" ==>>> check to see amounts are different  {new_amount_to_check} vs. {latest_sbf_dict[Sbf_enum.AMOUNT.value]}<<=\n")
        if latest_sbf_dict[Sbf_enum.AMOUNT.value] != new_amount_to_check:
#            print (f" ======>>> found amounts to be different and will update {new_amount_to_check}  vs.  {latest_sbf_dict[Sbf_enum.AMOUNT.value]}<<==")
            latest_sbf_dict[Sbf_enum.AMOUNT.value] = new_amount_to_check
            pass
#       sbf_dict[statement_headings_dict["AMOUNT"][bank_num]] += sbf_dict[statement_headings_dict["AMOUNT"][bank_num]]

        return


#*********************************************************
#category Sum
#*********************************************************

#*******************************************
# Statement headings for each bank
#*******************************************
#excel_credit_card_headings = {"transaction_date": 1, "card_number": 2, "description":3, "category": 4, "debit": 5, "credit":6}
#excel_first_tech_headings = {"transaction_date": 2, "card_number": "n/a", "check_number": 5, "type":9, "description":7, "category": 8, "debit": 4, "credit":"n/a, "}
#excel_us_bank_headings = {"transaction_date":0, "type":1, "description" : 2, "download": 3, "debit" : 4, "credit" : 5}

#in dict {key: [CAPITAL ONE, FIRST TECH, US BANK]}

#banks = ["capital_one", "first_tech"," US_bank"]

bank_choice = "9"
bank_name = "none picked yet"
total = 0
special_cat_num = 0


#++++++++++++++++++++++++++++++++++++++++++++++++++++++
#   Start of Main
#+++++++++++++++++++++++++++++++++++++++++++++++++++++
t_or_f_is_it_special = False
bank_num, bank_name = intro()
month_wanted = month_intro()
bank_file_1 = Bank(bank_name_dict[bank_num], bank_num, bank_file_name_dict[bank_name], month_wanted, statement_directory)
bank_file_1.csv_to_xl()
#--------------------------------------------------
# find and load existing data file into bank_data
#--------------------------------------------------
if os.path.isfile(save_file):
    wb = load_workbook(save_file)
    ws = wb.active

    pass
    for row in list(ws.rows)[1:]:
        pass
        bank_data[row[0].value] = [c.value for c in row[1:]]
    bank_data.pop ("null", "not_found")
    ws["J2"] = "fuck"
    pass
else:  #create one
    wb = openpyxl.Workbook()
    ws = wb.active
    wb.save (filename=save_file)
    pass
pass



while bank_choice !=  "NONE" :
    pass
    total_debit = 0
    total_deposit = 0
    category_from_bank_data = 0
    debit = 0
    deposit = 0
    bank_file = Bank(bank_name_dict[bank_num], bank_num, bank_file_name_dict[bank_name],month_wanted, statement_directory)
    y = bank_file.get_wb_and_ws()  # tupple of file_name, wb and ws
    if y == None:
        print ("NO MONTH FOUND")
        exit(f" month: {month_wanted} not found")
    ws = y[2]   #work sheet
    wb = y[1]   #work book
    bank_file_name = y[0]
#    df_sheet_dict = pd.read_excel(bank_file_name, sheet_name=None)
#    df = pd.concat(df_sheet_dict.values(), ignore_index=True)#file name
    size = Bank.get_statement_length(bank_file_name)
    print (f"Pandas size: {size[4] + 1} ")

    new_amount = 0
    print (f"number of rows in statement is :{size[1]}")
#    print (type(category_sums))
#    print (category_sums)
    for tupple_new_row in ws.iter_rows(min_row=2, max_row=size[4]+1, min_col=0, max_col=9, values_only = True):
        new_row = list(tupple_new_row)
        if bank_name == "CAPITAL_ONE":
            if new_row[statement_headings_dict['AMOUNT'][bank_num]] == None:
                new_row[statement_headings_dict['AMOUNT'][bank_num]] =round(float(new_row[statement_headings_dict['PAYMENT'][bank_num]]),2)
            else:
                new_row[statement_headings_dict['AMOUNT'][bank_num]] = round(float ((new_row[statement_headings_dict['AMOUNT'][bank_num]])),2) * -1.00

        pass
        a_new_statement_row = New_statement_row( new_row, bank_name, bank_num)
 #       print (f"new_row->key is:{new_row}\n")
        exists_in_a_dictonary_or_not,the_statement_key,the_statement_amount,the_statement_date, category_from_bank_data = a_new_statement_row.is_it_new_to_bank_data(sbf_dict, new_row)
        pass
        if (exists_in_a_dictonary_or_not == "BANK_DATA EXACTLY SAME AS NEW STATEMENT"):
            continue
        if (exists_in_a_dictonary_or_not == "EXISTS IN BANK_DATA" ):

            print (f"{the_statement_key} of ${the_statement_amount} Is already in bank_data for  with category wich one??")
            if the_statement_amount > 0:
                pass
                credit = the_statement_amount
            if the_statement_amount <= 0:
#                debit = abs(the_statement_amount)
                debit = (the_statement_amount)
                total_deposit += deposit
                total_debit += debit


 #               category_sums[]
                #SHOULD ADD THE NEW STATEMENT AMOUNT TO THE VALUE IN BANK DATA!!! DO IT
                bank_data[the_statement_key][sbf_headings_dict["AMOUNT"]] += the_statement_amount

        if  exists_in_a_dictonary_or_not == "NEW ENTRY":
            a_new_sbf_dict_key = a_new_statement_row.new_sbf_dict()
            new_cat_inst = Add_categories()
            print (f"\nCategory for:{list(sbf_dict.keys())[-1]} in {bank_name_dict[bank_num]} for $ {new_row[statement_headings_dict['AMOUNT'][bank_num]]}\n")
            new_cat_inst.print_category_menu()
            cat_num, special_cat_num = new_cat_inst.get_category()
            cat_num_text_str = category_dict[cat_num]
            cat_sums = Master_sum(new_amount,cat_num,sbf_dict)
            debit, deposit, category = cat_sums.fix_new_amount_signs(a_new_sbf_dict_key)
            new_cat_inst.add_new_cat_to_sbf_dict(cat_num, sbf_dict, cat_num_text_str)
            total_debit += debit
            total_deposit += deposit
#            category_sums[cat_num] += debit
            category_sums_dict[str(cat_num)][0] += debit
            pass
        elif (exists_in_a_dictonary_or_not == "EXISTING ENTRY" and t_or_f_is_it_special != True):
            pass
            print ("Existing Entry")
            new_amount = new_row[statement_headings_dict["AMOUNT"][bank_num]]
            a_new_statement_row.update_amounts(sbf_dict[my_list[-1]], new_amount)
            debit, deposit, category = cat_sums.fix_new_amount_signs(a_new_sbf_dict_key, total)
            print (f" Existing Vendor will add {debit} to {a_new_sbf_dict_key}'s category which is: {cat_num}")
        pass
        if 20 <= special_cat_num <= 30:
            break
    s = Merge_and_save(save_file, sbf_dict)
    s.merge_data()
    s.save()  # line 145
    wb.close()


    if special_cat_num == int(23):
        exit()
        sys.exit()

    print("SAVED")
  #  close()

    more_stuff = input ("Hey buddy, do you want to do anthing else?\n yes or no:\n")
    accepted_strings = {'y', 'yes', 'Yes', 'YES'}
    if more_stuff in accepted_strings:
        more_input = input ("Guess a number i.e 3  ro 2")
        if int(more_input) == 2:
            print ("PRINTING BANK_DATA then continue")
            pass
            print(json.dumps(bank_data, indent=4))
            continue
        elif int(more_input) == 3:
            print ("PRINTING CATEGORY SUMS then continue")
            ms = Master_sum(0,0,0,0)
            ms.print_cat_sums()
            continue
        else:
            print ("you fucked up the number you entered")
        input ("print db or category sums or total debit and total income")
    print("CLOSED")

    bank_choice = Bank_Names.NONE.value
    exit()








