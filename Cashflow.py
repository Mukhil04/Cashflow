import csv
from datetime import date
import datetime
import pandas as pd
import numpy as np
import pandas as pd
from sqlalchemy import null


# ----------------------------------------------------- CONVERTING FILES INTO CSV ----------------------------------------------------- 

read_file = pd.read_excel ("Receivable.xlsx")
read_file.to_csv ("Receivable.csv", index = None, header=True)
read_file = pd.read_excel ("Payable.xlsx")
read_file.to_csv ("Payable.csv", index = None, header=True)

receiavble_csv = "Receivable.csv"
payable_csv = "Payable.csv"
output_csv = "Output.csv"
# ----------------------------------------------------- READING RECEIVABLE ENTRIES FROM CSV INTO LIST ------------------------------------

receivable_entries = []
receivable_and_payable_on_date = {}
ingram_payable_on_date = {}
ark_payable_on_date = {}
dict3 = {}
dict4 = {}
dates = []

bank_balance = input("Enter bank balance: ")
'''
yes_bank_balance = input ("Enter Yes bank balance: ")
incred_balance = input("Enter INCRED balance: ")
'''
non_ark_ingram_bank_balance = []

with open (receiavble_csv, 'r') as csvfile:
    csvreader = csv.reader(csvfile)
      
    # Ignoring first two rows
    next(csvreader)                                         
    next(csvreader)                                                
  
    # Extracting each data row one by one
    for row in csvreader:
        receivable_entries.append(row)


# ----------------------------------------------------- INITIALIZING DICTIONARIES WITH 60 DATES AS KEY ------------------------------------

today_date = date.today()

for i in range(1,61):
    today_date += datetime.timedelta(days=1)
    receivable_and_payable_on_date[today_date] = []
    ingram_payable_on_date[today_date] = []
    ark_payable_on_date [today_date] = []
    dict3[today_date] = [0, 0, 0, datetime.date(2100,1,1)]
    dict4[today_date] = 0
    dates.append(today_date)

# Given date as key into receivable_on_date it will return a list of all outstanding values for that day

for row in receivable_entries:
    due_date = row[3][:-9].split("-")
    if len(due_date) == 3:
        year = int(due_date[0])
        month = int(due_date[1])
        date1 = int(due_date[2])
        due_date = datetime.date(year, month, date1)
        if due_date in receivable_and_payable_on_date:
            due_amount = row[6].replace(',','')
            if due_amount.find('.') == -1:
                receivable_and_payable_on_date[due_date].append(int(due_amount))
            else:
                due_amount = due_amount[:due_amount.find('.')]
                receivable_and_payable_on_date[due_date].append(int(due_amount))

payable_entries = []
today = date.today()
with open (payable_csv, 'r') as csvfile:
    csvreader1 = csv.reader(csvfile)
      
    # extracting field names through first row
    next(csvreader1)
    next(csvreader1)
  
    # extracting each data row one by one
    for row in csvreader1:
        payable_entries.append(row)

# PAYABLE INCLUSION IN receivable_on_date THAT KEEPS TRACK OF ALL CASHFLOW OF THAT DATE

# INGRAM PAYABLE ALONE IN DICTIONARY ingram_dict THAT KEEPS TRACK OF PAYABLE TO INGRAM OF THAT DATE

# ARK PAYABLE ALONE IN DICTIONARY ark_dict THAT KEEPS TRACK OF PAYABLE TO ARK OF THAT DATE
for row in payable_entries:
    if row[0] == "ARK Infosolutions Private Limited" :
        due_date = row[9][:-9].split("-")
        if len(due_date) == 3:
            year = int(due_date[0])
            month = int(due_date[1])
            date1 = int(due_date[2])
            due_date = datetime.date(year, month, date1)
            if due_date in ark_payable_on_date:
                due_amount = row[7].replace(',','')
                if due_amount.find('.') == -1:
                    ark_payable_on_date[due_date].append(int(due_amount))
                else:
                    due_amount = due_amount[:due_amount.find('.')]
                    ark_payable_on_date[due_date].append(int(due_amount))

    elif row[0] == "Ingram Micro India Private Limited":
        due_date = row[9][:-9].split("-")
        if len(due_date) == 3:
            year = int(due_date[0])
            month = int(due_date[1])
            date1 = int(due_date[2])
            due_date = datetime.date(year, month, date1)
            if due_date in ingram_payable_on_date:
                due_amount = row[7].replace(',','')
                if due_amount.find('.') == -1:
                    ingram_payable_on_date[due_date].append(int(due_amount))
                else:
                    due_amount = due_amount[:due_amount.find('.')]
                    ingram_payable_on_date[due_date].append(int(due_amount))
    else:
        due_date = row[9][:-9].split("-")
        if len(due_date) == 3:
            year = int(due_date[0])
            month = int(due_date[1])
            date1 = int(due_date[2])
            due_date = datetime.date(year, month, date1)
            if due_date in receivable_and_payable_on_date:
                due_amount = row[7].replace(',','')
                if due_amount.find('.') == -1:
                    receivable_and_payable_on_date[due_date].append(-1*int(due_amount))
                else:
                    due_amount = due_amount[:due_amount.find('.')]
                    receivable_and_payable_on_date[due_date].append(-1*int(due_amount))


# OD CALCULATION WITHOUT INGRAM AND ARK PAYABLE
for key in receivable_and_payable_on_date.keys():
    non_ark_ingram_bank_balance.append(sum(receivable_and_payable_on_date[key]) + int(bank_balance))
    bank_balance = sum(receivable_and_payable_on_date[key]) + int(bank_balance)

ingram_increment = 0
on_day_ingram_payment = []
ingram_payment_balance = []


ark_increment = 0
on_day_ark_payment = []
ark_payment_balance = []
OD =[]


reddington_ingram_using_incred_yesbank = 0
reddington_ingram_using_od = 0

yesbank_incred_used = []
counter = 0


# CALCULATES ON DAY INGRAM PAYMENT AND TOTAL INGRAM PAYMENT AMOUNT TILL DATE

for key in ingram_payable_on_date.keys():
    ingram_payment_balance.append(sum(ingram_payable_on_date[key]) + ingram_increment)
    ingram_increment = sum(ingram_payable_on_date[key]) + ingram_increment
    on_day_ingram_payment.append(sum(ingram_payable_on_date[key]))

# CALCULATES ON DAY ARK PAYMENT AND TOTAL ARK PAYMENT AMOUNT TILL DATE

for key in ark_payable_on_date.keys():
    ark_payment_balance.append(sum(ark_payable_on_date[key]) + ark_increment)
    ark_increment = sum(ark_payable_on_date[key]) + ark_increment
    on_day_ark_payment.append(sum(ark_payable_on_date[key]))


# OD CALCULATION

for i in range(60):
    OD.append(non_ark_ingram_bank_balance_data[i]-ingram_payment_balance[i]-ark_payment_balance[i])




fields3 = ["Date", "OD", "Ingram on day payment", "Total ingram payment", "Ark on day payment", "Total Ark payment", "Total yesbank and incred used", "Ingram loan today", "Ark loan today", "Date that causes the money to be borrwed today", "Repay laon?"]
with open(filename2, 'w') as csvfile: 
    # parsing each column of a row
    csvwriter = csv.writer(csvfile) 
        
    # writing the fields 
    csvwriter.writerow(fields3) 
    for i in range(60):
        if (OD[i]) < -47500000:
            yesbank_incred_used.append(-1*OD[i])
        elif (OD[i]) < 0:
            yesbank_incred_used.append((-1*OD[i]))
        else:
            yesbank_incred_used.append(0)
        
    # writing the data rows 
    for i in range(60):
        if i != 0:                                                                                  # MAYBE CHECK FOR i == 0
            counter = i
            number_of_days_before = 0
            if yesbank_incred_used[i-1] < yesbank_incred_used[i]:
                val = yesbank_incred_used[i]
                if (ingram_payment_balance[i] >= 40000000) and (val > 40000000):
                    # Cannot be covered by ingram alone
                    # four cases for ARK
                    ingram_val = 40000000
                    ark_val = val - ingram_val
                    
                    while ingram_val > 0 and counter >= 0:
                        if ingram_val <= on_day_ingram_payment[counter]:
                            if dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][0] > ingram_val:
                                pass
                            else:
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][0] = ingram_val 
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                            ingram_val = 0
                        else:
                            dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][0] = on_day_ingram_payment[counter]
                            dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                            number_of_days_before += 1
                            ingram_val -= on_day_ingram_payment[counter]
                            counter -= 1


                    counter = i
                    number_of_days_before = 0
                    if (ark_val <= 7500000) and (ark_payment_balance[i] >= ark_val):
                        while ark_val > 0 and counter >= 0:
                            if ark_val <= on_day_ark_payment[counter]:
                                if dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] > ark_val:
                                    pass
                                else:
                                    dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] = ark_val 
                                    dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                                ark_val = 0
                            else:
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] = on_day_ark_payment[counter]
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                                number_of_days_before += 1
                                ark_val -= on_day_ark_payment[counter]
                                counter -= 1
                        # Ark payment enough to cover rest and no debt


                    elif (ark_val <= 7500000) and (ark_payment_balance[i] < ark_val):
                        dict4[today+datetime.timedelta(days=i+1)] = max(dict4[today+datetime.timedelta(days=i+1)],ark_val - ark_payment_balance[i])
                        ark_val = ark_payment_balance[i]
                        while ark_val > 0 and counter >= 0 :
                            if ark_val <= on_day_ark_payment[counter]:
                                if dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] > ark_val:
                                    pass
                                else:
                                    dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] = ark_val 
                                    dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                                ark_val = 0
                            else:
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] = on_day_ark_payment[counter]
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                                number_of_days_before += 1
                                ark_val -= on_day_ark_payment[counter]
                                counter -= 1
                        
                        # Ark payment not enough to cover rest and yes debt
                    
                    elif (ark_val > 7500000) and (ark_payment_balance[i] >= 7500000):
                        dict4[today+datetime.timedelta(days=i+1)] = max(dict4[today+datetime.timedelta(days=i+1)],ark_val - 7500000)
                        ark_val = 7500000
                        while ark_val > 0 and counter >= 0:
                            if ark_val <= on_day_ark_payment[counter]:
                                if dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] > ark_val:
                                    pass
                                else:
                                    dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] = ark_val 
                                    dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                                ark_val = 0
                            else:
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] = on_day_ark_payment[counter]
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                                number_of_days_before += 1
                                ark_val -= on_day_ark_payment[counter]
                                counter -= 1
                        
                        # Ark payment not enough but full ark loan can be used

                    elif (ark_val > 7500000) and (ark_payment_balance[i] < 7500000):
                        dict4[today+datetime.timedelta(days=i+1)] = max(dict4[today+datetime.timedelta(days=i+1)],ark_val - ark_payment_balance[i])
                        ark_val = ark_payment_balance[i]
                        while ark_val > 0 and counter >= 0:
                            if ark_val <= on_day_ark_payment[counter]:
                                if dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] > ark_val:
                                    pass
                                else:
                                    dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] = ark_val 
                                    dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                                ark_val = 0
                            else:
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] = on_day_ark_payment[counter]
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                                number_of_days_before += 1
                                ark_val -= on_day_ark_payment[counter]
                                counter -= 1
                        # Ark payment not enough but only part of ark loan can be used
                    
                

                elif (ingram_payment_balance[i] >= 40000000) and (val <= 40000000 ):
                    # No ark
                    while val > 0 and counter >= 0:
                        if val < on_day_ingram_payment[counter]:
                            if dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][0] > val:
                                pass
                            else:
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][0] = val 
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                            val = 0
                        else:
                            dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][0] = on_day_ingram_payment[counter]
                            dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                            number_of_days_before += 1
                            val -= on_day_ingram_payment[counter]
                            counter -= 1



                elif (ingram_payment_balance[i] >= 30000000) and (val > ingram_payment_balance[i]):
                    # Cannot be covered by ingram alone
                    # four cases for ARK

                    ingram_val = ingram_payment_balance[i]
                    current_ingram_loan = ingram_val
                    ark_val = val - ingram_val
                    while ingram_val > 0 and counter >= 0:
                        if ingram_val <= on_day_ingram_payment[counter]:
                            if dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][0] > ingram_val:
                                pass
                            else:
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][0] = ingram_val 
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                            ingram_val = 0
                        else:
                            dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][0] = on_day_ingram_payment[counter]
                            dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                            number_of_days_before += 1
                            ingram_val -= on_day_ingram_payment[counter]
                            counter -= 1

                    counter = i
                    number_of_days_before = 0
                    if (ark_val <= 7500000) and (ark_payment_balance[i] >= ark_val):
                        while ark_val > 0 and counter >= 0:
                            if ark_val <= on_day_ark_payment[counter]:
                                if dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] > ark_val:
                                    pass
                                else:
                                    dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] = ark_val 
                                    dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                                ark_val = 0
                            else:
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] = on_day_ark_payment[counter]
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                                number_of_days_before += 1
                                ark_val -= on_day_ark_payment[counter]
                                counter -= 1
                        # Ark payment enough to cover rest and no debt


                    elif (ark_val <= 7500000) and (ark_payment_balance[i] < ark_val):
                        dict4[today+datetime.timedelta(days=i+1)] = max(dict4[today+datetime.timedelta(days=i+1)],ark_val - ark_payment_balance[i])
                        ark_val = ark_payment_balance[i]
                        while ark_val > 0 and counter >= 0 :
                            if ark_val <= on_day_ark_payment[counter]:
                                if dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] > ark_val:
                                    pass
                                else:
                                    dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] = ark_val 
                                    dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                                ark_val = 0
                            else:
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] = on_day_ark_payment[counter]
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                                number_of_days_before += 1
                                ark_val -= on_day_ark_payment[counter]
                                counter -= 1
                        
                        # Ark payment not enough to cover rest and yes debt
                    
                    elif (ark_val > 7500000) and (ark_payment_balance[i] >= 7500000):
                        dict4[today+datetime.timedelta(days=i+1)] = max(dict4[today+datetime.timedelta(days=i+1)],ark_val - 7500000)
                        ark_val = 7500000
                        while ark_val > 0 and counter >= 0:
                            if ark_val <= on_day_ark_payment[counter]:
                                if dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] > ark_val:
                                    pass
                                else:
                                    dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] = ark_val 
                                    dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                                ark_val = 0
                            else:
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] = on_day_ark_payment[counter]
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                                number_of_days_before += 1
                                ark_val -= on_day_ark_payment[counter]
                                counter -= 1
                        
                        # Ark payment not enough but full ark loan can be used

                    elif (ark_val > 7500000) and (ark_payment_balance[i] < 7500000):
                        dict4[today+datetime.timedelta(days=i+1)] = max(dict4[today+datetime.timedelta(days=i+1)],ark_val - ark_payment_balance[i])
                        ark_val = ark_payment_balance[i]
                        while ark_val > 0 and counter >= 0:
                            if ark_val <= on_day_ark_payment[counter]:
                                if dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] > ark_val:
                                    pass
                                else:
                                    dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] = ark_val 
                                    dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                                ark_val = 0
                            else:
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] = on_day_ark_payment[counter]
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                                number_of_days_before += 1
                                ark_val -= on_day_ark_payment[counter]
                                counter -= 1
                        # Ark payment not enough but only part of ark loan can be used

                elif (ingram_payment_balance[i] >= 30000000) and (val <= ingram_payment_balance[i]):
                    # No ark

                    while val > 0 and counter >= 0:
                        if val < on_day_ingram_payment[counter]:
                            if dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][0] > val:
                                pass
                            else:
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][0] = val 
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                            val = 0
                        else:
                            dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][0] = on_day_ingram_payment[counter]
                            dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                            number_of_days_before += 1
                            val -= on_day_ingram_payment[counter]
                            counter -= 1
                    pass

                elif (ingram_payment_balance[i] < 30000000) and (val > ingram_payment_balance[i]):
                    # Cannot be covered by ingram alone
                    # four cases for ARK

                    ingram_val = ingram_payment_balance[i]
                    ark_val = val - ingram_val
                    
                    while ingram_val > 0 and counter >= 0:
                        if ingram_val <= on_day_ingram_payment[counter]:
                            if dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][0] > ingram_val:
                                pass

                            else:
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][0] = ingram_val 
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                            ingram_val = 0
                        else:
                            dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][0] = on_day_ingram_payment[counter]
                            dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                            number_of_days_before += 1
                            ingram_val -= on_day_ingram_payment[counter]
                            counter -= 1

                    counter = i
                    number_of_days_before = 0
                    if (ark_val <= 7500000) and (ark_payment_balance[i] >= ark_val):
                        while ark_val > 0 and counter >= 0:
                            if ark_val <= on_day_ark_payment[counter]:
                                if dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] > ark_val:
                                    pass
                                else:
                                    dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] = ark_val 
                                    dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                                ark_val = 0
                            else:
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] = on_day_ark_payment[counter]
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                                number_of_days_before += 1
                                ark_val -= on_day_ark_payment[counter]
                                counter -= 1
                        # Ark payment enough to cover rest and no debt


                    elif (ark_val <= 7500000) and (ark_payment_balance[i] < ark_val):
                        dict4[today+datetime.timedelta(days=i+1)] = max(dict4[today+datetime.timedelta(days=i+1)],ark_val - ark_payment_balance[i])
                        ark_val = ark_payment_balance[i]
                        while ark_val > 0 and counter >= 0 :
                            if ark_val <= on_day_ark_payment[counter]:
                                if dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] > ark_val:
                                    pass
                                else:
                                    dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] = ark_val 
                                    dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                                ark_val = 0
                            else:
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] = on_day_ark_payment[counter]
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                                number_of_days_before += 1
                                ark_val -= on_day_ark_payment[counter]
                                counter -= 1
                        
                        # Ark payment not enough to cover rest and yes debt
                    
                    elif (ark_val > 7500000) and (ark_payment_balance[i] >= 7500000):
                        dict4[today+datetime.timedelta(days=i+1)] = max(dict4[today+datetime.timedelta(days=i+1)],ark_val - 7500000)
                        ark_val = 7500000
                        while ark_val > 0 and counter >= 0:
                            if ark_val <= on_day_ark_payment[counter]:
                                if dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] > ark_val:
                                    pass
                                else:
                                    dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] = ark_val 
                                    dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                                ark_val = 0
                            else:
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] = on_day_ark_payment[counter]
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                                number_of_days_before += 1
                                ark_val -= on_day_ark_payment[counter]
                                counter -= 1
                        
                        # Ark payment not enough but full ark loan can be used

                    elif (ark_val > 7500000) and (ark_payment_balance[i] < 7500000):
                        dict4[today+datetime.timedelta(days=i+1)] = max(dict4[today+datetime.timedelta(days=i+1)],ark_val - ark_payment_balance[i])
                        ark_val = ark_payment_balance[i]
                        while ark_val > 0 and counter >= 0:
                            if ark_val <= on_day_ark_payment[counter]:
                                if dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] > ark_val:
                                    pass
                                else:
                                    dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] = ark_val 
                                    dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                                ark_val = 0
                            else:
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][1] = on_day_ark_payment[counter]
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                                number_of_days_before += 1
                                ark_val -= on_day_ark_payment[counter]
                                counter -= 1
                        # Ark payment not enough but only part of ark loan can be used
                    

                elif (ingram_payment_balance[i] < 30000000) and (val < ingram_payment_balance[i]):
                    # No ark
                    while val > 0 and counter >= 0:
                        if val < on_day_ingram_payment[counter]:
                            if dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][0] > val:
                                pass
                            else:
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][0] = val 
                                dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                            val = 0
                        else:
                            dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][0] = on_day_ingram_payment[counter]
                            dict3[today+datetime.timedelta(days=i+1)-datetime.timedelta(days=number_of_days_before)][3] = today+datetime.timedelta(days=i+1)
                            number_of_days_before += 1
                            val -= on_day_ingram_payment[counter]
                            counter -= 1
            

                    

            
    for i in range(60):
        if dict3[today+datetime.timedelta(days=i+1)][3] == datetime.date(2100,1,1) and yesbank_incred_used[i] < 47500000 and yesbank_incred_used[i] > 0:
            csvwriter.writerow([dates[i], "{:,}".format(OD[i]), "{:,}".format(on_day_ingram_payment[i]), "{:,}".format(ingram_payment_balance[i]), "{:,}".format(on_day_ark_payment[i]), "{:,}".format(ark_payment_balance[i]), yesbank_incred_used[i], dict3[today+datetime.timedelta(days=i+1)][0], dict3[today+datetime.timedelta(days=i+1)][1], dict3[today+datetime.timedelta(days=i+1)][3], "Repay Incred/yesbank with bank balance"])
        else:
            csvwriter.writerow([dates[i], "{:,}".format(OD[i]), "{:,}".format(on_day_ingram_payment[i]), "{:,}".format(ingram_payment_balance[i]), "{:,}".format(on_day_ark_payment[i]), "{:,}".format(ark_payment_balance[i]), yesbank_incred_used[i], dict3[today+datetime.timedelta(days=i+1)][0], dict3[today+datetime.timedelta(days=i+1)][1], dict3[today+datetime.timedelta(days=i+1)][3], "Don't repay"])


filename3 = "Received_file.csv"
fields3 = ["Date", "OD", "Ingram loan today", "ARK loan today", "Repay?"]
with open(filename3, 'w') as csvfile: 
    # parsing each column of a row
    csvwriter = csv.writer(csvfile) 
        
    # writing the fields 
    csvwriter.writerow(fields3) 

    for i in range(60):
        if dict3[today+datetime.timedelta(days=i+1)][3] == datetime.date(2100,1,1) and yesbank_incred_used[i] < 47500000 and yesbank_incred_used[i] > 0:
            csvwriter.writerow([dates[i], "{:,}".format(OD[i]), dict3[today+datetime.timedelta(days=i+1)][0], dict3[today+datetime.timedelta(days=i+1)][1], "Repay Incred/yesbank with bank balance"])
        else:
            csvwriter.writerow([dates[i], "{:,}".format(OD[i]),dict3[today+datetime.timedelta(days=i+1)][0], dict3[today+datetime.timedelta(days=i+1)][1], "Don't repay"])


df_new = pd.read_csv('Output.csv')
df_new1 = pd.read_csv('Received_file.csv')
    # saving xlsx file
GFG = pd.ExcelWriter('Output.xlsx')
GFG1 = pd.ExcelWriter('Received_file.xlsx')
df_new.to_excel(GFG, index = False)
df_new.to_excel(GFG1, index = False)

    
GFG.save()
GFG1.save()


'''
Write short code proprly not as easy as keeping track of 
current_ark_loan and current_ingram_loan within the yes_bank_incred_used[i-1] > yes_bank_incred_used[i] if condition
as yes_bank_incred_used[i] > yes_bank_incred_used[i-1] could happen and make OD positive at which point we should not have any loans from either
'''

'''
Need column that keeps track of current ingram and ark loan
Also answer the question should I pay upfront
'''


'''
When deciding when to split reddington and ingram payment into OD or yes bank and incred

    1. When graph has positive slope,
       Calculate OD and yesbank and incred split until the days in the past necessary for all such days
       Decide payment split on xth day giving first priority to payment that requires more amount from yesbank and incred.
       If the amonut required from yesbank and incred is the same on a day give priority to the split determined by the payment from a later date
       
    2. If a date has no split payment then fill yesbank and incred with OD money as much as possible, ideally 0 yesbank and incred
'''

'''
INCRED: INGRAM- 1CRORES, ARK -75 LAKHS
YESBANK- INGRAM- 3 CRORES

ingram yesbank
ingram INCRED
ARK INCRED
'''