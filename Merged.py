import os
import csv
import requests
import pandas as pd
from datetime import datetime
from dateutil import parser
from dateutil.relativedelta import relativedelta
from dateutil.rrule import MO, TU, WE, TH, FR, SA, SU
from urllib.parse import urlencode,urljoin
from urllib.parse import unquote
from urllib.parse import quote
from urllib.parse import quote_plus

def getCardsNumbers():
    accounts = pd.read_csv("accounts.csv")
    login= ''
    password= ''
    directory='.\CardsNumbers'
    if not os.path.exists(directory):
        os.makedirs(directory)
    for index, row in accounts.iterrows():
        login=row['username']
        password= row['password']
        # Step 1: Log in to the website
        login_url = 'https://fleetcard.vivoenergy.com/WP/v1/login'
        payload = {
            'login': login,
            'password': password,
            }
        session = requests.Session()
        session.post(login_url, data=payload)
        # Step 2: Get RAW data

        file_url = 'https://fleetcard.vivoenergy.com/WP/v1/card-list'
        query_params = {
            "refresh": "true",
            "itemsPerPage": 0,
            "currentPage": 0,
            "idCardStatus": "",
            "purseList": "",
            "query": "",
            "depAndCC": ""
        }

        data = (session.get(file_url,params=query_params)).json()
        # Extract the relevant data from the dictionary object
        rows = [[data['cardNumber'], data['holder']['name'],data['holder']['carPlateNumber']] for data in data['cardList']]

        # Write the data to a CSV file

        with open(f'.\CardsNumbers\CardsInAccount {index+1}.csv', 'w', newline='') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(['Card number', 'Vehicle registration number' , 'Card holder'])
            writer.writerow(['Acoount :' , login , password])
            writer.writerows(rows)

def downloadAllData():
    accounts = pd.read_csv("accounts.csv")
    for index, row in accounts.iterrows():
        # Step 1: Log in to the website
        start_date= (datetime.now() + relativedelta(day=1,months=-1,hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S')
        end_date= (datetime.now() + relativedelta(day=1,days=-1,hour=23,minute=59,second=59)).strftime('%Y-%m-%d %H:%M:%S')
        login_url = 'https://fleetcard.vivoenergy.com/WP/v1/login'
        payload = {
            'login': row['username'],
            'password': row['password'],
            }
        session = requests.Session()
        session.post(login_url, data=payload)

        # Step 2: Download the excel file
        print(f'Download for Account {index+1} Started.') 
        file_url = 'https://fleetcard.vivoenergy.com/WP/v1/reports/card-purchase-list'
        query_params = {
            "refresh": "true",
            "itemsPerPage": 25,
            "currentPage": 1,
            "startDate": start_date,
            "endDate": end_date,
            "isProcessing": 1,
            "cardList": "",
            "service": "",
            "pos": "",
            "invoiceNumber": "",
            "format": "CSV",
            "additionalColumns": "consumption,holder,plate-nr"
        }

        response=session.get(file_url, params=unquote(urlencode(query_params)))
        with open('downloadingData.csv', 'a') as f:
            f.write(response.text)
        print(f'Download for Account {index+1} is done')    
    df = pd.read_csv('downloadingData.csv')
    value_index = df[df['Transaction date'] == 'Transaction date'].index
    df = df.drop(value_index)
    df.to_csv('downloadingData.csv', index=False)
    os.rename('downloadingData.csv', "AllCardPurchasesIn{}.csv".format((datetime.now() + relativedelta(months=-1)).strftime('%b')))    
    print('All Purchases Data From All Accounts Download is Done.')


def downloadStcData():
    # Step 1: Log in to the website
    start_date= (datetime.now() + relativedelta(day=1, months=-1,hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S')
    end_date= (datetime.now() + relativedelta(day=1,days=-1,hour=23,minute=59,second=59)).strftime('%Y-%m-%d %H:%M:%S')
    login_url = 'https://fleetcard.vivoenergy.com/WP/v1/login'
    payload = {
        'login': '',
        'password': '',
        }
    session = requests.Session()
    session.post(login_url, data=payload)

    # Step 2: Download the excel file
    print('STC Cards Purchases Data Download Started.')
    file_url = 'https://fleetcard.vivoenergy.com/WP/v1/reports/card-purchase-list'
    query_params = {
        "refresh": "true",
        "itemsPerPage": 25,
        "currentPage": 1,
        "startDate": start_date,
        "endDate": end_date,
        "isProcessing": 1,
        "cardList": "",
        "service": "",
        "pos": "",
        "invoiceNumber": "",
        "format": "CSV",
        "additionalColumns": "consumption,holder,plate-nr"
    }
    response=session.get(file_url, params=unquote(urlencode(query_params)))
    with open('downloadingData.xls', 'a') as f:
        f.write(response.text)   
    os.rename('downloadingData.xls', "STCPurchasesIn{}.xls".format((datetime.now() + relativedelta(months=-1)).strftime('%b')))
    print('STC Cards Purchases Data Download is Done.')


def getUserNPass(cardNumber):
    dir_path = '.\CardsNumbers'
    csv_files = [f for f in os.listdir(dir_path) if f.endswith('.csv')]
    for csv_file in csv_files:
        df = pd.read_csv(os.path.join(dir_path, csv_file))
        if cardNumber in df.values:
            return df.iloc[0][1] , df.iloc[0][2]
    return None


def checkAccount():
    print("Choose an Account: \n1. maher.haddad@ranamotors.com\n2. maher.haddad@ranamotors.com1\n3. maher.haddad@ranamotors.com3\n4. maher.haddad@ranamotors.com4\n5. maher.haddad@ranamotors.com5\n6. maher.haddad@ranamotors.com8\n7. maher.haddad@ranamotors.com10\n8. maher.haddad@ranamotors.com14\n9. maher@ogroupafrica.com")

    choice = int(input("Enter your choice [1-9]: "))

    if choice == 1:
        account= 'maher.haddad@ranamotors.com'

    elif choice == 2:
        account= 'maher.haddad@ranamotors.com1'

    elif choice == 3:
        account= 'maher.haddad@ranamotors.com3'

    elif choice == 4:
        account= 'maher.haddad@ranamotors.com4'

    elif choice == 5:
        account= 'maher.haddad@ranamotors.com5'

    elif choice == 6:
        account= 'maher.haddad@ranamotors.com8'

    elif choice == 7:
        account= 'maher.haddad@ranamotors.com10'

    elif choice == 8:
        account= 'maher.haddad@ranamotors.com14'

    elif choice == 9:
        account= 'maher@ogroupafrica.com'

    else:
        return None

    dir_path = '.\CardsNumbers'
    csv_files = [f for f in os.listdir(dir_path) if f.endswith('.csv')]
    for csv_file in csv_files:
        df = pd.read_csv(os.path.join(dir_path, csv_file))
        if account in df.values:
            return df.iloc[0][1] , df.iloc[0][2]
    


def getUserDateChoice():
    print("Choose an option:\n1. Custom date and time\n2. This week\n3. Last week\n4. This month\n5. Last month\n6. Go back X days from today\n7. Go back X weeks from this week\n8. Choose a specific month(current year)")
    userInput=int(input("Enter your choice: "))
    # Custom date and time
    if userInput == 1:           

        while True:
            while True:
                try:
                    start_date= parser.parse(input("Enter Start Date and Time: ")).strftime('%Y-%m-%d %H:%M:%S')
                    break
                except ValueError:
                    print("Invalid input, enter a valid date and time.")
            while True:
                try:
                    end_date= parser.parse(input("Enter End Date and Time: ")).strftime('%Y-%m-%d %H:%M:%S')
                    break
                except ValueError:
                    print("Invalid input, please enter a valid date and time.")
            if(start_date < end_date):
                return start_date , end_date
                break 
            print("Invalid input, end date must be after start date, enter a valid date and time")

    # This week
    elif userInput == 2:
        start_date= (datetime.now() + relativedelta(weekday=MO(-1),hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S')
        end_date= (datetime.now()).strftime('%Y-%m-%d %H:%M:%S')
        return start_date , end_date

    # Last week
    elif userInput == 3:
        start_date= (datetime.now() + relativedelta(weekday=MO(-2),hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S')
        end_date= (datetime.now() + relativedelta(weekday=SU(-1))).strftime('%Y-%m-%d %H:%M:%S')
        return start_date, end_date

    # This month
    elif userInput == 4:
        start_date= (datetime.now() + relativedelta(day=1,hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S')
        end_date= (datetime.now()).strftime('%Y-%m-%d %H:%M:%S')
        return start_date , end_date

    # Last month
    elif userInput == 5:
        start_date= (datetime.now() + relativedelta(day=1, months=-1,hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S')
        end_date= (datetime.now() + relativedelta(day=1,days=-1)).strftime('%Y-%m-%d %H:%M:%S')
        return start_date , end_date

    # Go back X days from today
    elif userInput == 6:
        start_date= (datetime.now() + relativedelta(days=-int(input("enter how many days back: ")),hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S')
        end_date= (datetime.now()).strftime('%Y-%m-%d %H:%M:%S')
        return start_date , end_date

    # Go back X weeks from this week
    elif userInput == 7:
        start_date= (datetime.now() + relativedelta(weekday=MO(-1), weeks=-int(input("enter how many weeks back: ")),hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S')
        end_date= (datetime.now() + relativedelta(weekday=SU(-1))).strftime('%Y-%m-%d %H:%M:%S')
        return start_date , end_date

    # Choose a specific month(current year)
    elif userInput == 8:
        month=int(input("enter month number: "))
        start_date= (datetime.now() + relativedelta(day=1, month=month,hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S')
        end_date= (datetime.now() + relativedelta(month=month+1,day=1,days=-1, hour=23,minute=59,second=59)).strftime('%Y-%m-%d %H:%M:%S')
        return start_date , end_date



def download_excel_file(login,password,start_date, end_date,card=''):
    # Step 1: Log in to the website
    login_url = 'https://fleetcard.vivoenergy.com/WP/v1/login'
    payload = {
        'login': login,
        'password': password,
        }
    session = requests.Session()
    session.post(login_url, data=payload)

    # Step 2: Download the excel file
    file_url = 'https://fleetcard.vivoenergy.com/WP/v1/reports/card-purchase-list'
    query_params = {
        "refresh": "true",
        "itemsPerPage": 25,
        "currentPage": 1,
        "startDate": start_date,
        "endDate": end_date,
        "isProcessing": 1,
        "cardList": "",
        "service": "",
        "pos": "",
        "invoiceNumber": "",
        "format": "CSV",
        "additionalColumns": "consumption,holder,plate-nr"
    }
    response=session.get(file_url, params=unquote(urlencode(query_params)))
    with open(f'{card}CardPurchases.csv', 'wb') as f:
        f.write(response.content)


#Main Code

if os.path.exists('.\CardsNumbers'):
    for i in range(9):
        if os.path.isfile(f'.\CardsNumbers\CardsInAccount {i+1}.csv'):
            pass
        else:
            getCardsNumbers()
else:
    getCardsNumbers()


while True:

    print("1. Download All Purchases Data From All Accounts.\n2. Download STC Cards Purchases.\n3. Enter a Card number.\n4. Choose an Account.\n5. Enter a Card number to identify it's account. \n6. Update Cards Numbers. \n7. Exit the program")
    choice = int(input("Enter your choice [1-7]: "))

    if choice == 1:
        downloadAllData()
    elif choice == 2:
        downloadStcData()
    elif choice == 3:
        while True:
            try:
                card = input("Enter your card number: ")
                login , password= getUserNPass(card)
                if login and password:
                    break
            except:
                print("Invalid card number. Try again.")    
        start_date, end_date= getUserDateChoice()
        download_excel_file(login,password,start_date, end_date,card)
    elif choice == 4:
        while True:
            try:
                login = checkAccount()
                if login[0]:
                    break
            except:
                print("Invalid account choice. Try again.")    
        start_date, end_date= getUserDateChoice()
        download_excel_file(login,password,start_date, end_date)
    elif choice == 5:
        while True:
            try:
                card = input("Enter your card number: ")
                login = getUserNPass(card)
                if login[0]:
                    break
            except:
                print("Invalid card number. Try again.")
        print(f'This card number belongs to this account:\n {login[0]}')
    elif choice == 6:
        getCardsNumbers()
    elif choice == 7:
        break    
    else:
        print("Invalid option. Try again.")