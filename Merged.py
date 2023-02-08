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

def getCardsNumbers():
    accounts = pd.read_csv("Credentials.csv")
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


        try:
            # code to attempt the login
            session = requests.Session()
            response = session.post(login_url, data=payload)
            if response.status_code != 200:
                raise Exception("Login failed, Problem with credentials")
        except requests.exceptions.ConnectionError:
            # handle the error when the connection is lost
            print("Error: Connection lost. Please check your internet connection and try again.")
            return
        except requests.exceptions.Timeout as e:
            print("The request timed out:", e)
            return
        except Exception as e:
            # handle the error if there is a problem with the credentials
            print("Error:", e)
            return


            # Step 2: Getting data

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
        rows = [[index,data['cardNumber'],data['holder']['carPlateNumber'], data['holder']['name']] for data in data['cardList']]

        # Write the data to a CSV file
        name=accounts.loc[index]['Account Name']
        with open(f'{directory}\CardsInAccount {name}.csv', 'w', newline='') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(['id','Card number', 'Vehicle registration number' , 'Card holder'])
            writer.writerows(rows)

def downloadAllData():
    accounts = pd.read_csv("Credentials.csv")
    for index, row in accounts.iterrows():
        # Step 1: Log in to the website
        start_date= parser.parse((datetime.now() + relativedelta(day=1,months=-1,hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S'))
        end_date= parser.parse((datetime.now() + relativedelta(day=1,days=-1,hour=23,minute=59,second=59)).strftime('%Y-%m-%d %H:%M:%S'))
        login_url = 'https://fleetcard.vivoenergy.com/WP/v1/login'
        payload = {
            'login': row['username'],
            'password': row['password'],
            }
        try:
            # Login attempt
            session = requests.Session()
            response = session.post(login_url, data=payload)
            if response.status_code != 200:
                raise Exception("Login failed, Problem with credentials")
        except requests.exceptions.ConnectionError:
            # handle the error when the connection is lost
            print("Error: Connection lost. Please check your internet connection and try again.")
            return
        except requests.exceptions.Timeout as e:
            print("The request timed out:", e)
            return
        except Exception as e:
            # handle the error if there is a problem with the credentials
            print("Error:", e)
            return

        # Step 2: Download the excel file
        accountName = accounts.loc[index]['Account Name']
        print(f'Download for Account {accountName} Started.') 
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
        print(f'Download for Account {accountName} is done')    
    df = pd.read_csv('downloadingData.csv')
    value_index = df[df['Transaction date'] == 'Transaction date'].index
    df = df.drop(value_index)
    df.to_excel(f'All Card Purchases Report {start_date.strftime("%b %d")} - {end_date.strftime("%b %d")}.xls',engine='openpyxl', index=False)
    os.remove('downloadingData.csv')
    print('All Purchases Data From All Accounts Download is Done.')


def downloadStcData():
    # Step 1: Log in to the website
    start_date= parser.parse((datetime.now() + relativedelta(day=1,months=-1,hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S'))
    end_date= parser.parse((datetime.now() + relativedelta(day=1,days=-1,hour=23,minute=59,second=59)).strftime('%Y-%m-%d %H:%M:%S'))
    accountsInfo= getUserNPass('WEST AFRICA TIRE SERV. LTD')
    login = accountsInfo[0]
    password = accountsInfo[1]
    login_url = 'https://fleetcard.vivoenergy.com/WP/v1/login'
    payload = {
        'login': login,
        'password': password,
        }
    try:
        # code to attempt the login
        session = requests.Session()
        response = session.post(login_url, data=payload)
        if response.status_code != 200:
            raise Exception("Login failed, Problem with credentials")
    except requests.exceptions.ConnectionError:
        # code to handle the error when the connection is lost
        print("Error: Connection lost. Please check your internet connection and try again.")
        return
    except requests.exceptions.Timeout as e:
        print("The request timed out:", e)
        return 
    except Exception as e:
        # handle the error if there is a problem with the credentials
        print("Error:", e)
        return

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
        "format": "XLS",
        "additionalColumns": "consumption,holder,plate-nr"
    }
    print('STC Cards Purchases Data Download Started.')
    response=session.get(file_url, params=unquote(urlencode(query_params)))
    with open(f'STC Report {start_date.strftime("%b %d")} - {end_date.strftime("%b %d")}.xls', 'wb') as f:
        f.write(response.content)   
    print('STC Cards Purchases Data Download is Done.')


def getCredentialsId(cardNumber):
    dir_path = '.\CardsNumbers'
    csv_files = [f for f in os.listdir(dir_path) if f.endswith('.csv')]
    for csv_file in csv_files:
        df = pd.read_csv(os.path.join(dir_path, csv_file))
        if cardNumber in df.values:
            return df.loc[0]['id']
    return None


def getUserNPass(choice):

    if choice == 1:
        account= 'RANA MOTORS'

    elif choice == 2:
        account= 'B.B.C INDUSTRIALS CO(GH) LTD'

    elif choice == 3:
        account= 'LAJJIMARK CO. LTD.'

    elif choice == 4 or choice == 'WEST AFRICA TIRE SERV. LTD':
        account= 'WEST AFRICA TIRE SERV. LTD'

    elif choice == 5:
        account= 'HIGHLAND SPRINGS (GH) LTD'

    elif choice == 6:
        account= 'KHOMARA PRINTING PRESS LTD'

    elif choice == 7:
        account= 'ODAYMAT INVESTMENTS LTD'

    elif choice == 8:
        account= 'RANA ATLAS'

    elif choice == 9:
        account= 'ELDACO'

    elif choice == 0:
        return 0,0;

    else:
        return None

    accounts = pd.read_csv("Credentials.csv")
    if account in accounts.values:
        index = accounts[accounts['Account Name'] == account].index[0]
        id= accounts.loc[index]['id']
        return accounts.loc[id]['username'],accounts.loc[id]['password'],accounts.loc[id]['Account Name']
    


def getUserDateChoice():
    print("Choose an option:\n1. Custom date and time\n2. This week\n3. Last week\n4. This month\n5. Last month\n6. Go back X days from today\n7. Go back X weeks from this week\n8. Choose a specific month(current year)\n\n  0. Exit")
    userInput=int(input("Enter your choice: "))
    # Custom date and time
    if userInput == 1:           

        while True:
            while True:
                try:
                    startDateInput=input("Enter Start Date and Time(0 to exit): ")
                    if startDateInput == '0':
                        return 0,0
                    else:
                        start_date= parser.parse(startDateInput).strftime('%Y-%m-%d %H:%M:%S')
                        break
                except ValueError:
                        print("Invalid input, enter a valid date and time.")
            while True:
                try:
                    endDateInput=input("Enter End Date and Time(0 to exit): ")
                    if endDateInput == '0':
                        return 0,0
                    else:
                        end_date= parser.parse(endDateInput).strftime('%Y-%m-%d %H:%M:%S')
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
        return parser.parse(start_date) , parser.parse(end_date)

    # Last week
    elif userInput == 3:
        start_date= (datetime.now() + relativedelta(weekday=MO(-2),hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S')
        end_date= (datetime.now() + relativedelta(weekday=SU(-1))).strftime('%Y-%m-%d %H:%M:%S')
        return parser.parse(start_date) , parser.parse(end_date)

    # This month
    elif userInput == 4:
        start_date= (datetime.now() + relativedelta(day=1,hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S')
        end_date= (datetime.now()).strftime('%Y-%m-%d %H:%M:%S')
        return parser.parse(start_date) , parser.parse(end_date)

    # Last month
    elif userInput == 5:
        start_date= (datetime.now() + relativedelta(day=1, months=-1,hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S')
        end_date= (datetime.now() + relativedelta(day=1,days=-1)).strftime('%Y-%m-%d %H:%M:%S')
        return parser.parse(start_date) , parser.parse(end_date)

    # Go back X days from today
    elif userInput == 6:
        start_date= (datetime.now() + relativedelta(days=-int(input("enter how many days back: ")),hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S')
        end_date= (datetime.now()).strftime('%Y-%m-%d %H:%M:%S')
        return parser.parse(start_date) , parser.parse(end_date)

    # Go back X weeks from this week
    elif userInput == 7:
        start_date= (datetime.now() + relativedelta(weekday=MO(-1), weeks=-int(input("enter how many weeks back: ")),hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S')
        end_date= (datetime.now() + relativedelta(weekday=SU(-1))).strftime('%Y-%m-%d %H:%M:%S')
        return parser.parse(start_date) , parser.parse(end_date)

    # Choose a specific month(current year)
    elif userInput == 8:
        month=int(input("enter month number: "))
        start_date= (datetime.now() + relativedelta(day=1, month=month,hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S')
        end_date= (datetime.now() + relativedelta(month=month+1,day=1,days=-1, hour=23,minute=59,second=59)).strftime('%Y-%m-%d %H:%M:%S')
        return parser.parse(start_date) , parser.parse(end_date)
    
    elif userInput == 0:
        return 0,0



def download_excel_file(login,password,start_date, end_date,card=''):
    # Step 1: Log in to the website
    login_url = 'https://fleetcard.vivoenergy.com/WP/v1/login'
    payload = {
        'login': login,
        'password': password,
        }
    try:
        # code to attempt the login
        session = requests.Session()
        response = session.post(login_url, data=payload)
        if response.status_code != 200:
            raise Exception("Login failed, Problem with credentials")
    except requests.exceptions.ConnectionError:
        # code to handle the error when the connection is lost
        print("Error: Connection lost. Please check your internet connection and try again.")
        return
    except requests.exceptions.Timeout as e:
        print("The request timed out:", e)
        return
    except Exception as e:
        # handle the error if there is a problem with the credentials
        print("Error:", e)
        return

    # Step 2: Download the excel file
    file_url = 'https://fleetcard.vivoenergy.com/WP/v1/reports/card-purchase-list'
    query_params = {
        "refresh": "true",
        "itemsPerPage": 25,
        "currentPage": 1,
        "startDate": start_date,
        "endDate": end_date,
        "isProcessing": 1,
        "cardList": card,
        "service": "",
        "pos": "",
        "invoiceNumber": "",
        "format": "XLS",
        "additionalColumns": "consumption,holder,plate-nr"
    }
    print("Fetching requested data...")
    response=session.get(file_url, params=unquote(urlencode(query_params)))
    # with open('temp.xls', 'wb') as f:
    #     f.write(response.content)
    return response


#Main Code

accounts = pd.read_csv("Credentials.csv")
if os.path.exists('.\CardsNumbers'):
    for i in range(9):
        name = accounts.loc[i]['Account Name']
        if os.path.isfile(f'.\CardsNumbers\CardsInAccount {name}.csv'):
            pass
        else:
            print('Fetching Cards Data...')
            getCardsNumbers()
else:
    getCardsNumbers()


while True:

    print("1. Download All Purchases Data From All Accounts.\n2. Download STC Cards Purchases.\n3. Enter a Card number.\n4. Choose an Account.\n5. Enter a Card number to identify it's account. \n6. Update Cards Numbers. \n\n0. Exit the program")
    choice = int(input("Enter your choice [1-6]: "))

    if choice == 1:
        downloadAllData()

        
    elif choice == 2:
        downloadStcData()


    elif choice == 3:
        while True:
            card = int(input("Enter your card number(0 to exit): "))

            try:
                if card == 0:
                    break
                else:
                    credID= getCredentialsId(card)
                    print(credID)
                    if credID or credID == 0:
                        login = accounts.loc[credID]['username']
                        password = accounts.loc[credID]['password']
                        start_date, end_date= getUserDateChoice()
                        response = download_excel_file(login,password,start_date, end_date,card)
                        with open(f'Card {card} Statements Report {start_date.strftime("%b %d")} - {end_date.strftime("%b %d")}.xls', 'wb') as f:
                            f.write(response.content)
                        break

            except:
                print("Invalid card number. Try again.")    


    elif choice == 4:
        while True:
            print("Choose an Account: \n1. RANA MOTORS\n2. B.B.C INDUSTRIALS CO(GH) LTD\n3. LAJJIMARK CO. LTD.\n4. WEST AFRICA TIRE SERV. LTD\n5. HIGHLAND SPRINGS (GH) LTD\n6. KHOMARA PRINTING PRESS LTD\n7. ODAYMAT INVESTMENTS LTD\n8. RANA ATLAS\n9. ELDACO\n\n0. Exit")
            choice = int(input("Enter your choice [1-9]: "))

            try:

                accountsInfo= getUserNPass('WEST AFRICA TIRE SERV. LTD')
                login = accountsInfo[0]
                password = accountsInfo[1]
                accountName=accountsInfo[2]
                if login==0:
                    break

                else:

                    if login:
                        start_date, end_date= getUserDateChoice()
                        if start_date==0:
                            break
                        else:
                            response = download_excel_file(login,password,start_date, end_date)
                            with open(f'{accountName} Statements Report {start_date.strftime("%b %d")} - {end_date.strftime("%b %d")}.xls', 'wb') as f:
                                f.write(response.content)
                            break
            except:
                print("An Error Occurred, Try Again.")    



    elif choice == 5:
        while True:

            card = int(input("Enter your card number (0 to exit): "))

            try:

                if int(card) == 0:
                    break;
                    
                else:
                    credID= getCredentialsId(card)
                    if credID or credID == 0:
                        accountName = accounts.loc[credID]['Account Name']
                        print(f'This card number belongs to this account:\n {accountName}')
                        break
            except:
                print("Invalid card number. Try again.")


    elif choice == 6:
        getCardsNumbers()


    elif choice == 0:
        break    
    else:
        print("Invalid option. Try again.")