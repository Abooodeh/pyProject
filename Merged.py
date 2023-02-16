import os
import csv
import requests
import pandas as pd
from datetime import datetime
from dateutil import parser
from dateutil.relativedelta import relativedelta
from dateutil.rrule import MO, TU, WE, TH, FR, SA, SU
from urllib.parse import urlencode,unquote
from tabulate import tabulate


def search(query,search='',login= '',password= '',df = pd.DataFrame()):
    # the search parameter purpose is to determine the way the function is going to be used when called.
    # search='' will return values.
    # search='displayInfo' will display search results only.
    for index, row in accounts.iterrows():
            login=row['username']
            password= row['password']
            accountName = row['Account Name']
            accountNameA = row['abbrevietedNames']
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
                    input("Error: Connection lost. Please check your internet connection and try again.")
                    return
            except requests.exceptions.Timeout as e:
                    input("The request timed out:", e)
                    return
            except Exception as e:
                    # handle the error if there is a problem with the credentials
                    input("Error:", e)
                    return
            # Step 2: Getting data
            file_url = 'https://fleetcard.vivoenergy.com/WP/v1/card-list'
            query_params = {
                    "refresh": "true",
                    "itemsPerPage": 0,
                    "currentPage": 0,
                    "idCardStatus": "",
                    "purseList": "",
                    "query": query,
                    "depAndCC": ""
                    }
            data=session.get(file_url,params=query_params).json()
            # Extract the relevant data from the dictionary object
            df1= pd.DataFrame([[data['cardNumber'], data['holder']['carPlateNumber'], data['holder']['name'],data['cardStatus']['desc'], data['purseList'][0]['limitFormatted'],'\n'.join([x['cardNumber'] for x in data['mifareList']]) if data['mifareList'] else 'None',accountNameA, login, password, accountName]for data in data['cardList']])
            df=pd.concat([df,df1])
    df.columns = ['Card number', 'Vehicle Number' , 'Card Holder' , 'Status' ,'Limit','Sticker Number', 'Account','Username', 'Password' , 'Account Name']
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    df.reset_index(drop=True, inplace=True)
    # if there is only one row, return the result 
    if df.shape[0] == 1:                                        
        if search == 'displayInfo' :
            os.system('cls')
            print(tabulate(df[['Card number', 'Vehicle Number' , 'Status' ,'Account','Limit','Sticker Number']], headers = 'keys', tablefmt = 'rounded_grid',showindex='never'))
            input('Press enter to continue...')
            return
        else:    
            return df.loc[0]['Username'],df.loc[0]['Password'],df.loc[0]['Card number'],df.loc[0]['Vehicle Number'],df.loc[0]['Account Name']
    # if there is more than 15 results, display them and user will be prompted to choose, with pages navigation
    elif df.shape[0] > 10:
        itemsPerPage=10
        totalPages = (df.shape[0] // itemsPerPage) + 1
        currentPage = 1
        while True:
            startIndex = (currentPage - 1) * itemsPerPage
            endIndex = startIndex + itemsPerPage
            page = df[startIndex:endIndex]
            os.system('cls')
            print('Multiple Matches Founds')
            print(f"Page {currentPage} of {totalPages}")
            if search == 'displayInfo':
                print(tabulate(page[['Card number', 'Vehicle Number' , 'Status' ,'Account','Limit','Sticker Number']], headers = 'keys', tablefmt = 'rounded_grid',showindex='never'))
            else:
                print(tabulate(page[['Card number', 'Vehicle Number' , 'Card Holder', 'Status' ,'Account','Limit']], headers = 'keys', tablefmt = 'rounded_grid'))
            try:
                    if search == 'displayInfo':
                        rowIndex = input("Enter 'n' for next page, 'p' for previous page, 'q' to quit: ")  
                        if rowIndex !='n' and rowIndex != 'p' and rowIndex != 'q':
                            input("Invalid input. Please enter 'n', 'p' or 'q'.\nPress Enter to continue...")
                    else:
                        rowIndex = input("Enter 'n' for next page, 'p' for previous page, 'q' to quit or Enter the index of a row to choose the card number: ")
                    if rowIndex == 'n':
                            currentPage += 1
                            if currentPage > totalPages:
                                    currentPage = 1
                    elif rowIndex == 'p':
                            currentPage -= 1
                            if currentPage < 1:
                                    currentPage = totalPages
                    elif rowIndex == 'q':
                            return 'q'
                    else:
                        if search != 'displayInfo':
                            rowIndex = int(rowIndex)
                            if rowIndex >= 0 and rowIndex < df.shape[0]:
                                    return df.loc[rowIndex]['Username'],df.loc[rowIndex]['Password'],df.loc[rowIndex]['Card number'],df.loc[rowIndex]['Vehicle Number']
            except ValueError:
                input("Invalid input. Please enter 'n', 'p', 'q', or a valid index number.\nPress Enter to continue...")
    # if there is less than 15 results, display them and user will be prompted to choose
    else:
        df.index = df.index + 1
        while True:
                os.system('cls')
                print('Multiple Matches Founds:')
                try:
                    if search == 'displayInfo':
                        print(tabulate(df[['Card number', 'Vehicle Number' , 'Status' ,'Account','Limit','Sticker Number']], headers = 'keys', tablefmt = 'rounded_grid',showindex='never'))
                        input('Press Enter to continue...')
                        break
                    else:  
                        print(tabulate(df[['Card number', 'Vehicle Number' , 'Card Holder', 'Status' ,'Account','Limit']], headers = 'keys', tablefmt = 'rounded_grid'))
                        rowIndex=input('Enter the index of a row to choose the card number(q to quit): ')
                        if rowIndex == 'q':
                                return 'q'
                        else:
                                rowIndex = int(rowIndex)
                                if rowIndex >= 0 and rowIndex < df.shape[0]:
                                        return df.loc[rowIndex]['Username'],df.loc[rowIndex]['Password'],df.loc[rowIndex]['Card number'],df.loc[rowIndex]['Vehicle Number']
                                else:
                                        input("Invalid row index. Please enter a valid number.\nPress Enter to continue...")
                except ValueError:
                        input("Invalid input. Please enter a valid index number.\nPress Enter to continue...")

def downloadAllData():
    os.system('cls')
    print('Downloading All Accounts Data:\n') 
    for index, row in accounts.iterrows():
        login=row['username']
        password= row['password']
        accountName=row['Account Name']
        # Step 1: Log in to the website
        start_date= parser.parse((datetime.now() + relativedelta(day=1,months=-1,hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S'))
        end_date= parser.parse((datetime.now() + relativedelta(day=1,days=-1,hour=23,minute=59,second=59)).strftime('%Y-%m-%d %H:%M:%S'))
        login_url = 'https://fleetcard.vivoenergy.com/WP/v1/login'
        payload = {
            'login': login,
            'password': password,
            }
        try:
            # Login attempt
            session = requests.Session()
            response = session.post(login_url, data=payload)
            if response.status_code != 200:
                raise Exception("Login failed, Problem with credentials")
        except requests.exceptions.ConnectionError:
            # handle the error when the connection is lost
            input("Error: Connection lost. Please check your internet connection and try again.")
            return
        except requests.exceptions.Timeout as e:
            input("The request timed out:", e)
            return
        except Exception as e:
            # handle the error if there is a problem with the credentials
            input("Error:", e)
            return

        # Step 2: Download the excel file
        print(f'Account {accountName} Started.') 
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
        with open(f'{downloadPath}\downloadingData.csv', 'a') as f:
            f.write(response.text)
    df = pd.read_csv(f'{downloadPath}\downloadingData.csv')
    value_index = df[df['Transaction date'] == 'Transaction date'].index
    df = df.drop(value_index)
    df.to_excel(f'{downloadPath}\All Card Purchases Report {start_date.strftime("%b %d")} - {end_date.strftime("%b %d")}.xls',engine='openpyxl', index=False)
    os.remove(f'{downloadPath}\downloadingData.csv')
    input('Downloading Data is Done.\nPress Enter to continue...')
    


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
        input("Error: Connection lost. Please check your internet connection and try again.")
        return
    except requests.exceptions.Timeout as e:
        input("The request timed out:", e)
        return 
    except Exception as e:
        # handle the error if there is a problem with the credentials
        input("Error:", e)
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
        "cardList": "1401053310,1401053311,1401053312,1401053313,1401053314,1401053315,1401053316,1401053317,1401053318,1401053319",
        "service": "",
        "pos": "",
        "invoiceNumber": "",
        "format": "XLS",
        "additionalColumns": "consumption,holder,plate-nr"
    }
    print('STC Cards Purchases Data Download Started.')
    response=session.get(file_url, params=unquote(urlencode(query_params)))
    with open(f'{downloadPath}\STC Report {start_date.strftime("%b %d")} - {end_date.strftime("%b %d")}.xls', 'wb') as f:
        f.write(response.content)   
    print('STC Cards Purchases Data Download is Done.')
    input("Press Enter to continue...")

def getUserNPass(choice,account=0):

    if choice == 1 or choice == 'RANA MOTORS' :
        account= 'RANA MOTORS'

    elif choice == 2 or choice == 'B.B.C INDUSTRIALS CO(GH) LTD' :
        account= 'B.B.C INDUSTRIALS CO(GH) LTD'

    elif choice == 3 or choice == 'LAJJIMARK CO. LTD.' :
        account= 'LAJJIMARK CO. LTD.'

    elif choice == 4 or choice == 'WEST AFRICA TIRE SERV. LTD':
        account= 'WEST AFRICA TIRE SERV. LTD'

    elif choice == 5 or choice == 'HIGHLAND SPRINGS (GH) LTD' :
        account= 'HIGHLAND SPRINGS (GH) LTD'

    elif choice == 6 or choice == 'KHOMARA PRINTING PRESS LTD' :
        account= 'KHOMARA PRINTING PRESS LTD'

    elif choice == 7 or choice == 'ODAYMAT INVESTMENTS LTD' :
        account= 'ODAYMAT INVESTMENTS LTD'

    elif choice == 8 or choice == 'RANA ATLAS' :
        account= 'RANA ATLAS'

    elif choice == 9 or choice == 'ELDACO' :
        account= 'ELDACO'

    elif choice == 10 :
        downloadAllData()

    elif choice == 0:
        return 0,0,0;

    else:
        return None
    if account==0:
        return 0,0,0;

    if account in accounts.values:
        index = accounts[accounts['Account Name'] == account].index[0]
        return accounts.loc[index]['username'],accounts.loc[index]['password'],accounts.loc[index]['Account Name']
    


def getUserDateChoice():
    os.system('cls')
    print("Choose an option:\n1. Custom date and time\n2. This week\n3. Last week\n4. This month\n5. Last month\n6. Go back X days from today\n7. Go back X weeks from this week\n8. Choose a specific month(current year)\n\n0. Exit\n")
    userInput=int(input("Enter your choice: "))
    # Custom date and time
    if userInput == 1:       
        os.system('cls')    

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
                        input("Invalid input, enter a valid date and time.\nPress enter to continue...")
            while True:
                try:
                    endDateInput=input("Enter End Date and Time(0 to exit): ")
                    if endDateInput == '0':
                        return 0,0
                    else:
                        end_date= parser.parse(endDateInput).strftime('%Y-%m-%d %H:%M:%S')
                        break
                except ValueError:
                    input("Invalid input, please enter a valid date and time.\nPress enter to continue...")
            if(start_date < end_date):
                return parser.parse(start_date) , parser.parse(end_date)
            input("Invalid input, end date must be after start date, enter a valid date and time.\nPress enter to continue...")

    # This week
    elif userInput == 2:
        os.system('cls')
        start_date= (datetime.now() + relativedelta(weekday=MO(-1),hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S')
        end_date= (datetime.now()).strftime('%Y-%m-%d %H:%M:%S')
        return parser.parse(start_date) , parser.parse(end_date)

    # Last week
    elif userInput == 3:
        os.system('cls')
        start_date= (datetime.now() + relativedelta(weekday=MO(-2),hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S')
        end_date= (datetime.now() + relativedelta(weekday=SU(-1))).strftime('%Y-%m-%d %H:%M:%S')
        return parser.parse(start_date) , parser.parse(end_date)

    # This month
    elif userInput == 4:
        os.system('cls')
        start_date= (datetime.now() + relativedelta(day=1,hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S')
        end_date= (datetime.now()).strftime('%Y-%m-%d %H:%M:%S')
        return parser.parse(start_date) , parser.parse(end_date)

    # Last month
    elif userInput == 5:
        os.system('cls')
        start_date= (datetime.now() + relativedelta(day=1, months=-1,hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S')
        end_date= (datetime.now() + relativedelta(day=1,days=-1)).strftime('%Y-%m-%d %H:%M:%S')
        return parser.parse(start_date) , parser.parse(end_date)

    # Go back X days from today
    elif userInput == 6:
        os.system('cls')
        start_date= (datetime.now() + relativedelta(days=-int(input("enter how many days back: ")),hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S')
        end_date= (datetime.now()).strftime('%Y-%m-%d %H:%M:%S')
        return parser.parse(start_date) , parser.parse(end_date)

    # Go back X weeks from this week
    elif userInput == 7:
        os.system('cls')
        start_date= (datetime.now() + relativedelta(weekday=MO(-1), weeks=-int(input("enter how many weeks back: ")),hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S')
        end_date= (datetime.now() + relativedelta(weekday=SU(-1))).strftime('%Y-%m-%d %H:%M:%S')
        return parser.parse(start_date) , parser.parse(end_date)

    # Choose a specific month(current year)
    elif userInput == 8:
        os.system('cls')
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
        input("Error: Connection lost. Please check your internet connection and try again.")
        return
    except requests.exceptions.Timeout as e:
        input("The request timed out:", e)
        return
    except Exception as e:
        # handle the error if there is a problem with the credentials
        input("Error:", e)
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
    return response

#Main Code 

downloadPath="D:\python"
realPath=os.path.realpath(os.path.dirname(__file__))
accounts = pd.read_csv(f'{realPath}\Credentials.csv')
os.system('cls')
while True:
    os.system('cls')
    print("1. Card Statments\n2. Card Info\n3. Account Statments\n4. STC Statments\n\n0. Exit the program\n")
    choice = input("Enter your choice [1-4]: ")

    if choice == '1':
        os.system('cls')
        while True:
            os.system('cls')
            inputCard = input("Enter Card Number or Vehicle Number(q to quit): ")
            try:
                if inputCard == 'q':
                    break
                else:
                    print('Fetching Data...')
                    accountsInfo=search(inputCard)
                    if accountsInfo=='q':
                        break
                    login=accountsInfo[0]
                    password=accountsInfo[1]
                    card=int(accountsInfo[2])
                    vehicleNumber=accountsInfo[3]
                    if login:
                        start_date, end_date= getUserDateChoice()
                        if start_date==0:
                            break
                        else:
                            response = download_excel_file(login,password,start_date, end_date,card)
                            with open(f'{downloadPath}\{vehicleNumber} {start_date.strftime("%b %d")} - {end_date.strftime("%b %d")}.xls', 'wb') as f:
                                f.write(response.content)
                            input("Data download is done\nPress Enter to continue...")   
                            break
            except:
                input('Invalid card or vehicle number. Try again.\nPress Enter to continue...')



    elif choice == '2':
        while True:
            os.system('cls')
            card =input("Enter Card Number or Vehicle Number(q to quit): ")
            try:
                if card == 'q':
                    break;
                else:
                    print('Fetching Data...')
                    search(card,'displayInfo')
            except:
                input('Invalid card or vehicle number. Try again.\nPress Enter to continue...')

    elif choice == '3':
        while True:
            os.system('cls')
            print("Choose an Account: \n1. RANA MOTORS\n2. B.B.C INDUSTRIALS CO(GH) LTD\n3. LAJJIMARK CO. LTD.\n4. WEST AFRICA TIRE SERV. LTD\n5. HIGHLAND SPRINGS (GH) LTD\n6. KHOMARA PRINTING PRESS LTD\n7. ODAYMAT INVESTMENTS LTD\n8. RANA ATLAS\n9. ELDACO\n\n10. All Accounts Data for last month\n\n0. Exit")
            choice = int(input("Enter your choice [1-10]: "))

            try:

                accountsInfo= getUserNPass(choice)
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
                            with open(f'{downloadPath}\{accountName} Statements Report {start_date.strftime("%b %d")} - {end_date.strftime("%b %d")}.xls', 'wb') as f:
                                f.write(response.content)
                            input("Data download is done\nPress Enter to continue...")
                            break
            except:
                input("An Error Occurred, Try Again.\nPress Enter to continue...")    

    elif choice == '4':
        os.system('cls')
        downloadStcData()

    elif choice == '0':
        break    
    else:
        input("Invalid option. Try again.\nPress Enter to continue...")