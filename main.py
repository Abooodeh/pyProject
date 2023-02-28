import os
import json
import requests
import pandas as pd
import concurrent.futures
from   rich                     import print
from   rich.table               import Table
from   rich.console             import Console
from   tabulate                 import tabulate
from   datetime                 import datetime
from   dateutil                 import parser
from   urllib.parse             import urlencode,unquote
from   dateutil.relativedelta   import relativedelta
from   dateutil.rrule           import MO,SU
import traceback
from rich.traceback import install
install() # this prints the errors in the terminal in a readable way

# Return lists of accounts informations
def constructLogins():
    usernamesL=[]
    passwordsL=[]
    namesL=[]
    aNamesL=[]
    for credential in loginInfo["Credentials"]:
        username=credential["username"]
        usernamesL.append(username)
        password= credential['password']
        passwordsL.append(password)
        accountName = credential['Account Name']
        namesL.append(accountName)
        accountNameA = credential['abbrevietedNames']
        aNamesL.append(accountNameA)
    return usernamesL,passwordsL,namesL,aNamesL

# Return a loggedIn session
def logIn(login,password):
        # Log in to the website
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
    return session

# Return a DataFrame with search results and account information
def getSearchData(login,password,accountName,accountNameA,query):
    session = logIn(login,password)
    # Getting data
    file_url = loginInfo['search']['fileURL']
    params = loginInfo['search']['params']
    params.update({'query':query})
    data=session.get(file_url,params=params).json()
    df= pd.DataFrame([[data['cardNumber'], data['holder']['carPlateNumber'], data['holder']['name'],data['cardStatus']['desc'], data['purseList'][0]['limitFormatted'],'\n'.join([x['cardNumber'] for x in data['mifareList']]) if data['mifareList'] else 'None',accountNameA, login, password, accountName]for data in data['cardList']])
    return df

# Will return a DataFrame that have SearchData for all accounts
def search(query):
    query = [query] * 9
    with concurrent.futures.ThreadPoolExecutor() as executor:
        results = executor.map(getSearchData, usernamesL, passwordsL, namesL, aNamesL,query)
    df = pd.concat(list(results))
    df.columns = ['Card number', 'Vehicle Number' , 'Card Holder' , 'Status' ,'Limit','Sticker Number', 'Account','Username', 'Password' , 'Account Name']
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    df.reset_index(drop=True, inplace=True)
    return df

# Returns card and account information based on user choice
def getUserCardChoice(df):   
    if df.shape[0] == 1:                                        
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
            print(tabulate(page[['Card number', 'Vehicle Number' , 'Card Holder', 'Status' ,'Account','Limit']], headers = 'keys', tablefmt = 'rounded_grid'))
            try:
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
                    rowIndex=int(rowIndex)
                    if rowIndex >= 0 and rowIndex < df.shape[0]:
                        return df.loc[rowIndex]['Username'],df.loc[rowIndex]['Password'],df.loc[rowIndex]['Card number'],df.loc[rowIndex]['Vehicle Number']
            except ValueError:
                input("Invalid input. Please enter 'n', 'p', 'q', or a valid index number.\nPress Enter to continue...")
    # if there is less than 15 results, display them and user will be prompted to choose
    else:
        while True:
                os.system('cls')
                print('Multiple Matches Founds:')
                try:
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

# Display the search results
def displaySearchInfo(df):   
    if df.shape[0] > 10:
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
            print(tabulate(page[['Card number', 'Vehicle Number' , 'Status' ,'Account','Limit','Sticker Number']], headers = 'keys', tablefmt = 'rounded_grid',showindex='never'))
            rowIndex = input("Enter 'n' for next page, 'p' for previous page, 'q' to quit: ")  
            if rowIndex !='n' and rowIndex != 'p' and rowIndex != 'q':
                input("Invalid input. Please enter 'n', 'p' or 'q'.\nPress Enter to continue...")
            if rowIndex == 'n':
                    currentPage += 1
                    if currentPage > totalPages:
                            currentPage = 1
            elif rowIndex == 'p':
                    currentPage -= 1
                    if currentPage < 1:
                            currentPage = totalPages
            elif rowIndex == 'q':
                    return
    # if there is less than 15 results, display them and user will be prompted to choose
    else:
        os.system('cls')
        print('Multiple Matches Founds:')
        print(tabulate(df[['Card number', 'Vehicle Number' , 'Status' ,'Account','Limit','Sticker Number']], headers = 'keys', tablefmt = 'rounded_grid',showindex='never'))
        input('Press Enter to continue...')

# Return account usage percentage and account name
def getAccountUsage(login,password, name):
    session = logIn(login,password)
    file_url = loginInfo['dashboard']['fileURL']
    #Getting data
    data=session.get(file_url).json()
    #Extract the relevant data and calculate the usage percentage
    creditLimit = abs(data['accountCreditLimit'])
    creditAvailable = abs(data['closingBalanceAvailableForCardRecharges'])
    percentage = round(((creditLimit - creditAvailable) / creditLimit * 100), 2)
    # Highlight the percentage in red if it's greater than 80%
    if percentage >= 80:
        percentage = f'[red bold]{percentage}%[/]'
        name=f'[red bold]{name}[/]'
    else:
        percentage = f'[cyan]{percentage}%[/]'
    
    return percentage, name

# Get usage percentage for all account and print them in a table
def dashboard():
    print('Refreshing Data...')
    #use concurrent to multithread the function
    with concurrent.futures.ThreadPoolExecutor() as executor:
        results = executor.map(getAccountUsage, usernamesL,passwordsL, aNamesL)
    percentagesL = []
    names = []
    #extracting data from results
    for data in results:
        percentagesL.append(data[0])
        names.append(data[1])
    table = Table(show_header=False)
    for i in range(9):
        table.add_column("",justify="center")
    table.add_row(*names)
    table.add_row(*percentagesL)
    return table
# Download a file from the server

def downloadFile(login,password,start_date, end_date,card=''):
    # Log in to the website
    session = logIn(login,password)
    file_url = loginInfo['file']['fileURL']
    params = loginInfo['file']['params']
    params.update({'startDate': start_date, 'endDate': end_date, 'cardList': card})
    print("Fetching requested data...")
    # Download the excel file
    response=session.get(file_url, params=unquote(urlencode(params)))
    return response

# Download all accounts statements to 1 file
def downloadAllData():
    os.system('cls')
    start_date , end_date = getUserDateChoice("LastMonth")
    file_url = loginInfo['file']['fileURL']
    params = loginInfo['file']['params']
    params.update({'startDate': start_date, 'endDate': end_date,"format": "CSV"})
    downloadPath=loginInfo["downloadsPath"]["allAccounts"]
    os.makedirs(downloadPath) if not os.path.exists(downloadPath) else None
    fileName=f'{downloadPath}\All Card Purchases Report {start_date.strftime("%b %d")} - {end_date.strftime("%b %d")}.xls' 
    fileName=f'{downloadPath}\All Card Purchases Report {start_date.strftime("%b %d")} - {end_date.strftime("%b %d)")} {datetime.now().strftime("(%I.%M.%S.%p)")}.xls' if os.path.exists(fileName) else fileName
    print('Downloading All Accounts Data:\n') 
    for i in range(9):
        login=usernamesL[i]
        password= passwordsL[i]
        accountName=namesL[i]
        # Log in to the website
        session = logIn(login,password)
        print(f'Account {accountName} Started.') 
        # Download the excel file
        response=session.get(file_url, params=unquote(urlencode(params)))
        with open(f'{downloadPath}\downloadingData.csv', 'a') as f:
            f.write(response.text)
    df = pd.read_csv(f'{downloadPath}\downloadingData.csv')
    value_index = df[df['Transaction date'] == 'Transaction date'].index
    df = df.drop(value_index)
    df.to_excel(fileName,engine='openpyxl', index=False)
    os.remove(f'{downloadPath}\downloadingData.csv')
    input('Downloading Data is Done.\nPress Enter to continue...')

# Download STC Data
def downloadStcData():
    start_date , end_date = getUserDateChoice("LastMonth")
    file_url = loginInfo['STC']['fileURL']
    params = loginInfo['STC']['params']
    params.update({'startDate': start_date, 'endDate': end_date})
    downloadPath=loginInfo["downloadsPath"]["STC"]
    os.makedirs(downloadPath) if not os.path.exists(downloadPath) else None
    fileName=f'{downloadPath}\STC Report {start_date.strftime("%b %d")} - {end_date.strftime("%b %d")}.xls' 
    fileName=f'{downloadPath}\STC Report {start_date.strftime("%b %d")} - {end_date.strftime("%b %d)")} {datetime.now().strftime("(%I.%M.%S.%p)")}.xls' if os.path.exists(fileName) else fileName
    accountsInfo= getUserNPass('WEST AFRICA TIRE SERV. LTD')
    login = accountsInfo[0]
    password = accountsInfo[1]
    # Log in to the website
    session = logIn(login,password)
    # Download the excel file
    print('STC Cards Purchases Data Download Started.')
    response=session.get(file_url, params=unquote(urlencode(params)))
    with open(fileName, 'wb') as f:
        f.write(response.content)   
    print('STC Cards Purchases Data Download is Done.')
    input("Press Enter to continue...")

# Return Username and Password
def getUserNPass(choice):

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
    index = namesL.index(account)
    return usernamesL[index],passwordsL[index],aNamesL[index]

# Return the user date choice after formatting it
def getUserDateChoice(userInput=''):
    os.system('cls')
    if userInput =='' :
        print("Choose an option:\n1. Custom date and time\n2. This week\n3. Last week\n4. Date of a week\n5. This month\n6. Last month\n7. Go back X days from today\n8. Go back X weeks from this week\n9. Choose a specific month (current year)\n\n0. Exit\n")
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
        end_date= (datetime.now() + relativedelta(weekday=SU(-1),hour=23,minute=59,second=59)).strftime('%Y-%m-%d %H:%M:%S')
        return parser.parse(start_date) , parser.parse(end_date)
    
    # Choose a date to return the date of the week 
    elif userInput== 4:
            os.system('cls')
            date=input('enter a date (0 to exit): ')
            if date!='0':
                start_date= (parser.parse(date) + relativedelta(weekday=MO(-1),hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S')
                end_date= (parser.parse(date) + relativedelta(weekday=SU,hour=23,minute=59,second=59)).strftime('%Y-%m-%d %H:%M:%S')
                return parser.parse(start_date) , parser.parse(end_date)
            else:
                return 0,0
            
    # This month
    elif userInput == 5:
        os.system('cls')
        start_date= (datetime.now() + relativedelta(day=1,hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S')
        end_date= (datetime.now()).strftime('%Y-%m-%d %H:%M:%S')
        return parser.parse(start_date) , parser.parse(end_date)

    # Last month
    elif userInput == 6 or userInput == "LastMonth" :
        os.system('cls')
        start_date= (datetime.now() + relativedelta(day=1, months=-1,hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S')
        end_date= (datetime.now() + relativedelta(day=1,days=-1,hour=23,minute=59,second=59)).strftime('%Y-%m-%d %H:%M:%S')
        return parser.parse(start_date) , parser.parse(end_date)

    # Go back X days from today
    elif userInput == 7:
        os.system('cls')
        days=int(input("enter how many days back (0 to exit): "))
        if days!=0:
            start_date= (datetime.now() + relativedelta(days=-days,hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S')
            end_date= (datetime.now()).strftime('%Y-%m-%d %H:%M:%S')
            return parser.parse(start_date) , parser.parse(end_date)
        else:
            return 0,0
        
    # Go back X weeks from this week
    elif userInput == 8:
        os.system('cls')
        weeks=int(input("enter how many weeks back (0 to exit): "))
        if weeks!=0:
            start_date= (datetime.now() + relativedelta(weekday=MO(-1), weeks=-weeks,hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S')
            end_date= (datetime.now() + relativedelta(weekday=SU(-1))).strftime('%Y-%m-%d %H:%M:%S')
            return parser.parse(start_date) , parser.parse(end_date)
        else: 
            return 0,0
        
    # Choose a specific month(current year)
    elif userInput == 9:
        os.system('cls')
        month=int(input("enter month number (0 to exit): "))
        if month!=0:
            start_date= (datetime.now() + relativedelta(day=1, month=month,hour=00,minute=00,second=00)).strftime('%Y-%m-%d %H:%M:%S')
            end_date= (datetime.now() + relativedelta(month=month+1,day=1,days=-1, hour=23,minute=59,second=59)).strftime('%Y-%m-%d %H:%M:%S')
            return parser.parse(start_date) , parser.parse(end_date)
        else:
            return 0,0
        
    elif userInput == 0:
        return 0,0

def main():
    table=dashboard()
    while True:
        os.system('cls')
        console = Console()
        console.print("[bold]Account Usage Percentage[/]", justify="center")
        console.print(table, justify="center")
        print("[bold]Main Menu[/]: \n")
        print("1. Card Statments\n2. Card Info\n3. Account Statments\n4. STC Statments\n\n0. Exit the program\nr. Refresh usage percentage ")
        choice = input("Enter your choice [1-4]: ")

        if choice == '1':
            os.system('cls')
            while True:
                os.system('cls')
                inputCard = input("Enter Card Number or Vehicle Number(q to quit): ")
                if inputCard == 'q':
                    break
                try:
                    print('Fetching Data...')
                    accountsInfo=getUserCardChoice(search(inputCard))
                    if accountsInfo!='q':
                        login=accountsInfo[0]
                        password=accountsInfo[1]
                        card=int(accountsInfo[2])
                        vehicleNumber=accountsInfo[3]
                        start_date, end_date= getUserDateChoice()
                        downloadPath=loginInfo["downloadsPath"]["card"]
                        if start_date!=0:
                            os.makedirs(downloadPath) if not os.path.exists(downloadPath) else None
                            fileName=f'{downloadPath}\{vehicleNumber} ({start_date.strftime("%b %d")} - {end_date.strftime("%b %d")}).xls'
                            fileName=f'{downloadPath}\{vehicleNumber} ({start_date.strftime("%b %d")} - {end_date.strftime("%b %d) (%I.%M.%S.%p)")}.xls' if os.path.exists(fileName) else fileName
                            response = downloadFile(login,password,start_date, end_date,card)
                            with open(fileName, 'wb') as f:
                                f.write(response.content)
                            input("Data download is done\nPress Enter to continue...")   
                            break
                except:
                    input('Invalid card or vehicle number. Try again.\nPress Enter to continue...')



        elif choice == '2':
            while True:
                os.system('cls')
                card =input("Enter Card Number or Vehicle Number(q to quit): ")
                if card == 'q':
                    break;
                try:
                    print('Fetching Data...')
                    displaySearchInfo(search(card))
                except:
                    input('Invalid card or vehicle number. Try again.\nPress Enter to continue...')

        elif choice == '3':
            while True:
                os.system('cls')
                print("Choose an Account: \n1. RANA MOTORS\n2. B.B.C INDUSTRIALS CO (GH) LTD\n3. LAJJIMARK CO. LTD.\n4. WEST AFRICA TIRE SERV. LTD\n5. HIGHLAND SPRINGS (GH) LTD\n6. KHOMARA PRINTING PRESS LTD\n7. ODAYMAT INVESTMENTS LTD\n8. RANA ATLAS\n9. ELDACO\n\n10. All Accounts Data for last month\n\n0. Exit")
                choice = int(input("Enter your choice [1-10]: "))
                if choice==0:
                    break
                if choice==10:
                    downloadAllData()
                    break
                try:
                    accountsInfo= getUserNPass(choice)
                    login = accountsInfo[0]
                    password = accountsInfo[1]
                    accountName=accountsInfo[2]
                    start_date, end_date= getUserDateChoice()
                    if start_date!=0:
                        downloadPath=loginInfo["downloadsPath"]["account"]
                        os.makedirs(downloadPath) if not os.path.exists(downloadPath) else None
                        fileName=f'{downloadPath}\{accountName} Statements Report {start_date.strftime("%b %d")} - {end_date.strftime("%b %d")}.xls'
                        fileName=f'{downloadPath}\{accountName} Statements Report {start_date.strftime("%b %d")} - {end_date.strftime("%b %d (%I.%M.%S.%p)")}.xls' if os.path.exists(fileName) else fileName
                        response = downloadFile(login,password,start_date, end_date)
                        with open(fileName, 'wb') as f:
                            f.write(response.content)
                        input("Data download is done\nPress Enter to continue...")
                        break
                except Exception as e:
                    input(f"{e}\nAn Error Occurred, Try Again.\nPress Enter to continue...")    

        elif choice == '4':
            os.system('cls')
            downloadStcData()
        elif choice == 'r':
            os.system('cls')
            table=dashboard()
        elif choice == '0':
            break    
        else:
            input("Invalid option. Try again.\nPress Enter to continue...")
    os.system('cls')

realPath=os.path.realpath(os.path.dirname(__file__))
with open(f'{realPath}\Login and Download Info.json', 'r') as f:
    loginInfo = json.load(f)
downloadPath=loginInfo["downloadsPath"]["defaultDownloadPath"]
usernamesL,passwordsL,namesL,aNamesL=constructLogins()
os.system('cls')

main()