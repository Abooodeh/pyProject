import requests
from datetime import datetime
from dateutil import parser
from dateutil.relativedelta import relativedelta
from dateutil.rrule import MO, TU, WE, TH, FR, SA, SU

#functions

def checkUserInput(userInput):
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



def download_excel_file(start_date, end_date,card=''):
    # Step 1: Log in to the website
    login_url = 'https://fleetcard.vivoenergy.com/WP/v1/login'
    payload = {
        'login': '',
        'password': ''
        }
    session = requests.Session()
    session.post(login_url, data=payload)

    # Step 2: Download the excel file
    file_url = f'https://fleetcard.vivoenergy.com/WP/v1/reports/card-purchase-list?refresh=true&itemsPerPage=25&currentPage=1&startDate={start_date}&endDate={end_date}&isProcessing=1&cardList={card}&service=&pos=&invoiceNumber=&format=XLS&additionalColumns=consumption,holder,plate-nr'
    response=session.get(file_url)
    with open('CardPurchases.xls', 'wb') as f:
        f.write(response.content)


#main code

print("Choose an option:\n1. Custom date and time\n2. This week\n3. Last week\n4. This month\n5. Last month\n6. Go back X days from today\n7. Go back X weeks from this week\n8. Choose a specific month(current year)")


# getting user choice
userInput=input()
#generating the dates based on user choice
start_date , end_date = checkUserInput(int(userInput))
download_excel_file(start_date, end_date)

print("Done")
