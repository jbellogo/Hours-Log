
# Easy to use Google Sheets API
import gspread
# Required credentials for acessing private data as is Google Sheets
from oauth2client.service_account import ServiceAccountCredentials

# Full, permissive scope to access all of a user's files
SCOPE = ['https://www.googleapis.com/auth/drive']
CREDENTIALS = ServiceAccountCredentials.from_json_keyfile_name('creds.json', SCOPE)
CLIENT = gspread.authorize(CREDENTIALS)
SHEET = CLIENT.open('SPREAD_SHEET_NAME').sheet1

# Dictionary to pandas data frame
import pandas as pd
df = pd.DataFrame(SHEET.get_all_records())


### Helper Methods ###

def front(self, n):
    '''
    like head() pandas method but for columns
    '''
    return self.iloc[:, :n]

def back(self, n):
    '''
    like tail() pandas method but for columns
    '''
    return self.iloc[:, -n:]

pd.DataFrame.front = front
pd.DataFrame.back = back


### WEEK STATS  ###
hours_per_day_data = df.loc[[13],:]


def print_week_info(hours_per_day_df):
    '''
    takes in a dataframe with the day totals, 2Row array
    Can't be generalized for any row yet
    '''
    message = "WEEK INFORMATION:\n\n"

    total = hours_per_day_df['TW'].values[0]
    message += f"Total hours this week: {total}\n"
    sorted_df = hours_per_day_df.loc[[13], "M":'Su'].sort_values(by=13, ascending=False, axis=1)

    max_frame = sorted_df.front(1)
    max_hours = max_frame.values[0][0]
    max_day = max_frame.columns[0]
    message += f"Week high of {max_hours}h on {max_day}\n"

    min_frame = sorted_df.back(1)
    min_hours = min_frame.values[0][0]
    min_day = min_frame.columns[0]
    message += f"Week low of {min_hours}h on {min_day}\n"

    average = hours_per_day_df.loc[[13], "M":'Su'].mean(axis=1)
    average = "{:.3f}".format(average.values[0])
    message += f"Average of hours working per day: {average}h\n"

    return message

### COURSE INFO ###
hours_per_course_data = df.loc[0:12,['Subject', 'TW']]


def print_course_info(hours_per_course_df):
    '''
    Calculates statistics based on hours per week spend on each course. Concatenates all into one single string.

    '''
    sorted_frame = hours_per_course_df.sort_values(by='TW', ascending=False)
    list_of_priorities = sorted_frame.head(3)['Subject']
    list_of_neglected = sorted_frame.tail(3)['Subject']

    ordered_times = sorted_frame['TW']
    sort_indices = list(sorted_frame.index) #list with the indices at the position of their sorted values

    message = "COURSE INFORMATION\n\nPrioritized:\n"
    count = 1
    for course in list_of_priorities:
        message += f"Your #{count} prioritized course this week was: "+ course + f" by: {ordered_times[sort_indices[count-1]]}h\n"
        count += 1
    count = 1
    message += "\nNeglected:\n"
    for course in list_of_neglected:
        message += f"Your #{count} neglected course this week was: "+ course + f" by: {ordered_times[sort_indices[-count]]}h\n"
        count+= 1

    school_course_average = hours_per_course_df.loc[6:12 ,['TW']].mean()
    school_course_average = "{:.3f}".format(school_course_average[0])

    message += f"\nMean:\nThe average time spent per school course this week was: {school_course_average}h"
    return message


def main():
    '''
    Creates the message to be sent as a single string
    '''
    week = print_week_info(hours_per_day_data)
    course = print_course_info(hours_per_course_data)
    return ('\n\n' + week + '\n\n' + course)


# Sending Email:
import smtplib, ssl, getpass #for invisible input

port = 465  # For SSL
smtp_server = "smtp.gmail.com"
sender_email = "youremail@gmail.com"  # Enter your address. need to enable "Less Secure Apps" on gmail
receiver_email = "reveiver@gmail.com"  # Enter receiver address
password = getpass.getpass("Type your password and press enter: ")
message = 'Week Summary: \n' + main()
print(message)


context = ssl.create_default_context()
with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
    try:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, message)
        print('message sent successfully')
    except:
        print('An error occured')

