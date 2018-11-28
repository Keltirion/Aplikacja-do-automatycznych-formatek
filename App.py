# Libraries to import
import datetime, xlrd, sys, traceback, time
from Emailer_class import *

data_list = []
# External files. They must have the same name as given in the code.
file = 'test2.xlsx'
file_email = 'email_body.txt'

# File information. Extract from the file first row and the number of rows.
xl_workbook = xlrd.open_workbook(file)
xl_sheet = xl_workbook.sheet_by_name('Sheet1')
xl_rows = xl_sheet.nrows

# Email data. Preparing the email body from .txt file
topic = "Questions regarding the delivery of your VELFAC ORDER ({})"
body = open(file_email, 'r', encoding='UTF-8')
contents = body.read()
# Dates needed for the email content.
today = datetime.date.today()
next_monday = today + datetime.timedelta(days=-today.weekday(), weeks=2)
prior_friday = next_monday - datetime.timedelta(days=3)

# Quick instruction printed for the user before running the script.
print('### Before starting please close all outlook windows, except main one ###\n')
# Asking the user to for starting the creation.
start = input('Do you want to start creating emails? There will be ' + str(xl_rows) + ' emails (yes/no)\n')

# Function for clearing the data from unnecessary informations.
def dataprep(max_range, storage):
    for i in range(1, max_range):
        xl_row = xl_sheet.row_values(i)
        del xl_row[1]
        del xl_row[6:15]
        storage.append(xl_row)

# Function for email creation.
def main():
    # Depending on the user input. Action will be taken.
    try:
        if start == 'yes' or 'Yes':
            # Clearing the data wit function.
            dataprep(xl_rows, data_list)
            # Adding the data from the file to email .txt file to subject and recipient.
            for i in range(1, xl_rows):
                address = data_list[0][5]
                topic_formated = topic.format(int(data_list[0][0]))
                print('Email nr {} was generated. {} Left. (Press CTRL + C to abort)'.format(i, (int(xl_rows) - i)))
                # Email creation with data extraced from the file. Email class was used here.
                Emailer(address, topic_formated, contents.format(int(data_list[0][0]),
                                                                 next_monday,
                                                                 data_list[0][1],
                                                                 data_list[0][2],
                                                                 data_list[0][4],
                                                                 data_list[0][3],
                                                                 prior_friday)).create()
                del data_list[0]
        # User selected "No" program closes.
        elif start == 'no' or 'No':
            quit()
    # User presed Ctr+c for aborting the script.
    except KeyboardInterrupt:
        print('Script aborted by the user - closing in 10 sec')
        time.sleep(10)
    # Error log for troubleshooting.
    except Exception:
        traceback.print_exc(file=sys.stdout)
    sys.exit(0)


if __name__ == '__main__':
     main()
