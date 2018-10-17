import datetime, xlrd, sys, traceback, time
from Emailer_class import *

data_list = []
# pliki zewnętrzne. Muszą mieć taką samą nazwe jak podana w kodzie.
file = 'test2.xlsx'
file_email = 'email_body.txt'

# Informacje o pliku. Wyciągniecie z excela pierwszego rzędu i liczby rzędów.
xl_workbook = xlrd.open_workbook(file)
xl_sheet = xl_workbook.sheet_by_name('Sheet1')
xl_rows = xl_sheet.nrows

# Dane do maila. Przygotowanie części tekstowych emaila.
topic = "Questions regarding the delivery of your VELFAC ORDER ({})"
body = open(file_email, 'r', encoding='UTF-8')
contents = body.read()
today = datetime.date.today()
next_monday = today + datetime.timedelta(days=-today.weekday(), weeks=2)
prior_friday = next_monday - datetime.timedelta(days=3)

print('### Before starting please close all outlook windows, except main one ###\n')
start = input('Do you want to start creating emails? There will be ' + str(xl_rows) + ' emails (yes/no)\n')


def dataprep(max_range, storage):
    for i in range(1, max_range):
        xl_row = xl_sheet.row_values(i)
        del xl_row[1]
        del xl_row[6:15]
        storage.append(xl_row)


def main():
    try:
        if start == 'yes' or 'Yes':
            dataprep(xl_rows, data_list)

            for i in range(1, xl_rows):
                address = data_list[0][5]
                topic_formated = topic.format(int(data_list[0][0]))
                print('Email nr {} was generated. {} Left. (Press CTRL + C to abort)'.format(i, (int(xl_rows) - i)))
                Emailer(address, topic_formated, contents.format(int(data_list[0][0]),
                                                                 next_monday,
                                                                 data_list[0][1],
                                                                 data_list[0][2],
                                                                 data_list[0][4],
                                                                 data_list[0][3],
                                                                 prior_friday)).create()
                del data_list[0]
        elif start == 'no' or 'No':
            quit()
    except KeyboardInterrupt:
        print('Script aborted by the user - closing in 10 sec')
        time.sleep(10)
    except Exception:
        traceback.print_exc(file=sys.stdout)
    sys.exit(0)


if __name__ == '__main__':
     main()
