import csv
from outlook import Message, create_csr, create_dsr, Outlook
from datetime import datetime
import os

tasks = []
def get_client_report_format():
    html = ''

def create_template(filename):
    with open(filename, 'r') as file:
        reader = csv.reader(file)
        if len(tasks) > 0:
            return
        for row in reader:
            tasks.append(row[0])

def generate_task_list_for_client():
    create_template('tasks.csv')
    result = '<td>'
    i = 0
    for task in tasks:
        result += '<p>' + str(i + 1) + '. ' + task + '</p>'
        i += 1
    result += '</td>'
    return result

def generate_task_list_for_dsr_1():
    create_template('tasks.csv')
    result = ''
    i=0
    for task in tasks:
        result += '<p>' + str(i+1) + '-</p>'
        result += '<p> Project Task       :    ' + task + '</p>'
        result += '<p> CRM                :     Rahul </p>'
        result += '<p> PM/TL              :     Rahul </p>'
        result += '<p> Current Status     :     Completed </p>'
        i += 1
    result += ''
    return result

def generate_task_list_for_dsr_2():
    create_template('tasks.csv')
    result = ''
    i = 0
    for task in tasks:
        result += '<p>' + str(i + 1) + '. ' + task + '</p>'
        i += 1
    result += ''
    return result

def remove_file(file_path):
    try:
        os.remove(file_path)
    except OSError as e:
        print(file_path + ' not exists')

def main():
    print('Generating client report')
    create_template('tasks.csv')
    remove_file('client_status_report.html');
    remove_file('daily_status_report.html');
    client_report_file = open('client_rep.html', 'r')
    work_for_tommorrow_file = open('work_for_tommorrow.csv', 'r')
    work_for_tommorrow = work_for_tommorrow_file.read()
    client_report_html = client_report_file.read()
    current_date = datetime.today().strftime('%d %B')
    task_list = generate_task_list_for_client()
    client_report_html = client_report_html.replace('TODAYS_DATE', current_date)
    client_report_html = client_report_html.replace('TASKS_LIST', task_list)
    client_report_html = client_report_html.replace('WORK_FOR_TOMORROW', work_for_tommorrow)
    print('Client report generated. Opening Outlook!')
    f = open("client_status_report.html", "a")
    f.write(client_report_html)
    f.close()
    print('Generating DSR')
    daily_status_report_file = open('daily_status_rep.html', 'r')
    daily_status_report_html = daily_status_report_file.read()
    task_list_1 = generate_task_list_for_dsr_1()
    task_list_2 = generate_task_list_for_dsr_2()
    daily_status_report_html = daily_status_report_html.replace('ACTIVITIES_PERFORMED', task_list_1)
    daily_status_report_html = daily_status_report_html.replace('ACTIVITY_LIST', task_list_2)
    print('DSR generated. Opening Outlook!')
    f = open("daily_status_report.html", "a")
    f.write(daily_status_report_html)
    f.close()

main()
outlook = Outlook()
create_dsr(outlook)
create_csr(outlook)