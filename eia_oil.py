import csv
import json
import os
import smtplib
import time
from datetime import datetime
from decimal import Decimal
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import psycopg2 as psycopg2
import requests
from psycopg2.extras import RealDictCursor

error = '503 - Sorry! This file is temporarily unavailable.'
dir_path = os.path.dirname(os.path.realpath(__file__))
current = datetime.now().strftime('OIL_%Y-%m-%d.csv')
start_time = time.time()

#########################
# CHANGE WAIT TIME BELOW
wait_sleep_time = 10
########################

path = f'{dir_path}{os.sep}Data{os.sep}Oil{os.sep}'
# investing_output_path = f'{dir_path}{os.sep}Data{os.sep}investing_output.csv'
if not os.path.exists(path):
    os.makedirs(path)


def database():
    """Function to connect to the PostgreSQL Database

    Returns:
        [object]: Cursor object
    """
    # Read Credentials.json
    json_data = dict(json.load(open(f'Control{os.sep}Credentials.json')))

    # Assign values from control.csv
    db = json_data["database"]
    user = json_data["user"]
    pw = json_data["password"]
    host = json_data["host"]
    port = json_data["port"]

    # Connecting to DB and Cursor creation
    conn = psycopg2.connect(database=db, user=user, password=pw, host=host, port=port)
    conn.autocommit = True

    return conn.cursor(cursor_factory=RealDictCursor)


# def is_valid(content):
#     temp_folder = f'{dir_path}{os.sep}.temp{os.sep}'
#     if not os.path.exists(temp_folder):
#         os.mkdir(temp_folder)
#     open(temp_folder + current, 'wb').write(content)
#     workbook = xlrd.open_workbook(temp_folder + current)
#     worksheet = workbook.sheet_by_index(0)
#     first_row_value = worksheet.row(1)[0].value
#     if first_row_value == error:
#         print(worksheet.row(1)[0].value)
#         shutil.rmtree(temp_folder)
#         return False
#     shutil.rmtree(temp_folder)
#     return True


def read_csv_and_get_data():
    csv_reader = csv.reader(open(path + current, 'r', newline='', encoding='ISO-8859-1'))
    next(csv_reader)
    result = {}
    for row in csv_reader:
        if row[0] == 'Crude Oil':
            result['Crude Oil'] = [row[3], row[4], row[7]]
        elif row[0] == 'Commercial (Excluding SPR)':
            result['Commercial (Excluding SPR)'] = [row[3], row[4], row[7]]
        elif row[0] == 'Total Motor Gasoline':
            result['Total Motor Gasoline'] = [row[3], row[4], row[7]]
        elif row[0] == 'Distillate Fuel Oil':
            result['Distillate Fuel Oil'] = [row[3], row[4], row[7]]
        elif row[0] == 'Total Stocks (Including SPR)':
            result['Total Stocks (Including SPR)'] = [row[3], row[4], row[7]]
        elif row[0] == 'Total Stocks (Excluding SPR)':
            result['Total Stocks (Excluding SPR)'] = [row[3], row[4], row[7]]
        elif row[0] == '> 500 ppm sulfur':
            result['Heating Oil'] = [row[3], row[4], row[7]]
    return result


def is_null(value):
    if value is None:
        return Decimal(0.00)
    return value * 1000


def fetch_eco_view_data():
    cursor = database()
    sql_query = """SELECT event_date,
                    event_time,
                    event_name,
                    actual,
                    forecast,
                    previous
                    FROM tradingview_eco
                    WHERE event_date >= CURRENT_DATE - 5 AND
                    event_date <= CURRENT_DATE AND
                    event_name like %(event_name_)s
                    ORDER BY event_date, event_time, event_name;
        """

    # data = cursor.fetchall()
    result = {}

    cursor.execute(sql_query, {'event_name_': 'EIA Wkly Crude Stk%'})
    weekly = cursor.fetchone()
    if weekly is None:
        raise Exception('No Data for EIA Wkly Crude Stk%')
    cursor.execute(sql_query, {'event_name_': 'API wkly crude Stk%'})
    api = cursor.fetchone()
    if api is None:
        raise Exception('No Data for API wkly crude Stk%')
    result['Crude Oil'] = [is_null(weekly['forecast']), is_null(weekly['previous']), is_null(api['actual'])]

    cursor.execute(sql_query, {'event_name_': 'EIA Wkly Gsln Stk%'})
    weekly = cursor.fetchone()
    if weekly is None:
        raise Exception('No Data for EIA Wkly Gsln Stk%')
    cursor.execute(sql_query, {'event_name_': 'API Wkly gsln stk%'})
    api = cursor.fetchone()
    if weekly is None:
        raise Exception('No Data for API Wkly gsln stk%')
    result['Total Motor Gasoline'] = [is_null(weekly['forecast']), is_null(weekly['previous']), is_null(api['actual'])]

    cursor.execute(sql_query, {'event_name_': 'EIA Weekly Heatoil Stock%'})
    weekly = cursor.fetchone()
    if weekly is None:
        raise Exception('No Data for EIA Weekly Heatoil Stock%')
    cursor.execute(sql_query, {'event_name_': 'API weekly heating oil%'})
    api = cursor.fetchone()
    if api is None:
        raise Exception('No Data for API weekly heating oil%')
    result['Heating Oil'] = [is_null(weekly['forecast']), is_null(weekly['previous']), is_null(api['actual'])]

    cursor.execute(sql_query, {'event_name_': 'EIA Wkly Dist. Stk%'})
    weekly = cursor.fetchone()
    if weekly is None:
        raise Exception('No Data for EIA Wkly Dist. Stk%')
    cursor.execute(sql_query, {'event_name_': 'API Wkly dist. Stk%'})
    api = cursor.fetchone()
    if api is None:
        raise Exception('No Data for API Wkly dist. Stk%')
    result['Distillate Fuel Oil'] = [is_null(weekly['forecast']), is_null(weekly['previous']), is_null(api['actual'])]

    # result['Total Stocks (Including SPR)'] = [row[3], row[4], row[7]]
    # result['Total Stocks (Excluding SPR)'] = [row[3], row[4], row[7]]
    return result


# def read_investing_output_csv_and_get_data():
#     csv_reader = csv.reader(open(investing_output_path, 'r', newline='', encoding='ISO-8859-1'))
#     next(csv_reader)
#     result = {}
#     for row in csv_reader:
#         if row[4] == 'API Weekly Crude Oil Stock':
#             result['API Weekly Crude Oil Stock'] = row[5]
#         elif row[4] == 'Crude Oil Inventories':
#             result['Crude Oil Inventories'] = row[6]
#         elif row[4] == 'EIA Weekly Distillates Stocks':
#             result['EIA Weekly Distillates Stocks'] = row[6]
#         elif row[4] == 'Gasoline Inventories':
#             result['Gasoline Inventories'] = row[6]
#     return result


# return red if positive else green
def color_based_on_value(value):
    if float(value) == 0:
        return 'white'
    else:
        return 'red' if float(value) > 0 else 'green'


# return as format #.##%
def to_percentage(value):
    return "{:.2f}%".format(float(value))


# return
def to_double(value):
    return float(value) * 1000


# returns as format ###,###
def to_currency_format_without_decimals(value):
    return '{:,.0f}'.format(float(value) * 1000)


def to_forecast_api_value(value):
    return float(value)


def to_str_forecast_api_value(value):
    return '{:,.2f}'.format(to_forecast_api_value(value))


def formula_api(diff, api):
    return ((float(diff)) - (float(api))) / abs(float(api))


def formula_forecast(diff, forecast, api):
    return ((float(diff)) - (float(forecast))) / abs(float(api))


def formula_last_wk(diff, previous, api):
    return ((float(diff)) - (float(previous))) / abs(float(api))


def generate_table(csv_result):
    eco_view_data = fetch_eco_view_data()

    crud_oil_spr_per_forecast = formula_forecast(to_double(csv_result['Commercial (Excluding SPR)'][0]),
                                                 eco_view_data['Crude Oil'][0], eco_view_data['Crude Oil'][2])
    crud_oil_spr_per_api = formula_api(to_double(csv_result['Commercial (Excluding SPR)'][0]),
                                       eco_view_data['Crude Oil'][2])
    crud_oil_spr_per_last_wk = formula_last_wk(to_double(csv_result['Commercial (Excluding SPR)'][0]),
                                               eco_view_data['Crude Oil'][1], eco_view_data['Crude Oil'][2])

    gasoline_per_forecast = formula_forecast(to_double(csv_result['Total Motor Gasoline'][0]),
                                             eco_view_data['Total Motor Gasoline'][0],
                                             eco_view_data['Total Motor Gasoline'][2])
    gasoline_per_api = formula_api(to_double(csv_result['Total Motor Gasoline'][0]),
                                   eco_view_data['Total Motor Gasoline'][2])
    gasoline_per_last_wk = formula_last_wk(to_double(csv_result['Total Motor Gasoline'][0]),
                                           eco_view_data['Total Motor Gasoline'][1],
                                           eco_view_data['Total Motor Gasoline'][2])

    heating_oil_forecast = formula_forecast(to_double(csv_result['Heating Oil'][0]), eco_view_data['Heating Oil'][0],
                                            eco_view_data['Heating Oil'][2])
    heating_oil_per_api = formula_api(to_double(csv_result['Heating Oil'][0]), eco_view_data['Heating Oil'][2])
    heating_oil_per_last_wk = formula_last_wk(to_double(csv_result['Heating Oil'][0]), eco_view_data['Heating Oil'][1],
                                              eco_view_data['Heating Oil'][2])

    # total_stock_ex_spr_per_api = formula(to_double(csv_result['Total Stocks (Excluding SPR)'][0]), investing_output_result['API Weekly Crude Oil Stock'])

    distillate_forecast = formula_forecast(to_double(csv_result['Distillate Fuel Oil'][0]),
                                           eco_view_data['Distillate Fuel Oil'][0],
                                           eco_view_data['Distillate Fuel Oil'][2])
    distillate_per_api = formula_api(to_double(csv_result['Heating Oil'][0]), eco_view_data['Distillate Fuel Oil'][2])
    distillate_per_last_wk = formula_last_wk(to_double(csv_result['Heating Oil'][0]),
                                             eco_view_data['Distillate Fuel Oil'][1],
                                             eco_view_data['Distillate Fuel Oil'][2])

    html = f"""<table style="width: 100%;border-collapse: collapse;border: 5px solid black; border-style: solid" border="5" cellpadding="5" >
    <tbody>
    <tr style="font-weight:bold;outline: thin solid;">
        <td><strong>Stub</strong></td> 
        <td>&nbsp;<strong>Difference</strong></td>
        <td>&nbsp;<strong>Forecast</strong></td>
        <td>&nbsp;<strong>API</strong></td>
        <td>&nbsp;<strong>% vs Forecast</strong></td>
        <td>&nbsp;<strong>% vs Last Wk</strong></td>
        <td><strong>% vs API</strong></td>
        <td>&nbsp;<strong>% vs Last Yr</strong></td>
    </tr>
    <tr style="font-weight:bold">
        <td>&nbsp;Crude Oil Ex-SPR</td>
        <td style="background-color: {color_based_on_value(csv_result['Commercial (Excluding SPR)'][0])};">&nbsp;{to_currency_format_without_decimals(csv_result['Commercial (Excluding SPR)'][0])}</td>
        <td style="background-color: {color_based_on_value(eco_view_data['Crude Oil'][0])};">&nbsp;{to_str_forecast_api_value(eco_view_data['Crude Oil'][0])}</td>
        <td style="background-color: {color_based_on_value(eco_view_data['Crude Oil'][2])};">&nbsp;{to_str_forecast_api_value(eco_view_data['Crude Oil'][2])}</td> 
        <td style="background-color: {color_based_on_value(crud_oil_spr_per_forecast)};">&nbsp;{to_percentage(crud_oil_spr_per_forecast)}</td>
        <td style="background-color: {color_based_on_value(crud_oil_spr_per_last_wk)};">&nbsp;{to_percentage(crud_oil_spr_per_last_wk)}</td>
        <td style="background-color: {color_based_on_value(crud_oil_spr_per_api)};">&nbsp;{to_percentage(crud_oil_spr_per_api)}</td>
        <td style="background-color: {color_based_on_value(csv_result['Commercial (Excluding SPR)'][2])};">&nbsp;{to_percentage(csv_result['Commercial (Excluding SPR)'][2])}</td>
    </tr>
    <tr style="font-weight:bold">
        <td>&nbsp;Gasoline</td>
        <td style="background-color: {color_based_on_value(csv_result['Total Motor Gasoline'][0])};">&nbsp;{to_currency_format_without_decimals(csv_result['Total Motor Gasoline'][0])}</td>
        <td style="background-color: {color_based_on_value(eco_view_data['Total Motor Gasoline'][0])};">&nbsp;{to_str_forecast_api_value(eco_view_data['Total Motor Gasoline'][0])}</td>
        <td style="background-color: {color_based_on_value(eco_view_data['Total Motor Gasoline'][2])};">&nbsp;{to_str_forecast_api_value(eco_view_data['Total Motor Gasoline'][2])}</td>
        <td style="background-color: {color_based_on_value(gasoline_per_forecast)};">&nbsp;{to_percentage(gasoline_per_forecast)}</td>
        <td style="background-color: {color_based_on_value(gasoline_per_last_wk)};">&nbsp;{to_percentage(gasoline_per_last_wk)}</td>
        <td style="background-color: {color_based_on_value(gasoline_per_api)};">&nbsp;{to_percentage(gasoline_per_api)}</td>
        <td style="background-color: {color_based_on_value(csv_result['Total Motor Gasoline'][2])};">&nbsp;{to_percentage(csv_result['Total Motor Gasoline'][2])}</td>
    </tr>
    <tr style="font-weight:bold">
        <td>&nbsp;Heating Oil</td>
        <td style="background-color: {color_based_on_value(csv_result['Heating Oil'][0])};">&nbsp;{to_currency_format_without_decimals(csv_result['Heating Oil'][0])}</td>
        <td style="background-color: {color_based_on_value(eco_view_data['Heating Oil'][0])};">&nbsp;{to_str_forecast_api_value(eco_view_data['Heating Oil'][0])}</td>
        <td style="background-color: {color_based_on_value(eco_view_data['Heating Oil'][2])};">&nbsp;{to_str_forecast_api_value(eco_view_data['Heating Oil'][2])}</td>
        <td style="background-color: {color_based_on_value(heating_oil_forecast)};">&nbsp;{to_percentage(heating_oil_forecast)}</td>
        <td style="background-color: {color_based_on_value(heating_oil_per_last_wk)};">&nbsp;{to_percentage(heating_oil_per_last_wk)}</td>
        <td style="background-color: {color_based_on_value(heating_oil_per_api)};">&nbsp;{to_percentage(heating_oil_per_api)}</td>
        <td style="background-color: {color_based_on_value(csv_result['Heating Oil'][2])};">&nbsp;{to_percentage(csv_result['Heating Oil'][2])}</td>
    </tr>
    <tr>
        <td>&nbsp;Crude Oil Total</td>
        <td style="background-color: {color_based_on_value(csv_result['Crude Oil'][0])};">&nbsp;{to_currency_format_without_decimals(csv_result['Crude Oil'][0])}</td>
        <td >&nbsp;</td>
        <td >&nbsp;</td>
        <td >&nbsp;</td>
        <td style="background-color: {color_based_on_value(csv_result['Crude Oil'][1])};">&nbsp;{to_percentage(csv_result['Crude Oil'][1])}</td>
        <td >&nbsp;</td>
        <td style="background-color: {color_based_on_value(csv_result['Crude Oil'][2])};">&nbsp;{to_percentage(csv_result['Crude Oil'][2])}</td>
    </tr>
    <tr>
        <td style="font-weight:bold">&nbsp;Total Stocks Ex-SPR</td>
        <td style="background-color: {color_based_on_value(csv_result['Total Stocks (Excluding SPR)'][0])};">&nbsp;{to_currency_format_without_decimals(csv_result['Total Stocks (Excluding SPR)'][0])}</td>
        <td >&nbsp;</td>
        <td >&nbsp;</td>
        <td >&nbsp;</td>
        <td style="background-color: {color_based_on_value(csv_result['Total Stocks (Excluding SPR)'][1])};">&nbsp;{to_percentage(csv_result['Total Stocks (Excluding SPR)'][1])}</td>
        <td >&nbsp;</td>
        <td style="background-color: {color_based_on_value(csv_result['Total Stocks (Excluding SPR)'][2])};">&nbsp;{to_percentage(csv_result['Total Stocks (Excluding SPR)'][2])}</td>
    </tr>
    <tr style="font-weight:bold">
        <td >&nbsp;Total Stocks</td>
        <td style="background-color: {color_based_on_value(csv_result['Total Stocks (Including SPR)'][0])};">&nbsp;{to_currency_format_without_decimals(csv_result['Total Stocks (Including SPR)'][0])}</td>
        <td >&nbsp;</td>
        <td >&nbsp;</td>
        <td >&nbsp;</td>
        <td style="background-color: {color_based_on_value(csv_result['Total Stocks (Including SPR)'][1])};">&nbsp;{to_percentage(csv_result['Total Stocks (Including SPR)'][1])}</td>
        <td >&nbsp;</td>
        <td style="background-color: {color_based_on_value(csv_result['Total Stocks (Including SPR)'][2])};">&nbsp;{to_percentage(csv_result['Total Stocks (Including SPR)'][2])}</td>
    </tr>
    <tr>
        <td>&nbsp;EIA Weekly Distillates Stocks</td>
        <td style="background-color: {color_based_on_value(csv_result['Distillate Fuel Oil'][0])};">&nbsp;{to_currency_format_without_decimals(csv_result['Distillate Fuel Oil'][0])}</td>
        <td style="background-color: {color_based_on_value(eco_view_data['Distillate Fuel Oil'][0])};">&nbsp;{to_str_forecast_api_value(eco_view_data['Distillate Fuel Oil'][0])}</td>
        <td style="background-color: {color_based_on_value(eco_view_data['Distillate Fuel Oil'][2])};">&nbsp;{to_str_forecast_api_value(eco_view_data['Distillate Fuel Oil'][2])}</td>
        <td style="background-color: {color_based_on_value(distillate_forecast)};">&nbsp;{to_percentage(distillate_forecast)}</td>
        <td style="background-color: {color_based_on_value(distillate_per_last_wk)};">&nbsp;{to_percentage(distillate_per_last_wk)}</td>
        <td style="background-color: {color_based_on_value(distillate_per_api)};">&nbsp;{to_percentage(distillate_per_api)}</td>
        <td style="background-color: {color_based_on_value(csv_result['Distillate Fuel Oil'][2])};">&nbsp;{to_percentage(csv_result['Distillate Fuel Oil'][2])}</td>
    </tr>
    </tbody>
    </table>"""
    return html


def send_email(body):
    control = dict(csv.reader(open(f'Control{os.sep}Control.csv')))

    msg = MIMEMultipart()
    msg['From'] = control['Email SMTP ID']
    msg['To'] = control['Email TO Email ID']
    msg['Subject'] = ':EIA_Oil:'
    msg.attach(MIMEText(body, 'html'))
    txt = msg.as_string()
    try:
        if control['Require logon using Secure Password Authentication (SPA)'].lower() == 'no':
            server = str(control['Email SMTP Server Name / IP Address']) + ":" + str(control['Email SMTP Server Port'])
            s = smtplib.SMTP(server)
        else:
            s = smtplib.SMTP(control['Email SMTP Server Name / IP Address'], int(control['Email SMTP Server Port']))
            s.starttls()

        s.login(control['Email SMTP ID'], control['Email SMTP Password'])
        s.sendmail(control['Email SMTP ID'], control['Email TO Email ID'], txt)
        s.quit()
        print('Sent mail')

    except smtplib.SMTPException as e:
        print(e)
        print("Email couldn't be sent")


def save_pdf(body):
    folder = f'{dir_path}{os.sep}PDF{os.sep}'
    if not os.path.exists(folder):
        os.mkdir(folder)
    with open(folder + 'EIA Oil Report.htm', 'w') as file:
        file.write(body)
        print(f'{folder}EIA Oil Report.htm saved!')


def save_pdf_and_send_email():
    # investing_output_result = read_investing_output_csv_and_get_data()
    csv_result = read_csv_and_get_data()
    body = generate_table(csv_result)
    save_pdf(body)
    send_email(body)


while True:
    try:
        response = requests.get('https://ir.eia.gov/wpsr/table1.csv', allow_redirects=True)
        if response.status_code == 200:
            open(path + current, 'wb').write(response.content)
            print(f'Found {path + current} with Data !!!')
            save_pdf_and_send_email()
            break
        else:
            print(error)
        print(f'Trying again in {wait_sleep_time} seconds!')
        time.sleep(wait_sleep_time)
    except Exception as e:
        print(e)
        print("Trying Again in 5 Seconds")
        if time.time() - start_time > 600:
            print('10 minutes have passed exiting')
            exit()
        time.sleep(wait_sleep_time)
        continue
exit()
