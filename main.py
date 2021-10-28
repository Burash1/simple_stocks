# Get stocks - web scrapping
# get data from the following websites
# https://smart-lab.ru/q/shares/
# excel / google sheet/document  https://docs.google.com/spreadsheets/d/1LOWCcyei54hwfliPAqU8M_J1hWC0UmrfGE0ShXqpfOY/edit?usp=sharing
# columns: company, international name, current price, % change
# Название, тикер, цена, изм %
# Checks !!!! try catch + reconnect
# pip install lxml
# pip install requests
# pip install beautifulsoup4
# pip3 install --upgrade google-api-python-client
# pip3 install oauth2client
# pip install pandas
# pip install openpyxl
import logging
#logging.basicConfig(level=logging.DEBUG)

import time  # timer
tic = time.perf_counter()  # start timer
logging.info('program start')

import requests
from bs4 import BeautifulSoup

# Подключаем библиотеки
#import httplib2
#import apiclient.discovery
#from oauth2client.service_account import ServiceAccountCredentials

#CREDENTIALS_FILE = 'C:/Users/Juli/PycharmProjects/from-pycharm-lerning-f445c7cbaff0.json'  # ключ

# Читаем ключи из файла
#credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])

#httpAuth = credentials.authorize(httplib2.Http()) # Авторизуемся в системе
#service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) # Выбираем работу с таблицами и 4 версию API

#spreadsheet = service.spreadsheets().create(body = {
 #   'properties': {'title': 'второй тестовый документ', 'locale': 'ru_RU'},
 #   'sheets': [{'properties': {'sheetType': 'GRID',
  #                             'sheetId': 0,
  #                             'title': 'Лист номер один',
  #                             'gridProperties': {'rowCount': 100, 'columnCount': 15}}}]
#}).execute()
#spreadsheetId = spreadsheet['spreadsheetId'] # сохраняем идентификатор файла

#driveService = apiclient.discovery.build('drive', 'v3', http = httpAuth) # Выбираем работу с Google Drive и 3 версию API
#access = driveService.permissions().create(
 #   fileId = spreadsheetId,
 #   body = {'type': 'user', 'role': 'writer', 'emailAddress': 'burash.fm@gmail.com'},  # Открываем доступ на редактирование
 #   fields = 'id'
#).execute()

#indexSheet = 'https://docs.google.com/spreadsheets/d/1gkUz2oyHcVZk6oFegSWWf1EBLWwvlZwaJOqOub6JpcA/edit'
#results = service.spreadsheets().values().batchUpdate(spreadsheetId = indexSheet, body = {
 #   "valueInputOption": "USER_ENTERED", # Данные воспринимаются, как вводимые пользователем (считается значение формул)
 #   "data": [
  #      {"range": "Лист номер один!B2:D5",
  #       "majorDimension": "ROWS",     # Сначала заполнять строки, затем столбцы
   #      "values": [
   ##                 ['25', "=6*6", "=sin(3,14/2)"]  # Заполняем вторую строку
    #               ]}
  #  ]
#}).execute()




#Trying to set up a parser and soup

import openpyxl
file = ("C:\\Users\\Juli\\Documents\\GitHub\\simple_stock_3\\trying\\settt.xlsx")
values = []
wb = openpyxl.load_workbook(file)
sheet = wb['Лист1']
for i in range(1,99):
    value = sheet.cell(row=i, column=1).value
    if value is not None:
        values.append(value)

print(values)
LN = len(values)
tikets = []
for i in range(LN):
    tikets.append(str(values[i]))

price = []

for i in range(LN):
    x = tikets[i] #tiket
    url = ("https://smart-lab.ru/forum/" + x)
    req = requests.get(url)
    src = req.text
    soup = BeautifulSoup(src, "lxml")
    MSK_trd = soup.find_all(class_="temp_micex_info_item")
    for item in MSK_trd:
        price += ((item.text).split())


print(price) # returns the ticket price
worksheet = wb['Лист1']
for i in range(LN):
    if i % 2 == 0:
        #???? как вызвать персчет координат в цикле
wb.save("C:\\Users\\Juli\\Documents\\GitHub\\simple_stock_3\\trying\\settt.xlsx") #Сохраняем измененный файл

#class Stocks:
  #  def get_data_from_smart_lab(self):
  #      print('Getting data has started:')
        #try catch, reconnect
        #web scrapping = data(columns + values)
   #     data = 'data'
    #    logging.info('data from smart lab')
    #    logging.debug(data)
     #   return data

   # def make_excel(self, data):
    #    logging.info('make excel')
        #formatting
     #   print('Formatting excel')
        #save excel
     #   print('Saving excel')

   # def make_google_sheet(self, data):
    #   logging.info('make google sheet')
        #formatting
      #  print('Formatting google sheet')
        #save google sheet to your drive
       # print('Saving google sheet')


#def main():
#    x1 = Stocks()
#    data = x1.get_data_from_smart_lab()
#    x1.make_excel(data)
#    x1.make_google_sheet(data)

toc = time.perf_counter()     # stop timer





#if __name__ == '__main__':
#print(f"The calculation took {toc - tic: 0.4f} seconds")
 #   main()
logging.info('Program stop')
