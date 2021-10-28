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
logging.basicConfig(level=logging.DEBUG)

import time  # timer
tic = time.perf_counter()  # start timer
logging.info('program start')

import requests
from bs4 import BeautifulSoup



import openpyxl
file = ("C:\\Users\\Juli\\Documents\\GitHub\\simple_stock_3\\trying\\settt.xlsx")
values = []
wb = openpyxl.load_workbook(file)
sheet = wb['Лист1']
for i in range(1,99):
    value = sheet.cell(row=i, column=1).value
    if value is not None:
        values.append(value)

logging.info(values)
LN = len(values)
tikets = []
for i in range(LN):
    tikets.append(str(values[i]))

price = []
price2 = ['0']  #подвечер не придумал как учесть начало итерации с нуля. поэтому костыль. прсто нулевая итерация равна нулю
for i in range(LN):
    x = tikets[i] #tiket
    url = ("https://smart-lab.ru/forum/" + x)
    req = requests.get(url)
    src = req.text
    soup = BeautifulSoup(src, "lxml")
    MSK_trd = soup.find_all(class_="temp_micex_info_item")
    for item in MSK_trd:
        price += ((item.text).split())

lnp=len(price)
for i in range(lnp):
    if i % 2 == 0:
        price2.append(price[i])

logging.info(price2) # returns the ticket price
worksheet = wb['Лист1']
for i in range(LN):
    if i != 0:
        worksheet['B'+ str(i)] = price2[i]
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
