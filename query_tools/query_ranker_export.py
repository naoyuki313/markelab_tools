import gspread
from oauth2client.service_account import ServiceAccountCredentials
import requests
import bs4

ROW_START = 2
TARGET = 2
QUERY_COL = 3
FIRST_TITLE_START = 4
SECOND_TITLE = 5
THIRD_TITLE = 6
FOURTH_TITLE = 7
ROOT_URL = 'https://www.google.co.jp/search?q='

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

credentials = ServiceAccountCredentials.from_json_keyfile_name('tidy-visitor-170300-ba74d7e1dcf8.json', scope)
gc = gspread.authorize(credentials)
wks = gc.open('クエリ出力ツールβ').worksheet('クエリ順位')

#The number of query on the sheet
query_cnt = len(wks.col_values(QUERY_COL))

for i in range(ROW_START, query_cnt + 1):
    if wks.cell(i, TARGET).value != "":
        #get the target queries
        trgt_query = wks.cell(i, QUERY_COL).value
        print(trgt_query)

        #send the get request
        res = requests.get(ROOT_URL + trgt_query)
        soup = bs4.BeautifulSoup(res.text, "html.parser")
        title_elem = soup.select('.r > a')
        title_index = 0
        for j in range(len(title_elem)):
            title = title_elem[j].get_text()
            url_source = title_elem[j].get('href').replace('/url?q=', '')
            url = url_source.split("&sa=")[0]
            wks.update_cell(i, title_index + FIRST_TITLE_START, title)
            wks.update_cell(i, title_index + FIRST_TITLE_START + 1, url)
            print(title)
            print(url)
            title_index = title_index + 2


#wks.update_acell('A2', 'Hello World!')
#print(wks.acell('A2'))
