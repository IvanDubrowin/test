import os
import textwrap
from datetime import datetime
from time import sleep
import xlsxwriter
from selenium.webdriver import Firefox
from selenium.webdriver.firefox.options import Options
from bs4 import BeautifulSoup


TARGET_URL = 'http://lists4priemka.fa.ru/listabits.aspx?fl=0&tl=%D0%B1%D0%BA%D0%BB&le=%D0%92%D0%9F%D0%9E'
NEXT_PAGE_ONCLICK = "aspxGVPagerOnClick('ASPxGridView1','PBN');"
PAGINATE_BUTTON_CLASS = 'dxpButton'
TITLE_TABLE_CLASS = 'dxgvHeader'
BODY_TABLE_ID = 'ASPxGridView1_DXMainTable'
DATA_ROW = 'dxgvDataRow'
ROOT_DIR = os.path.dirname(os.path.realpath(__file__))
GECKODRIVER_PATH = os.path.join(ROOT_DIR, 'geckodriver')
PAGE_PROGRESS = 'dxpSummary'


class TableParser:
    def __init__(self, url):
        self.url = url
        self.browser = self._setup_firefox()

    title = []
    body = []
    next_page = False

    def _parse(self):
        self.browser.get(self.url)
        while True:
            soup = BeautifulSoup(self.browser.page_source, 'html.parser')
            if not self.title:
                self.title = [td.get_text() for td in soup.find_all('td', class_=TITLE_TABLE_CLASS)]
            table = soup.find('table', id=BODY_TABLE_ID)
            print(soup.find('td', class_=PAGE_PROGRESS).get_text())
            if table:
                for row in table.find_all('tr', class_=DATA_ROW):
                    if row:
                        data_row = [td.get_text() or ' ' for td in row.find_all('td')]
                        if self.body:
                            if self.body[-1] != [data_row]:
                                self.body += [data_row]
                        else:
                            self.body += [data_row]
            buttons = soup.find_all('td', class_=PAGINATE_BUTTON_CLASS)
            if buttons:
                for button in buttons:
                    if button.get('onclick') == NEXT_PAGE_ONCLICK:
                        self.browser.execute_script('javascript:{}'.format(NEXT_PAGE_ONCLICK))
                        sleep(3)
                        self.next_page = True
            if self.next_page:
                self.next_page = False
                continue
            self.browser.close()
            break

    @staticmethod
    def _setup_firefox():
        opts = Options()
        opts.set_headless()
        assert opts.headless
        browser = Firefox(executable_path=GECKODRIVER_PATH, options=opts)
        return browser

    def get_table(self):
        self._parse()
        return [self.title] + self.body


def convert_to_xlsx(data):
    title_row = data[0]
    body_rows = data[1::]

    row_index = 0
    col_index = 0

    workbook = xlsxwriter.Workbook(
        '{}.xlsx'.format(
            datetime.strftime(datetime.now(), "%d.%m.%Y %H.%M.%S")
            )
        )
    worksheet = workbook.add_worksheet()

    row_format = workbook.add_format({
        'align': 'left',
        'valign': 'vcenter',
        'font_size': 10,
        'border': 1,
        'text_wrap': True
    })

    title_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'bold': True
    })

    def cell(text):
        base_height = 18
        wrap_text = '\n'.join(textwrap.wrap(text))

        return {
            'text': wrap_text,
            'width': len(wrap_text.split('\n')[0]),
            'height': len(wrap_text.split('\n')) * base_height
            }

    def setup_column_width(title, body):
        indent = 5
        columns_width = [len(t) + indent for t in title]
        for row in body:
            for i, c in enumerate(row):
                item = cell(c)
                if columns_width[i] < item['width']:
                    columns_width[i] = item['width'] + indent
        return columns_width

    width_map = setup_column_width(title_row, body_rows)

    for i, title in enumerate(title_row):
        worksheet.set_column(col_index, col_index, width_map[i])
        worksheet.write(row_index, col_index, title, title_format)
        col_index += 1
    row_index += 1

    for row in body_rows:
        col_index = 0
        max_height = 0
        for text in row:
            item = cell(text)
            worksheet.write(row_index, col_index, item['text'], row_format)
            if max_height < item['height']:
                max_height = item['height']
            worksheet.set_row(row_index, max_height)
            col_index += 1
        row_index += 1

    workbook.close()


def run():
    data = TableParser(TARGET_URL).get_table()
    if data:
        print('Загрузка данных завершена!')
        convert_to_xlsx(data)
        print('Запись в файл завершена!')
    else:
        print('Не удалось завершить загрузку!')


if __name__ == '__main__':
    while True:
        arg = input('Для запуска программы введите "start", а для выхода "exit": ')
        if arg == 'start':
            run()
        elif arg == 'exit':
            break
