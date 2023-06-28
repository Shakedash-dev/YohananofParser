# Name: YohananofAPI.py
# Author: Shkedo
# Description: Use selenium, gspread and macros to parse the online receipt of Yohananof.
# 			   The goal was to make it easier to share and split the reciept.
# Date: 28.6.23

# Imports.
from html.parser import HTMLParser
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import date
from googleapiclient.discovery import build
from time import sleep

# Constants.
GOOGLE_API_AUTH_JSON_FILE = r"google_api_auth_file.json"
XLSX_PATH = "YohananofReciept.xlsx"
IRRELEVANT_ROWS_START_HERE = 'סהכ הנחות'
GARBAGE_ROW = ['קוד', 'כמות', 'שם']
RAW_RECIEPT_SHEET_NAME = "תשלום גולמי"
SUMMARY_SHEET_NAME = "סיכום"
RAW_EQUATION_SUM = '=IF({0}7="","",ROUNDDOWN(SUMIF(INDIRECT(CONCAT("\'תשלום גולמי\'!",CONCAT(CHAR(64+MATCH({0}7,\'תשלום גולמי\'!1:1,0)),"2"))):INDIRECT(CONCAT("\'תשלום גולמי\'!",CONCAT(CHAR(64+MATCH({0}7,\'תשלום גולמי\'!1:1,0)),MATCH("סהכ הנחות:",\'תשלום גולמי\'!$A:$A,0)))),"v",INDIRECT(CONCAT("\'תשלום גולמי\'!",CONCAT(CHAR(64+MATCH("סהכ מחולק",\'תשלום גולמי\'!1:1,0)),"2"))):INDIRECT(CONCAT("\'תשלום גולמי\'!",CONCAT(CHAR(64+MATCH("סהכ מחולק",\'תשלום גולמי\'!1:1,0)),MATCH("סהכ הנחות:",\'תשלום גולמי\'!$A:$A,0))))),0))'
RAW_EQUATION_PRICE_DEVIDED = "=D{0}/COUNTA(F{0}:R{0})"
ALL_DEVIDED_COLUMN_NAME = "סהכ מחולק"
MERGED_SUM_CELL_LOC = "D10:{0}10"
MERGED_SUM_EQUATION = "=SUM(D8:O8)"

# Change me! (the names that will show on the reciept)
NAMES = ["שקדו", "יובל", "רום", "גהרו"]


def get_full_html(url):
    """Uses selenium to get the fully loaded HTML because the site is dynamic and
    can't be retrieved fully using a simple requests.get request"""

    # Set up Selenium WebDriver (provide path to your webdriver executable)
    driver = webdriver.Chrome()

    # Navigate to the target website
    driver.get(url)

    # Wait for the page to finish loading
    wait = WebDriverWait(driver, 3)  # Adjust the timeout as needed
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))

    # Get the HTML content of the loaded page
    html_content = driver.page_source

    # Close the Selenium WebDriver
    driver.quit()

    return html_content


class TableParser(HTMLParser):
    """This class is based on HTMLParser to parse the content of the online Yohananof reciept."""

    def __init__(self):
        super().__init__()
        self.in_table = False
        self.in_row = False
        self.in_cell = False
        self.table = []
        self.current_row = []
        self.isdiscount = False

    def handle_starttag(self, tag, attrs):
        if tag == "table":
            self.in_table = True
        elif tag == "tr":
            self.in_row = True
            self.current_row = []

            # Check if the row is a discount row.
            if attrs == [("class", "spaceUnder"), ("style", "color:red")]:
                self.isdiscount = True
            else:
                self.isdiscount = False

        elif tag == "td" or tag == "th":
            self.in_cell = True

    def handle_endtag(self, tag):
        if tag == "table":
            self.in_table = False

            # Further parsing of the table.
            # I used a "int i=0;i<=x;i++" style loop because some of the rows are connected to eachother.
            # To be able to access the next element and delete it I needed an index.
            # Also I needed the loop to be a dynamic size, because some of the rows will be deleted mid way.
            row_i = 0
            while row_i < len(self.table):
                # Check if we reached the end of the reciept.
                if IRRELEVANT_ROWS_START_HERE in self.table[row_i][0]:
                    break

                # If the row has only 2 elements before we get to the reciept summary,
                # it must mean that we have a scalable item and the next row will indicate
                # how much it weigh and how much it cost.
                if len(self.table[row_i]) == 2:

                    # So we combine them, into a single readable row.
                    # The format will be ITEM_NAME, ITEM_COST_PER_KG, ITEM_WEIGHT, TOTAL_COST.
                    self.table[row_i] = [self.table[row_i][0], self.table[row_i+1][0], self.table[row_i+1][1], self.table[row_i+1][2]]
                    del self.table[row_i+1]

                # Continue to the next element.
                row_i+=1

            # The last row is garbage.
            if self.table[-1] == GARBAGE_ROW:
                del self.table[-1]

        elif tag == "tr":
            self.in_row = False
            self.table.append(self.current_row)
        elif tag == "td" or tag == "th":
            self.in_cell = False

    def handle_data(self, data):
        if self.in_cell:
            data = data.replace("\"", "")
            data = data.replace("₪", "")

            # If this is a normal row, append normaly.
            if not self.isdiscount:
                self.current_row.append(data.strip())
            
            # If this is a discount row, that means it misses 2 columns. add them after the first one.
            else:
                self.current_row.append(data.strip())
                if len(self.current_row) == 1:
                    self.current_row.append("")
                    self.current_row.append("")

def insert_to_excel(table):
    """Insert the data into a local xlsx file."""
    workbook = Workbook()
    sheet = workbook.active

    for row in table:
        sheet.append(row)

    workbook.save(XLSX_PATH)

def spreadsheet_rtl(spreadsheet):
    """This function changes the direction of the spreadsheet to Right-To-Left."""
    data = {
        "requests": [
            {
                "updateSheetProperties": {
                    "properties": {"rightToLeft": True, "sheetId": spreadsheet.get_worksheet(0).id},
                    "fields": "rightToLeft",
                }
            },
            {
                "updateSheetProperties": {
                    "properties": {"rightToLeft": True, "sheetId": spreadsheet.get_worksheet(1).id},
                    "fields": "rightToLeft",
                }
            }
        ]
    }
    spreadsheet.batch_update(data)

def create_google_sheet():
    """This function uses google API to create a new google sheet according to the program format."""

    # Define the credentials and scope
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_API_AUTH_JSON_FILE, scope)

    # Authenticate and open the Google Sheet
    client = gspread.authorize(credentials)
    spreadsheet = client.create(f'Yohananof Reciept {date.today().strftime("%d-%m-%y")}')

    # Make the spreadsheet public
    spreadsheet.share("", perm_type="anyone", role="writer")

    # Print the URL of the spreadsheet.
    print("URL of the new spreadsheet:", spreadsheet.url)

    # Rename first sheet.
    spreadsheet.sheet1.update_title(RAW_RECIEPT_SHEET_NAME)

    # Create the summary sheet.
    spreadsheet.add_worksheet(title=SUMMARY_SHEET_NAME, rows="100", cols="10")

    # Set the spreadsheet to RTL reading.
    spreadsheet_rtl(spreadsheet)

    # Anchor the first row.
    spreadsheet.sheet1.freeze(rows=1)

    return spreadsheet


def insert_reciept_to_sheet(spreadsheet, reciept):
    """This function will insert reciept data into the raw reciept sheet."""
    raw_reciept_sheet = spreadsheet.get_worksheet(0)
    summary_sheet = spreadsheet.get_worksheet(1)

    raw_reciept_sheet.insert_rows(reciept, row=1)

    # Insert names into the reciept sheet and into summary sheet.
    col_name_index = 0
    for name in NAMES:

        # Update the cell and move on to the next name.
        raw_reciept_sheet.update_cell(1, 6+col_name_index, name)
        summary_sheet.update_cell(7, 4+col_name_index, name)
        col_name_index += 1

    raw_reciept_sheet.update_cell(1,5,ALL_DEVIDED_COLUMN_NAME)

def insert_equations_to_sheet(spreadsheet, num_of_rows):
    """This function will insert all equations into the summary sheet."""
    raw_reciept_sheet = spreadsheet.get_worksheet(0)
    summary_sheet = spreadsheet.get_worksheet(1)
    
    # Create the "devided price" column.
    devided_by_column = []
    for row_i in range(num_of_rows):
        devided_by_column.append(RAW_EQUATION_PRICE_DEVIDED.format(row_i+2))
    range_str = f"E2:E{len(devided_by_column)+1}"

    # Update the cells with formulas in column E
    raw_reciept_sheet.update(range_str, [[equation] for equation in devided_by_column], value_input_option="USER_ENTERED")

    # Insert an equation for each person.
    column_index = 0
    for name in NAMES:
        summary_sheet.update_cell(8, 4+column_index, RAW_EQUATION_SUM.format(chr(ord("D") + column_index)))
        column_index += 1

    # Add mereged SUM cell.
    merge_range = MERGED_SUM_CELL_LOC.format(chr(ord("D") + len(NAMES)-1))
    summary_sheet.merge_cells(merge_range)
    merged_cell = summary_sheet.range(merge_range)[0]
    merged_cell.value = MERGED_SUM_EQUATION
    summary_sheet.update_cells([merged_cell], value_input_option="USER_ENTERED")


def main():
    # Get the HTML data out of the URL using selenium.
    response = get_full_html(input("Insert reciept URL: "))

    # Parse the data into a table (list of lists)
    parser = TableParser()
    parser.feed(response)
    reciept = parser.table

    # Insert the reciept into a new google sheet.
    spreadsheet = create_google_sheet()
    insert_reciept_to_sheet(spreadsheet,reciept)

    # Insert equations into summary sheet.
    number_of_rows = len(reciept)-4
    insert_equations_to_sheet(spreadsheet, number_of_rows)


if __name__ == "__main__":
    main()
