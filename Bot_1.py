import time
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.options import Options
import gspread
import datetime
from oauth2client.service_account import ServiceAccountCredentials
from tabulate import tabulate
from selenium.webdriver.common.action_chains import ActionChains

# Konfigurasi Google Spreadsheet
spreadsheet_key = 'xxx'
credentials = ServiceAccountCredentials.from_json_keyfile_name('xxx', ['https://www.googleapis.com/auth/spreadsheets'])
client = gspread.authorize(credentials)
sheet = client.open_by_key(spreadsheet_key).sheet1

def show_data():
    # Mendapatkan data dari spreadsheet
    data = sheet.get_all_values()
    headers = data[0]
    rows = data[1:]

    # Kolom dan baris yang ingin ditampilkan
    selected_columns = [1]
    selected_rows = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29]

    # Filter kolom yang dipilih
    headers = [headers[col] for col in selected_columns]
    rows = [[row[col] for col in selected_columns] for row in rows]

    # Filter baris yang dipilih
    rows = [row for idx, row in enumerate(rows) if idx in selected_rows]

    return headers, rows

def show_data2():
    # Mendapatkan data dari spreadsheet
    data = sheet.get_all_values()
    headers = data[0]
    rows = data[1:]

    # Kolom dan baris yang ingin ditampilkan
    selected_columns = [1]
    selected_rows = [ 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45]

    # Filter kolom yang dipilih
    headers = [headers[col] for col in selected_columns]
    rows = [[row[col] for col in selected_columns] for row in rows]

    # Filter baris yang dipilih
    rows = [row for idx, row in enumerate(rows) if idx in selected_rows]

    return headers, rows
    
def show_data3():
    # Mendapatkan data dari spreadsheet
    data = sheet.get_all_values()
    headers = data[0]
    rows = data[1:]

    # Kolom dan baris yang ingin ditampilkan
    selected_columns = [1]
    selected_rows = [ 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60]

    # Filter kolom yang dipilih
    headers = [headers[col] for col in selected_columns]
    rows = [[row[col] for col in selected_columns] for row in rows]

    # Filter baris yang dipilih
    rows = [row for idx, row in enumerate(rows) if idx in selected_rows]

    return headers, rows

# Inisialisasi ChromeDriver
options = Options()
options.add_argument('--headless')  # Jalankan Chrome dalam mode headless (tanpa tampilan browser)
driver = webdriver.Edge()
driver.get("https://web.whatsapp.com")
time.sleep(20)  # Tunggu 20 detik untuk login manual

# Fungsi untuk mengirim pesan ke WhatsApp
def send_message_to_whatsapp(headers, rows):
    # Ganti 'Nama Grup WA' dengan nama grup WhatsApp yang dituju
    contact = 'xxx'

    # Temukan elemen input teks untuk pencarian grup
    input_box_search = WebDriverWait(driver, 50).until(lambda driver: driver.find_element("xpath", "/html/body/div[1]/div/div/div[4]/div/div[1]/div/div/div[2]/div/div[1]/p"))
    input_box_search.click()
    time.sleep(2)
    input_box_search.send_keys(contact)
    time.sleep(2)
    selected_contact = driver.find_element(by=By.XPATH, value="//span[@title='"+contact+"']")
    selected_contact.click()
    inp_xpath = '//*[@id="side"]/div[1]/div/div/div[2]/div/div[1]/p'
    input_box = driver.find_element(by=By.XPATH, value=inp_xpath)
    time.sleep(2)

    # Format the rows as a table
    table = tabulate(rows, headers)

    # Split the table into separate lines
    message_lines = table.split('\n')

    # Send each line as a WhatsApp message
    for line in message_lines:
        ActionChains(driver).send_keys(line).perform()
        ActionChains(driver).key_down(Keys.SHIFT).key_down(Keys.ENTER).key_up(Keys.SHIFT).key_up(Keys.ENTER).perform()

    # Press Enter to send the message
    ActionChains(driver).send_keys(Keys.RETURN).perform()
    time.sleep(2)

# Fungsi untuk menjalankan show_data secara otomatis setiap jam
def schedule_show_data():
    already_executed = False

    while True:
        now = datetime.datetime.now()
        if now.minute == 0 and 8 <= now.hour <= 23 and not already_executed:
            # Panggil show_data dan simpan outputnya dalam variabel 'output'
            headers, rows = show_data()
            send_message_to_whatsapp(headers, rows)
            
            # Panggil show_data2 dan simpan outputnya dalam variabel 'output2'
            headers2, rows2 = show_data2()
            send_message_to_whatsapp(headers2, rows2)
            
            # Panggil show_data3 dan simpan outputnya dalam variabel 'output3'
            headers3, rows3 = show_data3()
            send_message_to_whatsapp(headers3, rows3)
            
            already_executed = True
        elif now.minute != 0:
            already_executed = False

# Jalankan fungsi schedule_show_data
schedule_show_data()





