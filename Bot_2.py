import gspread
from oauth2client.service_account import ServiceAccountCredentials
from telegram.ext import Updater, CommandHandler
from tabulate import tabulate
from telegram import Update
import time
import datetime
import schedule

# Konfigurasi Google Spreadsheet
spreadsheet_key = 'xxx'
credentials = ServiceAccountCredentials.from_json_keyfile_name('xxx', ['https://www.googleapis.com/auth/spreadsheets'])
client = gspread.authorize(credentials)
sheet = client.open_by_key(spreadsheet_key).sheet1
sheet2 = client.open_by_key(spreadsheet_key).get_worksheet(1)
sheet3 = client.open_by_key(spreadsheet_key).get_worksheet(2)

# Fungsi untuk menangani perintah /start
def start(update, context):
    update.message.reply_text('Halo! Bot sedang berjalan.')

# Fungsi untuk menangani perintah /data
def show_data(context):
    # Mendapatkan data dari spreadsheet
    data = sheet.get_all_values()
    headers = data[0]
    rows = data[1:]

    # Kolom dan baris yang ingin ditampilkan
    selected_columns = [1]
    selected_rows = [ 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29]

    # Filter kolom yang dipilih
    headers = [headers[col] for col in selected_columns]
    rows = [[row[col] for col in selected_columns] for row in rows]

    # Filter baris yang dipilih
    rows = [row for idx, row in enumerate(rows) if idx in selected_rows]
    
    # Menghilangkan garis-garis pada awal data
    rows = [[cell.strip('- ') for cell in row] for row in rows]

    table = tabulate(rows, headers)
    context.bot.send_message(chat_id='-xxx', text=f'```\n{table}\n```', parse_mode='MarkdownV2')
    
def show_data2(context):
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
    
    # Menghilangkan garis-garis pada awal data
    rows = [[cell.strip('- ') for cell in row] for row in rows]

    table = tabulate(rows, headers)
    context.bot.send_message(chat_id='-xxx', text=f'```\n{table}\n```', parse_mode='MarkdownV2')
    
def show_data3(context):
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
    
    # Menghilangkan garis-garis pada awal data
    rows = [[cell.strip('- ') for cell in row] for row in rows]

    table = tabulate(rows, headers)
    context.bot.send_message(chat_id='-xxx', text=f'```\n{table}\n```', parse_mode='MarkdownV2')
    
def show_data_report(context):
    # Mendapatkan data dari spreadsheet
    data = sheet3.get_all_values()
    headers = data[0]
    rows = data[1:]

    # Kolom yang ingin ditampilkan
    selected_columns = [0]
    
    # Filter kolom yang dipilih
    headers = [headers[col] for col in selected_columns]
    rows = [[row[col] for col in selected_columns] for row in rows]

    # Filter baris yang dipilih
    rows = [row for row in rows if any(row)]

    table = tabulate(rows, headers)
    context.bot.send_message(chat_id='xxx', text=f'```\n{table}\n```', parse_mode='MarkdownV2')

def record_data(context):
    # Define Range dari Column (Total WO)
    column_range = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]
    starting_row = 4
    current_datetime = datetime.datetime.now()
    current_date = current_datetime.strftime('%m/%d/%Y')

    # mencari column dan row yang kosong
    empty_row = None
    empty_col = None

    # membaca row semua sekaligus
    all_rows = sheet2.get_all_values()

    for row in range(starting_row, len(all_rows) + 1):
        row_values = all_rows[row-1]
        for col_index in column_range:
            if col_index < len(row_values) and row_values[col_index] == '':
                empty_row = row
                empty_col = col_index
                break
        if empty_row is not None:
            break

    # jika tidak ada cell yang kosong dari range yang tertentu maka dibuat row baru
    if empty_row is None:
        empty_row = len(all_rows) + 1
        empty_col = 0
        # masukan/insert row baru
        sheet2.insert_row([''] * len(column_range), empty_row)
        # update column untuk row baru
        starting_row += 1
        
    # Insert/masukan tanggal sekarang di cell pertama di row baru
    sheet2.update_cell(empty_row, 1, current_date)

    # dapatkan data dari sheet 1, row 7, dan column 2
    data2 = sheet.cell(8, 3).value

    # masukan data kedalam cell yang kosong
    while True:
        try:
            sheet2.update_cell(empty_row, empty_col + 1, data2)
            break
        except gspread.exceptions.APIError as e:
            if e.response.status_code == 429: 
                time.sleep(1) 
            else:
                raise e

    # Define Range dari column2 (Total PS)
    column_range2 = [16,17,18,19,20,21,22,23,24,25,26,27,28,29,30]
    starting_row2 = 4

    # mencari column dan row yang kosong
    empty_row2 = None
    empty_col2 = None

    for row in range(starting_row2, len(all_rows) + 1):
        row_values = all_rows[row-1]
        for col_index in column_range2:
            if col_index < len(row_values) and row_values[col_index] == '':
                empty_row2 = row
                empty_col2 = col_index
                break
        if empty_row2 is not None:
            break

    # jika tidak ada cell yang kosong dari range yang tertentu maka dibuat row baru
    if empty_row2 is None:
        empty_row2 = len(all_rows) + 1
        empty_col2 = 0
        # masukan/insert row baru
        sheet2.insert_row([''] * len(column_range2), empty_row2)
        # update column untuk index row baru
        starting_row2 += 1

    # dapatkan data dari sheet 1, row 16, and column 3
    data3 = sheet.cell(16, 3).value

    # masukan data kedalam cell yang kosong
    while True:
        try:
            sheet2.update_cell(empty_row2, empty_col2 + 1, data3)
            break
        except gspread.exceptions.APIError as e:
            if e.response.status_code == 429:  
                time.sleep(1)  
            else:
                raise e

    context.bot.send_message(chat_id='-xxx', text='Data telah direkam.')

def main():
    # Inisialisasi bot
    updater = Updater(token='xxx', use_context=True)
    dispatcher = updater.dispatcher

    # Daftarkan handler perintah /data
    dispatcher.add_handler(CommandHandler('data', show_data))
    dispatcher.add_handler(CommandHandler('data2', show_data2))
    dispatcher.add_handler(CommandHandler('data3', show_data3))
    dispatcher.add_handler(CommandHandler('data4', show_data_report))
    dispatcher.add_handler(CommandHandler('start', start))
    dispatcher.add_handler(CommandHandler('record', record_data))

    # Jalankan bot
    updater.start_polling()

    # Jadwalkan fungsi record_data untuk run pada waktu/jam tertentu
    schedule.every().hour.at(":00").do(show_data, context=updater)
    schedule.every().hour.at(":00").do(show_data2, context=updater)
    schedule.every().hour.at(":00").do(show_data3, context=updater)
    schedule.every().hour.at(":00").do(record_data, context=updater)
    schedule.every().hour.at(":00").do(show_data_report, context=updater)

    while True:
        current_time = datetime.datetime.now().time()
        if current_time >= datetime.time(8) and current_time <= datetime.time(23):
            schedule.run_pending()
        time.sleep(1)

if __name__ == '__main__':
    main()
