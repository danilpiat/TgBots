import time
from threading import Thread

import telebot
import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from queue import Queue

check_list = ['1\. Назначен ответственный за адаптацию сотрудника',
        '2\. Сформирован адаптационный документ',
        '3\. Подготовлено тестовое задание \(не обязательно\)',
        '4\. Согласована и подготовлена техника',
        '5\. Подготовлена учетная запись в Б24 и других необходимых программах',
        '6\. Назначена задача “Добро пожаловать”',
        '7\. Озвучен и подтвержден день выхода на стажировку']

my_queue = Queue()

for item in check_list:
    my_queue.put(item)

third_day_check_list = ['8\. Предоставлены документы для трудоустройства, копии \(Паспорт, ИНН, СНИЛС\)',
        '9\. Подписана NDA',
        '10\. Двусторонняя ОС: кандидат\+АА \(по окончании 3\-го дня адаптации\)',
        '11\. Кандидат добавлен в орг\. структуру компании']

third_day_my_queue = Queue()

for third_day_item in third_day_check_list:
    third_day_my_queue.put(third_day_item)

# Авторизация в Google Sheets API
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(credentials)

# Установите ID таблицы Google Sheets
spreadsheet_id = '1GPsEjY7p9vo5_fEH-vuZposjf41zg1z9EEIjXlG7YHw'
# Получение данных из таблицы
sheet = client.open_by_key(spreadsheet_id).worksheet('Лист1')
data = sheet.get_all_values()

# Инициализация бота
bot = telebot.TeleBot('5764050861:AAHgihUhnffhSQAr2UQAQtixaz5Wi7VBulk')

chat_ids = []

# Функция для отправки сообщения в личные сообщения
# функция для отправки сообщения о кандидате
def send_candidate_message(candidate):

    date = datetime.datetime.strptime(candidate[4], '%d.%m.%Y').strftime('%d.%m.%Y')
    name = candidate[0]
    department = candidate[1]
    position = candidate[2]
    manager = candidate[3]
    message = f'Завтра {date} на стажировочные дни выходит кандидат {name} в отдел {department} на должность {position}. Руководитель - {manager}.'
    markup = telebot.types.InlineKeyboardMarkup()
    text = date + '|' + name
    accept_button = telebot.types.InlineKeyboardButton(text='Принято', callback_data=text)
    markup.add(accept_button)
    bot.send_message('692712592', message, reply_markup=markup)
    sheet.update_cell(sheet.find(in_column=1, query=name).row, 9, 'Отправлено')


def get_three_days_offset(date):
    today = datetime.datetime.strptime(date, '%d.%m.%Y')
    days_to_count = 2  # количество дней для отсчета
    counted_days = 0  # количество просчитанных дней

    while counted_days < days_to_count:
        today += datetime.timedelta(days=1)
        if today.weekday() < 5:  # проверяем, что это рабочий день (0 - понедельник, 4 - пятница)
            counted_days += 1
    return today.strftime('%d.%m.%Y')


# обработчик кнопки "Принять"
@bot.callback_query_handler(func=lambda call: True)
def accept_candidate(call):
    date_name = call.data
    if '|' in date_name:
        date, name = date_name.split('|')
        three_days_offset = get_three_days_offset(date)
        sheet.update_cell(sheet.find(in_column=1, query=name).row, 7, three_days_offset)
        bot.delete_message('692712592', call.message.message_id)
        send_checklist(name)

    else:
        bot.delete_message('692712592', call.message.message_id)
        send_checklist(date_name)


def split_queue(status):
    index = my_queue.queue.index(status)
    before = list(my_queue.queue)[:index]
    current = my_queue.queue[index]
    after = list(my_queue.queue)[index + 1:]
    return before, current, after

def split_3_days_queue(status):
    index = third_day_my_queue.queue.index(status)
    before = list(third_day_my_queue.queue)[:index]
    current = third_day_my_queue.queue[index]
    after = list(third_day_my_queue.queue)[index + 1:]
    return before, current, after

def get_next(status):
    index = my_queue.queue.index(status)
    if index < len(my_queue.queue) - 1:
        return my_queue.queue[index+1]
    else:
        return None
def get3_next(status):
    index = third_day_my_queue.queue.index(status)
    if index < len(third_day_my_queue.queue) - 1:
        return third_day_my_queue.queue[index+1]
    else:
        return None

def send_3_days_checklist(name):
    curr_status = sheet.cell(sheet.find(in_column=1, query=name).row, 8).value
    if curr_status is None:
        curr_status = third_day_my_queue.queue[0]
        before, current, after = split_3_days_queue(curr_status)
    else:
        next_stat = get3_next(curr_status)
        if next_stat is None:
            return
        else:
            before, current, after = split_3_days_queue(next_stat)
    bef_text = ''
    if before:
        if len(before) == 1:
            bef_text = f'~{before[0]}~'
        else:
            before = [f'~{b}~' for b in before]
            bef_text = '\n'.join(before)
    cur_text = f'*{current}*'
    after_text = ''
    if after:
        if len(after) == 1:
            after_text = after[0]
        else:
            after_text = '\n'.join(after)

    curr_stage = 'Текущий этап: \n' + f'*{current}*'
    text = f'Кандидат {name}\n' + bef_text + '\n' + cur_text + '\n' + after_text + '\n\n' + curr_stage
    markup = telebot.types.InlineKeyboardMarkup()
    nm = name + '^' + '(third_day)'
    accept_button = telebot.types.InlineKeyboardButton(text='Сделано', callback_data=nm)
    markup.add(accept_button)
    bot.send_message('692712592', text, reply_markup=markup, parse_mode='MarkdownV2')
    sheet.update_cell(sheet.find(in_column=1, query=name).row, 8, current)

def send_checklist(name):
    curr_status = sheet.cell(sheet.find(in_column=1, query=name).row, 6).value
    if curr_status is None:
        curr_status = my_queue.queue[0]
        before, current, after = split_queue(curr_status)
    else:
        next_stat = get_next(curr_status)
        if next_stat is None:
            return
        else:
            before, current, after = split_queue(next_stat)
    bef_text = ''
    if before:
        if len(before) == 1:
            bef_text = f'~{before[0]}~'
        else:
            before = [f'~{b}~' for b in before]
            bef_text = '\n'.join(before)
    cur_text = f'*{current}*'
    after_text = ''
    if after:
        if len(after) == 1:
            after_text = after[0]
        else:
            after_text = '\n'.join(after)

    curr_stage = 'Текущий этап: \n' + f'*{current}*'
    text = f'Кандидат {name}\n' + bef_text + '\n' + cur_text + '\n' + after_text + '\n\n' + curr_stage
    markup = telebot.types.InlineKeyboardMarkup()
    accept_button = telebot.types.InlineKeyboardButton(text='Сделано', callback_data=name)
    markup.add(accept_button)
    bot.send_message('692712592', text, reply_markup=markup, parse_mode='MarkdownV2')
    sheet.update_cell(sheet.find(in_column=1, query=name).row, 6, current)

def check_3_days():
    while True:
        candidates = sheet.get_all_values()[1:]

        today = datetime.date.today()
        for candidate in candidates:
            try:
                candidate_stage_date = datetime.datetime.strptime(candidate[6], '%d.%m.%Y').date()
            except:
                continue
            name = candidate[0]
            if candidate_stage_date == today and candidate[7] == '':
                send_3_days_checklist(name)
        time.sleep(10)


# функция для проверки кандидатов на следующий день
def check_candidates():
    while True:
        today = datetime.date.today()
        tomorrow = today + datetime.timedelta(days=1)
        candidates = sheet.get_all_values()[1:]
        for candidate in candidates:
            candidate_date = datetime.datetime.strptime(candidate[4], '%d.%m.%Y').date()
            if candidate_date == tomorrow and candidate[5] == '' and candidate[8] != 'Отправлено':
                send_candidate_message(candidate)
        time.sleep(30)


# запуск бота и проверка кандидатов каждый день в определенное время
if __name__ == '__main__':
    t1 = Thread(target=check_candidates,
               args=())  # передать переменную message. Обратите внимание на запятую после message!
    t1.start()
    t2 = Thread(target=check_3_days,
                args=())  # передать переменную message. Обратите внимание на запятую после message!
    t2.start()
    bot.polling()