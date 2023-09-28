import time

import schedule as schedule
import telebot
import gspread
from telebot.types import Message
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import dateparser


months = {
    'январь': 'January',
    'февраль': 'February',
    'март': 'March',
    'апрель': 'April',
    'май': 'May',
    'июнь': 'June',
    'июль': 'July',
    'август': 'August',
    'сентябрь': 'September',
    'октябрь': 'October',
    'ноябрь': 'November',
    'декабрь': 'December'
}

max_length = 4096

# Установите токен своего Telegram бота
bot = telebot.TeleBot('6441360109:AAHlY8DUoWMRbWCRMLPruyE6quvSCwpC1Do')

# Установите аккаунт сервисного аккаунта для Google Sheets
# Следуйте инструкциям по ссылке https://gspread.readthedocs.io/en/latest/oauth2.html#service-account
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(credentials)

# Установите ID таблицы Google Sheets
spreadsheet_id = '1u1fF_ORdVLxzvYHXqSyAXoHrUhu83_kSK6WpP1zjAeM'

text_array = []

expire_array = []
delivery_array = []
refresh_array = []

losted = []

def handle_text():
    # bot_chat_id = bot.get_me().id
    bot_chat_id = '-859736295'
    check_contract_status(bot_chat_id)

def send_message(chat_id, text):
    global sent_messages
    if sent_messages % 20 == 0 and sent_messages != 0:
        print("Limit of 20 messages per minute reached. Waiting...")
        time.sleep(60)
        sent_messages = 0
    # проверяем длину текста
    if len(text) > 4096:
        # если текст слишком длинный, разбиваем его на части
        parts = [text[i:i+4096] for i in range(0, len(text), 4096)]
        for part in parts:
            send_message(chat_id, part)

        return
    # отправляем сообщение
    bot.send_message(chat_id, text, parse_mode='Markdown')
    sent_messages += 1

# функция отправки массива сообщений
def send_messages_array(chat_id, messages, remaining_messages=0):

    # объединяем сообщения в один текст
    text = '\n\n'.join(messages)

    losted_text = '\n\n'.join(losted)
    # проверяем лимит на количество сообщений в минуту

    # отправляем сообщение
    send_message(chat_id, text)
    send_message(chat_id, losted_text)
    # увеличиваем счетчик отправленных сообщений


sent_messages = 0

def check_contract_status(bot_chat_id):
    sheet = client.open_by_key(spreadsheet_id).worksheet('ТН.')
    contracts = sheet.get_all_records(expected_headers=["Статус заказа", "Срок по договору", "номер ТН"])

    today = datetime.now().date()

    for contract in contracts:
        status = contract['Статус заказа']
        status = status.lower()
        try:
            expiry_date = datetime.strptime(contract['Срок по договору'], '%d.%m.%Y').date()
        except:
            continue
        try:
            if status != '':
                all_mas = losted + expire_array + refresh_array + delivery_array
                if not any(contract['номер ТН'] in sub for sub in all_mas):
                    if (status in ['закрыта', 'отказ клиента', 'устарел', '']) or (expiry_date == ''):
                        continue
                    elif status == 'потеряшка':
                        losted.append('Договор *{}* - потерялся'.format(contract['номер ТН']))

                    elif today > expiry_date:
                        expire_array.append('Срок договора *{}* истек. \n !Немедленно обновите договор! '.format(contract['номер ТН']))
                        # bot.send_message(chat_id=bot_chat_id, text='Срок договора {} истек. \n !Немедленно обновите договор! ({}) ({})'.format(contract['номер ТН'], expiry_date.strftime('%d.%m.%Y'), status))
                    elif (expiry_date - today).days <= 14:
                        refresh_array.append('До окончания срока данного договора *{}* осталось менее 14 дней. Просьба – обновите договор. '.format(contract['номер ТН']))
                        # bot.send_message(chat_id=bot_chat_id, text='До окончания срока данного договора {} осталось менее 14 дней. Просьба – обновите договор. ({}) ({})'.format(contract['номер ТН'], expiry_date.strftime('%d.%m.%Y'), status))
                    elif status.startswith('доставка на'):
                        delivery_month = status.split(' ')[-1]
                        if delivery_month in months:
                            delivery_date = datetime.strptime('{}.{}'.format(months[delivery_month], today.year), '%B.%Y').date()
                            if expiry_date <= delivery_date:
                                delivery_array.append('Срок по договору меньше запланированной доставки для договора *{}* '.format(contract['номер ТН']))
                                # bot.send_message(chat_id=bot_chat_id, text='Срок по договору меньше запланированной доставки для договора {} ({}) ({})'.format(contract['номер ТН'], expiry_date.strftime('%d.%m.%Y'), status))
                        else:
                            print(delivery_month)
        except Exception as e:
            print(e)
            continue
    text_array = expire_array + delivery_array + refresh_array
    send_messages_array(bot_chat_id, text_array)

schedule.every().day.at('08:00').do(handle_text)

# бесконечный цикл для выполнения расписания
while True:
    schedule.run_pending()
    time.sleep(30)




