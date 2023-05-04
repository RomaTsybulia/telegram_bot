import os
import time

import pandas as pd
import telebot
from dotenv import load_dotenv
load_dotenv()

GROUP_SIZE = 1000
USER_BASE = "users.xlsx"

bot = telebot.TeleBot(os.environ["TOKEN"])


def prepare_user_list(user_base_file):
    users_file = pd.read_excel(user_base_file)
    users = users_file['telegram_id'].tolist()

    users_groups = [
        users[i:i + GROUP_SIZE]
        for i, user in enumerate(users)
        if i % GROUP_SIZE == 0
    ]
    return users_groups


def send_message(message):
    groups = prepare_user_list(USER_BASE)
    users_get_message = 0
    for group in groups:
        for user in group:
            try:
                bot.send_message(user, message)
                print(user)
                users_get_message += 1
            except Exception as e:
                print(f'Error: {e}')

    generate_report(users_get_message)


def generate_report(num_active_users):
    report = f'<html><body><h1>Звітність</h1><p>Активні користувачі: {num_active_users}</p></body></html>'
    with open("report.html", "w") as file:
        file.write(report)



@bot.message_handler(commands=['send_message'])
def handle_send_message(message):
    msg_text = message.text.replace('/send_message', '').strip()
    send_message(msg_text)


@bot.message_handler(commands=['report'])
def handle_report(message):
    document = open('report.html', 'rb')
    bot.send_document(message.chat.id, document=document)


bot.polling()
