#!/usr/bin/env pipenv-shebang
import os
import sys
from threading import Thread

import yaml
import pretty_errors
from telegram.ext import Updater, CommandHandler, Filters

# read secret.yaml
with open('secret.yaml', 'r') as file:
    secret = yaml.load(file, Loader=yaml.FullLoader)
    SETTLEMENT = secret['telegram']['tokens']['settlement']
    HORIREVENS = secret['telegram']['tokens']['horirevens']
    YOGI = secret['telegram']['users']['yogi']
    ACHMAD = secret['telegram']['users']['achmad']

# init updater
updater = Updater(token=SETTLEMENT, use_context=True)


def send_message(chat_id=YOGI, text='...', parse_mode=None):
    """send message to client"""
    updater.bot.send_message(
        chat_id=chat_id, text=text, parse_mode=parse_mode)


def start(update, context):
    """post bot command list"""
    message = f'''
    \n*Settlement Command*\
    \n/pbb \- get settlement PBB\
    \n/pdam \- get settlement PDAM\
    \n/databasepbb \- get database PBB (.xlsx)\
    \n/databasepdam \- get database PDAM (.xlsx)\
    \n*Bot Command*\
    \n/getid \- get telegram ID\
    \n/start \- show this message\
    \n/restart \- restart the bot\
    \n/stop \- stop the bot\
    '''
    update.message.reply_text(message, parse_mode='MarkdownV2')


def getid(update, context):
    """get telegram ID"""
    chat_id = update.message.chat.id
    first_name = update.message.chat.first_name
    last_name = update.message.chat.last_name
    last_name = f' {last_name}' if last_name is not None else ''
    message = f'Hello {first_name}{last_name}.. your ID is {chat_id}'
    update.message.reply_text(message)


def stop_and_restart():
    updater.stop()
    os.execl(sys.executable, sys.executable, *sys.argv)


def restart(update, context):
    """restart telegram bot"""
    update.message.reply_text('Bot is restarting...')
    Thread(target=stop_and_restart).start()


def stop_and_shutdown():
    updater.stop()
    updater.is_idle = False


def stop(update, context):
    """stop telegram bot"""
    update.message.reply_text('Bot is shutting down...')
    Thread(target=stop_and_shutdown).start()


def pbb(update, context):
    """run pbb.py script"""
    os.system('pipenv-shebang pbb.py')


def pdam(update, context):
    """run pdam.py script"""
    os.system('pipenv-shebang pdam.py')


def database_pbb(update, context):
    """download file 'Database PBB.xlsx"""
    chat_id = update.message.chat.id
    filename = 'Database PBB.xlsx'
    filepath = os.path.join('assets', filename)
    if os.path.isfile(filepath):
        updater.bot.send_document(
            chat_id=chat_id, document=open(filepath, 'rb'))
    else:
        message = f'{filename} is not exist'
        send_message(chat_id=chat_id, text=message)


def database_pdam(update, context):
    """download file 'Database PDAM.xlsx"""
    chat_id = update.message.chat.id
    filename = 'Database PDAM.xlsx'
    filepath = os.path.join('assets', filename)
    if os.path.isfile(filepath):
        updater.bot.send_document(
            chat_id=chat_id, document=open(filepath, 'rb'))
    else:
        message = f'{filename} is not exist'
        send_message(chat_id=chat_id, text=message)


def main():
    # filter by username
    yogitrismayana = Filters.user(username='@yogitrismayana')

    # init handlers
    pbb_handler = CommandHandler('pbb', pbb)
    pdam_handler = CommandHandler('pdam', pdam)
    dbpbb_handler = CommandHandler('databasepbb', database_pbb)
    dbpdam_handler = CommandHandler('databasepdam', database_pdam)
    getid_handler = CommandHandler('getid', getid)
    start_handler = CommandHandler('start', start)
    restart_handler = CommandHandler(
        'restart', restart, filters=yogitrismayana)
    stop_handler = CommandHandler(
        'stop', stop, filters=yogitrismayana)

    # set handler to dispacher
    dispatcher = updater.dispatcher
    dispatcher.add_handler(pbb_handler)
    dispatcher.add_handler(pdam_handler)
    dispatcher.add_handler(dbpbb_handler)
    dispatcher.add_handler(dbpdam_handler)
    dispatcher.add_handler(getid_handler)
    dispatcher.add_handler(start_handler)
    dispatcher.add_handler(restart_handler)
    dispatcher.add_handler(stop_handler)

    # start the bot
    updater.start_polling()
    updater.idle()


if __name__ == '__main__':
    main()
