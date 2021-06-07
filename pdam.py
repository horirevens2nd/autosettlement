#!/usr/bin/env pipenv-shebang
import os
import logging
import imaplib
import datetime

import pretty_errors

import main
import dateformat
from emailhandler import get_attachment
from filehandler import get_content, xlsx_template_3
from telegrambot import send_message, YOGI, ACHMAD


def main():
    logger = logging.getLogger('main')

    try:
        today = datetime.date.today()
        yesterday = today - datetime.timedelta(days=1)
        yesterday1 = yesterday.strftime('%d-%m-%Y')
        yesterday2 = yesterday.strftime('%Y%m%d')
        yesterday2 = dateformat.format_1(yesterday2)

        filename, filepath = get_attachment(
            subject=f'ENGINE PDAM KAB. NGANJUK [D039P] Rekon Tanggal {yesterday1}')
        if filename is None and filepath is None:
            message1 = 'Email \"SETTLEMENT PDAM KAB NGANJUK\" has not been received'
            logger.info(message1)
            message2 = f'Email *SETTLEMENT PDAM KAB NGANJUK* untuk transaksi tanggal {yesterday2} belum diterima'
            send_message(text=message2, parse_mode='MarkdownV2')
            send_message(chat_id=ACHMAD, text=message2,
                         parse_mode='MarkdownV2')
        else:
            filesize = os.path.getsize(filepath)
            if filesize > 5:
                contents, date, count, total = get_content(filepath)

                message = f'Data available in attachment {filename}, file is saved'
                logger.info(message)

                # export as xlsx file
                filename, filepath = xlsx_template_3(contents)
                if filepath is not None:
                    message = f'''
                    \n*SETTLEMENT PDAM KAB NGANJUK*\
                    \nTanggal Trx : {date}\
                    \nJumlah Trx : {count}\
                    \nTotal BSU : Rp {total}\
                    '''
                    send_message(text=message,
                                 parse_mode='MarkdownV2')
                    send_message(chat_id=ACHMAD, text=message,
                                 parse_mode='MarkdownV2')
            else:
                if os.path.exists(filepath):
                    os.remove(filepath)
                    message = f'No data available in attachment {filename}, file is removed'
                    logger.info(message)

                message = 'Tidak ada *SETTLEMENT PDAM KAB NGANJUK* untuk transaksi hari kemarin'
                send_message(text=message, parse_mode='MarkdownV2')
                send_message(chat_id=ACHMAD, text=message,
                             parse_mode='MarkdownV2')
    except(imaplib.IMAP4.error) as e:
        logger.exception(e)
        message = 'Unexpected error is occurs'
        send_message(text=message)


if __name__ == '__main__':
    main()
