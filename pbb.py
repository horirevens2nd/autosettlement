#!/usr/bin/env pipenv-shebang
import os
import logging
import imaplib
import datetime
from time import sleep

import pretty_errors

import main as main_app
import dateformat
from emailhandler import get_attachment, send_email
from filehandler import get_content, xlsx_template_1, xlsx_template_2
from telegrambot import send_message, ACHMAD


def main():
    logger = logging.getLogger("main")

    try:
        today = datetime.date.today()
        yesterday = today - datetime.timedelta(days=1)
        yesterday1 = yesterday.strftime("%d-%m-%Y")
        yesterday2 = yesterday.strftime("%Y%m%d")
        yesterday2 = dateformat.format_1(yesterday2)

        filename, filepath = get_attachment(
            subject=f"ENGINE PBB KAB NGANJUK Tgl {yesterday1}"
        )
        if filename is None and filepath is None:
            message1 = 'Email "SETTLEMENT PBB KAB NGANJUK" has not been received'
            logger.info(message1)
            message2 = f"Email *SETTLEMENT PBB KAB NGANJUK* untuk transaksi tanggal {yesterday2} belum diterima"
            send_message(text=message2, parse_mode="MarkdownV2")
            send_message(chat_id=ACHMAD, text=message2, parse_mode="MarkdownV2")
        else:
            filesize = os.path.getsize(filepath)
            if filesize > 5:
                contents, date, count, total = get_content(filepath)

                message = f"Data available in attachment {filename}, file is saved"
                logger.info(message)

                # export as xlsx file
                filename1, filepath1 = xlsx_template_1(contents)
                sleep(3)
                filename2, filepath2 = xlsx_template_2(contents)

                if filepath1 is not None and filepath2 is not None:
                    date = dateformat.format_1(filename[:8])

                    message = f"""
                    \n*SETTLEMENT PBB KAB NGANJUK*\
                    \nTanggal Trx : {date}\
                    \nJumlah Trx : {count}\
                    \nTotal BSU : Rp {total}\
                    """
                    send_message(text=message, parse_mode="MarkdownV2")
                    send_message(chat_id=ACHMAD, text=message, parse_mode="MarkdownV2")

                    sender = "644spv@posindonesia.co.id"
                    # receiver = 'hori.juventini@gmail.com'
                    receiver = "bapenda@nganjukkab.go.id"
                    subject = f"Trx PBB Kantor Pos Nganjuk {date}"
                    body = """
                    Dengan hormat,

                    Berikut ini kami kirimkan file sesuai dengan subject diatas.
                    Tks

                    ===============================================
                    # Yogi Trismayana 991483728
                    # HP. 082140513878
                    # Kantor Pos Nganjuk 64400
                    # Jl. Supriyadi No. 19 Kauman - Nganjuk 64411
                    ===============================================
                    """
                    send_email(sender, receiver, subject, body, filepath2)
                else:
                    message = f"Unexpected error occurs. {filename1} and {filename2} is not created"
                    logger.info(message)
                    send_message(text=message)
            else:
                if os.path.exists(filepath):
                    os.remove(filepath)
                    message = (
                        f"No data available in attachment {filename}, file is removed"
                    )
                    logger.info(message)

                message = f"Tidak ada *SETTLEMENT PBB KAB NGANJUK* untuk transaksi tanggal {yesterday2}"
                send_message(text=message, parse_mode="MarkdownV2")
                send_message(chat_id=ACHMAD, text=message, parse_mode="MarkdownV2")
    except (imaplib.IMAP4.error) as e:
        logger.exception(e)
        message = "Unexpected error is occurs"
        send_message(text=message)


if __name__ == "__main__":
    main()
