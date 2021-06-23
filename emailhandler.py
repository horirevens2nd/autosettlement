#!/usr/bin/env pipenv-shebang
import os
import ssl
import logging
from email import encoders, message_from_string
from email.message import EmailMessage
from email.mime.multipart import MIMEMultipart, MIMEBase
from email.mime.text import MIMEText
from imaplib import IMAP4_SSL
from smtplib import SMTP_SSL

import yaml
import pretty_errors

import dateformat
from filehandler import create_dirs, create_subdirs
from telegrambot import send_message


logger = logging.getLogger("main")

# read secret
with open("secret.yaml", "r") as file:
    secret = yaml.load(file, Loader=yaml.FullLoader)
    USEREMAIL = secret["email"]["username"]
    PASSEMAIL = secret["email"]["password"]


def get_attachment(sender="mohamad.arif.supriyanto", subject=""):
    """get attachment from email

    :param sender: user email, defaults to 'mohamad.arif.supriyanto'
    :type sender: str, optional
    :param subject: email subject, defaults to ''
    :type subject: str, optional
    :return: filename, filepath
    :rtype: str
    """
    with IMAP4_SSL(host="mail.posindonesia.co.id") as server:
        server.login(USEREMAIL, PASSEMAIL)

        # search with specified criteria in inbox folder
        server.select(readonly=True)
        criteria = f'SUBJECT "{subject}" FROM "{sender}"'
        result, data = server.search(None, criteria)
        ids = data[0].split()
        latest_id = ids[-1:]

        if not latest_id:
            filename = None
            filepath = None
        else:
            # fetch newest email
            result, data = server.fetch(latest_id[0], "(RFC822)")
            raw_email = data[0][1]
            raw_email_string = raw_email.decode("utf-8")
            message = message_from_string(raw_email_string)

            # download attachment
            for part in message.walk():
                if part.get_content_maintype() == "multipart":
                    continue
                if part.get("Content-Disposition") is None:
                    continue
                filename = part.get_filename()

                # create a new file if not exist or overwrite it if exist
                if bool(filename):
                    assets_dirpath = create_dirs(["assets"])
                    base_dirpath = create_subdirs(assets_dirpath[0], ["txt"])
                    sub_dirpath = create_subdirs(base_dirpath[0], ["pbb", "pdam"])

                    product_id = filename[9:14]
                    if product_id == "P014B":
                        filepath = os.path.join(sub_dirpath[0], filename)
                    elif product_id == "D039P":
                        filepath = os.path.join(sub_dirpath[1], filename)

                    with open(filepath, "wb") as f:
                        f.write(part.get_payload(decode=True))
                    message = f"Attachment {filename} has been downloaded"
                    logger.info(message)

        return filename, filepath


def send_email(sender, receiver, subject, body, attachment):
    """send an email

    :param sender: user email
    :type sender: str
    :param receiver: user email
    :type receiver: str
    :param subject: email subject
    :type subject: str
    :param body: email message
    :type body: str
    :param attachment: email attachment
    :type attachment: str
    """
    try:
        # create a multipart message and set headers
        email = MIMEMultipart()
        email["From"] = sender
        email["To"] = receiver
        email["Subject"] = subject

        # add body to email
        email.attach(MIMEText(body, "plain"))

        # get attachment
        with open(attachment, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())

        # encode fie in ASCII
        encoders.encode_base64(part)

        # add header as key/value pair to attachment part
        filename = os.path.basename(attachment)
        part.add_header("Content-Disposition", f"attachment; filename={filename}")

        # add attachment to message and convert message to string
        email.attach(part)
        text = email.as_string()

        # login in to server
        context = ssl.create_default_context()
        with SMTP_SSL(host="mail.posindonesia.co.id", context=context) as server:
            server.login(USEREMAIL, PASSEMAIL)
            server.sendmail(sender, receiver, text)
            message = f"Sending file {filename} to {receiver} is succeed"
            logger.info(message)
            send_message(text=message)
    except Exception as e:
        message = f"Sending file {filename} to {receiver} is failed"
        logger.info(message)
        send_message(text=message)
