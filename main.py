#!/usr/bin/env pipenv-shebang
import os
import logging.config

import yaml
import pretty_errors


def setup_logs():
    dirpath = os.path.join(os.path.dirname(__file__), 'logs')
    if not os.path.isdir(dirpath):
        os.mkdir(dirpath)

    logs = ['info.log', 'error.log', 'cron.log']
    for log in logs:
        filepath = os.path.join(dirpath, log)
        if not os.path.isfile(filepath):
            with open(filepath, 'w') as f:
                pass


def logger():
    with open('logging.yaml', 'r') as file:
        config = yaml.load(file, Loader=yaml.FullLoader)
        logging.config.dictConfig(config)
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)
    return logger


if __name__ != '__main__':
    setup_logs()
    logger = logger()
