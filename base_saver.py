# -=- coding: utf-8 -=-

from module_for_base_saver import *
from configparser import ConfigParser
from logging.handlers import RotatingFileHandler
import logging
import schedule

logger = logging.getLogger('my_logger')
logging.basicConfig(handlers=[RotatingFileHandler('export_data.log', maxBytes=2000000, backupCount=10)],
            format=u'[%(asctime)s] %(levelname)s %(message)s')

try:
    config = ConfigParser()
    config.read('export_data.ini')
except Exception as e:
    logger.critical(str(e))
    logger.critical('Ошибка чтения export_data.ini')

if loglevel == '1':
    logger.setLevel(logging.DEBUG)
elif loglevel == '0':
    logger.setLevel(logging.CRITICAL)

check_date()

try:
    if to_work == 00:
        schedule.every(1).day.at('00:01').do(job)
    else:
        schedule.every(int(to_work)).minutes.do(job)

except:
    logger.critical('Не выбран период проверки обновления файла!')

while True:
    schedule.run_pending()
    sleep(1)
