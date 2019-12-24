# -=- coding: utf-8 -=-
from decode_license import msgd
from pyautogui import hotkey, typewrite
from time import sleep
from os import system, remove, path
from pywinauto import Application
from psutil import process_iter
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime
from sys import exit
from configparser import ConfigParser

logger = logging.getLogger('my_logger')
logging.basicConfig(handlers=[RotatingFileHandler('export_data.log', maxBytes=2000000, backupCount=10)],
                    format=u'[%(asctime)s] %(levelname)s %(message)s')

try:
    config = ConfigParser()
    config.read('export_data.ini')
except Exception as e:
    logger.critical(str(e))
    logger.critical('Ошибка чтения export_data.ini')
try:
    path_dbf = config['Settings']['PathDBF']
    path_PDS = config['Settings']['PathPDS']
    timeout_open = config['Settings']['OpenTimeout']
    timeout_save = config['Settings']['SaveTimeout']
    timeout_export = config['Settings']['ExportTimeout']
    password = config['Settings']['PassPDS']
    login = config['Settings']['LoginPDS']
    to_work = int(config['Settings']['Period'])
    to_work_time = (config['Settings']['Time'])
    loglevel = config['Settings']['LogLevel']
except Exception as e:
    logger.critical(str(e))
    logger.critical('Проверьте корректность параметров конфигурации')

if loglevel == '1':
    logger.setLevel(logging.DEBUG)
elif loglevel == '0':
    logger.setLevel(logging.CRITICAL)


def job():
    '''  основной блок программы  '''
    print('Программа запущена в {}'.format(datetime.now().strftime("%H:%M:%S")))
    logger.critical('Запуск программы')
    date_create = modification_date(path_dbf + 'CARDS.csv').strftime("%d.%m.%Y")
    date_now = datetime.now().strftime("%d.%m.%Y")

    if date_now > date_create:
        stop_process()

        del_file = path_dbf + 'CARDS.dbf'
        del_old_file(del_file)

        try:
            export_in_pds(path_PDS, path_dbf, timeout_open, timeout_export, password, login)
        except:
            logger.debug('Файл dbf не найден')
            stop_process()
        try:
            sleep(10)
            dfb_to_csv(timeout_save)
        except:
            logger.debug('Файл dbf не найден')
            stop_process()
        logger.critical('Экспорт файла завершён')
        print('Программа завершена в {}'.format(datetime.now().strftime("%H:%M:%S")))
    elif date_now <= date_create:
        logger.critical('Файл сегодня уже был экспортирован')


def modification_date(file):
    '''  проверка даты создания файла  '''
    try:
        t = path.getmtime(file)
        return datetime.fromtimestamp(t)
    except:
        logger.debug('Не удается найти указанный файл {}'.format(file))
        tmp_time = datetime(2000, 1, 1, 0, 0, 1, 1)
        return tmp_time


def check_date():
    '''  сравниваем дату  '''
    cur_date = datetime.now()
    to_date = datetime.strptime(msgd, '%d-%m-%y')
    if to_date < cur_date:
        logger.critical('Bad license')
        exit()


def stop_process():
    '''  остановка процессов  '''
    for proc in process_iter():
        name = proc.name()
        if name == 'PCARDS.EXE' or name == 'EXCEL.EXE':
            system('taskkill /F /IM {}'.format(name))
            logger.debug('Процесс {} остановлен'.format(name))


def export_in_pds(path_PDS, path_dbf, timeout_open, timeout_export, password, login):
    '''  открываем и воодим пароли  '''
    person_card = path_PDS + 'PCARDS.EXE'
    app = Application(backend="uia").start(person_card)
    logger.debug('Успешный запуск PDS')
    app_name = app.Person_Cards
    app_name.Edit.type_keys(password, with_spaces=True)
    hotkey('tab')
    hotkey('tab')
    hotkey('tab')
    typewrite(login)
    app_name.OKButton.click()
    logger.debug('Введёные учётные данные корректны')
    sleep(int(timeout_open))
    movement_cursors(timeout_open, timeout_export, path_dbf)
    stop_process()


def dfb_to_csv(timeout_save):
    '''  конвертируем файлы '''
    del_file = 'CARDS.csv'
    del_old_file(del_file)
    i = 0
    system('start CARDS.dbf')
    sleep(25)
    hotkey('enter')
    sleep(3)
    hotkey('f12')
    sleep(3)
    hotkey('tab')
    sleep(3)
    while True:
        hotkey('down')
        i += 1
        if i == 20:
            break
    sleep(3)
    hotkey('enter')
    sleep(3)
    hotkey('enter')
    sleep(3)
    hotkey('enter')
    logger.debug('Файл CARDS.csv сохранен!')
    sleep(int(timeout_save))
    hotkey('alt', 'f4')
    sleep(3)
    hotkey('right')
    sleep(3)
    hotkey('enter')
    logger.critical('Завершение работы программы')


def movement_cursors(timeout_open, timeout_export, path_dbf):
    '''  перемещаемся внутри программы  '''
    _n = 0
    sleep(5)
    hotkey('alt')
    hotkey('enter')
    hotkey('enter')
    sleep(int(timeout_open))
    hotkey('tab')
    sleep(2)
    hotkey('apps')
    sleep(2)
    while _n != 9:
        hotkey('down')
        _n += 1
    sleep(2)
    hotkey('Enter')
    sleep(10)
    typewrite(path_dbf + 'CARDS.dbf')
    hotkey('Enter')
    sleep(int(timeout_export))
    hotkey('Enter')
    logger.debug('Файл {} импортирован!'.format('CARDS.dbf'))


def del_old_file(file):
    '''  удаление старых файлов  '''
    try:
        if path.exists(file):
            remove(file)
            logger.debug('Файл {} успешно удалён!'.format(file))
    except:
        logger.debug('Файл {} найден!'.format(file))
