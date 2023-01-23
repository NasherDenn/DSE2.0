from docx import Document
import openpyxl
import os
import logging
import traceback
import utilities

# получаем имя машины с которой был осуществлён вход в программу
uname = os.environ.get('USERNAME')
# инициализируем logger
logger = logging.getLogger()
logger_with_user = logging.LoggerAdapter(logger, {'user': uname})


def add_table():
# def add_table(name_dir):
    # временная переменная со списком тестовых файлов UTT и PAUT 2019-2022
    name_dir = (['C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-A1-210-HF-104-UTT-22-01 rev.01.docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-A1-210-VC-101-UTT-22-01.docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-A1-331-HA-104-UTT-22-01.docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-PAUT-20-015.docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-PAUT-20-030.docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-PAUT-20-113 A1-4000-CM-071.docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-PAUT-20-160 Oil Tr3_rev.1 - Copy.docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-PAUT-20-160 Oil Tr3_rev.1',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-PAUT-21-308.docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-PAUT-21-309.docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-PAUT-21-336 (Sulfur Tr2).docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-PAUT-22-106 WELD (Gas1 Tr1).docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-PAUT-22-141.docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-PAUT-22-740_Corrosion_mapping.docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-UT-19-011.docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-UT-19-081',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-UT-19-360 A1-690-FG-051A.docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-UT-19-388 (A1-600-VA-006).docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-UT-20-020 S.docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-UT-20-039.docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-UT-20-143.docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-UT-20-196.docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-UT-20-397.docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-UT-20-432.docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-UTT-21-419 (C3-160-VA-003).docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-UTT-21-605 (Unit 520).docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-UTT-21-770.docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-UTT-21-823(Unit 210 Tr3).docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-UTT-22 -473 (Unit 332 Tr1).docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-UTT-22-064 (Unit 332 Tr1).docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-UTT-22-112(A1-210-HX-201G).docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-UTT-22-305 (CON64).docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-UTT-22-336-CON35.docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-UTT-22-341(KUT 560).docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-UTT-22-618.docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-UTT-22-674  Gas 2 Tr1  330.docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-UTT-22-714.docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-UTT-22-756.docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-ON-UTT-22-811.docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-OS-UT-21-794 (M2-400-VA-310).docx',
                 'C:/Users/asus\Documents/NDT YKR/Тестовые данные/04-YKR-OS-UTT-21-499 (M2-530-TA-002).docx',
                'docx(*.docx)'])
    utilities.name_dir(name_dir)


