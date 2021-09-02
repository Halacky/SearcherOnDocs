from os import getenv
from os import getcwd
from os import path
from os import makedirs
from os import system
from os import popen
from pytz import UTC
from lxml import etree
from datetime import datetime as dt
from datetime import timedelta


class ModulePath:
    PATH_OCR = 'OCR'  # Папка с распознанными файлами
    PATH_FILES = 'files'  # Папка с файлами для распознания
    PATH_USER = getenv('USERPROFILE')  # Считываем путь к папке пользователя
    PATH_HF = '\\AppData\\Local\\ABBYY\\FineReader\\14.00\\HotFolder'  # Путь для сохранения файла задач
    PATH_HF_TEMP = '\\AppData\\Local\\Temp\\ABBYY\\FineReader\\14.00\\HotFolder'  # Путь к временной папке задач
    PATH_FROM = getcwd()  # Текущее расположение файлов


class MakeHotFolderTask:
    PATH_HF_task = ModulePath.PATH_FROM + '\\' + 'HotFolder_Task\TASKForPars.hft'
    HF_task = PATH_HF_task.rsplit("\\", 1)[1]  # Шаблон задачи для HotFolder

    @classmethod
    def read_task(self):
        with open(self.PATH_HF_task, encoding='utf-16') as task:
            text = task.read()
            tree = etree.fromstring(text)
        return tree

    @classmethod
    def change_attrib(self):
        tree = self.read_task()
        tree.attrib['name'] = 'Digital-аудит (ОАОП)'
        tree.attrib['status'] = 'scheduled'
        tree.attrib['startTime'] = (dt.now(UTC) + timedelta(minutes=1)).strftime('%H:%M:%S.000 %d.%m.%Y UTC')

        # Папка для сохранения файла задач для импорта
        for level_1 in tree.getchildren():
            for level_2 in level_1.getchildren():
                if 'batchToCreateFolder' in level_2.keys():
                    level_2.attrib['batchToCreateFolder'] = ModulePath.PATH_USER + ModulePath.PATH_HF_TEMP

        # Папка для сохранения результатов распознания
        for level_1 in tree.getchildren():
            for level_2 in level_1.getchildren():
                for level_3 in level_2.getchildren():
                    if 'savePath' in level_3.keys():
                        level_3.attrib['savePath'] = ModulePath.PATH_FROM + '\\' + ModulePath.PATH_OCR

        # Папка с файлами для распознания
        for level_1 in tree.getchildren():
            for level_2 in level_1.getchildren():
                if 'folderPath' in level_2.keys():
                    level_2.attrib['folderPath'] = ModulePath.PATH_FROM + '\\' + ModulePath.PATH_FILES
        return tree

    @classmethod
    def save_new_task(self):
        with open(ModulePath.PATH_USER + ModulePath.PATH_HF + '\\' + self.HF_task, 'w', encoding='utf-16') as task:
            for_save = etree.tostring(self.change_attrib(), encoding='utf-16', pretty_print=True)
            task.write(for_save.decode('utf-16'))
        return None

    @classmethod
    def check(self):
        if not path.exists(ModulePath.PATH_USER + ModulePath.PATH_HF):
            makedirs(ModulePath.PATH_USER + ModulePath.PATH_HF)

        if not path.exists(ModulePath.PATH_USER + ModulePath.PATH_HF_TEMP):
            makedirs(ModulePath.PATH_USER + ModulePath.PATH_HF_TEMP)

        if not path.exists(ModulePath.PATH_FROM + '\\' + ModulePath.PATH_OCR):
            makedirs(ModulePath.PATH_FROM + '\\' + ModulePath.PATH_OCR)

        if not path.exists(ModulePath.PATH_FROM + '\\' + ModulePath.PATH_FILES):
            makedirs(ModulePath.PATH_FROM + '\\' + ModulePath.PATH_FILES)
        return None


class HotFolderCloseOpen:
    @classmethod
    def closeHF(self):
        '''Закрываем HotFolder чтобы подгрузить задачу при следующем открытии'''
        system("TASKKILL /F /IM \"HotFolder.exe\"")
        return None

    @classmethod
    def openHF(self):
        '''Запускаем HotFolder'''
        popen(r'C:\Program Files (x86)\ABBYY FineReader 14\HotFolder.exe')
        return None
