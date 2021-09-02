import subprocess
import time
from FileHundler.MakeHFTask import HotFolderCloseOpen, MakeHotFolderTask
import pandas as pd
import glob
import os
import shutil
import win32com.client as wc
import FileHundler.Extracter as Extracter
from datetime import datetime
from multiprocessing import Process, cpu_count
import pymorphy2


class PathCollection:
    path_to_unreadable_files = r'files\\'  # Путь до папки с файлами ДЛЯ РАСПОЗНАВАНИЯ
    output_folder = r'output\\'
    path_to_recognized = r'OCR\\' # Папка С РАСПОЗНАННЫМИ файлами
    res_excel = r'output\\Res_test.xlsx' # Путь то результрующего Excel файла
    zdp_sort = r"C:\Users\Oshchepkov-VA\Desktop\Обработка"  # Корневая папка с файлами для парсинга
    #zdp_sort = r"C:\Users\Oshchepkov-VA\Desktop\Ппп\test\индексация"  # Корневая папка с файлами для парсинга
    temp_folder = r"temp\\"


## Создаем датафрейм на основе словаря. Словарь типа: {ПУТЬ:ТЕКСТ}
def createDf(dict, name):
    df = pd.DataFrame()
    path = []
    text = []
    tokens = []
    normal_text = []
    for key, data in dict.items():
        path.append(key)
        text.append(data["text"])
        tokens.append(data["tokens_lem"])
        normal_text.append(data["text_lem"])
    df["path"] = path
    df["text"] = text
    df["tokens"] = tokens
    df["text_lem"] = normal_text
    create_excel(df, name)  ## Вызов метода для создания результирующего Экселя


## Получаем расширение файла
def getExpan(file):
    return file.split('.')[-1]


## Конввертируем файл типа doc в docx
def convertDocToDocx(file):
    word = wc.Dispatch('Word.Application')
    doc = word.Documents.Open(file)
    name = file.split("\\")[-1]  # Получаем имя файла, типа Pumpurum.doc
    new_name = name.replace("doc", "docx").replace("DOC", "docx")  # Получаем новое имя файла, путем замены расширения файла (почему бы не заменить просто расширение в пути?
    new_path = file.replace(name, new_name)                        # Потому что в папках путивстречаются расширения файлов
    doc.SaveAs(new_path, 16)  # Сохраняем файл с новым именем и укаываем код типа файла (16 = docx)
    text = Extracter.readTextWord(new_path)  ## Читаем текст из файла с типом docx
    doc.Close()
    word.Application.Quit()
    return [text, new_path]  ## Возвращаем прочитанный текст и путь до нового файла

## Функция перемещения (копирования) файлов из текущей директорию, в директорию нераспознанных файлов
def moveUnreadableFile(file):
    print("Файл не читается\n")
    new_path = PathCollection.path_to_unreadable_files + file.split('\\')[-1]
    shutil.copyfile(file, new_path)


# Метод запуска задачи в HotFolder
def HF_recognition(files, name):
    print(name+".xlsx")
    morph_analyzer = pymorphy2.MorphAnalyzer()
    dict_with_text_and_path_hf = {}
    if len(os.listdir(PathCollection.path_to_unreadable_files)) > 0:  # Если нераспознанных файлов больше 0

        MakeHotFolderTask.check()  # Проверяем наличие необходимых директорий
        MakeHotFolderTask.save_new_task()  # Создаем задачу для HotFolder-a
        HotFolderCloseOpen.closeHF()  # Закрываем HotFolder
        HotFolderCloseOpen.openHF()  # Открываем HotFolder

        print('''
                   ================== ABBYY Hot Folder =====================
                   Ожидайте завершение ABBYY Hot Folder и сообщения о завершении.
                   Статус обработки отображается в ABBYY Hot Folder.
                   ================== ABBYY Hot Folder =====================
                   ''')

        while not 'Hot Folder Log.txt' in os.listdir(PathCollection.path_to_recognized):  # Если не появился файл логов Hot Folder-a значит он еще работает и надо подождать
            time.sleep(1)

    list_of_recognized_files = [y for x in os.walk(PathCollection.path_to_recognized) for y in glob.glob(
        os.path.join(x[0], '*.docx'))]  # Получаем список распознаных файлов (В папке с распознанными файлами)

    for rec_file in list_of_recognized_files:  # Читаем все из распознанных файлов
        print("docx")
        if Extracter.readTextWord(rec_file) is False:
            print("Файл не читается\n")
            continue
        else:
            text_from_file = Extracter.readTextWord(rec_file)

        for fl in files:  # Получаем путь до НЕ распознаного файла из родной директории, чтобы в итоговом доке указать путь до изначального докуммента
            name = rec_file.split("\\")[-1]  # Получаем имя распознанного файла
            exp = '.' + name.split('.')[-1]  # Получаем расширение из файла
            name = name.replace(exp, "")   # Получаем имя без расширения
            if name in fl:  # Ищем по имени оригинальный файл (до расаознавания)
                #dict_with_text_and_path_hf[fl] = text_from_file  # В словарь вместо пути до распозанного файла, указываем путь до оригинального файла
                tokens_text, normal_text = tokenizer_n_normalize(morph_analyzer, text_from_file)
                dict_with_text_and_path_hf[fl] = {}
                dict_with_text_and_path_hf[fl]["text"] = text_from_file
                dict_with_text_and_path_hf[fl]["tokens_lem"] = tokens_text
                dict_with_text_and_path_hf[fl]["text_lem"] = normal_text
                break
        os.remove(rec_file)  # Удаялем распознанные файлы

    HotFolderCloseOpen.closeHF()  # Закрывам HotFolder
    if (os.path.isdir(PathCollection.path_to_recognized + '\\Hot Folder Log.txt')):
       os.remove(PathCollection.path_to_recognized + '\\Hot Folder Log.txt')  # И файл логов тоже удаляем
    createDf(dict_with_text_and_path_hf, name)


def tokenizer_n_normalize(morph_analyzer, text):
    lines = []
    split_text = text.replace(",", ".").split(" ")
    for word in split_text:
        index_by_dot = word.find(".")
        count_dot = word.count(".")
        if index_by_dot != -1:
            if index_by_dot == 0 or index_by_dot == len(word)-1:
                word = word.replace(".", " ")
            elif count_dot > 1:
                word = word[::-1].replace(".", "", count_dot - 1)[::-1]
            elif word[index_by_dot + 1].isdigit() and word[index_by_dot - 1].isdigit():
                lines.append(word.rstrip(".").rstrip(" "))
                continue
            else:
                word = word.replace(".", " ")
        if word != '':
            lines.append(word.rstrip(".").rstrip(" "))

    tokens = [morph_analyzer.parse(word)[0].normal_form for word in lines]
    normal_text = " ".join(tokens)
    return tokens, normal_text


def main(files, name):
    print(name+".xlsx")
    morph_analyzer = pymorphy2.MorphAnalyzer()
    dict_with_text_and_path = {}  # Словарь текста из файлов и пути до этого файла
    for file in files:

        check_about_chek = file.lower().split("\\")[-1]
        if 'чек' in check_about_chek:  # Пропускаем все файлы, в названии которых, присутсвтует слово чек
            continue

        text_from_file = ''
        if ("img" in file.lower() or "png" in file.lower() or "jpg" in file.lower() or "jpeg" in file.lower() and ".txt" not in file.lower()):
            print("Файл не читается\n")
            moveUnreadableFile(file)
            continue
        if (getExpan(file) == "pdf" or getExpan(file) == "PDF") and not os.path.isdir(file):  # Если расширение файла pdf и файл НЕ папка
            parsing_file = Extracter.readTextPdf(file)
            if parsing_file is False:  ## Если в тексте меньше 100 символов создаем копию этого файла в заранее созданной папке, для дальнейшего распознавания Hot Folder-ом
                print("Файл не читается\n")
                moveUnreadableFile(file)
                continue
            else:
                text_from_file = parsing_file  # Вызов функции парсинга документа
                tokens_text, normal_text = tokenizer_n_normalize(morph_analyzer, text_from_file)


        elif (getExpan(file) == "doc" or getExpan(
                file) == "DOC") and "~$" not in file and not os.path.isdir(file):  ## Если файл типа doc, конвертируем его в Docx (сразу же читаем) и удаляем старую версию с расширением doc
            try:
                text_and_path = convertDocToDocx(file)
            except Exception as e:
                text_and_path = convertDocToDocx(file)

            os.remove(file)
            if (text_and_path[0] is False):
                print("Файл не читается\n")
                continue
            else:
                text_from_file = text_and_path[0]
                file = text_and_path[1]
                tokens_text, normal_text = tokenizer_n_normalize(morph_analyzer, text_from_file)

        elif (getExpan(file) == "docx" or getExpan(file) == "DOCX") and "~$" not in file and not os.path.isdir(file):
            parsing_file = Extracter.readTextWord(file)
            if parsing_file is False:
                print("Файл не читается\n")
                continue
            else:
                text_from_file = parsing_file
                tokens_text, normal_text = tokenizer_n_normalize(morph_analyzer, text_from_file)

        elif (getExpan(file) == "xlsx" or getExpan(file) == "XLSX" or getExpan(file) == "xls") and not os.path.isdir(file):
            parsing_file = Extracter.readTextExcel(file)
            if parsing_file is False:
                print("Файл не читается\n")
                continue
            else:
                text_from_file = parsing_file
                tokens_text, normal_text = tokenizer_n_normalize(morph_analyzer, text_from_file)

        elif (getExpan(file) == "txt" or getExpan(file) == "TXT") and "~$" not in file and not os.path.isdir(file):
            parsing_file = Extracter.readTextTXT(file)
            if parsing_file is False:
                print("Файл не читается\n")
                continue
            else:
                text_from_file = parsing_file
                tokens_text, normal_text = tokenizer_n_normalize(morph_analyzer, text_from_file)

        else:
            continue

        dict_with_text_and_path[file] = {} # Создадим следующую структуру словаря "путь до файла" : {текст из файла, токены, лем. текст}
        dict_with_text_and_path[file]["text"] = text_from_file
        dict_with_text_and_path[file]["tokens_lem"] = tokens_text
        dict_with_text_and_path[file]["text_lem"] = normal_text

    createDf(dict_with_text_and_path, name)  # Создаем дата фрейм

## Функция создания результрующего Excel-я на основе ранее созданного словаря
def create_excel(df, name):
    writer = pd.ExcelWriter(PathCollection.temp_folder + name + ".xlsx", engine='xlsxwriter')  # Результирующий файл
    df.to_excel(writer, sheet_name='Sheet')
    writer.save()

## Функция разархивации архивов
def __unrar__(pathunrarfiles, dir_unrar):
    '''
        Обозначение команд и переключателей
        r = рекурсивный обход
        y = отвечать на все вопросы YES
        o+ = разрешение перезаписи
        ibck = запускать winrar фоном
        inul = деактивировать все ошибки
    '''

    par = r'"C:\Program Files\WinRAR\WinRAR.exe" ' \
          '-r ' \
          '-y ' \
          '-o+ ' \
          '-ibck ' \
          '-inul '\
          'x '
    command = par + '"' + pathunrarfiles + '" "' + dir_unrar + '"'
    FNULL = open(os.devnull, 'w')
    subprocess.call(command, stdout=FNULL, stderr=FNULL, shell=False)
    FNULL.close()


def createDirForUnrar(file):
    # Получаем новый путь для архивов (этот путь на пару директорий выше, коэф. 5 это порядок директории от корневой)
    # Например \\irkut.ca.sbrf.ru\vol1\BKB_UVA\ОАОП\03. Планы и отчеты\2021\03. Бэклог ОАОП\13725_Залоги\02_DATA\01_docs\7000 и если мы хотим перенести все архивы в папку 7000
    # То надо указать 10
    promised_land_for_arch = (file.split(list(filter(None, file.split("\\")))[5])[0] + list(filter(None, file.split("\\")))[5] + "\\")

    if not os.path.exists(promised_land_for_arch + file.split("\\")[-1]):
        shutil.move(file, promised_land_for_arch)

    new_name = file.split("\\")[-1][-10:] + '_' + str(datetime.now().time()).replace(':', '_')  # Новое имя папки

    if not (os.path.isdir(promised_land_for_arch + new_name)):
        os.mkdir(promised_land_for_arch + new_name)

    print("Новый путь до архива: " + promised_land_for_arch + "\n")
    return [promised_land_for_arch, new_name]


def workWirhArch(files):  # Распаковка архивов
    for file in files:
        if getExpan(file) == 'rar' or getExpan(file) == 'RAR' or getExpan(file) == 'zip' or getExpan(file) == 'ZIP':
            print("Происходит распаковка архива: "+file)
            try:
                created_dirs = createDirForUnrar(file)
                __unrar__(created_dirs[0] + file.split("\\")[-1], created_dirs[0] + created_dirs[1])
            except Exception as e:
                print(e)
                print("С архивом что-то не так")
                continue


def checkFoldeers():
    if not (os.path.isdir(PathCollection.path_to_unreadable_files)):
        os.mkdir(PathCollection.path_to_unreadable_files)
    if not (os.path.isdir(PathCollection.path_to_recognized)):
        os.mkdir(PathCollection.path_to_recognized)
    if not (os.path.isdir(PathCollection.output_folder)):
        os.mkdir(PathCollection.output_folder)
    if not (os.path.isdir(PathCollection.temp_folder)):
        os.mkdir(PathCollection.temp_folder)


def split_list(file_list, n):
    return [file_list[i:i + n] for i in range(0, len(file_list), n)]


if __name__ == '__main__':
    print("------------Обработка началась------------\n")

    checkFoldeers()
    start_time = datetime.now()
    list_of_files = [y for x in os.walk(PathCollection.zdp_sort) for y in glob.glob(os.path.join(x[0], '*.*'))]  # Получаем ВСЕ файлы
    workWirhArch(list_of_files)
    list_of_files = [y for x in os.walk(PathCollection.zdp_sort) for y in glob.glob(
        os.path.join(x[0], '*.*'))]  # Еще раз получаем все файлы (потому что добавились еще те, что из архивов)

    group_files_for_hundler = split_list(list_of_files, len(list_of_files) // cpu_count())
    procs = []

    for i, group in enumerate(group_files_for_hundler):
        proc = Process(target=main, args=(group, str(i)))
        procs.append(proc)
        proc.start()

    for proc in procs:
        proc.join()

    HF_recognition(list_of_files, str(len(group_files_for_hundler)))

    temp_files = glob.glob(PathCollection.temp_folder + "*.xlsx")

    dfs = [pd.read_excel(file) for file in temp_files]

    concat_df = pd.concat(dfs)
    writer = pd.ExcelWriter(PathCollection.res_excel, engine="xlsxwriter")
    concat_df.to_excel(writer)
    writer.save()
    writer.close()

    shutil.rmtree(PathCollection.temp_folder)
    shutil.rmtree(PathCollection.path_to_unreadable_files)
    shutil.rmtree(PathCollection.path_to_recognized)

    print("The End!")
    print(datetime.now() - start_time)
