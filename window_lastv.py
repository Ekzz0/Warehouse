import threading

import wx
import os
import time
from scripts import start
from threading import Thread
from pubsub import pub
from elevate import elevate

BUTTON_ADD = wx.NewIdRef()
BUTTON_START = wx.NewIdRef()
BUTTON_OPEN_EXCEL = wx.NewIdRef()
BUTTON_SET_NAME = wx.NewIdRef()
BUTTON_CLEAR = wx.NewIdRef()


# Класс для создания диалогового окна
class MyDialog(wx.Dialog):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.SetSize(400, 300)
        self.parent = parent  # parent - родительское окно. В данном случае это главныц экран приложения
         # Создание окна ввода и консоли:
        vbv = wx.BoxSizer(wx.VERTICAL)
        self.txt = wx.TextCtrl(self, wx.ID_ANY)
        vbv.Add(self.txt, flag=wx.EXPAND | wx.ALL, border=10)
        btn_ok = wx.Button(self, wx.ID_ANY, label='Ввести')
        vbv.Add(btn_ok, flag=wx.EXPAND | wx.ALL, border=10)
        self.control = wx.TextCtrl(self, style=wx.TE_READONLY | wx.TE_MULTILINE)
        self.control.WriteText('Введите имя нового файла')
        vbv.Add(self.control, flag=wx.EXPAND | wx.ALL, border=10, proportion=1)
        self.SetSizer(vbv)
        btn_ok.Bind(wx.EVT_BUTTON, self.on_btn_ok)

    # Функция, которая срабатывает при указании нового имени в диалоговом окне
    def on_btn_ok(self, event):
        self.control.SetValue('')
        if self.txt.GetValue() == '':
            self.control.WriteText(f'{time.asctime()[11:19]} - Нужно ввести имя нового файла!\n')
        else:
            self.parent.tc2.SetValue(self.txt.GetValue() + ".xlsx")
            self.parent.console.WriteText(f"{time.asctime()[11:19]} - Название нового файла установлено: {self.parent.tc2.GetValue()}\n")
            self.EndModal(wx.ID_OK)


# Основной класс для создания приложения
class MyFrame(wx.Frame):
    def __init__(self, parent, title):
        super().__init__(parent, title=title, size=(500, 700))
        """НУЖНО ЗАПРЕТИТЬ ИЗМЕНЕНИЕ РАЗМЕРОВ ОКНА"""
        self.path_name = 0

        # Создание иконки
        ico = wx.Icon('my_ico.ico', wx.BITMAP_TYPE_ICO)
        self.SetIcon(ico)

        # Тут начинается проектирование основного окна
        self.panel = wx.Panel(self)
        self.panel.SetBackgroundColour('#f8a05f')
        vbox = wx.BoxSizer(wx.VERTICAL)

        # Создание 1й сверху части (Добавить файл)
        vbox1, self.btn_add, st1, self.tc1 = self.my_vbox_create("Выбрать файл", "Путь к файлу: ", BUTTON_ADD, 14,
                                                                 self.panel)
        self.btn_add.SetBackgroundColour('#f6f6f6')
        vbox.Add(vbox1, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, border=20)

        # Создание 2й сверху части (Ввести имя файла)
        vbox2, self.btn_set_name, st2, self.tc2 = self.my_vbox_create("Ввести имя файла", "Имя файла: ",
                                                                      BUTTON_SET_NAME, 14,
                                                                      self.panel)
        self.btn_set_name.SetBackgroundColour('#f6f6f6')
        self.tc2.SetValue('Report.xlsx (Чтобы изменить имя, нажми \"Ввести имя файла\")')
        vbox.Add(vbox2, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, border=20)

        # Создание 3й сверху части: кнопки "Начать расчет"
        vbox3 = wx.BoxSizer(wx.VERTICAL)
        self.btn_start = wx.Button(self.panel, BUTTON_START, label='Начать расчет', size=(360, 30))
        self.btn_start.SetBackgroundColour('#f6f6f6')
        vbox3.Add(self.btn_start, flag=wx.EXPAND | wx.BOTTOM, border=20)
        vbox.Add(vbox3, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, border=20)

        # Создание 4й сверху части: кнопки "Открыть отчет"
        vbox4 = wx.BoxSizer(wx.VERTICAL)
        self.btn_open_excel = wx.Button(self.panel, BUTTON_OPEN_EXCEL, label='Открыть отчет', size=(360, 30))
        self.btn_open_excel.SetBackgroundColour('#f6f6f6')
        vbox4.Add(self.btn_open_excel, flag=wx.EXPAND | wx.BOTTOM)
        vbox.Add(vbox4, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, border=20)

        # Создание 5й свехру части: консоли
        st5 = wx.StaticText(self.panel, label="Консоль:")
        vbox.Add(st5, flag=wx.EXPAND | wx.LEFT, border=20)
        self.console = wx.TextCtrl(self.panel, style=wx.TE_READONLY | wx.TE_MULTILINE)
        self.console.SetBackgroundColour('#f4f4f4')
        vbox.Add(self.console, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, border=20, proportion=1)
        self.btn_clear = wx.Button(self.panel, BUTTON_CLEAR, label='Очистить консоль', size=(150, 30))
        self.btn_clear.SetBackgroundColour('#f6f6f6')
        vbox.Add(self.btn_clear, flag=wx.ALIGN_RIGHT | wx.RIGHT | wx.BOTTOM, border=30)
        self.panel.SetSizer(vbox)

        # Установка биндов для всех кнопок
        self.Bind(wx.EVT_BUTTON, self.add_file, id=BUTTON_ADD)
        self.Bind(wx.EVT_BUTTON, self.start_main, id=BUTTON_START)
        self.Bind(wx.EVT_BUTTON, self.onDialog, id=BUTTON_SET_NAME)
        self.Bind(wx.EVT_BUTTON, self.open_file, id=BUTTON_OPEN_EXCEL)
        self.Bind(wx.EVT_BUTTON, self.clear_console, id=BUTTON_CLEAR)

        # Создаёт ресивер пабсаба
        pub.subscribe(self.updateDisplay, "update")

    # Функция для очищения консоли
    def clear_console(self, event):
        self.console.SetValue('')

    # Функция, которая отвечает за открытие диалога
    def onDialog(self, event):
        with MyDialog(self, title="Ввод имени нового файла") as dlg:
            res = dlg.ShowModal()
            # if res == wx.ID_OK:
            # print("Нажата кнопка да")

    @staticmethod
    def my_vbox_create(button_label, st_label, button_id, border, panel):
        # Создание 1й сверху части
        vbox = wx.BoxSizer(wx.VERTICAL)
        # Создание кнопки 'Выбрать файл'
        button = wx.Button(panel, button_id, label=f'{button_label}', size=(360, 30))

        vbox.Add(button, flag=wx.EXPAND | wx.ALL, )
        # Создание сообщения: путь к файлу
        hbox = wx.BoxSizer(wx.HORIZONTAL)
        st = wx.StaticText(panel, label=f"{st_label}")
        tc = wx.TextCtrl(panel, style=wx.TE_READONLY)  # КАК ЗАПРЕТИТЬ РЕДАКТИРОВАНИЕ???
        hbox.Add(st, flag=wx.RIGHT)
        hbox.Add(tc, proportion=1)
        vbox.Add(hbox, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, border=border)

        return vbox, button, st, tc

    # Диалог для выбора файла
    def add_file(self, event):
        with wx.FileDialog(self, 'Открыть файл...', style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fileDialog:
            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return
            self.path_name = fileDialog.GetPath()  # получаем путь к файлу. Далее он используется для открытия
            self.tc1.SetValue(f"{self.path_name}")  # отображаем в текстовом поле
            self.console.WriteText(f"{time.asctime()[11:19]} - Выбран новый файл по пути: {self.tc1.GetValue()}\n")
            # print(self.path_name)

    #  Функция для запуска расчета
    def start_main(self, event):
        if self.tc1.GetValue() == '':
            self.console.WriteText(f'{time.asctime()[11:19]} - Попытка начать расчет... Нужно указать путь к файлу!\n')
        else:
            dlg1 = wx.MessageBox(
                'Вы хотите ввести свой путь для сохранения файла?\nПо умолчанию файл будет сохранен в папку, откуда запускается программа',
                "Вопрос", wx.YES | wx.NO, self)
            if dlg1 == wx.NO:  # Если пользователь не хочет ввести новый путь для файла
                thr = threading.Thread(target=self.analyze_func)  # создание потока для начала расчета
                thr.start()
                self.disable_buttons()
                # self.analyze_func()
            else:  # Если пользователь  хочет ввести новый путь для файла
                dlg2 = wx.DirDialog(self, "Выбор папки...", "C:\\", wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST)
                res = dlg2.ShowModal()
                self.console.WriteText(f'{time.asctime()[11:19]} - Выбрана папка: {dlg2.GetPath()} \n')
                time.sleep(0.5)
                thr = threading.Thread(target=self.analyze_func,
                                       args=(dlg2.GetPath(),))  # создание потока для начала расчета
                thr.start()
                self.disable_buttons()  # Отключение всех кнопок на время работы потока
                # self.analyze_func(dlg2.GetPath())

    # Функция для анализа excel - файла, которая вызывается в start_main()
    def analyze_func(self, new_path=''):
        if self.tc2.GetValue() == 'Report.xlsx (Чтобы изменить имя, нажми \"Ввести имя файла\")':
            try:
                if new_path == '':
                    msg1, msg2, msg3, msg4 = start(self.path_name, "Report.xlsx")
                else:
                    new_path = new_path + '\\'
                    msg1, msg2, msg3, msg4 = start(self.path_name, new_path + "Report.xlsx")
                self.console.WriteText(
                    str(f"{time.asctime()[11:19]} - Отчет по работе программы:" + "\n" + msg1 + "\n" + msg2 + "\n" + msg3 + "\n" + msg4 + "\n"))
                wx.CallAfter(pub.sendMessage, "update", msg="Thread finished!")  # Завершение текущего потока
                self.console.WriteText(f'{time.asctime()[11:19]} - Создан файл: Report.xlsx \n')
            except:
                self.console.WriteText(f"{time.asctime()[11:19]} - НЕИЗВЕСТНАЯ ОШИБКА! Возможно вы выбрали неверно оформленный файл\n")
                wx.CallAfter(pub.sendMessage, "update", msg="Thread finished!")  # Завершение текущего потока
        else:
            try:
                if new_path == '':
                    msg1, msg2, msg3, msg4 = start(self.path_name, self.tc2.GetValue())
                else:
                    new_path = new_path + '\\'
                    msg1, msg2, msg3, msg4 = start(self.path_name, new_path + self.tc2.GetValue())
                self.console.WriteText(
                    str(f"{time.asctime()[11:19]} - Отчет по работе программы:" + "\n" + msg1 + "\n" + msg2 + "\n" + msg3 + "\n" + msg4 + "\n"))
                wx.CallAfter(pub.sendMessage, "update", msg="Thread finished!")  # Завершение текущего потока
                self.console.WriteText(f'- Создан файл: {self.tc2.GetValue()}\n')
            except:
                self.console.WriteText(f"{time.asctime()[11:19]} - НЕИЗВЕСТНАЯ ОШИБКА! Возможно вы выбрали неверно оформленный файл\n")

    # Функция для открытия файла
    def open_file(self, event):
        if self.tc1.GetValue() == '':
            self.console.WriteText(f'{time.asctime()[11:19]} - Попытка открыть отчет... Нужно снала создать файл а, потом открывать его!\n')
        else:
            thr = threading.Thread(target=self.threading_open)  # создание потока для открытия excel - файла
            thr.start()
            self.console.WriteText(f"{time.asctime()[11:19]} - Открыт файл: Report.xlsx\n")
            self.disable_buttons()  # Отключение всех кнопок на время работы потока
            # btn = event.GetEventObject()
            # btn.Disable()

    def disable_buttons(self):
        self.btn_open_excel.Disable()
        self.btn_add.Disable()
        self.btn_clear.Disable()
        self.btn_start.Disable()
        self.btn_set_name.Disable()
    # Функция, которая вызывается отдельным потоком
    def threading_open(self):
        if self.tc2.GetValue() == 'Report.xlsx (Чтобы изменить имя, нажми \"Ввести имя файла\")':

            # subprocess.Popen(f'Report.xlsx')
            os.system('Report.xlsx')
            wx.CallAfter(pub.sendMessage, "update", msg="Thread finished!")
        else:
            self.console.WriteText(f"- Открыт файл: {self.tc2.GetValue()}\n")
            os.system(f'{self.tc2.GetValue()}')
            wx.CallAfter(pub.sendMessage, "update", msg="Thread finished!")  # Завершение текущего потока
            # subprocess.Popen(f'{self.tc2.GetValue()}')
            # os.startfile(f'{self.tc2.GetValue()}')
            return

    def updateDisplay(self, msg):
        """
        Получает данные из потока и обновляет дисплей
        """
        t = msg
        if type(msg) == str:
            self.btn_open_excel.Enable()
            self.btn_add.Enable()
            self.btn_clear.Enable()
            self.btn_start.Enable()
            self.btn_set_name.Enable()


# Основная функция, с которой происходит запуск программы
if __name__ == "__main__":
    elevate()
    app = wx.App()
    frame = MyFrame(None, title='Pallet Counter')
    frame.Center()
    frame.Show()
    app.MainLoop()
