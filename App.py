import os

from tkinter import filedialog as fd
import customtkinter

from PIL import Image

import file_analize as fa

class App(customtkinter.CTk):
    """
    Класс главного окна приложения

    :param APP_WIDTH: Ширина главного окна приложения
    :type APP_WIDTH: int

    :param APP_HEIGHT: Высота главного окна приложения
    :type APP_HEIGHT: int

    :param X_APP: Смещение для центровки по горизонтали
    :type X_APP: int

    :param Y_APP: Смещение для центровки по вертикали
    :type Y_APP: int

    :param VERSION: Версия приложения
    :type VERSION: str

    :param main_bar_frame: Верхняя панель главного окна с логотипом
    :type main_bar_frame: custom class (Main_bar), from CTkFrame

    :param check_info_frame: Панель для получения данных от пользователя
    :type check_info_frame: custom class (Check_info), from CTkFrame

    :param instruction: Инструкция пользования приложением
    :type instruction: custom class (Instruction), from CTkTextbox

    :param buttons_font: Шрифт для кнопок
    :type buttons_font: CTkFont

    :param labels_font: Шрифт для подписей
    :type labels_font: CTkFont
    
    :param instructions_font: Шрифт для инструкций
    :type instructions_font: CTkFont

    """

    def __init__(self):
        super().__init__()
        self.params()
        self.find_center()
        self.title(f'| Converter v {self.VERSION} |')
        self.geometry(f"{self.APP_WIDTH}x{self.APP_HEIGHT}+{int(self.X_APP)}+{int(self.Y_APP)}")
        self.minsize(780,550)
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.put_main_frames()
        
        self.keyboard_bind()

    def params(self) -> None:
        """
        Создание основных параметров приложения
        
        :return: None
        """

        customtkinter.set_appearance_mode('Dark')
        customtkinter.set_default_color_theme('blue')
        self.VERSION = '0.0.1.2'
        self.APP_WIDTH = 780
        self.APP_HEIGHT = 800
        
        self.buttons_font = customtkinter.CTkFont("Avenir Next", 14, 'normal')
        self.labels_font = customtkinter.CTkFont("Avenir Next", 14, 'normal')
        self.instructions_font = customtkinter.CTkFont("Avenir Next", 13, 'normal')

        return 0
    
    def find_center(self) -> None:
        """
        Функция для нахождения смещений на экране для центровки
        
        :return: None
        """
        SCREEN_WIDTH = self.winfo_screenwidth()
        SCREEN_HEIGHT = self.winfo_screenheight()

        self.X_APP = (SCREEN_WIDTH / 2) - (self.APP_WIDTH / 2)
        self.Y_APP = (SCREEN_HEIGHT / 2) - (self.APP_HEIGHT / 2)
        
        return None

    def put_main_frames(self) -> None:
        """
        Размещение виджетов главного окна

        :return: None
        """
        self.main_bar_frame = Main_bar(self)
        self.main_bar_frame.pack(padx=5, pady=[5,0], fill='x')

        self.check_info_frame = Check_info(self)
        self.check_info_frame.pack(padx=5, pady=5, side='left', fill='both', expand=True)

        self.instruction = Instruction(self)
        self.instruction.pack(padx=[0,5], pady=5, side='right', fill='both', expand=True)

        return None
    
    def on_closing(self, event=0) -> None:
        """
        Функция для объявления закрытия окна
        
        :return: None
        """
        self.destroy()
        return None

    def keyboard_bind(self) -> None:
        """
        Привязка нажатий комбинаций клавиш к функциям приложения
        
        :return: None
        """
        self.bind('<Control-q>', lambda event : self.quit())
        
        return None

    def reload(self) -> None:
        """
        Функция перезагрузки приложения до начального состояния
        
        :return: None
        """
        self.check_info_frame.destroy()
        self.check_info_frame = Check_info(self)
        self.check_info_frame.pack(padx=5, pady=5, side='left', fill='both', expand=True)
        
        return None

    def add_correct_save_label(self) -> None:
        """
        Функция отображения сообщения о корректном сохранении файла
        
        :return: None
        """
        self.reload()
        self.main_bar_frame.reload_button.configure(state='normal')
        self.save_img = customtkinter.CTkImage(light_image=Image.open(os.path.abspath("./Design/save.png")), size=(48, 60))
        self.correct_save_label = customtkinter.CTkLabel(self.check_info_frame, image=self.save_img,
                                                         text=' Обработка успешно завершена, файл сохранен. ', font=self.labels_font,
                                                         compound='top', padx=5, pady=10)
        self.correct_save_label.pack(fill='both', expand=True)

class Main_bar(customtkinter.CTkFrame):
    """
    Класс верхнего тулл-бара приложения

    :param logo_img: Изображение логотипа
    :type logo_img: CTkImage

    :param logo: Фрейм с логотипом
    :type logo: CTkLabel 

    :param reload_img: Изображение для кнопки перезагрузки
    :type reload_img: CTkImage 

    :param reload_button: Кнопка перезагрузки
    :type reload_button: CTkButton
    """

    def __init__(self, master):
        super().__init__(master, height=50, corner_radius=5)

        self.logo_img = customtkinter.CTkImage(light_image=Image.open(os.path.abspath("./Design/logo.png")), size=(120, 40))
        self.logo = customtkinter.CTkLabel(self, image=self.logo_img, text='')
        self.logo.pack(padx=5, pady=5, side='left')

        self.scaling_optionemenu = customtkinter.CTkOptionMenu(self, height=30, width=50,
                                                               values=["100%", "90%", "80%"],
                                                               font=self.master.buttons_font,
                                                               command=self.change_scaling_event)
        self.scaling_optionemenu.pack(padx=[0,5], pady=5, side='right')
        
        self.scaling_label = customtkinter.CTkLabel(self, height=30, width=50,
                                                    text="Маштаб интерфейса:", font=self.master.labels_font,)
        self.scaling_label.pack(padx=[0,5], pady=5, side='right')

        self.reload_img = customtkinter.CTkImage(light_image=Image.open(os.path.abspath("./Design/restart.png")), size=(20, 20))
        self.reload_button = customtkinter.CTkButton(self, height=30,
                                                     text='| Очистить поле ввода и выбора', font=master.buttons_font,
                                                     image=self.reload_img, compound='right',
                                                     command=lambda:master.reload())
        self.reload_button.pack(padx=[0,5], pady=5, side='right')

    def change_scaling_event(self, new_scaling: str)-> None:
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        customtkinter.set_widget_scaling(new_scaling_float)
        customtkinter.set_window_scaling(new_scaling_float)
        return None

class Instruction(customtkinter.CTkTextbox):
    """
    Класс фрейма инструкции для приложения

    """

    def __init__(self, master):
        super().__init__(master, font=master.instructions_font, corner_radius=5, wrap='word')
        self.insert('0.0', ''.join(open(os.path.abspath('.//instruction.txt'), encoding='utf-8').readlines()))
        self.configure(state='disabled')

class Check_info(customtkinter.CTkFrame):
    """
    Класс окна получения информации от пользователя

    :param choose_file_button: Кнопка выбора файла
    :type choose_file_button: CTkButton

    :param file_path: Абсолютный путь к выбранному файлу
    :type file_path: str

    :param sheets: Список названий листов в excel -  файле
    :type sheets: list

    :param source_label: Подпись к окну выбора листов для получения информации
    :type source_label: CTkLabel

    :param source_frame: Фрейм, содержащий интерфейс для выбора нужных листов
    :type source_frame: CTkScrollableFrame

    :param process_label: Подпись к окну выбора листов для обработки
    :type process_label: CTkLabel

    :param process_frame: Фрейм, содержащий интерфейс для выбора нужных листов
    :type process_frame: CTkScrollableFrame

    :param error_label: Подпись к окну выбора листов для ошибок
    :type error_label: CTkLabel

    :param error_frame: Фрейм, содержащий интерфейс для выбора нужных листов
    :type error_frame: CTkScrollableFrame

    :param commit_button: Фрейм, содержащий интерфейс для выбора нужных листов
    :type commit_button: CTkButton

    :param selected_sheets: Список, содержащий названия листов, с меткой типа
    :type selected_sheets: list
    """

    def __init__(self, master):
        super().__init__(master, corner_radius=5)

        self.choose_source_button = customtkinter.CTkButton(self, height=20,
                                                          text='Выбор файла с источником данных:', font=self.master.buttons_font, anchor='w',
                                                          command=lambda: self.choose_source_file())
        self.choose_source_button.pack(padx=5, pady=[5,10], fill='x', expand=False)

    def choose_source_file(self) -> None:
        """
        Получение абсолютного пути к файлу с источником данных

        :return: None
        """

        self.source_file_path = fd.askopenfile(filetypes=[('Excel tables', '*.xlsx')])
        if self.source_file_path != '':
            try: self.master.correct_save_label.destroy()
            except Exception as e: pass
            self.choose_source_button.configure(text=f'Выбрано: {os.path.basename(self.source_file_path.name)}', state='disabled')
            self.sheets = fa.source_sheets(self.source_file_path.name)
            self.create_source_selections()
        
        return None

    def create_source_selections(self) -> None:
        """
        Создание дополнительных кнопок для выбора листов источников
        
        :return: None
        """

        self.source_label = customtkinter.CTkLabel(self, height=20,
                                                   text='–-- Выбор листов с источниками данных ---', font=self.master.labels_font)
        self.source_label.pack(padx=0, pady=[0,5], fill='both', expand=False)

        self.source_frame = customtkinter.CTkScrollableFrame(self, height=40)
        self.source_frame.pack(padx=5, pady=[0,5], fill='both', expand=False)

        self.source_switches = []
 
        for sheet in self.sheets:
            self.source_switches.append(customtkinter.CTkCheckBox(self.source_frame,
                                                                    text=sheet[0], 
                                                                    font=self.master.buttons_font,
                                                                    command=None))
        
        for i in range(len(self.sheets)):
            self.source_switches[i].grid(row=i, column=0, padx=[1,10], pady=1, sticky='w')
            self.source_info_label = customtkinter.CTkLabel(self.source_frame,
                                                       text=f'| Кол-во столбцов: {self.sheets[i][1]}, строк: {self.sheets[i][2]}',
                                                       font=self.master.buttons_font)
            self.source_info_label.grid(row=i, column=1, padx=1, pady=1, sticky='e')
        
        self.choose_lsr_directory_button = customtkinter.CTkButton(self, height=20,
                                                          text='Выбор папки с файлами ЛСР:', font=self.master.buttons_font, anchor='w',
                                                          command=lambda: self.choose_lsr_directory())
        self.choose_lsr_directory_button.pack(padx=5, pady=[10,5], fill='x', expand=False)
        
        return None
    
    def choose_lsr_directory(self) -> None:
        """
        Получение абсолютного пути к папке с ЛСР

        :return: None
        """

        self.lsr_directory_path = fd.askdirectory()
        if self.lsr_directory_path != '':
            self.choose_lsr_directory_button.configure(text=f'Выбрано: {os.path.basename(self.lsr_directory_path)}', state='disabled')
            fa.lsr_documents(self.lsr_directory_path)
            self.lsr_documents = fa.lsr_documents(self.lsr_directory_path)
            self.create_lsr_selections()
        
        return None
    
    def create_lsr_selections(self) -> None:
        """
        Создание дополнительных кнопок для выбора листов источников и папки с ЛСР
        
        :return: None
        """

        self.lsr_label = customtkinter.CTkLabel(self, height=20,
                                                   text='–-- документы ЛСР для обработки ---', font=self.master.labels_font)
        self.lsr_label.pack(padx=0, pady=[0,5], fill='both', expand=False)

        self.lsr_frame = customtkinter.CTkScrollableFrame(self, height=40)
        self.lsr_frame.pack(padx=5, pady=[0,5], fill='both', expand=False)

        self.lsr_switches = []
 
        for lsr in self.lsr_documents:
            self.lsr_switches.append(customtkinter.CTkCheckBox(self.lsr_frame,
                                                                    text=lsr[0], 
                                                                    font=self.master.buttons_font,
                                                                    command=None))
        
        for i in range(len(self.lsr_documents)):
            self.lsr_switches[i].grid(row=i, column=0, padx=[1,10], pady=1, sticky='w')
            self.lsr_info_label = customtkinter.CTkLabel(self.lsr_frame,
                                                       text='| '+self.lsr_documents[i][1],
                                                       font=self.master.buttons_font)
            self.lsr_info_label.grid(row=i, column=1, padx=10, pady=1, sticky='w')
        
        self.choose_all_lsr_button = customtkinter.CTkCheckBox(self, height=20, width=70,
                                                          text='Выбрать все', font=self.master.buttons_font,
                                                          command=lambda: self.choose_all_lsr())
        self.choose_all_lsr_button.pack(padx=5, pady=[0,10], anchor='w', expand=False)


        self.commit_button = customtkinter.CTkButton(self, height=20,
                                                          text='Запустить обработку файла', font=self.master.buttons_font, anchor='n',
                                                          command=lambda: self.start_process())
        self.commit_button.pack(padx=5, pady=5, fill='x', expand=False, side='bottom')
        
        return None

    def choose_all_lsr(self) -> None:
        """
        Функция выбора всех ЛСР
        
        :return: None
        """
        if self.choose_all_lsr_button.get():
            for elem in self.lsr_switches:
                elem.select()
        else:
            for elem in self.lsr_switches:
                elem.deselect()
        return None
    
    
    def get_states(self) -> None:
        """
        Служебная функция для проверки выбранных листов
        
        :return: None
        """
        self.source_sheets = []
        self.lsr_documents_states = []
        for button in self.source_switches:
            if button.get(): self.source_sheets.append(button.cget('text'))
        for i in range(len(self.lsr_documents)):
            if self.lsr_switches[i].get(): self.lsr_documents_states.append(self.lsr_documents[i])
        
        return None
    
    
    def start_process(self) -> None:
        """
        Служебная функция для запуска процесса обработки файла
        
        :return: None
        """
        self.commit_button.configure(state='disabled')
        self.master.main_bar_frame.reload_button.configure(state='disabled')
        self.get_states()
        fa.insert_info(self.source_file_path.name, self.source_sheets, self.lsr_documents_states)
        self.master.add_correct_save_label()
        return None

if __name__ == "__main__":
    app = App()
    app.mainloop()