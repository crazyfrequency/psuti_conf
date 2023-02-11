import customtkinter
from PSUTI import PSUTI, check_url
import sys, signal, pathlib
from logic import check_path, Download, MakeConf
from threading import Thread

INFO = """Инструкция
Если конференция была скаченена не из приложения, то выберите галочку \"не стандартная конференция\".
В этом режиме программа не учитывает данные в сформированных файлах(созданные при загрузке конференции через приложение).
Заголовки секций - это название папок(если секций нет поместить все файлы в основную папку).

\"Добавить огловление по авторам\" - в конец сборника добавляется оглавление по авторам(доступно только при скачивании через приложение).
Пример:
Фамилия И.О. с. 12-13,14,54-56.

\"путь\" - путь к файлу куда сохранить итоговый файл.
"""

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("dark-blue")

class MainWindow(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.resizable(False,False)
        self.menu()
        self.psuti = None

    def menu(self):
        self.title("psuti conf")
        self.geometry("400x150")
        self.clear()
        button1 = customtkinter.CTkButton(master=self, text="Скачать файлы", command=self.download)
        button1.pack(pady=20, padx=10, fill="both", side="left", expand=True)
        button2 = customtkinter.CTkButton(master=self, text="Собрать сборник", command=self.conf_settings)
        button2.pack(pady=20, padx=10, fill="both", side="right", expand=True)

    def download(self):
        self.title("psuti conf - download")
        self.geometry("400x150")
        self.clear()
        frame = customtkinter.CTkFrame(self)
        frame.pack(pady=5, padx=10, fill="both", expand=True)
        label = customtkinter.CTkLabel(frame, text="Ссылка на конференцию")
        label.pack(side="top", fill="x", expand=True, padx=10, pady=10)
        entry = customtkinter.CTkEntry(
            frame,
            placeholder_text="https://conf.psuti.ru/snk61"
        )
        entry.pack(pady=5, padx=10, side="bottom", fill="x", expand=True)
        button1 = customtkinter.CTkButton(self, text="ок", command=lambda: self.login(entry.get()) if check_url(entry.get()) else self.error("неверный адрес сайта"))
        button1.pack(pady=10, padx=10, fill="both", side="left", expand=True)
        button2 = customtkinter.CTkButton(self, text="отмена", command=self.menu)
        button2.pack(pady=10, padx=10, fill="both", side="right", expand=True)
        
    def login(self, url: str):
        self.psuti = PSUTI(url)
        window = customtkinter.CTkToplevel(self)
        window.geometry("400x230")
        window.resizable(False,False)
        window.title("login")
        window.grab_set()
        window.focus()

        frame = customtkinter.CTkFrame(window)
        frame.pack(pady=5, padx=10, fill="both", expand=True)
        entry1 = customtkinter.CTkEntry(
            frame,
            placeholder_text="test@gmail.com"
        )
        entry1.pack(pady=5, padx=10, fill="x", expand=True)
        label1 = customtkinter.CTkLabel(frame, text="почта")
        label1.pack(before=entry1, fill="x", expand=True, padx=10, pady=10)
        entry2 = customtkinter.CTkEntry(frame,show='*', width=300)
        entry2.pack(pady=5, padx=10, side="left", fill="x", expand=True)
        label2 = customtkinter.CTkLabel(frame, text="пароль")
        label2.pack(before=entry2, fill="x", expand=True, padx=10, pady=10)


        checkbox = customtkinter.CTkCheckBox(frame, text="", width=0, command=lambda:entry2.configure(show='*') if checkbox.get()==0 else entry2.configure(show=''))
        checkbox.pack(pady=5, padx=1, side="right", expand=True)

        button1 = customtkinter.CTkButton(window, text="ок", command=lambda:self.pre_download(window) if self.psuti.login(entry1.get(),entry2.get()) else self.error("неверный логин или пароль", window))
        button1.pack(pady=10, padx=10, fill="both", side="left", expand=True)
        button2 = customtkinter.CTkButton(window, text="отмена", command=window.destroy)
        button2.pack(pady=10, padx=10, fill="both", side="right", expand=True)
        
    def pre_download(self, window:customtkinter.CTkToplevel=None):
        if(window!=None):window.destroy()
        self.geometry("400x180")
        self.clear()
        frame = customtkinter.CTkFrame(self)
        frame.pack(pady=5, padx=10, fill="both", expand=True)
        label1 = customtkinter.CTkLabel(frame, text="Папка для загрузки")
        label1.pack(fill="x", expand=True, side="top", padx=10, pady=10)
        entry = customtkinter.CTkEntry(
            frame,
            width=300,
            placeholder_text=str(pathlib.Path.cwd()).replace("\\","/")
        )
        entry.configure(state="disabled")
        entry.pack(pady=5, padx=10, side="left", expand=True)

        def set_text(text):
            if(not text):return
            entry.configure(state="normal")
            entry.configure(placeholder_text=text)
            entry.configure(state="disabled")

        button = customtkinter.CTkButton(frame, text="путь", width=28, command=lambda:set_text(customtkinter.filedialog.askdirectory()))
        button.pack(pady=10, padx=5, side="right", expand=True)
        button1 = customtkinter.CTkButton(self, text="ок", command=lambda: self.downloading(entry._placeholder_text) if check_path(entry._placeholder_text) else self.error("Папка не пуста"))
        button1.pack(pady=10, padx=10, fill="both", side="left", expand=True)
        button2 = customtkinter.CTkButton(self, text="отмена", command=self.menu)
        button2.pack(pady=10, padx=10, fill="both", side="right", expand=True)

    def downloading(self, path:str):
        self.title("psuti conf - downloading")
        self.geometry("400x130")
        self.clear()
        frame = customtkinter.CTkFrame(self)
        frame.pack(pady=5, padx=10, fill="both", expand=True)
        label = customtkinter.CTkLabel(frame, text="Загрузка")
        label.pack(fill="x", expand=True, side="top", padx=10, pady=10)
        progressbar = customtkinter.CTkProgressBar(frame)
        progressbar.pack(pady=1, padx=10, fill="both", expand=True)
        progressbar.set(0)
        label1 = customtkinter.CTkLabel(frame, text="0/-")
        label1.pack(expand=True, padx=1, side="bottom", pady=1)

        download = Download(path,self.psuti)

        button = customtkinter.CTkButton(self, text="отмена", command=download.stop)
        button.pack(pady=5, padx=10, side="bottom", expand=True)

        def f(downloaded:int, count:int):
            label1.configure(text=f'{downloaded}/{count}')
            progressbar.set(downloaded/count)
        
        Thread(target=download.start,args=[f,self.menu]).start()
        
    def downloaded(self):
        self.title("psuti conf - downloaded")
        self.geometry("400x130")
        label = customtkinter.CTkLabel(self, text="Успешно загружено")
        label.pack(side="top", fill="both", expand=True, padx=5, pady=5)
        button = customtkinter.CTkButton(self, height=28, text="ок", command=self.menu)
        button.pack(pady=10, padx=10, fill="x", side="bottom", expand=True)

    def conf_settings(self):
        self.title("psuti conf - make settings")
        self.geometry("400x470")
        self.clear()
        frame = customtkinter.CTkFrame(self)
        frame.pack(pady=5, padx=10, fill="both", expand=True)
        label = customtkinter.CTkLabel(frame, text="Параметры сборки")
        label.pack(fill="x", expand=True, side="top", padx=10, pady=10)
        button = customtkinter.CTkButton(frame, text="пояснение", command=lambda:self.info(text=INFO))
        button.pack(pady=5, padx=10, side="bottom", expand=True)
        checkbox2 = customtkinter.CTkCheckBox(frame, text="Добавить огловление по авторам")

        def f(bool:bool):
            if(bool):
                checkbox2.deselect()
                checkbox2.configure(state="disabled")
            else:
                checkbox2.configure(state="normal")

        checkbox1 = customtkinter.CTkCheckBox(frame, text="Не стандартная конференция", command=lambda:f(not not checkbox1.get()))
        checkbox1.pack(pady=5, padx=10, expand=True)
        checkbox2.pack(pady=5, padx=10, expand=True)
        label1 = customtkinter.CTkLabel(frame, text="Папка с файлами")
        label1.pack(fill="x", expand=True, after=checkbox2, padx=10, pady=10)
        frame1 = customtkinter.CTkFrame(frame)
        frame1.pack(pady=5, padx=10, fill="both", expand=True)
        entry1 = customtkinter.CTkEntry(
            frame1,
            width=300,
            placeholder_text=str(pathlib.Path.cwd()).replace("\\","/")
        )
        entry1.configure(state="disabled")
        entry1.pack(side="left", expand=True)

        def set_text1(text):
            if(not text):return
            entry1.configure(state="normal")
            entry1.configure(placeholder_text=text)
            entry1.configure(state="disabled")

        button3 = customtkinter.CTkButton(frame1, text="путь", width=28, command=lambda:set_text1(customtkinter.filedialog.askdirectory()))
        button3.pack(side="right", expand=True)

        label2 = customtkinter.CTkLabel(frame, text="Итоговый файл")
        label2.pack(fill="x", expand=True, after=frame1, padx=10, pady=10)
        frame2 = customtkinter.CTkFrame(frame)
        frame2.pack(pady=5, padx=10, fill="both", expand=True)
        entry2 = customtkinter.CTkEntry(
            frame2,
            width=300,
            placeholder_text=str(pathlib.Path.cwd()).replace("\\","/")+"/res.docx"
        )
        entry2.configure(state="disabled")
        entry2.pack(side="left", expand=True)

        def set_text2(text:str):
            if(not text):return
            if(not text.endswith(tuple([".doc",".docx",".pdf"]))):text+=".docx"
            entry2.configure(state="normal")
            entry2.configure(placeholder_text=text)
            entry2.configure(state="disabled")

        button4 = customtkinter.CTkButton(frame2, text="путь", width=28, command=lambda:set_text2(customtkinter.filedialog.asksaveasfilename(filetypes=[("Word files",".docx .doc .pdf")])))
        button4.pack(side="right", expand=True)

        label3 = customtkinter.CTkLabel(frame, text="Первые 3 страницы(без оглавления)")
        label3.pack(fill="x", expand=True, after=frame2, padx=10, pady=10)
        frame3 = customtkinter.CTkFrame(frame)
        frame3.pack(pady=5, padx=10, fill="both", expand=True)
        entry3 = customtkinter.CTkEntry(
            frame3,
            width=300,
            placeholder_text=str(pathlib.Path.cwd()).replace("\\","/")+"/res.docx"
        )
        entry3.configure(state="disabled")
        entry3.pack(side="left", expand=True)

        def set_text3(text:str):
            if(not text):return
            if(not text.endswith(tuple([".doc",".docx",".pdf"]))):text+=".docx"
            entry3.configure(state="normal")
            entry3.configure(placeholder_text=text)
            entry3.configure(state="disabled")

        button5 = customtkinter.CTkButton(frame3, text="путь", width=28, command=lambda:set_text3(customtkinter.filedialog.askopenfilename(filetypes=[("Word files",".docx .doc")])))
        button5.pack(side="right", expand=True)

        button1 = customtkinter.CTkButton(self, text="ок", command=lambda: self.conf_maker([
                not checkbox1.get(),
                checkbox2.getboolean(1),
                entry1._placeholder_text,
                entry2._placeholder_text,
                entry3._placeholder_text
            ]))
        button1.pack(pady=10, padx=10, fill="both", side="left", expand=True)
        button2 = customtkinter.CTkButton(self, text="отмена", command=self.menu)
        button2.pack(pady=10, padx=10, fill="both", side="right", expand=True)

    def conf_maker(self, data:list):
        self.title("psuti conf - making")
        self.geometry("400x130")
        self.clear()
        frame = customtkinter.CTkFrame(self)
        frame.pack(pady=5, padx=10, fill="both", expand=True)
        label = customtkinter.CTkLabel(frame, text="Обработка")
        label.pack(fill="x", expand=True, side="top", padx=10, pady=10)
        progressbar = customtkinter.CTkProgressBar(frame)
        progressbar.pack(pady=1, padx=10, fill="both", expand=True)
        progressbar.set(0)
        label1 = customtkinter.CTkLabel(frame, text="0/-")
        label1.pack(expand=True, padx=1, side="bottom", pady=1)

        download = MakeConf()

        button = customtkinter.CTkButton(self, text="отмена", command=download.stop)
        button.pack(pady=5, padx=10, side="bottom", expand=True)

        def f(downloaded:int, count:int):
            label1.configure(text=f'{downloaded}/{count}')
            progressbar.set(downloaded/count)
        
        Thread(target=download.start,args=[f,*data,self.conf_maked,self.menu,lambda err: self.conf_make_error(err)]).start()

    def conf_maked(self):
        self.title("psuti conf - maked")
        self.geometry("400x130")
        self.clear()
        label = customtkinter.CTkLabel(self, text="Успешно сформировано")
        label.pack(side="top", fill="both", expand=True, padx=5, pady=5)
        button = customtkinter.CTkButton(self, height=28, text="ок", command=self.menu)
        button.pack(pady=10, padx=10, fill="x", side="bottom", expand=True)

    def conf_make_error(self, text):
        self.geometry("200x130")
        self.clear()
        self.title("error")
        label = customtkinter.CTkLabel(self, text="error")
        label.pack(side="top", fill="both", expand=True, padx=5, pady=5)
        button = customtkinter.CTkButton(self, height=28, text="ок", command=self.menu)
        button.pack(pady=10, padx=10, fill="x", side="bottom", expand=True)
        self.info("error",text)

    def info(self, title:str="info", text:str="", root=None):
        if(not root):root = self
        window = customtkinter.CTkToplevel(root)
        window.geometry("300x250")
        window.minsize(250,200)
        window.title(title)
        text_obg = customtkinter.CTkTextbox(window,)
        text_obg.insert("0.0", text)
        text_obg.configure(state="disabled")
        text_obg.pack(side="top", fill="both", expand=True, padx=5, pady=5)

    def error(self, text:str="error", root=None):
        if(not root):root = self
        window = customtkinter.CTkToplevel(root)
        window.geometry("200x150")
        window.resizable(False,False)
        window.title("error")
        label = customtkinter.CTkLabel(window, text=text)
        label.pack(side="top", fill="both", expand=True, padx=5, pady=5)
        button = customtkinter.CTkButton(window, height=28, text="ок", command=window.destroy)
        button.pack(pady=10, padx=10, fill="x", side="bottom", expand=True)
        window.focus()

    def new_window(self):
        window = customtkinter.CTkToplevel(self)
        window.geometry("550x400")
        window.resizable(False,False)
        label = customtkinter.CTkLabel(window, text="CTkToplevel window")
        label.pack(side="top", fill="both", expand=True, padx=40, pady=40)
        window.grab_set()
        window.focus()
    
    def clear(self):
        keys = self.children.copy().keys()
        for i in keys:
            self.children.get(i).destroy()

if __name__ == "__main__":
    def signal_handler(sig, frame):
        sys.exit(0)

    signal.signal(signal.SIGINT, signal_handler)
    signal.signal(signal.SIGTERM, signal_handler)

    MainWindow().mainloop()