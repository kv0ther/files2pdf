import tkinter
import tkinter.filedialog
import win32com
import win32com.client
root = tkinter.Tk()


class Application:

    def __init__(self):
        self.root = root
        self.tela()
        self.frame_da_tela()
        root.mainloop()

    def tela(self):
        self.root.title('desgraça')
        self.root.configure(background='#0D214F')
        self.root.geometry('740x440')
        self.root.resizable(True, True)
        self.root.minsize(width=740, height=440)

    def buscar(self):
        self.orig_path= tkinter.filedialog.askopenfile(initialdir='/Documentos')
        self.path_text.delete(1.0, 'end')
        self.path_text.insert(1.0, self.orig_path.name)

    def new_buscar(self):
        self.path_new= tkinter.filedialog.askopenfile(initialdir='/Documentos')
        self.new_path.delete(1.0, 'end')
        self.new_path.insert(1.0, self.path_new.name)



    def frame_da_tela(self):
        self.frame_1 = tkinter.Frame(self.root, bd=4, bg='#ddd9ce', highlightbackground='black', highlightthickness=3)
        self.frame_1.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.46)

        # buscar e texto do caminho do arquivo
        self.bt_buscar_1 = tkinter.Button(self.frame_1, text='buscar', command=self.buscar)
        self.bt_buscar_1.place(relx=0.89, rely=0.08, relwidth=0.1, relheight=0.15)

        # buscar e texto do caminho do arquivo
        self.bt_buscar_1 = tkinter.Button(self.frame_1, text='buscar', command=self.new_buscar)
        self.bt_buscar_1.place(relx=0.89, rely=0.33, relwidth=0.1, relheight=0.15)

        self.frame_text = tkinter.Frame(self.frame_1)
        self.frame_text.place(relx=0.01, rely=0.096, relwidth=0.87, relheight=0.12)
        self.path_text = tkinter.Text(self.frame_text)
        self.path_text.pack(expand=True, fill='x')

        #buscar e novo caminho
        self.frame_new_path = tkinter.Frame(self.frame_1)
        self.frame_new_path.place(relx=0.01, rely=0.35, relwidth=0.87, relheight=0.12)
        self.new_path = tkinter.Text(self.frame_new_path)
        self.new_path.pack(expand=True, fill='x')

        self.botão = tkinter.Button(self.frame_1, text='print', command=self.conversion)
        self.botão.place(relx=0.3, rely=0.85, relwidth=0.3, relheight=0.12)

    def conversion(self):
        wdFormatPDF = 17

        self.entrada = self.path_text.get(1.0, 'end')
        self.saida = self.new_path.get(1.0, 'end')

        print(self.path_text)
        print(self.new_path)
        print(self.entrada)
        print(self.saida)

Application()
