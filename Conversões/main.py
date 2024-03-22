import pathlib
from tkinter import *
import tkinter.filedialog
import win32com
root = Tk()


class Application:

    def __init__(self):
        self.root = root
        self.tela()
        self.frame_da_tela()
        self.widgets_frame1()
        self.widgets_frame2()
        root.mainloop()

    def tela(self):
        self.root.title('Conversor')
        self.root.configure(background='#0D214F')
        self.root.geometry('740x440')
        self.root.resizable(True, True)
        self.root.minsize(width=740, height=440)

    def frame_da_tela(self):
        self.frame_1 = Frame(self.root, bd=4, bg='#ddd9ce', highlightbackground='black', highlightthickness=3)
        self.frame_1.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.46)

        self.frame_2= Frame(self.root, bd=4, bg='#ddd9ce', highlightbackground='black', highlightthickness=3)
        self.frame_2.place(relx=0.02, rely=0.5, relwidth=0.96, relheight=0.46)

    def widgets_frame1(self):
        self.frame_1 = tkinter.Frame(self.root, bd=4, bg='#ddd9ce', highlightbackground='black', highlightthickness=3)
        self.frame_1.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.46)

        # botão de busca 1
        self.bt_buscar_1 = tkinter.Button(self.frame_1, text='buscar', command=self.buscar_1)
        self.bt_buscar_1.place(relx=0.89, rely=0.09, relwidth=0.1, relheight=0.15)

        # botão de buscar 2
        self.bt_buscar_2 = tkinter.Button(self.frame_1, text='buscar', command=self.buscar_2)
        self.bt_buscar_2.place(relx=0.89, rely=0.37, relwidth=0.1, relheight=0.15)

        # label e text do caminho
        self.lb_path = Label(self.frame_1, text='caminho do arquivo original', background='#ddddce')
        self.lb_path.place(relx=0.01, rely=0.001, relwidth=0.22, relheight=0.1)
        self.frame_text = Frame(self.frame_1)
        self.frame_text.place(relx=0.01, rely=0.1, relwidth=0.87, relheight=0.12)
        self.path_text = Text(self.frame_text)
        self.path_text.pack(expand=True)

        # label e text do novo caminho
        self.lb_new_path = Label(self.frame_1, text='local do novo arquivo', background='#ddddce')
        self.lb_new_path.place(relx=0.01, rely=0.24, relwidth=0.17, relheight=0.14)
        self.frame_new_path = Frame(self.frame_1)
        self.frame_new_path.place(relx=0.01, rely=0.39, relwidth=0.87, relheight=0.12)
        self.new_path_text = Text(self.frame_new_path)
        self.new_path_text.configure(state='disabled')
        self.new_path_text.pack(expand=True, fill='x')


        # label e entrada do nome do aquivo convertido
        self.lb_name = Label(self.frame_1, text='nome do arquivo', background='#ddddce')
        self.lb_name.place(relx=0.01, rely=0.55, relwidth=0.14, relheight=0.08)

        self.new_name = Entry(self.frame_1)
        self.new_name.place(relx=0.01, rely=0.65, relwidth=0.87, relheight=0.1)

    def widgets_frame2(self):
        # criando botão de word2pdf
        self.bt_word2pddf = Button(self.frame_2, text='Word para PDF', command=self.word_for_pdf)
        self.bt_word2pddf.place(relx=0.01, rely=0.02, relwidth=0.13, relheight=0.15)

        self.bt_excel2pdf = Button(self.frame_2, text='Excel para PDF', command=self.excel_for_pdf)
        self.bt_excel2pdf.place(relx=0.15, rely=0.02, relwidth=0.13, relheight=0.15)

# criando função de conversão
    def buscar_1(self):
        self.path_orig = tkinter.filedialog.askopenfile(mode='r', initialdir='/Documentos',
                                                        initialfile='C:/Users/rafae/OneDrive/Área de Trabalho'
                                                                    '/Memorial_descritivo.docx',
                                                        title="Selecione um arquivo",
                                                        filetypes=(
                                                            ("Arquivos doc", ".docx"), ("arquivos xlsx", ".xlsx")))
        self.path_text.delete(1.0, 'end')
        self.path_text.insert(1.0, self.path_orig.name)

    def buscar_2(self):
        self.new_path_text.configure(state='normal')
        self.new_path = tkinter.filedialog.askdirectory(initialdir='/Documentos', title="Selecione um arquivo")
        self.new_path_text.delete(1.0, 'end')
        self.new_path_text.insert(1.0, self.new_path)
        self.new_path_text.configure(state='disabled')

    def excel_for_pdf(self):
        excel = win32com.client.Dispatch("Excel.Application")

        self.entrada = self.path_text.get(1.0)
        self.saida = self.new_path_text.get(1.0)

        sheets = excel.Workbooks.Open(self.entrada)
        work_sheets = sheets.Worksheets[0]

        work_sheets.ExportAsFixedFormat(0, self.saida)


    def word_for_pdf(self):
        wdFormatPDF = 17

        self.entrada = self.path_text.get(1.0, 'end')
        self.saida = self.new_path_text.get(1.0, 'end')

        self.x = pathlib.PurePath(self.entrada)
        self.y = self.saida

        print(self.x)
        print(type(self.x))

        word = win32com.client.Dispatch('Word.Application')
        doc = word.Documents.Open(self.x)
        doc.SaveAs(self.y, FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()


Application()
