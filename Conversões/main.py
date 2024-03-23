from tkinter import *
import tkinter.filedialog
import win32com
from win32com import client
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
        self.frame_1.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.56)

        self.frame_2= Frame(self.root, bd=4, bg='#ddd9ce', highlightbackground='black', highlightthickness=3)
        self.frame_2.place(relx=0.02, rely=0.6, relwidth=0.96, relheight=0.36)

    def widgets_frame1(self):
        # botões de busca
        self.bt_buscar_1 = tkinter.Button(self.frame_1, text='buscar', command=self.buscar_1)
        self.bt_buscar_1.place(relx=0.01, rely=0.24, relwidth=0.1, relheight=0.1)

        self.bt_buscar_2 = tkinter.Button(self.frame_1, text='buscar', command=self.buscar_2)
        self.bt_buscar_2.place(relx=0.01, rely=0.61, relwidth=0.1, relheight=0.1)

        # Entry
        self.Entry_1 = Entry(self.frame_1)
        self.Entry_1.place(relx=0.01, rely=0.115, relwidth=0.98, relheight=0.09)

        self.Entry_2 = Entry(self.frame_1)
        self.Entry_2.place(relx=0.01, rely=0.49, relwidth=0.98, relheight=0.09)

        self.Entry_name = Entry(self.frame_1)
        self.Entry_name.place(relx=0.01, rely=0.85, relwidth=0.98, relheight=0.09)

        # label e text de busca do caminho
        self.lb_path = Label(self.frame_1, text='Caminho do arquivo original', background='#dddddd')
        self.lb_path.place(relx=0.01, rely=0.005, relwidth=0.22, relheight=0.08)
        self.frame_text = Frame(self.frame_1)
        self.frame_text.place(relx=0.115, rely=0.25, relwidth=0.86, relheight=0.08)
        self.path_text = Text(self.frame_text, background='#dddddd')
        self.path_text.configure(state='disabled')
        self.path_text.pack(expand=True)

        #label do nome
        self.lb_name = Label(self.frame_1, text='Nome do arquivo', background='#dddddd')
        self.lb_name.place(relx=0.01, rely=0.74, relwidth=0.14, relheight=0.09)

        # label e text de busca do novo caminho
        self.lb_new_path = Label(self.frame_1, text='Local do novo arquivo', background='#dddddd')
        self.lb_new_path.place(relx=0.01, rely=0.38, relwidth=0.17, relheight=0.08)
        self.frame_new_path = Frame(self.frame_1)
        self.frame_new_path.place(relx=0.115, rely=0.62, relwidth=0.87, relheight=0.08)
        self.new_path_text = Text(self.frame_new_path, background='#dddddd')
        self.new_path_text.configure(state='disabled')
        self.new_path_text.pack(expand=True, fill='x')

    def widgets_frame2(self):
        # criando botão de word2pdf
        self.bt_word2pdf = Button(self.frame_2, text='Word para PDF', command=self.word_for_pdf)
        self.bt_word2pdf.place(relx=0.01, rely=0.02, relwidth=0.13, relheight=0.15)

        self.bt_excel2pdf = Button(self.frame_2, text='Excel para PDF', command=self.excel_for_pdf)
        self.bt_excel2pdf.place(relx=0.15, rely=0.02, relwidth=0.13, relheight=0.15)

# criando função de conversão
    def buscar_1(self):
        self.path_text.configure(state='normal')
        self.path_orig = tkinter.filedialog.askopenfile(mode='w', initialdir='/Documentos',
                                                        initialfile='C:/Users/rafae/OneDrive/Área de Trabalho'
                                                                    '/Memorial_descritivo.docx',
                                                        title="Selecione um arquivo",
                                                        filetypes=(
                                                            ("Arquivos doc", ".docx"), ("arquivos xlsx", ".xlsx")))
        a = self.path_orig.name.replace('/', '\\')
        self.path_text.delete(1.0, 'end')
        self.path_text.insert(1.0, a)
        self.path_text.configure(state='disabled')

    def buscar_2(self):
        self.new_path_text.configure(state='normal')
        self.new_path = tkinter.filedialog.askdirectory(initialdir='/Documentos', title="Selecione um arquivo")
        a = self.new_path.replace('/', '\\')
        self.new_path_text.delete(1.0, 'end')
        self.new_path_text.insert(1.0, a)
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

        self.entrada = self.path_text.get(1.0, 'end-1c')
        self.saida = self.new_path_text.get(1.0, 'end-1c') + r'\{}'.format(self.Entry_name.get())

        print(self.entrada)
        print(self.saida)

        word = win32com.client.Dispatch('Word.Application')
        doc = word.Documents.Open(self.entrada)
        doc.SaveAs(self.saida, FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()


Application()
