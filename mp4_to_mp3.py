from moviepy.editor import VideoFileClip
import tkinter as tk
from tkinter import *
from tkinter import font, messagebox
from tkinter import filedialog
import time
from concurrent.futures import ThreadPoolExecutor

#---------------------------------------------------------------------------------------------------------------
#Para que se cierre automaticamente la consola del cmd
import win32gui, win32con

consola = win32gui.GetForegroundWindow()
win32gui.ShowWindow(consola , win32con.SW_HIDE)

#---------------------------------------------------------------------------------------------------------------

class app():

    def __init__(self):
        _bgcolor = '#d9d9d9'  # X11 color: 'gray85'
        _fgcolor = '#000000'  # X11 color: 'black'
        _compcolor = '#d9d9d9' # X11 color: 'gray85'
        _ana1color = '#d9d9d9' # X11 color: 'gray85'
        _ana2color = '#ececec' # Closest X11 color: 'gray92'
        font11 = "-family {Segoe UI} -size 11 -weight bold"
        font12 = "-family {Courier New} -size 12 -weight bold"
        font13 = "-family {Segoe UI} -size 14 -weight bold"

        self.archivos = []
        self.mostrar_en_listbox = ()
        self.DIR = 'C:/Users/FRANK BOJORQUEZ/Downloads'
        self.ejecutor = ThreadPoolExecutor(max_workers=2)
        self.convirtiendo = False
        self.f_mp4 = ''
        self.f_mp3 = ''
        self.archivo_valido = False

        self.root = tk.Tk()
        self.root.geometry("600x462")
        self.root.resizable(0, 0)
        self.root.title("MP3 Conversor por Frank Bojorque")
        self.root.iconbitmap('musica.ico')
        self.root.configure(background="#330066")

        self.Listbox1 = tk.Listbox(self.root)
        self.Listbox1.place(relx=0.100, rely=0.043, relheight=0.556,relwidth=0.800)
        self.Listbox1.configure(background="#ffffe1")
        self.Listbox1.configure(disabledforeground="#a3a3a3")
        self.Listbox1.configure(font=font12)
        self.Listbox1.configure(foreground="#000000")
        self.Listbox1.configure(justify='left')

        self.B1 = tk.Button(self.root, command=self.agrear_archivos)
        self.B1.place(relx=0.233, rely=0.645, height=50, width=87)
        self.B1.configure(activebackground="#ececec")
        self.B1.configure(activeforeground="#000000")
        self.B1.configure(background="#ffff91")
        self.B1.configure(borderwidth="8")
        self.B1.configure(disabledforeground="#a3a3a3")
        self.B1.configure(font=font11)
        self.B1.configure(foreground="#000000")
        self.B1.configure(highlightbackground="#d9d9d9")
        self.B1.configure(highlightcolor="black")
        self.B1.configure(pady="0")
        self.B1.configure(text='''Agregar''')

        self.B2 = tk.Button(self.root, command=self.inicia_convertir)
        self.B2.place(relx=0.433, rely=0.645, height=50, width=87)
        self.B2.configure(activebackground="#ececec")
        self.B2.configure(activeforeground="#000000")
        self.B2.configure(background="#ffff91")
        self.B2.configure(borderwidth="8")
        self.B2.configure(disabledforeground="#a3a3a3")
        self.B2.configure(font=font11)
        self.B2.configure(foreground="#000000")
        self.B2.configure(highlightbackground="#d9d9d9")
        self.B2.configure(highlightcolor="black")
        self.B2.configure(pady="0")
        self.B2.configure(text='''Convertir''')

        self.B3 = tk.Button(self.root, command=self.eliminar_de_listbox)
        self.B3.place(relx=0.633, rely=0.645, height=50, width=87)
        self.B3.configure(activebackground="#ececec")
        self.B3.configure(activeforeground="#000000")
        self.B3.configure(background="#ffff91")
        self.B3.configure(borderwidth="8")
        self.B3.configure(disabledforeground="#a3a3a3")
        self.B3.configure(font=font11)
        self.B3.configure(foreground="#000000")
        self.B3.configure(highlightbackground="#d9d9d9")
        self.B3.configure(highlightcolor="black")
        self.B3.configure(pady="0")
        self.B3.configure(text='''Eliminar''')

        self.Label1 = tk.Label(self.root)
        self.Label1.place(relx=0.147, rely=0.801, height=80, width=450)
        self.Label1.configure(background="#330066")
        self.Label1.configure(disabledforeground="#a3a3a3")
        self.Label1.configure(font=font13)
        self.Label1.configure(foreground="#ffffff")    

        self.root.mainloop()


    def agrear_archivos(self):
        
        self.mostrar_en_listbox = filedialog.askopenfilenames(title='Seleccionar los archivos a convertir',
                                                initialdir=self.DIR,
                                                filetypes=[('Archivos MP4','*.mp4'),
                                                            ('Archivos AVI','*.avi'),
                                                            ('Archivos MKV','*.mkv'),
                                                            ('Archicos FLV','*.flv'),
                                                            ('Archivos MOV','*.mov'),
                                                            ('Archivos WMV','*.wmv')])
        if self.archivos == '':
            messagebox.showwarning('Advertencia', 'No se seleccionó ningún archivo...')
        else:
            for self.n_archivo in self.mostrar_en_listbox:
                self.archivos.append(self.n_archivo)
                self.n_archivo = self.n_archivo.split('/')[-1]
                self.Listbox1.insert(END, self.n_archivo)

    def eliminar_de_listbox(self):
        
        try:

            self.seleccion = self.Listbox1.curselection()[0]#Obtenemos el indice del elemento en el listbox
            self.seleccion = self.Listbox1.get(self.seleccion)#Obtenemos el valor del elemento partiendo de su indice dentro del listbox
            
            self.indice = [self.archivos.index(elemento) for elemento in self.archivos if self.seleccion in elemento]
            self.indice = int(self.indice[0]) #Lo convertimos a entero por que con la linea de arriba queda como una lista
    
            del self.archivos[self.indice]#Eliminarmos el elemento de la lista para no tomarlo en cuenta en la conversion
            
            self.Listbox1.delete(ANCHOR)#Eliminamos el elemento del listbox

        except:
            pass

    def informa_proceso(self):

        Helvfont = font.Font(family="Helvetica", size=13, weight="bold")

        while self.convirtiendo == False:
            self.Label1.config(text='CONVERSION EN PROCESO POR FAVOR ESPERA', font=Helvfont)
            time.sleep(0.9)
            self.Label1.config(text='')
            time.sleep(0.9)

        self.convirtiendo = False

    def convertir_a_mp3(self):

        try:
            for self.f_mp4 in self.archivos:
                
                self.extension = ''
                
                for c in self.f_mp4[::-1]:
                    self.extension += c
                    if c == '.':
                        break 

                self.extension = self.extension[::-1]

                self.f_mp3 = self.f_mp4.replace(self.extension, '.mp3')

                self.video_clip = VideoFileClip(self.f_mp4)
                self.audio_clip = self.video_clip.audio

                self.audio_clip.write_audiofile(self.f_mp3)

                self.audio_clip.close()
                self.video_clip.close()

            self.convirtiendo = True
            messagebox.showinfo('Aviso','CONVERSION EXITOSA!!!')
            self.archivos = []
            self.Listbox1.delete(0, tk.END)
        except:
            messagebox.showerror('Error','Debes seleccionar un archivo de audio')

    def inicia_convertir(self):

        if not self.archivos or self.archivos == '':
            messagebox.showwarning('Advertencia', 'No se seleccionó ningún archivo...')
        
        else:
            self.ejecutor.submit(self.convertir_a_mp3)
            self.ejecutor.submit(self.informa_proceso)
        
if __name__ == "__main__":
    app()

