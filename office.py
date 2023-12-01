from tkinter import Tk, Label, ttk, PhotoImage, Text, Scrollbar, messagebox
import sys
import os
import shutil
import getpass


# extraer el texto que esta en la LicenseOffice
with open("ライセンス/LicenseOffice.txt", "r", encoding="utf-8") as _office:
    LicenseOffice = _office.read()


# funciones
def center_window(win: any, ancho: int, alto: int):
    """
    Función para centrar todas la ventas
    """
    #  win     =>   ! es la ventana que usaremos para la configuración y poder cambiarle la posición
    #  ancho   =>   ! es el ancho de la ventana
    #  alto    =>   ! es el alto de la ventana
    x = (win.winfo_screenwidth() - ancho) // 2
    y = (win.winfo_screenheight() - alto) // 2
    win.geometry(f"{ancho}x{alto}+{x}+{y}")


def instalar(win: any, progress: any, bt_inst: any, bt_reg: any):
    """
    Función para instalar todos los programas
    """
    #  win        =>   ! es la ventana que se esta mostrando
    #  progress   =>   ! es la barra de progreso
    #  bt_inst    =>   ! es el botón para instalar la aplicación
    #  bt_reg     =>   ! es el botón para regresar al panel de términos y condiciones
    bt_inst["state"] = "disabled"
    bt_reg["state"] = "disabled"
    bt_inst["cursor"] = "arrow"
    bt_reg["cursor"] = "arrow"

    def progress_chargee():
        valor = progress["value"]
        if valor < 100:
            progress["value"] += 20
            progress.update()
            win.after(1000, progress_chargee)
        else:
            # _pack_images = f"{os.getcwd()}/画像"
            # _pack_files = f"{os.getcwd()}/ライセンス"
            # _pack_reactivador = f"{os.getcwd()}/書類/アクティベータ.zip"
            # _usuario = getpass.getuser()
            # _escritorio = f"C:/Users/{_usuario}/OneDrive/Escritorio"
            # shutil.move(_pack_images, _escritorio)
            # shutil.move(_pack_files, _escritorio)
            # shutil.copy(_pack_reactivador, _escritorio)
            win.destroy()
            Finish()

    win.protocol(
        "WM_DELETE_WINDOW",
        lambda: [
            {
                messagebox.showwarning(
                    "Peligro de Instalación",
                    "No puedes cerrar esta ventana mientras se esta instalando el programa",
                )
            }
        ],
    )
    progress_chargee()


def desempaquetado():
    pass


# aplicación
class Finish:
    """
    Inicio de aplicación donde se aceptaran los términos y condiciones
    """

    def __init__(self):
        # configuración de la pagina principal
        window = Tk()
        window.withdraw()
        window.iconbitmap("画像/office.ico")
        window.title("Office LTSC  |  Instalador")
        window.resizable(False, False)
        window.config(background="white")
        center_window(window, 650, 450)
        # información
        ## panel superior
        _linea_superior = Label(
            window,
            background="#b5b5b5",
        )
        _linea_superior.place(width=650, height=1, y=70)
        ### > logo del panel superior
        _office_img = PhotoImage(file=r"画像/logo.png")
        _office = Label(window, image=_office_img, borderwidth=0, background="white")
        _office.place(x=5, y=3)
        ### > información
        _info_ltsc = Label(
            window,
            text=" | Office LTSC - Finalizar",
            font=("Cascadia Code", 10),
            background="white",
            foreground="#eb3b00",
        )
        _info_ltsc.place(x=60, y=12)
        _info_panel = Label(
            window,
            text=" | Panel de agradecimiento por instalar la aplicación",
            font=("Cascadia Code", 10),
            background="white",
            foreground="#848486",
        )
        _info_panel.place(x=60, y=34)
        ## información
        _info = Label(
            window,
            text="Muchas Gracias Por Instalar",
            font=("Cascadia Code", 22),
            background="white",
            foreground="#eb3b00",
            justify="center",
        )
        _info.place(width=650, y=180)
        _info = Label(
            window,
            text="Office LTSC",
            font=("Cascadia Code", 22),
            background="white",
            foreground="#848486",
            justify="center",
        )
        _info.place(width=650, y=220)
        ## botón finalizar
        _bt_finalizar_img = PhotoImage(file=r"画像/aceptar.png")
        _bt_finalizar = ttk.Button(
            window,
            image=_bt_finalizar_img,
            compound="left",
            text="Finalizar",
            padding=(20, 15, 20, 15),
            cursor="hand2",
            takefocus=False,
            command=lambda: [{window.destroy(), sys.exit()}],
        )
        _bt_finalizar.place(x=480, y=365)
        ## imagen de windows
        window_img = PhotoImage(file=r"画像/windows.png")
        window_label = Label(window, image=window_img, background="white")
        window_label.place(x=15, y=400)
        # mostrar la ventana al usuario
        window.protocol("WM_DELETE_WINDOW", lambda: [{window.destroy(), sys.exit()}])
        window.deiconify()
        window.mainloop()


class Install:
    """
    Inicio de aplicación donde se aceptaran los términos y condiciones
    """

    def __init__(self):
        # configuración de la pagina principal
        window = Tk()
        window.withdraw()
        window.iconbitmap("画像/office.ico")
        window.title("Office LTSC  |  Instalador")
        window.resizable(False, False)
        window.config(background="white")
        center_window(window, 650, 450)
        # información
        ## panel superior
        _linea_superior = Label(
            window,
            background="#b5b5b5",
        )
        _linea_superior.place(width=650, height=1, y=70)
        ### > logo del panel superior
        _office_img = PhotoImage(file=r"画像/logo.png")
        _office = Label(window, image=_office_img, borderwidth=0, background="white")
        _office.place(x=5, y=3)
        ### > información
        _info_ltsc = Label(
            window,
            text=" | Office LTSC - Instalación",
            font=("Cascadia Code", 10),
            background="white",
            foreground="#eb3b00",
        )
        _info_ltsc.place(x=60, y=12)
        _info_panel = Label(
            window,
            text=" | Panel de instalación de Office LTSC",
            font=("Cascadia Code", 10),
            background="white",
            foreground="#848486",
        )
        _info_panel.place(x=60, y=34)
        ## barra de progreso
        _barra_progreso = ttk.Progressbar(window, takefocus=False, orient="horizontal")
        _barra_progreso.place(width=500, height=40, x=75, y=210)
        ## botón regresar
        _bt_regresar_img = PhotoImage(file=r"画像/regresar.png")
        _bt_regresar = ttk.Button(
            window,
            image=_bt_regresar_img,
            compound="left",
            text="Regresar",
            padding=(20, 15, 20, 15),
            cursor="hand2",
            takefocus=False,
            command=lambda: [{window.destroy(), Main()}],
        )
        _bt_regresar.place(x=320, y=365)
        ## botón instalar
        _bt_instalar_img = PhotoImage(file=r"画像/descargar.png")
        _bt_instalar = ttk.Button(
            window,
            image=_bt_instalar_img,
            compound="left",
            text="Instalar",
            padding=(20, 15, 20, 15),
            cursor="hand2",
            takefocus=False,
            command=lambda: [
                {instalar(window, _barra_progreso, _bt_instalar, _bt_regresar)}
            ],
        )
        _bt_instalar.place(x=480, y=365)
        ## imagen de windows
        window_img = PhotoImage(file=r"画像/windows.png")
        window_label = Label(window, image=window_img, background="white")
        window_label.place(x=15, y=400)
        # mostrar la ventana al usuario
        window.protocol("WM_DELETE_WINDOW", lambda: [{window.destroy(), Main()}])
        window.deiconify()
        window.mainloop()


class Main:
    """
    Inicio de aplicación donde se aceptaran los términos y condiciones
    """

    os.system(f"attrib +h {os.getcwd()}/書類")
    os.system(f"attrib +h {os.getcwd()}/ライセンス")

    def __init__(self):
        # configuración de la pagina principal
        # os.system('')
        window = Tk()
        window.withdraw()
        window.iconbitmap("画像/office.ico")
        window.title("Office LTSC  |  Instalador")
        window.resizable(False, False)
        window.config(background="white")
        center_window(window, 650, 450)
        # información
        ## panel superior
        _linea_superior = Label(
            window,
            background="#b5b5b5",
        )
        _linea_superior.place(width=650, height=1, y=70)
        ### > logo del panel superior
        _office_img = PhotoImage(file=r"画像/logo.png")
        _office = Label(window, image=_office_img, borderwidth=0, background="white")
        _office.place(x=5, y=3)
        ### > información
        _info_ltsc = Label(
            window,
            text=" | Office LTSC - Términos y Condiciones",
            font=("Cascadia Code", 10),
            background="white",
            foreground="#eb3b00",
        )
        _info_ltsc.place(x=60, y=12)
        _info_panel = Label(
            window,
            text=" | Términos & Condiciones que aceptaras al instalar la aplicación",
            font=("Cascadia Code", 10),
            background="white",
            foreground="#848486",
        )
        _info_panel.place(x=60, y=34)
        ### > texto
        _scrollbar = Scrollbar(window, takefocus=False)
        _term_condiciones = Text(
            window,
            background="#f0f0f0",
            foreground="black",
            font=("Cascadia Code", 10),
            borderwidth=0,
            wrap="word",
            selectbackground="#eb3b00",
            selectforeground="white",
        )
        _scrollbar.config(command=_term_condiciones.yview)
        _term_condiciones.insert("end", f"{LicenseOffice}")
        _term_condiciones.config(yscrollcommand=_scrollbar.set, state="disabled")
        _term_condiciones.place(width=576, height=250, x=25, y=90)
        _scrollbar.place(width=18, height=250, x=600, y=90)
        ## botón cancelar
        _bt_cancelar_img = PhotoImage(file=r"画像/cancelar.png")
        _bt_cancelar = ttk.Button(
            window,
            image=_bt_cancelar_img,
            compound="left",
            text="Cancelar",
            padding=(20, 15, 20, 15),
            cursor="hand2",
            takefocus=False,
            command=lambda: [{sys.exit()}],
        )
        _bt_cancelar.place(x=320, y=365)
        ## botón aceptar
        _bt_aceptar_img = PhotoImage(file=r"画像/aceptar.png")
        _bt_aceptar = ttk.Button(
            window,
            image=_bt_aceptar_img,
            compound="left",
            text="Aceptar",
            padding=(20, 15, 20, 15),
            cursor="hand2",
            takefocus=False,
            command=lambda: [{window.destroy(), Install()}],
        )
        _bt_aceptar.place(x=480, y=365)
        ## imagen de windows
        window_img = PhotoImage(file=r"画像/windows.png")
        window_label = Label(window, image=window_img, background="white")
        window_label.place(x=15, y=400)
        # mostrar la ventana al usuario
        window.deiconify()
        window.mainloop()


if __name__ == "__main__":
    Main()
