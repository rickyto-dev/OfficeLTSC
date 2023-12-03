from tkinter import Tk, Label, ttk, PhotoImage, Text, Scrollbar, messagebox
import sys
import os
import shutil
import getpass
import zipfile
import subprocess
import traceback
import winshell
import time

# extraer el texto que esta en la LicenseOffice
with open("documents/LicenseOffice.txt", "r", encoding="utf-8") as _office:
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
            try:
                messagebox.showinfo(
                    "Instalación Correcta",
                    "Se ha instalado el office correctamente preciosa 'Enter' o 'Aceptar' para iniciar la instalación",
                )
                # > [ usuario de la computadora ]
                _usuario = getpass.getuser()
                # > [ rutas o paths de las ubicaciones a donde se moverán los archivos ]
                _office_path = f"C:/Program Files/Microsoft Office/Office16/"
                _exe = f"{_office_path}/office-reactivador/OfficeLTSC Reactivador.exe"
                _images = f"{_office_path}/office-reactivador/images"
                _documentos = f"{_office_path}/office-reactivador/documents"
                _app_path = f"{_office_path}/OfficeLTSC Reactivador.exe"
                _name_link = "OfficeLTSC Reactivador.lnk"
                _menu_escritorio = f"C:/Users/{_usuario}/OneDrive/Escritorio"
                _images_office = f"{_office_path}images"
                _documentos_office = f"{_office_path}documents"
                _office_arch_exe = "office-setup/setup.exe"
                _office_arch_xml = "office-setup/config.xml"
                _comando_install_office = "setup /configure config.xml"

                with zipfile.ZipFile("app/office-setup.zip", "r") as zip_ref_office:
                    zip_ref_office.extractall(os.getcwd())
                    shutil.move(_office_arch_exe, os.getcwd())
                    shutil.move(_office_arch_xml, os.getcwd())
                    shutil.rmtree(f"{os.getcwd()}/office-setup", ignore_errors=True)
                    win.destroy()
                    os.system(f'attrib +h "{os.getcwd()}/setup.exe"')
                    os.system(f'attrib +h "{os.getcwd()}/config.xml"')
                    time.sleep(1)
                    subprocess.run(_comando_install_office, shell=True)
                    subprocess.run(
                        ["taskkill", "/IM", "OfficeC2RClient.exe", "/F"], check=True
                    )
                    with zipfile.ZipFile(
                        "app/office-reactivador.zip", "r"
                    ) as zip_ref_office_reactivador:
                        zip_ref_office_reactivador.extractall(_office_path)
                        os.remove(f"setup.exe")
                        os.remove(f"config.xml")
                        time.sleep(0.2)
                        shutil.move(_exe, _office_path)
                        shutil.move(_images, _office_path)
                        shutil.move(_documentos, _office_path)
                        os.system(f'attrib +h "{_images_office}"')
                        os.system(f'attrib +h "{_documentos_office}"')
                        shutil.rmtree(
                            f"{_office_path}office-reactivador", ignore_errors=True
                        )
                    # crear el acceso directo en el escritorio
                    _ESCRITORIO = os.path.join(_menu_escritorio, _name_link)
                    _escritorio_comando = f'mklink "{_ESCRITORIO}" "{_app_path}" '
                    subprocess.run(_escritorio_comando, shell=True)
                    # crear el acceso directo en el buscador de aplicaciones
                    winshell.CreateShortcut(
                        os.path.join(
                            os.environ["ProgramData"],
                            "Microsoft",
                            "Windows",
                            "Start Menu",
                            "Programs",
                            _name_link,
                        ),
                        _app_path,
                    )

                Finish()
            except Exception as e:
                traceback_str = traceback.format_exc()
                messagebox.showerror(
                    "Error", f"Se ha producido un error:\n\n{traceback_str}"
                )

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
        window.iconbitmap("images/office.ico")
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
        _office_img = PhotoImage(file=r"images/logo.png")
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
        ## estilo botones
        estilos_bts = ttk.Style()
        estilos_bts.configure(
            "ButtonsStyles.TButton", background="white", bordercolor="#d0d0d0"
        )
        estilos_bts.map(
            "ButtonsStyles.TButton",
            background=[("active", "#e0eef9")],
            bordercolor=[("active", "#0078d4")],
        )
        ## botón finalizar
        _bt_finalizar_img = PhotoImage(file=r"images/aceptar.png")
        _bt_finalizar = ttk.Button(
            window,
            image=_bt_finalizar_img,
            compound="left",
            text="Finalizar",
            padding=(20, 15, 20, 15),
            cursor="hand2",
            takefocus=False,
            state="ButtonsStyles.TButton",
            command=lambda: [{window.destroy(), sys.exit()}],
        )
        _bt_finalizar.place(x=480, y=365)
        ## imagen de windows
        window_img = PhotoImage(file=r"images/windows.png")
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
        window.iconbitmap("images/office.ico")
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
        _office_img = PhotoImage(file=r"images/logo.png")
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
        estilo_progressbar = ttk.Style()
        estilo_progressbar.theme_use("clam")
        estilo_progressbar.configure(
            "TProgressbar",
            troughcolor="white",
            bordercolor="#eb3b00",
            background="#eb3b00",
            darkcolor="#eb3b00",
            lightcolor="#eb3b00",
        )
        _barra_progreso = ttk.Progressbar(window, takefocus=False, orient="horizontal")
        _barra_progreso.place(width=500, height=40, x=75, y=210)
        ## estilo botones
        estilos_bts = ttk.Style()
        estilos_bts.configure(
            "ButtonsStyles.TButton", background="white", bordercolor="#d0d0d0"
        )
        estilos_bts.map(
            "ButtonsStyles.TButton",
            background=[("active", "#e0eef9")],
            bordercolor=[("active", "#0078d4")],
        )
        ## botón regresar
        _bt_regresar_img = PhotoImage(file=r"images/regresar.png")
        _bt_regresar = ttk.Button(
            window,
            image=_bt_regresar_img,
            compound="left",
            text="Regresar",
            padding=(20, 15, 20, 15),
            cursor="hand2",
            takefocus=False,
            style="ButtonsStyles.TButton",
            command=lambda: [{window.destroy(), Main()}],
        )
        _bt_regresar.place(x=320, y=365)
        ## botón instalar
        _bt_instalar_img = PhotoImage(file=r"images/descargar.png")
        _bt_instalar = ttk.Button(
            window,
            image=_bt_instalar_img,
            compound="left",
            text="Instalar",
            padding=(20, 15, 20, 15),
            cursor="hand2",
            takefocus=False,
            style="ButtonsStyles.TButton",
            command=lambda: [
                {instalar(window, _barra_progreso, _bt_instalar, _bt_regresar)}
            ],
        )
        _bt_instalar.place(x=480, y=365)
        ## imagen de windows
        window_img = PhotoImage(file=r"images/windows.png")
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

    def __init__(self):
        # configuración de pagina principal
        window = Tk()
        window.withdraw()
        window.iconbitmap("images/office.ico")
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
        _office_img = PhotoImage(file=r"images/logo.png")
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
        ## estilo botones
        estilos_bts = ttk.Style()
        estilos_bts.configure(
            "ButtonsStyles.TButton", background="white", bordercolor="#d0d0d0"
        )
        estilos_bts.map(
            "ButtonsStyles.TButton",
            background=[("active", "#e0eef9")],
            bordercolor=[("active", "#0078d4")],
        )
        ## botón cancelar
        _bt_cancelar_img = PhotoImage(file=r"images/cancelar.png")
        _bt_cancelar = ttk.Button(
            window,
            image=_bt_cancelar_img,
            compound="left",
            text="Cancelar",
            padding=(20, 15, 20, 15),
            cursor="hand2",
            takefocus=False,
            style="ButtonsStyles.TButton",
            command=lambda: [{sys.exit()}],
        )
        _bt_cancelar.place(x=320, y=365)
        ## botón aceptar
        _bt_aceptar_img = PhotoImage(file=r"images/aceptar.png")
        _bt_aceptar = ttk.Button(
            window,
            image=_bt_aceptar_img,
            compound="left",
            text="Aceptar",
            padding=(20, 15, 20, 15),
            cursor="hand2",
            style="ButtonsStyles.TButton",
            takefocus=False,
            command=lambda: [{window.destroy(), Install()}],
        )
        _bt_aceptar.place(x=480, y=365)
        ## imagen de windows
        window_img = PhotoImage(file=r"images/windows.png")
        window_label = Label(window, image=window_img, background="white")
        window_label.place(x=15, y=400)
        # mostrar la ventana al usuario
        window.deiconify()
        window.mainloop()


if __name__ == "__main__":
    Main()
