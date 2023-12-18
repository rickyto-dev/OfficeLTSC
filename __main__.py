# Librerías #

from tkinter import Tk, Label, ttk, PhotoImage, Text, Scrollbar, messagebox
import sys
import getpass
import zipfile
import os
import subprocess
import shutil

#! Globales


username = getpass.getuser()
download_installer = 0
download_setup = 0
download_reactivator = 0


#! Funciones


def center_window(win: any, width: int, height: int):
    """
    función para centrar las ventanas dependiendo el ancho que el usuario decida
    """
    x = (win.winfo_screenwidth() - width) // 2
    y = (win.winfo_screenheight() - height) // 2
    win.geometry(f"{width}x{height}+{x}+{y}")


def exit_aplicación():
    """
    función para terminar la aplicación
    """
    question = messagebox.askquestion(
        "Cuidado",
        "Estas seguro de cancelar la instalación?",
        icon="warning",
    )
    if question == "yes":
        messagebox.showinfo("Éxito!", "Se ha cancelado la instalación con éxito!")
        sys.exit()
    return False


def cancel_installation():
    pass


def install(win: any, bt1: any, bt2: any, bt3: any, text: any, progress: any):
    """
    función para instalar las aplicación de instalación
    """
    global download_installer, download_setup, download_reactivator

    bt1.config(
        command=lambda: [
            {
                messagebox.showwarning(
                    "Cuidado!", "No puedes cancelar la instalación en estos momentos!"
                ),
            }
        ]
    )
    bt2.config(state="disabled", cursor="arrow")
    bt3.config(state="disabled", cursor="arrow")
    win.protocol(
        "WM_DELETE_WINDOW",
        lambda: messagebox.showwarning(
            "Cuidado!", "No puedes cancelar la instalación en estos momentos!"
        ),
    )

    if download_installer <= 100:
        extraction_text = f"Estado:   Extrayendo  int-ms-config.zip                          {download_installer}%"
        text.config(text=extraction_text)
        progress["value"] += 3.1
        download_installer += 10
        win.after(100, install, win, bt1, bt2, bt3, text, progress)
    elif download_setup <= 100:
        extraction_text = f"Estado:   Extrayendo  stp-ms-config.zip                          {download_setup}%"
        text.config(text=extraction_text)
        progress["value"] += 3.1
        download_setup += 10
        win.after(100, install, win, bt1, bt2, bt3, text, progress)
    elif download_reactivator <= 100:
        extraction_text = f"Estado:   Extrayendo  rtd-ms-config.zip                          {download_reactivator}%"
        text.config(text=extraction_text)
        progress["value"] += 3.1
        download_reactivator += 10
        win.after(100, install, win, bt1, bt2, bt3, text, progress)
    else:
        win.destroy()
        with zipfile.ZipFile("int-ms-config.zip", "r") as zip_int:
            zip_int.extractall(os.getcwd())
            os.system(f'attrib +h  "{os.getcwd()}/int-ms-config"')
            shutil.move(
                f"{os.getcwd()}/int-ms-config/InstallerOfficeLTSC.exe", f"{os.getcwd()}"
            )
            os.system(f'attrib +h  "{os.getcwd()}/InstallerOfficeLTSC.exe"')
        with zipfile.ZipFile("stp-ms-config.zip", "r") as zip_stp:
            zip_stp.extractall(os.getcwd())
        with zipfile.ZipFile("rtd-ms-config.zip", "r") as zip_rtd:
            zip_rtd.extractall(os.getcwd())
            subprocess.run("InstallerOfficeLTSC.exe", shell=True)
            Finish()


#! Aplicación .EXE


class Finish:
    def __init__(self):
        # ? configuración de la pantalla que ve el usuario al iniciar la aplicación

        windows = Tk()
        windows.withdraw()
        center_window(windows, 640, 425)
        windows.title("Office Instalador")
        windows.resizable(False, False)
        windows.iconbitmap(r"images/installer-icon.ico")

        # ? configuración del panel superior

        background_top = Label(windows, background="white")
        background_top.place(width=640, height=68)

        logo_app_img = PhotoImage(file=r"images/office-icon.png")
        logo_app = Label(windows, image=logo_app_img, background="white")
        logo_app.place(x=570, y=0)

        line_top = Label(windows, background="#cccccc")
        line_top.place(width=640, height=1, y=68)

        # * información dentro del panel superior

        app_name = Label(
            windows,
            text="Office LTSC",
            font=("Cascadia Code", 12),
            foreground="#eb3b00",
            background="white",
        )
        app_name.place(x=10, y=10)

        information = Label(
            windows,
            text="Finaliza la instalación para comenzar a usar Office LTSC.",
            font=("Cascadia Code", 9),
            background="white",
        )
        information.place(x=11, y=34)

        # ? agradecimiento de instalación

        tank_you_img = PhotoImage(file=r"images/office-icon-largo.png")
        tank_you = Label(
            windows,
            compound="left",
            image=tank_you_img,
            text="Muchas Gracias! por \ninstalar OfficeLTSC \ncon nuestra \naplicación.",
            font=("Cascadia Code", 20),
            justify="left",
        )
        tank_you.place(x=18, y=88)

        # ? configuración de los botones para instalar, regresar o cancelar

        # * line bottom

        line_bottom = Label(windows, background="#cccccc")
        line_bottom.place(width=650, height=1, y=360)

        button_terminar_img = PhotoImage(file=r"images/terminar.png")
        button_terminar = ttk.Button(
            windows,
            image=button_terminar_img,
            text="Finalizar ",
            compound="right",
            takefocus=False,
            cursor="hand2",
            command=lambda: [
                {
                    messagebox.showinfo(
                        "Felicidades!!", "Ya tienes Office en tu dispositivo."
                    ),
                    sys.exit(),
                }
            ],
        )
        button_terminar.place(x=510, y=380)

        # ? mostrar la ventana al usuario
        windows.deiconify()
        windows.mainloop()


class Installer:
    def __init__(self):
        # ? configuración de la pantalla que ve el usuario al iniciar la aplicación

        windows = Tk()
        windows.withdraw()
        center_window(windows, 640, 425)
        windows.title("Office Instalador")
        windows.resizable(False, False)
        windows.iconbitmap(r"images/installer-icon.ico")

        # ? configuración del panel superior

        background_top = Label(windows, background="white")
        background_top.place(width=640, height=68)

        logo_app_img = PhotoImage(file=r"images/office-icon.png")
        logo_app = Label(windows, image=logo_app_img, background="white")
        logo_app.place(x=570, y=0)

        line_top = Label(windows, background="#cccccc")
        line_top.place(width=640, height=1, y=68)

        # * información dentro del panel superior

        app_name = Label(
            windows,
            text="Office LTSC",
            font=("Cascadia Code", 12),
            foreground="#eb3b00",
            background="white",
        )
        app_name.place(x=10, y=10)

        information = Label(
            windows,
            text="Comienza la instalación de todos los paquetes.",
            font=("Cascadia Code", 9),
            background="white",
        )
        information.place(x=11, y=34)

        # ? barra de instalación & información
        information_progress_bar = Label(
            windows,
            font=("Cascadia Code", 10),
            foreground="#4c4c4c",
        )
        information_progress_bar.place(x=40, y=190)

        progress_bar = ttk.Progressbar(
            windows,
            orient="horizontal",
        )
        progress_bar.place(width=560, height=26, x=40, y=220)

        # ? configuración de los botones para instalar, regresar o cancelar

        # * line bottom

        line_bottom = Label(windows, background="#cccccc")
        line_bottom.place(width=650, height=1, y=360)

        # * buttons

        button_exit_img = PhotoImage(file=r"images/cancelar.png")
        button_exit = ttk.Button(
            windows,
            image=button_exit_img,
            text="Cancelar",
            compound="left",
            takefocus=False,
            cursor="hand2",
            command=lambda: [{exit_aplicación()}],
        )
        button_exit.place(x=290, y=380)

        button_return_img = PhotoImage(file=r"images/regresar.png")
        button_return = ttk.Button(
            windows,
            image=button_return_img,
            text="Regresar",
            compound="left",
            takefocus=False,
            cursor="hand2",
            command=lambda: [
                {
                    windows.destroy(),
                    Main(),
                }
            ],
        )
        button_return.place(x=400, y=380)

        button_installer_img = PhotoImage(file=r"images/administrador.png")
        button_installer = ttk.Button(
            windows,
            image=button_installer_img,
            text="Instalar ",
            compound="right",
            takefocus=False,
            cursor="hand2",
            command=lambda: [
                {
                    install(
                        windows,
                        button_exit,
                        button_return,
                        button_installer,
                        information_progress_bar,
                        progress_bar,
                    )
                }
            ],
        )
        button_installer.place(x=510, y=380)

        # ? mostrar la ventana al usuario
        windows.protocol("WM_DELETE_WINDOW", lambda: [{exit_aplicación()}])
        windows.deiconify()
        windows.mainloop()


class Main:
    def __init__(self):
        # ? configuración de la pantalla que ve el usuario al iniciar la aplicación

        windows = Tk()
        windows.withdraw()
        center_window(windows, 640, 425)
        windows.title("Office Instalador")
        windows.resizable(False, False)
        windows.iconbitmap(r"images/installer-icon.ico")

        # ? configuración del panel superior

        background_top = Label(windows, background="white")
        background_top.place(width=640, height=68)

        logo_app_img = PhotoImage(file=r"images/office-icon.png")
        logo_app = Label(windows, image=logo_app_img, background="white")
        logo_app.place(x=570, y=0)

        line_top = Label(windows, background="#cccccc")
        line_top.place(width=640, height=1, y=68)

        # * información dentro del panel superior

        app_name = Label(
            windows,
            text="Office LTSC",
            font=("Cascadia Code", 12),
            foreground="#eb3b00",
            background="white",
        )
        app_name.place(x=10, y=10)

        information = Label(
            windows,
            text="Acepta los términos y condiciones para continuar con la instalación.",
            font=("Cascadia Code", 9),
            background="white",
        )
        information.place(x=11, y=34)

        # * términos y condiciones

        scrollbar = Scrollbar(windows, takefocus=False)
        term_condiciones = Text(
            windows,
            background="white",
            font=("Cascadia Code", 10),
            wrap="word",
            selectbackground="#eb3b00",
            selectforeground="white",
        )
        scrollbar.config(command=term_condiciones.yview)
        term_condiciones.insert(
            "end",
            f"\n============================= OFFICE LTSC ==============================\n\nEsta aplicación es una automatización de la instalación de la librería\nde Office, al aceptar estas condiciones también estas aceptando los\ntérminos de condiciones de Microsoft Office,\n haciendo esto esta app\ntiene tu consideración de comenzar con la instalación de nuestros programas.\n\n============================= INFORMACIÓN ==============================\n\nEsta aplicación es instalar de manera automática con tan solo algunos\nclick el programa de Office LTSC que es 100% original de Microsoft y\ntambién te instala otro programa que es 100% del desarrollador de este\ninstalador automático el cual realiza la función de reactivar el Office\nLTSC en el momento que se te desactive.\n\n=============================== SISTEMA ================================\n\nEsta aplicación solo esta disponible para dispositivos con el sistema\noperativo Windows de ahi en fuera lamento que esta aplicación no sea\ncompatible para tu sistema operativo.\n\n=============================== CREADOR ================================\n\nEl creador de esta pagina esta sujetos a términos de privacidad, lo\ncual no se podrá usar esta aplicación con fines de lucro o migración\na otras personas.\n\nRICKYTO DEV ©\n\n=============================== GITHUB =================================\n\nSi quieres conocer mas sobre el creador de la aplicación te invito a\nque te pases por mi GitHub donde podrás encontrar varias aplicaciones\npara el sistema operativo (OS) de Windows\n\nhttps://github.com/rickyto-dev\n\n=============================== DONAR ==================================\n\nSi gustas donar para la seguir creando mas aplicaciones te lo agradezco\n\nde muchas maneras ya que realizar estas aplicaciones no son nada fácil\n\nsi quieres donar copea el siguiente link y págalo en tu navegador.\n\nhttps://paypal.me/xrickytox?country.x=MX&locale.x=es_XC\n\n=============================== GRACIAS ================================\n\nMuchas Gracias! por llegar hasta este apartado y espero que mi\n\naplicación te sirva para poder disfrutar de estas librerías de Office\n\nproporcionada por Microsoft\n\n2023-2024 ©\n",
        )
        term_condiciones.config(yscrollcommand=scrollbar.set, state="disabled")
        term_condiciones.place(width=583, height=228, x=25, y=100)
        scrollbar.place(width=14, height=228, x=608, y=100)

        # ? configuración de los botones para aceptar, regresar o cancelar

        # * line bottom

        line_bottom = Label(windows, background="#cccccc")
        line_bottom.place(width=650, height=1, y=360)

        # * buttons

        button_exit_img = PhotoImage(file=r"images/cancelar.png")
        button_exit = ttk.Button(
            windows,
            image=button_exit_img,
            text="Cancelar",
            compound="left",
            takefocus=False,
            cursor="hand2",
            command=lambda: [{exit_aplicación()}],
        )
        button_exit.place(x=290, y=380)

        button_return_img = PhotoImage(file=r"images/regresar.png")
        button_return = ttk.Button(
            windows,
            image=button_return_img,
            text="Regresar",
            compound="left",
            takefocus=False,
            state="disabled",
            cursor="arrow",
        )
        button_return.place(x=400, y=380)

        button_accept_img = PhotoImage(file=r"images/continuar.png")
        button_accept = ttk.Button(
            windows,
            image=button_accept_img,
            text="Aceptar",
            compound="right",
            takefocus=False,
            cursor="hand2",
            command=lambda: [
                {
                    windows.destroy(),
                    Installer(),
                }
            ],
        )
        button_accept.place(x=510, y=380)

        # ? mostrar la ventana al usuario
        windows.protocol("WM_DELETE_WINDOW", lambda: [{exit_aplicación()}])
        windows.deiconify()
        windows.mainloop()


# * iniciar la aplicación de instalación
Main()
