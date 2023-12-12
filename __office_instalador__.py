# PAQUETES QUE VAMOS A UTILIZAR
from tkinter import Tk, PhotoImage, Label, Scrollbar, Text, ttk, messagebox
import urllib.request
import sys, os, shutil, getpass, zipfile, subprocess, winshell, time, socket

# GLOBALES
descarga = 0

# * IMÁGENES QUE ESTÁN EN LA NUBE * #


# # LINKS
class Imagenes:
    def __init__(self):
        self.logo_y_icon_web = urllib.request.urlopen(
            "https://i.imgur.com/gjzczCA.png"
        ).read()
        self.logo_aceptar = urllib.request.urlopen(
            "https://i.imgur.com/ugALV6A.png"
        ).read()
        self.logo_cancelar = urllib.request.urlopen(
            "https://i.imgur.com/uGbO1DH.png"
        ).read()
        self.logo_regresar = urllib.request.urlopen(
            "https://i.imgur.com/GlRjUmV.png"
        ).read()
        self.logo_descargar = urllib.request.urlopen(
            "https://i.imgur.com/nspcR2J.png"
        ).read()


# FUNCIONES & ACCIONES QUE REALIZARA LA APLICACIÓN
def center_window(win: any, ancho: int, alto: int):
    """
    Función para centrar todas la ventas
    """

    #  win     =>   ! ES LA VENTANA QUE CAMBIAREMOS DE POSICIÓN
    #  ancho   =>   ! ES EL ANCHO DE LA VENTANA
    #  alto    =>   ! ES EL ALTO DE LA VENTANA

    x = (win.winfo_screenwidth() - ancho) // 2
    y = (win.winfo_screenheight() - alto) // 2
    win.geometry(f"{ancho}x{alto}+{x}+{y}")


def exit_app():
    """
    Muestra una ventana donde se le pregunta al usuario si quiere terminar la instalación
    """

    mensaje = messagebox.askquestion(
        "Cuidado la Instalación a sido Interrumpida",
        "Estas seguro que quieres salir de la instalación ??",
        # MENSAJE QUE SE MOSTRARA AL USUARIO
        icon="warning",
    )
    if mensaje == "yes":
        # SI SE OBTIENE QUE EL USUARIO PRESIONO EL BOTÓN QUE SI QUISO SALIR, SE CERRARA TODA LA APLICACIÓN
        sys.exit()

    # SI EL USUARIO PRESIONO QUE NO SE REGRESARA A LA VENTA EN DONDE ESTABA
    else:
        return False


def instalar(win: any, progress: any, bt_inst: any, bt_reg: any, text_install: int):
    """
    Función para instalar todos los programas
    """

    #  win        =>   ! ES LA VENTANA QUE SE MUESTRA AL USUARIO
    #  progress   =>   ! ES LA BARRA DE PROGRESO
    #  bt_inst    =>   ! ES EL BOTÓN PARA INSTALAR LA APLICACIÓN
    #  bt_reg     =>   ! ES EL BOTÓN PARA REGRESAR A LA PAGINA PRINCIPAL

    # GLOBALES
    bt_inst["state"] = "disabled"
    bt_reg["state"] = "disabled"
    bt_inst["cursor"] = "arrow"
    bt_reg["cursor"] = "arrow"

    # DESACTIVAR EL CERRADO DE LA APLICACIÓN
    win.protocol(
        "WM_DELETE_WINDOW",
        lambda: [
            {
                messagebox.showwarning(
                    "Cuidado",
                    "No se puede cancelar la inhalación por favor espere mientras termina la instalación",
                )
            }
        ],
    )

    def progress_chargee():
        """
        Esta función lo que hace es mover la barra de progreso añadiendo nuevos parámetros
        """

        # GLOBALES
        global descarga

        # OBTIENE EL VALOR DE LA BARRA DE PROGRESO
        valor = progress["value"]

        if valor < 100:
            # AGREGAR MAS 20 AL VALOR ACTUAL DE LA BARRA DE PROGRESO
            progress["value"] += 20

            # CAMBIA LA CONFIGURACIÓN DEL ESTADO DEL PORCENTAJE DE INSTALACIÓN
            descarga += 20
            text_install.place(x=80, y=195)
            text_install.config(
                text=f"Instalando                                              {descarga}%"
            )

            # ACTUALIZA LA BARRA DE PROGRESO EN EL MOMENTO REAL
            progress.update()
            win.after(1000, progress_chargee)

        else:
            # DESTRUIR LA VENTANA EN LA QUE SE ESTA EJECUTANDO LA INSTALACIÓN
            win.destroy()

            # DESCOMPRIMIR EL INSTALADOR DE OFFICE
            with zipfile.ZipFile("stp-installer-manager.zip", "r") as setup:
                setup.extractall(os.getcwd())

            # OCULTAR EL ARCHIVO EXTRAÍDO
            os.system(f'attrib +h "{os.getcwd()}/office-setup"')

            # MOVER LOS ARCHIVOS DE LA CARPETA EXTRAÍDA A LA CARPETA EN LA QUE ESTAMOS
            shutil.move(f"{os.getcwd()}/office-setup/setup.exe", os.getcwd())
            shutil.move(f"{os.getcwd()}/office-setup/config.xml", os.getcwd())

            # OCULTAR LOS ARCHIVOS MOVIDOS
            os.system(f'attrib +h "{os.getcwd()}/setup.exe"')
            os.system(f'attrib +h "{os.getcwd()}/config.xml"')

            # BORRAR LA CARPETA EXTRAÍDA
            time.sleep(0.5)
            shutil.rmtree(f"{os.getcwd()}/office-setup")

            # INICIAR LA INSTALACIÓN DE OFFICE POR MEDIO DE COMANDOS
            ## COMANDOS ##
            _comando = "setup /configure config.xml"

            ## EJECUTAR PROGRAMA POR CONSOLA ##
            subprocess.run(_comando, shell=True)

            ## TERMINAR EL PROCESO DE LA APLICACIÓN ##
            subprocess.run(["taskkill", "/IM", "OfficeC2RClient.exe", "/F"], check=True)

            ## ELIMINAR TODOS LOS ARCHIVOS QUE USAMOS PARA LA INSTALACIÓN ##
            os.remove(f"setup.exe")
            os.remove(f"config.xml")

            # EXTRAER EL EJECUTABLE DE REACTIVADOR
            with zipfile.ZipFile("rct-installer-manager.zip", "r") as reactivador:
                reactivador.extractall("C:/Program Files/Microsoft Office/Office16/")

            # MOVER EL CONTENIDO QUE ESTA DENTRO DE LA LA CARPETA DENTRO DE OFFICE 16
            shutil.move(
                "C:/Program Files/Microsoft Office/Office16/office-reactivador/OfficeLTSC-Reactivador.exe",
                "C:/Program Files/Microsoft Office/Office16/",
            )

            # REMOVER LA CARPETA DONDE SE ENCONTRABA EL PROGRAMA
            shutil.rmtree(
                "C:/Program Files/Microsoft Office/Office16/office-reactivador"
            )

            # OBTENER EL NOMBRE DE LA COMPUTADORA
            _usuario = getpass.getuser()

            # CONFIGURAR EL NOMBRE DEL ACCESO DIRECTO
            _nombre_app = "OfficeLTSC-Reactivador.lnk"

            ## ESCRITORIO
            _menu_escritorio = f"C:/Users/{_usuario}/OneDrive/Escritorio"
            _ESCRITORIO = os.path.join(_menu_escritorio, _nombre_app)
            _escritorio_comando = f'mklink "{_ESCRITORIO}" f"C:/Users/{_usuario}/AppData/Roaming/Microsoft/Windows/Start Menu/Programs/OfficeLTSC-Reactivador.exe" '
            subprocess.run(_escritorio_comando, shell=True)

            ## MENU APLICACIONES
            winshell.CreateShortcut(
                os.path.join(
                    os.environ["ProgramData"],
                    "Microsoft",
                    "Windows",
                    "Start Menu",
                    "Programs",
                    _nombre_app,
                ),
                "C:/Program Files/Microsoft Office/Office16/OfficeLTSC-Reactivador.exe",
            )

            # INFORMAR AL USUARIO QUE LA APLICACIÓN DE REACTIVADOR SOLO SE PUEDE USAR SI TIENE INTERNET
            messagebox.showinfo(
                "Información Reactivador",
                "El programa de reactivador solo funciona si tienes internet ya que hace peticiones, de lo contrario el programa te arrogara un error al momento de querer abrirlo",
            )

            Finish()

    # INICIAR LA BARRA DE PROGRESO
    progress_chargee()


# * APLICACIÓN * #
class Finish:
    """
    Este es el apartado de finalización de la aplicación
    """

    def __init__(self):
        # GLOBAL
        imagenes = Imagenes()

        ## * CONFIGURACIÓN DE LA VENTANA QUE VERA EL USUARIO * ##
        window = Tk()

        # ESTO HACE QUE LA VENTANA SE OCULTE AL USUARIO AL MOMENTO DE QUE SE INICIALICE
        window.withdraw()

        # TITULO DE LA VENTANA DE LA APLICACIÓN
        window.title("Office LTSC")

        # AQUÍ ESTA LLAMANDO A LA FUNCIÓN (center_window) PARA CENTRAR LA VENTANA EN MEDIO DE TODA LA PANTALLA DEL USUARIO & LE ESTAMOS PASANDO ALGUNOS VALORES
        center_window(window, 650, 460)

        # ICONO DE LA APLICACIÓN
        icon = PhotoImage(data=imagenes.logo_y_icon_web)
        window.iconphoto(False, icon)
        window.config(background="white")

        # EN ESTA PARTE ES PARA QUE EL USUARIO NO PUEDA CAMBIAR EL TAMAÑO DE LA VENTANA
        window.resizable(0, 0)

        ## * PANEL SUPERIOR  * ##

        # SE DIBUJARA UNA LINEA DIVISORA PARA DIVIDIR LA INFORMACIÓN
        linea_superior = Label(
            window,
            background="#647cdc",
        )
        linea_superior.place(width=650, height=1, y=70)

        # SE MUESTRA EL LOGO DE LA APLICACIÓN PARA QUE SEA MAS ENTENDIBLE PARA EL USUARIO
        office_img = PhotoImage(data=imagenes.logo_y_icon_web)
        office = Label(window, image=office_img, borderwidth=0, background="white")
        office.place(x=5, y=3)

        # SE MUESTRA INFORMACIÓN ALADO DEL LOGO PARA QUE EL USUARIO ENTIENDA EN QUE APARTADO DE ENCUENTRA
        info_ltsc = Label(
            window,
            text=" | Office LTSC - Finalizar",
            font=("Cascadia Code", 10),
            background="white",
            foreground="#647cdc",
        )
        info_ltsc.place(x=60, y=12)
        info_pagina = Label(
            window,
            text=" | Finalizamos la instalación con éxito!",
            font=("Cascadia Code", 10),
            background="white",
            foreground="#8560b2",
        )
        info_pagina.place(x=60, y=34)

        ## * TEXTO DE AGRADECIMIENTO * ##

        info = Label(
            window,
            text="Muchas Gracias Por Instalar",
            font=("Cascadia Code", 22),
            background="white",
            foreground="#647cdc",
            justify="center",
        )
        info.place(width=650, y=180)
        info = Label(
            window,
            text="Office LTSC",
            font=("Cascadia Code", 22),
            background="white",
            foreground="#8560b2",
            justify="center",
        )
        info.place(width=650, y=230)

        ## * BOTONES * ##

        # ESTILOS DEL BOTÓN FINAL
        estilos_bts = ttk.Style()
        estilos_bts.theme_use("clam")
        estilos_bts.configure(
            "ButtonsStyles.TButton", background="white", bordercolor="#647cdc"
        )
        estilos_bts.map(
            "ButtonsStyles.TButton",
            background=[("active", "#e0eef9")],
            bordercolor=[("active", "#647cdc")],
        )

        # BOTÓN PARA FINALIZAR
        bt_finalizar_img = PhotoImage(data=imagenes.logo_aceptar)
        bt_finalizar = ttk.Button(
            window,
            image=bt_finalizar_img,
            compound="left",
            text="Finalizar",
            padding=(20, 15, 20, 15),
            cursor="hand2",
            takefocus=False,
            style="ButtonsStyles.TButton",
            command=lambda: [{window.destroy(), sys.exit()}],
        )
        bt_finalizar.place(x=480, y=377)

        # * MOSTRAR LA VENTANA CON TODA LA CONFIGURACIÓN * #
        window.protocol("WM_DELETE_WINDOW", lambda: [{window.destroy(), sys.exit()}])
        window.deiconify()
        window.mainloop()


class Install:
    """
    Este es el apartado de instalación
    """

    def __init__(self):
        # GLOBAL
        imagenes = Imagenes()

        ## * CONFIGURACIÓN DE LA VENTANA QUE VERA EL USUARIO * ##
        window = Tk()

        # ESTO HACE QUE LA VENTANA SE OCULTE AL USUARIO AL MOMENTO DE QUE SE INICIALICE
        window.withdraw()

        # TITULO DE LA VENTANA DE LA APLICACIÓN
        window.title("Office LTSC")

        # AQUÍ ESTA LLAMANDO A LA FUNCIÓN (center_window) PARA CENTRAR LA VENTANA EN MEDIO DE TODA LA PANTALLA DEL USUARIO & LE ESTAMOS PASANDO ALGUNOS VALORES
        center_window(window, 650, 460)

        # ICONO DE LA APLICACIÓN
        icon = PhotoImage(data=imagenes.logo_y_icon_web)
        window.iconphoto(False, icon)
        window.config(background="white")

        # EN ESTA PARTE ES PARA QUE EL USUARIO NO PUEDA CAMBIAR EL TAMAÑO DE LA VENTANA
        window.resizable(0, 0)

        ## * PANEL SUPERIOR  * ##

        # SE DIBUJARA UNA LINEA DIVISORA PARA DIVIDIR LA INFORMACIÓN
        linea_superior = Label(
            window,
            background="#647cdc",
        )
        linea_superior.place(width=650, height=1, y=70)

        # SE MUESTRA EL LOGO DE LA APLICACIÓN PARA QUE SEA MAS ENTENDIBLE PARA EL USUARIO
        office_img = PhotoImage(data=imagenes.logo_y_icon_web)
        office = Label(window, image=office_img, borderwidth=0, background="white")
        office.place(x=5, y=3)

        # SE MUESTRA INFORMACIÓN ALADO DEL LOGO PARA QUE EL USUARIO ENTIENDA EN QUE APARTADO DE ENCUENTRA
        info_ltsc = Label(
            window,
            text=" | Office LTSC - Instalación",
            font=("Cascadia Code", 10),
            background="white",
            foreground="#647cdc",
        )
        info_ltsc.place(x=60, y=12)
        info_pagina = Label(
            window,
            text=" | Panel de instalación de Office LTSC",
            font=("Cascadia Code", 10),
            background="white",
            foreground="#8560b2",
        )
        info_pagina.place(x=60, y=34)

        ## * TEXTO DE INSTALACIÓN & PORCENTAJES DE INSTALACIÓN * ##

        # TEXTO DE INSTALACIÓN
        text_install = Label(
            window,
            font=("Cascadia Code", 10),
            background="white",
            foreground="#8560b2",
        )

        ## *  BARRA DE PROGRESO  * ##

        # CONFIGURACIÓN DE LA BARRA DE PROGRESO DONDE SE VERA EL PROGRESO DE INSTALACIÓN
        estilo_progressbar = ttk.Style()
        estilo_progressbar.theme_use("clam")
        estilo_progressbar.configure(
            "TProgressbar",
            troughcolor="white",
            bordercolor="#647cdc",
            background="#647cdc",
            darkcolor="#647cdc",
            lightcolor="#647cdc",
        )
        barra_progreso = ttk.Progressbar(window, takefocus=False, orient="horizontal")
        barra_progreso.place(width=500, height=40, x=75, y=220)

        ## *  BOTONES  * ##

        # CONFIGURACIÓN DE LOS ESTILOS DE LOS BOTONES
        estilos_bts = ttk.Style()
        estilos_bts.theme_use("clam")
        estilos_bts.configure(
            "ButtonsStyles.TButton", background="white", bordercolor="#647cdc"
        )
        estilos_bts.map(
            "ButtonsStyles.TButton",
            background=[("active", "#e0eef9")],
            bordercolor=[("active", "#647cdc")],
        )

        # BOTÓN PARA REGRESAR A LA PAGINA PRINCIPAL QUE ES EL HOME
        bt_regresar_img = PhotoImage(data=imagenes.logo_regresar)
        bt_regresar = ttk.Button(
            window,
            image=bt_regresar_img,
            compound="left",
            text=" Regresar",
            padding=(20, 15, 20, 15),
            cursor="hand2",
            takefocus=False,
            style="ButtonsStyles.TButton",
            command=lambda: [{window.destroy(), Main()}],
        )
        bt_regresar.place(x=320, y=377)

        # BOTÓN DE QUE SE INICIE LA INSTALACIÓN
        bt_instalar_img = PhotoImage(data=imagenes.logo_descargar)
        bt_instalar = ttk.Button(
            window,
            image=bt_instalar_img,
            compound="left",
            text=" Instalar",
            padding=(20, 15, 20, 15),
            cursor="hand2",
            takefocus=False,
            style="ButtonsStyles.TButton",
            command=lambda: [
                {
                    instalar(
                        window, barra_progreso, bt_instalar, bt_regresar, text_install
                    )
                }
            ],
        )
        bt_instalar.place(x=480, y=377)

        # * MOSTRAR LA VENTANA CON TODA LA CONFIGURACIÓN * #
        window.protocol(
            "WM_DELETE_WINDOW",
            lambda: [{window.destroy(), Main()}],
        )
        window.deiconify()
        window.mainloop()


class Main:
    """Ventana que vera el usuario al momento de abrir la aplicación"""

    def __init__(self):
        # GLOBAL
        imagenes = Imagenes()

        ## * CONFIGURACIÓN DE LA VENTANA QUE VERA EL USUARIO * ##
        window = Tk()

        # ESTO HACE QUE LA VENTANA SE OCULTE AL USUARIO AL MOMENTO DE QUE SE INICIALICE
        window.withdraw()

        # TITULO DE LA VENTANA DE LA APLICACIÓN
        window.title("Office LTSC")

        # AQUÍ ESTA LLAMANDO A LA FUNCIÓN (center_window) PARA CENTRAR LA VENTANA EN MEDIO DE TODA LA PANTALLA DEL USUARIO & LE ESTAMOS PASANDO ALGUNOS VALORES
        center_window(window, 650, 460)

        # ICONO DE LA APLICACIÓN
        icon = PhotoImage(data=imagenes.logo_y_icon_web)
        window.iconphoto(False, icon)
        window.config(background="white")

        # EN ESTA PARTE ES PARA QUE EL USUARIO NO PUEDA CAMBIAR EL TAMAÑO DE LA VENTANA
        window.resizable(0, 0)

        ## * PANEL SUPERIOR  * ##

        # SE DIBUJARA UNA LINEA DIVISORA PARA DIVIDIR LA INFORMACIÓN
        linea_superior = Label(
            window,
            background="#647cdc",
        )
        linea_superior.place(width=650, height=1, y=70)

        # SE MUESTRA EL LOGO DE LA APLICACIÓN PARA QUE SEA MAS ENTENDIBLE PARA EL USUARIO
        office_img = PhotoImage(data=imagenes.logo_y_icon_web)
        office = Label(window, image=office_img, borderwidth=0, background="white")
        office.place(x=5, y=3)

        # SE MUESTRA INFORMACIÓN ALADO DEL LOGO PARA QUE EL USUARIO ENTIENDA EN QUE APARTADO DE ENCUENTRA
        info_ltsc = Label(
            window,
            text=" | Office LTSC - Términos y Condiciones",
            font=("Cascadia Code", 10),
            background="white",
            foreground="#647cdc",
        )
        info_ltsc.place(x=60, y=12)
        info_pagina = Label(
            window,
            text=" | Términos & Condiciones que aceptaras al instalar la aplicación",
            font=("Cascadia Code", 10),
            background="white",
            foreground="#8560b2",
        )
        info_pagina.place(x=60, y=34)

        ## *  TEXTO  * ##

        # CONFIGURACIÓN DEL SCROLLBAR Y EL CUADRO DE TEXTO QUE MOSTRARA LA INFORMACIÓN AL USUARIO
        scrollbar = Scrollbar(window, takefocus=False)
        term_condiciones = Text(
            window,
            background="#f0f0f0",
            foreground="black",
            font=("Cascadia Code", 10),
            borderwidth=0,
            wrap="word",
            selectbackground="#647cdc",
            selectforeground="white",
        )
        scrollbar.config(command=term_condiciones.yview)
        term_condiciones.insert(
            "end",
            f"\n============================= OFFICE LTSC =============================\n\nEsta aplicación es una automatización de la instalación de la librería\nde Office, al aceptar estas condiciones también estas aceptando los\ntérminos de condiciones de Microsoft Office,\n haciendo esto esta app\ntiene tu consideración de comenzar con la instalación de nuestros programas.\n\n============================= INFORMACIÓN =============================\n\nEsta aplicación es instalar de manera automática con tan solo algunos\nclick el programa de Office LTSC que es 100% original de Microsoft y\ntambién te instala otro programa que es 100% del desarrollador de este\ninstalador automático el cual realiza la función de reactivar el Office\nLTSC en el momento que se te desactive.\n\n=============================== SISTEMA ===============================\n\nEsta aplicación solo esta disponible para dispositivos con el sistema\noperativo Windows de ahi en fuera lamento que esta aplicación no sea\ncompatible para tu sistema operativo.\n\n=============================== CREADOR ===============================\n\nEl creador de esta pagina esta sujetos a términos de privacidad, lo\ncual no se podrá usar esta aplicación con fines de lucro o migración\na otras personas.\n\nRICKYTO DEV ©\n\n=============================== GITHUB ================================\n\nSi quieres conocer mas sobre el creador de la aplicación te invito a\nque te pases por mi GitHub donde podrás encontrar varias aplicaciones\npara el sistema operativo (OS) de Windows\n\nhttps://github.com/rickyto-dev\n\n=============================== DONAR =================================\n\nSi gustas donar para la seguir creando mas aplicaciones te lo agradezco\n\nde muchas maneras ya que realizar estas aplicaciones no son nada fácil\n\nsi quieres donar copea el siguiente link y págalo en tu navegador.\n\nhttps://paypal.me/xrickytox?country.x=MX&locale.x=es_XC\n\n=============================== GRACIAS ===============================\n\nMuchas Gracias! por llegar hasta este apartado y espero que mi\n\naplicación te sirva para poder disfrutar de estas librerías de Office\n\nproporcionada por Microsoft\n\n2023-2024 ©\n",
        )
        term_condiciones.config(yscrollcommand=scrollbar.set, state="disabled")
        term_condiciones.place(width=584, height=250, x=25, y=100)
        scrollbar.place(width=13, height=250, x=608, y=100)

        ## *  BOTONES  * ##

        # CONFIGURACIÓN DE LOS ESTILOS DE LOS BOTONES
        estilos_bts = ttk.Style()
        estilos_bts.theme_use("clam")
        estilos_bts.configure(
            "ButtonsStyles.TButton", background="white", bordercolor="#647cdc"
        )
        estilos_bts.map(
            "ButtonsStyles.TButton",
            background=[("active", "#e0eef9")],
            bordercolor=[("active", "#647cdc")],
        )

        # BOTÓN PARA CANCELAR LA INSTALACIÓN
        bt_cancelar_img = PhotoImage(data=imagenes.logo_cancelar)
        bt_cancelar = ttk.Button(
            window,
            image=bt_cancelar_img,
            compound="left",
            text=" Cancelar",
            padding=(20, 15, 20, 15),
            cursor="hand2",
            takefocus=False,
            style="ButtonsStyles.TButton",
            command=lambda: [{exit_app()}],
        )
        bt_cancelar.place(x=320, y=377)

        # BOTÓN DE QUE SE ACEPTARON LOS TÉRMINOS Y CONDICIONES
        bt_aceptar_img = PhotoImage(data=imagenes.logo_aceptar)
        bt_aceptar = ttk.Button(
            window,
            image=bt_aceptar_img,
            compound="left",
            text=" Aceptar",
            padding=(20, 15, 20, 15),
            cursor="hand2",
            style="ButtonsStyles.TButton",
            takefocus=False,
            command=lambda: [{window.destroy(), Install()}],
        )
        bt_aceptar.place(x=480, y=377)

        # * MOSTRAR LA VENTANA CON TODA LA CONFIGURACIÓN * #
        window.protocol("WM_DELETE_WINDOW", lambda: [{exit_app()}])
        window.deiconify()
        window.mainloop()


class ChargeApp:
    """Esto carga la aplicación para que cargue todos los datos como las [imagenes y demás cosas]"""

    def __init__(self):
        # GLOBAL
        imagenes = Imagenes()

        ## * CONFIGURACIÓN DE LA VENTANA QUE VERA EL USUARIO * ##
        window = Tk()

        # ESTO HACE QUE LA VENTANA SE OCULTE AL USUARIO AL MOMENTO DE QUE SE INICIALICE
        window.withdraw()

        # TITULO DE LA VENTANA DE LA APLICACIÓN
        window.title("Office LTSC")

        # AQUÍ ESTA LLAMANDO A LA FUNCIÓN (center_window) PARA CENTRAR LA VENTANA EN MEDIO DE TODA LA PANTALLA DEL USUARIO & LE ESTAMOS PASANDO ALGUNOS VALORES
        center_window(window, 414, 94)

        # ICONO DE LA APLICACIÓN
        icon = PhotoImage(data=imagenes.logo_y_icon_web)
        window.iconphoto(False, icon)
        window.config(background="white")

        # TEXTO DE QUE SE ESTA INICIANDO
        texto_iniciando = Label(
            window,
            text="Iniciando",
            font=("Cascadia Code", 10),
            background="white",
            foreground="#647cdc",
        )
        texto_iniciando.place(x=84, y=15)

        # BARRA DE PROGRESO
        estilo_progressbar = ttk.Style()
        estilo_progressbar.theme_use("clam")
        estilo_progressbar.configure(
            "TProgressbar",
            troughcolor="white",
            bordercolor="#f0f0f0",
            background="#647cdc",
            darkcolor="#647cdc",
            lightcolor="#647cdc",
        )
        self.barra_progreso = ttk.Progressbar(
            window, takefocus=False, orient="horizontal"
        )
        self.barra_progreso.place(width=300, height=30, x=84, y=40)

        # EN ESTA PARTE ES PARA QUE EL USUARIO NO PUEDA CAMBIAR EL TAMAÑO DE LA VENTANA
        window.resizable(0, 0)

        # LOGO DE LA APLICACIÓN
        logo_img = PhotoImage(data=imagenes.logo_y_icon_web)
        logo = Label(window, image=logo_img, background="white")
        logo.place(x=10, y=10)

        # INICIAR FUNCIÓN

        # * MOSTRAR LA VENTANA CON TODA LA CONFIGURACIÓN * #
        window.protocol("WM_DELETE_WINDOW", lambda: False)
        window.deiconify()
        window.after(100, lambda: self.iniciar_app(window))
        window.mainloop()

    def iniciar_app(self, win):
        valor = self.barra_progreso["value"]

        if valor < 100:
            # AGREGAR MAS 20 AL VALOR ACTUAL DE LA BARRA DE PROGRESO
            self.barra_progreso["value"] += 20
            self.barra_progreso.update()
            win.after(1000, lambda: self.iniciar_app(win))

        else:
            win.destroy()
            Main()


# * CHECA SI TIENE INTERNET EL USUARIO * #
def check_wifi():
    try:
        urllib.request.urlopen("http://www.google.com", timeout=1)
        return True
    except:
        return False


if check_wifi() == True:
    ChargeApp()
else:
    messagebox.showerror(
        "Error al Iniciar la Aplicación",
        "No se pudo iniciar la aplicación porque no tienes conexión a Internet en estos momentos. Por favor, actívalo o conéctate a una red móvil para instalar Office LTSC.",
    )
