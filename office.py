from tkinter import Tk, Label, ttk, PhotoImage, messagebox, Text, Scrollbar
import sys, subprocess

descarga = 0


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
        messagebox.showinfo(
            "Instalación Exitosa",
            "Se ha instalado la aplicación de Office LTSC de manera correcta en tu dispositivo Windows",
        )


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
            win.destroy()
            subprocess.check_call(
                [
                    "runas",
                    "/savecred",
                    "/user:Administrator",
                    "aplicaciones/OfficeInstallerClienteLTSCServices.exe",
                ]
            )
            Finish()

    # INICIAR LA BARRA DE PROGRESO
    progress_chargee()


with open("licencias/LicenseOffice.txt", "r", encoding="utf-8") as _licencia:
    LICENCIA = _licencia.read()


# APLICACIÓN
class Finish:
    """
    Este es el apartado de finalización de la aplicación
    """

    def __init__(self):
        ## * CONFIGURACIÓN DE LA VENTANA QUE VERA EL USUARIO * ##

        window = Tk()
        # ESTO HACE QUE LA VENTANA SE OCULTE AL USUARIO AL MOMENTO DE QUE SE INICIALICE
        window.withdraw()
        # TITULO DE LA VENTANA DE LA APLICACIÓN
        window.title("Office LTSC | Instalador")
        # AQUÍ ESTA LLAMANDO A LA FUNCIÓN (center_window) PARA CENTRAR LA VENTANA EN MEDIO DE TODA LA PANTALLA DEL USUARIO & LE ESTAMOS PASANDO ALGUNOS VALORES
        center_window(window, 650, 460)
        # ICONO DE LA APLICACIÓN
        window.iconbitmap(f"imagenes/office.ico")
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
        office_img = PhotoImage(file=r"imagenes/logo.png")
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
        bt_finalizar_img = PhotoImage(file=r"imagenes/aceptar.png")
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

        # FUNCIONES PARA REALIZAR AL MOMENTO QUE EL USUARIO INTENTE CERRAR LA VENTANA
        window.protocol("WM_DELETE_WINDOW", lambda: [{window.destroy(), sys.exit()}])

        # MOSTRAR LA VENTANA CON TODA LA INFORMACIÓN & CON TODA LA CONFIGURACIÓN
        window.deiconify()
        window.mainloop()


class Install:
    """
    Este es el apartado de instalación
    """

    def __init__(self):
        ## * CONFIGURACIÓN DE LA VENTANA QUE VERA EL USUARIO * ##

        window = Tk()
        # ESTO HACE QUE LA VENTANA SE OCULTE AL USUARIO AL MOMENTO DE QUE SE INICIALICE
        window.withdraw()
        # TITULO DE LA VENTANA DE LA APLICACIÓN
        window.title("Office LTSC | Instalador")
        # AQUÍ ESTA LLAMANDO A LA FUNCIÓN (center_window) PARA CENTRAR LA VENTANA EN MEDIO DE TODA LA PANTALLA DEL USUARIO & LE ESTAMOS PASANDO ALGUNOS VALORES
        center_window(window, 650, 460)
        # ICONO DE LA APLICACIÓN
        window.iconbitmap(f"imagenes/office.ico")
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
        office_img = PhotoImage(file=r"imagenes/logo.png")
        office = Label(window, image=office_img, borderwidth=0, background="white")
        office.place(x=5, y=3)
        # SE MUESTRA INFORMACIÓN ALADO DEL LOGO PARA QUE EL USUARIO ENTIENDA EN QUE APARTADO DE ENCUENTRA
        info_ltsc = Label(
            window,
            text=" | Office LTSC - Office LTSC - Instalación",
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
        bt_regresar_img = PhotoImage(file=r"imagenes/regresar.png")
        bt_regresar = ttk.Button(
            window,
            image=bt_regresar_img,
            compound="left",
            text=" Regresar",
            padding=(20, 15, 20, 15),
            cursor="hand2",
            takefocus=False,
            style="ButtonsStyles.TButton",
            command=lambda: [{window.destroy(), Home()}],
        )
        bt_regresar.place(x=320, y=377)
        # BOTÓN DE QUE SE INICIE LA INSTALACIÓN
        bt_instalar_img = PhotoImage(file=r"imagenes/administrador.png")
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

        # FUNCIONES PARA REALIZAR AL MOMENTO QUE EL USUARIO INTENTE CERRAR LA VENTANA
        window.protocol("WM_DELETE_WINDOW", lambda: [{window.destroy(), Home()}])

        # MOSTRAR LA VENTANA CON TODA LA INFORMACIÓN & CON TODA LA CONFIGURACIÓN
        window.deiconify()
        window.mainloop()


class Home:
    """
    Esto es lo primero que vera el usuario al iniciar la aplicación
    """

    def __init__(self):
        ## * CONFIGURACIÓN DE LA VENTANA QUE VERA EL USUARIO * ##

        window = Tk()
        # ESTO HACE QUE LA VENTANA SE OCULTE AL USUARIO AL MOMENTO DE QUE SE INICIALICE
        window.withdraw()
        # TITULO DE LA VENTANA DE LA APLICACIÓN
        window.title("Office LTSC | Instalador")
        # AQUÍ ESTA LLAMANDO A LA FUNCIÓN (center_window) PARA CENTRAR LA VENTANA EN MEDIO DE TODA LA PANTALLA DEL USUARIO & LE ESTAMOS PASANDO ALGUNOS VALORES
        center_window(window, 650, 460)
        # ICONO DE LA APLICACIÓN
        window.iconbitmap(f"imagenes/office.ico")
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
        office_img = PhotoImage(file=r"imagenes/logo.png")
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
        term_condiciones.insert("end", f"{LICENCIA}")
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
        bt_cancelar_img = PhotoImage(file=r"imagenes/cancelar.png")
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
        bt_aceptar_img = PhotoImage(file=r"imagenes/aceptar.png")
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

        # FUNCIONES PARA REALIZAR AL MOMENTO QUE EL USUARIO INTENTE CERRAR LA VENTANA
        window.protocol("WM_DELETE_WINDOW", lambda: [{exit_app()}])

        # MOSTRAR LA VENTANA CON TODA LA INFORMACIÓN & CON TODA LA CONFIGURACIÓN
        window.deiconify()
        window.mainloop()


# INICIAR LA APLICACIÓN EN EL SIGUIENTE APARTADO
Home()
