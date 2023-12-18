from tkinter import Tk, ttk, PhotoImage, Label, Text, Scrollbar, messagebox
import subprocess
from urllib import request


class Imagenes:
    def __init__(self):
        self.icon_img = request.urlopen("https://i.imgur.com/0liU7zg.png").read()
        self.reactive_img_path = request.urlopen("https://imgur.com/or2385J.png").read()


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


# aplicación
class Main:
    """
    Aplicación de inicio donde se podrá realizar la activación del programa
    """

    def __init__(self):
        # GLOBAL
        imagenes = Imagenes()

        ## * CONFIGURACIÓN DE LA VENTANA QUE VERA EL USUARIO * ##
        window = Tk()

        # ESTO HACE QUE LA VENTANA SE OCULTE AL USUARIO AL MOMENTO DE QUE SE INICIALICE
        window.withdraw()

        # TITULO DE LA VENTANA DE LA APLICACIÓN
        window.title("Office Reactivador")

        # AQUÍ ESTA LLAMANDO A LA FUNCIÓN (center_window) PARA CENTRAR LA VENTANA EN MEDIO DE TODA LA PANTALLA DEL USUARIO & LE ESTAMOS PASANDO ALGUNOS VALORES
        center_window(window, 448, 458)

        # ICONO DE LA APLICACIÓN
        icon = PhotoImage(data=imagenes.icon_img)
        window.iconphoto(False, icon)

        # EN ESTA PARTE ES PARA QUE EL USUARIO NO PUEDA CAMBIAR EL TAMAÑO DE LA VENTANA
        window.resizable(0, 0)

        ## * PANEL SUPERIOR  * ##

        background_top = Label(window, background="white")
        background_top.place(width=450, height=70)

        _linea_superior = Label(
            window,
            background="#b5b5b5",
        )
        _linea_superior.place(width=650, height=1, y=70)

        # LOGO DEL PANEL SUPERIOR
        _office_img = PhotoImage(data=imagenes.icon_img)
        _office = Label(window, image=_office_img, borderwidth=0, background="white")
        _office.place(x=380, y=3)

        # INFORMACIÓN DE LA APLICACIÓN
        _info_ltsc = Label(
            window,
            text="Office LTSC",
            font=("Cascadia Code", 12),
            foreground="#eb3b00",
            background="white",
        )
        _info_ltsc.place(x=10, y=12)
        _info_panel = Label(
            window,
            text="Reactiva tu OfficeLTSC con un solo click.",
            font=("Cascadia Code", 9),
            background="white",
        )
        _info_panel.place(x=11, y=34)

        # CONDICIONES DE USO POR EL USUARIO
        _scrollbar = Scrollbar(window, takefocus=False)
        _term_condiciones = Text(
            window,
            background="white",
            foreground="black",
            font=("Cascadia Code", 10),
            selectbackground="#eb3b00",
            selectforeground="white",
        )
        _scrollbar.config(command=_term_condiciones.yview)
        _term_condiciones.insert(
            "end",
            f"\n================= OFFICE LTSC =================\n\nEsta aplicación esta encargada de la\nreactivación del Office por si se te llegase a\ndesactivar pero con esta aplicación con solo un\n'click' lo puedes realizar las veces que se te\nllegue a desactivar\n\n=================    PASOS    =================\n\nSi quieres realizar la operación que realiza\nesta aplicación sigue los siguientes pasos que\nte dejo a continuación.\n\n1. Abre el buscador de archivos y copea y pega\nen la barra de direcciones la siguiente ruta\n\nC:/Program Files/Microsoft Office/Office16\n\n2. Ejecuta el programa con el nombre\nOSPPREARM.EXE tres veces con el modo\nadministrador.\n\n3. Listo ya quedo otra ves Office LTSC activado\nnuevamente.\n\n4. (Opcional) Reinicia tu computadora para que\nno aya ningún inconveniente al abrir la\npaquetería de Office LTSC activado.\n\n=================   SISTEMA   =================\n\nEsta aplicación solo esta disponible para\ndispositivos con el sistema operativo Windows\nde ahi en fuera lamento que esta aplicación no\nsea compatible para tu sistema operativo.\n\n=================   CREADOR   =================\n\nEl creador de esta pagina esta sujetos a\ntérminos de privacidad, lo cual no se podrá\nusar esta aplicación con fines de lucro o\nmigración a otras personas.\n\nRICKYTO DEV ©\n\n=================   GITHUB    =================\n\nSi quieres conocer mas sobre el creador de la\naplicación te invito a que te pases por mi\nGitHub donde podrás encontrar varias\naplicaciones para el sistema operativo (OS) de\nWindows\n\nhttps://github.com/rickyto-dev\n\n=================    DONAR    =================\n\nSi gustas donar para la seguir creando mas\naplicaciones te lo agradezco de muchas maneras\nya que realizar estas aplicaciones no son nada\n fácil si quieres donar copea el siguiente link\ny págalo en tu navegador.\n\nhttps://paypal.me/xrickytox?country.x=MX&locale.x=es_XC\n\n=================   GRACIAS   =================\n\nMuchas Gracias! por llegar hasta este apartado\n\ny espero que mi aplicación te sirva para poder\ndisfrutar de estas librerías de Office\nproporcionada por Microsoft\n\n2023-2024 ©\n",
        )
        _term_condiciones.config(yscrollcommand=_scrollbar.set, state="disabled")
        _term_condiciones.place(width=382, height=250, x=27, y=98)
        _scrollbar.place(width=18, height=250, x=408, y=98)

        # BOTÓN PARA REACTIVAR EL OFFICE LTSC
        _reactivar_office_img = PhotoImage(data=imagenes.reactive_img_path)
        _reactivar_office = ttk.Button(
            window,
            image=_reactivar_office_img,
            compound="right",
            text="Reactivar  ",
            cursor="hand2",
            takefocus=False,
            padding=(20, 15, 20, 15),
            command=lambda: [
                {
                    messagebox.showwarning(
                        "Cuidado en Ejecución",
                        "Se ejecutara el programa de reactivación espera mientras termina y le aparezca el mensaje que ya termino la reactivación y que se realizo correctamente.\n\nPreciosa el botón de 'Aceptar' o 'Enter' para que se ejecute el programa",
                    ),
                    subprocess.run("OSPPREARM.EXE"),
                    subprocess.run("OSPPREARM.EXE"),
                    subprocess.run("OSPPREARM.EXE"),
                    messagebox.showinfo(
                        "Reactivación Realizada Correctamente",
                        "Se realizo la reactivación de la paquetería de Office LTSC de manera correcta, yo puede volver a disfrutar de sus aplicaciones que son las siguientes:\n\n > Word\n\n > PowerPoint\n\n > Excel\n\nSi presenta algún error con esta aplicación comunicarse al siguiente correo: \n\nhelp.aplications.github.rickyto@gmail.com",
                    ),
                }
            ],
        )
        _reactivar_office.place(x=280, y=373)

        # mostrar la ventana
        window.deiconify()
        window.mainloop()


# * CHECA SI TIENE INTERNET EL USUARIO * #
def check_wifi():
    try:
        request.urlopen("http://www.google.com", timeout=1)
        return True
    except:
        return False


if check_wifi() == True:
    Main()
else:
    messagebox.showerror(
        "Error al Iniciar la Aplicación",
        "Nose a podido iniciar la aplicación porque no tienes internet en estos momentos, por favor activarlo o conectarse a una red móvil para Reactivar Office LTSC",
    )
