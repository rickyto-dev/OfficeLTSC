from tkinter import Tk, ttk, PhotoImage, Label, Text, Scrollbar, messagebox
import subprocess
import os

try:
    icon_img = (
        "C:/Program Files/Microsoft Office/Office16/images/office_reactivador.ico"
    )
    if not os.path.isfile(icon_img):
        raise FileNotFoundError
except FileNotFoundError:
    icon_img = "images/office_reactivador.ico"

try:
    logo_img_path = "C:/Program Files/Microsoft Office/Office16/images/logo.png"
    if not os.path.isfile(logo_img_path):
        raise FileNotFoundError
except FileNotFoundError:
    logo_img_path = "images/logo.png"

try:
    windows_img_path = "C:/Program Files/Microsoft Office/Office16/images/windows.png"
    if not os.path.isfile(windows_img_path):
        raise FileNotFoundError
except FileNotFoundError:
    windows_img_path = "images/windows.png"

try:
    terms_path = (
        "C:/Program Files/Microsoft Office/Office16/documents/LicenseReactivador.txt"
    )
    if not os.path.isfile(terms_path):
        raise FileNotFoundError
except FileNotFoundError:
    terms_path = "documents/LicenseReactivador.txt"

try:
    reactive_img_path = (
        "C:/Program Files/Microsoft Office/Office16/images/reactivar.png"
    )
    if not os.path.isfile(reactive_img_path):
        raise FileNotFoundError
except FileNotFoundError:
    reactive_img_path = "images/reactivar.png"


# leer términos y condiciones de la aplicación
with open(terms_path, "r", encoding="utf-8") as _license:
    LicenseOfficeActivador = _license.read()


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
        # configuración de la ventana
        window = Tk()
        window.withdraw()
        center_window(window, 450, 460)
        window.title("Office LTSC  |  Reactivador")
        window.iconbitmap(icon_img)
        window.resizable(False, False)
        window.config(background="white")
        # información
        ## panel superior
        _linea_superior = Label(
            window,
            background="#b5b5b5",
        )
        _linea_superior.place(width=650, height=1, y=70)
        ### > logo del panel superior
        _office_img = PhotoImage(file=logo_img_path)
        _office = Label(window, image=_office_img, borderwidth=0, background="white")
        _office.place(x=5, y=3)
        ### > información
        _info_ltsc = Label(
            window,
            text=" | Office LTSC - Reactivador",
            font=("Cascadia Code", 10),
            background="white",
            foreground="#eb3b00",
        )
        _info_ltsc.place(x=60, y=12)
        _info_panel = Label(
            window,
            text=" | Reactivar Office LTSC con un click",
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
        _term_condiciones.insert("end", f"{LicenseOfficeActivador}")
        _term_condiciones.config(yscrollcommand=_scrollbar.set, state="disabled")
        _term_condiciones.place(width=382, height=250, x=25, y=90)
        _scrollbar.place(width=18, height=250, x=405, y=90)
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
        ## botón para reactivar office
        _reactivar_office_img = PhotoImage(file=reactive_img_path)
        _reactivar_office = ttk.Button(
            window,
            image=_reactivar_office_img,
            compound="left",
            text="Reactivar",
            cursor="hand2",
            takefocus=False,
            style="ButtonsStyles.TButton",
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
        _reactivar_office.place(x=280, y=370)
        ## imagen de windows
        window_img = PhotoImage(file=windows_img_path)
        window_label = Label(window, image=window_img, background="white")
        window_label.place(x=15, y=405)
        # mostrar la ventana
        window.deiconify()
        window.mainloop()


if __name__ == "__main__":
    Main()
