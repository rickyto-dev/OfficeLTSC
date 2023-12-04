# IMPORTACIÓN DE TODAS LAS LIBRERÍAS QUE SE UTILIZARAN PARA LA INSTALACIÓN
import os, shutil, getpass, zipfile, subprocess, traceback, winshell, time

# OBTIENE EL NOMBRE DE USUARIO DE LA COMPUTADORA
_usuario = getpass.getuser()

# TODAS LAS RUTAS QUE SE UTILIZARAN PARA LOS MOVIMIENTOS Y CREACIÓN DE ACCESOS DIRECTOS
_office_path = f"C:/Program Files/Microsoft Office/Office16/"
_exe = f"{_office_path}/office-reactivador/OfficeLTSC Reactivador.exe"
_images = f"{_office_path}/office-reactivador/images"
_documentos = f"{_office_path}/office-reactivador/documents"

# NOMBRE DE LA APLICACIÓN & EN QUE RUTA SE ENCUENTRA ESTA
_app_path = f"{_office_path}/OfficeLTSC Reactivador.exe"

# CREACIÓN DE ACCESOS DIRECTOS
_name_link = "OfficeLTSC Reactivador.lnk"
_menu_escritorio = f"C:/Users/{_usuario}/OneDrive/Escritorio"

# ARCHIVOS DENTRO DEL EMPAQUETADOR DE OFFICE
_images_office = f"{_office_path}images"
_documentos_office = f"{_office_path}documentos"

# RUTAS PARA LA INSTALACIÓN DE OFFICE LTSC POR CONSOLA & EL COMANDO DE INSTALACIÓN
_office_arch_exe = "office-setup/setup.exe"
_office_arch_xml = "office-setup/config.xml"
_comando_install_office = "setup /configure config.xml"


# EXTRAE EL PRIMER ARCHIVO & REALIZA LAS FUNCIONES QUE SE TIENEN QUE REALIZAR EN CADA APARTADO
with zipfile.ZipFile("documentos/office-setup.zip", "r") as zip_ref_office:
    zip_ref_office.extractall(os.getcwd())

    # MUEVE LOS ARCHIVOS A LA CARPETA ACTUAL
    shutil.move(_office_arch_exe, os.getcwd())
    shutil.move(_office_arch_xml, os.getcwd())

    # ELIMINA DE MANERA RÁPIDA LA CARPETA EXTRAÍDA
    shutil.rmtree(f"{os.getcwd()}/office-setup", ignore_errors=True)

    # OCULTA LOS ARCHIVOS EXTRAÍDOS
    os.system(f'attrib +h "{os.getcwd()}/setup.exe"')
    os.system(f'attrib +h "{os.getcwd()}/config.xml"')

    # TOMA UN TIEMPO DE 1s PARA CONTINUAR CON LAS SIGUIENTE FUNCIONES Y NO AYA NINGÚN PROBLEMA
    time.sleep(1)

    # EJECUTA EL COMANDO PARA INICIAR CON LA INSTALACIÓN DEL PROGRAMA DE OFFICE LTSC POR MEDIO DE CONSOLA POR LO QUE SE MOSTRARA UNA CONSOLA AL USUARIO DE MANERA RÁPIDA Y DESPUÉS SE OCULTARA
    subprocess.run(_comando_install_office, shell=True)

    # TERMINA EL PROCESO PARA QUE SE EJECUTE LAS SIGUIENTES FUNCIONES
    subprocess.run(["taskkill", "/IM", "OfficeC2RClient.exe", "/F"], check=True)


# EXTRAE EL SEGUNDO ARCHIVO & REALIZA LAS FUNCIONES QUE SE TIENEN QUE REALIZAR EN CADA APARTADO
with zipfile.ZipFile(
    "documentos/office-reactivador.zip", "r"
) as zip_ref_office_reactivador:
    zip_ref_office_reactivador.extractall(_office_path)

    # DESPUÉS DE TODA LA INSTALACIÓN SE REMUEVEN LOS ARCHIVOS DE INSTALACIÓN
    os.remove(f"setup.exe")
    os.remove(f"config.xml")

    # TOMA UNA PAUSA DE .2s PARA CONTINUAR CON LAS FUNCIONES
    time.sleep(0.2)
    shutil.move(_exe, _office_path)
    shutil.move(_images, _office_path)
    shutil.move(_documentos, _office_path)
    os.system(f'attrib +h "{_images_office}"')
    os.system(f'attrib +h "{_documentos_office}"')
    shutil.rmtree(f"{_office_path}office-reactivador", ignore_errors=True)

# CREACIÓN DEL ACCESO DIRECTO EN EL ESCRITORIO
_ESCRITORIO = os.path.join(_menu_escritorio, _name_link)
_escritorio_comando = f'mklink "{_ESCRITORIO}" "{_app_path}" '
subprocess.run(_escritorio_comando, shell=True)

# CREACIÓN DEL ACCESO DIRECTO EN MENU DE APLICACIONES
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

# PARA FINALIZAR TERMINA TODA LA APLICACIÓN
subprocess.run(
    ["taskkill", "/IM", "OfficeInstallerClienteLTSCServices.exe", "/F"], check=True
)
