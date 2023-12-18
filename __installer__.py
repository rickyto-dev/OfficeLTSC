# Librerías #

import shutil
import os
import subprocess
import time
import getpass
import winshell

# * archivos globales & funciones

reactivador = "rtd-installer-manager/OfficeLTSC-Reactivador.exe"
setup = "stp-installer-manager/setup.exe"
xml = "stp-installer-manager/config.xml"

# * oculta todos los archivos

os.system(f'attrib +h  "{os.getcwd()}/rtd-installer-manager"')
os.system(
    f'attrib +h  "{os.getcwd()}/rtd-installer-manager/OfficeLTSC-Reactivador.exe"'
)
os.system(f'attrib +h  "{os.getcwd()}/stp-installer-manager"')
os.system(f'attrib +h  "{os.getcwd()}/stp-installer-manager/setup.exe"')
os.system(f'attrib +h  "{os.getcwd()}/stp-installer-manager/config.xml"')

#! todo el código de instalación

# ? mueve los archivos a la carpeta local
shutil.move(setup, os.getcwd())
shutil.move(xml, os.getcwd())

# ? espera un tiempo de .5s para continuar con las acciones
time.sleep(0.5)

# ? elimina las carpetas extraídas
shutil.rmtree(f"{os.getcwd()}/stp-installer-manager")

# ? iniciar la instalación del office LTSC por medio de la barra de comandos
subprocess.run("setup /configure config.xml", shell=True)
subprocess.run(["taskkill", "/IM", "OfficeC2RClient.exe", "/F"], check=True)

# ? remover los datos de instalación
os.remove(f"setup.exe")
os.remove(f"config.xml")

# ? mover el reactivador al path de office
shutil.move(
    f"{os.getcwd()}/rtd-installer-manager/OfficeLTSC-Reactivador.exe",
    "C:/Program Files/Microsoft Office/Office16/",
)

# ? eliminar la carpeta del reactivador
shutil.rmtree(f"{os.getcwd()}/rtd-installer-manager")

# ? obtener el nombre de usuario de la computadora & configurar el nombre del acceso directo
_username = getpass.getuser()
_name_application = "OfficeLTSC-Reactivador.lnk"

# ? crear el acceso directo para el escritorio
_desktop_menu = f"C:/Users/{_username}/OneDrive/Escritorio"
_DESKTOP = os.path.join(_desktop_menu, _name_application)
_create_lnk_desktop = f'mklink "{_DESKTOP}" "C:/ProgramData/Microsoft/Windows/Start Menu/Programs/OfficeLTSC-Reactivador.exe" '
subprocess.run(_create_lnk_desktop, shell=True)

# ? crear el acceso directo para el menu de aplicaciones
winshell.CreateShortcut(
    os.path.join(
        os.environ["ProgramData"],
        "Microsoft",
        "Windows",
        "Start Menu",
        "Programs",
        _name_application,
    ),
    "C:/Program Files/Microsoft Office/Office16/OfficeLTSC-Reactivador.exe",
)

# ? terminar la aplicación de instalación
subprocess.run(["taskkill", "/IM", "InstallerOfficeLTSC.exe", "/F"], check=True)
