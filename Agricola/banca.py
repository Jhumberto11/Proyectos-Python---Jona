from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from getpass import getpass
import os


URL = "https://www.ebanca.com/login"

# Perfil independiente para esta automatización
PERFIL = os.path.join(
    os.environ["USERPROFILE"],
    "Chrome_Ebanca"
)


def iniciar():

    print("=" * 50)
    print("E-BANCA BANCO AGRICOLA")
    print("=" * 50)

    usuario = input("Usuario: ").strip()

    # No se muestra mientras escribes
    password = getpass("Contraseña: ")

    options = Options()

    # --------------------------------------------------
    # PERFIL PERSISTENTE
    # Conserva cookies/configuración del navegador
    # --------------------------------------------------

    options.add_argument(
        f"--user-data-dir={PERFIL}"
    )

    # --------------------------------------------------
    # NO CERRAR CHROME CUANDO TERMINE PYTHON
    # --------------------------------------------------

    options.add_experimental_option(
        "detach",
        True
    )

    # Ventana maximizada
    options.add_argument("--start-maximized")


    print("\nAbriendo e-banca...")


    # Selenium Manager administra ChromeDriver automáticamente
    driver = webdriver.Chrome(
        options=options
    )

    driver.get(URL)


    print()
    print("✓ E-banca abierta")
    print("✓ Chrome permanecerá abierto")
    print("✓ Perfil del navegador persistente")
    print()
    print("Usuario ingresado en el programa:", usuario)
    print()
    print(
        "Completa el acceso en la página del banco, "
        "incluyendo Clave Dinámica si te la solicita."
    )

    # IMPORTANTE:
    # No usamos driver.quit()
    # porque queremos dejar Chrome abierto.


if __name__ == "__main__":
    iniciar()