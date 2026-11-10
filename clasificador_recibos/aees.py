import imaplib
import email
import os
import re
import csv
import shutil

from datetime import datetime
from getpass import getpass
from pypdf import PdfReader
from email.header import decode_header, make_header



# ==========================================================
# CONFIGURACIÓN
# ==========================================================

SERVIDOR = "imap.mail.yahoo.com"
PUERTO = 993

REMITENTE_AES = "facturaelectronica@aeselsalvador.com"

CARPETA_BASE = "Recibos_AES"

# Aquí pondremos posteriormente tus cuentas reales
PROPIETARIOS = {
    "200000412312": "PLANTA BAJA - BOBBY -- FRANCO",
    "200001204283": "APA 1",
    "200001154568": "2DA PLANTA",
    "200000411162": "Apa5 - Lavadora- Secadora",
    "200000949845": "Apa6 -7 -8"
}


# ============================================================
# FECHA INICIO DEL MES
# ============================================================

def obtener_inicio_mes_imap():

    hoy = datetime.now()

    meses = [
        "Jan", "Feb", "Mar", "Apr",
        "May", "Jun", "Jul", "Aug",
        "Sep", "Oct", "Nov", "Dec"
    ]

    return f"01-{meses[hoy.month - 1]}-{hoy.year}"



# ============================================================
# DECODIFICAR NOMBRE DE ARCHIVO
# ============================================================

def decodificar_texto(texto):

    if not texto:
        return ""

    try:
        return str(
            make_header(
                decode_header(texto)
            )
        )
    except:
        return texto


# ============================================================
# LIMPIAR NOMBRE DE CARPETA PARA WINDOWS
# ============================================================

def limpiar_nombre_carpeta(nombre):

    nombre = re.sub(
        r'[<>:"/\\|?*]',
        ' ',
        nombre
    )

    nombre = re.sub(
        r'\s+',
        ' ',
        nombre
    )

    nombre = nombre.strip().rstrip(". ")

    if not nombre:
        nombre = "SIN_NOMBRE"

    return nombre


# ============================================================
# LEER PDF
# ============================================================

def leer_pdf(ruta_pdf):

    try:

        lector = PdfReader(ruta_pdf)

        texto = ""

        for pagina in lector.pages:

            contenido = pagina.extract_text()

            if contenido:

                texto += contenido + "\n"

        return texto

    except Exception as error:

        print(
            f"ERROR leyendo PDF: {error}"
        )

        return ""


# ============================================================
# EXTRAER NÚMERO DE CUENTA
# ============================================================

def extraer_nc(texto, nombre_archivo=""):

    # Primero intentar obtenerlo del contenido

    resultado = re.search(
        r'\bNC\s+(\d{12})\b',
        texto,
        re.IGNORECASE
    )

    if resultado:

        return resultado.group(1)


    # Si falla, intentar obtenerlo del nombre:
    # AESACE_200000411162_....

    resultado = re.search(
        r'AESACE_(\d{12})_',
        nombre_archivo,
        re.IGNORECASE
    )

    if resultado:

        return resultado.group(1)


    return ""


# ============================================================
# EXTRAER CLIENTE
# ============================================================

def extraer_cliente(texto):

    resultado = re.search(
        r'Cliente:\s*\n?\s*([^\n\r]+)',
        texto,
        re.IGNORECASE
    )

    if resultado:

        return resultado.group(1).strip()


    return ""


# ============================================================
# EXTRAER FECHA DE VENCIMIENTO
# ============================================================

def extraer_vencimiento(texto):

    resultado = re.search(
        r'Fecha Vencimiento:.*?'
        r'(\d{2}/\d{2}/\d{4})',
        texto,
        re.IGNORECASE | re.DOTALL
    )

    if resultado:

        return resultado.group(1)


    return ""


# ============================================================
# EXTRAER TOTAL
# ============================================================

def extraer_total(texto):

    patrones = [

        r'TOTAL AES CAESS\s+\$?\s*([0-9]+\.\d{2})',

        r'TOTAL CAESS \+ OTROS SERVICIOS \+'
        r'\s*ALCALDIA\s+\$?\s*([0-9]+\.\d{2})'
    ]


    for patron in patrones:

        resultado = re.search(
            patron,
            texto,
            re.IGNORECASE
        )

        if resultado:

            return resultado.group(1)


    return ""


# ============================================================
# EXTRAER NPE AES
# ============================================================

def extraer_npe(texto):

    """
    Formato observado:

    0189 0000 1221 2026 0822 0200 0004 1116 25

    Son:

    4 + 4 + 4 + 4 + 4 + 4 + 4 + 4 + 2 dígitos

    pypdf algunas veces concatena antes del NPE
    los 12 dígitos del Doc. Interno.

    Ejemplo:

    2500247681060189 0000 1221 ...
    """

    patron = re.compile(

        r'(?:\d{12})?'

        r'(\d{4})\s+'
        r'(\d{4})\s+'
        r'(\d{4})\s+'
        r'(\d{4})\s+'
        r'(\d{4})\s+'
        r'(\d{4})\s+'
        r'(\d{4})\s+'
        r'(\d{4})\s+'
        r'(\d{2})(?!\d)'
    )


    resultado = patron.search(texto)


    if resultado:

        # Unir todos los grupos sin espacios

        npe = "".join(
            resultado.groups()
        )

        return npe


    return ""


# ============================================================
# PROCESAR
# ============================================================

def procesar():

    print()
    print("==========================================")
    print("      EXTRACTOR DE RECIBOS AES / CAESS")
    print("==========================================")
    print()


    correo = "imosa13@yahoo.com"
    password = "lkojvljjevjcjtuh"


    print()
    print("Conectando a Yahoo...")


    mail = imaplib.IMAP4_SSL(
        SERVIDOR,
        PUERTO
    )


    mail.login(
        correo,
        password
    )


    print("✓ Conectado correctamente")


    # --------------------------------------------------------
    # ABRIR BANDEJA EN SOLO LECTURA
    # --------------------------------------------------------

    mail.select(
        "INBOX",
        readonly=True
    )


    fecha_inicio = obtener_inicio_mes_imap()


    print()
    print(
        f"Buscando correos AES desde: {fecha_inicio}"
    )


    # --------------------------------------------------------
    # SOLO CORREOS AES DEL MES ACTUAL
    # --------------------------------------------------------

    estado, resultado = mail.search(

        None,

        'FROM',
        f'"{REMITENTE_AES}"',

        'SINCE',
        fecha_inicio
    )


    if estado != "OK":

        print(
            "No se pudieron buscar los correos."
        )

        mail.logout()

        return


    mensajes = resultado[0].split()


    print(
        f"Correos AES encontrados: {len(mensajes)}"
    )


    # --------------------------------------------------------
    # CARPETA TEMPORAL
    # --------------------------------------------------------

    carpeta_temp = os.path.join(
        CARPETA_BASE,
        "_TEMP"
    )


    os.makedirs(
        carpeta_temp,
        exist_ok=True
    )


    lista_npes = []


    # ========================================================
    # RECORRER CORREOS
    # ========================================================

    for numero in mensajes:


        estado, datos_email = mail.fetch(

            numero,

            "(BODY.PEEK[])"
        )


        if estado != "OK":
            continue


        mensaje = email.message_from_bytes(
            datos_email[0][1]
        )


        asunto = decodificar_texto(
            mensaje.get("Subject", "")
        )


        # ====================================================
        # BUSCAR PDFs
        # ====================================================

        for parte in mensaje.walk():


            nombre_archivo = parte.get_filename()


            if not nombre_archivo:
                continue


            nombre_archivo = decodificar_texto(
                nombre_archivo
            )


            if not nombre_archivo.lower().endswith(
                ".pdf"
            ):
                continue


            # ------------------------------------------------
            # DESCARGAR TEMPORALMENTE
            # ------------------------------------------------

            ruta_temp = os.path.join(

                carpeta_temp,

                nombre_archivo
            )


            contenido = parte.get_payload(
                decode=True
            )


            with open(
                ruta_temp,
                "wb"
            ) as archivo:

                archivo.write(
                    contenido
                )


            # ------------------------------------------------
            # LEER PDF
            # ------------------------------------------------

            texto_pdf = leer_pdf(
                ruta_temp
            )


            # ------------------------------------------------
            # EXTRAER DATOS
            # ------------------------------------------------

            nc = extraer_nc(
                texto_pdf,
                nombre_archivo
            )


            cliente = extraer_cliente(
                texto_pdf
            )


            npe = extraer_npe(
                texto_pdf
            )


            total = extraer_total(
                texto_pdf
            )


            vencimiento = extraer_vencimiento(
                texto_pdf
            )


            # ------------------------------------------------
            # IDENTIFICAR PROPIETARIO
            # ------------------------------------------------

            propietario = PROPIETARIOS.get(

                nc,

                cliente if cliente else f"SIN_ASIGNAR_{nc}"
            )


            # ------------------------------------------------
            # MOSTRAR
            # ------------------------------------------------

            print()
            print(
                "------------------------------------------"
            )

            print(
                f"Propietario: {propietario}"
            )

            print(
                f"NC:          {nc}"
            )

            print(
                f"Cliente:     {cliente}"
            )

            print(
                f"Total:       ${total}"
            )

            print(
                f"Vencimiento: {vencimiento}"
            )

            print(
                f"NPE:         {npe}"
            )


            # ------------------------------------------------
            # CREAR CARPETA DEL PROPIETARIO
            # ------------------------------------------------

            nombre_carpeta = limpiar_nombre_carpeta(
                propietario
            )


            carpeta_persona = os.path.join(

                CARPETA_BASE,

                nombre_carpeta
            )


            os.makedirs(
                carpeta_persona,
                exist_ok=True
            )


            # ------------------------------------------------
            # MOVER PDF
            # ------------------------------------------------

            destino = os.path.join(

                carpeta_persona,

                nombre_archivo
            )


            if os.path.exists(destino):

                os.remove(destino)


            shutil.move(
                ruta_temp,
                destino
            )


            # ------------------------------------------------
            # GUARDAR NPE
            # ------------------------------------------------

            if npe:

                lista_npes.append({

                    "propietario": propietario,

                    "nc": nc,

                    "npe": npe,

                    "total": total,

                    "vencimiento": vencimiento
                })

            else:

                print()
                print(
                    "⚠ NO SE PUDO EXTRAER EL NPE"
                )


    # ========================================================
    # CERRAR YAHOO
    # ========================================================

    mail.logout()


    # ========================================================
    # QUITAR NPES DUPLICADOS
    # ========================================================

    npes_unicos = {}


    for item in lista_npes:

        npes_unicos[
            item["npe"]
        ] = item


    # ========================================================
    # CREAR TXT
    # ========================================================

    os.makedirs(
        CARPETA_BASE,
        exist_ok=True
    )


    ruta_txt = os.path.join(

        CARPETA_BASE,

        "NPE'S_AES.txt"
    )


    with open(

        ruta_txt,

        "w",

        encoding="utf-8"

    ) as archivo:


        archivo.write(
            "NPE'S AES / CAESS\n"
        )

        archivo.write(
            "==============================\n\n"
        )


        for item in npes_unicos.values():


            archivo.write(

                f"{item['propietario']} - "
                f"{item['npe']}\n"
            )


    print()
    print(
        "=========================================="
    )

    print(
        "PROCESO FINALIZADO"
    )

    print(
        "=========================================="
    )

    print(
        f"NPE encontrados: {len(npes_unicos)}"
    )

    print(
        f"Archivo: {ruta_txt}"
    )


# ============================================================
# INICIAR
# ============================================================

if __name__ == "__main__":

    procesar()