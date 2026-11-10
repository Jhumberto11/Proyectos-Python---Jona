import imaplib
import email
import os
import re
import csv
from datetime import datetime

from email.header import decode_header, make_header
from getpass import getpass
from bs4 import BeautifulSoup


# ============================================================
# CONFIGURACIÓN
# ============================================================

SERVIDOR = "imap.mail.yahoo.com"
PUERTO = 993

REMITENTE_DELSUR = "factura@delsur.com.sv"

CARPETA_BASE = "Recibos_DELSUR"


# ============================================================
# AQUÍ PONDREMOS A QUIÉN PERTENECE CADA CUENTA
# ============================================================

PROPIETARIOS = {
    "511829301": "ESTELA PORTILLO",
    "104857201": "COLOCHO",
    "101407801": "FLOR",
    "515511201": "JAZMIN",
    "101407701": "BLOQUE 1 - ABAJO",
    "513010401": "BODEGA BLOQUE 2 - ABAJO",
    "516123301" : "BOMBA ARRIBA"
}



# ============================================================
# DECODIFICAR ASUNTO
# ============================================================

def decodificar_asunto(asunto):

    if not asunto:
        return ""

    try:
        return str(make_header(decode_header(asunto)))
    except:
        return asunto


# ============================================================
# OBTENER TEXTO DEL CORREO
# ============================================================

def obtener_texto_correo(mensaje):

    texto_total = ""

    for parte in mensaje.walk():

        tipo = parte.get_content_type()

        # Ignorar adjuntos
        if parte.get_content_disposition() == "attachment":
            continue

        if tipo not in ["text/plain", "text/html"]:
            continue

        contenido = parte.get_payload(decode=True)

        if contenido is None:
            continue

        charset = parte.get_content_charset() or "utf-8"

        try:
            contenido = contenido.decode(
                charset,
                errors="ignore"
            )
        except:
            contenido = contenido.decode(
                "utf-8",
                errors="ignore"
            )

        # Si viene en HTML
        if tipo == "text/html":

            soup = BeautifulSoup(
                contenido,
                "html.parser"
            )

            contenido = soup.get_text(
                " ",
                strip=True
            )

        texto_total += " " + contenido


    # Limpiar espacios
    texto_total = texto_total.replace("\xa0", " ")

    texto_total = re.sub(
        r"\s+",
        " ",
        texto_total
    )

    return texto_total.strip()


# ============================================================
# EXTRAER DATOS DEL RECIBO
# ============================================================

def extraer_datos(texto):

    datos = {

        "nc": "",
        "cliente": "",
        "direccion": "",
        "total": "",
        "npe": "",
        "vencimiento": ""
    }


    # ------------------------------
    # NC
    # ------------------------------

    resultado = re.search(
        r"\bNC:\s*(\d+)",
        texto,
        re.IGNORECASE
    )

    if resultado:
        datos["nc"] = resultado.group(1)


    # ------------------------------
    # NOMBRE CLIENTE
    # ------------------------------

    resultado = re.search(
        r"Nombre del cliente:\s*(.*?)\s*Direcci[oó]n del suministro:",
        texto,
        re.IGNORECASE
    )

    if resultado:
        datos["cliente"] = resultado.group(1).strip()


    # ------------------------------
    # DIRECCIÓN
    # ------------------------------

    resultado = re.search(
        r"Direcci[oó]n del suministro:\s*(.*?)\s*Total a pagar:",
        texto,
        re.IGNORECASE
    )

    if resultado:
        datos["direccion"] = resultado.group(1).strip()


    # ------------------------------
    # TOTAL
    # ------------------------------

    resultado = re.search(
        r"Total a pagar:\s*\$?\s*([0-9,]+\.\d{2})",
        texto,
        re.IGNORECASE
    )

    if resultado:
        datos["total"] = resultado.group(1).replace(",", "")


    # ------------------------------
    # NPE
    # ------------------------------

    resultado = re.search(
        r"\bNPE:\s*(\d+)",
        texto,
        re.IGNORECASE
    )

    if resultado:
        datos["npe"] = resultado.group(1)


    # ------------------------------
    # FECHA VENCIMIENTO
    # ------------------------------

    resultado = re.search(
        r"Fecha de vencimiento:\s*(\d{2}/\d{2}/\d{4})",
        texto,
        re.IGNORECASE
    )

    if resultado:
        datos["vencimiento"] = resultado.group(1)


    return datos


# ============================================================
# GUARDAR PDF
# ============================================================

def guardar_pdf(mensaje, carpeta, nc):

    os.makedirs(
        carpeta,
        exist_ok=True
    )

    archivos_guardados = []


    for parte in mensaje.walk():

        nombre = parte.get_filename()

        if not nombre:
            continue

        nombre = decodificar_asunto(nombre)

        if not nombre.lower().endswith(".pdf"):
            continue


        contenido = parte.get_payload(decode=True)

        nombre_nuevo = f"Recibo_{nc}_{nombre}"

        ruta = os.path.join(
            carpeta,
            nombre_nuevo
        )


        with open(ruta, "wb") as archivo:

            archivo.write(
                contenido
            )


        archivos_guardados.append(ruta)


    return archivos_guardados




# ============================================================
# PROCESAR YAHOO
# ============================================================

def obtener_inicio_mes_imap():

    hoy = datetime.now()

    meses = [
        "Jan", "Feb", "Mar", "Apr",
        "May", "Jun", "Jul", "Aug",
        "Sep", "Oct", "Nov", "Dec"
    ]

    mes = meses[hoy.month - 1]

    return f"01-{mes}-{hoy.year}"

def obtener_mes():
    mes = datetime.now().month
    return f"{mes}"





def procesar():

    print()
    print("====================================")
    print(" CLASIFICADOR DE RECIBOS DELSUR")
    print("====================================")
    print()


    #correo = input(
    #    "Correo Yahoo: "
    #).strip()

    correo = "imosa13@yahoo.com"
    password = "lkojvljjevjcjtuh"

    

    #password = getpass(
     #   "Contraseña: "
    #)


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
    print()


    # Solo lectura
    mail.select(
        "INBOX",
        readonly=True
    )


    # Buscar únicamente correos DELSUR

    fecha_inicio = obtener_inicio_mes_imap()

    print(
        f"Buscando facturas desde: {fecha_inicio}"
    )

    estado, resultado = mail.search(
        None,
        'FROM',
        f'"{REMITENTE_DELSUR}"',
        'SINCE',
        fecha_inicio
    )


    if estado != "OK":

        print(
            "No fue posible buscar los correos."
        )

        return


    mensajes = resultado[0].split()


    print(
        f"Correos DELSUR encontrados: {len(mensajes)}"
    )


    resultados = []


    # ========================================================
    # RECORRER CORREOS
    # ========================================================

    lista_npes = []
    for numero in mensajes:


        estado, datos_email = mail.fetch(

            numero,

            "(BODY.PEEK[])"
        )


        if estado != "OK":
            continue


        contenido = datos_email[0][1]


        mensaje = email.message_from_bytes(
            contenido
        )


        asunto = decodificar_asunto(
            mensaje.get("Subject", "")
        )


        # Validar asunto DELSUR

        if not asunto.startswith(
            "DELSUR Factura Electrónica"
        ):
            continue


        texto = obtener_texto_correo(
            mensaje
        )


        datos = extraer_datos(
            texto
        )


        nc = datos["nc"]


        # Si no encontró NC
        if not nc:

            print()
            print(
                "⚠ Recibo sin número de cuenta"
            )

            print(
                asunto
            )

            continue


        # ====================================================
        # IDENTIFICAR PROPIETARIO
        # ====================================================



        propietario = PROPIETARIOS.get(
            nc,
            f"SIN_ASIGNAR_{nc}"
        )

        if datos["npe"]:

            lista_npes.append({
                "propietario": propietario,
                "npe": datos["npe"],
                "nc": nc
            })

        # ====================================================
        # CREAR CARPETA
        # ====================================================

        carpeta_persona = os.path.join(

            CARPETA_BASE,

            propietario
        )


        archivos = guardar_pdf(

            mensaje,

            carpeta_persona,

            nc
        )


        resultado_recibo = {

            "persona": propietario,

            "nc": datos["nc"],

            "cliente": datos["cliente"],

            "direccion": datos["direccion"],

            "total": datos["total"],

            "npe": datos["npe"],

            "vencimiento": datos["vencimiento"],

            "asunto": asunto
        }


        resultados.append(
            resultado_recibo
        )


        # ====================================================
        # MOSTRAR RESULTADO
        # ====================================================

        print()
        print(
            "----------------------------------------"
        )

        print(
            f"Persona:     {propietario}"
        )

        print(
            f"NC:          {datos['nc']}"
        )

        print(
            f"Cliente:     {datos['cliente']}"
        )

        print(
            f"Total:       ${datos['total']}"
        )

        print(
            f"NPE:         {datos['npe']}"
        )

        print(
            f"Vencimiento: {datos['vencimiento']}"
        )


        if archivos:

            print(
                f"PDF:         {archivos[0]}"
            )


    # ========================================================
    # GUARDAR CSV
    # ========================================================

    os.makedirs(
        CARPETA_BASE,
        exist_ok=True
    )


    ruta_csv = os.path.join(

        CARPETA_BASE,

        "resumen_recibos.csv"
    )


    with open(

        ruta_csv,

        "w",

        newline="",

        encoding="utf-8-sig"

    ) as archivo:


        columnas = [

            "persona",
            "nc",
            "cliente",
            "direccion",
            "total",
            "npe",
            "vencimiento",
            "asunto"

        ]


        escritor = csv.DictWriter(

            archivo,

            fieldnames=columnas
        )


        escritor.writeheader()

        escritor.writerows(
            resultados
        )


    mail.logout()


    print()
    print(
        "========================================"
    )

    print(
        "PROCESO FINALIZADO"
    )

    print(
        "========================================"
    )

    print(
        f"Recibos procesados: {len(resultados)}"
    )

    print(
        f"Resumen guardado en: {ruta_csv}"
    )

    # ========================================================
    # CREAR ARCHIVO DE NPES
    # ========================================================
    mes = obtener_mes()
    ruta_npes = os.path.join(
        CARPETA_BASE,
        f"{mes}-NPE'S.txt"
    )


    # Evitar NPE duplicados
    npes_unicos = {}

    for item in lista_npes:

        npe = item["npe"]

        npes_unicos[npe] = item


    with open(
        ruta_npes,
        "w",
        encoding="utf-8"
    ) as archivo:

        archivo.write("NPE'S DELSUR\n")
        archivo.write("==============================\n\n")

        for item in npes_unicos.values():

            archivo.write(
                f"{item['propietario']} - {item['npe']}\n"
            )
# ============================================================
# INICIO
# ============================================================

if __name__ == "__main__":

    procesar()