import re, requests, hashlib, operator, ftplib, sys, subprocess, os, copy
import time
import json, webbrowser, mysql.connector
from decimal import Decimal
from datetime import date, datetime, timedelta
from PyQt5 import uic, QtCore, QtGui
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from diccionario_sunat import *
import urllib.request
from PyQt5 import *

import pathlib
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

probandoFactElect=False
incialesElectronica=["F", "B", "T"]
FacturacionElectronica=["FACTURA","BOLETA",'NOTA DE CRÉDITO','NOTA DE DÉBITO']
TipComprobante={'FACTURA':'1','BOLETA':'2','NOTA DE CRÉDITO':'3','NOTA DE DÉBITO':'4'}
porcentaje_de_igv=18.00
rutaBase=os.getcwd()

db=mysql.connector.connect(host="67.23.254.35",user="multipla_admin",passwd="multiplay123",database="multipla_pruebas")

url = 'https://www.multiplay.com.pe/consultas/consulta-prueba.php'
rutaFacturacion="https://api.nubefact.com/api/v1/bad6db4a-3f1c-4f93-b364-bcd30093df5f"
tokenFacturacion="c3992ada327745ac9034af1af4ab879ff6119a5d8d1b46ab8404069fadee556a"

#-------------------------------- Funciones Generales ----------------------------------

def ejecutarSql(sql):
    print(sql)
    datos = {'accion':'ejecutar','sql': sql}
    x = requests.post(url, data = datos)
    if x.text!="":
        respuesta=x.json()
        if respuesta!=[]:
            print(respuesta)
    else:
        print("respuesta vacía")
    return respuesta

def ejecutarSqlDB(sql, values):
    try:
        db.reconnect()
        mycursor=db.cursor()
        mycursor.executemany(sql, values)
        db.commit()
        db.close()
        return {'respuesta':'correcto', 'resultado':str(len(values)) + " filas insertadas" }
    except Exception as e:
        print(e)
        return {'respuesta':'incorrecto', 'resultado':e}

def consultarSql(sql):
    datos = {'accion':'leer','sql': sql}
    x = requests.post(url, data=datos)
    respuesta=x.json()
    myresult=[]
    if respuesta!=[]:
        for datos in respuesta:
            contenido=[]
            for k,dato in datos.items():
                contenido.append(dato)
            myresult.append(contenido)
    return myresult

def cargarLogo(lb, codSoc):
    try:
        if codSoc == 'multiplay':
            codSoc = 'Mp_st'
        folderLogo = '''Logos/Logo'''+ codSoc +'.png'
        logoSoc = QPixmap(folderLogo)
        ratio = QtCore.Qt.KeepAspectRatio
        logoSoc = logoSoc.scaled(250, 35, ratio)
        lb.setPixmap(logoSoc)
    except:
        ""

def cargarIcono(obj, tipoIcono):
    try:
        iconos = {
        'erp': "organization",
        'banco':"banco",
        'grabar': "diskette",
        'modificar': "edit",
        'nuevo':"new_record",
        'direccion':"location",
        'salir': "logout",
        'buscar': "loupe",
        'compra':"purchasing",
        'usuario': "user",
        'darbaja': "x-button",
        'habilitar':"verify",
        'cargar': "sand-clock",
        'liberar': "liberar",
        'activar':"check",
        'cerrar':"close",
        'agregar_texto':'add_text',
        'consultar':"query",
        'con_texto':"with_text",
        'imprimir': "printer",
        'visualizar': "visualizar",
        'registrar': "clipboard",
        'condicion': "conditions",
        'depositar':"deposit",
        'importar':"data-transfer",
        'continuar':"right-arrow",
        'guia':"destino",
        'pdf':"pdf",
        'enviar':"send",
        'correo':"mail"}

        icono = iconos[tipoIcono]
        folderIcono = '''IconosLocales/'''+ icono +'.png'
        icon = QPixmap(folderIcono)
        if tipoIcono != 'erp':
            obj.setIcon(QIcon(icon))
        else:
            obj.setWindowIcon(QIcon(icon))
    except:
        ""

def mensajeDialogo(tipo, titulo, mensaje):
    try:
        msg = QMessageBox()
        cargarIcono(msg, 'erp')
        if tipo == 'error':
            msg.setIcon(QMessageBox.Critical)
        elif tipo == 'informacion':
            msg.setIcon(QMessageBox.Information)
        elif tipo == 'advertencia':
            msg.setIcon(QMessageBox.Warning)
        elif tipo == 'pregunta':
            msg.setIcon(QMessageBox.Question)
            msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        msg.setWindowTitle(titulo)
        msg.setText(mensaje)
        msg.setStyleSheet('''QMessageBox {background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgba(160, 160, 160, 255), stop:1 rgba(255, 255, 255, 255));} ''')
        valor = msg.exec_()
        if valor == 1024:
            valor = 'Ok'
        if valor == 16384:
            valor = 'Yes'
        if valor == 65536:
            valor = 'No'
        return valor
    except:
        ""

def convlist(sql):
    informacion=consultarSql(sql)
    lista = []
    for info in informacion:
        for elemento in info:
            lista.append(elemento)
    return lista

def insertarDatos(cb, Datos):
    cb.clear()
    for dato in Datos:
        cb.addItem(dato[0])

def buscarTabla(tw, texto, columnas):
    try:
        rango = range(tw.topLevelItemCount())
        palabras=re.sub(' +', ' ', texto).split(" ")
        patrones=[]
        for palabra in palabras:
            patrones.append(re.compile(palabra.upper()))
        if texto=="":
            for i in rango:
                tw.topLevelItem(i).setHidden(False)
        else:
            for i in rango:
                busqueda=True
                for j in columnas:
                    subBusqueda=False
                    for patron in patrones:
                        subBusqueda=subBusqueda or (patron.search(tw.topLevelItem(i).text(j).upper()) is None)
                    busqueda=busqueda and subBusqueda
                if busqueda:
                    tw.topLevelItem(i).setHidden(True)
                else:
                    tw.topLevelItem(i).setHidden(False)
    except Exception as e:
        mensajeDialogo("error", "buscarTabla", e)

def validarCorreo(lineEdit):
    correo = lineEdit.text()
    reg_exp = '^[(a-z0-9\_\-\.)]+@[(a-z0-9\_\-\.)]+\.[(a-z)]{2,15}$'
    if re.match(reg_exp, correo.lower()):
    	print("Correo correcto")
    else:
        lineEdit.clear()
        print("Correo incorrecto")

def validarNumero(lineEdit):
    valor = lineEdit.text()
    reg_exp = '^[1-9]\d*(,\d+)*(\.\d+)?$'
    if re.match(reg_exp, valor):
        print("Es un valor numérico")
    else:
        lineEdit.clear()
        print("No es un valor numérico")

def formatearFecha(fecha):
    if fecha=="":
        return ""
    fecha=fecha.split("-")
    fecha.reverse()
    return "-".join(fecha)

def formatearDecimal(str, nro):
    try:
        decimal = float(str)
        decimalRound = round(decimal,int(nro))
        cantDecimales = "{:,." + nro + "f}"
        decimalStr = cantDecimales.format(decimalRound)
        return decimalStr
    except:
        ""

def QDateToStrView(Qdate):
    a1=str(Qdate.date().year())
    m1=str(Qdate.date().month())
    d1=str(Qdate.date().day())

    if len(d1)==1:
        d1='0'+d1
    if len(m1)==1:
        m1='0'+m1
    # strFecha="%s-%s-%s" % (a1,m1,d1)
    strFecha="%s-%s-%s" % (d1,m1,a1)
    return strFecha

def crearCarpeta(nombre):
    if sys.platform == "win32":
        ruta=rutaBase[0:-len(rutaBase.split("\\")[-1])] + nombre + "\\"
    else:
        ruta=rutaBase[0:-len(rutaBase.split("/")[-1])] + nombre + "/"
    try:
        os.mkdir(ruta)
        print("Nueva carpeta creada")
    except OSError as e:
        print("La Carpeta ya existe")
    return ruta

def abrirArchivo(ruta):
    try:
        if sys.platform == "win32":
            os.startfile(ruta)
        else:
            opener ="open" if sys.platform == "darwin" else "xdg-open"
            subprocess.call([opener, ruta])
    except Exception as e:
        mensaje("Critico", "abrirArchivo", e)

def EnviarCorreo(CorreoRemitente,Archivo,Asunto,Cuerpo):

    remitente = 'no-responder@multiplay.com.pe'
    Clave="MPqap4ca8SBjFuN"
    destinatarios = [CorreoRemitente]
    asunto = Asunto
    cuerpo = Cuerpo
    ruta_adjunto = Archivo
    NombreArchivo=Archivo.split('\\')
    NombreArchivo=NombreArchivo[-1]
    nombre_adjunto = NombreArchivo

    # Creamos el objeto mensaje
    mensaje = MIMEMultipart()

    # Establecemos los atributos del mensaje
    mensaje['From'] = remitente
    mensaje['To'] = ", ".join(destinatarios)
    mensaje['Subject'] = asunto

    # Agregamos el cuerpo del mensaje como objeto MIME de tipo texto
    mensaje.attach(MIMEText(cuerpo, 'plain'))

    # Abrimos el archivo que vamos a adjuntar
    archivo_adjunto = open(ruta_adjunto, 'rb')

    # Creamos un objeto MIME base
    adjunto_MIME = MIMEBase('application', 'octet-stream')
    # Y le cargamos el archivo adjunto
    adjunto_MIME.set_payload((archivo_adjunto).read())
    # Codificamos el objeto en BASE64
    encoders.encode_base64(adjunto_MIME)
    # Agregamos una cabecera al objeto
    adjunto_MIME.add_header('Content-Disposition', "attachment; filename= %s" % nombre_adjunto)
    # Y finalmente lo agregamos al mensaje
    mensaje.attach(adjunto_MIME)

    # Creamos la conexión con el servidor
    #Correo gmail

    sesion_smtp = smtplib.SMTP('mail.multiplay.com.pe: 587')
    # sesion_smtp = smtplib.SMTP('smtp.live.com', 587)

    # Ciframos la conexión
    sesion_smtp.starttls()

    # Iniciamos sesión en el servidor
    sesion_smtp.login(remitente,Clave)

    # Convertimos el objeto mensaje a texto
    texto = mensaje.as_string().encode()

    # Enviamos el mensaje
    sesion_smtp.sendmail(remitente, destinatarios, texto)

    # Cerramos la conexión
    sesion_smtp.quit()
    # except Exception as e:
    #
    #     print("ERROR:", e)
    #     mensaje("Critico", "enviarCorreo", e)

def ValidarNumero(texto):
    if texto != "":
        texto=str(texto)
        texto=re.sub('[^0123456789.-]', "", texto, flags=re.IGNORECASE)
        if texto.count(".")>1:
            texto=texto[0:texto.find(".",texto.find(".")+1)] + texto[texto.find(".",texto.find(".")+1)+1:len(texto)+1]
        if texto.find("-")>0:
            texto=texto[0:-1]
    if texto=="":
        texto="0"
    return texto

def cargarCb(sql, cb): #ACTUALIZADO
    try:
        myresult=consultarSql(sql)
        temporal=cb.currentText()
        cb.clear()
        for dato in myresult:
            cb.addItem(str(dato[0]))
        cb.setCurrentText(temporal)
    except Exception as e:
        print("cargarCb: ",e)

def bloquearCb(cb):
    cb.setEnabled(False)
    cb.setStyleSheet("color: rgb(0,0,0);\n""background-color: rgb(255,255,255);")

def bloquearLe(le):
    le.setEnabled(False)
    le.setStyleSheet("color: rgb(0,0,0);\n""background-color: rgb(255,255,255);")

###############################################################################################################################################################

MONEDA_SINGULAR = 'DÓLAR AMERICANO'
MONEDA_PLURAL = 'DÓLARES AMERICANOS'

CENTIMOS_SINGULAR = 'CENTAVO'
CENTIMOS_PLURAL = 'CENTAVOS'

MAX_NUMERO = 999999999999

UNIDADES = (
    'cero',
    'uno',
    'dos',
    'tres',
    'cuatro',
    'cinco',
    'seis',
    'siete',
    'ocho',
    'nueve'
)

DECENAS = (
    'diez',
    'once',
    'doce',
    'trece',
    'catorce',
    'quince',
    'dieciseis',
    'diecisiete',
    'dieciocho',
    'diecinueve'
)

DIEZ_DIEZ = (
    'cero',
    'diez',
    'veinte',
    'treinta',
    'cuarenta',
    'cincuenta',
    'sesenta',
    'setenta',
    'ochenta',
    'noventa'
)

CIENTOS = (
    '_',
    'ciento',
    'doscientos',
    'trescientos',
    'cuatroscientos',
    'quinientos',
    'seiscientos',
    'setecientos',
    'ochocientos',
    'novecientos'
)


def StrToDate(FechaString):
    if FechaString == "None" or FechaString == None or FechaString == "":
        return date(2000,1,1)
    else:
        return date(int(FechaString.split("-")[0]), int(FechaString.split("-")[1]),int(FechaString.split("-")[2]))

def numero_a_letras(numero):
    numero_entero = int(numero)
    if numero_entero > MAX_NUMERO:
        raise OverflowError('Número demasiado alto')
    if numero_entero < 0:
        return 'menos %s' % numero_a_letras(abs(numero))
    letras_decimal = ''
    parte_decimal = int(round((abs(numero) - abs(numero_entero)) * 100))
    if parte_decimal > 9:
        letras_decimal = 'punto %s' % numero_a_letras(parte_decimal)
    elif parte_decimal > 0:
        letras_decimal = 'punto cero %s' % numero_a_letras(parte_decimal)
    if (numero_entero <= 99):
        resultado = leer_decenas(numero_entero)
    elif (numero_entero <= 999):
        resultado = leer_centenas(numero_entero)
    elif (numero_entero <= 999999):
        resultado = leer_miles(numero_entero)
    elif (numero_entero <= 999999999):
        resultado = leer_millones(numero_entero)
    else:
        resultado = leer_millardos(numero_entero)
    resultado = resultado.replace('uno mil', 'un mil')
    resultado = resultado.strip()
    resultado = resultado.replace(' _ ', ' ')
    resultado = resultado.replace('  ', ' ')
    if parte_decimal > 0:
        resultado = '%s %s' % (resultado, letras_decimal)
    return resultado

def numero_a_moneda(numero):
    numero_entero = int(numero)
    parte_decimal = int(round((abs(numero) - abs(numero_entero)) * 100))
    centimos = ''
    if parte_decimal == 1:
        centimos = CENTIMOS_SINGULAR
    else:
        centimos = CENTIMOS_PLURAL
    moneda = ''
    if numero_entero == 1:
        moneda = MONEDA_SINGULAR
    else:
        moneda = MONEDA_PLURAL
    letras = numero_a_letras(numero_entero)
    letras = letras.replace('uno', 'un')
    # letras_decimal = 'con %s %s' % (numero_a_letras(parte_decimal).replace('uno', 'un'), centimos)
    # letras = '%s %s %s' % (letras, moneda, letras_decimal)
    # return letras

    # Para céntimos 00/100
    if parte_decimal == 0:
        parte_decimal = '00'
    else:
        parte_decimal = str(parte_decimal)
    letras_parte_decimal = 'con '+ parte_decimal + '/100'
    letras = '%s %s %s' % (letras, letras_parte_decimal, moneda)
    return letras

def leer_decenas(numero):
    if numero < 10:
        return UNIDADES[numero]
    decena, unidad = divmod(numero, 10)
    if numero <= 19:
        resultado = DECENAS[unidad]
    elif 21 <= numero <= 29: #Corregido
        resultado = 'veinti%s' % UNIDADES[unidad]
    else:
        resultado = DIEZ_DIEZ[decena]
        if unidad > 0:
            resultado = '%s y %s' % (resultado, UNIDADES[unidad])
    return resultado

def leer_centenas(numero):
    centena, decena = divmod(numero, 100)
    if decena == 0 and centena == 1: #Corregido
        resultado = 'cien'
    else:
        resultado = CIENTOS[centena]
        if decena > 0:
            resultado = '%s %s' % (resultado, leer_decenas(decena))
    return resultado

def leer_miles(numero):
    millar, centena = divmod(numero, 1000)
    resultado = ''
    if (millar == 1):
        resultado = ''
    if (millar >= 2) and (millar <= 9):
        resultado = UNIDADES[millar]
    elif (millar >= 10) and (millar <= 99):
        resultado = leer_decenas(millar)
    elif (millar >= 100) and (millar <= 999):
        resultado = leer_centenas(millar)
    resultado = '%s mil' % resultado
    if centena > 0:
        resultado = '%s %s' % (resultado, leer_centenas(centena))
    return resultado

def leer_millones(numero):
    millon, millar = divmod(numero, 1000000)
    resultado = ''
    if (millon == 1):
        resultado = ' un millon '
    if (millon >= 2) and (millon <= 9):
        resultado = UNIDADES[millon]
    elif (millon >= 10) and (millon <= 99):
        resultado = leer_decenas(millon)
    elif (millon >= 100) and (millon <= 999):
        resultado = leer_centenas(millon)
    if millon > 1:
        resultado = '%s millones' % resultado
    if (millar > 0) and (millar <= 999):
        resultado = '%s %s' % (resultado, leer_centenas(millar))
    elif (millar >= 1000) and (millar <= 999999):
        resultado = '%s %s' % (resultado, leer_miles(millar))
    return resultado

def leer_millardos(numero):
    millardo, millon = divmod(numero, 1000000)
    return '%s millones %s' % (leer_miles(millardo), leer_millones(millon))

def formatoFechaTexto(fecha):
    meses = ("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
    dia = fecha.day
    mes = meses[fecha.month - 1]
    año = fecha.year
    fecha_texto = "{} de {} del {}".format(dia, mes, año)
    return fecha_texto

##################################################################################################################################################

def tipoDocumento(RUC, self):
    if len(RUC)==11:
        return dict_cliente_tipo_de_documento["RUC - REGISTRO ÚNICO DE CONTRIBUYENTE"]
    elif len(RUC)==8:
        return dict_cliente_tipo_de_documento["DNI - DOC. NACIONAL DE IDENTIDAD"]
    else:
        mensajeDialogo("informacion", "Tipo de Documento", "No se pudo identificar el tipo de documento")

def subirNubeFact(dataJson, GuiaT002, self):
    try:
        headers = {"Authorization" : tokenFacturacion, 'content-type': 'application/json'}
        r = requests.post(rutaFacturacion, data=json.dumps(dataJson), headers=headers)
        print(r, r.status_code)
        data=r.json()
        if r.status_code==200:
            print(data)
            enlace_del_pdf=data["enlace_del_pdf"]
            sunat_description=data["sunat_description"]
            sunat_note=data["sunat_note"]
            sunat_responsecode=data["sunat_responsecode"]
            sunat_soap_error=data["sunat_soap_error"]
            mensaje_error=[]
            if sunat_description!=None:
                mensaje_error.append(sunat_description)
            if sunat_note!=None:
                mensaje_error.append(sunat_note)
            if sunat_responsecode!=None:
                mensaje_error.append(sunat_responsecode)
            if sunat_soap_error!='' and sunat_soap_error!=None:
                mensaje_error.append(sunat_soap_error)
            if data["aceptada_por_sunat"]:
                mensajeDialogo("informacion", "Operación exitosa", sunat_description)
            else:
                if mensaje_error==[]:
                    mensajeDialogo("informacion", "Pendiente de validación Sunat", "El documento está en proceso de validación de Sunat\n\nRevisar el portal de NubeFact")
                else:
                    mensajeDialogo("informacion", "Documento rechazado", " | ".join(mensaje_error))
            webbrowser.open(enlace_del_pdf)

        elif r.status_code==400:
            mensajeDialogo("error", "Solicitud incorrecta", data["errors"])
            print(data)
        elif r.status_code==401:
            mensajeDialogo("error", "No autorizado", data["errors"])
            print(data)
        elif r.status_code==500:
            mensajeDialogo("error", "Error de servidor interno", data["errors"])
            print(data)
        else:
            mensajeDialogo("advertencia", "Error desconocido", "Error fuera de la lista")
            print(data)
        return r

    except Exception as e:
        mensajeDialogo("error", "subirNubeFact", e)
        print(e)
        return False

#--------------------------------PROGRAMA N° 1 - ERP_DATA_REF_P002----------------------------------


#--------------------------------PROGRAMA N° 2 - ERP_ORG_P004----------------------------------

def actualizar(tw,sql):
    tw.clearContents()
    informacion=consultarSql(sql)
    flags = (QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
    row=0
    for fila in informacion:
        col=0
        for i in fila:
            if i!=fila[3]:
                item=QTableWidgetItem(i)
                item.setFlags(flags)
                if tw.rowCount()<=row:
                    tw.insertRow(tw.rowCount())
                tw.setItem(row, col, item)
                col += 1
        if fila[3]=="1":
            C4=QTableWidgetItem(str("ACTIVO"))
            C4.setFlags(flags)
            if tw.rowCount()<=row:
                tw.insertRow(tw.rowCount())
            tw.setItem(row, 4, C4)
        else:
            C4=QTableWidgetItem(str("BAJA"))
            C4.setFlags(flags)
            font = QtGui.QFont()
            font.setPointSize(12)
            C4.setFont(font)
            brush = QtGui.QBrush(QtGui.QColor(255, 0, 0))
            brush.setStyle(QtCore.Qt.SolidPattern)
            C4.setForeground(brush)
            if tw.rowCount()<=row:
                tw.insertRow(tw.rowCount())
            tw.setItem(row, 4, C4)
        row+=1

def NombreUbigeo(CodPais,CodDepartamento,CodProvincia,CodDistrito,TablaUbigeo):
    NombreUBI={}
    try:
        NombreUBI["Pais"]=TablaUbigeo[CodPais+"-0-0-0"]
    except:
        NombreUBI["Pais"]=""
    if NombreUBI["Pais"]=="Peru":
        try:
            if CodDepartamento!="0":
                NombreUBI["Departamento"]=TablaUbigeo[CodPais+"-"+CodDepartamento+"-0-0"]
            else:
                NombreUBI["Departamento"]=""
        except:
            NombreUBI["Departamento"]=""
        try:
            if CodProvincia!="0" and CodDepartamento!="0":
                NombreUBI["Provincia"]=TablaUbigeo[CodPais+"-"+CodDepartamento+"-"+CodProvincia+"-0"]
            else:
                NombreUBI["Provincia"]=""
        except:
            NombreUBI["Provincia"]=""
        try:
            if CodDistrito!="0" and CodProvincia!="0" and CodDepartamento!="0":
                NombreUBI["Distrito"]=TablaUbigeo[CodPais+"-"+CodDepartamento+"-"+CodProvincia+"-"+CodDistrito]
            else:
                NombreUBI["Distrito"]=""
        except:
            NombreUBI["Distrito"]=""
    else:
        try:
            if CodDepartamento!="0":
                NombreUBI["Departamento"]=TablaUbigeo[CodPais+"-"+CodDepartamento+"-0-0"]
            else:
                NombreUBI["Departamento"]=""
        except:
            NombreUBI["Departamento"]=""
        NombreUBI["Provincia"]=""
        NombreUBI["Distrito"]=""
    return NombreUBI

def TablaUbigeo(sql):
    ubicacion=consultarSql(sql)
    tablaUbigeo={}
    for item in ubicacion:
        tablaUbigeo[item[0]+"-"+item[1]+"-"+item[2]+"-"+item[3]]=item[4]
    return tablaUbigeo


#--------------------------------PROGRAMA N° 3 - ERP_PROV_P001----------------------------------

def actualizarInter(self,tw,sql,Tipo_Inter,dicTipoInter):
    tw.clearContents()
    informacion=consultarSql(sql)
    if informacion!=[]:
        rows=tw.rowCount()
        for r in range(rows):
            tw.removeRow(1)
        flags = (QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
        row=0
        for fila in informacion:
            if tw.rowCount()<=row:
                tw.insertRow(tw.rowCount())
            cb = QComboBox(tw)
            tw.setCellWidget(row, 0, cb)
            insertarDatos(cb,Tipo_Inter)
            cb.setEditable(True)
            for k,v in dicTipoInter.items():
                if fila[0]==k:
                    tw.cellWidget(row, 0).setEditText(v)
            font = QtGui.QFont()
            font.setPointSize(12)
            cb.setFont(font)
            # font = QFont('Times', 12)
            le = cb.lineEdit()
            le.setFont(font)
            tw.resizeColumnToContents(0)
            tw.cellWidget(row, 0).setEnabled(False)
            tw.cellWidget(row, 0).setStyleSheet("color: rgb(0,0,0);")

            col=1
            for i in fila:
                if i!=fila[0] and i!=fila[7]:
                    item=QTableWidgetItem(i)
                    item.setFlags(flags)
                    if tw.rowCount()<=row:
                        tw.insertRow(tw.rowCount())
                    tw.setItem(row, col, item)
                    col += 1

            if fila[7]=="1":
                C7=QTableWidgetItem(str("ACTIVO"))
                C7.setFlags(flags)
                if tw.rowCount()<=row:
                    tw.insertRow(tw.rowCount())
                tw.setItem(row, 7, C7)
            else:
                C7=QTableWidgetItem(str("BAJA"))
                C7.setFlags(flags)
                font = QtGui.QFont()
                font.setPointSize(12)
                C7.setFont(font)
                brush = QtGui.QBrush(QtGui.QColor(255, 0, 0))
                brush.setStyle(QtCore.Qt.SolidPattern)
                C7.setForeground(brush)
                if tw.rowCount()<=row:
                    tw.insertRow(tw.rowCount())
                tw.setItem(row, 7, C7)
            row+=1
        rowPosi=tw.rowCount()
        tw.insertRow(rowPosi)
        item=QTableWidgetItem()
        item.setFlags(flags)
        tw.setItem(rowPosi,7, item)

        cb0 = QComboBox(tw)
        tw.setCellWidget(rowPosi, 0, cb0)
        insertarDatos(cb0,Tipo_Inter)
        cb0.setCurrentIndex(-1)
        font = QtGui.QFont()
        font.setPointSize(12)
        cb0.setFont(font)
        tw.resizeColumnToContents(0)

    else:
        cb = QComboBox(tw)
        tw.setCellWidget(0, 0, cb)
        insertarDatos(cb,Tipo_Inter)
        cb.setCurrentIndex(-1)
        font = QtGui.QFont()
        font.setPointSize(12)
        cb.setFont(font)
        tw.resizeColumnToContents(0)

        flags = (QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
        item=QTableWidgetItem()
        item.setFlags(flags)
        tw.setItem(0,7, item)

def actualizarBan(self,tw,sql,datos,TCta,dicbanco,banco,dicmoneda,mon):
    tw.clearContents()

    informacion=consultarSql(sql)

    if informacion!=[]:
        rows=tw.rowCount()
        for r in range(rows):
            tw.removeRow(1)
        flags = (QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
        row=0
        for fila in informacion:
            C0=QTableWidgetItem(fila[0])
            C0.setFlags(flags)
            if tw.rowCount()<=row:
                tw.insertRow(tw.rowCount())
            tw.setItem(row,0, C0)
            tw.resizeColumnToContents(0)

            if tw.rowCount()<=row:
                tw.insertRow(tw.rowCount())
            cb1 = QComboBox(tw)
            # cb1.setEditable(True)
            tw.setCellWidget(row, 1, cb1)
            llenarPais(datos,cb1)
            cb1.setCurrentIndex(int(fila[1])-1)
            # tw.cellWidget(row, 1).setEditText(fila[1])
            font = QtGui.QFont()
            font.setPointSize(12)
            cb1.setFont(font)
            tw.resizeColumnToContents(1)
            tw.cellWidget(row, 1).setEnabled(False)
            tw.cellWidget(row, 1).setStyleSheet("color: rgb(0,0,0);")
            cb1.activated.connect(self.cargarDepartamento)

            if tw.rowCount()<=row:
                tw.insertRow(tw.rowCount())
            cb2 = QComboBox(tw)
            # cb2.setEditable(True)
            tw.setCellWidget(row, 2, cb2)
            Paisx=fila[1]
            llenarDepartamento(datos,cb2,Paisx)
            cb2.setCurrentText(fila[2])
            # tw.cellWidget(row, 2).setEditText(fila[2])
            font = QtGui.QFont()
            font.setPointSize(12)
            cb2.setFont(font)
            tw.resizeColumnToContents(2)
            tw.cellWidget(row, 2).setEnabled(False)
            tw.cellWidget(row, 2).setStyleSheet("color: rgb(0,0,0);")

            if tw.rowCount()<=row:
                tw.insertRow(tw.rowCount())
            cb3 = QComboBox(tw)
            # cb3.setEditable(True)
            tw.setCellWidget(row, 3, cb3)
            insertarDatos(cb3,banco)
            cb3.setCurrentIndex(int(fila[3])-1)
            # tw.cellWidget(row, 3).setEditText(fila[3])
            font = QtGui.QFont()
            font.setPointSize(12)
            cb3.setFont(font)
            tw.resizeColumnToContents(3)
            tw.cellWidget(row, 3).setEnabled(False)
            tw.cellWidget(row, 3).setStyleSheet("color: rgb(0,0,0);")

            if tw.rowCount()<=row:
                tw.insertRow(tw.rowCount())
            cb4 = QComboBox(tw)
            # cb4.setEditable(True)
            tw.setCellWidget(row, 4, cb4)
            for k,v in TCta.items():
                cb4.addItem(k)
            if fila[4]=="CA":
                cb4.setCurrentIndex(0)
            elif fila[4]=="CC":
                cb4.setCurrentIndex(1)
            font = QtGui.QFont()
            font.setPointSize(12)
            cb4.setFont(font)
            tw.resizeColumnToContents(4)
            tw.cellWidget(row,4).setEnabled(False)
            tw.cellWidget(row,4).setStyleSheet("color: rgb(0,0,0);")

            C5=QTableWidgetItem(fila[5])
            C5.setFlags(flags)
            if tw.rowCount()<=row:
                tw.insertRow(tw.rowCount())
            tw.setItem(row,5, C5)

            if tw.rowCount()<=row:
                tw.insertRow(tw.rowCount())
            cb6 = QComboBox(tw)
            # cb6.setEditable(True)
            tw.setCellWidget(row, 6, cb6)
            insertarDatos(cb6,mon)

            cb6.setCurrentIndex(int(fila[6])-1)
            # tw.cellWidget(row, 8).setEditText(fila[6])
            font = QtGui.QFont()
            font.setPointSize(12)
            cb6.setFont(font)
            tw.resizeColumnToContents(6)
            tw.cellWidget(row, 6).setEnabled(False)
            tw.cellWidget(row, 6).setStyleSheet("color: rgb(0,0,0);")

            C7=QTableWidgetItem(fila[7])
            C7.setFlags(flags)
            if tw.rowCount()<=row:
                tw.insertRow(tw.rowCount())
            tw.setItem(row,7, C7)

            if fila[8]=="1":
                C8=QTableWidgetItem(str("ACTIVO"))
                C8.setFlags(flags)
                if tw.rowCount()<=row:
                    tw.insertRow(tw.rowCount())
                tw.setItem(row,8, C8)
            else:
                C8=QTableWidgetItem(str("BAJA"))
                C8.setFlags(flags)
                font = QtGui.QFont()
                font.setPointSize(12)
                C8.setFont(font)
                brush = QtGui.QBrush(QtGui.QColor(255, 0, 0))
                brush.setStyle(QtCore.Qt.SolidPattern)
                C8.setForeground(brush)
                if tw.rowCount()<=row:
                    tw.insertRow(tw.rowCount())
                tw.setItem(row, 8, C8)
            row+=1

        rowPosi=tw.rowCount()
        tw.insertRow(rowPosi)

        CB0 = QComboBox(tw)
        tw.setCellWidget(rowPosi, 1, CB0)
        for k,v in datos.items():
            codigo=k.split("-")
            if "-".join(codigo[1:])=="0-0-0":
                CB0.addItem(v)
        CB0.setCurrentIndex(-1)
        # CB0.setStyleSheet("background-color: rgb(255,255,255);")
        font = QtGui.QFont()
        font.setPointSize(12)
        CB0.setFont(font)
        tw.resizeColumnToContents(1)
        CB0.activated.connect(self.cargarDepartamento)

        #creacion combo departamento...
        CB1 = QComboBox(tw)
        tw.setCellWidget(rowPosi, 2, CB1)
        CB1.setCurrentIndex(-1)
        # CB1.setStyleSheet("background-color: rgb(255,255,255);")
        font = QtGui.QFont()
        font.setPointSize(12)
        CB1.setFont(font)
        tw.resizeColumnToContents(2)

        #creacion combo tipo de banco...
        CB2 = QComboBox(tw)
        tw.setCellWidget(rowPosi, 3, CB2)
        insertarDatos(CB2,banco)
        CB2.setCurrentIndex(-1)
        # CB2.setStyleSheet("background-color: rgb(255,255,255);")
        font = QtGui.QFont()
        font.setPointSize(12)
        CB2.setFont(font)
        tw.resizeColumnToContents(3)

        #creacion combo tipo de cuenta...
        CB3 = QComboBox(tw)
        tw.setCellWidget(rowPosi, 4, CB3)
        for k,v in TCta.items():
            CB3.addItem(k)
        CB3.setCurrentIndex(-1)
        # CB3.setStyleSheet("background-color: rgb(255,255,255);")
        font = QtGui.QFont()
        font.setPointSize(12)
        CB3.setFont(font)
        tw.resizeColumnToContents(4)

        CB4 = QComboBox(tw)
        tw.setCellWidget(rowPosi, 6, CB4)
        insertarDatos(CB4,mon)
        CB4.setCurrentIndex(-1)
        # CB4.setStyleSheet("background-color: rgb(255,255,255);")
        font = QtGui.QFont()
        font.setPointSize(12)
        CB4.setFont(font)
        tw.resizeColumnToContents(6)

        Nro=QTableWidgetItem(str(rowPosi+1))
        Nro.setFlags(flags)
        tw.setItem(rowPosi,0, Nro)

        item=QTableWidgetItem()
        item.setFlags(flags)
        tw.setItem(rowPosi,8, item)
        tw.resizeColumnToContents(0)

    else:
        cb0 = QComboBox(tw)
        tw.setCellWidget(0, 1, cb0)
        for k,v in datos.items():
            codigo=k.split("-")
            if "-".join(codigo[1:])=="0-0-0":
                cb0.addItem(v)
        cb0.setCurrentIndex(-1)
        # cb0.setStyleSheet("background-color: rgb(255,255,255);")
        font = QtGui.QFont()
        font.setPointSize(12)
        cb0.setFont(font)
        tw.resizeColumnToContents(1)
        cb0.activated.connect(self.cargarDepartamento)

        #creacion combo departamento...
        cb1 = QComboBox(tw)
        tw.setCellWidget(0, 2, cb1)
        cb1.setCurrentIndex(-1)
        # cb1.setStyleSheet("background-color: rgb(255,255,255);")
        font = QtGui.QFont()
        font.setPointSize(12)
        cb1.setFont(font)
        tw.resizeColumnToContents(2)

        #creacion combo tipo de banco...
        cb2 = QComboBox(tw)
        tw.setCellWidget(0, 3, cb2)
        insertarDatos(cb2,banco)
        cb2.setCurrentIndex(-1)
        # cb2.setStyleSheet("background-color: rgb(255,255,255);")
        font = QtGui.QFont()
        font.setPointSize(12)
        cb2.setFont(font)
        tw.resizeColumnToContents(3)

        #creacion combo tipo de cuenta...
        cb3 = QComboBox(tw)
        tw.setCellWidget(0, 4, cb3)
        for k,v in TCta.items():
            cb3.addItem(k)
        cb3.setCurrentIndex(-1)
        # cb3.setStyleSheet("background-color: rgb(255,255,255);")
        font = QtGui.QFont()
        font.setPointSize(12)
        cb3.setFont(font)
        tw.resizeColumnToContents(4)

        cb4 = QComboBox(tw)
        tw.setCellWidget(0, 6, cb4)
        insertarDatos(cb4,mon)
        cb4.setCurrentIndex(-1)
        # cb4.setStyleSheet("background-color: rgb(255,255,255);")
        font = QtGui.QFont()
        font.setPointSize(12)
        cb4.setFont(font)
        tw.resizeColumnToContents(6)

        flags = (QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)

        Nro=QTableWidgetItem("1")
        Nro.setFlags(flags)
        tw.setItem(0,0, Nro)

        item=QTableWidgetItem()
        item.setFlags(flags)
        tw.setItem(0,8, item)
        tw.resizeColumnToContents(0)

def actualizarComp(tw,sql):
    tw.clearContents()
    informacion=consultarSql(sql)
    if informacion!=[]:
        rows=tw.rowCount()
        for r in range(rows):
            tw.removeRow(1)
        flags = (QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
        row=0
        for fila in informacion:
            tw.resizeColumnToContents(0)
            tw.resizeColumnToContents(5)
            col=0
            for i in fila:
                item=QTableWidgetItem(i)
                item.setFlags(flags)
                # item.setTextAlignment(QtCore.Qt.AlignHCenter)
                if tw.rowCount()<=row:
                    tw.insertRow(tw.rowCount())
                tw.setItem(row, col, item)
                # tw.resizeColumnToContents(col)
                col += 1
            row+=1
        rowPosi=tw.rowCount()
        tw.insertRow(rowPosi)
        Nro=QTableWidgetItem(str(rowPosi+1))
        Nro.setFlags(flags)
        tw.setItem(rowPosi,0, Nro)
        tw.resizeColumnToContents(0)
        tw.resizeColumnToContents(5)
    else:
        flags = (QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
        Nro=QTableWidgetItem("1")
        Nro.setFlags(flags)
        tw.setItem(0,0, Nro)
        tw.resizeColumnToContents(0)
        tw.resizeColumnToContents(5)

def llenarPais(TablaUbigeo,cbPais):  #Codigo pais va 0
    for ubigeo,nombre in TablaUbigeo.items():
        au=ubigeo.find("-")
        bu=ubigeo.find("-",au+1)
        cu=ubigeo.find("-",bu+1)
        if ubigeo[au+1:]=="0-0-0":
            cbPais.addItem(nombre)

def llenarDepartamento(TablaUbigeo,cbDepartamento,codigoPais):  #Codigo pais va 0
    for ubigeo,nombre in TablaUbigeo.items():
        au=ubigeo.find("-")
        bu=ubigeo.find("-",au+1)
        cu=ubigeo.find("-",bu+1)
        if ubigeo[:au]==codigoPais and ubigeo[au+1:]!="0-0-0" and ubigeo[bu+1:]=="0-0":
            cbDepartamento.addItem(nombre)

def llenarDep(TablaUbigeo,cbDepartamento,codigoPais):  #Codigo pais va 0
    cbDepartamento.clear()
    for ubigeo,nombre in TablaUbigeo.items():
        au=ubigeo.find("-")
        bu=ubigeo.find("-",au+1)
        cu=ubigeo.find("-",bu+1)
        if ubigeo[:au]==codigoPais and ubigeo[au+1:]!="0-0-0" and ubigeo[bu+1:]=="0-0":
            cbDepartamento.addItem(ubigeo[au+1:bu]+" - "+nombre)
            cbDepartamento.setCurrentIndex(-1)

def verificarTIP(tw):
    TIP=[]
    A=tw.rowCount()
    B=tw.currentRow()
    for fila in range(A-(A-B)):
        item=tw.cellWidget(fila, 0).currentText()
        TIP.append(item)
    return TIP

###############################################################

def consultaRuc(mostrar, RUC):
    try:
        print("Primer método")
        respuesta=consultaRucPeruApis(mostrar, RUC) #Razon Social y Nombre Comercial
        if respuesta!=False:
            return respuesta
    except Exception as e:
        print(e)
        # mensajeDialogo("error", "Primer método", "%s\nRUC:%s" % (e, RUC))

    try:
        print("Segundo método")
        respuesta=consultaRucApiSPeru(mostrar, RUC) #Menos Nombre Comercial
        if respuesta!=False:
            return respuesta
    except Exception as e:
        print(e)
        # mensajeDialogo("error", "Segundo método", "%s\nRUC:%s" % (e, RUC))

    try:
        print("Tercer método")
        respuesta=consultaRucApiPeruDev(mostrar, RUC) #Direccion completa falta Distrito Departamento Provincia
        if respuesta!=False:
            return respuesta
    except Exception as e:
        print(e)
        # mensajeDialogo("error", "Tercer método", "%s\nRUC:%s" % (e, RUC))

    try:
        print("Cuarto método")
        respuesta=consultaRucMigo(mostrar, RUC)
        if respuesta!=False:
            return respuesta
    except Exception as e:
        print(e)
        # mensajeDialogo("error", "Cuarto método", "%s\nRUC:%s" % (e, RUC))

def consultaRucPeruApis(mostrar, RUC): #Razon Social y Nombre Comercial Check
    continuar=True
    if len(RUC)<8: return False
    try:
        tokenRUC="QLDGyWjlG4EZU11WJ0gcVO7wcBC5mzPR8CWEBx09N0qXNk7GvPerz6Y8WPXK" #Token Personal
        headers = {"Authorization" : "Bearer %s" % tokenRUC, 'Content-Type':'application/json'}

        if len(RUC)==8:
            url="https://api.peruapis.com/v1/dni"
            r = requests.post(url, data=json.dumps({"document":"%s" % RUC}), headers=headers)
            print(r.text)
            if not r.ok:
                if mostrar:
                    print("1")
                    print("No se encontró el RUC")
                    # mensajeDialogo("informacion", "Consulta RUC", "No se encontró el RUC " + str(RUC))
                return False
            data=r.json()
            Ruc=data["data"]["dni"]
            RazonSocial=data["data"]["fullname"]
            Codigo=data["data"]["verification_code"]
            NombreComercial="-"
            Direccion=""
            Ubigeo=""
            Descripcion=""
            Estado=""

            url="https://api.peruapis.com/v1/ruc"
            r = requests.post(url, data=json.dumps({"document":"10%s%s" % (RUC, Codigo)}), headers=headers)
            print(r.text)

            if not r.ok:
                continuar=False

            if continuar:
                data=r.json()
                Ruc=data["data"]["ruc"]
                RazonSocial=data["data"]["name"]
                NombreComercial=data["data"]["commercial_name"]
                if NombreComercial==None:
                    NombreComercial="-"
                Distrito=data["data"]["district"]
                Provincia=data["data"]["province"]
                Departamento=data["data"]["region"]
                if Distrito==None:
                    Descripcion=""
                    Direccion="-"
                else:
                    Descripcion="%s-%s-%s" % (Distrito, Provincia, Departamento)
                    Direccion=data["data"]["address"] + " " + Descripcion
                Ubigeo=data["data"]["location"]
                Estado=data["data"]["status"]
        else:
            url="https://api.peruapis.com/v1/ruc"
            r = requests.post(url, data=json.dumps({"document":"%s" % RUC}), headers=headers)
            print(r.text)
            if not r.ok:
                if mostrar:
                    print("2")
                    print("No se encontró el RUC")
                    # mensajeDialogo("informacion", "Consulta RUC", "No se encontró el RUC " + str(RUC))
                return False
            data=r.json()
            Ruc=data["data"]["ruc"]
            RazonSocial=data["data"]["name"]
            NombreComercial=data["data"]["commercial_name"]
            if NombreComercial==None:
                NombreComercial="-"
            Distrito=data["data"]["district"]
            Provincia=data["data"]["province"]
            Departamento=data["data"]["region"]
            if Distrito==None:
                Descripcion=""
                Direccion="-"
            else:
                Descripcion="%s-%s-%s" % (Distrito, Provincia, Departamento)
                Direccion=data["data"]["address"] + " " + Descripcion
            Ubigeo=data["data"]["location"]
            Estado=data["data"]["status"]
    except Exception as e:
        if mostrar:
            print(e)
            # mensajeDialogo("error", "consultaRucPeruApis", "%s\nRUC:%s" % (e, RUC))
        return False
    return [Ruc, RazonSocial, NombreComercial, Direccion, Estado, Ubigeo, Descripcion]

def consultaRucApiSPeru(mostrar, RUC): #Nombre Comercial Check
    continuar=True
    if len(RUC)<8: return False
    try:
        token="eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJlbWFpbCI6InJvbm55Lm9hLjE0QGdtYWlsLmNvbSJ9.8rOVuY6r1UK1GJlimZjxSH3u8ONw83kQoBV85V6mQeE" # Token Personal
        if len(RUC)==8:
            tipo="dni"
        else:
            tipo="ruc"
        url="https://dniruc.apisperu.com/api/v1/%s/%s?token=%s" % (tipo, RUC, token)
        r = requests.get(url)
        print(r.text)
        if not r.ok:
            if "504 Gateway Time-out" in r.text:
                print("Error en la página")
                a=float("hola")
            if "Ha excedido" in r.text:
                print("Exceso de consultas")
                a=float("hola")
            print("3")
            print("No se encontró el RUC")
            # mensajeDialogo("informacion", "Consulta RUC", "No se encontró el RUC " + str(RUC))
            return False
        data=r.json()
        if tipo=="ruc":
            Ruc=data["ruc"]
            RazonSocial=data["razonSocial"]
            NombreComercial=data["nombreComercial"]
            if data["direccion"]!="-":
                Departamento=data["departamento"]
                Provincia=data["provincia"]
                Distrito=data["distrito"]
                Direccion="%s %s - %s - %s" % (data["direccion"], Departamento, Provincia, Distrito)
                respuesta=consultarSql(sqlUbigeoRuc % (Departamento, Provincia, Distrito))
                if respuesta==[]:
                    Ubigeo=""
                    Descripcion=""
                else:
                    Ubigeo=respuesta[0][0]
                    Descripcion=respuesta[0][1]
            else:
                Direccion="-"
                Ubigeo=""
                Descripcion=""
            Estado=data["estado"]
        else:
            ruc="10" + data["dni"] + data["codVerifica"]
            url="https://dniruc.apisperu.com/api/v1/ruc/%s?token=%s" % (ruc, token)
            r = requests.get(url)
            if not r.ok:
                Ruc=data["dni"]
                RazonSocial="%s %s %s" % (data["apellidoPaterno"], data["apellidoMaterno"], data["nombres"])
                NombreComercial="-"
                Direccion="-"
                Ubigeo=""
                Descripcion=""
                Estado="SIN RUC"
            else:
                data=r.json()
                Ruc=data["ruc"]
                RazonSocial=data["razonSocial"]
                NombreComercial=data["nombreComercial"]
                Direccion="-"
                Ubigeo=""
                Descripcion=""
                Estado=data["estado"]
    except Exception as e:
        if mostrar:
            print(e)
            # mensajeDialogo("error", "consultaRucApiSPeru", "%s\nRUC:%s" % (e, RUC))
        return False
    return [Ruc, RazonSocial, NombreComercial, Direccion, Estado, Ubigeo, Descripcion]

def consultaRucApiPeruDev(mostrar, RUC): #Direccion completa Check
    continuar=True
    if len(RUC)<8: return False
    try:
        tokenRUC="dea2d1928a82aaaac3dd17c37d9647f2de56cc835d77fb11d3404a9133efa0c3" #Token Personal
        headers = {"Authorization" : "Bearer %s" % tokenRUC, 'Content-Type': 'application/json'}

        if len(RUC)==8:
            url="https://apiperu.dev/api/dni/%s" % RUC
            r = requests.get(url, headers=headers)
            print(r.text)
            if not r.ok:
                print("4")
                print("No se encontró el RUC")
                # mensajeDialogo("informacion", "Consulta RUC", "No se encontró el RUC " + str(RUC))
                return False
            data=r.json()
            Ruc=data["data"]["ruc"]
            RazonSocial=data["data"]["nombre_completo"]
            Codigo=data["data"]["codigo_verificacion"]
            NombreComercial=""
            Direccion=""
            Ubigeo=""
            Descripcion=""
            Estado=""

            url="https://apiperu.dev/api/ruc/10%s%s" % (RUC, Codigo)
            r = requests.get(url, headers=headers)
            print(r.text)
            if not r.ok:
                continuar=False
            if continuar:
                data=r.json()
                Ruc=data["data"]["ruc"]
                RazonSocial=data["data"]["nombre_o_razon_social"]
                NombreComercial=""
                if "direccion_completa" in data["data"]:
                    Direccion=data["data"]["direccion_completa"]
                else:
                    Direccion=""
                Ubigeo=data["data"]["ubigeo"][2]
                Descripcion=""
                Estado=data["data"]["estado"]
        else:
            url="https://apiperu.dev/api/ruc/%s" % RUC
            r = requests.get(url, headers=headers)
            print(r.text)
            if not r.ok:
                print("5")
                print("No se encontró el RUC")
                # mensajeDialogo("informacion", "Consulta RUC", "No se encontró el RUC " + str(RUC))
                return False
            data=r.json()
            Ruc=data["data"]["ruc"]
            RazonSocial=data["data"]["nombre_o_razon_social"]
            NombreComercial=""
            if "direccion_completa" in data["data"]:
                Direccion=data["data"]["direccion_completa"]
            else:
                Direccion=""
            Ubigeo=data["data"]["ubigeo"][2]
            Descripcion=""
            Estado=data["data"]["estado"]

    except Exception as e:
        if mostrar:
            print(e)
            # mensajeDialogo("error", "consultaRucApiPeruDev", "%s\nRUC:%s" % (e, RUC))
        return False
    return [Ruc, RazonSocial, NombreComercial, Direccion, Estado, Ubigeo, Descripcion]

def consultaRucMigo(mostrar, RUC):
    if len(RUC)<8: return False
    try:
        tokenRUC="jRdhFs3PVRqow7h9Aiev7JjkX2gtELMFXb7yfhbBaCVmqGgtPlSGYWTxt6wG" #Api Migo token Personal 2021/07/14
        headers = {'Content-Type':'application/json'}

        if len(RUC)==8:
            url="https://api.migo.pe/api/v1/dni"
            data = '{"token":"%s", "dni":"%s"}' % (tokenRUC, RUC)
            print(data, url)
            r = requests.post(url, data=data, headers=headers)
            print(r.text)
            if not r.ok:
                if mostrar:
                    print("6")
                    print("No se encontró el RUC ")
                    # mensajeDialogo("informacion", "Consulta RUC", "No se encontró el RUC " + str(RUC))
                return False
            data=r.json()
            if data["success"]:
                Ruc=data["dni"]
                RazonSocial=data["nombre"]
                Codigo="-"
                NombreComercial="-"
                Direccion="-"
                Ubigeo=""
                Descripcion=""
                Estado=""
            else:
                if mostrar:
                    print("7")
                    print("No se encontró el RUC ")
                    # mensajeDialogo("informacion", "Consulta RUC", "No se encontró el RUC " + str(RUC))
                return False
        else:
            url="https://api.migo.pe/api/v1/ruc"
            data = '{"token":"%s", "ruc":"%s"}' % (tokenRUC, RUC)
            print(data, url)
            r = requests.post(url, data=data, headers=headers)
            print(r.text)
            if not r.ok:
                if mostrar:
                    mensajeDialogo("informacion", "Consulta RUC", "No se encontró el RUC " + str(RUC))
                return False
            data=r.json()
            if data["success"]:
                Ruc=data["ruc"]
                RazonSocial=data["nombre_o_razon_social"]
                NombreComercial="-"
                Distrito=data["distrito"]
                Provincia=data["provincia"]
                Departamento=data["departamento"]
                if Distrito==None:
                    Descripcion=""
                    Direccion="-"
                else:
                    Descripcion="%s-%s-%s" % (Distrito, Provincia, Departamento)
                    Direccion=data["direccion_simple"] + " " + Descripcion
                Ubigeo=data["ubigeo"]
                Estado=data["estado_del_contribuyente"]
            else:
                if mostrar:
                    print("No se encontró el RUC ")
                    # mensajeDialogo("informacion", "Consulta RUC", "No se encontró el RUC " + str(RUC))
                return False

    except Exception as e:
        if mostrar:
            print(e)
            # mensajeDialogo("error", "consultaRucMigo", "%s\nRUC:%s" % (e, RUC))
        return False
    return [Ruc, RazonSocial, NombreComercial, Direccion, Estado, Ubigeo, Descripcion]

#################################################################################################################

#--------------------------------PROGRAMA N° 4 - ERP_REQ_P002----------------------------------

def Cargar(self,tw,sql,Inicio,Final,Fec_Inicial,Fec_Final,Cod_Soc,Año):
    tw.clearContents()
    informacion=consultarSql(sql)
    if informacion!=[]:
        rows=tw.rowCount()
        for r in range(rows):
            tw.removeRow(1)
        flags = (QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
        row=0
        for fila in informacion:
            fila[1]=formatearFecha(fila[1])
            fila[2]=formatearFecha(fila[2])
            sqlMonto="SELECT SUM(Cant_Mat_Serv*Precio_ref) FROM TAB_SOLP_002_Detalle_Solicitud_Pedido WHERE Cod_Soc='%s' AND Año='%s' AND Nro_Solp='%s'"%(Cod_Soc,Año,fila[0])
            Monto=convlist(sqlMonto)
            mont=formatearDecimal(Monto[0],'3')
            fila.insert(6,mont)
            col=0
            for i in fila:
                item=QTableWidgetItem(i)
                item.setFlags(flags)
                if tw.rowCount()<=row:
                    tw.insertRow(tw.rowCount())
                tw.setItem(row,col, item)
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                tw.resizeColumnToContents(3)
                tw.resizeColumnToContents(4)
                col += 1

            pb = QPushButton("Consultar",tw)
            tw.setCellWidget(row, 6, pb)
            cargarIcono(pb,'consultar')
            font = QtGui.QFont()
            font.setPointSize(11)
            font.setBold(True)
            pb.setFont(font)
            pb.setStyleSheet("background-color: rgb(255, 213, 79);")
            pb.clicked.connect(self.Consultar)
            tw.resizeColumnToContents(6)

            row+=1
    else:
        mensajeDialogo("error", "Error","No se encontraron solicitudes de pedido en este rango")
        Inicio.setCurrentIndex(-1)
        Final.setCurrentIndex(-1)
        Fec_Inicial.clear()
        Fec_Final.clear()

def actualizarSOLP(self,tw,sql,Estado_Doc,Cod_Soc,NroSOLP,Año):
    tw.clearContents()
    informacion=consultarSql(sql)
    if informacion!=[]:
        rows=tw.rowCount()
        for r in range(rows):
            tw.removeRow(1)
        flags = (QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
        row=0
        for fila in informacion:
            fila[6]=formatearDecimal(fila[6],'3')
            fila[7]=formatearDecimal(fila[7],'2')
            fila[8]=formatearFecha(fila[8])
            for k,v in Estado_Doc.items():
                if fila[0]==k:
                    C0=QTableWidgetItem(v)
                    C0.setFlags(flags)
                    C0.setTextAlignment(QtCore.Qt.AlignCenter)
                    if v=='Anulado':
                        brush = QtGui.QBrush(QtGui.QColor(255, 0, 0))
                        brush.setStyle(QtCore.Qt.SolidPattern)
                        C0.setForeground(brush)
                    elif v=='Aprobado':
                        brush = QtGui.QBrush(QtGui.QColor(0, 0, 127))
                        brush.setStyle(QtCore.Qt.SolidPattern)
                        C0.setForeground(brush)
                    if tw.rowCount()<=row:
                        tw.insertRow(tw.rowCount())
                    tw.setItem(row, 0, C0)
                    tw.resizeColumnToContents(0)

            C1=QTableWidgetItem(fila[1])
            C1.setFlags(flags)
            C1.setTextAlignment(QtCore.Qt.AlignCenter)
            brush = QtGui.QBrush(QtGui.QColor(85, 255, 255))
            brush.setStyle(QtCore.Qt.SolidPattern)
            C1.setBackground(brush)
            if tw.rowCount()<=row:
                tw.insertRow(tw.rowCount())
            tw.resizeColumnToContents(1)
            tw.setItem(row, 1, C1)

            if tw.rowCount()<=row:
                tw.insertRow(tw.rowCount())
            cb0=QComboBox(tw)
            tw.setCellWidget(row,2,cb0)
            tw.cellWidget(row,2).addItem(fila[2])
            font = QtGui.QFont()
            font.setPointSize(12)
            cb0.setFont(font)
            tw.resizeColumnToContents(2)
            tw.cellWidget(row,2).setEnabled(False)

            material=[fila[3],fila[4],fila[5],fila[6],fila[7],fila[8],fila[9]]
            col=3
            for i in material:
                item=QTableWidgetItem(i)
                item.setFlags(flags)
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                if tw.rowCount()<=row:
                    tw.insertRow(tw.rowCount())
                tw.setItem(row, col, item)
                tw.resizeColumnToContents(col)
                col+=1

            combo=[fila[10],fila[11],fila[12],fila[13]]
            c=10
            for i in combo:
                if tw.rowCount()<=row:
                    tw.insertRow(tw.rowCount())
                cb1=QComboBox(tw)
                tw.setCellWidget(row,c,cb1)
                tw.cellWidget(row,c).addItem(i)
                font = QtGui.QFont()
                font.setPointSize(12)
                cb1.setFont(font)
                tw.resizeColumnToContents(c)
                tw.cellWidget(row,c).setEnabled(False)
                c+=1

            sqlTexto="SELECT Texto FROM TAB_SOC_019_Texto_Proceso WHERE Cod_Soc='%s'AND Año='%s' AND Tipo_Proceso='1'AND Nro_Doc='%s'AND Item_Doc='%s'"%(Cod_Soc,Año,NroSOLP,fila[1])
            texto=consultarSql(sqlTexto)

            btTexto=QPushButton(tw)
            if tw.rowCount()<=row:
                tw.insertRow(tw.rowCount())
            tw.setCellWidget(row, 14, btTexto)

            if texto!=[]:
                cargarIcono(btTexto,'con_texto')
            elif texto==[]:
                cargarIcono(btTexto,'agregar_texto')

            tw.resizeColumnToContents(14)
            font = QtGui.QFont()
            font.setPointSize(11)
            font.setBold(True)
            btTexto.setFont(font)
            btTexto.setStyleSheet("background-color: rgb(255, 213, 79);")
            btTexto.clicked.connect(self.TextoPosicion)
            row+=1

def actualizarboton(self,tw,Cod_Soc,Año,Numero_Solp,Item_Solp,row):
    sqlTexto="SELECT Texto FROM TAB_SOC_019_Texto_Proceso WHERE Cod_Soc='%s'AND Año='%s' AND Tipo_Proceso='1'AND Nro_Doc='%s'AND Item_Doc='%s'"%(Cod_Soc,Año,Numero_Solp,Item_Solp)
    texto=consultarSql(sqlTexto)

    btTexto=QPushButton(tw)
    tw.setCellWidget(row, 14, btTexto)

    if texto!=[]:
        cargarIcono(btTexto,'con_texto')
    elif texto==[]:
        cargarIcono(btTexto,'agregar_texto')

    tw.resizeColumnToContents(14)
    font = QtGui.QFont()
    font.setPointSize(11)
    font.setBold(True)
    btTexto.setFont(font)
    btTexto.setStyleSheet("background-color: rgb(255, 213, 79);")
    btTexto.clicked.connect(self.TextoPosicion)

#--------------------------------PROGRAMA N° 5 - ERP_COMP_P001----------------------------------

def CargarCotApro(self,tw,sql):
    tw.clearContents()
    informacion=consultarSql(sql)
    if informacion!=[]:
        rows=tw.rowCount()
        for r in range(rows):
            tw.removeRow(1)
        flags = (QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
        row=0
        for fila in informacion:
            if fila[8]==None:
                fila[3]=formatearDecimal(fila[3],'3')
                fila[4]=formatearFecha(fila[4])
                fila[6]=formatearFecha(fila[6])
                fila[8]="Ganador"
                col=0
                for i in fila:
                    if i=='0':
                        i='-'
                    item=QTableWidgetItem(i)
                    item.setFlags(flags)
                    item.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter)
                    if tw.rowCount()<=row:
                        tw.insertRow(tw.rowCount())
                    tw.setItem(row, col, item)
                    tw.resizeColumnToContents(col)
                    col += 1
                row+=1
    else:
        mensajeDialogo("informacion","Informacion","No se encontraron cotizaciones aprobadas en este rango, verifique")

def CargarPC(self,tw,sql,dicTipPed,dicEstado):
    tw.clearContents()
    informacion=consultarSql(sql)
    if informacion!=[]:
        rows=tw.rowCount()
        for r in range(rows):
            tw.removeRow(1)
        flags = (QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
        row=0
        for fila in informacion:
            fila[3]=dicTipPed[fila[3]]
            fila[5]=formatearFecha(fila[5])
            fila[6]=dicEstado[fila[6]]
            # if fila[8]==None:
            #     fila[8]="GANADOR"
            col=0
            for i in fila:
                if i=='0':
                    i='-'
                item=QTableWidgetItem(i)
                item.setFlags(flags)
                item.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter)
                if tw.rowCount()<=row:
                    tw.insertRow(tw.rowCount())
                tw.setItem(row, col, item)
                tw.resizeColumnToContents(col)
                col += 1
            row+=1
    else:
        mensajeDialogo("informacion","Informacion","No se encontraron cotizaciones aprobadas en este rango, verifique")

def CargarPedComp(self,tw,sql,Cod_Soc,Año,Nro_Doc):
    tw.clearContents()
    informacion=consultarSql(sql)
    if informacion!=[]:
        rows=tw.rowCount()
        for r in range(rows):
            tw.removeRow(1)
        flags = (QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
        item_ped=1
        row=0
        for fila in informacion:
            fila.insert(0,str(item_ped))
            fila[6]=formatearDecimal(fila[6],'3')
            fila[5]=formatearDecimal(fila[5],'3')
            col=0
            for i in fila:
                item=QTableWidgetItem(i)
                item.setFlags(flags)
                item.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter)
                if tw.rowCount()<=row:
                    tw.insertRow(tw.rowCount())
                tw.setItem(row, col, item)
                tw.resizeColumnToContents(col)

                col += 1

            sqlTexto="SELECT Texto FROM TAB_SOC_019_Texto_Proceso WHERE Cod_Soc='%s'AND Año='%s' AND Tipo_Proceso='3'AND Nro_Doc='%s'AND Item_Doc='%s'"%(Cod_Soc,Año,Nro_Doc,fila[0])
            texto=consultarSql(sqlTexto)

            btTexto=QPushButton(tw)
            if tw.rowCount()<=row:
                tw.insertRow(tw.rowCount())
            tw.setCellWidget(row, 11, btTexto)

            if texto!=[]:
                cargarIcono(btTexto,'con_texto')
            elif texto==[]:
                cargarIcono(btTexto,'agregar_texto')

            # tw.resizeColumnToContents(11)
            font = QtGui.QFont()
            font.setPointSize(11)
            font.setBold(True)
            btTexto.setFont(font)
            btTexto.setStyleSheet("background-color: rgb(255, 213, 79);")
            btTexto.clicked.connect(self.TextoPosicion)
            row+=1
            item_ped=int(item_ped)+1
    # else:
    #     QMessageBox.critical(self, "Informacion","No se encontraron datos", QMessageBox.Ok)

def actualizarboton2(self,tw,Cod_Soc,Año,Nro_Doc,Item_Solp,row):
    sqlTexto="SELECT Texto FROM TAB_SOC_019_Texto_Proceso WHERE Cod_Soc='%s'AND Año='%s' AND Tipo_Proceso='3'AND Nro_Doc='%s'AND Item_Doc='%s'"%(Cod_Soc,Año,Nro_Doc,Item_Solp)
    texto=consultarSql(sqlTexto)

    btTexto=QPushButton(tw)
    tw.setCellWidget(row, 11, btTexto)

    if texto!=[]:
        cargarIcono(btTexto,'con_texto')
    elif texto==[]:
        cargarIcono(btTexto,'agregar_texto')

    # tw.resizeColumnToContents(11)
    font = QtGui.QFont()
    font.setPointSize(11)
    font.setBold(True)
    btTexto.setFont(font)
    btTexto.setStyleSheet("background-color: rgb(255, 213, 79);")
    btTexto.clicked.connect(self.TextoPosicion)

def cargarInter(tw,sql):
    tw.clearContents()
    informacion=consultarSql(sql)
    if informacion!=[]:
        rows=tw.rowCount()
        for r in range(rows):
            tw.removeRow(r)
        flags = (QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
        row=0
        for fila in informacion:
            col=0
            for i in fila:
                if i!=fila[7]:
                    item=QTableWidgetItem(i)
                    item.setFlags(flags)
                    item.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter)
                    if tw.rowCount()<=row:
                        tw.insertRow(tw.rowCount())
                    tw.setItem(row, col, item)
                    col += 1

            if fila[7]=="1":
                C7=QTableWidgetItem(str("ACTIVO"))
                C7.setFlags(flags)
                C7.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter)
                if tw.rowCount()<=row:
                    tw.insertRow(tw.rowCount())
                tw.setItem(row, 7, C7)
            row+=1
    else:
        mensajeDialogo("informacion","Informacion","No se encontraron interlocutores registrados, verifique")

def condPos(self,tw,sql,Cond_Comp,Cond_Comp1,Cond_Comp2,Cond_Comp3,dicCond_Comp,Precio,Valor,Moneda,Tipo_Pedido):
    informacion=consultarSql(sql)
    flags = (QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
    if informacion!=[]:
        row=0
        for fila in informacion:
            for k,v in dicCond_Comp.items():
                if fila[0]==v:
                    condcomp=k
            if int(condcomp)<6:
                if tw.rowCount()-3<=row:
                    tw.insertRow(tw.rowCount()-3)
                cb = QComboBox(tw)
                tw.setCellWidget(row, 0, cb)
                insertarDatos(cb,Cond_Comp)
                tw.cellWidget(row, 0).setCurrentText(fila[0])
                font = QtGui.QFont()
                font.setPointSize(12)
                cb.setFont(font)
                tw.resizeColumnToContents(0)
                tw.cellWidget(row, 0).setEnabled(False)

                col=1
                for i in fila:
                    if i=='0.00':
                        i=""
                    if i!=fila[0]:
                        celda=QTableWidgetItem(i)
                        celda.setFlags(flags)
                        celda.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter)
                        if tw.rowCount()-3<=row:
                            tw.insertRow(tw.rowCount()-3)
                        tw.setItem(row, col, celda)
                        col += 1
                row+=1

            else:
                row+=1
                if tw.rowCount()-1<=row:
                    tw.insertRow(tw.rowCount()-1)
                cb = QComboBox(tw)
                tw.setCellWidget(row, 0, cb)
                insertarDatos(cb,Cond_Comp)
                tw.cellWidget(row, 0).setCurrentText(fila[0])
                font = QtGui.QFont()
                font.setPointSize(12)
                cb.setFont(font)
                tw.resizeColumnToContents(0)
                tw.cellWidget(row, 0).setEnabled(False)

                col=1
                for i in fila:
                    if i=='0.00':
                        i=""
                    if i!=fila[0]:
                        celda=QTableWidgetItem(i)
                        celda.setFlags(flags)
                        celda.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter)
                        if tw.rowCount()-1<=row:
                            tw.insertRow(tw.rowCount()-1)
                        tw.setItem(row, col, celda)
                        col += 1

    else:

        cb0 = QComboBox(tw)
        tw.setCellWidget(0, 0, cb0)
        insertarDatos(cb0,Cond_Comp1)
        cb0.setCurrentIndex(0)
        font = QtGui.QFont()
        font.setPointSize(12)
        cb0.setFont(font)
        tw.resizeColumnToContents(0)

        item1=QTableWidgetItem(Precio)
        item1.setFlags(flags)
        item1.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter)
        tw.setItem(0, 2, item1)

        item2=QTableWidgetItem(Valor)
        item2.setFlags(flags)
        item2.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter)
        tw.setItem(0, 3, item2)

        item3=QTableWidgetItem(Moneda)
        item3.setFlags(flags)
        item3.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter)
        tw.setItem(0, 4, item3)

        item22=QTableWidgetItem(Valor)
        item22.setFlags(flags)
        brush = QtGui.QBrush(QtGui.QColor(255, 213, 79))
        brush.setStyle(QtCore.Qt.SolidPattern)
        item22.setBackground(brush)
        item22.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter)
        tw.setItem(1, 3, item22)

        item33=QTableWidgetItem(Moneda)
        item33.setFlags(flags)
        brush = QtGui.QBrush(QtGui.QColor(255, 213, 79))
        brush.setStyle(QtCore.Qt.SolidPattern)
        item33.setBackground(brush)
        item33.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter)
        tw.setItem(1, 4, item33)

        cb1 = QComboBox(tw)
        tw.setCellWidget(2, 0, cb1)
        if Tipo_Pedido=='Importaciones':
            insertarDatos(cb1,Cond_Comp2)
            cb1.setCurrentIndex(-1)
            font = QtGui.QFont()
            font.setPointSize(12)
            cb1.setFont(font)
            tw.resizeColumnToContents(0)
            cb1.activated.connect(self.Condicion2)
        else:
            insertarDatos(cb1,Cond_Comp3)
            cb1.setCurrentIndex(-1)
            font = QtGui.QFont()
            font.setPointSize(12)
            cb1.setFont(font)
            tw.resizeColumnToContents(0)
            cb1.activated.connect(self.Condicion3)

def buscarTablaPC(self,tbw):
    rango = range(tbw.rowCount())
    try:
        descrip = self.lePalabra.text()
    except:
        descrip = ''

    patrones_descrip = re.sub(' +',' ', descrip).split(" ")
    list_descrip = []
    for palabra in patrones_descrip:
        list_descrip.append(re.compile(palabra.upper()))

    patron_codmat = re.compile(self.leMaterial.text().upper())
    fechaInicial = formatearFecha(self.leInicial.text())
    fechaFinal = formatearFecha(self.leFinal.text())

    if self.lePalabra.text() == "" and self.leMaterial.text() == "" and self.leInicial.text() == "" and self.leFinal.text() == "":
        for i in rango:
            tbw.setRowHidden(i,False)
    else: # ^ $
        for i in rango:
            if (patron_codmat.search(tbw.item(i,7).text().upper()) is None):
                tbw.setRowHidden(i,True)
            elif (formatearFecha(tbw.item(i,5).text()) > fechaFinal and self.leFinal.text() != ''):
                tbw.setRowHidden(i,True)
            elif (formatearFecha(tbw.item(i,5).text()) < fechaInicial and self.leInicial.text() != ''):
                tbw.setRowHidden(i,True)
            elif list_descrip != [re.compile('')]:
                ocultar=False
                for patron_descrip in list_descrip:
                    ocultar=ocultar or ((patron_descrip.search(tbw.item(i,0).text().upper()) is None) and (patron_descrip.search(tbw.item(i,1).text().upper()) is None) and (patron_descrip.search(tbw.item(i,3).text().upper()) is None) and (patron_descrip.search(tbw.item(i,4).text().upper()) is None))
                if ocultar:
                    tbw.setRowHidden(i,True)
                else:
                    tbw.setRowHidden(i,False)
            else:
                tbw.setRowHidden(i,False)

#--------------------------------PROGRAMA N° 6 - ERP_CALIDAD_P001----------------------------------

#--------------------------------PROGRAMA N° 7 - ERP_ACTIVOS_P002----------------------------------

def CargarActivo(sql,tw,self):
    tw.clearContents()
    rows=tw.rowCount()
    for r in range(rows):
        tw.removeRow(1)
    informacion=consultarSql(sql)
    if informacion!=[]:
        flags = (QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
        row=0
        for fila in informacion:
            if fila[4]=='1':
                fila[4]='ACTIVO'
            if fila[4]=='2':
                fila[4]='BAJA'
            fila[3]=formatearDecimal(fila[3],'2')
            col=0
            for i in fila:
                item=QTableWidgetItem(i)
                item.setFlags(flags)
                if i=='BAJA':
                    font = QtGui.QFont()
                    font.setPointSize(12)
                    item.setFont(font)
                    brush = QtGui.QBrush(QtGui.QColor(255, 0, 0))
                    brush.setStyle(QtCore.Qt.SolidPattern)
                    item.setForeground(brush)
                item.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter)
                if tw.rowCount()<=row:
                    tw.insertRow(tw.rowCount())
                tw.setItem(row, col, item)
                tw.resizeColumnToContents(1)
                # tw.resizeColumnToContents(5)
                col += 1
            row+=1
    else:
        mensajeDialogo("informacion", "Información", "No se encontraron Activos Fijos")

#--------------------------------PROGRAMA N° 8 - ERP_ALM_P027----------------------------------

def CargarPedComCli(sql,tbw,estregalm,self):
    tbw.clearContents()
    rows=tbw.rowCount()
    for r in range(rows):
        tbw.removeRow(1)
    data=consultarSql(sql)
    informacion= []
    for info in data:
        if info not in informacion:
            informacion.append(info)
    if informacion!=[]:
        flags = (QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
        row=0
        for fila in informacion:
            fila[3]=formatearFecha(fila[3])
            fila[4]=estregalm[fila[4]]
            col=0
            for i in fila:
                item=QTableWidgetItem(i)
                item.setFlags(flags)
                item.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter)
                if tbw.rowCount()<=row:
                    tbw.insertRow(tbw.rowCount())
                tbw.setItem(row, col, item)
                if len(fila[1])>25:
                    tbw.resizeColumnToContents(1)
                else:
                    tbw.setColumnWidth( 1, 265)
                col += 1
            row+=1
    else:
        mensajeDialogo("informacion", "Información", "No se encontraron registros")

def CargarDetalle(self, tbw, sql, dict_serie, dict_lote):
    tbw.clearContents()
    informacion=consultarSql(sql)
    if informacion!=[]:
        rows=tbw.rowCount()
        for r in range(rows):
            tbw.removeRow(1)
        flags = (QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
        row=0
        for fila in informacion:
            # Cant_Restante=float(fila[4])-float(fila[5])
            fila[4]=formatearDecimal(fila[4],'3')
            fila[5]=formatearDecimal(fila[5],'3')
            fila[6]=formatearDecimal(fila[6],'3')

            col=0
            for i in fila:
                if col == 1: ## Código de material
                    serie = dict_serie[i]
                    btSerie=QPushButton(tbw)
                    if tbw.rowCount()<=row:
                        tbw.insertRow(tbw.rowCount())
                    tbw.setCellWidget(row, col+7, btSerie)

                    if serie == '1': ## Si control de series
                        cargarIcono(btSerie,'activar')
                        btSerie.clicked.connect(self.SeriePosicion)
                    else:
                        cargarIcono(btSerie,'cerrar')

                    # tbw.resizeColumnToContents(11)
                    font = QtGui.QFont()
                    font.setPointSize(11)
                    font.setBold(True)
                    btSerie.setFont(font)
                    btSerie.setStyleSheet("background-color: rgb(255, 213, 79);")

                    lote = dict_lote[i]
                    if lote == '1': ## Si maneja Lote
                        lote = ''
                        info2 = QTableWidgetItem(lote)
                        tbw.setItem(row, col+6, info2)
                    else:
                        lote = "---"
                        info2 = QTableWidgetItem(lote)
                        info2.setFlags(flags)
                        info2.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                        tbw.setItem(row, col+6, info2)

                if col == 5:
                    if i == '0.000': ## Saldo por entregar
                        i = ""
                    info = QTableWidgetItem(item)
                    info.setFlags(flags)
                    info.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                    tbw.setItem(row, col, info)
                else:
                    item=QTableWidgetItem(i)
                    item.setFlags(flags)
                    item.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter)
                    if tbw.rowCount()<=row:
                        tbw.insertRow(tbw.rowCount())
                    tbw.setItem(row, col, item)
                    tbw.resizeColumnToContents(col)

                col += 1
            row+=1

def CargarCot_Guia(sql,tw,self):
    tw.clearContents()
    informacion=consultarSql(sql)
    if informacion!=[]:
        rows=tw.rowCount()
        for r in range(rows):
            tw.removeRow(1)
        flags = (QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
        row=0
        for fila in informacion:
            fila[5]=formatearDecimal(fila[5],'3')
            fila[6]=formatearDecimal(fila[6],'2')
            col=0
            for i in fila:
                item=QTableWidgetItem(i)
                item.setFlags(flags)
                if tw.rowCount()<=row:
                    tw.insertRow(tw.rowCount())
                tw.setItem(row,col, item)
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                tw.resizeColumnToContents(col)
                col += 1
            row+=1
    else:
        mensajeDialogo("informacion", "Información","No se encontraron registros")


def CargarFact_Guia(sql,tw,self):
    tw.clearContents()
    informacion=consultarSql(sql)
    if informacion!=[]:
        rows=tw.rowCount()
        for r in range(rows):
            tw.removeRow(1)
        flags = (QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
        row=0
        for fila in informacion:
            fila[5]=formatearDecimal(fila[5],'3')
            fila[6]=formatearDecimal(fila[6],'2')
            col=0
            for i in fila:
                item=QTableWidgetItem(i)
                item.setFlags(flags)
                if tw.rowCount()<=row:
                    tw.insertRow(tw.rowCount())
                tw.setItem(row,col, item)
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                tw.resizeColumnToContents(col)
                col += 1
            row+=1
    else:
        mensajeDialogo("informacion", "Información","No se encontraron registros")

def generarGuia(Cod_Soc,tipo_de_comprobante,serie,numero,cliente_tipo_de_documento,ruc,razonSocial,direccion,correo,fechaEmision,observaciones,motivo_de_traslado,peso_bruto_total,numero_de_bultos,tipo_de_transporte,fecha_de_inicio_de_traslado,transportista_documento_tipo,transportista_documento_numero,transportista_denominacion,transportista_placa_numero,conductor_documento_tipo,conductor_documento_numero,conductor_denominacion,punto_de_partida_ubigeo,punto_de_partida_direccion, punto_de_llegada_ubigeo, punto_de_llegada_direccion, tbwCotizacion_Cliente, nombreArchivo, self):

    if Cod_Soc!="1000":
        mensajeDialogo("informacion", "Facturación Electrónica", "Función válida para Multicable")
        return
    if self.leURL.text()=="":
        resultado=mensajeDialogo("pregunta", "Generar Documento", "¿Seguro que desea enviar a Sunat el documento?\n\nUna vez enviado no se podrá modificar el documento")
        if resultado=='Yes':
            try:
                data = {}
                data['operacion'] = "generar_guia"
                data['tipo_de_comprobante'] = dict_tipo_de_comprobante[tipo_de_comprobante]
                data["serie"] = serie
                data["numero"] = numero
                data["cliente_tipo_de_documento"] = cliente_tipo_de_documento
                data["cliente_numero_de_documento"] = ruc
                data["cliente_denominacion"] = razonSocial
                data["cliente_direccion"] = direccion
                correos=correo.split(" / ")
                if len(correos)==3:
                    correo1=correos[0]
                    correo2=correos[1]
                    correo3=correos[2]
                elif len(correos)==2:
                    correo1=correos[0]
                    correo2=correos[1]
                    correo3=""
                elif len(correos)==1:
                    correo1=correos[0]
                    correo2=""
                    correo3=""
                else:
                    correo1=""
                    correo2=""
                    correo3=""
                data["cliente_email"] = correo1
                data["cliente_email_1"] = correo2
                data["cliente_email_2"] = correo3
                data["fecha_de_emision"] = fechaEmision
                data["observaciones"] = observaciones
                data["motivo_de_traslado"] = motivo_de_traslado
                data["peso_bruto_total"] = peso_bruto_total
                data["numero_de_bultos"] = numero_de_bultos
                data["tipo_de_transporte"] = tipo_de_transporte
                data["fecha_de_inicio_de_traslado"] = fecha_de_inicio_de_traslado
                data["transportista_documento_tipo"] = transportista_documento_tipo
                data["transportista_documento_numero"] = transportista_documento_numero
                data["transportista_denominacion"] = transportista_denominacion
                data["transportista_placa_numero"] = transportista_placa_numero
                data["conductor_documento_tipo"] = conductor_documento_tipo
                data["conductor_documento_numero"] = conductor_documento_numero
                data["conductor_denominacion"] = conductor_denominacion
                data["punto_de_partida_ubigeo"] = punto_de_partida_ubigeo
                data["punto_de_partida_direccion"] = punto_de_partida_direccion
                data["punto_de_llegada_ubigeo"] = punto_de_llegada_ubigeo
                data["punto_de_llegada_direccion"] = punto_de_llegada_direccion
                data["enviar_automaticamente_a_la_sunat"] = True
                data["enviar_automaticamente_al_cliente"] = False if correo=="" else True
                data["codigo_unico"] = ""
                data["formato_de_pdf"] = ""
                data["items"]=[]

                for row in range(tbwCotizacion_Cliente.rowCount()):

                    Cod_Mat=tbwCotizacion_Cliente.item(row,1).text()
                    Descripcion=tbwCotizacion_Cliente.item(row,2).text()
                    Marca=tbwCotizacion_Cliente.item(row,3).text()
                    Unidad=tbwCotizacion_Cliente.item(row,4).text()
                    Cantidad=tbwCotizacion_Cliente.item(row,5).text()

                    NomMat= Descripcion + (" / MARCA: " + Marca if Marca!="" else "")
                    item={}
                    item['unidad_de_medida'] = Unidad
                    item["codigo"] = Cod_Mat
                    item["descripcion"] = NomMat
                    item["cantidad"] = Cantidad
                    print(item)

                    data['items'].append(item)

                crearCarpeta("%s JSON" % tipo_de_comprobante)
                if sys.platform == "win32":
                    ruta=rutaBase[0:-len(rutaBase.split("\\")[-1])]
                    with open('%s JSON\\%s.json' % (ruta + tipo_de_comprobante, nombreArchivo), 'w') as file:
                        json.dump(data, file, indent=4)
                else:
                    ruta=rutaBase[0:-len(rutaBase.split("/")[-1])]
                    with open('%s JSON/%s.json' % (ruta + tipo_de_comprobante, nombreArchivo), 'w') as file:
                        json.dump(data, file, indent=4)

                if serie=="T002":
                    respuesta=subirNubeFact(data, True, self)
                else:
                    respuesta=subirNubeFact(data, False, self)

                if respuesta.status_code==200:
                    respuesta=respuesta.json()
                    if respuesta["aceptada_por_sunat"]:
                        enlace=respuesta["enlace"]
                    else:
                        sunat_description=respuesta["sunat_description"]
                        sunat_note=respuesta["sunat_note"]
                        sunat_responsecode=respuesta["sunat_responsecode"]
                        sunat_soap_error=respuesta["sunat_soap_error"]
                        mensaje_error=[]
                        if sunat_description!=None:
                            mensaje_error.append(sunat_description)
                        if sunat_note!=None:
                            mensaje_error.append(sunat_note)
                        if sunat_responsecode!=None:
                            mensaje_error.append(sunat_responsecode)
                        if sunat_soap_error!='':
                            mensaje_error.append(sunat_soap_error)

                        if mensaje_error==[]:
                            enlace="Pendiente de validación Sunat"
                        else:
                            enlace="Documento rechazado"
                    self.leURL.setText(enlace)
                    # sql="UPDATE TAB_ALM_001_Mov_Almacenes SET URL='%s' WHERE Cod_Emp='%s'AND Tipo_Comprobante='%s' AND Serie='%s' AND Nro_Facturacion='%s'" % (enlace,Cod_Soc,dict_tipo_de_comprobante[tipo_de_comprobante], serie, numero)
                    # respuesta=ejecutarSql(sql)
                    # if respuesta["respuesta"]=="incorrecto":
                    #     mensaje("Peligro", "Error", str(respuesta["respuesta"]))
                    #     return
                elif respuesta.status_code==400:
                    consultarGuiaError(Cod_Soc, TipComprobante[tipo_de_comprobante], serie, numero, self)

            except Exception as e:
                mensajeDialogo("error", "generarGuia", e)

def consultarGuia(Cod_Soc, tipo_de_comprobante, serie, numero, self):
    try:
        if Cod_Soc!="1000":
            mensajeDialogo("informacion", "Facturación Electrónica", "Función válida para Multicable")
            return
        data = {}
        data['operacion'] = "consultar_guia"
        data['tipo_de_comprobante'] = dict_tipo_de_comprobante[tipo_de_comprobante]
        data["serie"] = serie
        data["numero"] = numero

        if serie=="T002":
            respuesta=subirNubeFact(data, True, self)
        else:
            respuesta=subirNubeFact(data, False, self)

        if respuesta.status_code==200:
            respuesta=respuesta.json()
            if respuesta["aceptada_por_sunat"]:
                enlace=respuesta["enlace"]
            else:
                enlace="Documento rechazado"

            if self.leURL.text()!=enlace:
                self.leURL.setText(enlace)
            # sql="UPDATE TAB_ALM_001_Mov_Almacenes SET URL='%s' WHERE Cod_Emp='%s'AND Tipo_Comprobante='%s' AND Serie='%s' AND Nro_Facturacion='%s'" % (enlace,Cod_Soc,dict_tipo_de_comprobante[tipo_de_comprobante], serie, numero)
            # respuesta=ejecutarSql(sql)
            #     if respuesta["respuesta"]=="incorrecto":
            #         mensaje("Peligro", "Error", str(respuesta["respuesta"]))
            #         return
            #     mensaje("Informacion", "Actualización", "Enlace actualizado")

    except Exception as e:
        mensajeDialogo("error", "consultarGuia", e)

def consultarGuiaError(Cod_Soc, tipo_de_comprobante, serie, numero, self):
    try:
        if Cod_Soc!="1000":
            mensajeDialogo("informacion", "Facturación Electrónica", "Función válida para Multicable")
            return
        data = {}
        data['operacion'] = "consultar_guia"
        data['tipo_de_comprobante'] = tipo_de_comprobante
        data["serie"] = serie
        data["numero"] = numero

        if serie=="T002":
            respuesta=subirNubeFact(data, True, self)
        else:
            respuesta=subirNubeFact(data, False, self)

        if respuesta.status_code==200:
            respuesta=respuesta.json()
            if respuesta["aceptada_por_sunat"]:
                enlace=respuesta["enlace"]
            else:
                enlace="Documento rechazado"
            if self.leURL.text()!=enlace:
                self.leURL.setText(enlace)
                # sql="UPDATE %s SET rutaNubeFact='%s' WHERE id=%s" % (tipo_de_comprobante.replace("id_", ""), enlace, id_documento)
                # respuesta=ejecutarSql(sql)
                # if respuesta["respuesta"]=="incorrecto":
                #     mensaje("Peligro", "Error", str(respuesta["respuesta"]))
                #     return
                # mensajeDialogo("informacion", "Actualización", "Enlace actualizado")

    except Exception as e:
        mensajeDialogo("error", "consultarGuiaError", e)

#--------------------------------PROGRAMA N° 9 - ERP_ALM_P029----------------------------------

def CargarBindCard(informacion,tbw,self):
    tbw.clearContents()
    rows=tbw.rowCount()
    for r in range(rows):
        tbw.removeRow(1)
    if informacion!=[]:
        flags = (QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
        row=0
        for fila in informacion:
            fila[2]=formatearFecha(fila[2])
            if fila[3]=='0.000':
                fila[3]=fila[10]+ ' (+)'
            else:
                fila[3]=fila[3]+ ' (-)'
            col=0
            for i in fila:
                if i==None or i=='0':
                    i=''
                item=QTableWidgetItem(i)
                item.setFlags(flags)
                item.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter)
                if tbw.rowCount()<=row:
                    tbw.insertRow(tbw.rowCount())
                tbw.setItem(row, col, item)
                tbw.resizeColumnToContents(col)

                col += 1
            row+=1
    else:
        mensajeDialogo("información","Información","No se encontraron registros")

#--------------------------------Programa N° 10 - ERP_FACTURA----------------------------------

def CargarCot(sql,tw,self):
    tw.clearContents()
    informacion=consultarSql(sql)
    if informacion!=[]:
        rows=tw.rowCount()
        for r in range(rows):
            tw.removeRow(0)
        flags = (QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
        row=0
        for fila in informacion:
            PrecioconIGV=float(fila[6])*(1+(porcentaje_de_igv/100))
            DescuentoconIGV=float(fila[7])*(1+(porcentaje_de_igv/100))
            PreciofinalconIGV=PrecioconIGV-(DescuentoconIGV/float(fila[5]))
            Subtotal=float(fila[5])*PreciofinalconIGV
            fila.insert(7,str(PrecioconIGV))
            fila.insert(9,str(DescuentoconIGV))
            fila.insert(10,str(PreciofinalconIGV))
            fila.insert(11,str(Subtotal))
            fila[5]=formatearDecimal(fila[5],'3')
            fila[6]=formatearDecimal(fila[6],'2')
            fila[7]=formatearDecimal(fila[7],'2')
            fila[8]=formatearDecimal(fila[8],'2')
            fila[9]=formatearDecimal(fila[9],'2')
            fila[10]=formatearDecimal(fila[10],'2')
            fila[11]=formatearDecimal(fila[11],'2')
            if fila[12]==None:
                fila[12]='0.000'
            else:
                fila[12]=formatearDecimal(fila[12],'3')
            if fila[13]==None:
                fila[13]='0.000'
            else:
                fila[13]=formatearDecimal(fila[13],'3')
            col=0
            for i in fila:
                item=QTableWidgetItem(i)
                item.setFlags(flags)
                if tw.rowCount()<=row:
                    tw.insertRow(tw.rowCount())
                tw.setItem(row,col, item)
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                tw.resizeColumnToContents(col)
                col += 1
            row+=1
    else:
        mensajeDialogo("informacion", "Información","No se encontraron registros")

def CargarFact(sql,tw,self):
    tw.clearContents()
    informacion=consultarSql(sql)
    if informacion!=[]:
        rows=tw.rowCount()
        for r in range(rows):
            tw.removeRow(1)
        flags = (QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
        row=0
        for fila in informacion:
            PrecioconIGV=float(fila[6])*(1+(porcentaje_de_igv/100))
            DescuentoconIGV=float(fila[7])*(1+(porcentaje_de_igv/100))
            fila.insert(7,str(PrecioconIGV))
            fila.insert(9,str(DescuentoconIGV))
            fila[5]=formatearDecimal(fila[5],'3')
            fila[6]=formatearDecimal(fila[6],'2')
            fila[7]=formatearDecimal(fila[7],'2')
            fila[8]=formatearDecimal(fila[8],'2')
            fila[9]=formatearDecimal(fila[9],'2')
            fila[10]=formatearDecimal(fila[10],'2')
            fila[11]=formatearDecimal(fila[11],'2')
            if fila[12]==None:
                fila[12]='0.000'
            else:
                fila[12]=formatearDecimal(fila[12],'3')
            if fila[13]==None:
                fila[13]='0.000'
            else:
                fila[13]=formatearDecimal(fila[13],'3')
            col=0
            for i in fila:
                item=QTableWidgetItem(i)
                item.setFlags(flags)
                if tw.rowCount()<=row:
                    tw.insertRow(tw.rowCount())
                tw.setItem(row,col, item)
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                tw.resizeColumnToContents(col)
                col += 1
            row+=1
    else:
        mensajeDialogo("informacion", "Información","No se encontraron registros")

def generarDocumento(Cod_Soc,Año,tipo_de_comprobante, serie, numero, sunat_transaction, cliente_tipo_de_documento, ruc, razonSocial, direccion, correo, fechaEmision, fechaVencimiento, moneda, tipoDeCambio, descuento_global, total_descuento, totalGravada, totalIgv, total, condicionesDePago, tbwCuotas, orden_compra_servicio, tipo_de_igv, guia_tipo, nombreArchivo, tbwCotizacion_Cliente, tipo_de_nota_de_credito, tipo_de_nota_de_debito, self):
    if Cod_Soc!="1000":
        mensajeDialogo("informacion", "Facturación Electrónica", "Función válida solo para Multicable")
        return
    if self.leURL.text()=="":
        resultado=mensajeDialogo("pregunta", "Generar Documento", "¿Seguro que desea enviar a Sunat el documento?\n\nUna vez enviado no se podrá modificar el documento")
        if resultado=='Yes':
            try:
                data = {}
                data['operacion'] = "generar_comprobante"
                data['tipo_de_comprobante'] = dict_tipo_de_comprobante[tipo_de_comprobante]
                data["serie"] = serie
                data["numero"] = int(numero)
                data["sunat_transaction"] = sunat_transaction
                data["cliente_tipo_de_documento"] = cliente_tipo_de_documento
                data["cliente_numero_de_documento"] = ruc
                data["cliente_denominacion"] = razonSocial
                data["cliente_direccion"] = direccion
                correos=correo.split(" / ")
                if len(correos)==3:
                    correo1=correos[0]
                    correo2=correos[1]
                    correo3=correos[2]
                elif len(correos)==2:
                    correo1=correos[0]
                    correo2=correos[1]
                    correo3=""
                elif len(correos)==1:
                    correo1=correos[0]
                    correo2=""
                    correo3=""
                else:
                    correo1=""
                    correo2=""
                    correo3=""
                data["cliente_email"] = correo1
                data["cliente_email_1"] = correo2
                data["cliente_email_2"] = correo3
                data["fecha_de_emision"] = fechaEmision
                data["fecha_de_vencimiento"] = fechaVencimiento
                data["moneda"] = moneda
                data["tipo_de_cambio"] = tipoDeCambio
                data["porcentaje_de_igv"] = porcentaje_de_igv
                data["descuento_global"] = descuento_global
                data["total_descuento"] = total_descuento
                data["total_anticipo"] = ""
                data["total_gravada"] = totalGravada
                data["total_inafecta"] = ""
                data["total_exonerada"] = ""
                data["total_igv"] = totalIgv
                data["total_gratuita"] = ""
                data["total_otros_cargos"] = ""
                data["total"] = total
                data["percepcion_tipo"] = ""
                data["percepcion_base_imponible"] = ""
                data["total_percepcion"] = ""
                data["total_incluido_percepcion"] = ""
                data["total_impuestos_bolsas"] = ""
                data["detraccion"] = False
                data["observaciones"] = ""
                if tipo_de_nota_de_credito!=[]:
                    data["documento_que_se_modifica_tipo"] = tipo_de_nota_de_credito[0]
                    data["documento_que_se_modifica_serie"] = tipo_de_nota_de_credito[1]
                    data["documento_que_se_modifica_numero"] = tipo_de_nota_de_credito[2]
                    data["tipo_de_nota_de_credito"] = tipo_de_nota_de_credito[3]
                    data["tipo_de_nota_de_debito"] = ""
                elif tipo_de_nota_de_debito!=[]:
                    data["documento_que_se_modifica_tipo"] = tipo_de_nota_de_debito[0]
                    data["documento_que_se_modifica_serie"] = tipo_de_nota_de_debito[1]
                    data["documento_que_se_modifica_numero"] = tipo_de_nota_de_debito[2]
                    data["tipo_de_nota_de_credito"] = ""
                    data["tipo_de_nota_de_debito"] = tipo_de_nota_de_debito[3]
                else:
                    data["documento_que_se_modifica_tipo"] = ""
                    data["documento_que_se_modifica_serie"] = ""
                    data["documento_que_se_modifica_numero"] = ""
                    data["tipo_de_nota_de_credito"] = ""
                    data["tipo_de_nota_de_debito"] = ""
                data["enviar_automaticamente_a_la_sunat"] = True
                data["enviar_automaticamente_al_cliente"] = False if correo=="" else True
                data["condiciones_de_pago"] = condicionesDePago
                data["medio_de_pago"] = ""
                data["placa_vehiculo"] = ""
                data["orden_compra_servicio"] = orden_compra_servicio
                data["formato_de_pdf"] = ""
                data["generado_por_contingencia"] = serie[0] not in incialesElectronica
                data["bienes_region_selva"] = ""
                data["servicios_region_selva"] = ""
                data["items"]=[]

                for row in range(tbwCotizacion_Cliente.rowCount()):
                #Detalle de Facturación
                    NomMat=tbwCotizacion_Cliente.item(row,2).text()
                    NomMarca=tbwCotizacion_Cliente.item(row,3).text()
                    Unidad=tbwCotizacion_Cliente.item(row,4).text()
                    CodigoMat=tbwCotizacion_Cliente.item(row,14).text()
                    CodigoSunat=tbwCotizacion_Cliente.item(row,15).text()

                    descripcion= NomMat + (" / MARCA: " + NomMarca if NomMarca!="" else "")
                    item={}
                    item['unidad_de_medida'] = Unidad
                    item["codigo"] = CodigoMat
                    item["descripcion"] = descripcion

                    cantidad = float(tbwCotizacion_Cliente.item(row,5).text().replace(",",""))
                    valor_unitario = float(tbwCotizacion_Cliente.item(row,6).text().replace(",",""))
                    precio_unitario = float(tbwCotizacion_Cliente.item(row,7).text().replace(",",""))
                    descuento = float(tbwCotizacion_Cliente.item(row,8).text().replace(",",""))
                    total = float(tbwCotizacion_Cliente.item(row,11).text().replace(",",""))
                    subtotal = round(total/(1+(porcentaje_de_igv/100)),2)
                    igv = round(total-subtotal, 2)

                    if moneda!=2:
                        precio_unitario = round(precio_unitario * tipoDeCambio, 3)
                        valor_unitario = round(precio_unitario / (1+(porcentaje_de_igv/100)), 3)
                        precioFinal=round(total/cantidad, 3)
                        precioFinal_S=round(precioFinal * tipoDeCambio, 3)
                        if str(precioFinal_S)=="-0.0": precioFinal_S=0

                        descuento=redondeoSunat((cantidad*(precio_unitario-precioFinal_S)/(1+(porcentaje_de_igv/100))))
                        descuentoConIGV=round(descuento*(1+(porcentaje_de_igv/100)), 3)
                        total=redondeoSunat(cantidad*precio_unitario-descuentoConIGV)
                        subtotal = redondeoSunat(total/(1+(porcentaje_de_igv/100)))
                        igv = redondeoSunat(total-subtotal)

                    item["cantidad"] = cantidad
                    item["valor_unitario"] = valor_unitario
                    item["precio_unitario"] = precio_unitario
                    item["descuento"] = descuento
                    item["subtotal"] = subtotal
                    item["tipo_de_igv"] = tipo_de_igv
                    item["igv"] = igv
                    item["total"] = total
                    item["codigo_producto_sunat"] = CodigoSunat
                    item["anticipo_regularizacion"] = False
                    item["anticipo_documento_serie"] = ""
                    item["anticipo_documento_numero"] = ""
                    print(item)

                    data['items'].append(item)

                data["guias"]=[]

                data["venta_al_credito"]=[]
                print(tbwCuotas)
                if tbwCuotas!=[]:
                    for fila in tbwCuotas:
                        item={}
                        item["cuota"] = fila[0]
                        item["importe"] = fila[1]
                        item["fecha_de_pago"] = formatearFecha(fila[2])
                        data["venta_al_credito"].append(item)

                crearCarpeta("%s JSON" % tipo_de_comprobante)
                if sys.platform == "win32":
                    ruta=rutaBase[0:-len(rutaBase.split("\\")[-1])]
                    with open('%s JSON\\%s.json' % (ruta + tipo_de_comprobante, nombreArchivo), 'w') as file:
                        json.dump(data, file, indent=4)
                else:
                    ruta=rutaBase[0:-len(rutaBase.split("/")[-1])]
                    with open('%s JSON/%s.json' % (ruta + tipo_de_comprobante, nombreArchivo), 'w') as file:
                        json.dump(data, file, indent=4)

                if serie=="T002":
                    respuesta=subirNubeFact(data, True, self)
                else:
                    respuesta=subirNubeFact(data, False, self)

                if respuesta.status_code==200:
                    respuesta=respuesta.json()
                    if respuesta["aceptada_por_sunat"]:
                        enlace=respuesta["enlace"]
                    else:
                        sunat_description=respuesta["sunat_description"]
                        sunat_note=respuesta["sunat_note"]
                        sunat_responsecode=respuesta["sunat_responsecode"]
                        sunat_soap_error=respuesta["sunat_soap_error"]
                        mensaje_error=[]
                        if sunat_description!=None:
                            mensaje_error.append(sunat_description)
                        if sunat_note!=None:
                            mensaje_error.append(sunat_note)
                        if sunat_responsecode!=None:
                            mensaje_error.append(sunat_responsecode)
                        if sunat_soap_error!='':
                            mensaje_error.append(sunat_soap_error)

                        if mensaje_error==[]:
                            enlace="Pendiente de validación Sunat"
                        else:
                            enlace="Documento rechazado"

                    self.leURL.setText(enlace)
                    self.leURL.setReadOnly(True)
                    sql="UPDATE TAB_VENTA_009_Cabecera_Facturacion SET URL='%s' WHERE Cod_Soc='%s' AND Año='%s' AND Tipo_Comprobante='%s' AND Serie='%s' AND Nro_Facturacion='%s'" % (enlace, Cod_Soc, Año, dict_tipo_de_comprobante[tipo_de_comprobante], serie, numero)
                    respuesta=ejecutarSql(sql)
                    if respuesta["respuesta"]=="incorrecto":
                        mensajeDialogo("advertencia", "Error", str(respuesta["respuesta"]))
                        return
                elif respuesta.status_code==400:
                    consultarDocumentoError(Cod_Soc,Año,dict_tipo_de_comprobante[tipo_de_comprobante],serie,numero,self)

                self.pbEnviar_SUNAT.setEnabled(False)
                self.pbAbrirPDF.setEnabled(True)
                self.pbAnular_Factura.setEnabled(True)

            except Exception as e:
                mensajeDialogo("error", "generarDocumento", e)
                print(e)

def consultarDocumento(Cod_Soc, Año, tipo_de_comprobante, serie, numero, self):
    try:
        if Cod_Soc!="1000":
            mensajeDialogo("informacion", "Facturación Electrónica", "Función válida solo para Multicable")
            return
        data = {}
        data['operacion'] = "consultar_comprobante"
        data['tipo_de_comprobante'] = tipo_de_comprobante
        data["serie"] = serie
        data["numero"] = int(numero)
        # print(data)

        if serie=="T002":
            respuesta=subirNubeFact(data, True, self)
        else:
            respuesta=subirNubeFact(data, False, self)

        if respuesta.status_code==200:
            respuesta=respuesta.json()
            if respuesta["aceptada_por_sunat"]:
                enlace=respuesta["enlace"]
            else:
                sunat_description=respuesta["sunat_description"]
                sunat_note=respuesta["sunat_note"]
                sunat_responsecode=respuesta["sunat_responsecode"]
                sunat_soap_error=respuesta["sunat_soap_error"]
                mensaje_error=[]
                if sunat_description!=None:
                    mensaje_error.append(sunat_description)
                if sunat_note!=None:
                    mensaje_error.append(sunat_note)
                if sunat_responsecode!=None:
                    mensaje_error.append(sunat_responsecode)
                if sunat_soap_error!='':
                    mensaje_error.append(sunat_soap_error)

                if mensaje_error==[]:
                    enlace="Pendiente de validación Sunat"
                else:
                    enlace="Documento rechazado"

            if self.leURL.text()!=enlace:
                self.leURL.setText(enlace)
                sql="UPDATE TAB_VENTA_009_Cabecera_Facturacion SET URL='%s' WHERE Cod_Soc='%s' AND Año='%s' AND Tipo_Comprobante='%s' AND Serie='%s' AND Nro_Facturacion='%s'" % (enlace, Cod_Soc,Año, tipo_de_comprobante, serie, numero)
                respuesta=ejecutarSql(sql)
                if respuesta["respuesta"]=="incorrecto":
                    mensajeDialogo("advertencia", "Error", str(respuesta["respuesta"]))
                    return
                mensajeDialogo("informacion", "Actualización", "Enlace actualizado")

    except Exception as e:
        mensajeDialogo("error", "consultarDocumento", e)

def consultarDocumentoError(Cod_Soc, Año, tipo_de_comprobante, serie, numero, self):
    try:
        if Cod_Soc!="1000":
            mensajeDialogo("informacion", "Facturación Electrónica", "Función válida solo para Multicable")
            return
        data = {}
        data['operacion'] = "consultar_comprobante"
        data['tipo_de_comprobante'] = tipo_de_comprobante
        data["serie"] = serie
        data["numero"] = int(numero)

        if serie=="T002":
            respuesta=subirNubeFact(data, True, self)
        else:
            respuesta=subirNubeFact(data, False, self)

        if respuesta.status_code==200:
            respuesta=respuesta.json()
            if respuesta["aceptada_por_sunat"]:
                enlace=respuesta["enlace"]
            else:
                sunat_description=respuesta["sunat_description"]
                sunat_note=respuesta["sunat_note"]
                sunat_responsecode=respuesta["sunat_responsecode"]
                sunat_soap_error=respuesta["sunat_soap_error"]
                mensaje_error=[]
                if sunat_description!=None:
                    mensaje_error.append(sunat_description)
                if sunat_note!=None:
                    mensaje_error.append(sunat_note)
                if sunat_responsecode!=None:
                    mensaje_error.append(sunat_responsecode)
                if sunat_soap_error!='':
                    mensaje_error.append(sunat_soap_error)

                if mensaje_error==[]:
                    enlace="Pendiente de validación Sunat"
                else:
                    enlace="Documento rechazado"
            if self.leURL.text()!=enlace:
                self.leURL.setText(enlace)
                self.leURL.setReadOnly(True)
                sql="UPDATE TAB_VENTA_009_Cabecera_Facturacion SET URL='%s' WHERE Cod_Soc='%s' AND Año='%s' AND Tipo_Comprobante='%s' AND Serie='%s' AND Nro_Facturacion='%s';" % (enlace, Cod_Soc, Año, tipo_de_comprobante, serie, numero)
                respuesta=ejecutarSql(sql)
                if respuesta["respuesta"]=="incorrecto":
                    mensajeDialogo("advertencia", "Error", str(respuesta["respuesta"]))
                    return
                mensajeDialogo("informacion", "Actualización", "Enlace actualizado")

    except Exception as e:
        mensajeDialogo("error", "consultarDocumentoError", e)

def anularDocumento(Cod_Soc, Año, tipo_de_comprobante, serie, numero, motivo, self):
    try:
        if Cod_Soc!="1000":
            mensajeDialogo("informacion", "Facturación Electrónica", "Función válida solo para Multicable")
            return
        if self.leURL.text()=="Documento rechazado":
            ruta="Anulación por rechazo de Sunat"
            motivo="Anulación por rechazo de Sunat"
        else:
            data = {}
            data['operacion'] = "generar_anulacion"
            data['tipo_de_comprobante'] = dict_tipo_de_comprobante[tipo_de_comprobante]
            data["serie"] = serie
            data["numero"] = int(numero)
            data["motivo"] = motivo
            data["codigo_unico"] = ""

            if tipo_de_comprobante in FacturacionElectronica and serie[0] in incialesElectronica:
                if probandoFactElect or int(self.leFecha_Emision.text()[6:10])>=2021:
                    if serie=="T002":
                        respuesta=subirNubeFact(data, True, self)
                    else:
                        respuesta=subirNubeFact(data, False, self)

                    if respuesta.status_code==200:
                        ruta=respuesta.json()["enlace"]
                    elif respuesta.status_code==400:
                        validarAnulacionError(Cod_Soc, Año, tipo_de_comprobante, serie, numero, motivo, self)
                        return
                    else:
                        ruta="Error"
                else:
                    ruta="No electrónica"
            else:
                ruta=""

        self.leURL.setText(ruta)
        tabla=tipo_de_comprobante.lower()

        anularDocumentoSQL(Cod_Soc, Año, tabla, dict_tipo_de_comprobante[tipo_de_comprobante], ruta, serie, numero, self)

    except Exception as e:
        mensajeDialogo("error", "anularDocumento", e)

def validarAnulacion(Cod_Soc, Año, tipo_de_comprobante, serie, numero, self):
    try:
        if Cod_Soc!="1000":
            mensajeDialogo("informacion", "Facturación Electrónica", "Función válida solo para Multicable")
            return
        data = {}
        data['operacion'] = "consultar_anulacion"
        data['tipo_de_comprobante'] = tipo_de_comprobante
        data["serie"] = serie
        data["numero"] = int(numero)

        if serie=="T002":
            respuesta=subirNubeFact(data, True, self)
        else:
            respuesta=subirNubeFact(data, False, self)

        if respuesta.status_code==200:
            respuesta=respuesta.json()
            if respuesta["aceptada_por_sunat"]:
                enlace=respuesta["enlace"]
            else:
                enlace="Documento rechazado"
            if self.leURL.text()!=enlace:
                self.leURL.setText(enlace)
                sql="UPDATE TAB_VENTA_009_Cabecera_Facturacion SET URL='%s' WHERE Cod_Soc='%s' AND Año='%s' AND Tipo_Comprobante='%s' AND Serie='%s' AND Nro_Facturacion='%s'" % (enlace, Cod_Soc, Año, tipo_de_comprobante, serie, numero)
                respuesta=ejecutarSql(sql)
                if respuesta["respuesta"]=="incorrecto":
                    mensajeDialogo("advertencia", "Error", str(respuesta["respuesta"]))
                    return
                mensajeDialogo("informacion", "Actualización", "Enlace actualizado")

    except Exception as e:
        mensajeDialogo("error", "validarAnulacion", e)

def validarAnulacionError(Cod_Soc, Año, tipo_de_comprobante, serie, numero, motivo, self):
    try:
        if Cod_Soc!="1000":
            mensajeDialogo("informacion", "Facturación Electrónica", "Función válida solo para Multicable")
            return
        data = {}
        data['operacion'] = "consultar_anulacion"
        data['tipo_de_comprobante'] = dict_tipo_de_comprobante[tipo_de_comprobante]
        data["serie"] = serie
        data["numero"] = int(numero)

        if tipo_de_comprobante in FacturacionElectronica and serie[0] in incialesElectronica:
            if serie=="T002":
                respuesta=subirNubeFact(data, True, self)
            else:
                respuesta=subirNubeFact(data, False, self)
            if respuesta.status_code==200:
                ruta=respuesta.json()["enlace"]
            else:
                ruta="Error"
        else:
            ruta=""

        self.leURL.setText(ruta)
        tabla=tipo_de_comprobante.lower()

        anularDocumentoSQL(Cod_Soc, Año, tabla, dict_tipo_de_comprobante[tipo_de_comprobante], ruta, serie, numero, self)

    except Exception as e:
        mensajeDialogo("error", "validarAnulacionError", e)

def anularDocumentoSQL(Cod_Soc, Año, tabla, tipo_de_comprobante, ruta, serie, numero, self):
    try:
        sql="UPDATE TAB_VENTA_009_Cabecera_Facturacion SET URL='%s', Estado_Factura='1' WHERE Cod_Soc='%s' AND Año='%s' AND Tipo_Comprobante='%s' AND Serie='%s' AND Nro_Facturacion='%s'" % (ruta, Cod_Soc, Año, tipo_de_comprobante, serie, numero)
        respuesta=ejecutarSql(sql)
        if respuesta["respuesta"]=="incorrecto":
            mensajeDialogo("advertencia", "Error", str(respuesta["respuesta"]))
            return

        mensajeDialogo("informacion", "Anular", tabla.capitalize() + " anulada")
        self.leEstado_Documento.setText("ANULADO")
        self.leEstado_Documento.setStyleSheet("color: rgb(255,130,130);\n""background-color: rgb(255,255,255);")

    except Exception as e:
        mensajeDialogo("error", "anularDocumentoSQL", e)

#--------------------------------Programa N° 11 - ERP_NOTAS----------------------------------

def CargarFactNota(sql,tw,DescuentoGlobal,self):
    tw.clearContents()
    informacion=consultarSql(sql)
    if informacion!=[]:
        rows=tw.rowCount()
        for r in range(rows):
            tw.removeRow(0)
        flags = (QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
        row=0
        for fila in informacion:
            fila[7]='0.00'
            fila[6]=fila[8]
            PreciosinIGV=float(fila[6])*(1/(1+(porcentaje_de_igv/100)))
            DescuentosinIGV=float(fila[7])*(1/(1+(porcentaje_de_igv/100)))
            fila.insert(6,str(PreciosinIGV))
            fila.insert(8,str(DescuentosinIGV))
            fila[5]=formatearDecimal(fila[5],'3')
            fila[6]=formatearDecimal(fila[6],'4')
            fila[7]=formatearDecimal(fila[7],'2')
            fila[8]=formatearDecimal(fila[8],'2')
            fila[9]=formatearDecimal(fila[9],'2')
            fila[10]=formatearDecimal(fila[10],'2')
            fila[11]=formatearDecimal(fila[11],'2')
            if fila[12]==None:
                fila[12]='0.000'
            else:
                fila[12]=formatearDecimal(fila[12],'3')
            if fila[13]==None:
                fila[13]='0.000'
            else:
                fila[13]=formatearDecimal(fila[13],'3')
            col=0
            for i in fila:
                item=QTableWidgetItem(i)
                item.setFlags(flags)
                if tw.rowCount()<=row:
                    tw.insertRow(tw.rowCount())
                tw.setItem(row,col, item)
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                tw.resizeColumnToContents(col)
                col += 1
            row+=1
    else:
        mensajeDialogo("informacion", "Información","No se encontraron registros")
