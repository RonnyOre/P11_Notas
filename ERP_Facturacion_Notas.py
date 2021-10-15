import sys
from datetime import datetime
from Funciones04 import *
from diccionario_sunat import *
from PyQt5 import uic, QtCore, QtGui
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
import urllib.request
import pandas as pd
from pylab import *
from fpdf import FPDF
import tkinter as tk
from tkinter import filedialog

TipNota={'NOTA DE CRÉDITO':'3','NOTA DE DÉBITO':'4'}
TipComprobante={'FACTURA':'1','BOLETA':'2'}
# TipSerie={'FACTURA':['F001','0001'],'BOLETA':['B001','0001']}
TipSerie=['F001','B001','0001']
EstadoFactura={'1':'ANULADO','2':'CANCELADO','3':'PENDIENTE','4':'VENCIDO'}
FormaPago={'1':'CONTADO','2':'CRÉDITO'}
sqlCliente="SELECT Razon_social,Cod_cliente FROM `TAB_COM_001_Maestro Clientes`"
sqlMoneda="SELECT Descrip_moneda,Cod_moneda FROM TAB_SOC_008_Monedas"

sqlMarca="SELECT Descrip_Marca,Cod_Marca  FROM TAB_MAT_010_Marca_de_Producto"
sqlMatSunat="SELECT Descrip_SUNAT,Cod_Sunat FROM TAB_SOC_026_Tabla_productos_SUNAT"

class IngresarMotivo(QDialog):
    def __init__(self, titulo, label, limite):
        QDialog.__init__(self)
        uic.loadUi("Ingresar_Motivo.ui",self)
        self.texto=""
        self.setWindowTitle(titulo)
        self.lblLabel.setText(label)
        self.leTexto.setMaxLength(limite)

        self.buttonBox.accepted.connect(self.aceptar)
        self.buttonBox.rejected.connect(self.rechazar)
        cargarIcono(self, 'erp')

    def aceptar(self):
        self.texto=self.leTexto.text()
        self.close()

    def rechazar(self):
        self.texto=""
        self.close()

class SeleccionarNota(QDialog):
    def __init__(self,Tiponota,Serienota):
        QDialog.__init__(self)
        uic.loadUi("ERP_Consulta_Nota.ui",self)

        self.twNotas.itemDoubleClicked.connect(self.Facturacion)
        self.lePalabra.textChanged.connect(self.buscar)

        cargarIcono(self, 'erp')

        sqlNota='''SELECT a.Serie,a.Nro_Facturacion,a.Fecha_Emision, b.Razon_social, b.RUC, SUM(c.Sub_Total), a.Estado_Factura
        FROM TAB_VENTA_009_Cabecera_Facturacion a
        LEFT JOIN `TAB_COM_001_Maestro Clientes`b ON a.Cod_Cliente=b.Cod_cliente
        LEFT JOIN TAB_VENTA_010_Detalle_Facturacion c ON a.Cod_Soc=c.Cod_Soc AND a.Año=c.Año AND a.Tipo_Comprobante=c.Tipo_Comprobante AND a.Serie=c.Serie AND a.Nro_Facturacion=c.Nro_Facturacion
        WHERE a.Cod_Soc='%s' AND a.Año='%s' AND a.Tipo_Comprobante='%s' AND a.Serie='%s'
        GROUP BY c.Nro_Facturacion'''%(Cod_Soc,Año,Tiponota,Serienota)
        Nota=consultarSql(sqlNota)

        self.twNotas.clear()
        for fila in Nota:
            fila[6]=EstadoFactura[fila[6]]
            item=QTreeWidgetItem(self.twNotas,fila)
            item.setFlags(item.flags() | QtCore.Qt.ItemIsEditable)
            for i in range(len(fila)):
                if i!=3 or i!=5:
                    item.setTextAlignment(i,QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                if i==5:
                    item.setTextAlignment(i,QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
                self.twNotas.resizeColumnToContents(i)
            self.twNotas.addTopLevelItem(item)

    def buscar(self):
        buscarTabla(self.twNotas, self.lePalabra.text(), [1,2,3,4])

    def Facturacion(self,item):
        global SerieNota,NroNota
        SerieNota=item.text(0)
        NroNota=item.text(1)
        self.close()

class SeleccionarFacturacion(QDialog):
    # def __init__(self,SerieComp):
    def __init__(self,TipoComprobante,SerieComp):
        QDialog.__init__(self)
        uic.loadUi("ERP_Consulta_Documento.ui",self)

        self.twFacturacion.itemDoubleClicked.connect(self.Facturacion)
        self.lePalabra.textChanged.connect(self.buscar)

        cargarIcono(self, 'erp')

        sqlFac='''SELECT a.Serie,a.Nro_Facturacion,a.Fecha_Emision, b.Razon_social, b.RUC, SUM(c.Sub_Total), a.Estado_Factura
        FROM TAB_VENTA_009_Cabecera_Facturacion a
        LEFT JOIN `TAB_COM_001_Maestro Clientes`b ON a.Cod_Cliente=b.Cod_cliente
        LEFT JOIN TAB_VENTA_010_Detalle_Facturacion c ON a.Cod_Soc=c.Cod_Soc AND a.Año=c.Año AND a.Tipo_Comprobante=c.Tipo_Comprobante AND a.Serie=c.Serie AND a.Nro_Facturacion=c.Nro_Facturacion
        WHERE a.Cod_Soc='%s' AND a.Año='%s' AND a.Tipo_Comprobante='%s'
        GROUP BY c.Serie,c.Tipo_Comprobante,c.Nro_Facturacion
        ORDER BY a.Serie DESC,a.Nro_Facturacion DESC, a.Fecha_Emision DESC;'''%(Cod_Soc,Año,TipoComprobante)
        Fac=consultarSql(sqlFac)

        self.twFacturacion.clear()
        for fila in Fac:
            fila[6]=EstadoFactura[fila[6]]
            item=QTreeWidgetItem(self.twFacturacion,fila)
            item.setFlags(item.flags() | QtCore.Qt.ItemIsEditable)
            for i in range(len(fila)):
                if i!=3 or i!=5:
                    item.setTextAlignment(i,QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                if i==5:
                    item.setTextAlignment(i,QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
                self.twFacturacion.resizeColumnToContents(i)
            self.twFacturacion.addTopLevelItem(item)

    def buscar(self):
        buscarTabla(self.twFacturacion, self.lePalabra.text(), [1,2,3,4])

    def Facturacion(self,item):
        global Serie,Nro_Facturacion
        Serie=item.text(0)
        Nro_Facturacion=item.text(1)
        self.close()

class VerCuotas(QDialog):
    def __init__(self,Cod_Cliente,Nro_Cotización):
        QDialog.__init__(self)
        uic.loadUi("ERP_Consulta_Cuotas.ui",self)

        self.pbSalir.clicked.connect(self.Salir)

        cargarIcono(self, 'erp')
        cargarIcono(self.pbGrabar, 'grabar')
        cargarIcono(self.pbModificar, 'modificar')
        cargarIcono(self.pbSalir, 'salir')

        sqlCuotas='''SELECT Nro_correlativo,Monto_Deposito,Fecha_Limite_Deposito
        FROM TAB_VENTA_005_Cabecera_Operaciones_Bancarias_Por_Cotizacion
        WHERE Cod_Soc='%s' AND Año_Cot_Client='%s' AND Cod_Cliente='%s' AND  Nro_Cot_Client='%s'
        ORDER BY Nro_correlativo ASC''' %(Cod_Soc,Año,Cod_Cliente,Nro_Cotización)
        informacion=consultarSql(sqlCuotas)

        self.tbwCuotas.clearContents()
        rows=self.tbwCuotas.rowCount()
        for r in range(rows):
            self.tbwCuotas.removeRow(0)
        if informacion!=[]:
            flags = (QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
            row=0
            for fila in informacion:
                fila[1]=formatearDecimal(fila[1],'2')
                fila[2]=formatearFecha(fila[2])
                col=0
                for i in fila:
                    item=QTableWidgetItem(i)
                    item.setFlags(flags)
                    item.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter)
                    if self.tbwCuotas.rowCount()<=row:
                        self.tbwCuotas.insertRow(self.tbwCuotas.rowCount())
                    self.tbwCuotas.setItem(row, col, item)
                    col += 1
                row+=1

    def Salir(self):
        self.close()

class ERP_Facturacion_Notas(QMainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        uic.loadUi("ERP_Facturacion_Notas.ui",self)

        global Cod_Soc,Nom_Soc,Cod_Usuario
        global Año,dicCliente,dicFormaPago,dicMoneda,dicMarca,dicMat,dicMatSUNAT

        Cod_Soc='1000'
        Nom_Soc='MULTICABLE PERU S.A.C.'
        Cod_Usuario='2021100004'

        self.cbTipo_Nota.activated.connect(self.cargarSerie)
        self.cbSerie.activated.connect(self.activarBusqueda)

        self.pbSelec_Nota.clicked.connect(self.SeleccionarNota)
        self.pbSelec_Factura.clicked.connect(self.SeleccionarFactura)
        self.pbSelec_Boleta.clicked.connect(self.SeleccionarBoleta)

        self.pbLimpiar.clicked.connect(self.Limpiar)

        self.leDescuento_Global.textChanged.connect(self.numeroDescuentoGlobal)
        self.leDescuento_Global.editingFinished.connect(self.descuentoGlobal)

        # self.pbVerCuotas.clicked.connect(self.VerCuotas)

        self.pbGrabar.clicked.connect(self.Grabar)
        self.pbEnviar_SUNAT.clicked.connect(self.botonEnviarSunat)
        self.pbAbrirPDF.clicked.connect(self.botonAbrirPdf)
        self.pbAnular_Factura.clicked.connect(self.botonAnular)
        self.pbConsulta_Anulacion.clicked.connect(self.botonConsultaAnulacion)
        self.pbImprimir.clicked.connect(self.nota_credito_o_debito)
        self.pbSalir.clicked.connect(self.Salir)

    # def datosGenerales(self, codSoc, empresa, usuario):
    #     global Cod_Soc, Nom_Soc, Cod_Usuario
    #     global Año,dicCliente,dicMoneda,dicMarca,dicMat,dicMatSUNAT
    #     Cod_Soc = codSoc
    #     Nom_Soc = empresa
    #     Cod_Usuario = usuario

        self.cargarCombos()
        self.tipoCambio()
        self.botones()
        self.cargarDiccionarios()
        self.bloquearDatos()

        cargarLogo(self.lbLogo_Mp,'multiplay')
        cargarLogo(self.lbLogo_Soc, Cod_Soc)
        cargarIcono(self, 'erp')
        cargarIcono(self.pbVerCuotas, 'consultar')
        cargarIcono(self.pbGrabar, 'grabar')
        cargarIcono(self.pbAbrirPDF, 'pdf')
        cargarIcono(self.pbSelec_Nota, 'buscar')
        cargarIcono(self.pbSelec_Factura,'buscar')
        cargarIcono(self.pbSelec_Boleta,'buscar')
        cargarIcono(self.pbLimpiar,'nuevo')
        cargarIcono(self.pbEnviar_SUNAT, 'enviar')
        cargarIcono(self.pbAnular_Factura,'darbaja')
        cargarIcono(self.pbImprimir, 'imprimir')
        cargarIcono(self.pbConsulta_Anulacion,'buscar')
        cargarIcono(self.pbSalir, 'salir')

        now = datetime.datetime.now()
        Año=str(now.year)

        Cliente=consultarSql(sqlCliente)
        dicCliente={}
        for dato in Cliente:
            dicCliente[dato[0]]=dato[1]

        Moneda=consultarSql(sqlMoneda)
        dicMoneda={}
        for dato in Moneda:
            dicMoneda[dato[0]]=dato[1]

        Marca=consultarSql(sqlMarca)
        dicMarca={}
        for dato in Marca:
            dicMarca[dato[0]]=dato[1]

        sqlMat="SELECT Descrip_Mat,Cod_Mat FROM TAB_MAT_001_Catalogo_Materiales WHERE Cod_Soc='%s';"%(Cod_Soc)
        Material=consultarSql(sqlMat)
        dicMat={}
        for dato in Material:
            dicMat[dato[0]]=dato[1]

        MaterialSUNAT=consultarSql(sqlMatSunat)
        dicMatSUNAT={}
        for dato in MaterialSUNAT:
            dicMatSUNAT[dato[0]]=dato[1]

    def botones(self):
        self.pbSelec_Factura.setEnabled(False)
        self.pbSelec_Boleta.setEnabled(False)
        self.pbGrabar.setEnabled(False)
        self.pbEnviar_SUNAT.setEnabled(False)
        self.pbAbrirPDF.setEnabled(False)
        self.pbAnular_Factura.setEnabled(False)
        self.pbConsulta_Anulacion.setEnabled(False)
        self.pbImprimir.setEnabled(False)

    def cargarDiccionarios(self):
        global dict_descripBreve
        sqlDescripBreve = '''SELECT Cod_Mat, Texto_Breve FROM `TAB_MAT_001_Catalogo_Materiales` WHERE Cod_Soc='%s';''' % (Cod_Soc) #Código de Sociedad
        dataDescripBreve = consultarSql(sqlDescripBreve)
        dict_descripBreve = {}
        for mat in dataDescripBreve:
            dict_descripBreve[mat[0]] = mat[1]
        # print(dict_descripBreve)

    def cargarCombos(self):
        for i in TipNota.keys():
            self.cbTipo_Nota.addItem(i)
            self.cbTipo_Nota.setCurrentIndex(-1)

        for i in dict_sunat_transaction.keys():
            self.cbTipo_Operacion.addItem(i)
            self.cbTipo_Operacion.setCurrentIndex(-1)

    def tipoCambio(self):
        Fecha=datetime.datetime.now().strftime("%Y-%m-%d")
        sqlTipoCambio="SELECT Tipo_Cambio FROM `TAB_SOC_007_Monedas Tipo de cambio` WHERE Cod_moneda = '%s' AND Fecha_Reg = '%s' AND Mon_Cambio = '%s' ORDER BY Fecha_Reg DESC, Hora_Reg DESC"%(1,Fecha,2)
        lista = convlist(sqlTipoCambio)
        if lista!=[]:
            self.leTipo_Cambio.setText(lista[0])
            self.leTipo_Cambio.setReadOnly(True)
        else:
            # reply=mensajeDialogo('pregunta','Información',"No hay un tipo de cambio registrado para la fecha actual.\n¿Desea continuar con la cotización?")
            # if reply=='Yes':
                #Retirar, estas lineas de abajo son provicionales!
            sqlProvicional="SELECT Tipo_Cambio FROM `TAB_SOC_007_Monedas Tipo de cambio` WHERE Cod_moneda = '%s' AND Fecha_Reg<'%s' AND Mon_Cambio = '%s' ORDER BY Fecha_Reg DESC, Hora_Reg DESC"%(1,Fecha,2)
            TipCamPro = convlist(sqlProvicional)
            self.leTipo_Cambio.setText(TipCamPro[0])
            self.leTipo_Cambio.setReadOnly(True)

    def cargarSerie(self):
        self.cbSerie.clear()
        self.cbMotivo.clear()
        tipo_nota=self.cbTipo_Nota.currentText()
        # listSeries=[]
        # for datos in TipSerie.values():
        #     for i in datos:
        #         if i not in listSeries:
        #             listSeries.append(i)
        # listSeries.sort(reverse=True)
        # for i in listSeries:
        for i in TipSerie:
            self.cbSerie.addItem(i)
            self.cbSerie.setCurrentIndex(-1)

        if tipo_nota=="NOTA DE CRÉDITO":
            for k,v in dict_tipo_de_nota_de_credito.items():
                self.cbMotivo.addItem(k)
                self.cbMotivo.setCurrentIndex(-1)

        elif tipo_nota=="NOTA DE DÉBITO":
            for k,v in dict_tipo_de_nota_de_debito.items():
                self.cbMotivo.addItem(k)
                self.cbMotivo.setCurrentIndex(-1)

        # if len(NroDocumento)!=0:
        #     if len(NroDocumento)==11:
        #         if self.cbSerie.currentText()=='B001':
        #             self.leRUC.setText(NroDocumento[2:10])

    def activarBusqueda(self):
        Serie=self.cbSerie.currentText()
        if Serie=='F001':
            self.pbSelec_Factura.setEnabled(True)
            self.pbSelec_Boleta.setEnabled(False)
        elif Serie=='B001':
            self.pbSelec_Factura.setEnabled(False)
            self.pbSelec_Boleta.setEnabled(True)
        if Serie=='0001':
            self.pbSelec_Factura.setEnabled(True)
            self.pbSelec_Boleta.setEnabled(True)

    def numeroDescuentoGlobal(self):
        validarNumero(self.leDescuento_Global)

    def descuentoGlobal(self):
        self.leDescuento_Global.setText(formatearDecimal(self.leDescuento_Global.text(),'2'))
        self.cargarMontos()

    def cargarMontos(self):
        try:
            list_cantidad = []
            list_precio_sinIGV  = []
            list_descuento_sinIGV= []

            for i in range(self.tbwCotizacion_Cliente.rowCount()):
                try:
                    list_cantidad.append(self.tbwCotizacion_Cliente.item(i,5).text())
                except:
                    list_cantidad.append("")
                try:
                    list_precio_sinIGV.append(self.tbwCotizacion_Cliente.item(i,6).text())
                except:
                    list_precio_sinIGV.append("")
                try:
                    desc=self.tbwCotizacion_Cliente.item(i,8).text()
                    d = desc.replace(",","")
                    list_descuento_sinIGV.append(float(d))
                except:
                    list_descuento_sinIGV.append("")

            list_cantxprecio = []
            for i in range(len(list_cantidad)):
                c = list_cantidad[i].replace(",","")
                p = list_precio_sinIGV[i].replace(",","")
                cant_x_precio= float(c) * float(p)
                list_cantxprecio.append(cant_x_precio)
                i += 1

            TipoCambio=self.leTipo_Cambio.text()
            Subtotal=sum(list_cantxprecio)
            DescuentoItem=sum(list_descuento_sinIGV)
            DescuentoGlobal=self.leDescuento_Global.text()
            if len(DescuentoGlobal)==0:
                DescuentoGlobal="0.00"
                self.leDescuento_Global_Soles.clear()
            else:
                DescuentoGlobal=DescuentoGlobal.replace(",","")
                self.leDescuento_Global_Soles.setText(formatearDecimal(str(float(DescuentoGlobal)*float(TipoCambio)),'2'))
            DescuentoTotal=DescuentoItem+float(DescuentoGlobal)
            TotalsinIGV=Subtotal-DescuentoTotal
            IGV=0.18*TotalsinIGV
            Total=TotalsinIGV+IGV

            self.leSub_Total.setText(formatearDecimal(str(Subtotal),'2'))
            self.leDescuento_Item.setText(formatearDecimal(str(DescuentoItem),'2'))
            self.leTotal_SinIGV.setText(formatearDecimal(str(TotalsinIGV),'2'))
            self.leIGV.setText(formatearDecimal(str(IGV),'2'))
            self.leTotal.setText(formatearDecimal(str(Total),'2'))

            self.leSub_Total_Soles.setText(formatearDecimal(str(Subtotal*float(TipoCambio)),'2'))
            self.leDescuento_Item_Soles.setText(formatearDecimal(str(DescuentoItem*float(TipoCambio)),'2'))
            self.leTotal_SinIGV_Soles.setText(formatearDecimal(str(TotalsinIGV*float(TipoCambio)),'2'))
            self.leIGV_Soles.setText(formatearDecimal(str(IGV*float(TipoCambio)),'2'))
            self.leTotal_Soles.setText(formatearDecimal(str(Total*float(TipoCambio)),'2'))

        except Exception as e:
            print(e)

    def bloquearDatos(self):

        self.leNumero.setReadOnly(True)
        self.leFecha_Emision.setReadOnly(True)
        self.leEstado_Documento.setReadOnly(True)
        self.leURL.setReadOnly(True)

        self.leSerie_Numero.setReadOnly(True)
        self.leRazon_Social.setReadOnly(True)
        self.leDireccion.setReadOnly(True)
        self.leRUC.setReadOnly(True)
        self.leInterlocutor.setReadOnly(True)
        self.leDNI.setReadOnly(True)
        self.leCorreo.setReadOnly(False)

        bloquearCb(self.cbTipo_Operacion)
        self.leForma_Pago.setReadOnly(True)
        self.leMoneda.setReadOnly(True)
        self.leTipo_Cambio.setReadOnly(True)

        self.leSub_Total.setReadOnly(True)
        self.leDescuento_Item.setReadOnly(True)
        self.leDescuento_Global.setEnabled(False)
        self.leTotal_SinIGV.setReadOnly(True)
        self.leIGV.setReadOnly(True)
        self.leTotal.setReadOnly(True)

        self.leSub_Total_Soles.setReadOnly(True)
        self.leDescuento_Item_Soles.setReadOnly(True)
        self.leDescuento_Global_Soles.setReadOnly(True)
        self.leTotal_SinIGV_Soles.setReadOnly(True)
        self.leIGV_Soles.setReadOnly(True)
        self.leTotal_Soles.setReadOnly(True)
        self.pbVerCuotas.setEnabled(False)

    def Limpiar(self):

        self.botones()
        #Limpieza de datos.....
        self.cbTipo_Nota.setCurrentIndex(-1)
        self.cbTipo_Nota.setEnabled(True)
        self.cbSerie.setCurrentIndex(-1)
        self.cbSerie.setEnabled(True)
        self.leNumero.clear()
        self.leFecha_Emision.clear()
        self.leEstado_Documento.clear()
        self.leURL.clear()
        self.cbMotivo.setCurrentIndex(-1)
        self.cbMotivo.setEnabled(True)

        self.leSerie_Numero.clear()
        self.leRazon_Social.clear()
        self.leDireccion.clear()
        self.leRUC.clear()
        self.leInterlocutor.clear()
        self.leDNI.clear()
        self.leCorreo.clear()

        self.cbTipo_Operacion.setCurrentIndex(-1)
        self.leForma_Pago.clear()
        self.leMoneda.clear()

        self.leSub_Total.clear()
        self.leDescuento_Item.clear()
        self.leDescuento_Global.clear()
        self.leTotal_SinIGV.clear()
        self.leIGV.clear()
        self.leTotal.clear()

        self.leSub_Total_Soles.clear()
        self.leDescuento_Item_Soles.clear()
        self.leDescuento_Global_Soles.clear()
        self.leTotal_SinIGV_Soles.clear()
        self.leIGV_Soles.clear()
        self.leTotal_Soles.clear()

        self.tbwCotizacion_Cliente.clearContents()
        rows=self.tbwCotizacion_Cliente.rowCount()
        for r in range(rows):
            self.tbwCotizacion_Cliente.removeRow(0)

    # def VerCuotas(self):
    #
    #     Tip_Comprobante=self.cbSerie.currentText()
    #     Nro_Facturacion=self.leSerie_Numero.text()
    #     Cod_Cliente=dicCliente[self.leRazon_Social.text()]
    #
    #     VerCuotas(Cod_Cliente,Nro_Cotización).exec_()

    def SeleccionarFactura(self):
        global Serie,Nro_Facturacion
        Serie=None
        Nro_Facturacion=None

        TipNota=self.cbTipo_Nota.currentText()
        Serienota=self.cbSerie.currentText()
        TipoComprobante='1'
        if len(TipNota)!=0 and len(Serienota)!=0:
            # for k,v in TipSerie.items():
            #     if Serienota in v:
            #         TipoComp=k
            # for k,v in TipComprobante.items():
            #     if TipoComp==k:
            #         TipoComprobante=v
            # print(TipoComprobante)
            # SeleccionarFacturacion(Serienota).exec_()
            SeleccionarFacturacion(TipoComprobante,Serienota).exec_()
            self.CargarFacturacion()
        elif len(TipNota)!=0 and len(Serienota)==0:
            mensajeDialogo('informacion','Información','Seleccione Serie')
        else:
            mensajeDialogo('informacion','Información','Seleccione Tipo de Comprobante y la Serie')

    def SeleccionarBoleta(self):
        global Serie,Nro_Facturacion
        Serie=None
        Nro_Facturacion=None

        TipNota=self.cbTipo_Nota.currentText()
        Serienota=self.cbSerie.currentText()
        TipoComprobante='2'
        if len(TipNota)!=0 and len(Serienota)!=0:
            # for k,v in TipSerie.items():
            #     if Serienota in v:
            #         TipoComp=k
            # for k,v in TipComprobante.items():
            #     if TipoComp==k:
            #         TipoComprobante=v
            # print(TipoComprobante)
            # SeleccionarFacturacion(Serienota).exec_()
            SeleccionarFacturacion(TipoComprobante,Serienota).exec_()
            self.CargarFacturacion()
        elif len(TipNota)!=0 and len(Serienota)==0:
            mensajeDialogo('informacion','Información','Seleccione Serie')
        else:
            mensajeDialogo('informacion','Información','Seleccione Tipo de Comprobante y la Serie')

    def CargarFacturacion(self):
        try:
            sqlCabFact='''SELECT b.Razon_social,b.Direcc_cliente,b.RUC,b.Representante_Cliente,b.DNI,b.Correo_Representante,a.Tipo_Operación,a.Forma_Pago,c.Descrip_moneda,a.Descuento_Global
            FROM TAB_VENTA_009_Cabecera_Facturacion a
            LEFT JOIN `TAB_COM_001_Maestro Clientes`b ON a.Cod_Cliente=b.Cod_cliente
            LEFT JOIN TAB_SOC_008_Monedas c ON a.Moneda=c.Cod_moneda
            WHERE a.Cod_Soc='%s' AND a.Año='%s' AND a.Tipo_Comprobante='%s' AND a.Serie='%s' AND a.Nro_Facturacion='%s';'''%(Cod_Soc,Año,Tipo_Comprobante,Serie,Nro_Facturacion)
            lista=convlist(sqlCabFact)

            sqlDetFact='''SELECT a.Item,c.Descrip_SUNAT, b.Descrip_Mat, d.Descrip_Marca, c.Unidad_SUNAT, a.Cantidad, a.Precio_sin_IGV, a.Descuento_con_IGV, SUM(e.Stock_disponible), SUM(e.Stock_Bloq_con_QA),a.Cod_Material,b.Cod_Prod_SUNAT
            FROM TAB_VENTA_010_Detalle_Facturacion a
            LEFT JOIN TAB_MAT_001_Catalogo_Materiales b ON b.Cod_Soc=a.Cod_Soc AND b.Cod_Mat=a.Cod_Material
            LEFT JOIN TAB_SOC_026_Tabla_productos_SUNAT c ON b.Cod_Prod_SUNAT=c.Cod_Sunat
            LEFT JOIN TAB_MAT_010_Marca_de_Producto d ON a.Marca=d.Cod_Marca
            LEFT JOIN TAB_MAT_002_Stock_Almacen e ON a.Cod_Soc=e.Cod_Soc AND a.Cod_Material=e.Cod_Mat
            WHERE a.Cod_Soc='%s' AND a.Año='%s' AND a.Tipo_Comprobante='%s' AND a.Serie='%s' AND a.Nro_Facturacion='%s'
            GROUP BY e.Cod_Mat ORDER BY a.Item ASC;'''%(Cod_Soc,Año,Tipo_Comprobante,Serie,Nro_Facturacion)

            for i in range(len(lista)):
                if lista[i]=='0':
                    lista[i]=""

            if Nro_Facturacion!=None:
                bloquearCb(self.cbTipo_Nota)
                bloquearCb(self.cbSerie)
                self.leNumero.clear()
                self.leFecha_Emision.clear()
                self.leEstado_Documento.clear()
                self.leEstado_Documento.setStyleSheet("")
                self.leEstado_Documento.setStyleSheet("background-color: rgb(255,255,255);")
                self.leURL.clear()
                self.cbMotivo.setEnabled(True)
                self.cbMotivo.setCurrentIndex(-1)
                self.leDescuento_Global.clear()
                self.leDescuento_Global_Soles.clear()
                self.cbTipo_Operacion.setEditable(False)
                self.leSerie_Numero.setText(Serie+"-"+Nro_Facturacion)
                self.botones()

            self.leRazon_Social.setText(lista[0])
            self.leDireccion.setText(lista[1])
            if Tipo_Comprobante=='2':
                if len(lista[2])==11:
                    lista[2]=lista[2][2:10]
            self.leRUC.setText(lista[2])
            self.leInterlocutor.setText(lista[3])
            self.leDNI.setText(lista[4])
            self.leCorreo.setText(lista[5])

            for k,v in dict_sunat_transaction.items():
                if int(lista[6])==v:
                    self.cbTipo_Operacion.setEditable(True)
                    font = QtGui.QFont()
                    font.setPointSize(11)
                    le = self.cbTipo_Operacion.lineEdit()
                    le.setFont(font)
                    self.cbTipo_Operacion.setCurrentText(k)

            self.leForma_Pago.setText(FormaPago[lista[7]])
            if FormaPago[lista[7]]=='CRÉDITO':
                self.pbVerCuotas.setEnabled(True)
            else:
                self.pbVerCuotas.setEnabled(False)

            self.leMoneda.setText(lista[8])

            if lista[9]=='0.00':
                lista[9]=""
            self.leDescuento_Global.setText(formatearDecimal(lista[9],'2'))

            CargarFactNota(sqlDetFact,self.tbwCotizacion_Cliente,self)
            self.cargarMontos()

            sqlVerificar="SELECT Nro_Cotización FROM TAB_VENTA_009_Cabecera_Facturacion WHERE Serie='%s';"%(self.cbSerie.currentText())
            guardadas=convlist(sqlVerificar)
            print(guardadas)
            if Serie+'-'+Nro_Facturacion in guardadas:
                self.pbGrabar.setEnabled(False)
                mensajeDialogo('informacion','Información','Esta Facturación ya fue Grabada')
            else:
                self.pbGrabar.setEnabled(True)

        except Exception as e:
            print(e)

    def Grabar(self):
        try:

            Fecha=datetime.datetime.now().strftime("%Y-%m-%d")
            Hora=datetime.datetime.now().strftime("%H:%M:%S.%f")

            tipo_nota=self.cbTipo_Nota.currentText()
            Serie=self.cbSerie.currentText()
            Motivo=self.cbMotivo.currentText()
            if len(tipo_nota)!=0 and len(Serie)!=0:
                Tipo_Comprobante=TipNota[tipo_nota]

                Nro_Cotización=self.leSerie_Numero.text()

                Cod_Cliente=dicCliente[self.leRazon_Social.text()]

                #Eliminar cuando funcione
                self.leCorreo.setReadOnly(False)

                Tipo_Operacion=dict_sunat_transaction[self.cbTipo_Operacion.currentText()]

                if self.leForma_Pago.text()=='CONTADO':
                    Forma_Pago='1'
                elif self.leForma_Pago.text()=='CRÉDITO':
                    Forma_Pago='2'

                Moneda=dicMoneda[self.leMoneda.text()]

                Descuento_Global=self.leDescuento_Global.text().replace(",","")

                if Tipo_Comprobante=='3':
                    Motivo_Nota=dict_tipo_de_nota_de_credito[Motivo]
                elif Tipo_Comprobante=='4':
                    Motivo_Nota=dict_tipo_de_nota_de_debito[Motivo]

                sqlCorrelativo="SELECT MAX(Nro_Facturacion) FROM TAB_VENTA_011_Cabecera_Notas WHERE Cod_Soc='%s' AND Año='%s' AND Tipo_Comprobante='%s' AND Serie='%s'"%(Cod_Soc, Año, Tipo_Comprobante, Serie)
                Nro_Actual=convlist(sqlCorrelativo)

                if Nro_Actual[0]==None:
                    Nro_Actual='1'
                    Num_Actual=Nro_Actual.zfill(8)
                else:
                    Nro_Actual=int(Nro_Actual[0])
                    Nro_Actual=str(Nro_Actual+1)
                    Num_Actual=Nro_Actual.zfill(8)

                self.leNumero.setText(Num_Actual)
                self.leNumero.setReadOnly(True)
                Nro_Facturacion=self.leNumero.text()

                Fecha_Emision=Fecha
                self.leFecha_Emision.setText(formatearFecha(Fecha_Emision))

                Estado_Factura='3'
                self.leEstado_Documento.setText(EstadoFactura['3'])

                sqlCabecera='''INSERT INTO TAB_VENTA_011_Cabecera_Notas(Cod_Soc, Año, Tipo_Comprobante, Serie, Nro_Facturacion, Fecha_Emision, Estado_Factura, Fecha_Vencimiento, Nro_Cotización, Cod_Cliente, Tipo_Operación, Forma_Pago, Moneda, Descuento_Global, Motivo_Nota, Fecha_Reg, Hora_Reg, Usuario_Reg)
                VALUES ('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')'''%(Cod_Soc, Año, Tipo_Comprobante, Serie, Nro_Facturacion, Fecha_Emision, Estado_Factura, Fecha_Emision, Nro_Cotización, Cod_Cliente, Tipo_Operacion, Forma_Pago, Moneda, Descuento_Global, Motivo_Nota, Fecha, Hora, Cod_Usuario)
                accion=ejecutarSql(sqlCabecera)

                if accion['respuesta']=='correcto':

                    d=self.tbwCotizacion_Cliente.rowCount()
                    for row in range(d):

                        #Detalle de Facturación
                        Item=self.tbwCotizacion_Cliente.item(row,0).text()
                        Cod_Material_Sunat=dicMatSUNAT[self.tbwCotizacion_Cliente.item(row,1).text()]
                        Cod_Material=dicMat[self.tbwCotizacion_Cliente.item(row,2).text()]
                        Marca=dicMarca[self.tbwCotizacion_Cliente.item(row,3).text()]
                        Unidad=self.tbwCotizacion_Cliente.item(row,4).text()
                        Cantidad=self.tbwCotizacion_Cliente.item(row,5).text().replace(",","")
                        Precio_sin_IGV=self.tbwCotizacion_Cliente.item(row,6).text().replace(",","")
                        Descuento_sin_IGV=self.tbwCotizacion_Cliente.item(row,8).text().replace(",","")
                        Precio_Final=self.tbwCotizacion_Cliente.item(row,10).text().replace(",","")
                        Sub_Total=self.tbwCotizacion_Cliente.item(row,11).text().replace(",","")

                        sqlDetalle='''INSERT INTO TAB_VENTA_012_Detalle_Notas(Cod_Soc, Año, Tipo_Comprobante, Serie, Nro_Facturacion, Item, Cod_Material, Cod_Material_Sunat, Marca, Unidad, Cantidad, Precio_sin_IGV, Descuento_sin_IGV, Precio_Final, Sub_Total, Fecha_Reg, Hora_Reg, Usuario_Reg)
                        VALUES ('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')'''%(Cod_Soc, Año, Tipo_Comprobante, Serie, Nro_Facturacion, Item, Cod_Material, Cod_Material_Sunat, Marca, Unidad, Cantidad, Precio_sin_IGV,Descuento_sin_IGV, Precio_Final, Sub_Total, Fecha, Hora, Cod_Usuario)
                        respuesta=ejecutarSql(sqlDetalle)

                    if respuesta['respuesta']=='correcto':
                        mensajeDialogo('informacion','Excelente','La ' + str(self.cbTipo_Nota.currentText()) + ' ' + str(self.cbSerie.currentText()) + '-' + str(self.leNumero.text()) + ' se grabo correctamente.')

                        self.pbGrabar.setEnabled(False)
                        if Cod_Soc=='1000'and Serie!='0001':
                            self.pbEnviar_SUNAT.setEnabled(True)
                        elif Cod_Soc=='1000'and Serie=='0001':
                            self.pbImprimir.setEnabled(True)

                    elif respuesta['respuesta']=='incorrecto':
                        mensajeDialogo('error','Error','Ocurrio un problema, comuniquese con soporte')

                elif accion['respuesta']=='incorrecto':
                    mensajeDialogo('error','Error','No se pudo Grabar, Vuelva a intentarlo')

            elif len(tipo_nota)!=0 and len(Serie)==0:
                mensajeDialogo('informacion','Información','Seleccione Serie')

            else:
                mensajeDialogo('informacion','Información','Seleccione Tipo de Comprobante y la Serie')

        except Exception as e:
            mensajeDialogo('error','Error','No se pudo Grabar, verifique los datos y vuelva a intentarlo')
            print(e)

    def SeleccionarNota(self):
        global SerieNota,NroNota
        SerieNota=None
        NroNota=None

        Tipnota=self.cbTipo_Nota.currentText()
        Serienota=self.cbSerie.currentText()
        if len(Tipnota)!=0 and len(Serienota)!=0:
            TipoNota=TipNota[Tipnota]
            SeleccionarNota(TipoNota,Serienota).exec_()
            self.CargarNota()
        elif len(Tipnota)!=0 and len(Serienota)==0:
            mensajeDialogo('informacion','Información','Seleccione Serie')
        else:
            mensajeDialogo('informacion','Información','Seleccione Tipo de Comprobante y la Serie')

    def CargarNota(self):
        try:
            TipoNota=TipNota[self.cbTipo_Nota.currentText()]
            sqlCabNota='''SELECT a.Fecha_Emision,a.Estado_Factura,a.URL,a.Nro_Cotización, a.Motivo_Nota,b.Razon_social,b.Direcc_cliente,b.RUC,b.Representante_Cliente,b.DNI,b.Correo_Representante,a.Tipo_Operación,a.Forma_Pago,c.Descrip_moneda,a.Descuento_Global
            FROM TAB_VENTA_011_Cabecera_Notas a
            LEFT JOIN `TAB_COM_001_Maestro Clientes`b ON a.Cod_Cliente=b.Cod_cliente
            LEFT JOIN TAB_SOC_008_Monedas c ON a.Moneda=c.Cod_moneda
            WHERE a.Cod_Soc='%s' AND a.Año='%s' AND a.Tipo_Comprobante='%s' AND a.Serie='%s'AND a.Nro_Facturacion='%s';'''%(Cod_Soc,Año,TipoNota, SerieNota,NroNota)
            lista=convlist(sqlCabNota)

            sqlDetNota='''SELECT a.Item,c.Descrip_SUNAT, b.Descrip_Mat, d.Descrip_Marca, c.Unidad_SUNAT, a.Cantidad, a.Precio_sin_IGV, a.Descuento_sin_IGV, a.Precio_Final, a.Sub_Total, SUM(e.Stock_disponible), SUM(e.Stock_Bloq_con_QA),a.Cod_Material,b.Cod_Prod_SUNAT
            FROM TAB_VENTA_012_Detalle_Notas a
            LEFT JOIN TAB_MAT_001_Catalogo_Materiales b ON b.Cod_Soc=a.Cod_Soc AND b.Cod_Mat=a.Cod_Material
            LEFT JOIN TAB_SOC_026_Tabla_productos_SUNAT c ON b.Cod_Prod_SUNAT=c.Cod_Sunat
            LEFT JOIN TAB_MAT_010_Marca_de_Producto d ON a.Marca=d.Cod_Marca
            LEFT JOIN TAB_MAT_002_Stock_Almacen e ON a.Cod_Soc=e.Cod_Soc AND a.Cod_Material=e.Cod_Mat
            WHERE a.Cod_Soc='%s' AND a.Año='%s' AND a.Tipo_Comprobante='%s' AND a.Serie='%s' AND a.Nro_Facturacion='%s'
            GROUP BY e.Cod_Mat ORDER BY a.Item ASC;'''%(Cod_Soc,Año,TipoNota,SerieNota,NroNota)

            for i in range(len(lista)):
                if lista[i]=='0':
                    lista[i]=""

            if SerieNota!=None and NroNota!=None:
                bloquearCb(self.cbTipo_Nota)
                bloquearCb(self.cbSerie)
                bloquearCb(self.cbMotivo)
                bloquearCb(self.cbTipo_Operacion)
                bloquearLe(self.leDescuento_Global)
                self.leDescuento_Global_Soles.clear()
                self.leNumero.setText(NroNota)
                self.leEstado_Documento.setStyleSheet("")
                self.leEstado_Documento.setStyleSheet("background-color: rgb(255,255,255);")
                self.botones()


            self.leFecha_Emision.setText(formatearFecha(lista[0]))
            self.leEstado_Documento.setText(EstadoFactura[lista[1]])
            self.leURL.setText(lista[2])

            self.leSerie_Numero.setText(lista[3])
            self.cbMotivo.setCurrentIndex(int(lista[4])-1)

            self.leRazon_Social.setText(lista[5])
            self.leDireccion.setText(lista[6])
            self.leRUC.setText(lista[7])
            self.leInterlocutor.setText(lista[8])
            self.leDNI.setText(lista[9])
            self.leCorreo.setText(lista[10])

            for k,v in dict_sunat_transaction.items():
                if int(lista[11])==v:
                    self.cbTipo_Operacion.setEditable(True)
                    font = QtGui.QFont()
                    font.setPointSize(11)
                    le = self.cbTipo_Operacion.lineEdit()
                    le.setFont(font)
                    self.cbTipo_Operacion.setCurrentText(k)

            self.leForma_Pago.setText(FormaPago[lista[12]])
            if FormaPago[lista[12]]=='CRÉDITO':
                self.pbVerCuotas.setEnabled(True)
            else:
                self.pbVerCuotas.setEnabled(False)

            self.leMoneda.setText(lista[13])

            if lista[14]=='0.00':
                lista[14]=""
            self.leDescuento_Global.setText(formatearDecimal(lista[14],'2'))

            if Cod_Soc=='1000'and self.cbSerie.currentText()!='0001':

                self.pbImprimir.setEnabled(False)

                if lista[1]=="1":
                    self.leEstado_Documento.setStyleSheet("color: rgb(255,130,130);\n""background-color: rgb(255,255,255);")
                    self.pbConsulta_Anulacion.setEnabled(True)

                elif lista[1]=="3":
                    if lista[2]=="":
                        self.pbEnviar_SUNAT.setEnabled(True)
                    else:
                        self.pbAbrirPDF.setEnabled(True)
                        self.pbAnular_Factura.setEnabled(True)

            elif Cod_Soc=='1000'and self.cbSerie.currentText()=='0001':

                self.pbImprimir.setEnabled(True)

                if lista[1]=="1":
                    self.leEstado_Documento.setStyleSheet("color: rgb(255,130,130);\n""background-color: rgb(255,255,255);")
                    self.pbConsulta_Anulacion.setEnabled(True)

                elif lista[1]=="3":
                    if lista[2]=="":
                        self.pbEnviar_SUNAT.setEnabled(True)
                    else:
                        self.pbAbrirPDF.setEnabled(True)
                        self.pbAnular_Factura.setEnabled(True)

            CargarFact(sqlDetNota,self.tbwCotizacion_Cliente,self)
            self.cargarMontos()

        except Exception as e:
            print(e)

    def botonEnviarSunat(self):
        if self.leMoneda.text()=="Dolar estadounidense":
            textoMoneda="DOLARES"

        tipo_de_comprobante=self.cbTipo_Nota.currentText()
        serie=self.cbSerie.currentText()
        numero=self.leNumero.text()
        sunat_transaction=dict_sunat_transaction[self.cbTipo_Operacion.currentText()]
        cliente_tipo_de_documento=tipoDocumento(self.leRUC.text(), self)
        ruc=self.leRUC.text()
        razonSocial=self.leRazon_Social.text()
        direccion=self.leDireccion.text()
        correo=self.leCorreo.text()
        fechaEmision=self.leFecha_Emision.text()
        fechaVencimiento=self.leFecha_Emision.text()
        moneda=dict_moneda[textoMoneda]
        tipoDeCambio=float(self.leTipo_Cambio.text())
        if textoMoneda=="DOLARES":
            if len(self.leDescuento_Global.text())!=0:
                descuento_global=float(self.leDescuento_Global.text().replace(",",""))
            else:
                descuento_global=float('0.00')
            total_descuento=descuento_global+float(self.leDescuento_Item.text().replace(",",""))
            totalGravada=float(self.leTotal_SinIGV.text().replace(",",""))
            totalIgv=float(self.leIGV.text().replace(",",""))
            total=float(self.leTotal.text().replace(",",""))
        else:
            if len(self.leDescuento_Global_Soles.text())!=0:
                descuento_global=float(self.leDescuento_Global_Soles.text().replace(",",""))
            else:
                descuento_global=float('0.00')
            total_descuento=descuento_global+float(self.leDescuento_Item_Soles.text().replace(",",""))
            totalGravada=float(self.leTotal_SinIGV_Soles.text().replace(",",""))
            totalIgv=float(self.leIGV_Soles.text().replace(",",""))
            total=float(self.leTotal_Soles.text().replace(",",""))


        condicionesDePago="CONTADO"
        tbwCuotas=[]
        orden_compra_servicio=''
        tipo_de_igv=dict_tipo_de_igv["Gravado - Operación Onerosa"]
        guia_tipo=dict_guia_tipo["GUÍA DE REMISIÓN REMITENTE"]
        ##############################################################
        for k,v in TipSerie.items():
            if serie in v:
                TipoComp=k

        tipo_de_nota_de_credito=[]
        tipo_de_nota_de_debito=[]

        if tipo_de_comprobante=='NOTA DE CRÉDITO':
            tipo_de_nota_de_credito.append(dict_tipo_de_comprobante[TipoComp]) # documento_que_se_modifica_tipo
            tipo_de_nota_de_credito.append(self.leSerie_Numero.text().split("-")[0]) # documento_que_se_modifica_serie
            tipo_de_nota_de_credito.append(int(self.leSerie_Numero.text().split("-")[1])) # documento_que_se_modifica_numero
            tipo_de_nota_de_credito.append(dict_tipo_de_nota_de_credito[self.cbMotivo.itemText(self.cbMotivo.currentIndex())]) # tipo_de_nota_de_credito

        elif tipo_de_comprobante=='NOTA DE DÉBITO':
            tipo_de_nota_de_debito.append(dict_tipo_de_comprobante[TipoComp]) # documento_que_se_modifica_tipo
            tipo_de_nota_de_debito.append(self.leSerie_Numero.text().split("-")[0]) # documento_que_se_modifica_serie
            tipo_de_nota_de_debito.append(int(self.leSerie_Numero.text().split("-")[1])) # documento_que_se_modifica_numero
            tipo_de_nota_de_debito.append(dict_tipo_de_nota_de_debito[self.cbMotivo.itemText(self.cbMotivo.currentIndex())]) # tipo_de_nota_de_debito
        #############################################################
        Doc=tipo_de_comprobante.lower()
        nombreArchivo = Doc.capitalize() + " %s-%s" % (serie, numero)

        generarDocumento(Cod_Soc,Año,tipo_de_comprobante, serie, numero, sunat_transaction, cliente_tipo_de_documento, ruc, razonSocial, direccion, correo, fechaEmision, fechaVencimiento, moneda, tipoDeCambio, descuento_global, total_descuento, totalGravada, totalIgv, total, condicionesDePago, tbwCuotas, orden_compra_servicio, tipo_de_igv, guia_tipo, nombreArchivo, self.tbwCotizacion_Cliente, tipo_de_nota_de_credito, tipo_de_nota_de_debito, self)

    def botonAbrirPdf(self):
        tipo_de_comprobante=TipNota[self.cbTipo_Nota.currentText()]
        serie=self.cbSerie.currentText()
        numero=self.leNumero.text()
        consultarDocumento(Cod_Soc,Año,tipo_de_comprobante, serie, numero, self)

    def botonAnular(self):
        if self.leNumero.text()=="":
            mensajeDialogo("informacion", "Anular", "Seleccionar factura")
        else:
            resultado=mensajeDialogo("pregunta","Anular","¿Seguro que desea anular la factura " + str(self.cbSerie.currentText()) + "-" + str(self.leNumero.text()) + "?")
            if resultado == 'Yes':
                Dialogo=IngresarMotivo("Motivo", "Ingrese el motivo de la anulación", 100)
                Dialogo.exec_()
                tipo_de_comprobante=self.cbTipo_Nota.currentText()
                serie=self.cbSerie.currentText()
                numero=self.leNumero.text()
                motivo=Dialogo.texto
                if motivo!="":
                    anularDocumento(Cod_Soc, Año, tipo_de_comprobante, serie, numero, motivo, self)
                    self.pbAbrirPDF.setEnabled(False)
                    self.pbAnular_Factura.setEnabled(False)
                    self.pbConsulta_Anulacion.setEnabled(True)

    def botonConsultaAnulacion(self):
        if self.leNumero.text()=="":
            mensajeDialogo("informacion", "Anular", "Seleccionar factura")
        else:
            tipo_de_comprobante=TipNota[self.cbTipo_Nota.currentText()]
            serie=self.cbSerie.currentText()
            numero=self.leNumero.text()
            validarAnulacion(Cod_Soc, Año, tipo_de_comprobante, serie, numero, self)

    def nota_credito_o_debito(self):
        if self.cbTipo_Nota.currentText() == 'NOTA DE CRÉDITO':
            self.notaCreditoManual()
        if self.cbTipo_Nota.currentText() == 'NOTA DE DÉBITO':
            self.notaDebitoManual()

    def notaCreditoManual(self):
        print('Impresión Manual Nota de Crédito')
        reply = mensajeDialogo('pregunta','Nota de Crédito',"¿Desea Imprimir la Nota de Crédito?")
        if reply == 'Yes':
            # Armar lista de listas
            list_productos = []
            for i in range(self.tbwCotizacion_Cliente.rowCount()):
                list_temporal = []
                for j in range(self.tbwCotizacion_Cliente.columnCount()):
                    if j==1:
                        dato = self.tbwCotizacion_Cliente.item(i,j).text()
                        try:
                            descrip_breve = dict_descripBreve[dato]
                        except:
                            descrip_breve = ''
                        list_temporal.append(descrip_breve)
                    elif j==7 or j==11:
                        dato = self.tbwCotizacion_Cliente.item(i,j).text()
                        list_temporal.append(dato)
                    elif j==5:
                        dato = self.tbwCotizacion_Cliente.item(i,j).text()
                        list_temporal.insert(0,dato)
                list_productos.append(list_temporal)
            # print(list_productos)

            pdf = FPDF('L', 'mm', (161, 220))
            pdf.set_auto_page_break(False)
            pdf.add_page()
            pdf.set_xy(0,0)
            # reply = mensajeDialogo('pregunta','Nota de Crédito', '¿Mostrar Imagen Referencial?')
            # if reply == 'Yes':
            #     pdf.image('Documentos/credito_21,95x16,23.png', 0,0, 219.5,162.3) # 220 x 161
            pdf.set_font('Helvetica', '', 8.5)
            pdf.set_fill_color(255, 255, 255)
            pdf.set_text_color(0, 0, 255)

            def dataComprobante(pdf):
                #### Datos Cliente
                razon_social = self.leRazon_Social.text()
                ruc = self.leRUC.text()

                fecha_texto = formatoFechaTexto(datetime.datetime.now()).split(" ")
                dia, mes_letras, año = fecha_texto[0], fecha_texto[2], fecha_texto[-1][-1]
                alto = 2.5
                desfase_up = 56 # Mide 56 mm
                desfase_left = 19 # Mide 22 mm
                pdf.set_xy(desfase_left, desfase_up)
                pdf.cell(7, alto, dia, 0, 0, 'C')
                pdf.cell(5, alto, '', 0, 0, 'C')
                pdf.cell(21, alto, mes_letras, 0, 0, 'C')
                pdf.cell(15, alto, '', 0, 0, 'C')
                pdf.cell(3, alto, año, 0, 0, 'C')
                pdf.set_xy(desfase_left+10, desfase_up+6.5)
                pdf.cell(123, alto, razon_social, 0, 0, 'L')
                pdf.cell(14, alto, '', 0, 0, 'C')
                pdf.cell(38, alto, ruc, 0, 2, 'L')

                ##### Fecha Cancelado
                alto = 2.5
                desfase_up = 148.5 # Mide 149.5 mm
                desfase_left = 77 # Mide 80 mm
                pdf.set_xy(desfase_left, desfase_up)
                pdf.cell(5, alto, dia, 0, 0, 'C')
                pdf.cell(5, alto, '', 0, 0, 'C')
                pdf.cell(23, alto, mes_letras, 0, 0, 'C')
                pdf.cell(12, alto, '', 0, 0, 'C')
                pdf.cell(2, alto, año, 0, 0, 'C')

                ##### Comprobante Relacionado
                nro_comprobante = self.leSerie_Numero.text()
                alto = 3
                desfase_up = 131.5 # Mide 132 mm
                desfase_left = 25.25 # Mide 30.75 mm
                pdf.set_xy(desfase_left, desfase_up)
                pdf.cell(123, alto, 'Nro. Factura / Boleta: ' + nro_comprobante, 0, 0, 'L')

            dataComprobante(pdf)

            def montoComprobante(pdf, total_f):
                ##### Monto Letras
                total_letras = numero_a_moneda(total_f).upper()
                alto = 3
                desfase_up = 135 # Mide 135 mm
                desfase_left = 25.25 # Mide 28.25 mm
                pdf.set_xy(desfase_left, desfase_up)
                pdf.cell(124, alto, total_letras, 0, 0, 'L')

                ##### Pago Total
                igv_f = total_f * 0.18
                sub_total_f = total_f - igv_f
                igv = formatearDecimal(str(igv_f),'2')
                sub_total = formatearDecimal(str(sub_total_f),'2')
                total = formatearDecimal(str(total_f),'2')
                alto = 5.75
                desfase_up = 137.75 # Mide 130.5 mm
                desfase_left = 175.75 # Mide 178.75 mm
                pdf.set_xy(desfase_left, desfase_up)
                pdf.cell(28, alto, '$ '+sub_total, 0, 2, 'R')
                pdf.cell(28, alto, '$ '+igv, 0, 2, 'R')
                pdf.cell(28, alto+0.25, '$ '+total, 0, 2, 'R')

            ##### Tabla
            desfase_up = 80 # Mide 80 mm
            desfase_left = 8.5 # Mide 11.5 mm
            pdf.set_xy(desfase_left, desfase_up)
            total_float = 0

            for lista in list_productos:
                largo_1, largo_2, largo_3, largo_4 = 16.75, 124, 26.5, 28
                alto = 2.75

                if list_productos.index(lista) != 0:
                    pdf.cell(-1.5)
                    num_items = 8 # Numero de Item´s por tabla
                    if list_productos.index(lista) % num_items == 0:
                        print('Salto de Página')
                        montoComprobante(pdf, total_float)
                        #### Nueva PAGINA
                        pdf.accept_page_break()
                        pdf.add_page()
                        dataComprobante(pdf)
                        total_float = 0
                        ##### Nueva Tabla
                        desfase_up = 80 # Mide 80 mm
                        desfase_left = 8.5 # Mide 11.5 mm
                        pdf.set_xy(desfase_left, desfase_up)

                valor_y = pdf.get_y()
                valor_x1 = pdf.get_x()
                pdf.multi_cell(largo_1, alto, lista[0], 0, 'C')
                fin_y = pdf.get_y()

                valor_x2 = valor_x1 + largo_1
                pdf.set_xy(valor_x2, valor_y)
                pdf.multi_cell(largo_2, alto, lista[1], 0, 'L')
                if pdf.get_y() > fin_y:
                    fin_y = pdf.get_y()
                else:
                    fin_y = fin_y

                valor_x3 = valor_x2 + largo_2
                pdf.set_xy(valor_x3, valor_y)
                pdf.multi_cell(largo_3, alto, '$ '+lista[2], 0, 'R')
                if pdf.get_y() > fin_y:
                    fin_y = pdf.get_y()
                else:
                    fin_y = fin_y

                valor_x4 = valor_x3 + largo_3
                pdf.set_xy(valor_x4, valor_y)
                pdf.multi_cell(largo_4, alto, '$ '+lista[3], 0, 'R')
                if pdf.get_y() > fin_y:
                    fin_y = pdf.get_y()
                else:
                    fin_y = fin_y

                pdf.set_y(fin_y + alto)

                total_float += float(lista[3].replace(',',''))

            montoComprobante(pdf, total_float)

            root = tk.Tk()
            root.withdraw()
            ruta = filedialog.asksaveasfilename(initialfile = 'NOTA_CREDITO_'+ self.cbSerie.currentText() + '_' + self.leNumero.text(), defaultextension=".pdf", filetypes=(("Documento PDF", "*.pdf"),))
            if ruta != "":
                pdf.output(ruta, 'F')
                mensajeDialogo('informacion','Reporte','Reporte PDF Generado con éxito')

    def notaDebitoManual(self):
        print('Impresión Manual Nota de Débito')
        reply = mensajeDialogo('pregunta','Nota de Débito',"¿Desea Imprimir la Nota de Débito?")
        if reply == 'Yes':
            list_productos = []
            for i in range(self.tbwCotizacion_Cliente.rowCount()):
                list_temporal = []
                for j in range(self.tbwCotizacion_Cliente.columnCount()):
                    if j==1:
                        dato = self.tbwCotizacion_Cliente.item(i,j).text()
                        try:
                            descrip_breve = dict_descripBreve[dato]
                        except:
                            descrip_breve = ''
                        list_temporal.append(descrip_breve)
                    elif j==7 or j==11:
                        dato = self.tbwCotizacion_Cliente.item(i,j).text()
                        list_temporal.append(dato)
                    elif j==5:
                        dato = self.tbwCotizacion_Cliente.item(i,j).text()
                        list_temporal.insert(0,dato)
                list_productos.append(list_temporal)
            # print(list_productos)

            pdf = FPDF('L', 'mm', (161, 220)) # 220 x 161
            pdf.set_auto_page_break(False)
            pdf.add_page()
            pdf.set_xy(0,0)
            # reply = mensajeDialogo('pregunta','Nota de Débito','¿Mostrar Imagen Referencial?')
            # if reply == 'Yes':
            #     pdf.image('Documentos/debito_21,81x15,76.png', 0,0, 218.1,157.6) # 220 x 161
            pdf.set_font('Helvetica', '', 8.5)
            pdf.set_fill_color(255, 255, 255)
            pdf.set_text_color(0, 0, 255)

            def dataComprobante(pdf):
                #### Datos Cliente
                razon_social = self.leRazon_Social.text()
                ruc = self.leRUC.text()
                fecha_texto = formatoFechaTexto(datetime.datetime.now()).split(" ")
                dia, mes_letras, año = fecha_texto[0], fecha_texto[2], fecha_texto[-1][-1]
                alto = 2.5
                desfase_up = 53.5 # Mide 56.5 mm
                desfase_left = 19  # Mide 23 mm
                pdf.set_xy(desfase_left, desfase_up)
                pdf.cell(12, alto, dia, 0, 0, 'C')
                pdf.cell(5, alto, '', 0, 0, 'C')
                pdf.cell(34, alto, mes_letras, 0, 0, 'C')
                pdf.cell(13, alto, '', 0, 0, 'C')
                pdf.cell(3, alto, año, 0, 0, 'C')
                pdf.set_xy(desfase_left+10, desfase_up+7)
                pdf.cell(123, alto, razon_social, 0, 0, 'L')
                pdf.cell(14, alto, '', 0, 0, 'C')
                pdf.cell(38, alto, ruc, 0, 2, 'L')

                ##### Fecha Cancelado
                alto = 2.5
                desfase_up = 142.25 # Mide 145 mm
                desfase_left = 77 # Mide 81 mm
                pdf.set_xy(desfase_left, desfase_up)
                pdf.cell(6, alto, dia, 0, 0, 'C')
                pdf.cell(5, alto, '', 0, 0, 'C')
                pdf.cell(25, alto, mes_letras, 0, 0, 'C')
                pdf.cell(12, alto, '', 0, 0, 'C')
                pdf.cell(2, alto, año, 0, 0, 'C')

                ##### Comprobante Relacionado
                nro_comprobante = self.leSerie_Numero.text()
                alto = 3
                desfase_up = 127 # Mide 127.5 mm
                desfase_left = 26.5 # Mide 30.75 mm
                pdf.set_xy(desfase_left, desfase_up)
                pdf.cell(123, alto, 'Nro. Factura / Boleta: ' + nro_comprobante, 0, 0, 'L')

            dataComprobante(pdf)

            def montoComprobante(pdf, total_f):
                ##### Monto Letras
                total_letras = numero_a_moneda(total_f).upper()
                alto = 3
                desfase_up = 130.5 # Mide 133 mm
                desfase_left = 26.5 # Mide 30.75 mm
                pdf.set_xy(desfase_left, desfase_up)
                pdf.cell(123, alto, total_letras, 0, 0, 'L')

                ##### Pago Total
                igv_f = total_f * 0.18
                sub_total_f = total_f - igv_f
                igv = formatearDecimal(str(igv_f),'2')
                sub_total = formatearDecimal(str(sub_total_f),'2')
                total = formatearDecimal(str(total_f),'2')
                alto = 5.5
                desfase_up = 133.5 # Mide 136 mm
                desfase_left = 176 # Mide 180.25 mm
                pdf.set_xy(desfase_left, desfase_up)
                pdf.cell(28, alto, '$ '+sub_total, 0, 2, 'R')
                pdf.cell(28, alto, '$ '+igv, 0, 2, 'R')
                pdf.cell(28, alto+0.5, '$ '+total, 0, 2, 'R')

            ##### Tabla
            desfase_up = 80 # Mide 82 mm
            desfase_left = 11 # Mide 14.25 mm -- Decía 10.5 +3(de la impresora)
            pdf.set_xy(desfase_left, desfase_up)

            total_float = 0
            for lista in list_productos:
                largo_1, largo_2, largo_3, largo_4 = 16.75, 122.75, 26.25, 28
                alto = 2.75

                if list_productos.index(lista) != 0:
                    pdf.cell(1)
                    num_items = 8 # Numero de Item´s por tabla
                    if list_productos.index(lista) % num_items == 0:
                        print('Salto de Página')
                        montoComprobante(pdf, total_float)
                        #### Nueva PAGINA
                        pdf.accept_page_break()
                        pdf.add_page()
                        dataComprobante(pdf)
                        total_float = 0
                        ##### Nueva Tabla
                        desfase_up = 80 # Mide 82 mm
                        desfase_left = 11 # Mide 14.25 mm -- Decía 10.5 +3(de la impresora)
                        pdf.set_xy(desfase_left, desfase_up)

                valor_y = pdf.get_y()
                valor_x1 = pdf.get_x()
                pdf.multi_cell(largo_1, alto, lista[0], 0, 'C')
                fin_y = pdf.get_y()

                valor_x2 = valor_x1 + largo_1
                pdf.set_xy(valor_x2, valor_y)
                pdf.multi_cell(largo_2, alto, lista[1], 0, 'L')
                if pdf.get_y() > fin_y:
                    fin_y = pdf.get_y()
                else:
                    fin_y = fin_y

                valor_x3 = valor_x2 + largo_2
                pdf.set_xy(valor_x3, valor_y)
                pdf.multi_cell(largo_3, alto, '$ '+lista[2], 0, 'R')
                if pdf.get_y() > fin_y:
                    fin_y = pdf.get_y()
                else:
                    fin_y = fin_y

                valor_x4 = valor_x3 + largo_3
                pdf.set_xy(valor_x4, valor_y)
                pdf.multi_cell(largo_4, alto, '$ '+lista[3], 0, 'R')
                if pdf.get_y() > fin_y:
                    fin_y = pdf.get_y()
                else:
                    fin_y = fin_y

                pdf.set_y(fin_y + alto)

                total_float += float(lista[3].replace(',',''))

            montoComprobante(pdf, total_float)

            root = tk.Tk()
            root.withdraw()
            ruta = filedialog.asksaveasfilename(initialfile = 'NOTA_DEBITO_' + self.cbSerie.currentText() + '_' + self.leNumero.text(), defaultextension=".pdf", filetypes=(("Documento PDF", "*.pdf"),))
            if ruta != "":
                pdf.output(ruta, 'F')
                mensajeDialogo('informacion','Reporte','Reporte PDF Generado con éxito')

    def Salir(self):
        self.close()

if __name__ == '__main__':
    app=QApplication(sys.argv)
    _main=ERP_Facturacion_Notas()
    _main.showMaximized()
    app.exec_()
