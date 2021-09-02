# coding=utf-8
dict_tipo_de_comprobante={}
dict_tipo_de_comprobante["FACTURA"]=1
dict_tipo_de_comprobante["BOLETA"]=2
dict_tipo_de_comprobante["NOTA DE CRÉDITO"]=3
dict_tipo_de_comprobante["NOTA DE DÉBITO"]=4
dict_tipo_de_comprobante["GUÍA"]=7

dict_sunat_transaction={}
dict_sunat_transaction["VENTA INTERNA"] = 1
dict_sunat_transaction["EXPORTACIÓN"] = 2
dict_sunat_transaction["VENTA INTERNA – ANTICIPOS"] = 4
dict_sunat_transaction["VENTAS NO DOMICILIADOS QUE NO CALIFICAN COMO   EXPORTACIÓN."] = 29
dict_sunat_transaction["OPERACIÓN SUJETA A DETRACCIÓN."] = 30
dict_sunat_transaction["DETRACCIÓN - SERVICIOS DE TRANSPORTE CARGA "] = 33
dict_sunat_transaction["OPERACIÓN SUJETA A PERCEPCIÓN"] = 34
dict_sunat_transaction["DETRACCIÓN - SERVICIOS DE TRANSPORTE DE PASAJEROS."] = 32
dict_sunat_transaction["DETRACCIÓN - RECURSOS HIDROBIOLÓGICOS"] = 31

dict_cliente_tipo_de_documento={}
dict_cliente_tipo_de_documento["RUC - REGISTRO ÚNICO DE CONTRIBUYENTE"] = "6"
dict_cliente_tipo_de_documento["DNI - DOC. NACIONAL DE IDENTIDAD"] = "1"
dict_cliente_tipo_de_documento["VARIOS - VENTAS MENORES A S/.700.00 Y OTROS"] = "-"
dict_cliente_tipo_de_documento["CARNET DE EXTRANJERÍA"] = "4"
dict_cliente_tipo_de_documento["PASAPORTE"] = "7"
dict_cliente_tipo_de_documento["CÉDULA DIPLOMÁTICA DE IDENTIDAD"] = "A"
dict_cliente_tipo_de_documento["NO DOMICILIADO SIN RUC (EXPORTACIÓN)"] = "0"

# robles.tim@gmail.com / rpaniura@multiplay.com.pe

dict_moneda={}
dict_moneda["SOLES"] = 1
dict_moneda["DOLARES"] = 2
dict_moneda["EUROS"] = 3

dict_tipo_de_igv={}
dict_tipo_de_igv["Gravado - Operación Onerosa"] = 1
dict_tipo_de_igv["Gravado – Retiro por premio"] = 2
dict_tipo_de_igv["Gravado – Retiro por donación"] = 3
dict_tipo_de_igv["Gravado – Retiro"] = 4
dict_tipo_de_igv["Gravado – Retiro por publicidad"] = 5
dict_tipo_de_igv["Gravado – Bonificaciones"] = 6
dict_tipo_de_igv["Gravado – Retiro por entrega a trabajadores"] = 7
dict_tipo_de_igv["Exonerado - Operación Onerosa"] = 8
dict_tipo_de_igv["Inafecto - Operación Onerosa"] = 9
dict_tipo_de_igv["Inafecto – Retiro por Bonificación"] = 10
dict_tipo_de_igv["Inafecto – Retiro"] = 11
dict_tipo_de_igv["Inafecto – Retiro por Muestras Médicas"] = 12
dict_tipo_de_igv["Inafecto - Retiro por Convenio Colectivo"] = 13
dict_tipo_de_igv["Inafecto – Retiro por premio"] = 14
dict_tipo_de_igv["Inafecto - Retiro por publicidad"] = 15
dict_tipo_de_igv["Exportación"] = 16
dict_tipo_de_igv["Exonerado - Transferencia Gratuita"] = 17

dict_guia_tipo={}
dict_guia_tipo["GUÍA DE REMISIÓN REMITENTE"] = 1
dict_guia_tipo["GUÍA DE REMISIÓN TRANSPORTISTA"] = 2

dict_tipo_de_nota_de_credito={}
dict_tipo_de_nota_de_credito["ANULACIÓN DE LA OPERACIÓN"] = 1
dict_tipo_de_nota_de_credito["ANULACIÓN POR ERROR EN EL RUC"] = 2
dict_tipo_de_nota_de_credito["CORRECCIÓN POR ERROR EN LA DESCRIPCIÓN"] = 3
dict_tipo_de_nota_de_credito["DESCUENTO GLOBAL"] = 4
dict_tipo_de_nota_de_credito["DESCUENTO POR ÍTEM"] = 5
dict_tipo_de_nota_de_credito["DEVOLUCIÓN TOTAL"] = 6
dict_tipo_de_nota_de_credito["DEVOLUCIÓN POR ÍTEM"] = 7
dict_tipo_de_nota_de_credito["BONIFICACIÓN"] = 8
dict_tipo_de_nota_de_credito["DISMINUCIÓN EN EL VALOR"] = 9
dict_tipo_de_nota_de_credito["OTROS CONCEPTOS"] = 10
dict_tipo_de_nota_de_credito["AJUSTES AFECTOS AL IVAP"] = 11
dict_tipo_de_nota_de_credito["AJUSTES DE OPERACIONES DE EXPORTACIÓN"] = 12

nota_credito_ingreso_cuentas=[1,2,4,5,6,7,8,9,10,11,12]
nota_credito_ingreso_producto=[1,2,6,7]

dict_tipo_de_nota_de_debito={}
dict_tipo_de_nota_de_debito["INTERESES POR MORA"] = 1
dict_tipo_de_nota_de_debito["AUMENTO DE VALOR"] = 2
dict_tipo_de_nota_de_debito["PENALIDADES"] = 3
dict_tipo_de_nota_de_debito["AJUSTES AFECTOS AL IVAP"] = 4
dict_tipo_de_nota_de_debito["AJUSTES DE OPERACIONES DE EXPORTACIÓN"] = 5

dict_motivo_de_traslado={}
dict_motivo_de_traslado["VENTA"]="01"
dict_motivo_de_traslado["VENTA SUJETA A CONFIRMACION DEL COMPRADOR"]="14"
dict_motivo_de_traslado["COMPRA"]="02"
dict_motivo_de_traslado["TRASLADO ENTRE ESTABLECIMIENTOS DE LA MISMA EMPRESA"]="04"
dict_motivo_de_traslado["TRASLADO EMISOR ITINERANTE CP"]="18"
dict_motivo_de_traslado["IMPORTACION"]="08"
dict_motivo_de_traslado["EXPORTACION"]="09"
dict_motivo_de_traslado["TRASLADO A ZONA PRIMARIA"]="19"
dict_motivo_de_traslado["OTROS"]="13"
