#coding:utf-8
from datetime import datetime
import sys,time,locale,traceback
sys.path.append("..\\..")
from libreria.conexionBD import *

class Ejecutor(object):
	def __init__(self,listaParametros,codSolic,rutaArchivo):
			#print "Metodo Constructor del Ejecutor Detalle General Precierre"
			try:
				query = "SELECT A.C_SOLICITUD, A.C_TIPO_CDR, A.C_OPERADORA,A.F_DESDE,A.F_HASTA,"
				query +="A.C_DIRECCION_CONTABLE,A.C_MONEDA,B.DNIO_SIGNIFICADO AS 'DIRECCION_CONTABLE',C.AB_MONEDA "
				query += "FROM ICX_SOLIC_PRECIERRE A "
				query +="INNER JOIN ICX_DOMINIO B ON B.DNIO_VALOR = A.C_DIRECCION_CONTABLE "
				query +="INNER JOIN ICX_MONEDA C ON C.C_MONEDA= A.C_MONEDA "
				query +=" WHERE A.C_SOLICITUD = " + str(listaParametros["Cod Solicitud Precierre"])
				query +=" and B.DNIO_NOMBRE = 'DNIO_DIRECCION_CONTABLE'"

				solicPrecierre = ejecutarQuery(query)
				solicPrecierre = solicPrecierre[0]

				query = "SELECT * FROM ICX_TRAFICO WHERE "
				query += "F_CDR BETWEEN '" + str(solicPrecierre["F_DESDE"])+ "' AND '" +str(solicPrecierre["F_HASTA"])
				query += "' AND TAS_CASO_TRAFICO_OPERADORA = '" + str(solicPrecierre["C_OPERADORA"])
				query += "' AND TAS_CASO_TRAFICO_DIR_CONTABLE = " + str(solicPrecierre["C_DIRECCION_CONTABLE"])
				query += " AND TAS_LISTA_PRECIO_MONEDA = " + str(solicPrecierre["C_MONEDA"])

				#registros = ejecutarQuery(query)
				'''
				print query
				print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
				print registros
				print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
				print len(registros)'''
				fecha = datetime.now().strftime('%Y%m%d%H%M%S')
				nombreReporte = "ICX_"+str(solicPrecierre["C_SOLICITUD"]) +"_"+solicPrecierre["C_OPERADORA"]
				nombreReporte+= "_FDESDE_"+str(solicPrecierre["F_DESDE"].strftime('%Y%m%d'))+"_FHASTA_"
				nombreReporte+= str(solicPrecierre["F_HASTA"].strftime('%Y%m%d'))+ "_" + solicPrecierre["DIRECCION_CONTABLE"].replace(" ","_") + "_" + str(fecha) +".txt"

				reporte = open(rutaArchivo + nombreReporte,"w")
				reporte.write("REPORTE DE |TASACION PRECIERRE |DETALLADO GENERAL" + "\r\n")
				reporte.write("Operadora|Tipo CDR|Direccion Contable|Moneda|Solicitud PreCierre" + "\r\n")
 				valores = str(solicPrecierre["C_OPERADORA"])+"|"+str(solicPrecierre["C_TIPO_CDR"])+"|"
 				valores += str(solicPrecierre["DIRECCION_CONTABLE"])+"|"+str(solicPrecierre["AB_MONEDA"])+"|"
 				valores += str(solicPrecierre["C_SOLICITUD"])

 				reporte.write(valores + "\r\n\r\n")

				nombreCampos = "F_CDR|C_TIPO_CDR|NB_ARCHIVO|ANO|BNO|DURATION|IA_ROUTE_IN_EXT|IA_ROUTE_OUT_EXT|"
				nombreCampos += "IA_SERV_CLASS_EXT|IA_TC|TRANSDATETIME|USER_FIELD1|USER_FIELD2|USER_FIELD3|USER_FIELD4|"
				nombreCampos += "USER_FIELD5|PREP_RUTA_ENT_RUTA|PREP_RUTA_ENT_RUTA_TASACION|PREP_RUTA_ENT_T_RUTA|"
				nombreCampos += "PREP_RUTA_ENT_OPERADORA|PREP_RUTA_SAL_RUTA|PREP_RUTA_SAL_RUTA_TASACION|PREP_RUTA_SAL_T_RUTA|"
				nombreCampos += "PREP_RUTA_SAL_OPERADORA|PREP_ANO_PREFIX|PREP_BNO_PREFIX|"
				nombreCampos += "PREP_ANO_ZONA|PREP_BNO_ZONA|PREP_ANO_OPERADORA|PREP_BNO_OPERADORA|"
				nombreCampos += "PREP_ZR_DESTINO|PREP_ZR_DESTINO_TASACION|PREP_ANO_ORIGEN_EXCEP|PREP_ANO_ORIGEN|"
				nombreCampos += "PREP_ANO_ORIGEN_TASACION|TAS_CASO_TRAFICO_ID|TAS_CASO_TRAFICO_RUTA_TAS_ENT|"
				nombreCampos += "TAS_CASO_TRAFICO_RUTA_TAS_SAL|TAS_NRO_CDR|TAS_CASO_TRAFICO_OPERADORA|"
				nombreCampos += "TAS_CASO_TRAFICO_RITEM|TAS_CASO_TRAFICO_DIR_CONTABLE|TAS_CASO_TRAFICO_MET_CONTABLE|"
				nombreCampos += "TAS_CASO_TRAFICO_CLASIF|TAS_CASO_TRAFICO_CLASE_TARIFA_GRUPO|TAS_CASO_TRAFICO_CLASE_TARIFA|"
				nombreCampos += "TAS_CASO_TRAFICO_RITEMDET|TAS_CASO_TRAFICO_COD_CONTABLE|TAS_LISTA_PRECIO|TAS_LISTA_PRECIO_DET|"
				nombreCampos += "TAS_LISTA_PRECIO_RITEM|TAS_LISTA_PRECIO_TPRECIO|TAS_LISTA_PRECIO_QMONTO_ICX|"
				nombreCampos += "TAS_LISTA_PRECIO_QTASA_MIN|TAS_LISTA_PRECIO_QRED_SEG_UNID|TAS_LISTA_PRECIO_QRED_MIN_UNID|"
				nombreCampos += "TAS_LISTA_PRECIO_QRED_UNID_ADIC|TAS_LISTA_PRECIO_IRED_AJUSTE|TAS_LISTA_PRECIO_MONEDA|"
				nombreCampos += "TAS_LISTA_PRECIO_DURATION_A_FACT|TAS_LISTA_PRECIO_MONTO|I_RETASADO|Q_RETASADO|F_INSERCION|F_ULT_ACT"

				reporte.write(nombreCampos + "\r\n")

				for reg in ejecutarQuery_v2(50000,query):
					linea = str(reg["F_CDR"]) + "|"
					linea += str(reg["C_TIPO_CDR"]) + "|"
					linea += str(reg["NB_ARCHIVO"]) + "|"
					linea += str(reg["ANO"]) + "|"
					linea += str(reg["BNO"]) + "|"
					linea += str(reg["DURATION"]) + "|"
					linea += str(reg["IA_ROUTE_IN_EXT"]) + "|"
					linea += str(reg["IA_ROUTE_OUT_EXT"]) + "|"
					linea += str(reg["IA_SERV_CLASS_EXT"]) + "|"
					linea += str(reg["IA_TC"]) + "|"
					linea += str(reg["TRANSDATETIME"]) + "|"
					linea += '"' + str(reg["USER_FIELD1"]) + '"' + "|"
					linea += '"' + str(reg["USER_FIELD2"]) + '"' + "|"
					linea += '"' + str(reg["USER_FIELD3"]) + '"' + "|"
					linea += '"' + str(reg["USER_FIELD4"]) + '"' + "|"
					linea += '"' + str(reg["USER_FIELD5"]) + '"' + "|"
					linea += str(reg["PREP_RUTA_ENT_RUTA"]) + "|"
					linea += str(reg["PREP_RUTA_ENT_RUTA_TASACION"]) + "|"
					linea += str(reg["PREP_RUTA_ENT_T_RUTA"]) + "|"
					linea += str(reg["PREP_RUTA_ENT_OPERADORA"]) + "|"
					linea += str(reg["PREP_RUTA_SAL_RUTA"]) + "|"
					linea += str(reg["PREP_RUTA_SAL_RUTA_TASACION"]) + "|"
					linea += str(reg["PREP_RUTA_SAL_T_RUTA"]) + "|"
					linea += str(reg["PREP_RUTA_SAL_OPERADORA"]) + "|"
					linea += str(reg["PREP_ANO_PREFIX"]) + "|"
					linea += str(reg["PREP_BNO_PREFIX"]) + "|"
					linea += str(reg["PREP_ANO_ZONA"]) + "|"
					linea += str(reg["PREP_BNO_ZONA"]) + "|"
					linea += str(reg["PREP_ANO_OPERADORA"]) + "|"
					linea += str(reg["PREP_BNO_OPERADORA"]) + "|"
					linea += str(reg["PREP_ZR_DESTINO"]) + "|"
					linea += str(reg["PREP_ZR_DESTINO_TASACION"]) + "|"
					linea += str(reg["PREP_ANO_ORIGEN_EXCEP"]) + "|"
					linea += str(reg["PREP_ANO_ORIGEN"]) + "|"
					linea += str(reg["PREP_ANO_ORIGEN_TASACION"]) + "|"
					linea += str(reg["TAS_CASO_TRAFICO_ID"]) + "|"
					linea += str(reg["TAS_CASO_TRAFICO_RUTA_TAS_ENT"]) + "|"
					linea += str(reg["TAS_CASO_TRAFICO_RUTA_TAS_SAL"]) + "|"
					linea += str(reg["TAS_NRO_CDR"]) + "|"
					linea += str(reg["TAS_CASO_TRAFICO_OPERADORA"]) + "|"
					linea += str(reg["TAS_CASO_TRAFICO_RITEM"]) + "|"
					linea += str(reg["TAS_CASO_TRAFICO_DIR_CONTABLE"]) + "|"
					linea += str(reg["TAS_CASO_TRAFICO_MET_CONTABLE"]) + "|"
					linea += str(reg["TAS_CASO_TRAFICO_CLASIF"]) + "|"
					linea += str(reg["TAS_CASO_TRAFICO_CLASE_TARIFA_GRUPO"]) + "|"
					linea += str(reg["TAS_CASO_TRAFICO_CLASE_TARIFA"]) + "|"
					linea += str(reg["TAS_CASO_TRAFICO_RITEMDET"]) + "|"
					linea += str(reg["TAS_CASO_TRAFICO_COD_CONTABLE"]) + "|"
					linea += str(reg["TAS_LISTA_PRECIO"]) + "|"
					linea += str(reg["TAS_LISTA_PRECIO_DET"]) + "|"
					linea += str(reg["TAS_LISTA_PRECIO_RITEM"]) + "|"
					linea += str(reg["TAS_LISTA_PRECIO_TPRECIO"]) + "|"
					linea += str(reg["TAS_LISTA_PRECIO_QMONTO_ICX"]) + "|"
					linea += str(reg["TAS_LISTA_PRECIO_QTASA_MIN"]) + "|"
					linea += str(reg["TAS_LISTA_PRECIO_QRED_SEG_UNID"]) + "|"
					linea += str(reg["TAS_LISTA_PRECIO_QRED_MIN_UNID"]) + "|"
					linea += str(reg["TAS_LISTA_PRECIO_QRED_UNID_ADIC"]) + "|"
					linea += str(reg["TAS_LISTA_PRECIO_IRED_AJUSTE"]) + "|"
					linea += str(reg["TAS_LISTA_PRECIO_MONEDA"]) + "|"
					linea += str(reg["TAS_LISTA_PRECIO_DURATION_A_FACT"]) + "|"
					linea += str(reg["TAS_LISTA_PRECIO_MONTO"]) + "|"
					linea += str(reg["I_RETASADO"]) + "|"
					linea += str(reg["Q_RETASADO"]) + "|"
					linea += str(reg["F_INSERCION"]) + "|"
					linea += str(reg["F_ULT_ACT"])

					reporte.write(linea + "\r\n")

				reporte.close()

				self.rutaArchivo = rutaArchivo + nombreReporte
				self.contenido = listaParametros["cuerpoCorreo"].replace("[/NB_ARCHIVO]",nombreReporte)
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].decode('latin1').encode('utf8')
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/C_OPERADORA]",str(solicPrecierre["C_OPERADORA"]))
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/F_DESDE]",str(solicPrecierre["F_DESDE"]))
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/F_HASTA]",str(solicPrecierre["F_HASTA"]))
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/C_DIRECCION_CONTABLE]",str(solicPrecierre["DIRECCION_CONTABLE"]))

				self.asunto = listaParametros["tituloCorreo"]

				query = "UPDATE ICX_SOLIC_REPORTE SET X_OBS = '{0}', F_ULT_ACT = NOW() WHERE C_SOLICITUD = {1}"
				ejecutarQuery(query.format("Ruta del archivo: " + self.rutaArchivo , codSolic))
				
			except Exception as ex:
				detalles = traceback.format_exc()
				observacion = "Excepcion de tipo {0} . Argumentos:\n{1!r}\nDetalles:\n{2} "
				observacion = observacion.format(type(ex).__name__, ex.args,detalles)
				observacion = observacion.replace("'","*")
				query = "UPDATE ICX_SOLIC_REPORTE SET X_OBS = '"+ str(observacion)
				query += "',F_ULT_ACT = NOW() WHERE C_SOLICITUD = "+ str(codSolic) +" AND X_OBS IS NULL;"
				respuesta = ejecutarQuery(query)
