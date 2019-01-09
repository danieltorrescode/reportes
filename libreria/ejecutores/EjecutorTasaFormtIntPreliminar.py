#coding:utf-8
from datetime import datetime,timedelta
from FuncionesComunes import *
import sys,time,locale,traceback,os
sys.path.append("..\\..")
from libreria.conexionBD import *

class Ejecutor(FuncionesComunes):
	def __init__(self,listaParametros,codSolic,rutaArchivo):
			#print "Metodo Constructor del Ejecutor Tasacion Detallado Segun Formato Int Preliminar"
			try:
				listParam = listaParametros
				listParam["Direccion_Contable"] = listParam["Dirección Contable".decode('utf8').encode('latin1')]
				query="SELECT a.*, b.DNIO_SIGNIFICADO AS 'DIRECCION_CONTABLE',c.AB_MONEDA as 'Moneda' "
				query+="FROM ICX_SOLIC_REPORTE a "
				query+="INNER JOIN ICX_DOMINIO b ON b.DNIO_VALOR='"+listParam["Direccion_Contable"]+"' "
				query+="AND b.DNIO_NOMBRE='DNIO_DIRECCION_CONTABLE' "
				query+="INNER JOIN ICX_MONEDA c ON c.C_MONEDA='"+listParam["Moneda"]+"' "
				query+="WHERE a.C_SOLICITUD ="+ str(codSolic)

				SolicFormtInt = ejecutarQuery(query)
				SolicFormtInt = SolicFormtInt[0]

				fecha = datetime.now().strftime('%Y%m%d%H%M%S')
				date_str = listaParametros["Fecha Desde"]
				format_str = '%Y-%m-%d'
				datetime_obj = datetime.strptime(date_str, format_str)
				fchDesdeFormt = datetime_obj.strftime('%Y%m%d')

				date_str = listaParametros["Fecha Hasta"]
				format_str = '%Y-%m-%d'
				datetime_obj = datetime.strptime(date_str, format_str)
				fchHastaFormt = datetime_obj.strftime('%Y%m%d')

				nombreReporte = "ICX_"+listParam["Operadora"]
				nombreReporte+= "_FDESDE_" + fchDesdeFormt + "_FHASTA_" + fchHastaFormt
				nombreReporte+=  "_" + SolicFormtInt["DIRECCION_CONTABLE"].replace(" ","_") + "_"+ str(fecha) +".txt"

				reporte = open(rutaArchivo + nombreReporte,"w")
				'''
				reporte.write("REPORTE DE |TASACION PRELIMINAR |DETALLADO SEGUN| FORMATO INT" + "\r\n")
				reporte.write("Operadora|Tipo CDR|Direccion Contable|Moneda|Solicitud Precierre" + "\r\n")
 				valores = str(listParam["Operadora"])+"|"+str(SolicFormtInt["C_TIPO_CDR"])+"|"
 				valores += str(solicPrecierre["C_DIRECCION_CONTABLE"])+"|"+str(solicPrecierre["AB_MONEDA"])+"|"
 				valores += str(solicPrecierre["C_SOLICITUD"])
 				reporte.write(valores + "\r\n\r\n")'''

				if str(listParam["Direccion_Contable"]) == "1":

					query = "SELECT B.NB_OPERADORA,A.USER_FIELD4,A.USER_FIELD3,A.ANO,A.BNO,A.PREP_BNO_PREFIX,C.NB_ZONA,"
					query +="A.TRANSDATETIME,D.DNIO_SIGNIFICADO AS 'T_ZONA',"
					query +="A.DURATION,A.TAS_LISTA_PRECIO_DURATION_A_FACT,E.AB_MONEDA,A.TAS_LISTA_PRECIO_MONTO,"
					query +="A.TAS_LISTA_PRECIO_QTASA_MIN, date_format(A.F_CDR,'%Y%m%d') AS 'F_CDR',"
					query +="CASE WHEN date_format(A.TRANSDATETIME,'%Y-%m-%d') between '"+str(listParam["Fecha Desde"])+"' "
					query +="and '"+str(listParam["Fecha Hasta"])+"' THEN 'N' ELSE 'Y' END AS Remanente"
					query +=" FROM ICX_TRAFICO A "
					query += "INNER JOIN ICX_OPERADORAS B ON B.C_OPERADORA ='"+listParam["Operadora"]
					query += "' AND B.C_TIPO_CDR ='"+ SolicFormtInt["C_TIPO_CDR"]
					query += "' INNER JOIN ICX_ZONAS C ON C.C_ZONA = A.PREP_BNO_ZONA AND C.C_TIPO_CDR = A.C_TIPO_CDR"
					query += " INNER JOIN ICX_DOMINIO D ON D.DNIO_NOMBRE = 'DNIO_ZONA_GMT' AND D.DNIO_VALOR = B.T_ZONA"
					query += " INNER JOIN ICX_MONEDA E ON E.C_MONEDA = A.TAS_LISTA_PRECIO_MONEDA"
					query += " WHERE A.F_CDR BETWEEN '"+str(listParam["Fecha Desde"])+"' AND '"+str(listParam["Fecha Hasta"])
					query += "' AND A.TAS_CASO_TRAFICO_OPERADORA = '" + str(listParam["Operadora"])
					query += "' AND A.TAS_CASO_TRAFICO_DIR_CONTABLE = '" + str(listParam["Direccion_Contable"])
					query += "' AND A.TAS_LISTA_PRECIO_MONEDA = " + str(listParam["Moneda"])

					#registros = ejecutarQuery(query)

					'''print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
					print query
					print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
					print len(registros)
					nombreCampos = "CLIENTE|PREFIJO_TECNICO_CLIENTE|IP_CLIENTE|NROA|NROB|NROB_PREFIJO_GEOGRAFICO|"
					nombreCampos += "NROB_PREFIJO_GEOGRAFICO_DESC|F_INICIO|GMT_CLIENTE|F_INICIO_GMT_CLIENTE|"
					nombreCampos += "DURACION_REAL|DURACION_FACTURABLE|MONEDA|TARIFA_BASE|MONTO|SOLICITUD_CIERRE|"
					nombreCampos += "REMANENTE|REMANENTE_CLIENTE"

					reporte.write(nombreCampos + "\r\n")'''

					for reg in ejecutarQuery_v2(50000,query):
						linea = str(reg["NB_OPERADORA"]) + "|"
						linea += self.userFild(str(reg["USER_FIELD4"])) + "|"
						linea += self.HexaToIp(str(reg["USER_FIELD3"])) + "|"
						linea += str(reg["ANO"]) + "|"
						linea += str(reg["BNO"]) + "|"
						linea += str(reg["PREP_BNO_PREFIX"]) + "|"
						linea += str(reg["NB_ZONA"]) + "|"
						linea += str(reg["TRANSDATETIME"]) + "|"
						linea += str(reg["T_ZONA"]) + "|"
						FechGmtCliente = self.FechaGmtCliente(reg["TRANSDATETIME"],reg["T_ZONA"])
						linea += str(FechGmtCliente) + "|"
						linea += str(reg["DURATION"]) + "|"
						linea += str(reg["TAS_LISTA_PRECIO_DURATION_A_FACT"]) + "|"
						linea += str(reg["AB_MONEDA"]) + "|"
						linea += str(reg["TAS_LISTA_PRECIO_QTASA_MIN"]) + "|"
						linea += str(reg["TAS_LISTA_PRECIO_MONTO"]) + "|"
						linea += str("N/A") + "|"
						linea += str(reg["Remanente"]) + "|"
						linea += self.RemanteCliente(FechGmtCliente,listParam["Fecha Desde"],listParam["Fecha Hasta"])

						reporte.write(linea + "\r\n")

				elif str(listParam["Direccion_Contable"]) == "2":

					query = "SELECT B.NB_OPERADORA,A.USER_FIELD5,A.USER_FIELD3,A.ANO,A.BNO,A.PREP_BNO_PREFIX,C.NB_ZONA,"
					query +="A.TRANSDATETIME,D.DNIO_SIGNIFICADO AS 'T_ZONA',"
					query +="A.DURATION,A.TAS_LISTA_PRECIO_DURATION_A_FACT,E.AB_MONEDA,A.TAS_LISTA_PRECIO_MONTO,"
					query +="A.TAS_LISTA_PRECIO_QTASA_MIN,date_format(A.F_CDR,'%Y%m%d') AS 'F_CDR',"
					query +="CASE WHEN date_format(A.TRANSDATETIME,'%Y-%m-%d') between '"+str(listParam["Fecha Desde"])+"' "
					query +="and '"+str(listParam["Fecha Hasta"])+"' THEN 'N' ELSE 'Y' END AS Remanente"
					query +=" FROM ICX_TRAFICO A "
					query += "INNER JOIN ICX_OPERADORAS B ON B.C_OPERADORA ='"+listParam["Operadora"]
					query += "' AND B.C_TIPO_CDR ='"+ SolicFormtInt["C_TIPO_CDR"]
					query += "' INNER JOIN ICX_ZONAS C ON C.C_ZONA = A.PREP_BNO_ZONA AND C.C_TIPO_CDR = A.C_TIPO_CDR"
					query += " INNER JOIN ICX_DOMINIO D ON D.DNIO_NOMBRE = 'DNIO_ZONA_GMT' AND D.DNIO_VALOR = B.T_ZONA"
					query += " INNER JOIN ICX_MONEDA E ON E.C_MONEDA = A.TAS_LISTA_PRECIO_MONEDA"
					query += " WHERE A.F_CDR BETWEEN '"+str(listParam["Fecha Desde"])+"' AND '"+str(listParam["Fecha Hasta"])
					query += "' AND A.TAS_CASO_TRAFICO_OPERADORA = '" + str(listParam["Operadora"])
					query += "' AND A.TAS_CASO_TRAFICO_DIR_CONTABLE = '" + str(listParam["Direccion_Contable"])
					query += "' AND A.TAS_LISTA_PRECIO_MONEDA = " + str(listParam["Moneda"])

					#registros = ejecutarQuery(query)

					'''print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
					print query
					print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
					print len(registros)

					nombreCampos = "PROVEEDOR|PREFIJO_TECNICO_PROVEEDOR|IP_PROVEEDOR|NROA|NROB|NROB_PREFIJO_GEOGRAFICO|"
					nombreCampos +="NROB_PREFIJO_GEOGRAFICO_DESC|F_INICIO|GMT_PROVEEDOR|F_INICIO_GMT_PROVEEDOR|"
					nombreCampos += "DURACION_REAL|DURACION_FACTURABLE|MONEDA|TARIFA_BASE|MONTO|SOLICITUD_PRECIERRE|"
					nombreCampos += "REMANENTE|REMANENTE_PROVEEDOR"

					reporte.write(nombreCampos + "\r\n") '''

					for reg in ejecutarQuery_v2(50000,query):

						linea = str(reg["NB_OPERADORA"]) + "|"
						linea += self.userFild(str(reg["USER_FIELD5"])) + "|"
						linea += self.HexaToIp(str(reg["USER_FIELD3"])) + "|"
						linea += str(reg["ANO"]) + "|"
						linea += str(reg["BNO"]) + "|"
						linea += str(reg["PREP_BNO_PREFIX"]) + "|"
						linea += str(reg["NB_ZONA"]) + "|"
						linea += str(reg["TRANSDATETIME"]) + "|"
						linea += str(reg["T_ZONA"]) + "|"
						FechGmtCliente = self.FechaGmtCliente(reg["TRANSDATETIME"],reg["T_ZONA"])
						linea += str(FechGmtCliente) + "|"
						linea += str(reg["DURATION"]) + "|"
						linea += str(reg["TAS_LISTA_PRECIO_DURATION_A_FACT"]) + "|"
						linea += str(reg["AB_MONEDA"]) + "|"
						linea += str(reg["TAS_LISTA_PRECIO_QTASA_MIN"]) + "|"
						linea += str(reg["TAS_LISTA_PRECIO_MONTO"] * -1) + "|"
						linea += str("N/A") + "|"
						linea += str(reg["Remanente"]) + "|"
						linea += self.RemanteCliente(FechGmtCliente,listParam["Fecha Desde"],listParam["Fecha Hasta"])

						reporte.write(linea + "\r\n")

				else:
					print("ERROR CON LA VARIBLE listaParametros")

				reporte.close()

				self.rutaArchivo = rutaArchivo + nombreReporte
				self.contenido = listParam["cuerpoCorreo"].replace("[/NB_ARCHIVO]",nombreReporte)
				listParam["tituloCorreo"] = listParam["tituloCorreo"].decode('latin1').encode('utf8')
				listParam["tituloCorreo"] = listParam["tituloCorreo"].replace("[/C_OPERADORA]",str(listParam["Operadora"]))
				listParam["tituloCorreo"] = listParam["tituloCorreo"].replace("[/F_DESDE]",str(listParam["Fecha Desde"]))
				listParam["tituloCorreo"] = listParam["tituloCorreo"].replace("[/F_HASTA]",str(listParam["Fecha Hasta"]))
				listParam["tituloCorreo"] = listParam["tituloCorreo"].replace("[/C_DIRECCION_CONTABLE]",str(SolicFormtInt["DIRECCION_CONTABLE"]))

				self.asunto = listParam["tituloCorreo"]
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
