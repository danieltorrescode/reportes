#coding:utf-8
from datetime import datetime,timedelta
from FuncionesComunes import *
import sys,time,locale,traceback
sys.path.append("..\\..")
from libreria.conexionBD import *

class Ejecutor(FuncionesComunes):
	def __init__(self,listaParametros,codSolic,rutaArchivo):
			#print "Metodo Constructor del Ejecutor Detallado Segun Formato Int"
			try:
				query = "SELECT A.C_SOLICITUD, A.C_TIPO_CDR, A.C_OPERADORA,A.C_PERIODO,A.C_PLAN_PERIODO,A.C_MONEDA,"
				query +="A.C_DIRECCION_CONTABLE AS 'CodDirCont',B.DNIO_SIGNIFICADO AS 'C_DIRECCION_CONTABLE',C.AB_MONEDA "
				query += "FROM ICX_SOLIC_CIERRE A "
				query +="INNER JOIN ICX_DOMINIO B ON B.DNIO_VALOR = A.C_DIRECCION_CONTABLE "
				query +="INNER JOIN ICX_MONEDA C ON C.C_MONEDA= A.C_MONEDA "
				query +=" WHERE A.C_SOLICITUD = " + str(listaParametros["Cod Solicitud Cierre"])
				query +=" and B.DNIO_NOMBRE = 'DNIO_DIRECCION_CONTABLE'"

				solicitudCierre = ejecutarQuery(query)
				solicitudCierre = solicitudCierre[0]

				fecha = datetime.now().strftime('%Y%m%d%H%M%S')
				nombreReporte = "ICX_" + str(solicitudCierre["C_SOLICITUD"]) +"_"+solicitudCierre["C_PERIODO"]
				nombreReporte+= "_" + solicitudCierre["C_OPERADORA"]+ "_" + solicitudCierre["C_DIRECCION_CONTABLE"].replace(" ","_") + "_" + str(fecha) +".txt"

				reporte = open(rutaArchivo + nombreReporte,"w")
				'''
				reporte.write("REPORTE DE |TASACION FINAL |DETALLADO SEGUN| FORMATO INT" + "\r\n")
				reporte.write("Operadora|Tipo CDR|Direccion Contable|Moneda|Solicitud Cierre" + "\r\n")
 				valores = str(solicitudCierre["C_OPERADORA"])+"|"+str(solicitudCierre["C_TIPO_CDR"])+"|"
 				valores += str(solicitudCierre["C_DIRECCION_CONTABLE"])+"|"+str(solicitudCierre["AB_MONEDA"])+"|"
 				valores += str(solicitudCierre["C_SOLICITUD"])
 				reporte.write(valores + "\r\n\r\n")'''


				query = "SELECT F_INICIO_PERIODO,F_FIN_PERIODO FROM ICX_PERIODOS_DET WHERE "
				query += "C_PLAN_PERIODO = '" + str(solicitudCierre["C_PLAN_PERIODO"])
				query += "' AND C_PERIODO = '" + str(solicitudCierre["C_PERIODO"])	+ "'"
				fecha = ejecutarQuery(query)
				fecha = fecha[0]

				if listaParametros["Dirección Contable"].decode('utf8').encode('latin1') == "1":

					query = "SELECT B.NB_OPERADORA,A.USER_FIELD4,A.USER_FIELD3,A.ANO,A.BNO,A.PREP_BNO_PREFIX,C.NB_ZONA,"
					query +="A.TRANSDATETIME,D.DNIO_SIGNIFICADO AS 'T_ZONA',"
					query +="A.DURATION,A.TAS_LISTA_PRECIO_DURATION_A_FACT,E.AB_MONEDA,A.TAS_LISTA_PRECIO_MONTO,"
					query +="A.TAS_LISTA_PRECIO_QTASA_MIN, date_format(A.F_CDR,'%Y%m%d') AS 'F_CDR',"
					query +="CASE WHEN date_format(A.TRANSDATETIME,'%Y-%m-%d') between F.F_INICIO_PERIODO "
					query +="and F.F_FIN_PERIODO THEN 'N' ELSE 'Y' END AS Remanente,F.F_INICIO_PERIODO,F.F_FIN_PERIODO"
					query +=" FROM ICX_TRAFICO A "
					query += "INNER JOIN ICX_OPERADORAS B ON B.C_OPERADORA ='"+solicitudCierre["C_OPERADORA"]
					query += "' AND B.C_TIPO_CDR ='"+ solicitudCierre["C_TIPO_CDR"]
					query += "' INNER JOIN ICX_ZONAS C ON C.C_ZONA = A.PREP_BNO_ZONA AND C.C_TIPO_CDR = A.C_TIPO_CDR"
					query += " INNER JOIN ICX_DOMINIO D ON D.DNIO_NOMBRE = 'DNIO_ZONA_GMT' AND D.DNIO_VALOR = B.T_ZONA"
					query += " INNER JOIN ICX_MONEDA E ON E.C_MONEDA = A.TAS_LISTA_PRECIO_MONEDA"
					query += " INNER JOIN ICX_PERIODOS_DET F ON F.C_TIPO_CDR = '" + solicitudCierre["C_TIPO_CDR"]
					query += "' AND F.C_PLAN_PERIODO ='"+solicitudCierre["C_PLAN_PERIODO"]+"' AND F.C_PERIODO = '"+solicitudCierre["C_PERIODO"]+"'"
					query += " WHERE A.F_CDR BETWEEN '" + str(fecha["F_INICIO_PERIODO"])+ "' AND '" +str(fecha["F_FIN_PERIODO"])
					query += "' AND A.TAS_CASO_TRAFICO_OPERADORA = '" + str(solicitudCierre["C_OPERADORA"])
					query += "' AND A.TAS_CASO_TRAFICO_DIR_CONTABLE = '" + str(solicitudCierre["CodDirCont"])
					query += "' AND A.TAS_LISTA_PRECIO_MONEDA = " + str(solicitudCierre["C_MONEDA"])

					#registros = ejecutarQuery(query)
					'''
					print query
					print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
					print registros
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
						linea += str(listaParametros["Cod Solicitud Cierre"]) + "|"
						linea += str(reg["Remanente"]) + "|"
						linea += self.RemanteCliente(FechGmtCliente,reg["F_INICIO_PERIODO"],reg["F_FIN_PERIODO"])

						reporte.write(linea + "\r\n")

				elif listaParametros["Dirección Contable"].decode('utf8').encode('latin1') == "2":

					query = "SELECT B.NB_OPERADORA,A.USER_FIELD5,A.USER_FIELD3,A.ANO,A.BNO,A.PREP_BNO_PREFIX,C.NB_ZONA,"
					query +="A.TRANSDATETIME,D.DNIO_SIGNIFICADO AS 'T_ZONA',"
					query +="A.DURATION,A.TAS_LISTA_PRECIO_DURATION_A_FACT,E.AB_MONEDA,A.TAS_LISTA_PRECIO_MONTO,"
					query +="A.TAS_LISTA_PRECIO_QTASA_MIN,date_format(A.F_CDR,'%Y%m%d') AS 'F_CDR',"
					query +="CASE WHEN date_format(A.TRANSDATETIME,'%Y-%m-%d') between F.F_INICIO_PERIODO and F.F_FIN_PERIODO THEN 'N' ELSE 'Y' END AS Remanente,F.F_INICIO_PERIODO,F.F_FIN_PERIODO"
					query +=" FROM ICX_TRAFICO A "
					query += "INNER JOIN ICX_OPERADORAS B ON B.C_OPERADORA ='"+solicitudCierre["C_OPERADORA"]
					query += "' AND B.C_TIPO_CDR ='"+ solicitudCierre["C_TIPO_CDR"]
					query += "' INNER JOIN ICX_ZONAS C ON C.C_ZONA = A.PREP_BNO_ZONA AND C.C_TIPO_CDR = A.C_TIPO_CDR"
					query += " INNER JOIN ICX_DOMINIO D ON D.DNIO_NOMBRE = 'DNIO_ZONA_GMT' AND D.DNIO_VALOR = B.T_ZONA"
					query += " INNER JOIN ICX_MONEDA E ON E.C_MONEDA = A.TAS_LISTA_PRECIO_MONEDA"
					query += " INNER JOIN ICX_PERIODOS_DET F ON F.C_TIPO_CDR = '" + solicitudCierre["C_TIPO_CDR"]
					query += "' AND F.C_PLAN_PERIODO ='"+solicitudCierre["C_PLAN_PERIODO"]+"' AND F.C_PERIODO = '"+solicitudCierre["C_PERIODO"]+"'"
					query += " WHERE A.F_CDR BETWEEN '" + str(fecha["F_INICIO_PERIODO"])+ "' AND '" +str(fecha["F_FIN_PERIODO"])
					query += "' AND A.TAS_CASO_TRAFICO_OPERADORA = '" + str(solicitudCierre["C_OPERADORA"])
					query += "' AND A.TAS_CASO_TRAFICO_DIR_CONTABLE = '" + str(solicitudCierre["CodDirCont"])
					query += "' AND A.TAS_LISTA_PRECIO_MONEDA = " + str(solicitudCierre["C_MONEDA"])

					registros = ejecutarQuery(query)
					'''
					print query
					print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
					print registros
					print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
					print len(registros)

					nombreCampos = "PROVEEDOR|PREFIJO_TECNICO_PROVEEDOR|IP_PROVEEDOR|NROA|NROB|NROB_PREFIJO_GEOGRAFICO|"
					nombreCampos +="NROB_PREFIJO_GEOGRAFICO_DESC|F_INICIO|GMT_PROVEEDOR|F_INICIO_GMT_PROVEEDOR|"
					nombreCampos += "DURACION_REAL|DURACION_FACTURABLE|MONEDA|TARIFA_BASE|MONTO|SOLICITUD_CIERRE|"
					nombreCampos += "REMANENTE|REMANENTE_PROVEEDOR"
					reporte.write(nombreCampos + "\r\n")	'''

					for reg in registros:

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
						linea += str(listaParametros["Cod Solicitud Cierre"]) + "|"
						linea += str(reg["Remanente"]) + "|"
						linea += self.RemanteCliente(FechGmtCliente,reg["F_INICIO_PERIODO"],reg["F_FIN_PERIODO"])

						reporte.write(linea + "\r\n")

				else:
					print("ERROR CON LA VARIBLE listaParametros")


				reporte.close()

				self.rutaArchivo = rutaArchivo + nombreReporte
				self.contenido = listaParametros["cuerpoCorreo"].replace("[/NB_ARCHIVO]",nombreReporte)
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].decode('latin1').encode('utf8')
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/C_OPERADORA]",str(solicitudCierre["C_OPERADORA"]))
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/C_PERIODO]",str(solicitudCierre["C_PERIODO"]))
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/C_DIRECCION_CONTABLE]",str(solicitudCierre["C_DIRECCION_CONTABLE"]))
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
