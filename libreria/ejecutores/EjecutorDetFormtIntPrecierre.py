#coding:utf-8
from datetime import datetime,timedelta
from FuncionesComunes import *
import sys,time,locale,traceback,os
sys.path.append("..\\..")
from libreria.conexionBD import *

class Ejecutor(FuncionesComunes):
	def __init__(self,listaParametros,codSolic,rutaArchivo):
			#print "Metodo Constructor del Ejecutor Detallado Segun Formato Int Precierre"
			try:
				query = "SELECT A.C_SOLICITUD, A.C_TIPO_CDR, A.C_OPERADORA,A.F_DESDE,A.F_HASTA,A.C_MONEDA,"
				query +="A.C_DIRECCION_CONTABLE AS 'CodDirCont',B.DNIO_SIGNIFICADO AS 'C_DIRECCION_CONTABLE',C.AB_MONEDA "
				query += "FROM ICX_SOLIC_PRECIERRE A "
				query +="INNER JOIN ICX_DOMINIO B ON B.DNIO_VALOR = A.C_DIRECCION_CONTABLE "
				query +="INNER JOIN ICX_MONEDA C ON C.C_MONEDA= A.C_MONEDA "
				query +=" WHERE A.C_SOLICITUD = " + str(listaParametros["Cod Solicitud Precierre"])
				query +=" and B.DNIO_NOMBRE = 'DNIO_DIRECCION_CONTABLE'"

				solicPrecierre = ejecutarQuery(query)
				solicPrecierre = solicPrecierre[0]

				fecha = datetime.now().strftime('%Y%m%d%H%M%S')
				nombreReporte = "ICX_"+str(solicPrecierre["C_SOLICITUD"])+"_"+solicPrecierre["C_OPERADORA"]
				nombreReporte+= "_FDESDE_" + str(solicPrecierre["F_DESDE"].strftime('%Y%m%d')) + "_FHASTA_"
				nombreReporte+= str(solicPrecierre["F_HASTA"].strftime('%Y%m%d'))
				nombreReporte+=  "_" + solicPrecierre["C_DIRECCION_CONTABLE"].replace(" ","_") + "_"+ str(fecha) +".txt"

				reporte = open(rutaArchivo + nombreReporte,"w")
				'''
				reporte.write("REPORTE DE |TASACION PRECIERRE |DETALLADO SEGUN| FORMATO INT" + "\r\n")
				reporte.write("Operadora|Tipo CDR|Direccion Contable|Moneda|Solicitud Precierre" + "\r\n")
 				valores = str(solicPrecierre["C_OPERADORA"])+"|"+str(solicPrecierre["C_TIPO_CDR"])+"|"
 				valores += str(solicPrecierre["C_DIRECCION_CONTABLE"])+"|"+str(solicPrecierre["AB_MONEDA"])+"|"
 				valores += str(solicPrecierre["C_SOLICITUD"])
 				reporte.write(valores + "\r\n\r\n")'''

				if str(solicPrecierre["CodDirCont"]) == "1":

					query = "SELECT B.NB_OPERADORA,A.USER_FIELD4,A.USER_FIELD3,A.ANO,A.BNO,A.PREP_BNO_PREFIX,C.NB_ZONA,"
					query +="A.TRANSDATETIME,D.DNIO_SIGNIFICADO AS 'T_ZONA',"
					query +="A.DURATION,A.TAS_LISTA_PRECIO_DURATION_A_FACT,E.AB_MONEDA,A.TAS_LISTA_PRECIO_MONTO,"
					query +="A.TAS_LISTA_PRECIO_QTASA_MIN, date_format(A.F_CDR,'%Y%m%d') AS 'F_CDR',"
					query +="CASE WHEN date_format(A.TRANSDATETIME,'%Y-%m-%d') between '"+str(solicPrecierre["F_DESDE"])+"' "
					query +="and '"+str(solicPrecierre["F_HASTA"])+"' THEN 'N' ELSE 'Y' END AS Remanente"
					query +=" FROM ICX_TRAFICO A "
					query += "INNER JOIN ICX_OPERADORAS B ON B.C_OPERADORA ='"+solicPrecierre["C_OPERADORA"]
					query += "' AND B.C_TIPO_CDR ='"+ solicPrecierre["C_TIPO_CDR"]
					query += "' INNER JOIN ICX_ZONAS C ON C.C_ZONA = A.PREP_BNO_ZONA AND C.C_TIPO_CDR = A.C_TIPO_CDR"
					query += " INNER JOIN ICX_DOMINIO D ON D.DNIO_NOMBRE = 'DNIO_ZONA_GMT' AND D.DNIO_VALOR = B.T_ZONA"
					query += " INNER JOIN ICX_MONEDA E ON E.C_MONEDA = A.TAS_LISTA_PRECIO_MONEDA"
					query += " WHERE A.F_CDR BETWEEN '"+str(solicPrecierre["F_DESDE"])+"' AND '"+str(solicPrecierre["F_HASTA"])
					query += "' AND A.TAS_CASO_TRAFICO_OPERADORA = '" + str(solicPrecierre["C_OPERADORA"])
					query += "' AND A.TAS_CASO_TRAFICO_DIR_CONTABLE = '" + str(solicPrecierre["CodDirCont"])
					query += "' AND A.TAS_LISTA_PRECIO_MONEDA = " + str(solicPrecierre["C_MONEDA"])

					#registros = ejecutarQuery(query)

					'''print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
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
						linea += str(listaParametros["Cod Solicitud Precierre"]) + "|"
						linea += str(reg["Remanente"]) + "|"
						linea += self.RemanteCliente(FechGmtCliente,solicPrecierre["F_DESDE"],solicPrecierre["F_HASTA"])

						reporte.write(linea + "\r\n")

				elif str(solicPrecierre["CodDirCont"]) == "2":

					query = "SELECT B.NB_OPERADORA,A.USER_FIELD5,A.USER_FIELD3,A.ANO,A.BNO,A.PREP_BNO_PREFIX,C.NB_ZONA,"
					query +="A.TRANSDATETIME,D.DNIO_SIGNIFICADO AS 'T_ZONA',"
					query +="A.DURATION,A.TAS_LISTA_PRECIO_DURATION_A_FACT,E.AB_MONEDA,A.TAS_LISTA_PRECIO_MONTO,"
					query +="A.TAS_LISTA_PRECIO_QTASA_MIN,date_format(A.F_CDR,'%Y%m%d') AS 'F_CDR',"
					query +="CASE WHEN date_format(A.TRANSDATETIME,'%Y-%m-%d') between '"+str(solicPrecierre["F_DESDE"])+"' "
					query +="and '"+str(solicPrecierre["F_HASTA"])+"' THEN 'N' ELSE 'Y' END AS Remanente"
					query +=" FROM ICX_TRAFICO A "
					query += "INNER JOIN ICX_OPERADORAS B ON B.C_OPERADORA ='"+solicPrecierre["C_OPERADORA"]
					query += "' AND B.C_TIPO_CDR ='"+ solicPrecierre["C_TIPO_CDR"]
					query += "' INNER JOIN ICX_ZONAS C ON C.C_ZONA = A.PREP_BNO_ZONA AND C.C_TIPO_CDR = A.C_TIPO_CDR"
					query += " INNER JOIN ICX_DOMINIO D ON D.DNIO_NOMBRE = 'DNIO_ZONA_GMT' AND D.DNIO_VALOR = B.T_ZONA"
					query += " INNER JOIN ICX_MONEDA E ON E.C_MONEDA = A.TAS_LISTA_PRECIO_MONEDA"
					query += " WHERE A.F_CDR BETWEEN '"+str(solicPrecierre["F_DESDE"])+"' AND '"+str(solicPrecierre["F_HASTA"])
					query += "' AND A.TAS_CASO_TRAFICO_OPERADORA = '" + str(solicPrecierre["C_OPERADORA"])
					query += "' AND A.TAS_CASO_TRAFICO_DIR_CONTABLE = '" + str(solicPrecierre["CodDirCont"])
					query += "' AND A.TAS_LISTA_PRECIO_MONEDA = " + str(solicPrecierre["C_MONEDA"])

					#registros = ejecutarQuery(query)

					'''print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
					print registros
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
						linea += str(listaParametros["Cod Solicitud Precierre"]) + "|"
						linea += str(reg["Remanente"]) + "|"
						linea += self.RemanteCliente(FechGmtCliente,solicPrecierre["F_DESDE"],solicPrecierre["F_HASTA"])

						reporte.write(linea + "\r\n")

				else:
					print("ERROR CON LA VARIBLE listaParametros")

				reporte.close()

				self.rutaArchivo = rutaArchivo + nombreReporte
				self.contenido = listaParametros["cuerpoCorreo"].replace("[/NB_ARCHIVO]",nombreReporte)
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].decode('latin1').encode('utf8')
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/C_OPERADORA]",str(solicPrecierre["C_OPERADORA"]))
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/F_DESDE]",str(solicPrecierre["F_DESDE"]))
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/F_HASTA]",str(solicPrecierre["F_HASTA"]))
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/C_DIRECCION_CONTABLE]",str(solicPrecierre["C_DIRECCION_CONTABLE"]))

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
