#coding:utf-8
from datetime import datetime,timedelta
from FuncionesComunes import *
import sys,time,locale,traceback
sys.path.append("..\\..")
from libreria.conexionBD import *

class Ejecutor(FuncionesComunes):
	def __init__(self,listaParametros,codSolic,rutaArchivo):
			#print "Metodo Constructor del Ejecutor Detallado Segun Formato 111"
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
				nombreReporte = "ICX_"+str(solicitudCierre["C_SOLICITUD"])+"_"+ solicitudCierre["C_OPERADORA"]
				nombreReporte+=  "_" + solicitudCierre["C_DIRECCION_CONTABLE"].replace(" ","_") + "_"+ str(fecha) +"_CTRANS.txt"
				# SE CREA EL ARCHVIO _CTRANS.txt: Para registros con "TIPO CARGO" igual a CTRF ó CTRM
				reporte_CTRANS = open(rutaArchivo + nombreReporte,"w")

				# SE CREA EL ARCHVIO _SIN_CTRANS.txt: Para registros con "TIPO CARGO" distinto a CTRF y CTRM
				nombreReporte2 = "ICX_"+str(solicitudCierre["C_SOLICITUD"])+"_"+ solicitudCierre["C_OPERADORA"]
				nombreReporte2 +=  "_" + solicitudCierre["C_DIRECCION_CONTABLE"].replace(" ","_") + "_" + str(fecha) +"_SIN_CTRANS.txt"
				reporte_SIN_CTRANS = open(rutaArchivo + nombreReporte2,"w")

				'''
				camposCabecera = "REPORTE DE TASACIÓN FINAL DETALLADO SEGUN FORMATO 111" + "\r\n"
				camposCabecera += "Operadora|Tipo CDR|Direccion Contable|Moneda|Solicitud Precierre" + "\r\n"
				# SE ESCRIBE LA CABECERA DEL ARCHIVO _CTRANS.txt Y _SIN_CTRANS.txt
				reporte_CTRANS.write(camposCabecera)
				reporte_SIN_CTRANS.write(camposCabecera)

 				valores = str(solicitudCierre["C_OPERADORA"])+"|"+str(solicitudCierre["C_TIPO_CDR"])+"|"
 				valores += str(solicitudCierre["C_DIRECCION_CONTABLE"])+"|"+str(solicitudCierre["AB_MONEDA"])+"|"
 				valores += str(solicitudCierre["C_SOLICITUD"])
 				# SE ESCRIBE LOS VALORES DE LOS CAMPOS DE LA CABECERA PARA AMBOS ARCHIVOS
 				reporte_CTRANS.write(valores + "\r\n")
 				reporte_SIN_CTRANS.write(valores + "\r\n")


				nombreCampos = "TIPO DE SERVICIO" + "|"
				nombreCampos += "COD OPERADORA ABONADO A" + "|"
				nombreCampos += "COD PAIS ABONADO A" + "|"
				nombreCampos += "COD ACCESO ABONADO A" + "|"
				nombreCampos += "NUMERO TELEFONICO ABONADO A" + "|"
				nombreCampos += "RESERVA 1" + "|"
				nombreCampos += "COD OPERADORA ABONADO B" + "|"
				nombreCampos += "COD PAIS ABONADO B" + "|"
				nombreCampos += "COD ACCESO ABONADO B" + "|"
				nombreCampos += "NUMERO TELEFONICO ABONADO B" + "|"
				nombreCampos += "RESERVA 2" + "|"
				nombreCampos += "FECHA LLAMADA" + "|"
				nombreCampos += "HORA LLAMADA" + "|"
				nombreCampos += "DURACIÓN LLAMADA" + "|"
				nombreCampos += "RESERVA 3" + "|"
				nombreCampos += "COSTO" + "|"
				nombreCampos += "RESERVA 4" + "|"
				nombreCampos += "TRONCAL" + "|"
				nombreCampos += "CENTRAL ENTREGA" + "|"
				nombreCampos += "RESERVA 5" + "|"
				nombreCampos += "TIPO ACCESO" + "|"
				nombreCampos += "CODIGO ACCESO OPERADOR LD" + "|"
				nombreCampos += "RESERVA 6" + "|"
				nombreCampos += "TIPO CARGO" + "|"

				reporte_CTRANS.write(nombreCampos + "\r\n")
				reporte_SIN_CTRANS.write(nombreCampos + "\r\n")'''

				query = "SELECT F_INICIO_PERIODO,F_FIN_PERIODO FROM ICX_PERIODOS_DET WHERE "
				query += "C_PLAN_PERIODO = '" + str(solicitudCierre["C_PLAN_PERIODO"])
				query += "' AND C_PERIODO = '" + str(solicitudCierre["C_PERIODO"])	+ "'"
				fecha = ejecutarQuery(query)
				fecha = fecha[0]

				query = "SELECT B.X_COD_OPERADORA X_COD_OPERADORA_A,IFNULL(BB.X_COD_OPERADORA,'29149') X_COD_OPERADORA_B,A.ANO,A.BNO,date_format(A.TRANSDATETIME,'%Y%m%d') AS 'YMD'"
				query += ",date_format(A.TRANSDATETIME,'%H%i%s') AS 'HMS',A.TAS_LISTA_PRECIO_DURATION_A_FACT"
				query += ",ABS(A.TAS_LISTA_PRECIO_MONTO),A.USER_FIELD1,A.IA_TC,A.IA_ROUTE_IN_EXT,A.IA_ROUTE_OUT_EXT"
				query += ",A.PREP_RUTA_ENT_OPERADORA,A.USER_FIELD2,IFNULL(BB.X_COD_YZ, '00') X_COD_YZ,A.TAS_LISTA_PRECIO"
				query += ",A.TAS_LISTA_PRECIO_DET,A.PREP_ANO_OPERADORA,A.PREP_BNO_OPERADORA,D.AB_PRECIO_DET"
				query += " FROM ICX_TRAFICO A "
				query += "INNER JOIN ICX_OPERADORAS B ON B.C_OPERADORA = A.PREP_ANO_OPERADORA "
				query += "LEFT JOIN ICX_OPERADORAS BB ON BB.C_OPERADORA = A.PREP_BNO_OPERADORA "
				query += "INNER JOIN ICX_SOLIC_CIERRE C ON C.C_SOLICITUD = "+ str(solicitudCierre["C_SOLICITUD"])
				query += " INNER JOIN ICX_NOMBRE_LISTA_PRECIO D ON (A.C_TIPO_CDR = D.C_TIPO_CDR AND D.C_LISTA_PRECIO = A.TAS_LISTA_PRECIO"
				query += " AND D.C_LISTA_PRECIO_DET = A.TAS_LISTA_PRECIO_DET)"
				query += " WHERE A.F_CDR BETWEEN '" + str(fecha["F_INICIO_PERIODO"])+ "' AND '" +str(fecha["F_FIN_PERIODO"])
				query += "' AND A.TAS_CASO_TRAFICO_OPERADORA = '" + str(solicitudCierre["C_OPERADORA"])
				query += "' AND A.TAS_CASO_TRAFICO_DIR_CONTABLE = '" + str(solicitudCierre["CodDirCont"])
				query += "' AND A.TAS_LISTA_PRECIO_MONEDA = " + str(solicitudCierre["C_MONEDA"])

				#registros = ejecutarQuery(query)
				'''
				print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
				print registros
				print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
				print len(registros)'''

				#for reg in registros:
				for reg in ejecutarQuery_v2(50000,query):

					linea = "01" #+ "|"
					linea += self.codOper(reg["X_COD_OPERADORA_A"],"A") #+ "|"
					ANO = self.codPais(reg["ANO"],reg["IA_ROUTE_IN_EXT"],reg["PREP_ANO_OPERADORA"],"A")
					linea += ANO[0:3]#+ "|"
					linea += ANO[3:6]#+ "|"
					linea += ANO[6:]#+ "|"
					linea += "00" #+ "|"
					linea += self.codOper(reg["X_COD_OPERADORA_B"],"B") #+ "|"
					BNO = self.codPais(reg["BNO"],reg["IA_ROUTE_IN_EXT"],reg["PREP_BNO_OPERADORA"],"B")
					linea += BNO[0:3]#+ "|"
					linea += BNO[3:6]#+ "|"
					linea += BNO[6:]#+ "|"
					linea += "00" #+ "|"
					linea += reg["YMD"] #+ "|"
					linea += reg["HMS"] #+ "|"
					linea += self.duracionLLamada(reg["TAS_LISTA_PRECIO_DURATION_A_FACT"]) #+ "|"
					linea += "0" #+ "|"
					linea += self.costo(reg["ABS(A.TAS_LISTA_PRECIO_MONTO)"]) #+ "|"
					linea += "00000000" #+ "|"
					linea += self.troncal(reg["IA_TC"],reg["IA_ROUTE_IN_EXT"],reg["IA_ROUTE_OUT_EXT"],
											reg["PREP_RUTA_ENT_OPERADORA"],solicitudCierre["C_OPERADORA"]) #+ "|"
					linea += self.centralEntrega(reg["USER_FIELD2"]) #+ "|"
					linea += "0" #+ "|"
					tipoAcceso = self.tipoAcceso(reg["USER_FIELD1"],reg["BNO"])
					linea += tipoAcceso #+ "|"
					linea += self.codAcceso(tipoAcceso,reg["IA_TC"],reg["X_COD_YZ"]) #+ "|"
					linea += "0000" #+ "|"
					TipoCargo = self.tipoCargo(reg["TAS_LISTA_PRECIO"],reg["AB_PRECIO_DET"],reg["TAS_LISTA_PRECIO_DET"])
					linea += TipoCargo #+ "|"

					if TipoCargo.find("CTRF") != -1 or TipoCargo.find("CTRM") != -1:
						reporte_CTRANS.write(linea + "\r\n")
					else:
						reporte_SIN_CTRANS.write(linea + "\r\n")

				reporte_CTRANS.close()
				reporte_SIN_CTRANS.close()

				self.rutaArchivo = rutaArchivo + nombreReporte
				nombreReporte = nombreReporte +" y "+ nombreReporte2
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
