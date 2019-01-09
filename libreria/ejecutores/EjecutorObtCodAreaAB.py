#coding:utf-8
from datetime import datetime,timedelta
from FuncionesComunes import *
import sys,time,locale,traceback,calendar
sys.path.append("..\\..")
from libreria.conexionBD import *

class Ejecutor(FuncionesComunes):
	def __init__(self,listaParametros,codSolic,rutaArchivo):
			#print "Metodo Constructor del Ejecutor Reporte para obtener código de área de A y B en soporte "
			try:
				fecha = datetime.now().strftime('%Y%m%d%H%M%S')
				nombreOriginal = listaParametros["Archivo"].split('/')
				nombreOriginal = nombreOriginal[-1].replace('.txt','')
				nombreReporte = nombreOriginal+"_ObtCodAreaAB_"+str(codSolic)+"_"+fecha+".txt"
				ObtCodAreaAB = open(rutaArchivo + nombreReporte,"w")

				query = "SELECT DNIO_VALOR, DNIO_ABREVIACION FROM ICX_DOMINIO WHERE DNIO_NOMBRE "
				query +="IN ('DNIO_CODAREA_VZLA_MOVIL', 'DNIO_CODAREA_VZLA_NGN', 'DNIO_CODAREA_VZLA_FIJO');"
				listaAbreviacion = ejecutarQuery(query)
				#ObtCodAreaAB.write('lineaCDR|COD_AREA_A|DESC_COD_AREA_A|COD_AREA_B|DESC_COD_AREA_B' + "\r\n")
				with open(listaParametros["Archivo"]) as f:
					for linea in f:
						lineaCDR = linea.strip()
						'''
						print "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
						print lineaCDR
						print lineaCDR[0:2]
						print lineaCDR[2:7]
						print lineaCDR[7:10]
						print lineaCDR[10:13]
						print lineaCDR[13:20]
						print lineaCDR[20:22]
						print lineaCDR[22:27]
						print lineaCDR[27:30]
						print lineaCDR[30:33]
						print lineaCDR[33:40]
						print lineaCDR[40:42]
						print lineaCDR[42:50]
						print lineaCDR[50:56]
						print lineaCDR[56:62]
						print lineaCDR[62:63]
						print lineaCDR[63:77]
						print lineaCDR[77:85]
						print lineaCDR[85:92]
						print lineaCDR[92:96]
						print lineaCDR[96:97]
						print lineaCDR[97:98]
						print lineaCDR[98:101]
						print lineaCDR[101:105]
						print lineaCDR[105:111]'''
						#**************************************************************************************************
						descripcionA = "NO DISPONIBLE"
						descripcionB = "NO DISPONIBLE"
						if "058"== lineaCDR[7:10] or "58"== lineaCDR[7:10]:
							for la in listaAbreviacion:
								if la["DNIO_VALOR"]	== lineaCDR[10:13]:
									descripcionA = la["DNIO_ABREVIACION"]
						else:
							pass

						if "058"==lineaCDR[27:30] or "58"==lineaCDR[27:30]:
							for la in listaAbreviacion:
								if la["DNIO_VALOR"]	== lineaCDR[30:33]:
									descripcionB = la["DNIO_ABREVIACION"]
						else:
							pass

						lineaCDR = lineaCDR +'|'+ lineaCDR[10:13]+'|'+descripcionA+'|'
						lineaCDR += lineaCDR[30:33]+'|'+descripcionB+ "\r\n"
						ObtCodAreaAB.write(lineaCDR)
				ObtCodAreaAB.close()

				CDRleidos = 0
				CDRaceptados = 0
				with open(listaParametros["Archivo"]) as f:
					CDRleidos = sum(1 for _ in f)

				with open(rutaArchivo + nombreReporte) as g:
					CDRaceptados = sum(1 for _ in g)

				self.rutaArchivo = rutaArchivo + nombreReporte
				obs ="Archivo Consultado: " + nombreOriginal + '|'
				obs +="Archivo Generado: " + nombreReporte + '|'
				obs +="CDRs leidos: "+ str(CDRleidos)+ '|'
				obs +="CDRs Aceptados: "+ str(CDRaceptados)

				query = "UPDATE ICX_SOLIC_REPORTE SET X_OBS = '{0}', F_ULT_ACT = NOW() WHERE C_SOLICITUD = {1}"
				ejecutarQuery(query.format("Ruta del archivo: " + self.rutaArchivo +". " + obs, codSolic))

				self.contenido = listaParametros["cuerpoCorreo"].replace("[/NB_ARCHIVO]",nombreReporte)
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].decode('latin1').encode('utf8')
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/ARCHIVO]",nombreOriginal)
				self.asunto = listaParametros["tituloCorreo"]

			except Exception as ex:
				detalles = traceback.format_exc()
				observacion = "Excepcion de tipo {0} . Argumentos:\n{1!r}\nDetalles:\n{2} "
				observacion = observacion.format(type(ex).__name__, ex.args,detalles)
				observacion = observacion.replace("'","*")
				query = "UPDATE ICX_SOLIC_REPORTE SET X_OBS = '"+ str(observacion)
				query += "',F_ULT_ACT = NOW() WHERE C_SOLICITUD = "+ str(codSolic) +" AND X_OBS IS NULL;"
				respuesta = ejecutarQuery(query)
