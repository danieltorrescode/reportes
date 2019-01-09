#coding:utf-8
from datetime import datetime,timedelta
from FuncionesComunes import *
import sys,time,locale,traceback,calendar
sys.path.append("..\\..")
from libreria.conexionBD import *

class Ejecutor(FuncionesComunes):
	def __init__(self,listaParametros,codSolic,rutaArchivo):
			#print "Metodo Constructor del Ejecutor  Búsqueda por CDR por Números Públicos"
			try:
				fecha = datetime.now().strftime('%Y%m%d%H%M%S')
				nombreOriginal = listaParametros["Archivo"].split('/')
				nombreOriginal = nombreOriginal[-1].replace('.txt','')
				nombreReporte = nombreOriginal+"_N1CDT_FIJO_"+str(codSolic)+"_"+fecha+".txt"
				nombreReporteFijo = nombreReporte
				N1CDT_FIJO = open(rutaArchivo + nombreReporte,"w")
				#*************************************************************
				nombreReporte2 = nombreOriginal+"_N1CDT_MOVIL_"+str(codSolic)+"_"+fecha+".txt"
				N1CDT_MOVIL = open(rutaArchivo + nombreReporte2,"w")
				# QUERY BUSQUEDA TABLA ICX_LISTA_MATCH_DET
				query = "SELECT X_VALOR_MATCH from ICX_LISTA_MATCH_DET WHERE NB_LISTA_MATCH = 'PUBLICOS' "
				query +="AND X_DESCRIPCION = 'TELEFONOS PUBLICOS NETUNO'"
				tuplaValores = ejecutarQuery(query)
				listaMatch=[]
				for valor in tuplaValores:
					listaMatch.append(valor['X_VALOR_MATCH'])

				prefijoMovil = ('412','414','416','424','426')

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
						#*************************************************************************************************
						# COD ACCESO ABONADO A
						#print lineaCDR[10:13]
						# NUMERO TELEFONICO ABONADO A
						#print lineaCDR[13:20]
						if listaMatch.count(lineaCDR[10:13]+lineaCDR[13:20]) == 1:
							if prefijoMovil.count(lineaCDR[30:33]) == 1:
								N1CDT_MOVIL.write(lineaCDR + "\r\n")
							else:
								N1CDT_FIJO.write(lineaCDR + "\r\n")
						else:
							continue
				N1CDT_FIJO.close()
				N1CDT_MOVIL.close()

				CDRleidos = 0
				CDRaceptados = 0
				CDRaceptados2 = 0
				with open(listaParametros["Archivo"]) as f:
					CDRleidos = sum(1 for _ in f)
				with open(rutaArchivo + nombreReporte) as g:
					CDRaceptados = sum(1 for _ in g)
				with open(rutaArchivo + nombreReporte2) as h:
					CDRaceptados2 = sum(1 for _ in h)

				#print CDRleidos
				#print CDRaceptados
				#print CDRaceptados2
				self.rutaArchivo = rutaArchivo + nombreReporte

				obs ="Archivo Consultado: " + nombreOriginal + '|'
				obs +="Archivo Generado: " + nombreReporte + ','
				obs += nombreReporteFijo + '|'
				obs +="CDRs leidos: "+ str(CDRleidos)+ '|'
				obs +="CDRs Aceptados: "+ str(CDRaceptados) + ","
				obs +="CDRs Aceptados2: "+ str(CDRaceptados2)

				query = "UPDATE ICX_SOLIC_REPORTE SET X_OBS = '{0}', F_ULT_ACT = NOW() WHERE C_SOLICITUD = {1}"
				ejecutarQuery(query.format("Ruta del archivo: " + self.rutaArchivo + ' . ' + obs, codSolic))

				nombreReporte = nombreReporte +" y "+nombreReporte2
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
