#coding:utf-8
from datetime import datetime,timedelta
from FuncionesComunes import *
import sys,time,locale,traceback,calendar
sys.path.append("..\\..")
from libreria.conexionBD import *

class Ejecutor(FuncionesComunes):
	def __init__(self,listaParametros,codSolic,rutaArchivo):
			#print "Metodo Constructor del Ejecutor Reportes Busqueda por CDR"
			try:
				fecha = datetime.now().strftime('%Y%m%d%H%M%S')
				nombreOriginal = listaParametros["Archivo"].split('/')
				nombreOriginal = nombreOriginal[-1].replace('.txt','')
				nombreReporte = nombreOriginal+"_ResultBusq_"+"_"+str(codSolic)+"_"+fecha+".txt"
				ResultBusq = open(rutaArchivo + nombreReporte,"w")

				fechaDesde = ""
				fDesde = ""
				fechaHasta = ""
				fHasta = ""
				durDesde = ""
				durHasta = ""
				NumeroA = ""
				NumeroB = ""
				CodAreaA = ""
				CodAreaB = ""
				CentralEntrega = ""
				TipoCargo = ""

				#*************************************************************************************************
				try:
					fechaDesde = listaParametros["Fecha Desde"]
					fechaHasta = listaParametros["Fecha Hasta"]

					# CAMBIO EM
					#fDesde = fechaDesde.split('/')
					#fDesde = datetime(int(fDesde[2]),int(fDesde[1]),int(fDesde[0]))
					fDesde = fechaDesde.replace('-','')

					# CAMBIO EM
					#fHasta = fechaHasta.split('/')
					#fHasta = datetime(int(fHasta[2]),int(fHasta[1]),int(fHasta[0]))
					fHasta = fechaHasta.replace('-','')
				except:
					pass
					#print 'No Existe listaParametros["Fecha Desde"] ni listaParametros["Fecha Hasta"]'

				#***********************************************************************************************
				try:
					#print listaParametros["Duración Desde"]
					durDesde = int(listaParametros["Duración Desde".decode('utf8').encode('latin1')])
					durHasta = int(listaParametros["Duración Hasta".decode('utf8').encode('latin1')])
				except:
					pass
					#print 'No Existe listaParametros["Duración Desde"] ni listaParametros["Duración Hasta"]'

				#****************************************************************************************************
				try:
					#print listaParametros["Número de A"]
					NumeroA = listaParametros["Número de A".decode('utf8').encode('latin1')].replace(' ','')
					NumeroA = tuple(NumeroA.split(','))
				except:
					pass
					#print 'No Existe listaParametros["Número de A"]'

				#*****************************************************************************************************
				try:
					#print listaParametros["Número de B"]
					NumeroB = listaParametros["Número de B".decode('utf8').encode('latin1')].replace(' ','')
					NumeroB = tuple(NumeroB.split(','))
				except:
					pass
					#print 'No Existe listaParametros["Número de B"]'

				#******************************************************************************************************
				try:
					#print listaParametros["Código Área de A"]
					CodAreaA = listaParametros["Código Área de A".decode('utf8').encode('latin1')].replace(' ','')
					CodAreaA = tuple(CodAreaA.split(','))
				except:
					pass
					#print 'No Existe listaParametros["Código Área de A"]'

				#*****************************************************************************************************
				try:
					#print listaParametros["Código Área de B"]
					CodAreaB = listaParametros["Código Área de B".decode('utf8').encode('latin1')].replace(' ','')
					CodAreaB = tuple(CodAreaB.split(','))
				except:
					pass
					#print 'No Existe listaParametros["Código Área de B"]'

				#*****************************************************************************************************
				try:
					#print listaParametros["Central Entrega"]
					CentralEntrega = listaParametros["Central Entrega"].replace(' ','')
					CentralEntrega = tuple(CentralEntrega.split(','))
				except:
					pass
					#print 'No Existe listaParametros["Central Entrega"]'

				#*****************************************************************************************************
				try:
					#print listaParametros["Tipo de Cargo"]
					TipoCargo = listaParametros["Tipo de Cargo"].replace(' ','')
					TipoCargo = tuple(TipoCargo.split(','))
				except:
					pass
				#*****************************************************************************************************

				#*****************************************************************************************************
				#print listaParametros["Duración Desde".decode('utf8').encode('latin1')]
				#print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
				#print fDesde
				#print fHasta
				#print durDesde
				#print durHasta
				#print NumeroA
				#print NumeroB
				#print CodAreaA
				#print CodAreaB
				#print CentralEntrega
				#print TipoCargo
				#print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"

				with open(listaParametros["Archivo"]) as f:
					for linea in f:
						lineaCDR = linea.strip()
						'''
						print '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!'
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
						print lineaCDR[105:111]
						'''

						#*************************************************************************************************
						if fDesde != "" and fHasta != "":
							fCDR = lineaCDR[42:50]

							# CAMBIO EM
							#fCDR = datetime(int(fCDR[0:4]),int(fCDR[4:6]),int(fCDR[6:8]))

							if fCDR >= fDesde and fCDR <= fHasta:
								pass
							else:
								continue
						else:
							pass
							#print 'No Existe listaParametros["Fecha Desde"] ni listaParametros["Fecha Hasta"]'

						#***********************************************************************************************
						if durDesde != "" and durHasta != "":
							#print lineaCDR[56:62]
							duracion = int(lineaCDR[56:62])
							#print duracion
							#print durDesde
							#print durHasta

							if duracion >= durDesde and duracion<= durHasta:
								pass
							else:
								continue
						else:
							pass
							#print 'No Existe listaParametros["Duración Desde"] ni listaParametros["Duración Hasta"]'

						#****************************************************************************************************
						if NumeroA != "":
							if lineaCDR[7:10].startswith(NumeroA) == False or lineaCDR[10:13].startswith(NumeroA) == False or lineaCDR[13:20].startswith(NumeroA) == False:
								continue
							else:
								pass
						else:
							pass
							#print 'No Existe listaParametros["Número de A"]'

						#*****************************************************************************************************
						if NumeroB != "":
							if lineaCDR[27:30].startswith(NumeroB) == False  or  lineaCDR[30:33].startswith(NumeroB) == False  or lineaCDR[33:40].startswith(NumeroB) == False:
								continue
							else:
								pass
						else:
							pass
							#print 'No Existe listaParametros["Número de B"]'

						#******************************************************************************************************
						if CodAreaA != "":
							if lineaCDR[10:13].startswith(CodAreaA) == False:
								continue
							else:
								pass
						else:
							pass
							#print 'No Existe listaParametros["Código Área de A"]'

						#*****************************************************************************************************
						if CodAreaB != "":
							if lineaCDR[30:33].startswith(CodAreaB) == False:
								continue
							else:
								pass
						else:
							pass
							#print 'No Existe listaParametros["Código Área de B"]'

						#*****************************************************************************************************
						if CentralEntrega != "":
							if lineaCDR[92:96].startswith(CentralEntrega) == False:
								continue
							else:
								pass
						else:
							pass
							#print 'No Existe listaParametros["Central Entrega"]'

						#*****************************************************************************************************
						if TipoCargo != "":
							if lineaCDR[105:111].startswith(TipoCargo) == False:
								continue
							else:
								pass
						else:
							pass
							#print 'No Existe listaParametros["Tipo de Cargo"]'

						# ESCRIBIR LA LINEA EN EL NUEVO ARCHIVO
						ResultBusq.write(lineaCDR + "\r\n")

				ResultBusq.close()
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
				obs +="CDRs Aceptados: "+ str(CDRaceptados) +"."
				query = "UPDATE ICX_SOLIC_REPORTE SET X_OBS = '{0}', F_ULT_ACT = NOW() WHERE C_SOLICITUD = {1}"
				ejecutarQuery(query.format("Ruta del archivo: " + self.rutaArchivo + ". " + obs , codSolic))

				self.contenido = listaParametros["cuerpoCorreo"].replace("[/NB_ARCHIVO]",nombreReporte)
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].decode('latin1').encode('utf8')
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/NB_ARCHIVO]",str(nombreOriginal))
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/F_DESDE]",str(fechaDesde))
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/F_HASTA]",str(fechaHasta))
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/D_DESDE]",str(durDesde))
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/D_HASTA]",str(durHasta))
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/NRO_A]",str(NumeroA))
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/NRO_B]",str(NumeroB))
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/COD_AREA_A]",str(CodAreaA))
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/COD_AREA_B]",str(CodAreaB))
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/CENTRAL]",str(CentralEntrega))
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/T_CARGO]",str(TipoCargo)).replace("\r\n","")
				self.asunto = listaParametros["tituloCorreo"]
			except Exception as ex:
				detalles = traceback.format_exc()
				observacion = "Excepcion de tipo {0} . Argumentos:\n{1!r}\nDetalles:\n{2} "
				observacion = observacion.format(type(ex).__name__, ex.args,detalles)
				observacion = observacion.replace("'","*")
				query = "UPDATE ICX_SOLIC_REPORTE SET X_OBS = '"+ str(observacion)
				query += "',F_ULT_ACT = NOW() WHERE C_SOLICITUD = "+ str(codSolic) +" AND X_OBS IS NULL;"
				respuesta = ejecutarQuery(query)
