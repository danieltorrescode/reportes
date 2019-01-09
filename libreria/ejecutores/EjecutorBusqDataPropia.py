#coding:utf-8
from datetime import datetime,timedelta
from FuncionesComunes import *
import sys,time,locale,traceback,calendar
sys.path.append("..\\..")
from libreria.conexionBD import *

class Ejecutor(FuncionesComunes):
	def __init__(self,listaParametros,codSolic,rutaArchivo):
			#print "Metodo Constructor del Ejecutor Reportes Busqueda Datos Propios"
			try:
				query = "SELECT * FROM ICX_SOLIC_REPORTE WHERE C_SOLICITUD = " + str(codSolic) +";"
				resultConsulta = ejecutarQuery(query)
				resultConsulta = resultConsulta[0]

				fecha = datetime.now().strftime('%Y%m%d%H%M%S')
				tipoCDR = "tipoCDR"
				if resultConsulta["C_TIPO_CDR"] == 'NEHP':
					tipoCDR = "NEHP"
				elif resultConsulta["C_TIPO_CDR"] == 'NEXTONE':
					tipoCDR = "NEXTONE"

				nombreReporte = tipoCDR+"_ResultBusq_"+str(codSolic)+"_"+fecha+".txt"
				ResultBusq = open(rutaArchivo + nombreReporte,"w")

				fDesde = ''
				fHasta = ''
				durDesde = ''
				durHasta = ''
				NumeroA = ''
				NumeroB = ''
				CodAreaA = ''
				CodAreaB = ''
				CentralEntrega = ''
				TipoCargo = ''

				#*************************************************************************************************
				#print listaParametros["Fecha Desde"]
				#fDesde = listaParametros["Fecha Desde"].split('-')
				#fDesde = datetime(int(fDesde[2]),int(fDesde[1]),int(fDesde[0]))
				#fDesde = fDesde.strftime('%Y-%m-%d')

				#fHasta = listaParametros["Fecha Hasta"].split('-')
				#fHasta = datetime(int(fHasta[2]),int(fHasta[1]),int(fHasta[0]))
				#fHasta = fHasta.strftime('%Y-%m-%d')

				fDesde = listaParametros["Fecha Desde"]
				fHasta = listaParametros["Fecha Hasta"]

				#***********************************************************************************************
				try:
					#print listaParametros["Duración Desde"]
					durDesde = listaParametros["Duración Desde".decode('utf8').encode('latin1')].replace(' ','')
					durHasta = listaParametros["Duración Hasta".decode('utf8').encode('latin1')].replace(' ','')
					if durDesde == "":
						durDesde="null"

					if durHasta == "":
						durHasta="null"

					#print lineaCDR[56:62]
					#duracion = int(lineaCDR[56:62])
					#print duracion
					#print durDesde
					#print durHasta
				except Exception as inst:
					print type(inst)     # the exception instance
					print inst.args      # arguments stored in .args
					print inst
					durDesde = "null"
					durHasta = "null"
					#print 'No Existe listaParametros["Duración Desde"] ni listaParametros["Duración Hasta"]'

				#****************************************************************************************************
				try:
					#print listaParametros["Número de A"]
					NumeroA = listaParametros["Número de A".decode('utf8').encode('latin1')].replace(' ','')
					NumeroA = NumeroA.split(',')
					NumeroA = str(tuple(NumeroA)).replace("(","").replace(",)","").replace(")","")
				except:
					pass
					print 'No Existe listaParametros["Número de A"]'

				#*****************************************************************************************************
				try:
					#print listaParametros["Número de B"]
					NumeroB = listaParametros["Número de B".decode('utf8').encode('latin1')].replace(' ','')
					NumeroB = NumeroB.split(',')
					NumeroB = str(tuple(NumeroB)).replace("(","").replace(",)","").replace(")","")
				except:
					pass
					#print 'No Existe listaParametros["Número de B"]'

				#******************************************************************************************************
				try:
					#print listaParametros["Código Área de A"]
					CodAreaA = listaParametros["Código Área de A".decode('utf8').encode('latin1')].replace(' ','')
					CodAreaA = CodAreaA.split(',')
					CodAreaA = str(tuple(CodAreaA)).replace("(","").replace(",)","").replace(")","")
				except:
					pass
					#print 'No Existe listaParametros["Código Área de A"]'

				#*****************************************************************************************************
				try:
					#print listaParametros["Código Área de B"]
					CodAreaB = listaParametros["Código Área de B".decode('utf8').encode('latin1')].replace(' ','')
					CodAreaB = CodAreaB.split(',')
					CodAreaB = str(tuple(CodAreaB)).replace("(","").replace(",)","").replace(")","")
				except:
					pass
					#print 'No Existe listaParametros["Código Área de B"]'

				#*****************************************************************************************************
				try:
					#print listaParametros["Central Entrega"]
					CentralEntrega = listaParametros["Central Entrega"].replace(' ','')
					CentralEntrega = CentralEntrega.split(',')
					CentralEntrega = str(tuple(CentralEntrega)).replace("(","").replace(",)","").replace(")","")
				except:
					pass
					#print 'No Existe listaParametros["Central Entrega"]'

				#*****************************************************************************************************
				try:
					#print listaParametros["Tipo de Cargo"]
					TipoCargo = listaParametros["Tipo de Cargo"].replace(' ','')
					TipoCargo = TipoCargo.split(',')
					TipoCargo = str(tuple(TipoCargo)).replace("(","").replace(",)","").replace(")","")
				except:
					pass
					#print 'No Existe listaParametros["Tipo de Cargo"]'
				# ESCRIBIR LA LINEA EN EL NUEVO ARCHIVO

				query = "SELECT * FROM ICX_TRAFICO T, ICX_NOMBRE_LISTA_PRECIO N "
				query +="WHERE T.F_CDR BETWEEN '" + str(fDesde)+ "' AND '" +str(fHasta)+"' "
				query +="AND T.TAS_LISTA_PRECIO_DURATION_A_FACT BETWEEN IFNULL("+durDesde+", 0) "
				query +="AND IFNULL("+durHasta+", 9999999999) "

				if NumeroA != "''":
					query +="AND T.ANO IN ("+NumeroA+") "

				if NumeroB != "''":
					query +="AND T.BNO IN ("+NumeroB+") "

				if CodAreaA != "''":
					query +="AND SUBSTR(T.ANO, 1, LENGTH("+CodAreaA+")) IN ("+CodAreaA+") "

				if CodAreaB != "''":
					query +="AND SUBSTR(T.BNO, 1, LENGTH("+CodAreaB+")) IN ("+CodAreaB+") "

				if CentralEntrega != "''":
					query +="AND T.USER_FIELD2 IN ("+CentralEntrega+") "

				#print codSolic

				if resultConsulta["C_TIPO_CDR"] == 'NEHP':
					query +="AND T.TAS_LISTA_PRECIO_DET = N.C_LISTA_PRECIO_DET "
					query +="AND T.TAS_LISTA_PRECIO = N.C_LISTA_PRECIO "
					if TipoCargo != "''":
						query +="AND N.AB_PRECIO_DET IN ("+TipoCargo+");"
				elif resultConsulta["C_TIPO_CDR"] == 'NEXTONE':
					if TipoCargo != "''":
						query +="AND T.TAS_LISTA_PRECIO_DET  IN ("+TipoCargo+");"

				#print query
				#registros = ejecutarQuery(query)
				#print registros

				cabecera ="F_CDR|C_TIPO_CDR|NB_ARCHIVO|ANO|BNO|DURATION|IA_ROUTE_IN_EXT|IA_ROUTE_OUT_EXT|"
				cabecera +="IA_SERV_CLASS_EXT|IA_TC|TRANSDATETIME|USER_FIELD1|USER_FIELD2|USER_FIELD3|USER_FIELD4|"
				cabecera +="USER_FIELD5|PREP_RUTA_ENT_RUTA|PREP_RUTA_ENT_RUTA_TASACION|PREP_RUTA_ENT_T_RUTA|"
				cabecera +="PREP_RUTA_ENT_OPERADORA|PREP_RUTA_SAL_RUTA|PREP_RUTA_SAL_RUTA_TASACION|PREP_RUTA_SAL_T_RUTA|"
				cabecera +=" PREP_RUTA_SAL_OPERADORA|PREP_ANO_PREFIX|PREP_BNO_PREFIX|PREP_ANO_ZONA|PREP_BNO_ZONA|"
				cabecera +=" PREP_ANO_OPERADORA|PREP_BNO_OPERADORA|PREP_ZR_DESTINO|PREP_ZR_DESTINO_TASACION|"
				cabecera +="PREP_ANO_ORIGEN_EXCEP|PREP_ANO_ORIGEN|PREP_ANO_ORIGEN_TASACION|TAS_CASO_TRAFICO_ID|"
				cabecera +="TAS_CASO_TRAFICO_RUTA_TAS_ENT|TAS_CASO_TRAFICO_RUTA_TAS_SAL|TAS_NRO_CDR|"
				cabecera +="TAS_CASO_TRAFICO_OPERADORA|TAS_CASO_TRAFICO_RITEM|TAS_CASO_TRAFICO_DIR_CONTABLE|"
				cabecera +="TAS_CASO_TRAFICO_MET_CONTABLE|TAS_CASO_TRAFICO_CLASIF|TAS_CASO_TRAFICO_CLASE_TARIFA_GRUPO|"
				cabecera +="TAS_CASO_TRAFICO_CLASE_TARIFA|TAS_CASO_TRAFICO_RITEMDET|TAS_CASO_TRAFICO_COD_CONTABLE|"
				cabecera +="TAS_LISTA_PRECIO|TAS_LISTA_PRECIO_DET|TAS_LISTA_PRECIO_RITEM|TAS_LISTA_PRECIO_TPRECIO|"
				cabecera +="TAS_LISTA_PRECIO_QMONTO_ICX|TAS_LISTA_PRECIO_QTASA_MIN|TAS_LISTA_PRECIO_QRED_SEG_UNID|"
				cabecera +="TAS_LISTA_PRECIO_QRED_MIN_UNID|TAS_LISTA_PRECIO_QRED_UNID_ADIC|TAS_LISTA_PRECIO_IRED_AJUSTE|"
				cabecera +="TAS_LISTA_PRECIO_MONEDA|TAS_LISTA_PRECIO_DURATION_A_FACT|TAS_LISTA_PRECIO_MONTO|I_RETASADO|"
				cabecera +="Q_RETASADO|F_INSERCION|F_ULT_ACT"

				ResultBusq.write(cabecera + "\r\n")

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
					linea += '"'+str(reg["USER_FIELD1"])+'"' + "|"
					linea += '"'+str(reg["USER_FIELD2"])+'"' + "|"
					linea += '"'+str(reg["USER_FIELD3"])+'"' + "|"
					linea += '"'+str(reg["USER_FIELD4"])+'"' + "|"
					linea += '"'+str(reg["USER_FIELD5"])+'"' + "|"
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
					ResultBusq.write(linea + "\r\n")

				ResultBusq.close()

				self.rutaArchivo = rutaArchivo + nombreReporte
				self.contenido = listaParametros["cuerpoCorreo"].replace("[/NB_ARCHIVO]",nombreReporte)
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].decode('latin1').encode('utf8')
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/F_DESDE]",str(fDesde))
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/F_HASTA]",str(fHasta))
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/D_DESDE]",str(durDesde))
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/D_HASTA]",str(durHasta))
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/NRO_A]",NumeroA)
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/NRO_B]",NumeroB)
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/COD_AREA_A]",CodAreaA)
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/COD_AREA_B]",CodAreaB)
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/CENTRAL]",CentralEntrega)
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/T_CARGO]",TipoCargo).replace("\n","")
				self.asunto = listaParametros["tituloCorreo"].replace("''","")

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
