#coding:utf-8
from datetime import datetime,timedelta
from FuncionesComunes import *
import sys,time,locale,traceback,calendar
sys.path.append("..\\..")
from libreria.conexionBD import *

class Ejecutor(FuncionesComunes):
	def __init__(self,listaParametros,codSolic,rutaArchivo):
			#print "Metodo Constructor del Ejecutor Estadistica Trafico Int"
			try:
				fecha = datetime.now().strftime('%Y%m%d%H%M%S')
				fechaHoy = datetime.now()
				mes = ""
				fch = ""
				if fechaHoy.strftime('%d') == "01":
					if fechaHoy.strftime('%m')== "01":
						fch = datetime(fechaHoy.year - 1,12,fechaHoy.day)
					else:
						fch = datetime(fechaHoy.year,fechaHoy.month - 1,fechaHoy.day)
				else:
					fch = fechaHoy

				mes = fch.strftime('%m')
				ultDiaMes = calendar.monthrange(fch.year,fch.month)
				ultDiaMes = ultDiaMes[1]
				yymm = str(fch.year) + str(mes)

				datosReporte = "REPORTE DE |ESTADISTICA DE TRAFICO |DE INTERCONEXION INTERNACIONAL|" + "\r\n"
				nombreCampos = "PROVEEDOR|DESCR PROVEEDOR|CONTRATO PROV|CLIENTE|DESCR CLIENTE|CONTRATO CLI|MES|PREF.GEO|"
				nombreCampos+= "DESTINO|LLAMADAS|SEGUNDOS|SEGUNDOS A FACTURAR|MINUTOS A FACTURAR|RED (SEG UNIDAD)|"
				nombreCampos+= "RED (MINIMA UNIDAD)|RED (UNIDAD ADICIONAL)|TARIFA|MONTO|MONEDA|ORIGEN|DESTINO" + "\r\n"

				# REPORTE DEL 01 AL UTLTMO DIA DEL MES
				nombreReporte = "ICX_COSTO_FDESDE_"+yymm+str("01")+"_FHASTA_"+yymm+str(ultDiaMes)+"_"+ str(fecha) +".txt"
				rep_COSTO_01_ult = open(rutaArchivo + nombreReporte,"w")
				rep_COSTO_01_ult.write(datosReporte)
				rep_COSTO_01_ult.write(nombreCampos)
				# REPORTE DEL 01 AL DIA 27 DEL MES
				nombreReporte = "ICX_COSTO_FDESDE_"+yymm+str("01")+"_FHASTA_"+yymm+str(27)+"_"+ str(fecha) +".txt"
				rep_COSTO_01_27 = open(rutaArchivo + nombreReporte,"w")
				rep_COSTO_01_27.write(datosReporte)
				rep_COSTO_01_27.write(nombreCampos)
				# REPORTE DEL 28 AL UTLTMO DIA DEL MES
				nombreReporte = "ICX_COSTO_FDESDE_"+yymm+str(28)+"_FHASTA_"+yymm+str(ultDiaMes)+"_"+ str(fecha) +".txt"
				rep_COSTO_28_ult = open(rutaArchivo + nombreReporte,"w")
				rep_COSTO_28_ult.write(datosReporte)
				rep_COSTO_28_ult.write(nombreCampos)
				# REPORTE DEL 01 AL UTLTMO DIA DEL MES
				nombreReporte = "ICX_FACT_FDESDE_"+yymm+str("01")+"_FHASTA_"+yymm+str(ultDiaMes)+"_"+ str(fecha) +".txt"
				rep_FACT_01_ult = open(rutaArchivo + nombreReporte,"w")
				rep_FACT_01_ult.write(datosReporte)
				rep_FACT_01_ult.write(nombreCampos)
				# REPORTE DEL 01 AL DIA 27 DEL MES
				nombreReporte = "ICX_FACT_FDESDE_"+yymm+str("01")+"_FHASTA_"+yymm+str(27)+"_"+ str(fecha) +".txt"
				rep_FACT_01_27 = open(rutaArchivo + nombreReporte,"w")
				rep_FACT_01_27.write(datosReporte)
				rep_FACT_01_27.write(nombreCampos)
				# REPORTE DEL 28 AL UTLTMO DIA DEL MES
				nombreReporte = "ICX_FACT_FDESDE_"+yymm+str(28)+"_FHASTA_"+yymm+str(ultDiaMes)+"_"+ str(fecha) +".txt"
				rep_FACT_28_ult = open(rutaArchivo + nombreReporte,"w")
				rep_FACT_28_ult.write(datosReporte)
				rep_FACT_28_ult.write(nombreCampos)


				query = "SELECT OP.C_OPERADORA,OP.NB_OPERADORA,OP.X_CAMPO_USUARIO1,OC.C_OPERADORA,OC.NB_OPERADORA,"
				query +="OC.X_CAMPO_USUARIO1,A.F_CDR,A.PREP_BNO_PREFIX,Z.NB_ZONA,COUNT(*),SUM(A.DURATION),"
				query +="SUM(A.TAS_LISTA_PRECIO_DURATION_A_FACT),SUM(A.TAS_LISTA_PRECIO_DURATION_A_FACT)/60,"
				query +="A.TAS_LISTA_PRECIO_QRED_SEG_UNID,A.TAS_LISTA_PRECIO_QRED_MIN_UNID,A.TAS_LISTA_PRECIO_IRED_AJUSTE,"
				query +="A.TAS_LISTA_PRECIO_QTASA_MIN,A.TAS_LISTA_PRECIO_MONTO,A.BNO,A.USER_FIELD5,M.AB_MONEDA"
				query +=" FROM ICX_TRAFICO A "
				query += "INNER JOIN ICX_ZONAS Z ON Z.C_TIPO_CDR = A.C_TIPO_CDR AND Z.C_ZONA = A.PREP_BNO_ZONA "
				query += "INNER JOIN ICX_MONEDA M ON M.C_MONEDA = A.TAS_LISTA_PRECIO_MONEDA "
				query += "INNER JOIN ICX_OPERADORAS OC ON A.TAS_CASO_TRAFICO_OPERADORA = OC.C_OPERADORA "
				query += "INNER JOIN ICX_OPERADORAS OP ON OP.C_TARDEST = "
				query += "IFNULL(SUBSTRING_INDEX(SUBSTRING_INDEX(A.USER_FIELD5,'|',2),'|',-1),'') "
				query += "WHERE MONTH(A.F_CDR) ="+str(mes)+" AND YEAR(A.F_CDR) ="+str(fch.year)
				query += " AND A.C_TIPO_CDR = 'NEXTONE' AND A.TAS_CASO_TRAFICO_DIR_CONTABLE = '1' "
				query += "GROUP BY OP.C_OPERADORA,OP.NB_OPERADORA,OP.X_CAMPO_USUARIO1,OC.C_OPERADORA,OC.NB_OPERADORA,"
				query += "OC.X_CAMPO_USUARIO1,A.F_CDR,A.PREP_BNO_PREFIX,Z.NB_ZONA,A.TAS_LISTA_PRECIO_QRED_SEG_UNID,"
				query += "A.TAS_LISTA_PRECIO_QRED_MIN_UNID,A.TAS_LISTA_PRECIO_IRED_AJUSTE,A.TAS_LISTA_PRECIO_QTASA_MIN,"
				query += "A.TAS_LISTA_PRECIO_MONTO,A.BNO,A.USER_FIELD5,M.AB_MONEDA ORDER BY A.F_CDR ASC"

				#registros = ejecutarQuery(query)
				'''
				print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
				print registros
				print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
				print len(registros)'''
				for reg in ejecutarQuery_v2(50000,query):

					linea = str(reg["C_OPERADORA"])+ "|"
					linea += str(reg["NB_OPERADORA"])+ "|"
					linea += str(reg["X_CAMPO_USUARIO1"])+ "|"
					linea += str(reg["C_OPERADORA"])+ "|"
					linea += str(reg["NB_OPERADORA"])+ "|"
					linea += str(reg["X_CAMPO_USUARIO1"])+ "|"
					linea += str(reg["F_CDR"].strftime('%m'))+ "|"
					linea += str(reg["PREP_BNO_PREFIX"])+ "|"
					linea += str(reg["NB_ZONA"])+ "|"
					linea += str(reg["COUNT(*)"])+ "|"
					linea += str(reg["SUM(A.DURATION)"])+ "|"
					linea += str(reg["SUM(A.TAS_LISTA_PRECIO_DURATION_A_FACT)"])+ "|"
					linea += str(reg["SUM(A.TAS_LISTA_PRECIO_DURATION_A_FACT)/60"])+ "|"
					linea += str(reg["TAS_LISTA_PRECIO_QRED_SEG_UNID"])+ "|"
					linea += str(reg["TAS_LISTA_PRECIO_QRED_MIN_UNID"])+ "|"
					linea += str(reg["TAS_LISTA_PRECIO_IRED_AJUSTE"])+ "|"
					linea += str(reg["TAS_LISTA_PRECIO_QTASA_MIN"])+ "|"
					linea += str(reg["TAS_LISTA_PRECIO_MONTO"])+ "|"
					linea += str(reg["AB_MONEDA"])+ "|"
					linea += self.origen(str(reg["USER_FIELD5"]))+ "|"
					linea += self.destino(str(reg["BNO"])) + "\r\n"

					# ESCRITURA DE LOS REPORTES ICX_COSTO
					F_CDR = reg["F_CDR"].strftime('%Y%m%d')
					if F_CDR >= yymm+str("01") and F_CDR <= yymm+str(27):
						# ESCRIBE EN EL REPORTE DEL 01 AL DIA 27 DEL MES
						rep_FACT_01_27.write(linea)
					elif F_CDR >= yymm+str(28) and F_CDR <= yymm+str(ultDiaMes):
						# ESCRIBE EN EL REPORTE DEL 28 AL UTLTMO DIA DEL MES
						rep_FACT_28_ult.write(linea)
					# ESCRIBE EN EL REPORTE DEL 01 AL UTLTMO DIA DEL MES
					rep_FACT_01_ult.write(linea)
				#############################################################################################################
				query = "SELECT OP.C_OPERADORA,OP.NB_OPERADORA,OP.X_CAMPO_USUARIO1,OC.C_OPERADORA,OC.NB_OPERADORA,"
				query +="OC.X_CAMPO_USUARIO1,A.F_CDR,A.PREP_BNO_PREFIX,Z.NB_ZONA,COUNT(*),SUM(A.DURATION),"
				query +="SUM(A.TAS_LISTA_PRECIO_DURATION_A_FACT),SUM(A.TAS_LISTA_PRECIO_DURATION_A_FACT)/60,"
				query +="A.TAS_LISTA_PRECIO_QRED_SEG_UNID,A.TAS_LISTA_PRECIO_QRED_MIN_UNID,A.TAS_LISTA_PRECIO_IRED_AJUSTE,"
				query +="A.TAS_LISTA_PRECIO_QTASA_MIN,A.TAS_LISTA_PRECIO_MONTO,A.BNO,A.USER_FIELD5,M.AB_MONEDA"
				query +=" FROM ICX_TRAFICO A "
				query += "INNER JOIN ICX_ZONAS Z ON Z.C_TIPO_CDR = A.C_TIPO_CDR AND Z.C_ZONA = A.PREP_BNO_ZONA "
				query += "INNER JOIN ICX_MONEDA M ON M.C_MONEDA = A.TAS_LISTA_PRECIO_MONEDA "
				query += "INNER JOIN ICX_OPERADORAS OC ON OC.C_TARDEST = "
				query += "IFNULL(SUBSTRING_INDEX(SUBSTRING_INDEX(A.USER_FIELD5,'|',2),'|',-1),'') "
				query += "INNER JOIN ICX_OPERADORAS OP ON A.TAS_CASO_TRAFICO_OPERADORA = OP.C_OPERADORA "
				query += "WHERE MONTH(A.F_CDR) ="+str(mes)+" AND YEAR(A.F_CDR) ="+str(fch.year)
				query += " AND A.C_TIPO_CDR = 'NEXTONE' AND A.TAS_CASO_TRAFICO_DIR_CONTABLE = '2' "
				query += "GROUP BY OP.C_OPERADORA,OP.NB_OPERADORA,OP.X_CAMPO_USUARIO1,OC.C_OPERADORA,OC.NB_OPERADORA,"
				query += "OC.X_CAMPO_USUARIO1,A.F_CDR,A.PREP_BNO_PREFIX,Z.NB_ZONA,A.TAS_LISTA_PRECIO_QRED_SEG_UNID,"
				query += "A.TAS_LISTA_PRECIO_QRED_MIN_UNID,A.TAS_LISTA_PRECIO_IRED_AJUSTE,A.TAS_LISTA_PRECIO_QTASA_MIN,"
				query += "A.TAS_LISTA_PRECIO_MONTO,A.BNO,A.USER_FIELD5,M.AB_MONEDA ORDER BY A.F_CDR ASC"

				#registros = ejecutarQuery(query)
				'''
				print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
				print registros
				print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
				print len(registros)'''
				for reg in ejecutarQuery_v2(50000,query):

					linea = str(reg["C_OPERADORA"])+ "|"
					linea += str(reg["NB_OPERADORA"])+ "|"
					linea += str(reg["X_CAMPO_USUARIO1"])+ "|"
					linea += str(reg["C_OPERADORA"])+ "|"
					linea += str(reg["NB_OPERADORA"])+ "|"
					linea += str(reg["X_CAMPO_USUARIO1"])+ "|"
					linea += str(reg["F_CDR"].strftime('%m'))+ "|"
					linea += str(reg["PREP_BNO_PREFIX"])+ "|"
					linea += str(reg["NB_ZONA"])+ "|"
					linea += str(reg["COUNT(*)"])+ "|"
					linea += str(reg["SUM(A.DURATION)"])+ "|"
					linea += str(reg["SUM(A.TAS_LISTA_PRECIO_DURATION_A_FACT)"])+ "|"
					linea += str(reg["SUM(A.TAS_LISTA_PRECIO_DURATION_A_FACT)/60"])+ "|"
					linea += str(reg["TAS_LISTA_PRECIO_QRED_SEG_UNID"])+ "|"
					linea += str(reg["TAS_LISTA_PRECIO_QRED_MIN_UNID"])+ "|"
					linea += str(reg["TAS_LISTA_PRECIO_IRED_AJUSTE"])+ "|"
					linea += str(reg["TAS_LISTA_PRECIO_QTASA_MIN"])+ "|"
					linea += str(reg["TAS_LISTA_PRECIO_MONTO"])+ "|"
					linea += str(reg["AB_MONEDA"])+ "|"
					linea += self.origen(str(reg["USER_FIELD5"]))+ "|"
					linea += self.destino(str(reg["BNO"])) + "\r\n"

					# ESCRITURA DE LOS REPORTES ICX_COSTO
					F_CDR = reg["F_CDR"].strftime('%Y%m%d')
					if F_CDR >= yymm+str("01") and F_CDR <= yymm+str(27):
						# ESCRIBE EN EL REPORTE DEL 01 AL DIA 27 DEL MES
						rep_COSTO_01_27.write(linea)
					elif F_CDR >= yymm+str(28) and F_CDR <= yymm+str(ultDiaMes):
						# ESCRIBE EN EL REPORTE DEL 28 AL UTLTMO DIA DEL MES
						rep_COSTO_28_ult.write(linea)
					# ESCRIBE EN EL REPORTE DEL 01 AL UTLTMO DIA DEL MES
					rep_COSTO_01_ult.write(linea)

				rep_COSTO_01_ult.close()
				rep_COSTO_01_27.close()
				rep_COSTO_28_ult.close()
				rep_FACT_01_ult.close()
				rep_FACT_01_27.close()
				rep_FACT_28_ult.close()

				self.rutaArchivo = rutaArchivo + nombreReporte
				self.contenido = ""#listaParametros["cuerpoCorreo"].replace("[/NB_ARCHIVO]",nombreReporte)
				self.asunto = ""
				
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
