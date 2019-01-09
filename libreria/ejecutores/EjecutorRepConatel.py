#coding:utf-8
from datetime import datetime,timedelta
from FuncionesComunes import *
import sys,time,locale,traceback,calendar
sys.path.append("..\\..")
from libreria.conexionBD import *

class Ejecutor(FuncionesComunes):
	def __init__(self,listaParametros,codSolic,rutaArchivo):
			#print "Metodo Constructor del Ejecutor Reporte Conatel"
			try:
				fDesde = listaParametros["Fecha Desde"]
				fHASTA = listaParametros["Fecha Hasta"]

				fDesde = fDesde.split('/')
				fDesde = datetime(int(fDesde[2]),int(fDesde[1]),int(fDesde[0]))
				fDesde1 = fDesde.strftime('%Y%m%d')
				fDesde2 = fDesde.strftime('%Y-%m-%d')

				fHASTA = fHASTA.split('/')
				fHASTA = datetime(int(fHASTA[2]),int(fHASTA[1]),int(fHASTA[0]))
				fHASTA1 = fHASTA.strftime('%Y%m%d')
				fHASTA2 = fHASTA.strftime('%Y-%m-%d')

				fecha = datetime.now().strftime('%Y%m%d%H%M%S')
				datosReporte = "REPORTE DE CONATEL ,DE TRAFICO VENEZUELA-INTERNACIONAL ,E INTERNACIONAL-VENEZUELA" + "\r\n"
				# REPORTE TRAFINTVE
				nombreReporte = "TRAFINTVE_FDESDE_"+fDesde1+"_FHASTA_"+fHASTA1+"_"+ str(fecha) +".csv"
				TRAFINTVE = open(rutaArchivo + nombreReporte,"w")
				TRAFINTVE.write(datosReporte)
				nombreCampos = "NRO_A,NRO_B,TRONCAL_SALIDA,FECHA,HORA,DURACION_SEG"+ "\r\n"
				TRAFINTVE.write(nombreCampos)
				# REPORTE TRAFVEINT
				nombreReporte2 = "TRAFVEINT_FDESDE_"+fDesde1+"_FHASTA_"+fHASTA1+"_"+ str(fecha) +".csv"
				TRAFVEINT = open(rutaArchivo + nombreReporte2,"w")
				TRAFVEINT.write(datosReporte)
				nombreCampos = "NRO_A,NRO_B,TRONCAL_SALIDA,FECHA,HORA,DURACION_SEG"+ "\r\n"
				TRAFVEINT.write(nombreCampos)
				# REPORTE TRAFRESUMEN
				nombreReporte3 = "TRAFRESUMEN_FDESDE_"+fDesde1+"_FHASTA_"+fHASTA1+"_"+ str(fecha) +".csv"
				TRAFRESUMEN = open(rutaArchivo + nombreReporte3,"w")
				TRAFRESUMEN.write(datosReporte)
				nombreCampos = "CLIENTE,CLIENTE_EMP_CONTRATANTE,TRONCAL_CLIENTE,PROVEEDOR,PROVEEDOR_EMP_CONTRATANTE,"
				nombreCampos += "TRONCAL_PROVEEDOR,ORIGEN,DESTINO,DURACION_SEG" + "\r\n"
				TRAFRESUMEN.write(nombreCampos)

				query = "SELECT A.TAS_LISTA_PRECIO_DURATION_A_FACT,A.TRANSDATETIME,A.BNO,A.USER_FIELD5 "
				query +="FROM ICX_TRAFICO A WHERE A.C_TIPO_CDR = 'NEXTONE' AND A.TAS_CASO_TRAFICO_DIR_CONTABLE = '2' "
				query += "AND A.F_CDR BETWEEN '"+fDesde2+"' AND '"+ fHASTA2+"'"

				#registros = ejecutarQuery(query)
				'''
				print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
				print registros
				print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
				print len(registros)'''
				for reg in ejecutarQuery_v2(50000,query):
					linea = self.NRO_A(str(reg["USER_FIELD5"]))+ ","
					linea += str(reg["BNO"])[2:]+ ","
					linea += "CNH"+ ","
					linea += reg["TRANSDATETIME"].strftime('%d/%m/%Y')+ ","
					linea += reg["TRANSDATETIME"].strftime('%H:%M:%S')+ ","
					linea += str(reg["TAS_LISTA_PRECIO_DURATION_A_FACT"])+ "\r\n"

					if self.origen(str(reg["USER_FIELD5"]))=="INTERNACIONAL" and self.destino(str(reg["BNO"]))!="INTERNACIONAL":
						TRAFINTVE.write(linea)
				#######################################################################################
				query = "SELECT A.TAS_LISTA_PRECIO_DURATION_A_FACT,A.TRANSDATETIME,A.BNO,A.USER_FIELD5,OP.X_CAMPO_USUARIO1 "
				query +="FROM ICX_TRAFICO A "
				query += "INNER JOIN ICX_OPERADORAS OP ON OP.C_TARDEST = "
				query += "IFNULL(SUBSTRING_INDEX(SUBSTRING_INDEX(A.USER_FIELD5,'|',2),'|',-1),'') "
				query += "WHERE A.C_TIPO_CDR = 'NEXTONE' AND A.TAS_CASO_TRAFICO_DIR_CONTABLE = '2' "
				query += "AND A.F_CDR BETWEEN '"+fDesde2+"' AND '"+ fHASTA2+"'"

				#registros = ejecutarQuery(query)
				'''
				print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
				print registros
				print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
				print len(registros)'''
				for reg in ejecutarQuery_v2(50000,query):
					linea = self.NRO_A(str(reg["USER_FIELD5"]))+ ","
					linea += str(reg["BNO"])[2:]+ ","
					linea += "CNH"+ ","
					linea += reg["TRANSDATETIME"].strftime('%d/%m/%Y')+ ","
					linea += reg["TRANSDATETIME"].strftime('%H:%M:%S')+ ","
					linea += str(reg["TAS_LISTA_PRECIO_DURATION_A_FACT"])+ "\r\n"
					if self.origen(str(reg["USER_FIELD5"]))!="INTERNACIONAL" and self.destino(str(reg["BNO"]))=="INTERNACIONAL" and ('VITCOM_RET','AUTENTICACION_MARC','DID_NUI').count(str(reg["X_CAMPO_USUARIO1"])) == 1:
						TRAFVEINT.write(linea)
				#######################################################################################
				query = "SELECT OP.C_OPERADORA,OP.X_CAMPO_USUARIO1,A.TAS_CASO_TRAFICO_OPERADORA,OC.X_CAMPO_USUARIO1,"
				query += "OP.NB_OPERADORA AS 'nombOper1',OC.NB_OPERADORA AS 'nombOper2',"
				query += "SUM(A.TAS_LISTA_PRECIO_DURATION_A_FACT),A.BNO,A.USER_FIELD5 "
				query += "FROM ICX_TRAFICO A "
				query += "INNER JOIN ICX_ZONAS Z ON Z.C_TIPO_CDR = A.C_TIPO_CDR AND Z.C_ZONA = A.PREP_BNO_ZONA "
				query += "INNER JOIN ICX_MONEDA M ON M.C_MONEDA = A.TAS_LISTA_PRECIO_MONEDA "
				query += "INNER JOIN ICX_OPERADORAS OC ON A.TAS_CASO_TRAFICO_OPERADORA = OC.C_OPERADORA "
				query += "INNER JOIN ICX_OPERADORAS OP ON OP.C_TARDEST = "
				query += "IFNULL(SUBSTRING_INDEX(SUBSTRING_INDEX(A.USER_FIELD5,'|',2),'|',-1),'') "
				query += " AND A.C_TIPO_CDR = 'NEXTONE' AND A.TAS_CASO_TRAFICO_DIR_CONTABLE = '2' "
				query += "AND A.F_CDR BETWEEN '"+fDesde2+"' AND '"+ fHASTA2+"' "
				query += "GROUP BY OP.C_OPERADORA,OP.X_CAMPO_USUARIO1,A.TAS_CASO_TRAFICO_OPERADORA,OC.X_CAMPO_USUARIO1,"
				query += "A.BNO,A.USER_FIELD5"

				#registros = ejecutarQuery(query)
				'''
				print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
				print registros
				print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
				print len(registros)'''
				for reg in ejecutarQuery_v2(50000,query):
					linea = str(reg["nombOper1"])+ ","
					linea += str(reg["X_CAMPO_USUARIO1"])+ ","
					linea += "CNH"+ ","
					linea += str(reg["nombOper2"])+ ","
					linea += str(reg["OC.X_CAMPO_USUARIO1"])+ ","
					linea += "CNH"+ ","
					linea += self.origen(str(reg["USER_FIELD5"]))+ ","
					linea += self.destino(str(reg["BNO"])) + ","
					linea += str(reg["SUM(A.TAS_LISTA_PRECIO_DURATION_A_FACT)"])+ "\r\n"
					if self.origen(str(reg["USER_FIELD5"]))!="INTERNACIONAL" and self.destino(str(reg["BNO"]))=="INTERNACIONAL":
						TRAFRESUMEN.write(linea)
				#######################################################################################
				TRAFINTVE.close()
				TRAFVEINT.close()
				TRAFRESUMEN.close()

				self.rutaArchivo = rutaArchivo + nombreReporte
				nombreReporte = nombreReporte +" , "+ nombreReporte2 +" y "+nombreReporte3
				self.contenido = listaParametros["cuerpoCorreo"].replace("[/NB_ARCHIVO]",nombreReporte)
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].decode('latin1').encode('utf8')
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/F_DESDE]",listaParametros["Fecha Desde"])
				listaParametros["tituloCorreo"] = listaParametros["tituloCorreo"].replace("[/F_HASTA]",listaParametros["Fecha Hasta"])
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

	def NRO_A(self,nroA):
		if nroA == '|' or nroA == "None":
			return "0000000000"
		else:
			nroA = nroA.split('|')
			nroA = nroA[2]
			return nroA
