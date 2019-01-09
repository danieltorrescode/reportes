from conexionBD import *
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.encoders import encode_base64
import traceback,importlib,time,shutil,smtplib,os,sys
sys.path.append('/home/settleradm/Aplicacion/libreria')
from datosCorreo import *

class SolicitudReporte(object):
	def __init__(self):
		query =  "SELECT * FROM ICX_CONFIG;"
		tuplaBuffer = ejecutarQuery(query)
		self.rutaArchivo = tuplaBuffer[0]['X_DIR_REPORTES']
		Q_MESES_MTTOREP = tuplaBuffer[0]['Q_MESES_MTTOREP']
		for (path, ficheros, archivos) in os.walk(self.rutaArchivo):
		    #print path
		    #print ficheros
		    #print archivos
		    #print "#############################################"
		    for archivo in archivos:
		        file_path = self.rutaArchivo + '/' + archivo
		        #print archivo
		        #print os.stat(file_path)
		        fechaCreacion = datetime.fromtimestamp(os.path.getmtime(file_path ))
		        fechaHoy = datetime.now()
		        #print "#############################################"
		        #print fechaCreacion
		        #print fechaHoy.month - fechaCreacion.month
		        if fechaHoy.month - fechaCreacion.month >= Q_MESES_MTTOREP:
		            #print "NB_ARCHIVO: " + archivo + " ,debe eliminarse"
		            try:
		                os.remove(file_path)
		            except OSError:
						print "Error al borrar reporte"
		        else:
					pass
		            #print "NB_ARCHIVO: " + archivo + " ,puede quedarse en la carpeta"
		    break

	def solicitudesAutomaticas(self):
		# NOMBRE Y CODIGO DE REPORTES DE SOLICITUD AUTOMATICA
		# Reporte Diario Total -> 21
		# SIEMPRE DEBE HABER MAS DE UN ELEMENTO EN LA TUPLA POR ESO hay un elemento 0
		# PARA EL MOMENTO SOLO HAY UN REPORTE AUTOMATICO
		#reportesAutomaticos = (0,21)
		reportesAutomaticos =  ejecutarQuery("SELECT C_REPORTE,C_TIPO_CDR FROM ICX_REPORTE WHERE I_AUTOMATICO = 1;")
		query = "SELECT a.C_REPORTE,a.C_SOLICITUD FROM ICX_SOLIC_REPORTE a INNER JOIN ICX_REPORTE b ON  a.C_REPORTE=b.C_REPORTE WHERE DATE(a.F_SOLIC) = DATE(NOW()) and b.I_AUTOMATICO = 1;"
		tuplaBuffer = ejecutarQuery(query)
		solicPendiente = []
		for i in tuplaBuffer:
			solicPendiente.append(i['C_REPORTE'])

		for codSolicAuto in reportesAutomaticos:
			codSolicAuto = codSolicAuto['C_REPORTE']
			tipoCDR = codSolicAuto['C_TIPO_CDR']
			if solicPendiente.count(codSolicAuto) == 0 and codSolicAuto !=0 :
				query = "INSERT INTO ICX_SOLIC_REPORTE (C_USUARIO,F_SOLIC,ST_SOLIC,C_TIPO_CDR,C_REPORTE,X_OBS,F_ULT_ACT) "
				query += "VALUES ('NETUNO',NOW(),0,'"+str(tipoCDR)+"',"+str(codSolicAuto)+",NULL,NOW());"
				ejecutarQuery(query)
				#print query
		return "***************"

	def consultarSolicitudes(self):
		query = "SELECT a.*, b.NB_REPORTE, b.I_ACTIVO, b.T_EJECUTOR, b.X_EJECUTOR, b.I_ENV_CORREO, "
		query +=" b.I_ANEXO_CORREO,b.X_TITULO_CORREO, b.X_CUERPO_CORREO , b.I_REQ_HILO, "
		query +=" b.I_AUTOMATICO,b.X_CORREO_REP_AUTOM,c.X_CORREO "
		query +=" FROM ICX_SOLIC_REPORTE a "
		query +=" INNER JOIN ICX_REPORTE b ON a.C_REPORTE = b.C_REPORTE AND a.C_TIPO_CDR = b.C_TIPO_CDR "
		query +=" LEFT JOIN ICX_USUARIO c ON c.C_USUARIO = a.C_USUARIO "
		query +=" WHERE a.ST_SOLIC = 0 ORDER BY a.C_SOLICITUD ASC LIMIT 3;"
		#print query
		tuplaBuffer = ejecutarQuery(query)
		return tuplaBuffer

	def iniciarEjecutor(self,nombEjec,listParam,codSolic):
		s = "libreria.ejecutores." + str(nombEjec)
		exe = importlib.import_module(s)
		Ejecutor = exe.Ejecutor(listParam,codSolic,self.rutaArchivo)
		datosEjecutor = {}
		datosEjecutor['Ruta'] = Ejecutor.rutaArchivo
		datosEjecutor['Asunto'] = Ejecutor.asunto
		datosEjecutor['Contenido'] = Ejecutor.contenido
		return datosEjecutor

	def actualizarEstatusSolic(self,codSolic,estatus):
		query = "UPDATE ICX_SOLIC_REPORTE SET ST_SOLIC = "+ str(estatus)
		query +=", F_ULT_ACT = NOW() WHERE C_SOLICITUD = "+ str(codSolic) +";"
		#print query
		tuplaBuffer = ejecutarQuery(query)
		return tuplaBuffer

	def getParametros(self,codReporte,tipoCDR,codSolic):
		query = "SELECT a.*, b.X_VALOR FROM ICX_REPORTE_DET a INNER JOIN ICX_SOLIC_REPORTE_DET b ON a.R_ITEM = b.R_ITEM AND "
		query += "a.C_TIPO_CDR = b.C_TIPO_CDR AND a.C_REPORTE = b.C_REPORTE WHERE a.C_REPORTE = "+ str(codReporte)
		query += " AND a.C_TIPO_CDR = '"+ str(tipoCDR) + "' AND b.C_SOLICITUD = "+ str(codSolic) + " ORDER BY a.I_ORDEN ASC;"
		#print query
		tuplaBuffer = ejecutarQuery(query)
		return tuplaBuffer

	def actualizarObs(self,codSolic,obs):
		query = "UPDATE ICX_SOLIC_REPORTE SET X_OBS = CONCAT('"+ str(obs) +"','  ',IFNULL(X_OBS,''))  "
		query += ", F_ULT_ACT = NOW() WHERE C_SOLICITUD = "+ str(codSolic) +" ;"
		#print query
		tuplaBuffer = ejecutarQuery(query)
		return tuplaBuffer

	def enviarEmail(self,codSolic,destinatario,remitente,asunto,mensaje,rutaArchivo):
		remitente = CORREO_USUARIO

		header = MIMEMultipart()
		header['Subject']= asunto
		header['From']= remitente
		header['To']= destinatario
		mensaje = MIMEText(mensaje,'html') #Content-type:text/html
		header.attach(mensaje)

		#######################################################################
		# Datos de la cuenta de correo remitente
		#######################################################################
		username = CORREO_USUARIO
		password = CORREO_PASSWORD

		if(os.path.isfile(rutaArchivo)):
			adjunto = MIMEBase('application','octet-stream')
			adjunto.set_payload(open(rutaArchivo, "rb").read())
			encode_base64(adjunto)
			adjunto.add_header('Content-Disposition','attachment;filename="%s"' % os.path.basename(rutaArchivo))
			header.attach(adjunto)

		#######################################################################
		# Enviar el correo
		# Para enviar de otro servidor hay que cambiar el servidor y el puerto por el que este escuchando
		#######################################################################
		try:
			listDirCorreo = []
			dirCorreoX = destinatario.split(",")
			for dirC in dirCorreoX:
				listDirCorreo.append(dirC)

			server = smtplib.SMTP(CORREO_SMTP,CORREO_PUERTO)

			# Habilitar estas dos (2) lineas si la instalacion se hace en un servidor de desarrollo (con salida a internet)
			server.starttls()
			server.login(username,password)

			server.sendmail(remitente, listDirCorreo, header.as_string())
			server.quit()

			print "El mensaje pudo enviarse con exito. Correo: " + str(destinatario) + ", Titulo: " + str(asunto)
			#os.remove(rutaArchivo)
		except Exception as ex:
			detalles = traceback.format_exc()
			observacion = "Excepcion de tipo {0} . Argumentos:\n{1!r}\nDetalles:\n{2} "
			observacion = observacion.format(type(ex).__name__, ex.args,detalles)
			observacion = observacion.replace("'","*")
			self.actualizarObs(codSolic,observacion)
