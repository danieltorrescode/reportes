#coding:utf-8
from libreria.SolicitudReporte import *
import pdb,os,sys,threading,logging,traceback
from datetime import datetime

atenderSolicitudes = SolicitudReporte()
#print '############################################'
#print 'Solicitudes Automaticas'
#print atenderSolicitudes.solicitudesAutomaticas()
#print '############################################'

solicitudesPendientes = atenderSolicitudes.consultarSolicitudes()
##print solicitudesPendientes
#print '############################################'
#print "Solicitudes Pendientes en cola: " + str (len(solicitudesPendientes))
#print '############################################'

def ProcesarSolicitud(solicitud):
    #atenderSolicitudes.actualizarEstatusSolic(solicitud["C_SOLICITUD"],1)
    archivo = ""
    #print "Procesando Solicitud: " + str(solicitud["C_SOLICITUD"])
    if solicitud["T_EJECUTOR"] == 1:
        #print "Activar a un Ejecutor tipo modulo python: " + str (solicitud["T_EJECUTOR"])
        tuplaParametros = atenderSolicitudes.getParametros(solicitud["C_REPORTE"],solicitud["C_TIPO_CDR"],solicitud["C_SOLICITUD"])
        listaParametros = {}
        for parametro in tuplaParametros:
            listaParametros[parametro["NB_PARAMETRO"]] = parametro["X_VALOR"]
        listaParametros["tituloCorreo"] = solicitud["X_TITULO_CORREO"]
        listaParametros["cuerpoCorreo"] = solicitud["X_CUERPO_CORREO"]
        try:
            datosEjecutor = atenderSolicitudes.iniciarEjecutor(solicitud["X_EJECUTOR"],listaParametros,solicitud["C_SOLICITUD"])
            atenderSolicitudes.actualizarEstatusSolic(solicitud["C_SOLICITUD"],2)
            #print "Solicitud atendida correctamente"
            try:
                # ENVIO DE CORREOS
                contenido = ""
                asunto = datosEjecutor["Asunto"]

                if solicitud["I_ENV_CORREO"] == 1:
                    if solicitud["I_ANEXO_CORREO"] == 1:
                        contenido = """<p style="color:black;">"""+ asunto +"""</p>"""
                        #print contenido
                    else:
                        contenido = datosEjecutor["Contenido"]
                        datosEjecutor["Ruta"] = ""
                        #print contenido

                    destinatario = ""
                    remitente = ""
                    if solicitud["I_AUTOMATICO"] == 1:
                        destinatario = solicitud["X_CORREO_REP_AUTOM"]
                    else:
                        destinatario = solicitud["X_CORREO"]


                    mensaje = """<br/><br/><p style="color:black;">""" + contenido + """</p>"""
                    atenderSolicitudes.enviarEmail(solicitud["C_SOLICITUD"],destinatario,remitente,asunto,mensaje,datosEjecutor["Ruta"])
                    observacion = "Solicitud atendida y correo enviado correctamente. "
                    atenderSolicitudes.actualizarObs(solicitud["C_SOLICITUD"],observacion)
                    #print observacion
                else:
                    pass
                    #print "No Enviar correo"
            except Exception as ex:
                print "Error durante el envio del correo"
                detalles = traceback.format_exc()
                observacion = "Excepcion de tipo {0} . Argumentos:\n{1!r}\nDetalles:\n{2} "
                observacion = observacion.format(type(ex).__name__, ex.args,detalles)
                observacion = observacion.replace("'","*")
                print observacion
                atenderSolicitudes.actualizarObs(solicitud["C_SOLICITUD"],observacion)
        except Exception as ex:
            atenderSolicitudes.actualizarEstatusSolic(solicitud["C_SOLICITUD"],3)
            print "Error al procesar la solicitud"
            detalles = traceback.format_exc()
            observacion = "Excepcion de tipo {0} . Argumentos:\n{1!r}\nDetalles:\n{2} "
            observacion = observacion.format(type(ex).__name__, ex.args,detalles)
            observacion = observacion.replace("'","*")
            print observacion
            atenderSolicitudes.actualizarObs(solicitud["C_SOLICITUD"],observacion)

    #print '################################################################################'

for solicitud in solicitudesPendientes:
    atenderSolicitudes.actualizarEstatusSolic(solicitud["C_SOLICITUD"],1)

for solicitud in solicitudesPendientes:
    flag = solicitud["I_REQ_HILO"]
    if flag == 1:
        # INICIA UN HILO PARALELO
        hilo = threading.Thread(target=ProcesarSolicitud, args=(solicitud,))
        hilo.start()
    else:
        # EL PROCESO SE MANTIENE EL HILO PRICNCIPAL
        ProcesarSolicitud(solicitud)
