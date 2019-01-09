import MySQLdb
import sys,traceback
#PARAMENTROS DE CONEXION A LA BASE DE DATOS SETTLER DE FORMA REMOTA
DB_HOST = '172.16.47.188'
DB_USER = 'icx'
DB_PASS = 's3ttl3r16?'
DB_NAME = 'icx'

def ejecutarQuery(query):
    datos = [DB_HOST, DB_USER, DB_PASS, DB_NAME]
    try:
        conn = MySQLdb.connect(*datos) # Conectar a la base de datos
        #cursor = conn.cursor() # Crear un cursor
        cursor = conn.cursor(MySQLdb.cursors.DictCursor) # Crear un cursor Dictionary
        cursor.execute(query) # Ejecutar una consulta
        #print cursor.rowcount
        if query.upper().startswith('SELECT'):
            data = cursor.fetchall() # Traer los resultados de un select
        else:
            conn.commit()
            data = None
    except Exception:
        print Exception
        traceback.print_exc()
        data = None
        #conn.rollback()
        sys.exit('Falla en conexion en la BD')

    cursor.close() # Cerrar el cursor
    conn.close()
    return data



def ejecutarQuery_v2(arraysize,query):
    datos = [DB_HOST, DB_USER, DB_PASS, DB_NAME]
    conn = MySQLdb.connect(*datos) # Conectar a la base de datos
    #cursor = conn.cursor() # Crear un cursor
    cursor = conn.cursor(MySQLdb.cursors.SSDictCursor) # Crear un cursor Dictionary
    cursor.execute(query) # Ejecutar una consulta

    try:
        ''' A generator that simplifies the use of fetchmany '''
        while True:
            results = cursor.fetchmany(arraysize)
            if not results: break
            for result in results:
                yield result
    except Exception as inst:
        print Exception
        print type(inst)     # the exception instance
        print inst.args      # arguments stored in .args
        print inst
        traceback.print_exc()
        data = None
        #conn.rollback()
        sys.exit('Falla en conexion en la BD. Error en ejecutarQuery_v2')
