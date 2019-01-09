#coding:utf-8 
class FuncionesComunes(object):
	def __init__(self):		
		pass
	#*****************************************************************************************************************
	# FUNCIONES UTILZADAS EN LOS EJECUTORES  'EjecutorMinxLin', 'EjecutorRepConatel' Y 'EjecutorEstadisticaTraficoInt'

	def NRO_A(self,nroA):
		if nroA == '|' or nroA == "None":
			return "NO DISPONIBLE"
		else:
			nroA = nroA.split('|')
			nroA = nroA[2]
			return nroA

	def origen(self,NRO_A):
		if NRO_A == "|":
			return "pipe"
		NRO_A = NRO_A.split('|')
		NRO_A = NRO_A[2]

		COD_FIJOS_VZLA = "'212','234','235','237','238','239','240','241','242','243','244','245','246','247',"
		COD_FIJOS_VZLA += "'248','249','251','252','253','254','255','256','257','258','259','261','262','263',"
		COD_FIJOS_VZLA += "'264','265','266','267','268','269','271','272','273','274','275','276','277','278',"
		COD_FIJOS_VZLA += "'279','281','282','283','284','285','286','287','288','289','291','292','293','294',"
		COD_FIJOS_VZLA += "'295','296','412','414','416','424','426','500','501','800'"

		COD_MOVIL_VZLA = "'412','414','416','424','426'"

		COD_NGN_VZLA = "'500','501','800'"		
	
		if len(NRO_A) == 12 and NRO_A[0:2] == "58" or len(NRO_A) == 10:
			if COD_FIJOS_VZLA.find(NRO_A[0:3]) != -1:
				return "VENEZUELA FIJO"
			elif COD_MOVIL_VZLA.find(NRO_A[0:3]) != -1:
				return "VENEZUELA MOVIL"
			elif COD_NGN_VZLA.find(NRO_A[0:3]) != -1:
				return "VENEZUELA NGN"		
			else:
				return "INTERNACIONAL"
		else:
			return "INTERNACIONAL"

	def destino(self,BNO):
		if BNO[0:3] == "582" :
			return "VENEZUELA FIJO"
		elif BNO[0:3] == "584" :
			return "VENEZUELA MOVIL"
		elif BNO[0:3] == "585" or BNO[0:3] == "588" :
			return "VENEZUELA NGN"
		elif BNO[0:3] == "581" :
			return "VENEZUELA 1XY"
		elif BNO[0:9] == "VENEZUELA" :
			return "VENEZUELA"
		else:
			return "INTERNACIONAL"

	#*****************************************************************************************************************
	# FUNCIONES UTILZADAS EN LOS EJECUTORES  'EjecutorDetFormt111' Y 'EjecutorDetFormt111Preliminar'

	def codOper(self,X_COD_OPERADORA,letra):
		X_COD_OPERADORA = str(X_COD_OPERADORA)
		if (X_COD_OPERADORA == 'None' or X_COD_OPERADORA == '29148' or X_COD_OPERADORA == '29149'): 
			if letra == "A":
				X_COD_OPERADORA = '29148'
			else:
				X_COD_OPERADORA = '29149'
		return X_COD_OPERADORA

	def codPais(self,XNO,IA_ROUTE_IN_EXT,PREP_OPERADORA,letra):

		codigoAreaValidos = "212,234,235,237,238,239,240,241,242,243,244,245,246,"
		codigoAreaValidos +="247,248,249,251,252,253,254,255,256,257,258,259,261,"
		codigoAreaValidos +="262,263,264,265,266,267,268,269,271,272,273,274,275,"
		codigoAreaValidos +="276,277,278,279,281,282,283,284,285,286,287,288,289,"
		codigoAreaValidos +="291,292,293,294,295,296,412,414,416,424,426,500,501,800"
		 
		prefijoInvalido = "110,112,114,116,135"

		XNO = str(XNO)
		if XNO == 'None':
			XNO = "0"

		if letra == "A":
			if XNO[0:3] == "199" and XNO != "1996388661":
				#XNO = XNO.lstrip('199')		# La funcion lstrip elimina 1991 cuando se le dice que elimine 199 de 19912128034888
				XNO = XNO[3:]

			if str(IA_ROUTE_IN_EXT)=='7859' and XNO[0:3]!="199" and  len(XNO)==10 and codigoAreaValidos.find(XNO[0:3]) != -1:
				XNO = "058" + XNO

			if prefijoInvalido.find(XNO[0:3]) != -1:
				#XNO = XNO.lstrip(XNO[0:3])		# La funcion lstrip elimina 1991 cuando se le dice que elimine 199 de 19912128034888
				XNO = XNO[3:]

		if len(XNO) >= 10:
			if str(PREP_OPERADORA) == 'UNKNOWN' or str(PREP_OPERADORA) == 'None':
				XNO = '' + XNO
			else:
				if codigoAreaValidos.find(XNO[0:3]) != -1:
					XNO = '058' + XNO
			
			while len(XNO) < 13:
				XNO=  '0' + XNO
			
			XNO = XNO[0:13]
		
		else:
			while len(XNO) < 13:
				XNO=  '0' + XNO

		return XNO

	def duracionLLamada(self,DURATION_A_FACT):
		DURATION_A_FACT = str(DURATION_A_FACT)
		while len(DURATION_A_FACT) < 6:
			DURATION_A_FACT= '0' + DURATION_A_FACT
		return DURATION_A_FACT


	def costo(self,TAS_LISTA_PRECIO_MONTO):
		TAS_LISTA_PRECIO_MONTO = str(TAS_LISTA_PRECIO_MONTO)
		TAS_LISTA_PRECIO_MONTO = TAS_LISTA_PRECIO_MONTO.split('.')

		parteEntera = TAS_LISTA_PRECIO_MONTO[0]
		parteDecimal = TAS_LISTA_PRECIO_MONTO[1]

		while len(parteEntera) < 7:
			parteEntera = '0' + parteEntera

		while len(parteDecimal) < 7:
			parteDecimal =  parteDecimal + '0'

		return parteEntera + parteDecimal


	def troncal(self,IA_TC,IA_ROUTE_IN_EXT,IA_ROUTE_OUT_EXT,PREP_RUTA_ENT_OPERADORA,C_OPERADORA):
		troncal = ""
		IA_TC = str(IA_TC)
		IA_ROUTE_IN_EXT = str(IA_ROUTE_IN_EXT)		
		IA_ROUTE_OUT_EXT = str(IA_ROUTE_OUT_EXT)
		PREP_RUTA_ENT_OPERADORA= str(PREP_RUTA_ENT_OPERADORA)
		C_OPERADORA = str(C_OPERADORA)

		if IA_TC == "10":
			if IA_ROUTE_IN_EXT == "None":
				troncal = '       '
			else:
				troncal = IA_ROUTE_IN_EXT
		elif IA_TC == "20":
			if IA_ROUTE_OUT_EXT == "None":
				troncal = '       '
			else:
				troncal = IA_ROUTE_OUT_EXT
		elif IA_TC == "30":
			if PREP_RUTA_ENT_OPERADORA == C_OPERADORA:
				troncal = IA_ROUTE_IN_EXT
			else:
				troncal = IA_ROUTE_OUT_EXT

		while len(troncal) < 7 :
			troncal =  troncal + ' ' 
		return troncal

	def centralEntrega(self,USER_FIELD2):
		USER_FIELD2 = str(USER_FIELD2)
		if USER_FIELD2 == 'None':
			USER_FIELD2 = '    '

		if len(USER_FIELD2) < 4 :
			while len(USER_FIELD2) < 4:
				USER_FIELD2 = '\t' + USER_FIELD2

		return USER_FIELD2

	def tipoAcceso(self,USER_FIELD1,BNO):
		USER_FIELD1 = str(USER_FIELD1)
		BNO = str(BNO)
		if USER_FIELD1 == 'P':
			USER_FIELD1 = '1'
		elif BNO[0:2] == '80' and len(BNO) == 10 or USER_FIELD1=='M':
			USER_FIELD1 = '0'
		else:
			USER_FIELD1 = ' '
		return USER_FIELD1

	def codAcceso(self,tipoAcceso,IA_TC,X_COD_YZ):
		IA_TC = str(IA_TC)
		X_COD_YZ = str(X_COD_YZ)
		codAcceso = ""
		if (tipoAcceso =='0' or tipoAcceso =='1') and (IA_TC == '10' or IA_TC == '30'):
			codAcceso = "199"
		elif IA_TC == "20":

			if X_COD_YZ == 'None':
				codAcceso = '1  '
			else:
				while len(X_COD_YZ) < 2:
					X_COD_YZ = '0' + X_COD_YZ
				codAcceso = '1'+ X_COD_YZ	
		else:
			codAcceso = '   '

		return codAcceso

	def tipoCargo(self,tasListaPrecio,abPrecioDet,tasListaPrecioDet):
		tipo_cargo_extP = ""
		if str(tasListaPrecio) == "881":
			tipo_cargo_extP = str(abPrecioDet)
		else:
			tipo_cargo_extP = str(tasListaPrecioDet)

		if tipo_cargo_extP != 'None':
			if len(tipo_cargo_extP) > 6:
				tipo_cargo_extP = tipo_cargo_extP[0:6]

			tipo_cargo_extP = tipo_cargo_extP.replace(' ','')

			while len(tipo_cargo_extP)<6:
				tipo_cargo_extP = tipo_cargo_extP + "0"
		else:
			tipo_cargo_extP = '000000'

		return tipo_cargo_extP

	#*****************************************************************************************************************
	# FUNCIONES UTILZADAS EN LOS EJECUTORES  'EjecutorDetFormtInt' Y 'EjecutorDetFormtIntPreliminar'

	def HexaToIp(self,hexa):
		hexa="00000000"
		ip = str(int(hexa[0:2],16)) +"."+ str(int(hexa[2:4],16))+"."
		ip += str(int(hexa[4:6],16))+"."+ str(int(hexa[6:8],16))
		return ip

	def userFild(self,userFild):
		userFild = str(userFild).split("|")
		return  userFild[0] 

	def FechaGmtCliente(self,fch,tZona):
		if tZona.find('-') != -1:
			tZona = tZona.replace(':',':-')
		tZona = tZona.split(':')
		i= fch + timedelta(hours=int(tZona[0]), minutes=int(tZona[1]))
		return i


	def RemanteCliente(self,FechaGmtCliente,F_ini,F_fin):
		if FechaGmtCliente.date() >= F_ini and FechaGmtCliente.date() <= F_fin:
			return 'N'
		else:
			return 'Y'
