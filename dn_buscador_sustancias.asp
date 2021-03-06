<!--#include file="./lib/listas.asp"-->
<%
	ordenacion = EliminaInyeccionSQL( request( "ordenacion" ) )
	sentido = EliminaInyeccionSQL( request( "sentido" ) )
	numRecordsByPage = EliminaInyeccionSQL( request( "numRecordsByPage" ) )

	'valores de busqueda por defecto
	if ordenacion = "" then ordenacion = "sus.nombre"
	if sentido = "" then sentido = ""
	if numRecordsByPage = "" then numRecordsByPage = 50

	if displayMode="" then

	else
	
		if isnumeric(numRecordsByPage) then
			numRecordsByPage=round(numRecordsByPage,0)
		else
			numRecordsByPage=50
		end if

		nombre = lcase(EliminaInyeccionSQL(request.form("nombre")))
		tipobus = EliminaInyeccionSQL(request.form("tipobus"))
		numero = EliminaInyeccionSQL(request.form("numero"))
		cas_alternativo=EliminaInyeccionSQL(request.form("cas_alternativo"))

		select case displayMode

			case "search":

				condicion=""

				if nombre <> "" or numero <> "" or cas_alternativo <> "" then
					condicion=""
				
					if nombre <> "" then	'busca en nombre, sinonimos, nombre ingles y nombre comercial
						nombre2 = h(nombre)
						nombre2 = quitartildes(nombre2)
						nombre2 = montartildes(nombre2)
						
						if tipobus = "exacto" then
							' La busqueda exacta tambien usa like, sin %, para no distinguir mayusculas

							' CONDICION NUEVA
							' Para que encuentre por nombre ingles en busqueda exacta hay que incluir los casos:
							' "nom" (exacta)
							' "nom@%" (al principio y seguido por otros)
							' "%@nom" (al final y antecedido por otros)
							' "%@nom@%" (en medio, seguido y precedido por otros)
							' No debe haber espacio junto a las @ (avisar al cliente)

							condicion = condicion &  " (sus.nombre like '" & nombre2 & "' or sin.nombre like '" & nombre2 & "' or sus.nombre_ing like '" & nombre2 & "' or sus.nombre_ing like '" & nombre2 & "@%' or sus.nombre_ing like '%@" & nombre2 & "' or sus.nombre_ing like '%@ " & nombre2 & "@%' or sus.nombre_ing like '%@" & nombre2 & "@%' or com.nombre like '" &nombre2& "')  "

						else
							condicion = condicion & " (sus.nombre like '%" & nombre2 & "%' or sin.nombre like '%" & nombre2 & "%' or sus.nombre_ing like '%" & nombre2 & "%' or com.nombre like '%" & nombre2 & "%')  "
						end if
						
					end if
				
					if numero <> "" then
						if nombre <> "" then condicion = condicion & " OR "
						condicion = condicion & " (num_ce_einecs = '" & numero & "' OR num_ce_elincs  = '" & numero & "' OR  num_rd = '" & numero & "' OR  num_cas = '" & numero & "' OR cas_alternativos like '%" & numero & "%')"
					end if
					
					if cas_alternativo <> "" then
						if (nombre <> "" or numero <> "") then condicion = condicion & " OR "
						condicion = condicion & " cas_alternativos like '%" & cas_alternativo & "%'"
					end if
					
				end if

				sqls = "select distinct sus.id, sus.nombre "
				sqls = sqls & " from dn_risc_sustancias as sus FULL OUTER JOIN dn_risc_sinonimos as sin ON (sus.id=sin.id_sustancia) "
				sqls = sqls & " FULL OUTER JOIN dn_risc_nombres_comerciales as com ON (sus.id=com.id_sustancia) "

				'según filtro, unimos a distintas tablas, nos indica de que buscador venimos
				if filtro <> "0" then sqls = sqls & get_string_tablas( filtro )

				if condicion <> "" then sqls = sqls & " WHERE (" & condicion & ")"
				
				'según filtro, agregamos distintas condiciones
				if filtro<>"0" then

					if condicion="" then
						sqls = sqls & " WHERE ("
					else
						sqls = sqls & " AND ("
					end if

					sqls = sqls & get_string_codicion( filtro ) & ")"

				end if

				sqls = sqls & " ORDER BY " & ordenacion &  " "
				response.write "<!--"&sqls&"-->"

				Set objRst = Server.CreateObject("ADODB.Recordset")
				objRst.Open sqls, objConnection2, adOpenStatic, adCmdText
				numRecordsFound = objRst.recordcount

				if not objRst.eof then
					arrayDatos = objRst.getrows

					for I = 0 to UBound(arrayDatos,2)
						arr=arr& arrayDatos(0,I) & ","
					next
					'esta sera la pagina 1
					currentPageNumber = 1
				end if

				objRst.Close
				Set objRst = Nothing

			case "pagination":

				numRecordsFound = EliminaInyeccionSQL(request("numRecordsFound"))
				currentPageNumber = EliminaInyeccionSQL(request("currentPageNumber"))
				arr = request("arr")

		end select 'cual busc

		'RESULTADOS DE BUSQUEDA (para busc 1 y busc 2)
		'seleccionamos datos a mostrar de los x registros que toquen
		if numRecordsFound>0 then

			'vemos que registros hay que mostrar
			currentPageInitialRecordNumber=(currentPageNumber*numRecordsByPage)-numRecordsByPage
			currentPageFinalRecordNumber=currentPageInitialRecordNumber+numRecordsByPage

			if currentPageFinalRecordNumber>=numRecordsFound-1 then
				currentPageFinalRecordNumber=numRecordsFound
			end if

			currentPageFinalRecordNumber=currentPageFinalRecordNumber-1

			arrayx = split(arr, ",")

			for i = currentPageInitialRecordNumber to currentPageFinalRecordNumber
				cadenaids = cadenaids & arrayx(i) & ","
			next
			
			cadenaids = left( cadenaids, len( cadenaids ) - 1 )

		%><!--#include file="dn_buscador_sustancias_lista.asp"--><%
		
		end if 'numRecordsFound>0
		
	end if 'busc

	if numRecordsFound = 1 then
		response.redirect( unico_enlace )
	end if
%>