<!--#include file="../listas.asp"-->
<%
function doSearch(displayMode)
	dim result : set result = Server.CreateObject("Scripting.Dictionary")
	dim numRecordsByPage
	numRecordsByPage = EliminaInyeccionSQL( request( "numRecordsByPage" ) )
	if numRecordsByPage = "" then numRecordsByPage = 50
	if isnumeric(numRecordsByPage) then
		numRecordsByPage=round(numRecordsByPage,0)
	else
		numRecordsByPage=50
	end if

	Dim nombre, numero, tipobus
	nombre = lcase(EliminaInyeccionSQL(request.form("nombre")))
	numero = EliminaInyeccionSQL(request.form("numero"))
	tipobus = EliminaInyeccionSQL(request.form("tipobus"))

	select case displayMode

		case "search":

			dim searchQuery : searchQuery = obtainSearchQuery(nombre, numero, tipobus)

			Set objRst = Server.CreateObject("ADODB.Recordset")
			objRst.Open searchQuery, objConnection2, adOpenStatic, adCmdText
			numRecordsFound = objRst.recordcount

			if not objRst.eof then
				arrayDatos = objRst.getrows

				substancesFoundedIdsSrz = serializeIds(arrayDatos)
				currentPageNumber = 1
			end if

			objRst.Close
			Set objRst = Nothing

		case "pagination":

			numRecordsFound = EliminaInyeccionSQL(request("numRecordsFound"))
			currentPageNumber = EliminaInyeccionSQL(request("currentPageNumber"))
			substancesFoundedIdsSrz = request("substancesFoundedIdsSrz")

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

		arrayx = split(substancesFoundedIdsSrz, ",")

		for i = currentPageInitialRecordNumber to currentPageFinalRecordNumber
			cadenaids = cadenaids & arrayx(i) & ","
		next
		
		cadenaids = left( cadenaids, len( cadenaids ) - 1 )

		sqlpag = "select id, nombre from dn_risc_sustancias as sus WHERE id IN(" & cadenaids & ") ORDER BY nombre"
		set rstpag = objConnection2.execute(sqlpag)
		arrayDatos = rstpag.GetRows
		rstpag.close
		set rstpag = nothing
		tablares = formatHtmlTable(arrayDatos _
						, currentPageFinalRecordNumber _
						, currentPageInitialRecordNumber _
					)

	end if 'numRecordsFound>0
		
	if numRecordsFound = 1 then
		dim unico_enlace : unico_enlace = "dn_risctox_ficha_sustancia.asp?id_sustancia=" & arrayDatos(0, 0)
		response.redirect( unico_enlace )
	end if

	result.add "currentPageNumber", currentPageNumber
	result.add "numRecordsByPage", numRecordsByPage
	result.add "nombre", nombre
	result.add "numero", numero
	result.add "numRecordsFound", numRecordsFound
	result.add "substancesFoundedIdsSrz", substancesFoundedIdsSrz
	result.add "tablares", tablares
	result.add "currentPageInitialRecordNumber", currentPageInitialRecordNumber
	result.add "currentPageFinalRecordNumber", currentPageFinalRecordNumber

	set doSearch = result
end function

function formatHtmlTable(arrayDatos _
			, currentPageFinalRecordNumber _
			, currentPageInitialRecordNumber _
		)
	dim i
	for i = 0 to currentPageFinalRecordNumber-currentPageInitialRecordNumber
		tableContent = tableContent &_
		"<tr>" &_
			"<td class='celda_risctox'>" &_
				"<a href='dn_risctox_ficha_sustancia.asp?id_sustancia=" & arrayDatos(0,i) & "'>" &_
					corta(arrayDatos(1,i), 100, "puntossuspensivos") &_
				"</a><br />" &_ 
				dameSinonimos(arrayDatos(0,i)) &_
				dameNombreIngles(arrayDatos(0,i)) &_
				dameNombreComercial(arrayDatos(0,i)) &_
			"</td>" &_
		"</tr>"
	next
	
	formatHtmlTable = "<table class='tabla3' width='90%' align='center' border='0' cellpadding='4' cellspacing='0'>" & tableContent & "</table>"
end function

function obtainSearchQuery(byVal nombre, byVal numero, tipobus)
	dim condicion : condicion = ""

	if nombre <> "" or numero <> "" then
		condicion = ""
	
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
		
	end if

	sqls = "select distinct sus.id, sus.nombre "
	sqls = sqls & " from dn_risc_sustancias as sus FULL OUTER JOIN dn_risc_sinonimos as sin ON (sus.id=sin.id_sustancia) "
	sqls = sqls & " FULL OUTER JOIN dn_risc_nombres_comerciales as com ON (sus.id=com.id_sustancia) "

	sqls = sqls & get_string_tablas(0) 'magic number 0 means basic type of buscador'

	if condicion <> "" then sqls = sqls & " WHERE (" & condicion & ")"
	
	sqls = sqls & " ORDER BY sus.nombre"

	obtainSearchQuery = sqls
end function

function serializeIds(list)
	dim i, result 
	for i = 0 to UBound(list,2)
		result = result & list(0, i) & ","
	next

	serializeIds = result
end function
%>