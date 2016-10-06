<!--#include file="../listas.asp"-->
<%
numRecordsByPage = EliminaInyeccionSQL( request( "numRecordsByPage" ) )
if numRecordsByPage = "" then numRecordsByPage = 50

if isnumeric(numRecordsByPage) then
	numRecordsByPage=round(numRecordsByPage,0)
else
	numRecordsByPage=50
end if

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

			for I = 0 to UBound(arrayDatos,2)
				arr=arr& arrayDatos(0,I) & ","
			next
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

	sqlpag = "select id, nombre from dn_risc_sustancias as sus WHERE id IN(" & cadenaids & ") ORDER BY nombre"
	set rstpag = objConnection2.execute(sqlpag)
	if not rstpag.eof then
		arrayDatos = rstpag.GetRows
		for contadorFilas = 0 to currentPageFinalRecordNumber-currentPageInitialRecordNumber
			tablares = tablares & "<tr>"
			enlazacon = "dn_risctox_ficha_sustancia.asp"
			'Sergio -> por si hay uno solo, lo cojo
			unico_enlace = enlazacon & "?id_sustancia=" & arrayDatos(0, contadorFilas)
			tablares = tablares & "<td class='celda_risctox'><a href='" & enlazacon & "?id_sustancia=" & arrayDatos(0,contadorFilas) & "'>" & corta(arrayDatos(1,contadorFilas),100, "puntossuspensivos") & "</a><br />" & dameSinonimos(arrayDatos(0,contadorFilas)) & dameNombreingles(arrayDatos(0,contadorFilas))& dameNombrecomercial(arrayDatos(0,contadorFilas)) & "</td>"
			tablares = tablares & "</tr>"
		next
	end if
	rstpag.close
	set rstpag = nothing

	tablares = "<table class='tabla3' width='90%' align='center' border='0' cellpadding='4' cellspacing='0'>" & tablares & "</table>"

end if 'numRecordsFound>0
	
if numRecordsFound = 1 then
	response.redirect( unico_enlace )
end if

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
%>