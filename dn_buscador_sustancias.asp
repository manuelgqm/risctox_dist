<!--#include file="./lib/listas.asp"-->
<%
	ordenacion = EliminaInyeccionSQL( request( "ordenacion" ) )
	sentido = EliminaInyeccionSQL( request( "sentido" ) )
	nregs = EliminaInyeccionSQL( request( "nregs" ) )

	'valores de busqueda por defecto
	if ordenacion = "" then ordenacion = "sus.nombre"
	if sentido = "" then sentido = ""
	if nregs = "" then nregs = 50

	if busc="" then

	else
	
		if isnumeric(nregs) then
			nregs=round(nregs,0)
		else
			nregs=50
		end if

		nombre = lcase(EliminaInyeccionSQL(request.form("nombre")))
		tipobus = EliminaInyeccionSQL(request.form("tipobus"))
		numero = EliminaInyeccionSQL(request.form("numero"))
		cas_alternativo=EliminaInyeccionSQL(request.form("cas_alternativo"))

		select case busc

			case 1: 'han dado a buscar

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
				hr = objRst.recordcount

				if not objRst.eof then
					arrayDatos = objRst.getrows

					for I = 0 to UBound(arrayDatos,2)
						arr=arr& arrayDatos(0,I) & ","
					next
					'esta sera la pagina 1
					pag = 1
				end if

				objRst.Close
				Set objRst = Nothing

			case 2: 'paginando

				hr = EliminaInyeccionSQL(request("hr"))
				pag = EliminaInyeccionSQL(request("pag"))
				arr = request("arr")

		end select 'cual busc


		'RESULTADOS DE BUSQUEDA (para busc 1 y busc 2)
		'seleccionamos datos a mostrar de los x registros que toquen
		if hr>0 then

			'vemos que registros hay que mostrar
			registroini=(pag*nregs)-nregs
			'response.write "<p>registroini=" &registroini& "</p>"

			registrofin=registroini+nregs
			'response.write "<p>registrofin=" &registrofin& "</p>"

			if registrofin>=hr-1 then
				registrofin=hr
				'response.write "<p>registrofin era mayor, ahora=" &registrofin& "</p>"
			end if

			registrofin=registrofin-1
			'response.write "<p>registrofin corregido=" &registrofin& "</p>"

			arrayx = split(arr, ",")

			for i = registroini to registrofin
				cadenaids = cadenaids & arrayx(i) & ","
			next
			
			'quitamos la ultima coma
			cadenaids = left( cadenaids, len( cadenaids ) - 1 )

		%><!--#include file="dn_buscador_sustancias_lista.asp"--><%
		
		end if 'hr>0
		
	end if 'busc

	if hr = 1 then
		response.redirect( unico_enlace )
	end if
%>
