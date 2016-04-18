<!--#include file="dn_conexion.asp"-->
<!--#include file="adovbs.inc"-->
<!--#include file="dn_restringida.asp"-->
<!--#include file="dn_funciones_comunes.asp"-->
<!--#include file="dn_funciones_texto.asp"-->

<%
response.ContentType="text/html; charset=iso-8859-1"
%>

<%
origen = request("origen")
numero_tipo = h(request("numero_tipo"))
numero = h(request("numero"))
nombre = h(quitarTildes(request("nombre")))

' Si hay condicion, buscamos
if ((numero <> "") or (nombre <> "")) then
	' Hay condición, la montamos
	condicion=""
	if (numero <> "") then
		condicion = "(num_"&numero_tipo&" LIKE '"&numero&"%')"
		orden_numero = "num_"&numero_tipo&", "
	else
		orden_numero = ""	
	end if

	if (nombre <> "") then
		if (condicion = "") then
			condicion = "(dn_risc_sustancias.nombre LIKE '%"&nombre&"%') or (dn_risc_sinonimos.nombre LIKE '%"&nombre&"%')"
		else
			condicion = condicion & " and ((dn_risc_sustancias.nombre LIKE '%"&nombre&"%') or (dn_risc_sinonimos.nombre LIKE '%"&nombre&"%'))"
		end if
	end if

	' Calculamos cuántos resultados hay

	sql0="SELECT COUNT(DISTINCT dn_risc_sustancias.id) AS numero FROM dn_risc_sustancias FULL OUTER JOIN dn_risc_sinonimos ON dn_risc_sustancias.id = dn_risc_sinonimos.id_sustancia WHERE "&condicion
	set objRst0=objConnection2.execute(sql0)
	numero_sustancias = objRst0("numero")
	objRst0.close()
	set objRst0=nothing

	response.write "<p><strong>"&numero_sustancias&" sustancias coincidentes</strong></p>"
	if (numero_sustancias > 10) then
		response.write "<p>Se muestran las 10 primeras. Para refinar la búsqueda, introduzca más caracteres.</p>"
	end if

	' Realizamos la consulta trayendo los datos
	sql="SELECT DISTINCT TOP 10 dn_risc_sustancias.id, dn_risc_sustancias.nombre, clasificacion_1, clasificacion_2, clasificacion_3, clasificacion_4, clasificacion_5, clasificacion_6, clasificacion_7, clasificacion_8, clasificacion_9, clasificacion_10, clasificacion_11, clasificacion_12, clasificacion_13, clasificacion_14, clasificacion_15, num_"&numero_tipo&" AS numero FROM dn_risc_sustancias FULL OUTER JOIN dn_risc_sinonimos ON dn_risc_sustancias.id = dn_risc_sinonimos.id_sustancia WHERE "&condicion&" ORDER BY "&orden_numero&"dn_risc_sustancias.nombre"
	'response.write sql&"<br/>"

	set objRst=objConnection2.execute(sql)
	if (not objRst.eof) then

		' Hay resultados
		response.write "<p>Las siguientes sustancias coinciden con el nombre y/o número introducido. Pulsa sobre la sustancia deseada para copiar sus datos.</p>"

		response.write "<ul>"
		do while (not objRst.eof)
			' Para cada sustancia tenemos su nombre y frases R.
			' Las frases R se sacan de las columnas clasificacion_1 a clasificacion_15. Vienen mezcladas con símbolos;
			' después se extraerá la información para evaluar, pues el símbolo no se emplea

			' Cogemos el nombre
			nombre = h(trim(objRst("nombre")))

			' Montamos la cadena para el numero si se indico
			if (numero <> "") then
				cadena_numero = "["&objRst("numero")&"] "
			else
				cadena_numero = ""
			end if

			' Montamos frases R
			frases_r=monta_frases_r(objRst("clasificacion_1"), objRst("clasificacion_2"), objRst("clasificacion_3"), objRst("clasificacion_4"), objRst("clasificacion_5"), objRst("clasificacion_6"), objRst("clasificacion_7"), objRst("clasificacion_8"), objRst("clasificacion_9"), objRst("clasificacion_10"), objRst("clasificacion_11"), objRst("clasificacion_12"), objRst("clasificacion_13"), objRst("clasificacion_14"), objRst("clasificacion_15"))

%>
			<li><a href="javascript:selecciona_sustancia('<%=hjs(nombre)%>', '<%=hjs(objRst("numero"))%>', '<%=frases_r%>', '<%=origen%>');"><strong><%=cadena_numero%><%=corta(nombre, 80, "puntossuspensivos")%></strong></a></li>
<%

				' Buscamos sinónimos de esta sustancia
				sql2="SELECT nombre FROM dn_risc_sinonimos WHERE id_sustancia="&objRst("id")&" ORDER BY nombre"
				set objRst2=objConnection2.execute(sql2)
				if (not objRst2.eof) then
					response.write "<ul>"
					do while (not objRst2.eof)
						response.write "<li>Sinónimo: "&corta(objRst2("nombre"), 90, "puntossuspensivos")&"</li>"
						objRst2.movenext
					loop
					response.write "</ul>"
				end if
				objRst2.close()
				set objRst2=nothing
			objRst.movenext
		loop
		response.write "</ul>"
	else
		' No hay resultados
		response.write "<p>No se encontraron sustancias con este número identificativo o nombre.</p>"
	end if

objRst.close()
set objRst = nothing

else
	' No hay condicion
		response.write "<p>Indica un número identificativo o nombre para realizar la búsqueda.</p>"	
end if
%>

<%
cerrarconexion
%>
<br/>
