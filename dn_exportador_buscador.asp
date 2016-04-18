<!--#include file="dn_conexion.asp"-->
<!--#include file="adovbs.inc"-->
<!--#include file="dn_funciones_comunes.asp"-->
<!--#include file="dn_funciones_texto.asp"-->

<%
response.ContentType="text/html; charset=ISO-8859-15"
%>

<%
nombre = h(quitarTildes(EliminaInyeccionSQL(request("nombre"))))
numero_tipo = h(EliminaInyeccionSQL(request("numero_tipo")))
numero = h(EliminaInyeccionSQL(request("numero")))

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

%>
	<p>
		<% if (numero_sustancias > 0) then %>
			<strong><%= numero_sustancias %> sustancias coincidentes.</strong>
		<% end if %>
		<%
		if (numero_sustancias > 25) then
			response.write " Se muestran las 25 primeras. Para refinar la b&uacute;squeda, introduzca m&aacute;s caracteres."
		end if
		%>
	</p>
<%

	' Realizamos la consulta trayendo los datos
	sql="SELECT DISTINCT TOP 25 dn_risc_sustancias.id AS id_sustancia, dn_risc_sustancias.nombre, clasificacion_1, clasificacion_2, clasificacion_3, clasificacion_4, clasificacion_5, clasificacion_6, clasificacion_7, clasificacion_8, clasificacion_9, clasificacion_10, clasificacion_11, clasificacion_12, clasificacion_13, clasificacion_14, clasificacion_15, conc_1, conc_2, conc_3, conc_4, conc_5, conc_5, conc_6, conc_7, conc_8, conc_9, conc_10, conc_11, conc_12, conc_13, conc_14, conc_15, eti_conc_1, eti_conc_2, eti_conc_3, eti_conc_4, eti_conc_5, eti_conc_6, eti_conc_7, eti_conc_8, eti_conc_9, eti_conc_10, eti_conc_11, eti_conc_12, eti_conc_13, eti_conc_14, eti_conc_15, num_"&numero_tipo&" AS numero, notas_rd_363, grupo_iarc FROM dn_risc_sustancias FULL OUTER JOIN dn_risc_sustancias_iarc ON dn_risc_sustancias.id = dn_risc_sustancias_iarc.id_sustancia FULL OUTER JOIN dn_risc_sinonimos ON dn_risc_sustancias.id = dn_risc_sinonimos.id_sustancia WHERE "&condicion&" ORDER BY "&orden_numero&"dn_risc_sustancias.nombre"
	'response.write sql&"<br/>"

	set objRst=objConnection2.execute(sql)
	if (not objRst.eof) then

		' Hay resultados
		response.write "<ul>"
		do while (not objRst.eof)
			' Cogemos el nombre
			nombre = h(trim(objRst("nombre")))
			grupos = dameGrupos(objRst("id_sustancia"))
			disolvente = es_disolvente(objRst("id_sustancia"))

			' Las siguientes llamadas devuelven true o false, queremos numero 1 o 0
			' ya que lo que modificamos es el selectedIndex
			cov = true_por_uno(esta_en_lista ("cov", objRst("id_sustancia")))
			cancer_rd = true_por_uno(esta_en_lista ("cancer_rd", objRst("id_sustancia")))
			cancer_iarc = true_por_uno(esta_en_lista ("cancer_iarc", objRst("id_sustancia")))		
			tpr = true_por_uno(esta_en_lista ("tpr", objRst("id_sustancia")))
			de = true_por_uno(esta_en_lista ("de", objRst("id_sustancia")))
			neurotoxica = true_por_uno(esta_en_lista ("neurotoxico", objRst("id_sustancia")))
			sensibilizante = true_por_uno(esta_en_lista ("sensibilizante", objRst("id_sustancia")))
			tpb = true_por_uno(esta_en_lista ("tpb", objRst("id_sustancia")))

			' Para saber nivel cancerigeno y mutageno RD, necesitamos las clasificaciones
			clasificacion_1 = objRst("clasificacion_1")
			clasificacion_2 = objRst("clasificacion_2")
			clasificacion_3 = objRst("clasificacion_3")
			clasificacion_4 = objRst("clasificacion_4")
			clasificacion_5 = objRst("clasificacion_5")
			clasificacion_6 = objRst("clasificacion_6")
			clasificacion_7 = objRst("clasificacion_7")
			clasificacion_8 = objRst("clasificacion_8")
			clasificacion_9 = objRst("clasificacion_9")
			clasificacion_10 = objRst("clasificacion_10")
			clasificacion_11 = objRst("clasificacion_11")
			clasificacion_12 = objRst("clasificacion_12")
			clasificacion_13 = objRst("clasificacion_13")
			clasificacion_14 = objRst("clasificacion_14")
			clasificacion_15 = objRst("clasificacion_15")

			nivel_cancerigeno_rd = dame_nivel_cancerigeno_rd()
			nivel_mutageno_rd = dame_nivel_mutageno_rd()

			grupo_iarc = objRst("grupo_iarc")

			conc_1 = replace(objRst("conc_1"), "&#37;", "<")
			eti_conc_1 = objRst("eti_conc_1")
			conc_2 = replace(objRst("conc_2"), "&#37;", "<")
			eti_conc_2 = objRst("eti_conc_2")
			conc_3 = replace(objRst("conc_3"), "&#37;", "<")
			eti_conc_3 = objRst("eti_conc_3")
			conc_4 = replace(objRst("conc_4"), "&#37;", "<")
			eti_conc_4 = objRst("eti_conc_4")
			conc_5 = replace(objRst("conc_5"), "&#37;", "<")
			eti_conc_5 = objRst("eti_conc_5")
			conc_6 = replace(objRst("conc_6"), "&#37;", "<")
			eti_conc_6 = objRst("eti_conc_6")
			conc_7 = replace(objRst("conc_7"), "&#37;", "<")
			eti_conc_7 = objRst("eti_conc_7")
			conc_8 = replace(objRst("conc_8"), "&#37;", "<")
			eti_conc_9 = objRst("eti_conc_9")
			conc_10 = replace(objRst("conc_10"), "&#37;", "<")
			eti_conc_10 = objRst("eti_conc_10")
			conc_11 = replace(objRst("conc_11"), "&#37;", "<")
			eti_conc_11 = objRst("eti_conc_11")
			conc_12 = replace(objRst("conc_12"), "&#37;", "<")
			eti_conc_12 = objRst("eti_conc_12")
			conc_13 = replace(objRst("conc_13"), "&#37;", "<")
			eti_conc_13 = objRst("eti_conc_13")
			conc_14 = replace(objRst("conc_14"), "&#37;", "<")
			eti_conc_14 = objRst("eti_conc_14")
			conc_15 = replace(objRst("conc_15"), "&#37;", "<")
			eti_conc_15 = objRst("eti_conc_15")

			num_risctox = objRst("id_sustancia")

			' Montamos la cadena para el numero si se indico
			if (numero <> "") then
				cadena_numero = "["&objRst("numero")&"] "
			else
				cadena_numero = ""
			end if

%>
			<li><a href="javascript:selecciona_sustancia('<%=hjs(nombre)%>', '<%=hjs(objRst("numero"))%>', '<%= grupos %>', '<%= disolvente %>', '<%= cov %>', '<%= cancer_rd %>', '<%= cancer_iarc %>', '<%= tpr %>', '<%= de %>', '<%= neurotoxica %>', '<%= sensibilizante %>', '<%= tpb %>', '<%= nivel_cancerigeno_rd %>', '<%= nivel_mutageno_rd %>', '<%= objRst("notas_rd_363") %>', '<%= grupo_iarc %>', '<%= conc_1 %>', '<%= eti_conc_1 %>', '<%= conc_2 %>', '<%= eti_conc_2 %>', '<%= conc_3 %>', '<%= eti_conc_3 %>', '<%= conc_4 %>', '<%= eti_conc_4 %>', '<%= conc_5 %>', '<%= eti_conc_5 %>', '<%= conc_6 %>', '<%= eti_conc_6 %>', '<%= conc_7 %>', '<%= eti_conc_7 %>', '<%= conc_8 %>', '<%= eti_conc_8 %>', '<%= conc_9 %>', '<%= eti_conc_9 %>', '<%= conc_10 %>', '<%= eti_conc_10 %>', '<%= conc_11 %>', '<%= eti_conc_11 %>', '<%= conc_12 %>', '<%= eti_conc_12 %>', '<%= conc_13 %>', '<%= eti_conc_13 %>', '<%= conc_14 %>', '<%= eti_conc_14 %>', '<%= conc_15 %>', '<%= eti_conc_15 %>', '<%= num_risctox %>');"><strong><%=cadena_numero%><%=corta(nombre, 80, "puntossuspensivos")%></strong></a></li>
<%

				' Buscamos sinónimos de esta sustancia
				sql2="SELECT nombre FROM dn_risc_sinonimos WHERE id_sustancia="&objRst("id_sustancia")&" ORDER BY nombre"
				set objRst2=objConnection2.execute(sql2)
				if (not objRst2.eof) then
					response.write "<ul>"
					do while (not objRst2.eof)
						response.write "<li>Sin&oacute;nimo: "&h(corta(objRst2("nombre"), 90, "puntossuspensivos"))&"</li>"
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
		'	response.write "<p>No se encontraron sustancias con este n&uacute;mero identificativo o nombre.</p>"
%>
		<script language="JavaScript">
			new Effect.Fade('rixresults');
		</script>
<%
	end if

objRst.close()
set objRst = nothing

else
	' No hay condicion
		response.write "<p>Indica un n&uacute;mero identificativo o nombre para realizar la b&uacute;squeda.</p>"	
end if



cerrarconexion
%>


<%
' ##########################################################################
function dameGrupos(byval id_sustancia)
	' Devuelve lista de grupos para la sustancia indicada

	lista = ""

	sql="SELECT nombre FROM dn_risc_sustancias_por_grupos INNER JOIN dn_risc_grupos ON dn_risc_sustancias_por_grupos.id_grupo = dn_risc_grupos.id WHERE id_sustancia="&id_sustancia&" ORDER BY nombre"
	set objRstGrupos=objConnection2.execute(sql)
	if (not objRstGrupos.eof) then
		do while (not objRstGrupos.eof)
			if (lista = "") then
				lista = objRstGrupos("nombre")
			else
				lista = lista&", "&objRstGrupos("nombre")
			end if

			objRstGrupos.movenext
		loop
	end if
	objRstGrupos.close()
	set objRstGrupos=nothing

	dameGrupos = h(lista)
end function

' #################################################################

function true_por_uno(byval booleano)
	if (booleano = true) then
		true_por_uno = 1
	else
		true_por_uno = 0
	end if
end function

' #############################################################################################

function dame_nivel_cancerigeno_rd()
	' Juntamos todas las clasificaciones
	clasificacion_rd = clasificacion_1 & clasificacion_2 & clasificacion_3 & clasificacion_4 & clasificacion_5 & clasificacion_6 & clasificacion_7 & clasificacion_8 & clasificacion_9 & clasificacion_10 & clasificacion_11 & clasificacion_12 & clasificacion_13 & clasificacion_14 & clasificacion_15

	' Sustituimos "Carc. Cat." por "Carc.Cat." para unificar
	clasificacion_rd = replace(clasificacion_rd, "Carc. Cat.", "Carc.Cat.")

	' Quitamos los espacios en blanco
	clasificacion_rd = replace(clasificacion_rd, " ", "")

	' Buscamos la primera aparicion de "Carc.Cat."
	posicion = instr(1,clasificacion_rd, "Carc.Cat.")

	' Sacamos el nivel como el caracter que hay justo detrás de la primera aparición de la subcadena

	if (posicion > 0) then
		dame_nivel_cancerigeno_rd = mid(clasificacion_rd, posicion+9, 1)
	else
		dame_nivel_cancerigeno_rd = ""
	end if
end function

' #############################################################################################

function dame_nivel_mutageno_rd()
	' Juntamos todas las clasificaciones
	clasificacion_rd = clasificacion_1 & clasificacion_2 & clasificacion_3 & clasificacion_4 & clasificacion_5 & clasificacion_6 & clasificacion_7 & clasificacion_8 & clasificacion_9 & clasificacion_10 & clasificacion_11 & clasificacion_12 & clasificacion_13 & clasificacion_14 & clasificacion_15

	' Sustituimos "Muta. Cat." por "Muta.Cat." para unificar
	clasificacion_rd = replace(clasificacion_rd, "Muta. Cat.", "Muta.Cat.")

	' Quitamos los espacios en blanco
	clasificacion_rd = replace(clasificacion_rd, " ", "")

	'response.write "["&clasificacion_rd&"]"

	' Buscamos la primera aparicion de "Muta.Cat."
	posicion = instr(1,clasificacion_rd, "Muta.Cat.")

	' Sacamos el nivel como el caracter que hay justo detrás de la primera aparición de la subcadena
	if (posicion > 0) then
		dame_nivel_mutageno_rd = mid(clasificacion_rd, posicion+9, 1)
	else
		dame_nivel_mutageno_rd = ""
	end if
end function
%>
