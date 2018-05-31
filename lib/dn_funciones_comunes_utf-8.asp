<!--#include file="../EliminaInyeccionSQL.asp"-->
<%
''on error resume next
' MENSAJES DE ERROR
' ##########################################################################
function flashMsgShow()

	'Si existe la variable de sessión "flashMsg", la muestra y después la borra
	if (not session("flashMsg")="") then
%><fieldset id="flashmsg"><legend class="<%=lcase(session("flashType"))%>"><strong><%=session("flashType")%></strong></legend><%=session("flashMsg")%></fieldset>
<%
		session("flashType")=""
		session("flashMsg")=""
	end if
end function

' ##########################################################################

function flashMsgCreate(msg, tipo)

	'Crea mensaje de error
	session("flashType")=tipo
	session("flashMsg")=msg

end function

function comprobarl(valor,max,nombre)
	if len(valor)>max then
		comprobarl="<br />-Se ha sobrepasado la longitud máxima (" &max& ") para el campo " &nombre& ": " &valor
	else
		comprobarl=""
	end if
end function

' ##########################################################################

sub dameOpciones(byval campo1, byval campo2, byval tabla, byval orderby, byval seleccionada, byval pordefectotxt, byval pordefectocod, byval concatenar)
	' Pinta las options para el select
	' Si se pasa una cadena no vacía para concatenar, en el texto de cada opción se muestra campo1+concatenar+campo2. Si no, solo campo2

	' Opción por defecto
	if (pordefectotxt <> "") then
%>

		<option value="<%=pordefectocod%>"><%=pordefectotxt%></option>
<%
	end if

	' Traemos listado de tabla
	sql="select "&campo1&", "&campo2&" from "&tabla&" ORDER BY "&orderby
	set objRst = objConnection2.execute(sql)

	do while (not objRst.eof)
		if (concatenar <> "") then
			texto = objRst(campo1)&concatenar&objRst(campo2)
		else
			texto = objRst(campo2)
		end if
%>
		<option value="<%=objRst(campo1)%>"><%=texto%></option>
<%
		objRst.movenext
	loop

	objRst.Close
	set objRst=nothing
end sub

' ##########################################################################

function dameSinonimos(byval id_sustancia)
	' Devuelve lista de sinónimos para la sustancia indicada
	sinonimos = ""

	sql="SELECT dn_risc_sinonimos.nombre AS sinonimo, dn_risc_sustancias.nombre FROM dn_risc_sinonimos INNER JOIN dn_risc_sustancias ON dn_risc_sinonimos.id_sustancia = dn_risc_sustancias.id WHERE dn_risc_sinonimos.id_sustancia="&id_sustancia&" AND dn_risc_sinonimos.nombre <> dn_risc_sustancias.nombre ORDER BY dn_risc_sinonimos.nombre"
	set objRst=objConnection2.execute(sql)
	if (not objRst.eof) then
		sinonimos = sinonimos & "<ul>"
		do while (not objRst.eof)
			sinonimos = sinonimos &"<li>"&h(espaciar(objRst("sinonimo")))&"</li>"
			objRst.movenext
		loop
		sinonimos = sinonimos & "</ul>"
	end if
	objRst.close()
	set objRst=nothing

	dameSinonimos = sinonimos
end function

function dameNombreingles(byval id_sustancia)
	' Devuelve nombreingles para la sustancia indicada
	cad = ""

	sql="SELECT Nombre_ing FROM dn_risc_sustancias  WHERE id="&id_sustancia
	set objRst=objConnection2.execute(sql)
	if objRst("Nombre_ing")<>"" then
		cad = cad & "<ul>"
		do while (not objRst.eof)
			cad = cad &"<li>"&quita_arroba(h(espaciar(elimina_repes(objRst("Nombre_ing"),"@"))))&"</li>"
			objRst.movenext
		loop
		cad = cad & "</ul>"
	end if
	objRst.close()
	set objRst=nothing

	dameNombreingles = cad
end function

function quita_arroba(byval cadena)
  cadena=replace(cadena, "@", "</li><li>")
  cadena=replace(cadena, "@", "</li><li>")
  cadena=replace(cadena, "@", "</li><li>")
  cadena=replace(cadena, "@", "</li><li>")
  cadena=replace(cadena, "@", "</li><li>")
  cadena=replace(cadena, "@", "</li><li>")
  cadena=replace(cadena, "@", "</li><li>")
  cadena=replace(cadena, "@", "</li><li>")
  cadena=replace(cadena, "@", "</li><li>")
  cadena=replace(cadena, "@", "</li><li>")
  quita_arroba=cadena
end function

function dameNombrecomercial(byval id_sustancia)
	' Devuelve lista de Nombrecomercial para la sustancia indicada
	cad = ""

	sql="SELECT dn_risc_nombres_comerciales.nombre AS nc, dn_risc_sustancias.nombre FROM dn_risc_nombres_comerciales INNER JOIN dn_risc_sustancias ON dn_risc_nombres_comerciales.id_sustancia = dn_risc_sustancias.id WHERE dn_risc_nombres_comerciales.id_sustancia="&id_sustancia&" AND dn_risc_nombres_comerciales.nombre <> dn_risc_sustancias.nombre ORDER BY dn_risc_nombres_comerciales.nombre"
	set objRst=objConnection2.execute(sql)
	if (not objRst.eof) then
		cad = cad & "<ul>"
		do while (not objRst.eof)
			cad = cad &"<li>"&h(espaciar(objRst("nc")))&"</li>"
			objRst.movenext
		loop
		cad = cad & "</ul>"
	end if
	objRst.close()
	set objRst=nothing

	dameNombrecomercial = cad
end function
%>

<%
'************BÚSQUEDA (TILDES)*********
function quitartildes(byVal termino)
	if (isnull(termino)) then
		quitartildes = null
	else

		' Reemplazamos todas las tildes.
		termino = replace(termino,"á","a")
		termino = replace(termino,"é","e")
		termino = replace(termino,"í","i")
		termino = replace(termino,"ó","o")
		termino = replace(termino,"ú","u")

		termino = replace(termino,"à","a")
		termino = replace(termino,"è","e")
		termino = replace(termino,"ì","i")
		termino = replace(termino,"ò","o")
		termino = replace(termino,"ù","u")

		termino = replace(termino,"ü","u")

		termino = replace(termino,"Á","A")
		termino = replace(termino,"É","E")
		termino = replace(termino,"Í","I")
		termino = replace(termino,"Ó","O")
		termino = replace(termino,"Ú","U")

		termino = replace(termino,"À","A")
		termino = replace(termino,"È","E")
		termino = replace(termino,"Ì","I")
		termino = replace(termino,"Ò","O")
		termino = replace(termino,"Ù","U")

		termino = replace(termino,"Ü","U")

		quitartildes=termino
	end if
end function

function montartildes(byVal termino)

	' pasamos a exp. regular con todas las posibilidades

	termino = replace(termino,"a","[aáà]")
	termino = replace(termino,"e","[eéè]")
	termino = replace(termino,"i","[iíì]")
	termino = replace(termino,"o","[oóò]")
	termino = replace(termino,"u","[uúùü]")

	montartildes=termino

end function
%>

<%
sub paginacion
%>
 <strong>Páginas: </strong><br />
<%
	totalpags=roundsup(hr/nregs)
	if pag>1 then
%>
	<a href='#' onclick='cambiapag(<%=pag-1%>)'>&lt; Anterior</a>
<%
	end if

	for i=1 to totalpags
		if (cint(i)=cint(pag)) then
			mipag=" <b>" &i& "</b>"
		else
			mipag=" <a href='#' onclick='cambiapag(" &i& ")'>" &i& "</a>"
		end if
		response.write mipag
	next

	if cint(pag)<cint(totalpags) then
%>
	<a href='#' onclick='cambiapag(<%=pag+1%>)'>Siguiente &gt;</a>
<%
	end if

end sub
%>

<%
'VARIOS************************************

function roundsup (pes) 'Redondea a enteros superiores.

	ncoma=instr(pes, ",")
	if ncoma>0 then
	parteentera=int(pes)
	pes=parteentera+1
	end if
	roundsup=(pes)

end function

function quitaultimoscar(cadena,ncars)
	quitaultimoscar=cadena
	if cadena<>"" then
		if len(cadena)>ncars then	quitaultimoscar=left(cadena,(len(cadena)-ncars))
	end if
end function
%>


<%
' #############################################################################
' FRASES R
' #############################################################################

function monta_frases(tipo, c1, c2, c3, c4, c5, c6, c7, c8, c9, c10, c11, c12, c13, c14, c15)

	' Cada llamada va concatenando a las frases acumuladas anteriormente
	frases = ""
	frases = extrae_frase(c1, frases, tipo)
	frases = extrae_frase(c2, frases, tipo)
	frases = extrae_frase(c3, frases, tipo)
	frases = extrae_frase(c4, frases, tipo)
	frases = extrae_frase(c5, frases, tipo)
	frases = extrae_frase(c6, frases, tipo)
	frases = extrae_frase(c7, frases, tipo)
	frases = extrae_frase(c8, frases, tipo)
	frases = extrae_frase(c9, frases, tipo)
	frases = extrae_frase(c10, frases, tipo)
	frases = extrae_frase(c11, frases, tipo)
	frases = extrae_frase(c12, frases, tipo)
	frases = extrae_frase(c13, frases, tipo)
	frases = extrae_frase(c14, frases, tipo)
	frases = extrae_frase(c15, frases, tipo)

	monta_frases=frases
end function

' #############################################################################

function monta_frases_r_danesa(byval frases_r)
	' Las frases R danesas vienen separadas por espacios, y para cada una si tiene simbolo, separado por punto y coma

	frases = ""
	array_1 = split (frases_r, " ")
	for i=0 to ubound(array_1)
		'response.write "<br />"&array_1(i)
		' Para cada frase sustituimos punto y coma por espacio para usar el mismo formato que RD y poder extraer la frase
		array_1(i) = replace(array_1(i), ";", " ")
		'response.write "<br />"&array_1(i)
		frases = extrae_frase(array_1(i), frases, "R")
		'response.write "<br />"&frases
	next

	' Devolvemos las frases R danesas
	monta_frases_r_danesa = frases
end function

' #############################################################################

function arregla_frases(byval c, tipo)
	' En casos como el DDT que tiene frases como R25-48/25, hay que convertir a R25 R48/25, o sea, cambiar "-" por " R",
	' pero solo en los casos en que tenga "-" y "/"

	' Lo dividimos primero separando por espacios, arreglamos cada una y lo volvemos a unir
	c2=""
	if isnull(c) then c = ""
	array_c = split(c, " ")
	for i=0 to ubound(array_c)
		if ((instr(array_c(i), "-") <> 0) and (instr(array_c(i), "/") <> 0)) then
			array_c(i)=replace(array_c(i), "-", " "+tipo)
		end if
		c2=c2&" "&array_c(i)
	next

'	response.write "<br />"&c&" se convierte en "&c2

	arregla_frases=c2
end function

' #############################################################################

function extrae_frase(c,f, tipo)
	' Saca las frases R, eliminando el símbolo

	' Arreglamos la frase en caso de que tenga "-" y "/"
	c=arregla_frases(c, tipo)

	' Limpiamos la clasificación para quedarnos con las frases
	array_frases = split(c, " ")
	nuevo_c = ""
	for i=0 to ubound(array_frases)
		' Para que sea frase R ha de tener longitud 2 o mayor y comenzar por R más un digito ( o H)
		' Ej.: "R1", "R10", "R1/6"

		if (es_frase(array_frases(i),tipo)) then
			if (nuevo_c="") then
				nuevo_c = array_frases(i)
			else
				nuevo_c = nuevo_c&", "&array_frases(i)
			end if
		end if
	next

	if (nuevo_c <> "") then
		' La clasificación no es vacía, concatenamos a la frase
		if (f <> "") then
			' Ya hay algo en las frases, concateno
			extrae_frase = f & ", " & nuevo_c
		else
			' No hay nada, devuelvo clasificación
			extrae_frase = nuevo_c
		end if
	else
		' La clasificacion es vacía, devolvemos la frase tal cual
		extrae_frase = f
	end if
end function

' #############################################################################

function describe_frase(tipo, byval frase)
	' Devuelve la descripción de la frase consultando la base de datos

	' Sustituye "-" por "/" para unificar formato
	frase = replace(frase, "-", "/")

	sql="SELECT texto FROM dn_risc_frases_"+tipo+" WHERE frase='"&frase&"'"
	set objRst=objConnection2.execute(sql)

	if (objRst.eof) then
		descripcion = ""
	else
		descripcion = objRst("texto")
	end if

	objRst.close()
	set objRst=nothing

	describe_frase = descripcion
end function

function describe_frase_international(tipo, byval frase, lang)
	frase = replace(frase, "-", "/")

	texto_field_name = "texto"
	if lang = "en" then
		texto_field_name = "texto_ing"
	end if

	sql = "SELECT " & texto_field_name & " FROM dn_risc_frases_" & tipo & " WHERE frase='" & frase & "'"
	set objRst = objConnection2.execute(sql)

	if (objRst.eof) then
		descripcion = ""
	else
		descripcion = objRst(texto_field_name)
	end if

	objRst.close()
	set objRst = nothing

	describe_frase_international = descripcion
end function

' #############################################################################

function describe_frase_s(byval frase)
	' Devuelve la descripción de la frase consultando la base de datos

	' Sustituye "-" por "/" para unificar formato
	frase = replace(frase, "-", "/")

	sql="SELECT texto FROM dn_risc_frases_s WHERE frase='"&frase&"'"
	set objRst=objConnection2.execute(sql)

	if (objRst.eof) then
		descripcion = ""
	else
		descripcion = objRst("texto")
	end if

	objRst.close()
	set objRst=nothing

	describe_frase_s = descripcion
end function

' #############################################################################

function describe_categoria_peligro(byval frase)
	' Devuelve la descripción de la frase consultando la base de datos

	' Sustituye "-" por "/" para unificar formato
	frase = replace(frase, "-", "/")

	sql="SELECT texto FROM dn_risc_categorias_peligro WHERE frase='"&frase&"'"
	set objRst=objConnection2.execute(sql)

	if (objRst.eof) then
		descripcion = ""
	else
		descripcion = objRst("texto")
	end if

	objRst.close()
	set objRst=nothing

	describe_categoria_peligro = descripcion
end function

function describe_categoria_peligro_international(byval frase, lang)
	frase = replace(frase, "-", "/")
	frase_field_name = "frase"
	texto_field_name = "texto"

	if lang = "en" then
		texto_field_name = "texto_ing"
		frase_field_name = "frase_ing"
	end if
	sql = "SELECT " & texto_field_name & ", " & frase_field_name & " FROM dn_risc_categorias_peligro WHERE frase='" & frase & "'"
	set objRst=objConnection2.execute(sql)

	if (objRst.eof) then
		frase = ""
		descripcion = ""
	else
		frase = objRst(frase_field_name)
		descripcion = objRst(texto_field_name)
	end if

	objRst.close()
	set objRst = nothing

	Dim fraseArray(1)
	fraseArray(0) = frase
	fraseArray(1) = descripcion

	describe_categoria_peligro_international = fraseArray
end function

' #############################################################################

function describe_simbolo(byval simbolo)
	' Devuelve la descripción del símbolo consultando la base de datos
	sql="SELECT descripcion FROM dn_simbolos WHERE simbolo='"&trim(simbolo)&"'"

	set objRst=objConnection2.execute(sql)

	if (objRst.eof) then
		descripcion = ""
	else
		descripcion = objRst("descripcion")
	end if

	objRst.close()
	set objRst=nothing

	describe_simbolo = descripcion
end function

function get_symbol_description(byval simbolo, lang)
	' Devuelve la descripción del símbolo consultando la base de datos
	dim description_field_name : description_field_name = "descripcion"

	if lang = "en" then
		description_field_name = "descripcion_ing"
	end if

	sql = "SELECT " & description_field_name & " FROM dn_simbolos WHERE simbolo='" & trim(simbolo) & "'"

	set objRst=objConnection2.execute(sql)

	if (objRst.eof) then
		descripcion = ""
	else
		descripcion = objRst(description_field_name)
	end if

	objRst.close()
	set objRst = nothing

	get_symbol_description = descripcion
end function

' #############################################################################

function imagen_simbolo(byval simbolo)
	' Devuelve la imagen del símbolo consultando la base de datos
	sql="SELECT imagen FROM dn_simbolos WHERE simbolo='"&trim(simbolo)&"'"
	set objRst=objConnection2.execute(sql)

	if (objRst.eof) then
		imagen = ""
	else
		imagen = objRst("imagen")
	end if

	objRst.close()
	set objRst=nothing

	imagen_simbolo = imagen
end function

' ##################################################################################

function dame_id_definicion(byval cadena)

	' Devuelve el id de la tabla rq_definiciones de la base antigua

	sql = "SELECT id FROM rq_definiciones where palabra='"&cadena&"'"

	set objRst=objConnection.execute(sql)
	if (objRst.eof) then
		id = -1
	else
		id = objRst("id")
	end if

	objRst.close()
	set objRst=nothing

	dame_id_definicion=id

end function

function get_definition(palabra, lang)
	Dim definition_fileld_name : definition_fileld_name = "definicion"
	Dim definition : definition = ""

	if lang = "en" then
		definition_fileld_name = "definicion_eng"
	end if
	sql = "SELECT " & definition_fileld_name & " FROM rq_definiciones where palabra='" & palabra & "'"
	set objRst = objConnection.execute(sql)
	if not objRst.eof then
		definition = objRst(definition_fileld_name)
	end if
	objRst.close()
	set objRst = nothing

	get_definition = definition
end function

' ##################################################################################

function parche_definicion(byval cadena, byval tipo)
	' Aplica parche a definiciones dependiendo del tipo
	nuevacadena = ""
	select case tipo
		case "VLA", "VLB":
			select case cadena
				case "1", "2", "3", "4", "5", "6", "7", "8", "o":
					nuevacadena = "("&cadena&")"
				case "F", "I", "S":
					nuevacadena = lcase(cadena)&"."
				case else:
					nuevacadena = cadena
			end select

		case "MMA":
			select case cadena
				case "1", "2", "3":
					nuevacadena = cadena & "."
				case else:
					nuevacadena = cadena
			end select

		case "MomentoVLB":
			' Me quedo con los tres ultimos caracteres, ejemplo: "Después de la jornada laboral (5)" -> "(5)"
			nuevacadena = right(cadena,3)

		case "MomentoVLBInicio":
			' Me quedo con todo menos los tres ultimos caracteres, ejemplo: "Después de la jornada laboral (5)" -> "Después de la jornada laboral "
			nuevacadena = left(cadena,(len(cadena)-3))
	end select

	if (nuevacadena <> "") then
		parche_definicion = nuevacadena
	else
		parche_definicion = cadena
	end if
end function

' ##################################################################################

function dame_id_uso(byval cadena)
	' Devuelve el id de la tabla dn_alternativas_usos para el uso indicado
	sql = "SELECT id FROM dn_risc_usos where nombre='"&cadena&"'"
	set objRstUso=objConnection2.execute(sql)
	if (objRstUso.eof) then
		id = -1
	else
		id = objRstUso("id")
	end if
	objRstUso.close()
	set objRstUso=nothing

	dame_id_uso=id
end function

' ##################################################################################

' Montamos las diferentes condiciones para ser usadas despues

' CANCER RD
' R40, R45, R49, R40/20, R40/21, R40/22, R40/20/21, R40/20/22, R40/21/22, R40/20/21/22
campos="sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15"
frases = "R40, R45, R49, R40/20, R40/21, R40/22, R40/20/21, R40/20/22, R40/21/22, R40/20/21/22"
sql_lista_cancer_rd="select distinct sus.id, sus.nombre from dn_risc_sustancias as sus WHERE "&monta_condicion(campos, frases)

' CANCER LISTA DANESA
' R40, R45, R49, R40/20, R40/21, R40/22, R40/20/21, R40/20/22, R40/21/22, R40/20/21/22
campos="sus.frases_r_danesa"
frases = "R40, R45, R49, R40/20, R40/21, R40/22, R40/20/21, R40/20/22, R40/21/22, R40/20/21/22"
sql_lista_cancer_danesa="select distinct sus.id, sus.nombre from dn_risc_sustancias as sus WHERE "&monta_condicion(campos, frases)


' MUTAGENO RD
' R46, R68, R68/20, R68/21, R68/22, R68/20/21, R68/20/22, R68/21/22, R68/20/21/22
campos="sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15"
frases = "R46, R68, R68/20, R68/21, R68/22, R68/20/21, R68/20/22, R68/21/22, R68/20/21/22"
sql_lista_mutageno_rd="select distinct sus.id, sus.nombre from dn_risc_sustancias as sus WHERE "&monta_condicion(campos, frases)


' MUTAGENO LISTA DANESA
' R46, R68, R68/20, R68/21, R68/22, R68/20/21, R68/20/22, R68/21/22, R68/20/21/22
campos="sus.frases_r_danesa"
frases = "R46, R68, R68/20, R68/21, R68/22, R68/20/21, R68/20/22, R68/21/22, R68/20/21/22"
sql_lista_mutageno_danesa="select distinct sus.id, sus.nombre from dn_risc_sustancias as sus WHERE "&monta_condicion(campos, frases)


' CANCER IARC
sql_lista_cancer_iarc = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_iarc ON (sus.id=dn_risc_sustancias_iarc.id_sustancia) WHERE (dn_risc_sustancias_iarc.grupo_iarc<>'')"

' CANCER IARC EXCEPTO GRUPO 3
sql_lista_cancer_iarc_excepto_grupo_3 = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_iarc ON (sus.id=dn_risc_sustancias_iarc.id_sustancia) WHERE (dn_risc_sustancias_iarc.grupo_iarc<>'' AND dn_risc_sustancias_iarc.grupo_iarc NOT LIKE '%3%')"

' CANCER OTRAS
sql_lista_cancer_otras = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_cancer_otras ON (sus.id=dn_risc_sustancias_cancer_otras.id_sustancia) WHERE (dn_risc_sustancias_cancer_otras.categoria_cancer_otras<>'')"

' CANCER MAMA
sql_lista_cancer_mama = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus FULL OUTER JOIN dn_risc_sinonimos as sin ON (sus.id=sin.id_sustancia) LEFT OUTER JOIN dn_risc_sustancias_mama_cop ON (sus.id=dn_risc_sustancias_mama_cop.id_sustancia) WHERE (dn_risc_sustancias_mama_cop.cancer_mama=1)"

' COP
sql_lista_cop = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus FULL OUTER JOIN dn_risc_sinonimos as sin ON (sus.id=sin.id_sustancia) LEFT OUTER JOIN dn_risc_sustancias_mama_cop ON (sus.id=dn_risc_sustancias_mama_cop.id_sustancia) WHERE (dn_risc_sustancias_mama_cop.cop<>'')"

' SALUD
sql_lista_salud = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_salud AS sal ON (sus.id=sal.id_sustancia) WHERE (sal.cardiocirculatorio=1 OR sal.rinyon=1 OR sal.respiratorio=1 OR sal.reproductivo=1 OR sal.piel_sentidos=1 OR sal.neuro_toxicos=1 OR sal.musculo_esqueletico=1 OR sal.sistema_inmunitario=1 OR sal.higado_gastrointestinal=1 OR sal.sistema_endocrino=1 OR sal.embrion=1 OR sal.cancer=1)"

' TPR
' Frases R: R60, R61, R62, R63
campos="sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15"
frases = "R60, R61, R62, R63"
sql_lista_tpr="select distinct sus.id, sus.nombre from dn_risc_sustancias as sus WHERE "&monta_condicion(campos, frases)

' TPR DANESA
' Frases R: R60, R61, R62, R63
campos="sus.frases_r_danesa"
frases = "R60, R61, R62, R63"
sql_lista_tpr_danesa="select distinct sus.id, sus.nombre from dn_risc_sustancias as sus WHERE "&monta_condicion(campos, frases)

' DE
sql_lista_de = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_neuro_disruptor ON (sus.id=dn_risc_sustancias_neuro_disruptor.id_sustancia) WHERE (dn_risc_sustancias_neuro_disruptor.nivel_disruptor<>'')"

' NEUROTOXICO
' Neurotoxico es si contiene nivel en la tabla auxiliar, o si tiene frase R67.
' Esta es la condicion antigua...
'sql_lista_neurotoxico = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_neuro_disruptor ON (sus.id=dn_risc_sustancias_neuro_disruptor.id_sustancia) WHERE (dn_risc_sustancias_neuro_disruptor.nivel_neurotoxico<>'')"

' Y esta es la nueva, buscando por los dos casos
sql_lista_neurotoxico_nivel = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_neuro_disruptor ON (sus.id=dn_risc_sustancias_neuro_disruptor.id_sustancia) WHERE (dn_risc_sustancias_neuro_disruptor.nivel_neurotoxico<>'')"

campos="sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15"
frases="R67"
sql_lista_neurotoxico_rd = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus WHERE "&monta_condicion(campos, frases)


' NEUROTOXICO SEGUN LISTA DANESA
campos="sus.frases_r_danesa"
frases="R67"
sql_lista_neurotoxico_danesa = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus WHERE "&monta_condicion(campos, frases)


' NEUROTOXICO PARA LISTADO (SIMILAR A FUNCIONAMIENTO DE LISTA NEGRA, SOLO CONDICION)
sql_lista_neurotoxico = "(sus.id IN ("&qn(sql_lista_neurotoxico_nivel)&") OR sus.id IN ("&qn(sql_lista_neurotoxico_rd)&") OR sus.id IN ("&qn(sql_lista_neurotoxico_danesa)&"))"


' SENSIBILIZANTE
' Que tengan R42, R43 o R42/43
campos="sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15"
frases = "R42, R43, R42/43, R42-43"
sql_lista_sensibilizante="select distinct sus.id, sus.nombre from dn_risc_sustancias as sus WHERE "&monta_condicion(campos, frases)

' SENSIBILIZANTE DANESA
' Que tengan R42, R43 o R42/43
campos="sus.frases_r_danesa"
frases = "R42, R43, R42/43"
sql_lista_sensibilizante_danesa="select distinct sus.id, sus.nombre from dn_risc_sustancias as sus WHERE "&monta_condicion(campos, frases)

'SENSIBILIZANTE REACH
sql_lista_sensibilizante_reach = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus FULL OUTER JOIN dn_risc_sinonimos as sin ON (sus.id=sin.id_sustancia) LEFT OUTER JOIN dn_risc_sensibilizantes_reach AS sen ON (sus.id=sen.id_sustancia)  WHERE (sus.id<>'' AND sen.id_sustancia <> '')"

' EEPP
sql_lista_eepp = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus FULL OUTER JOIN dn_risc_sinonimos as sin ON (sus.id=sin.id_sustancia) LEFT OUTER JOIN dn_risc_sustancias_por_grupos AS spg ON (sus.id=spg.id_sustancia) LEFT OUTER JOIN dn_risc_grupos_por_enfermedades AS gpe ON (spg.id_grupo=gpe.id_grupo) LEFT OUTER JOIN dn_risc_sustancias_por_enfermedades AS spe ON sus.id = spe.id_sustancia WHERE ((sus.id<>'' AND (spe.id_enfermedad IS NOT NULL) OR (gpe.id_enfermedad IS NOT NULL)))"

' TPB
sql_lista_tpb = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (dn_risc_sustancias_ambiente.anchor_tpb<>'')"

' DIRECTIVA AGUAS
sql_lista_directiva_aguas = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (dn_risc_sustancias_ambiente.directiva_aguas=1)"

' ALEMANA
sql_lista_alemana = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus FULL OUTER JOIN dn_risc_sinonimos as sin ON (sus.id=sin.id_sustancia) LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (dn_risc_sustancias_ambiente.clasif_MMA<>''and DN_RISC_SUSTANCIAS_AMBIENTE.CLASIF_MMA <> 'nwg')"

' OZONO
sql_lista_ozono = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (dn_risc_sustancias_ambiente.dano_ozono=1)"

' CLIMA
sql_lista_clima = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (dn_risc_sustancias_ambiente.dano_cambio_clima=1)"

' AIRE
sql_lista_aire = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (dn_risc_sustancias_ambiente.dano_calidad_aire=1)"

'SUELOS Sergio
sql_lista_suelos = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (dn_risc_sustancias_ambiente.toxicidad_suelo=1)"


' COV
sql_lista_cov = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (dn_risc_sustancias_ambiente.cov=1)"

' VERTIDOS
sql_lista_vertidos = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_iarc iarc ON sus.id = iarc.id_sustancia LEFT OUTER JOIN dn_risc_sustancias_cancer_otras otras ON sus.id = otras.id_sustancia LEFT OUTER JOIN dn_risc_sustancias_neuro_disruptor neuro ON sus.id = neuro.id_sustancia LEFT OUTER JOIN dn_risc_sustancias_ambiente ambiente ON sus.id = ambiente.id_sustancia WHERE ((sus.num_rd <> '') OR (sus.frases_r_danesa <> '') OR (iarc.grupo_iarc <> '') OR (otras.categoria_cancer_otras <> '') OR (neuro.nivel_disruptor <> '') OR (ambiente.enlace_tpb <> '') OR (ambiente.directiva_aguas <> '') OR (ambiente.clasif_mma <> ''))"

' LPCIC
sql_lista_lpcic = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (eper_agua<>'' or eper_aire<>'' or eper_suelo<>'')"

' LPCIC-AGUA
sql_lista_lpcic_agua = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (eper_agua<>'')"

' LPCIC-AIRE
sql_lista_lpcic_aire = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (eper_aire<>'')"

' LPCIC-SUELO
sql_lista_lpcic_suelo = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (eper_suelo<>'')"

' SUSTANCIA PRIORITARIA
sql_lista_sustancia_prioritaria = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (sustancia_prioritaria=1)"


' RESIDUOS
sql_lista_residuos = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus FULL OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) FULL OUTER JOIN dn_risc_sustancias_cancer_otras ON (sus.id=dn_risc_sustancias_cancer_otras.id_sustancia) FULL OUTER JOIN dn_risc_sustancias_iarc ON (sus.id=dn_risc_sustancias_iarc.id_sustancia) FULL OUTER JOIN dn_risc_sustancias_neuro_disruptor ON (sus.id=dn_risc_sustancias_neuro_disruptor.id_sustancia) FULL OUTER JOIN dn_risc_sustancias_vl ON (sus.id=dn_risc_sustancias_vl.id_sustancia) WHERE ((sus.num_rd<>'' or sus.frases_r_danesa <> '' or dn_risc_sustancias_ambiente.id is not null or dn_risc_sustancias_cancer_otras.id is not null or dn_risc_sustancias_iarc.id is not null or dn_risc_sustancias_neuro_disruptor.id is not null or dn_risc_sustancias_vl.id is not null ) AND sus.id is not null)"

' ACCIDENTES
sql_lista_accidentes = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (seveso<>'')"

' EMISIONES
sql_lista_emisiones = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (dn_risc_sustancias_ambiente.emisiones_atmosfera=1)"

' PROHIBIDAS
sql_lista_prohibidas = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_prohibidas as pro ON (sus.id=pro.id_sustancia) WHERE "

' RESTRINGIDAS
sql_lista_restringidas = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_restringidas as rest ON (sus.id=rest.id_sustancia) WHERE "


' LISTA NEGRA
' Emplear selects anidados para buscar las sustancias cuyo id esté en cualquiera de las tablas que son de la lista negra
' (esta_en_lista_cancer_rd or esta_en_lista_cancer_iarc_excepto_grupo_3 or esta_en_lista_cancer_otras or esta_en_lista_de or esta_en_lista_neurotoxico or  esta_en_lista_tpb or esta_en_lista_sensibilizante or esta_en_lista_tpr or esta_en_lista_mutageno_rd)

' Solo va la condicion...
'sql_lista_negra = "(sus.id IN ("&qn(sql_lista_cancer_rd)&") OR sus.id IN ("&qn(sql_lista_cancer_danesa)&") OR sus.id IN ("&qn(sql_lista_cancer_iarc_excepto_grupo_3)&") OR sus.id IN ("&qn(sql_lista_cancer_otras)&") OR sus.id IN ("&qn(sql_lista_mutageno_rd)&") OR sus.id IN ("&qn(sql_lista_mutageno_danesa)&") OR sus.id IN ("&qn(sql_lista_de)&") OR sus.id IN ("&qn(sql_lista_neurotoxico)&") OR sus.id IN ("&qn(sql_lista_tpb)&") OR sus.id IN ("&qn(sql_lista_sensibilizante)&") OR sus.id IN ("&qn(sql_lista_sensibilizante_danesa)&") OR sus.id IN ("&qn(sql_lista_tpr)&") OR sus.id IN ("&qn(sql_lista_tpr_danesa)&"))"

' NUEVA VERSION: simplificado en un campo, para agilizar
sql_lista_negra = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus WHERE (dn_risc_sustancias.negra=1)"

' *** INICIO SPL

' PROHIBIDAS EMBARAZADAS
sql_lista_prohibidas_embarazadas = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN spl_risc_sustancias_prohibidas_embarazadas as pro_emb ON (sus.id=pro_emb.id_sustancia) WHERE "

' PROHIBIDAS LACTANTES
sql_lista_prohibidas_lactantes = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN spl_risc_sustancias_prohibidas_lactantes as pro_lac ON (sus.id=pro_lac.id_sustancia) WHERE "

' CANDIDATAS REACH
sql_lista_candidatas_reach = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN spl_risc_sustancias_candidatas_reach as candidatas_reach ON (sus.id=candidatas_reach.id_sustancia) WHERE "

' AUTORIZACION REACH
sql_lista_autorizacion_reach = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN spl_risc_sustancias_autorizacion_reach as autorizacion_reach ON (sus.id=autorizacion_reach.id_sustancia) WHERE "

' BIOCIDAS AUTORIZADAS
sql_lista_biocidas_autorizadas = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN spl_risc_sustancias_biocidas_autorizadas as biocidas_autorizadas ON (sus.id=biocidas_autorizadas.id_sustancia) WHERE "

' BIOCIDAS PROHIBIDAS
sql_lista_biocidas_prohibidas = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN spl_risc_sustancias_biocidas_prohibidas as biocidas_prohibidas ON (sus.id=biocidas_prohibidas.id_sustancia) WHERE "

' PESTICIDAS AUTORIZADAS
sql_lista_pesticidas_autorizadas = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN spl_risc_sustancias_pesticidas_autorizadas as pesticidas_autorizadas ON (sus.id=pesticidas_autorizadas.id_sustancia) WHERE "

' PESTICIDAS PROHIBIDAS
sql_lista_pesticidas_prohibidas = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN spl_risc_sustancias_pesticidas_prohibidas as pesticidas_prohibidas ON (sus.id=pesticidas_prohibidas.id_sustancia) WHERE "


' *** FIN SPL

' ##LOLO
sql_lista_corap = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus left outer join ist_risc_sustancias_corap as sustancias_corap ON (sus.id = sustancias_corap.id_sustancia) WHERE "

' ##FIN LOLO

'response.write sql_lista_negra

' ##################################################################################
function qn(byval cadena)
  qn=replace(cadena, ", sus.nombre", "")
end function

' ##################################################################################

function esta_en_lista(byval lista, byval id_sustancia)

	' Montamos condicion inicial dependiendo de lista, como en buscador publico de risctox pero sin sinónimos
	select case lista
		case "cancer_rd": ' Cancerigeno según RD
      sql_lista = parentesis_where(sql_lista_cancer_rd) & " OR ("&monta_condicion_grupo("asoc_cancer_rd")&") )"

		case "cancer_danesa": ' Cancerigeno según lista danesa
      sql_lista = sql_lista_cancer_danesa

		case "mutageno_rd": ' Mutágeno según RD
      sql_lista = sql_lista_mutageno_rd


		case "mutageno_danesa": ' Mutágeno según lista danesa
      sql_lista = sql_lista_mutageno_danesa

		case "cancer_iarc": ' Cancerígena según IARC
      sql_lista = parentesis_where(sql_lista_cancer_iarc) & " OR ("&monta_condicion_grupo("asoc_cancer_iarc")&") )"

		case "cancer_iarc_excepto_grupo_3": ' Cancerígena según IARC, excepto Grupo 3
      sql_lista = sql_lista_cancer_iarc_excepto_grupo_3

		case "cancer_otras": ' Cancerígena según otras fuentes
      sql_lista = parentesis_where(sql_lista_cancer_otras) & " OR ("&monta_condicion_grupo("asoc_cancer_otras")&") )"

		case "cancer_otras_excepto_grupo_4":
      sql_lista = parentesis_where(sql_lista_cancer_otras) & " OR ("&monta_condicion_grupo("asoc_cancer_otras")&") ) AND dn_risc_sustancias_cancer_otras.categoria_cancer_otras not like '%G-A4%'"

		case "cancer_mama": ' Cancerígena mama
      sql_lista = parentesis_where(sql_lista_cancer_mama) & " OR ("&monta_condicion_grupo("asoc_cancer_mama")&") )"

		case "cop": ' COP
      sql_lista = parentesis_where(sql_lista_cop) & " OR ("&monta_condicion_grupo("asoc_cop")&") )"

		case "salud": ' Efectos para la salud y órganos afectados
      sql_lista = sql_lista_salud

		case "tpr": ' Tóxicos para la reproducción
      sql_lista = parentesis_where(sql_lista_tpr) & " OR ("&monta_condicion_grupo("asoc_reproduccion")&") )"

		case "tpr_danesa": ' Tóxicos para la reproducción según lista danesa
      sql_lista = sql_lista_tpr_danesa

		case "de": ' Disruptor endocrino
      sql_lista = parentesis_where(sql_lista_de) & " OR ("&monta_condicion_grupo("asoc_disruptores")&") )"

		case "neurotoxico": ' Neurótoxico (RD o Danesa o por nivel)
'      sql_lista = parentesis_where(sql_lista_neurotoxico) & " OR ("&monta_condicion_grupo("asoc_neuro_oto")&") )"
      sql_lista = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus WHERE " & " ("&monta_condicion_grupo("asoc_neuro_oto")&") "

		case "neurotoxico_rd": ' Neurótoxico RD
      sql_lista = sql_lista_neurotoxico_rd

		case "neurotoxico_danesa": ' Neurótoxico Danesa
      sql_lista = sql_lista_neurotoxico_danesa

		case "neurotoxico_nivel": ' Neurótoxico Danesa
      sql_lista = sql_lista_neurotoxico_nivel

		case "sensibilizante": ' Sensibilizante
      sql_lista = sql_lista_sensibilizante

		case "sensibilizante_danesa": ' Sensibilizante según lista danesa
      sql_lista = sql_lista_sensibilizante_danesa

	  case "sensibilizante_reach": ' Sensibilizante según reach
      'sql_lista = sql_lista_sensibilizante_reach
      sql_lista = parentesis_where(sql_lista_sensibilizante_reach) & " OR ("&monta_condicion_grupo("asoc_alergenos")&") )"

		case "eepp": ' Enfermedades profesionales relacionadas
      sql_lista = sql_lista_eepp

		case "tpb": ' Tóxicas, persistentes y bioacumulativas
      'sql_lista = sql_lista_tpb
      sql_lista = parentesis_where(sql_lista_tpb) & " OR ("&monta_condicion_grupo("asoc_tpb")&") )"

		case "directiva_aguas": ' Directiva de aguas
      sql_lista = parentesis_where(sql_lista_directiva_aguas) & " OR ("&monta_condicion_grupo("asoc_directiva_aguas")&") )"
	  'response.write sql_lista

		case "sustancias_prioritarias": '
      sql_lista = sql_lista_sustancia_prioritaria

		case "alemana": ' Alemana de aguas
      'sql_lista = sql_lista_alemana
      sql_lista = parentesis_where(sql_lista_alemana) & " OR ("&monta_condicion_grupo("asoc_peligrosas_agua_alemania")&") )"

		case "ozono": ' Capa de ozono
      'sql_lista = sql_lista_ozono
      sql_lista = parentesis_where(sql_lista_ozono) & " OR ("&monta_condicion_grupo("asoc_capa_ozono")&") )"

		case "clima": ' Cambio climático
      sql_lista = sql_lista_clima
      sql_lista = parentesis_where(sql_lista_clima) & " OR ("&monta_condicion_grupo("asoc_cambio_climatico")&") )"

		case "aire": ' Calidad del aire
      sql_lista = parentesis_where(sql_lista_aire) & " OR ("&monta_condicion_grupo("asoc_calidad_aire")&") )"

		case "cov": ' COV
      'sql_lista = sql_lista_cov
      sql_lista = parentesis_where(sql_lista_cov) & " OR ("&monta_condicion_grupo("asoc_cov")&") )"

	  case "suelos": ' Contaminante suelos
      'sql_lista = sql_lista_suelos
      sql_lista = parentesis_where(sql_lista_suelos) & " OR ("&monta_condicion_grupo("asoc_contaminantes_suelo")&") )"

		case "vertidos": ' Vertidos
      sql_lista = sql_lista_vertidos

		case "lpcic": ' LPCIC
      'sql_lista = parentesis_where(sql_lista_lpcic) & " OR ("&monta_condicion_grupo("asoc_eper")&") )"
      sql_lista = sql_lista_lpcic

		case "lpcic-agua": ' LPCIC Agua
      'sql_lista = parentesis_where(sql_lista_lpcic_agua) & " OR ("&monta_condicion_grupo("asoc_eper")&") )"
      sql_lista = sql_lista_lpcic_agua

		case "lpcic-aire": ' LPCIC Aire
      'sql_lista = parentesis_where(sql_lista_lpcic_aire) & " OR ("&monta_condicion_grupo("asoc_eper")&") )"
      sql_lista = sql_lista_lpcic_aire

	  case "lpcic-suelo": ' LPCIC Aire
      'sql_lista = parentesis_where(sql_lista_lpcic_aire) & " OR ("&monta_condicion_grupo("asoc_eper")&") )"
      sql_lista = sql_lista_lpcic_suelo

		case "residuos": ' Residuos peligrosos
      sql_lista = sql_lista_residuos

		case "accidentes": ' Accidentes graves
      sql_lista = parentesis_where(sql_lista_accidentes) & " OR ("&monta_condicion_grupo("asoc_seveso")&") )"

		case "emisiones": ' Emisiones atmosféricas
      sql_lista = parentesis_where(sql_lista_emisiones) & " OR ("&monta_condicion_grupo("asoc_emisiones_atmosfericas")&") )"

		case "prohibidas": ' Sustancias prohibidas
      		sql_lista = parentesis_where(sql_lista_prohibidas) & "(sus.id=pro.id_sustancia) OR ("&monta_condicion_grupo("asoc_prohibidas")&"))"

		case "restringidas": ' Sustancias restringidas
      		sql_lista = parentesis_where(sql_lista_restringidas) & "(sus.id=rest.id_sustancia) OR ("&monta_condicion_grupo("asoc_restringidas")&"))"

	' SPL NUEVAS LISTAS
		case "prohibidas_embarazadas": '
			campos="sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15"
      		frases="R60, R61"
      		sql_lista = parentesis_where(sql_lista_prohibidas_embarazadas)
			sql_lista = sql_lista & "(sus.id=pro_emb.id_sustancia) OR  " ' Lista de sustancias prohibidas para embarazadas
			sql_lista = sql_lista & "( " & monta_condicion(campos, frases) & " OR " ' Sustancias con R60 y R61
			sql_lista = sql_lista & " (sus.num_rd='082-001-00-6' ) " ' Sustancias con rd=082-001-00-6
			sql_lista = sql_lista & " OR ("&monta_condicion_grupo("asoc_prohibidas_embarazadas")&" )"

			sql_lista = sql_lista & " OR (sus.num_rd='650-017-00-8' AND sus.num_rd='650-016-00-2')"
			sql_lista = sql_lista & "))"


'response.write sql_lista
		case "prohibidas_lactantes": '
			campos="sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15"
     		frases="R64"
      		sql_lista = parentesis_where(sql_lista_prohibidas_lactantes)
			sql_lista = sql_lista & "(sus.id=pro_lac.id_sustancia) OR  " ' Lista de sustancias prohibidas para lactantes
			sql_lista = sql_lista & "( " & monta_condicion(campos, frases) & " OR " ' Sustancias con R60 y R61
			sql_lista = sql_lista & " (sus.num_rd='082-001-00-6' ) " ' Sustancias con rd=082-001-00-6
			sql_lista = sql_lista & " OR ("&monta_condicion_grupo("asoc_prohibidas_lactantes")&" )"

			sql_lista = sql_lista & " OR (sus.num_rd='650-017-00-8' AND sus.num_rd='650-016-00-2')"
			sql_lista = sql_lista & "))"

		case "candidatas_reach":
			sql_lista = parentesis_where(sql_lista_candidatas_reach) & "(sus.id=candidatas_reach.id_sustancia)  OR ("&monta_condicion_grupo("asoc_candidatas_reach")&"))"

		case "autorizacion_reach":
			sql_lista = parentesis_where(sql_lista_autorizacion_reach) & "(sus.id=autorizacion_reach.id_sustancia)  OR ("&monta_condicion_grupo("asoc_autorizacion_reach")&"))"

		case "biocidas_autorizadas":
			sql_lista = parentesis_where(sql_lista_biocidas_autorizadas) & "(sus.id=biocidas_autorizadas.id_sustancia)  OR ("&monta_condicion_grupo("asoc_biocidas_autorizadas")&"))"

		case "biocidas_prohibidas":
			sql_lista = parentesis_where(sql_lista_biocidas_prohibidas) & "(sus.id=biocidas_prohibidas.id_sustancia)  OR ("&monta_condicion_grupo("asoc_biocidas_prohibidas")&"))"

		case "pesticidas_autorizadas":
			sql_lista = parentesis_where(sql_lista_pesticidas_autorizadas) & "(sus.id=pesticidas_autorizadas.id_sustancia)  OR ("&monta_condicion_grupo("asoc_pesticidas_autorizadas")&"))"

		case "pesticidas_prohibidas":
			sql_lista = parentesis_where(sql_lista_pesticidas_prohibidas) & "(sus.id=pesticidas_prohibidas.id_sustancia)  OR ("&monta_condicion_grupo("asoc_pesticidas_prohibidas")&"))"

		case "corap"
			sql_lista = sql_lista_corap & "(sus.id = sustancias_corap.id_sustancia)"

	'-- /SPL NUEVAS LISTAS

	end select

	sql_lista = sql_lista & " AND sus.id = " & id_sustancia
	'if lista="alemana" then response.write sql_lista
	'response.write "<br /><br />"&lista&": "&sql_lista
	'response.write "<!--"&sql_lista&"-->"

	set obj_rst_lista = objConnection2.execute(sql_lista)
	if (obj_rst_lista.eof) then
		esta = false
	else
    'response.write "<br /><br />Está en lista <strong>"&lista&"</strong>: "&sql_lista&"<br />"
    'if (lista = "cancer_rd") then
    '  response.write "<br />"&clasificacion_1
    '  response.write "<br />"&clasificacion_2
    '  response.write "<br />"&clasificacion_3
    '  response.write "<br />"&clasificacion_4
    '  response.write "<br />"&clasificacion_5
    '  response.write "<br />"&clasificacion_6
    'end if

		esta = true
	end if

	obj_rst_lista.close()
	set obj_rst_lista = nothing

	esta_en_lista = esta
end function

' ##########################################################################
function parentesis_where(byval cadena)
  ' Añade otro paréntesis al principio del WHERE
  cadena = ucase(cadena)
  parentesis_where = replace(cadena, "WHERE", "WHERE (")
end function

' ##########################################################################

function es_disolvente(byval id_sustancia)
	' Una sustancia es disolvente si está asociada al uso Disolvente
	id_uso_disolvente = dame_id_uso("DISOLVENTE")
	sql="SELECT COUNT(*) AS num FROM dn_risc_sustancias_por_usos WHERE id_sustancia="&id_sustancia&" AND id_uso =" &id_uso_disolvente
	set objRstDis=objConnection2.execute(sql)
	if (objRstDis("num") <> 0) then
		disolvente=1
	else
		disolvente=0
	end if

	objRstDis.close()
	set objRstDis=nothing

  ' También puede estar asociada al uso disolvente a través de un grupo. En caso de que haya fallado el test anterior, buscaremos
  ' por esa condición
  if (disolvente = 0) then
    sql = "SELECT COUNT(*) AS num FROM dn_risc_sustancias AS s INNER JOIN dn_risc_sustancias_por_grupos AS spg ON s.id = spg.id_sustancia INNER JOIN dn_risc_grupos AS g ON spg.id_grupo = g.id INNER JOIN dn_risc_grupos_por_usos AS gpu ON g.id = gpu.id_grupo WHERE gpu.id_uso = "&id_uso_disolvente&" AND s.id="&id_sustancia

  	set objRstDis=objConnection2.execute(sql)
  	if (objRstDis("num") <> 0) then
  		disolvente=1
  	else
  		disolvente=0
  	end if

  	objRstDis.close()
	  set objRstDis=nothing
  end if

	es_disolvente = disolvente
end function

' ##########################################################################

function monta_condicion(byval campos, byval frases)
  ' Helper para montar la parte de SQL donde se buscan frases R en los campos clasificacion_xx y/o frases_r_danesa,
  ' indicando en qué campos buscar (separados por comas) y qué frases (tambien separados por comas)

  ' Ejemplo:
  ' monta_condicion("sus.clasificacion_1, sus_clasificacion_2, sus_clasificacion_3", "R42, R43, R42/43") devuelve:
  ' "((sus.clasificacion_1 LIKE '%R42') OR (sus.clasificacion_1 LIKE '%R42;%') OR (sus.clasificacion_1 LIKE '%R43') OR (sus.clasificacion_1 LIKE '%R43;%') OR (sus.clasificacion_1 LIKE '%R42/43') OR (sus.clasificacion_1 LIKE '%R42/43;%') OR (sus.clasificacion_2 LIKE '%R42') OR (sus.clasificacion_2 LIKE '%R42;%') OR (sus.clasificacion_2 LIKE '%R43') OR (sus.clasificacion_2 LIKE '%R43;%') OR (sus.clasificacion_2 LIKE '%R42/43') OR (sus.clasificacion_2 LIKE '%R42/43;%') OR (sus.clasificacion_3 LIKE '%R42') OR (sus.clasificacion_3 LIKE '%R42;%') OR (sus.clasificacion_3 LIKE '%R43') OR (sus.clasificacion_3 LIKE '%R43;%') OR (sus.clasificacion_3 LIKE '%R42/43') OR (sus.clasificacion_3 LIKE '%R42/43;%'))"

  ' Creamos array campos y de frases con split
  array_campos = split(campos, ",")
  array_frases = split(frases, ",")

  ' Bucleamos para ir montando la condición
  condicion = ""
  for c=0 to ubound(array_campos)
    ' Para cada campo montamos la variante de frase limpia o acabada en punto y coma
    campo = trim(array_campos(c))

    'Bucleamos para cada frase
    for f=0 to ubound(array_frases)
      frase = trim(array_frases(f))
      if (condicion <> "") then
        condicion = condicion&" OR "
      end if
      ' Buscamos al inicio del campo, o separado por ; o separado por espacio (lista danesa)
      condicion = condicion&"("&campo&" LIKE '%"&frase&"') OR ("&campo&" LIKE '%"&frase&";%') OR ("&campo&" LIKE '%"&frase&" %')"
    next

  next

  monta_condicion = "("&condicion&")"

end function

function monta_condicion_grupo(byval check_lista)
  ' Devuelve una cadena para incluir sustancias asociadas a través de grupo en las listas de risctox
  ' Se le debe indicar por parámetro el nombre del checkbox correspondiente a la lista en el formulario de grupo
  ' de la herramienta (ejemplo, "asoc_cop")

  monta_condicion_grupo = "sus.id IN (SELECT DISTINCT spg.id_sustancia FROM dn_risc_grupos AS g INNER JOIN dn_risc_sustancias_por_grupos AS spg ON spg.id_grupo = g.id WHERE g."&check_lista&"=1)"

end function

function monta_condicion_grupo_por_nombre(byval nombre_grupo)
  ' Devuelve una cadena para incluir sustancias asociadas a través de grupo en las listas de risctox
  ' Se le debe indicar por parámetro el nombre del grupo
  ' (ejemplo, "PLOMO Y SUS COMPUESTOS")

  monta_condicion_grupo_por_nombre = "sus.id IN (SELECT DISTINCT spg.id_sustancia FROM dn_risc_grupos AS g INNER JOIN dn_risc_sustancias_por_grupos AS spg ON spg.id_grupo = g.id WHERE g.nombre='"&nombre_grupo&"')"

end function


function navegador()
  if instr(Request.servervariables("HTTP_USER_AGENT"),"Firefox") then
    navegador="FF"
  else
    navegador="noFF"
  end if
end function



%>
