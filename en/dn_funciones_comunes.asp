<!--#include file="../EliminaInyeccionSQL.asp"-->
<%
sub paginacion
%>
 <strong>Pages: </strong><br />
<%
	totalpags=roundsup(hr/nregs)
	if pag>1 then
%>
	<a href='#' onclick='cambiapag(<%=pag-1%>)'>&lt; Previous</a>
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
	<a href='#' onclick='cambiapag(<%=pag+1%>)'>Next &gt;</a>
<%
	end if

end sub

' #############################################################################

function describe_simbolo(byval simbolo)
	' Devuelve la descripción del símbolo consultando la base de datos
	sql="SELECT descripcion_ing FROM dn_simbolos WHERE simbolo='"&trim(simbolo)&"'"
	set objRst=objConnection2.execute(sql)

	if (objRst.eof) then
		descripcion = ""
	else
		descripcion = objRst("descripcion_ing")
	end if

	objRst.close()
	set objRst=nothing

	describe_simbolo = descripcion
end function

' #############################################################################

function describe_frase(tipo, byval frase)
	' Devuelve la descripción de la frase consultando la base de datos

	' Sustituye "-" por "/" para unificar formato
	frase = replace(frase, "-", "/")

	sql="SELECT texto_ing FROM dn_risc_frases_"+tipo+" WHERE frase='"&frase&"'"
	set objRst=objConnection2.execute(sql)

	if (objRst.eof) then
		descripcion = ""
	else
		descripcion = objRst("texto_ing")
	end if

	objRst.close()
	set objRst=nothing

	describe_frase = descripcion
end function


' #############################################################################

function describe_categoria_peligro(byval frase)
	' Devuelve la descripción de la frase consultando la base de datos

	' Sustituye "-" por "/" para unificar formato
	frase = replace(frase, "-", "/")

	sql="SELECT texto_ing,frase_ing FROM dn_risc_categorias_peligro WHERE frase='"&frase&"'"
	set objRst=objConnection2.execute(sql)

	if (objRst.eof) then
		frase= ""
		descripcion = ""
	else
		frase = objRst("frase_ing")
		descripcion = objRst("texto_ing")
	end if

	objRst.close()
	set objRst=nothing

	Dim fraseArray(1)
	fraseArray(0) = frase
	fraseArray(1) = descripcion

	describe_categoria_peligro = fraseArray

end function


' ##################################################################################

function dame_definicion(byval cadena)
	' Devuelve la definicion de la tabla rq_definiciones de la base antigua
	sql = "SELECT definicion_eng FROM rq_definiciones where palabra='"&cadena&"'"
	set objRst=objConnection.execute(sql)
	if (objRst.eof) then
		definicion = "No definition found for: <b>"&cadena&"</b>"
	else
		definicion = objRst("definicion_eng")
	end if
	objRst.close()
	set objRst=nothing

	dame_definicion=definicion
end function

function dame_nombre_en_ingles_definicion(byval cadena)
	' Devuelve la palabra en inglés
	sql = "SELECT palabra_eng FROM rq_definiciones where palabra='"&cadena&"'"
'response.write sql
	set objRst=objConnection.execute(sql)
	if (objRst.eof) then
		nombre = "No name found for: <b>"&cadena&"</b>"
	else
		nombre = objRst("palabra_eng")
	end if
	objRst.close()
	set objRst=nothing

	dame_nombre_en_ingles_definicion=nombre
end function



%>
