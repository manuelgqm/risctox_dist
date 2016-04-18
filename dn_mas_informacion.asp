<!--#include file="adovbs.inc"-->
<!--#include file="dn_conexion.asp"-->
<!--#include file="dn_funciones_comunes.asp"-->
<!--#include file="dn_funciones_texto.asp"-->
<!--#include file="dn_restringida.asp"-->
  
<%
	id = EliminaInyeccionSQL(request("id"))
	tipo = EliminaInyeccionSQL(request("tipo"))
	listado = EliminaInyeccionSQL(request("listado"))

function arregla(texto)
	texto_ = replace(texto,chr(10),"<br />")
	texto_ = replace(texto_,"FUENTE","<br /><br /><b>FUENTE</b>")
	arregla = texto_
end function

	comentarios = ""
	if listado = "prohibidas" then comentarios = generaComentarios("dn_risc_sustancias_"&listado,"Comentario","comentario_prohibida")
	if listado = "restringidas" then comentarios = generaComentarios("dn_risc_sustancias_"&listado,"Comentario","comentario_restringida")
	if listado = "biocidas_prohibidas" then comentarios = generaComentarios("spl_risc_sustancias_"&listado,"Fuente,Fecha l&iacute;mite,Usos no autorizados","fuente,fecha_limite,usos")
	if listado = "biocidas_autorizadas" then comentarios = generaComentarios("spl_risc_sustancias_"&listado,"Fuente,Pureza m&iacute;nima,Condiciones,Usos autorizados","fuente,pureza_minima,condiciones,usos")
	if listado = "pesticidas_prohibidas" then comentarios = generaComentarios("spl_risc_sustancias_"&listado,"Fuente,Exenciones","fuente,exenciones")
	if listado = "pesticidas_autorizadas" then comentarios = generaComentarios("spl_risc_sustancias_"&listado,"Fuente,Plazo renovaci&oacute;n,Pureza m&iacute;nima,Usos autorizados","fuente,plazo_renovacion,pureza_minima,usos")

	if listado = "prohibidas_embarazadas" then comentarios = generaComentarios("spl_risc_sustancias_"&listado,"Comentario","comentario_prohibida")
	if listado = "prohibidas_lactantes" then comentarios = generaComentarios("spl_risc_sustancias_"&listado,"Comentario","comentario_prohibida")
	if listado = "candidatas_reach" then comentarios = generaComentarios("spl_risc_sustancias_"&listado,"",,"")
	if listado = "autorizacion_reach" then comentarios = generaComentarios("spl_risc_sustancias_"&listado,"","")
	if listado = "" then comentarios = generaComentarios("spl_risc_sustancias_"&listado,"","comentario_restringida")
	if listado = "corap" then comentarios = generaComentarios_sin_grupo("ist_risc_sustancias_" & listado, "Año, Estado miembro, Motivos iniciales de preocupación", "año,estado_miembro,motivos_preocupacion")

	objRecordset.close
	set objRecordset=nothing
	cerrarconexion


%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>M&aacute;s informaci&oacute;n</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="Title" content="Risctox" />
<meta name="Author" content="SPL Sistemas de Informaci?n - www.spl-ssi.com" />
<meta name="description" content="Informaci?n, formaci?n y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Subject" content="Informaci?n, formaci?n y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Keywords" content="Informaci?n, formaci?n y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Language" content="Spanish" />
<meta name="Revisit" content="15 days" />
<meta name="Distribution" content="Global" />
<meta name="Robots" content="All" />
<link rel="stylesheet" type="text/css" href="estructura.css">
<link rel="stylesheet" type="text/css" href="dn_estilos.css">
<body>
<br />
<table class="tabla3" width="95%" align="center" height="100%" valign="middle" cellpadding="5">

<tr>
  <td class=texto colspan="2" style='text-align:justify'>
  	<%
		if tipo=1 then
			 response.write arregla(comentarios)
    else
       response.write arregla(comentarios)
    end if
	%>
  </td>
  </tr>
</table>
<br />
</body>
</html>
<%

' Obtiene de una lista en BD todos los comentarios, con los campos y nombres de campos en dos argumentos,
' y monta el string de comentarios
function generaComentarios(lista, nombresCampos,campos)
	dim c
	comentarios = ""

	' Convertimos las cadenas de campos y nombres en arrays
	nombresCamposArray = split(nombresCampos,",")
	camposArray = split(campos,",")


	sql = "(SELECT "&campos&" FROM "&lista&" AS lst WHERE lst.id_sustancia="&id&")"
	sql = sql & " UNION ALL "
	sql = sql & "(SELECT  "
	' A los campos dentro de la tabla grupo se les a?ade el prefijo asoc_lista_seleccionada
	for i = 0 to UBound(camposArray)
		c = camposArray(i)
		if c<>"" then
			sql = sql & "asoc_"&listado&"_"&c&","
		end if
	next
	sql = left(sql,	len(sql)-1)
	sql = sql & " FROM dn_risc_grupos AS grp, dn_risc_sustancias_por_grupos as sg WHERE sg.id_sustancia="&id&" AND sg.id_grupo=grp.id AND grp.asoc_"&listado&"=1)"

'comentarios = sql
	set objRecordset = Server.CreateObject ("ADODB.Recordset")
	objRecordset.Open sql,objConnection,adOpenKeyset


	do while not objRecordset.eof
		if comentarios <> "" then comentarios = comentarios & "---------------------------------------------------------------------------<br>"
		for i = 0 to UBound(camposArray)
			c = camposArray(i)
			if c<>"" then
				valor = objRecordset(c)
				if (len(valor)>0) then
					comentarios = comentarios & "<b>" &  nombresCamposArray(i) & ":</b> " & valor & "<br>"
				end if
			end if
		next

		objRecordset.movenext
	loop
	generaComentarios = comentarios
end function

function generaComentarios_sin_grupo(lista, nombresCampos, campos)
	
	dim c, str_out, nombresCamposArray, camposArray
	str_out = ""
	c = ""

	nombresCamposArray = split(nombresCampos,",")
	camposArray = split(campos,",")

	sql = "(SELECT " & campos & " FROM " & lista & " AS lst WHERE lst.id_sustancia=" & id & ")"
	
	set objRecordset = Server.CreateObject ("ADODB.Recordset")
	
	objRecordset.Open sql, objConnection, adOpenKeyset
	
	do while not objRecordset.eof
		
		if str_out <> "" then str_out = str_out & "---------------------------------------------------------------------------<br>"
		for i = 0 to ubound( camposArray )
			c = camposArray(i)
			if c<>"" then
				valor = objRecordset(camposArray(i))
				if (len(valor)>0) then
					str_out = str_out & "<b>" &  nombresCamposArray(i) & ":</b> " & valor & "<br>"
				end if
			end if
		next

		objRecordset.movenext
	loop
	
	objRecordset.close
	set objRecordset=nothing
	cerrarconexion
	
	generaComentarios_sin_grupo = str_out
	
end function

%>