<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	'OBJConnection.Open "driver={sql server};server=osiris.servidoresdns.net;database=qc507;UID=qc507;PWD=sql"
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"

	
	'----- Si es restringida y no estás identificado no puedes entrar
	'if session("id_ecogente")="" then response.redirect "acceso.asp?idpagina="&idpagina
	'---- ATENCIÓN: ponerlo cuando publiquemos en abierto
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ECOinformas: Base de datos de sustancias tóxicas y peligrosas RISCTOX</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="Title" content="ECOinformas" />
<meta name="Author" content="XiP multimèdia" />
<meta name="description" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Subject" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Keywords" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Language" content="Spanish" />
<meta name="Revisit" content="15 days" />
<meta name="Distribution" content="Global" />
<meta name="Robots" content="All" />
<link rel="stylesheet" type="text/css" href="estructura.css"  />
<SCRIPT LANGUAGE="JavaScript">
<!--
function imprimir() 
{

		if  (confirm('¿Imprimir este documento?')) { print(); close();}
}

// -->
</SCRIPT>
</head>
<body onload="imprimir()">
<p class="titulo3" align="center">
<b><%=request("titulo")%></b>
<% if request("buscar")<>"" then %>
<br>(con el término&nbsp;<%=request("buscar")%>)
<% end if %>
</p>

<p class="texto" align="center">
<%=formatdatetime(now(),1)&" a las "&time()%>
</p>

<table class="tabla3" width="90%" align="center" border=0 cellpadding=4 cellspacing=0>
<tr>
<% if request("nombre1")<>"" then %>
<td class="subtitulo3" align="left"><b><%=ucase(request("nombre1"))%></b></td>
<% end if %>
<% if request("nombre2")<>"" then %>
<td class="subtitulo3" align="left"><b><%=ucase(request("nombre2"))%></b></td>
<% end if %>
<% if request("nombre3")<>"" then %>
<td class="subtitulo3" align="left"><b><%=ucase(request("nombre3"))%></b></td>
<% end if %>
<% if request("nombre4")<>"" then %>
<td class="subtitulo3" align="left"><b><%=ucase(request("nombre4"))%></b></td>
<% end if %>
<% if request("nombre5")<>"" then %>
<td class="subtitulo3" align="left"><b><%=ucase(request("nombre5"))%></b></td>
<% end if %>
<% if request("nombre6")<>"" then %>
<td class="subtitulo3" align="left"><b><%=ucase(request("nombre6"))%></b></td>
<% end if %>
<% if request("nombre7")<>"" then %>
<td class="subtitulo3" align="left"><b><%=ucase(request("nombre7"))%></b></td>
<% end if %>

</tr>
<%
	sql = request("sql")
	'response.redirect "orden.asp?orden="&sql
				
	set objRecordset = Server.CreateObject ("ADODB.Recordset")
	objRecordset.Open sql,objConnection,adOpenKeyset
	do while not objRecordset.eof %>
<tr>
<% if request("nombre1")="nombre/sustancia" then
	if objRecordset("nombre")="" or isnull(objRecordset("nombre"))then
		nombre_actual = objRecordset("sustancia")
	else
		nombre_actual = objRecordset("nombre")
	end if
   else
	nombre_actual = eval("objRecordset(" & """" & request("campo1") & """" & ")" )
	if nombre_actual="" or isnull(nombre_actual) then nombre_actual = "--"
   end if %>
   <td class="celda_imprimir" align="left"><%=nombre_actual%>&nbsp;</td>
<% if request("campo2")<>"" then %>
   <td class="celda_imprimir" align="left" nowrap><% if instr(objRecordset("cas"),"-XX-")=0 then response.write eval("objRecordset(" & """" & request("campo2") & """" & ")" ) %>&nbsp;</td>
<% end if %>   
<% if request("campo3")<>"" then %>
   <td class="celda_imprimir" align="left"><%=eval("objRecordset(" & """" & request("campo3") & """" & ")" ) %>&nbsp;</td>
<% end if %>
<% if request("campo4")<>"" then %>
   <td class="celda_imprimir" align="left"><%=eval("objRecordset(" & """" & request("campo4") & """" & ")" ) %>&nbsp;</td>
<% end if %>
<% if request("campo5")<>"" then %>
   <td class="celda_imprimir" align="left"><%=eval("objRecordset(" & """" & request("campo5") & """" & ")" ) %>&nbsp;</td>
<% end if %>		   
<% if request("campo6")<>"" then %>
   <td class="celda_imprimir" align="left"><%=eval("objRecordset(" & """" & request("campo6") & """" & ")" ) %>&nbsp;</td>
<% end if %>
<% if request("campo7")<>"" then %>
   <td class="celda_imprimir" align="left"><%=eval("objRecordset(" & """" & request("campo7") & """" & ")" ) %>&nbsp;</td>
<% end if %>
</tr>
<%	objrecordset.movenext
	loop %>
</table>
</body>
</html>