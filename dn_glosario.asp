<!--#include file="dn_conexion.asp"-->
<!--#include file="adovbs.inc"-->
<!--#include file="dn_restringida.asp"-->
<!--#include file="dn_funciones_comunes.asp"-->
<!--#include file="dn_funciones_texto.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Definición</title>
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
<body>
&nbsp;
<table class="tabla3" width="90%" align="center" height="100%" valign="middle" cellpadding="5">

<%
tabla = "dn_risc_"&EliminaInyeccionSQL(request("tabla"))
id = EliminaInyeccionSQL(request("id"))

sql = "SELECT nombre, descripcion FROM "&tabla&" WHERE id="&id
'response.write "<br>"&sql
set objRst = objConnection2.execute(sql)
if (objRst.eof) then
	' No se encontró
%>
  <tr>
    <td class=titulo3 align=right valign=top>?</td>
    <td class=texto align=left>No encontrado en el glosario</td>
  </tr>
<%
else
	' Se encontró
%>

  <tr>
    <td class=titulo3 align=right valign=top width="20%"><%=objRst("nombre")%></td>
    <td class=texto align=left><%=nl2br(objRst("descripcion"))%></td>
  </tr>

<% 
end if

objRst.close()
set objRst=nothing
%>


<%
cerrarconexion
%>


</table>
&nbsp;
</body>
</html>






