<!--#include file="dn_restringida.asp"-->
<!--#include file="dn_conexion.asp"-->
<!--#include file="dn_funciones_comunes.asp"-->
<!--#include file="dn_funciones_texto.asp"-->

<%
id_request = request("id")
id_request = EliminaInyeccionSQL(id_request)
sql = "SELECT * FROM dn_risc_companias WHERE id="&id_request
set objRst=objConnection2.execute(sql)
if(objRst.eof) then
	nombre = "Compa��a no encontrada"
	direccion = "No se ha encontrado la compa��a indicada"
  fuente = ""
  productora = false
  distribuidora = false
else
	nombre = objRst("nombre")
	direccion = objRst("direccion")
  fuente = objRst("fuente")
  productora = objRst("productora")
  distribuidora = objRst("distribuidora")
end if

cerrarconexion
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Definici�n</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="Title" content="ECOinformas" />
<meta name="Author" content="XiP multim�dia" />
<meta name="description" content="Informaci�n, formaci�n y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Subject" content="Informaci�n, formaci�n y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Keywords" content="Informaci�n, formaci�n y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Language" content="Spanish" />
<meta name="Revisit" content="15 days" />
<meta name="Distribution" content="Global" />
<meta name="Robots" content="All" />

<link rel="stylesheet" type="text/css" href="estructura.css"  />
<body>
&nbsp;
<table class="tabla3" width="90%" align="center" height="100%" valign="middle" cellpadding="5">
<tr>
	<td class=titulo3 align=left valign=top><%=nombre%></td>
</tr>
<tr>
	<td class=texto align=left>
	<%
		' Al mostrar direccion, intentamos parsear URLs dividiendolo todo por espacios. Para cada palabra, si comienza por "http://", es un enlace
		array_direccion = split(direccion, " ")
		for i=0 to ubound(array_direccion)
			dire = array_direccion(i)
			if (left(dire, 7) = "http://") then
				response.write "<a href='"&dire&"' target='_blank'>"&dire&"</a> "
			else
				response.write dire & " "
			end if
		next
	 %>

    <%
    if (productora and distribuidora) then
    %>
      <br /><br />Compa��a productora y distribuidora.
    <%
    elseif (productora) then
    %>
      <br /><br />Compa��a productora.
    <%
    elseif (distribuidora) then
    %>
      <br /><br />Compa��a distribuidora.
    <%
    end if
    %>

    <%
    if (fuente <> "") then
    %>
    <br /><br /><strong>Fuente: </strong><%= fuente %>
    <%
    end if
    %>
	</td>
</tr>
</table>
&nbsp;
</body>
</html>

