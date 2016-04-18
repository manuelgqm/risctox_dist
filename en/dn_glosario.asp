<!--#include file="../EliminaInyeccionSQL.asp"-->
<!--#include file="../dn_conexion.asp"-->


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>RISCTOX: Glossary</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="Title" content="Risctox" />
<meta name="Author" content="SPL Sistemas de Información - www.spl-ssi.com" />
<meta name="description" content="RISCTOX: Toxic and hazardous substances database" />
<meta name="Subject" content="RISCTOX: Toxic and hazardous substances database" />
<meta name="Keywords" content="RISCTOX: Toxic and hazardous substances database" />
<meta name="Language" content="English" />
<meta name="Revisit" content="15 days" />
<meta name="Distribution" content="Global" />
<meta name="Robots" content="All" />

<link rel="stylesheet" type="text/css" href="../estructura.css">
<link rel="stylesheet" type="text/css" href="css/en.css">
<body>
&nbsp;
<table class="tabla3" width="90%" align="center" height="100%" valign="middle" cellpadding="5">

<%
tabla = "dn_risc_"&EliminaInyeccionSQL(request("tabla"))
id = EliminaInyeccionSQL(request("id"))

sql = "SELECT nombre_ing, descripcion_ing FROM "&tabla&" WHERE id="&id
'response.write "<br>"&sql
set objRst = objConnection2.execute(sql)
if (objRst.eof) then
	' No se encontr�
%>
  <tr>
    <td class=titulo3 align=right valign=top>?</td>
    <td class=texto align=left>Not found</td>
  </tr>
<%
else
	' Se encontr�
%>

  <tr>
    <td class=titulo3 align=right valign=top width="20%"><%=objRst("nombre_ing")%></td>
    <td class=texto align=left><%=nl2br(objRst("descripcion_ing"))%></td>
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






