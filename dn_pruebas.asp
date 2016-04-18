<!--#include file="dn_funciones_comunes.asp"-->
<!--#include file="adovbs.inc"--><!--#include file="dn_conexion.asp"-->

<%
log_this "probando el logger"
%>

<table border="1">
<tr>
  <th>Script</th>
  <th>Fecha</th>
  <th>Descripción</th>
</tr>
<%
sql="SELECT TOP 100 * FROM dn_log ORDER BY id DESC"
set objRst=objConnection2.execute(sql)

do while(not objRst.eof)
%>
<tr>
  <td><%= objRst("script") %></td>
  <td><%= objRst("fecha") %></td>
  <td><%= objRst("descripcion") %></td>
</tr>
<%  
  objRst.movenext
loop
%>
</table>

<%
cerrarconexion
%>
