<!--#include file="eco_conexion.asp"-->
<%

id = clng(EliminaInyeccionSQL(request("id")))
	
sql3 = "SELECT numeracion FROM WEBISTAS_PAGINAS WHERE idpagina="&id
Set objRecordset3 = Server.CreateObject ("ADODB.Recordset")
objRecordset3.Open sql3,OBJConnection,adOpenKeyset
numeracion = objrecordset3("numeracion")

sql = "SELECT idpagina FROM WEBISTAS_PAGINAS WHERE numeracion LIKE '"&numeracion&"%' AND len(numeracion)="&cstr(len(numeracion)+1)
Set objRecordset = Server.CreateObject ("ADODB.Recordset")
objRecordset.Open sql,OBJConnection,adOpenKeyset
hijos = objrecordset.recordcount

if hijos=0 then
	orden = "DELETE FROM WEBISTAS_PAGINAS WHERE idpagina="&id
	Set DSql = Server.CreateObject("ADODB.Recordset")
	Set DSql = OBJConnection.Execute(orden)
end if

%>
<html>
<head><title>Borrar Página definitivamente</title>
<link rel="stylesheet" type="text/css" href="panelcontrol.css">
</head>
<body class="cue_fondo" topmargin="20" leftmargin="20">
<p class="cue_fuente">
<% if hijos=0 then %>
ok, página eliminada definitivamente
<script LANGUAGE="JScript">
 location.href='eco_editarpagina.asp?id=1';
 parent.frames.izquierda.location.reload();
</script>

<% else %>

Hay <%=hijos%> páginas incluidas en esta página
</p>
<script LANGUAGE="JScript">
 alert('No se puede borrar por contener otras páginas');
 location.href='eco_editarpagina.asp?id=<%=id%>';
</script>

<% end if %>
</body>
</html>
