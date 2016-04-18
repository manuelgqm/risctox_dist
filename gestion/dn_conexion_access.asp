<%
dim objConnAccess

set objConnAccess = Server.CreateObject("ADODB.Connection")
'objConnAccess.open "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("estructuras\risctox.mdb")
objConnAccess.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath("estructuras\risctox.mdb")&";Persist Security Info=False;Jet OLEDB:Database ;"
%>


<%
sub cerrar_conexion_access
	objConnAccess.close
	Set objConnAccess=nothing
end sub
%>
