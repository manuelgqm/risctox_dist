<%
response.ContentType="text/html; charset=iso-8859-1"
%>

<!--#include file="../dn_restringida.asp"-->
<!--#include file="../dn_conexion.asp"-->
<%
' Devuelve el listado de sustancias no toxicas asociadas al uso indicado
id_uso=EliminaInyeccionSQL(request("id_uso"))

sql = "SELECT dn_risc_sustancias_por_usos.id_sustancia AS id_sustancia, dn_risc_sustancias.nombre AS nombre FROM dn_risc_sustancias_por_usos INNER JOIN dn_risc_sustancias ON dn_risc_sustancias_por_usos.id_sustancia = dn_risc_sustancias.id WHERE dn_risc_sustancias_por_usos.toxico = 0 AND dn_risc_sustancias_por_usos.id_uso="&id_uso
set objRst = objConnection2.execute(sql)
if (not objRst.eof) then
%>
	<ul>
<%
	do while (not objRst.eof)
		id_sustancia = objRst("id_sustancia")
		nombre = objRst("nombre")
%>
			<li><a href="dn_risctox_ficha_sustancia.asp?id_sustancia=<%=id_sustancia%>"><%=nombre%></a></li>
<%
		objRst.movenext
	loop
%>
	</ul>
<%
else
	response.write "<p>No se han encontrado sustancias alternativas asociadas al uso.</p>"
end if
%>

<%
cerrarconexion
%>
