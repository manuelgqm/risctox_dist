<!--#include file="../../EliminaInyeccionSQL.asp"-->
<%
seccion = 1
if (Instr(request.servervariables("PATH_INFO"),"/evalua/") > 0 or idpagina=1175 or idpagina=961) then seccion = 3
if (Instr(request.servervariables("PATH_INFO"),"/alternativas/") > 0 or idpagina=576 or idpagina=1174 ) then seccion = 2
idpagina=EliminaInyeccionSQL(request("idpagina"))
%>


<%
seccion = 3
if 1=0 and session("id_ecogente2")<>"" then %>
<div class="textsubmenu" id="submenusup<% response.write (seccion) %>">
	<table width="100%" border="0" cellspacing="4" cellpadding="0">
	<% sql = "SELECT nombre,apellidos,sexo FROM ECOINFORMAS_GENTE WHERE idgente="&session("id_ecogente")
		set objRecordset = Server.CreateObject ("ADODB.Recordset")
		set objRecordset = OBJConnection.Execute(sql)
		usuario = objRecordset("nombre")&" "&objRecordset("apellidos")
		usuario_sexo = "o"
		if objRecordset("sexo")=75 then usuario_sexo = "a"
		objRecordset.close
		set objRecordset=nothing
	%>
		<tr>
			<td align="right">Usuari<%=usuario_sexo%> identificad<%=usuario_sexo%>:&nbsp;<%=usuario%>&nbsp;</td></tr>
	</table>
			</div>

<% end if %>
<!-- modo: <%=session("modo") %>-->
