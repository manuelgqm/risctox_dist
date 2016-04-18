<!--#include file="EliminaInyeccionSQL.asp"-->
<%
seccion = 1
if (Instr(request.servervariables("PATH_INFO"),"/evalua/") > 0 or idpagina=1175 or idpagina=961) then seccion = 3
if (Instr(request.servervariables("PATH_INFO"),"/alternativas/") > 0 or idpagina=576 or idpagina=1174 ) then seccion = 2

	if 1=1 then
			idpagina=EliminaInyeccionSQL(request("idpagina"))
%>

			<div id="encabezado_nuevo_risctox<%=seccion%>">
			</div>
            <div class="textsubmenu" id="submenusup3">
            	<table cellspacing=1 cellpadding=1 border=0 style='padding:3px' width="100%">
            	<tr>
            	<%
                if (seccion=3) then 'evalua
					response.write "<td style='padding:5px'><a href='../index.asp'>bbdd risctox</a></td>"
					response.write "<td>|</td>"
					response.write "<td style='padding:5px'><a href='/alternativas/'>bbdd alternativas</a></td>"
					response.write "<td>|</td>"
					response.write "<td bgcolor='#FFFFFF' style='padding:5px'>eval&uacute;a lo que usas</td>"
				elseif (seccion=2) then 'alternativas
					response.write "<td style='padding:5px'><a href='../index.asp'>bbdd risctox</a></td>"
					response.write "<td>|</td>"
					response.write "<td bgcolor='#FFFFFF' style='padding:5px'>bbdd alternativas</td>"
					response.write "<td >|</td>"
					response.write "<td style='padding:5px'><a href='/evalua/'>eval&uacute;a lo que usas</a></td>"
				else 'risctox
					response.write "<td style='padding:5px' bgcolor='#FFFFFF'>bbdd risctox</td>"
					response.write "<td>|</td>"
					response.write "<td style='padding:5px'><a href='/alternativas/'>bbdd alternativas</a></td>"
					response.write "<td>|</td>"
					response.write "<td style='padding:5px'><a href='/evalua/'>eval&uacute;a lo que usas</a></td>"
					'response.write "<td style='padding:5px'><a href='http://www.istas.net/risctox/evalua/'>eval&uacute;a lo que usas</a></td>"
				end if
                'response.write

				%>
<!--				<td width="30%">-->
<%

'dir = request.servervariables("HTTP_HOST") & request.servervariables("URL") & "?" & request.servervariables("QUERY_STRING")
'response.write(dir)
'for each x in Request.ServerVariables
'  response.write(x & " = " & request.servervariables(x)& "<br />")
'next
%>
<!--</td>-->
				<td ><b>es<b>|<a href="en/">en</a></td>
				</tr>
                </table>
            </div>

<% else %>
<div id="encabezado_nuevo3">
	<table width="100%" cellpadding=0 border=0>
		<tr>
			<td width="215" height="78" onclick="location.href='index.asp?idpagina=548'" style="cursor:hand">&nbsp;</td>
			<td width="142" height="78" onclick="location.href='index.asp?idpagina=548'" style="cursor:hand">&nbsp;</td>
			<td width="166" height="78" onclick="location.href='index.asp?idpagina=549'" style="cursor:hand">&nbsp;</td>
			<td width="160" height="78" onclick="location.href='index.asp?idpagina=550'" style="cursor:hand">&nbsp;</td>
			<td width="25"  height="78" align="center">
				<a href="mailto:salvira@istas.ccoo.es?subject=Contacto ECOinformas"><img src="imagenes/ico_contacto.gif" border="0" alt="Contacto"></a><br>
				<a href="busqueda.asp"><img src="imagenes/ico_busqueda.gif" border="0" alt="busqueda"></a><br>
				<a href="index.asp?idpagina=560"><img src="imagenes/ico_ayuda.gif" border="0" alt="ayuda"></a>
			</td>
		</tr>
	</table>
</div>

			<div id="menusup3">
			<table border="0" cellspacing="5" cellpadding="0">
            			<tr class="textmenusup">
<%              				sql = "SELECT titulo,idpagina,numeracion FROM WEBISTAS_PAGINAS WHERE visible=1 AND numeracion LIKE 'AIC%' AND len(numeracion)=4 ORDER BY numeracion"
			   	   	   	set objRecordset = Server.CreateObject ("ADODB.Recordset")
		   	   	           	set objRecordset = OBJConnection.Execute(sql)
		   	   	           	do while not objRecordset.eof
              						response.write "<td class=textmenusup>"
							if mid(numeracion,1,4)=mid(objRecordset("numeracion"),1,4) then
								response.write lcase(objRecordset("titulo"))
              						else
              							response.write "<a href=index.asp?idpagina="&objRecordset("idpagina")&" style='text-decoration:none'>"&lcase(objRecordset("titulo"))&"</a>"
              						end if
              						response.write "</td><td class=textmenusup>|</td>"
							objrecordset.movenext
 						loop %>
              			</tr>
          		</table>
			</div>

<% end if %>
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
