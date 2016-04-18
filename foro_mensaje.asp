<!--#include file="EliminaInyeccionSQL.asp"-->
<%
iden = session("id_ecogente")
tipo = clng(EliminaInyeccionSQL(Request("tipo")))
valor = clng(EliminaInyeccionSQL(request("id")))
fecha = now()

Const adOpenKeyset = 1
DIM objConnection
DIM objRecordset

Set OBJConnection = Server.CreateObject("ADODB.Connection")
OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"

orden = "INSERT into ECO_FOROS_LEIDOS (idmensaje,idgente,fecha) values ("&valor&","&iden&",'"&fecha&"')"
Set objRecordset = Server.CreateObject ("ADODB.Recordset")
Set Dorga = OBJConnection.Execute(orden)

SQLQuery = "SELECT ECO_FOROS.*,ECOINFORMAS_GENTE.nombre,ECOINFORMAS_GENTE.apellidos FROM ECO_FOROS LEFT JOIN ECOINFORMAS_GENTE ON ECO_FOROS.idgente=ECOINFORMAS_GENTE.idgente WHERE id="&valor
Set objRecordset = Server.CreateObject ("ADODB.Recordset")
set objRecordset = OBJConnection.Execute(SQLQuery)

texto = objRecordSet("texto")
salto = chr("13")
texto = replace (texto,salto,"<br>")
valor2 = objRecordset("tipo")
nombrecompleto = objRecordSet("nombre")&" "&objRecordSet("apellidos")

sqlquery2 = "SELECT * FROM ECO_FOROS WHERE tipo="&tipo&" AND nivel=0"
set objRecordset2 = OBJConnection.Execute(SQLQuery2)

sqlquery4 = "SELECT ECO_FOROS_LEIDOS.idmensaje, ECO_FOROS_LEIDOS.idgente FROM ECO_FOROS_LEIDOS GROUP BY ECO_FOROS_LEIDOS.idmensaje, ECO_FOROS_LEIDOS.idgente HAVING (ECO_FOROS_LEIDOS.idmensaje)="&valor
Set objRecordset4 = Server.CreateObject ("ADODB.Recordset")
objRecordSet4.Open sqlquery4,objConnection,1
leidopor = objRecordSet4.recordCount

sqlquery5 = "SELECT confirmado_web FROM ECOINFORMAS_GENTE WHERE idgente="&iden
set objRecordset5 = OBJConnection.Execute(SQLQuery5)
if objRecordset5("confirmado_web")=3 then permiso = 1

%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="Description" content="ISTAS">
<meta name="Author" content="XiP multimedia">
<title>Mensaje del Foro</title>
<link rel="stylesheet" type="text/css" href="estructura.css">
<SCRIPT LANGUAGE="JavaScript">
<!--
function enviar() {
	var asunto = document.formulario.asunto.value;
	var texto = document.formulario.texto.value;
	
	if ((asunto == "") || (texto == ""))
	{ alert("Escribe el asunto y el texto del mensaje antes de enviarlo"); }
	else
	{ if (texto.length >= 3999)
		{ alert('El mensaje es demasiado largo. Tiene '+texto.length+' caracteres y caben 4000 como máximo'); }
	  else	
		{ document.formulario.submit(); }
	}
}
// -->
</SCRIPT>
</head>

<body bgcolor="#FFFFFF" topmargin="10" leftmargin="10" class="cuerpo">

<table border="0" cellpadding="4" cellspacing="2" class="tabla">
  <tr>
    <td class="celda2" valign="top" align="right">De..&nbsp;</td>
    <td class="campo" valign="top"><%=nombrecompleto%></td>
  </tr>
  <tr>
    <td class="celda2" valign="top" align="right">Fecha..&nbsp;</td>
    <td class="campo"><%=objrecordset("fecha")%></td>
  </tr>
  <tr>
    <td class="celda2" valign="top" align="right">Asunto..&nbsp;</font></b></td>
    <td class="campo"><b><%=objrecordset("asunto")%></b>&nbsp;<font color="#555555">
    <% if permiso=1 then %>
    <br>leído por <%=leidopor%>&nbsp;personas&nbsp;[<a style="text-decoration:none;cursor:hand" onClick="window.open('foro_mensaje_leidopor.asp?actualiza=<%=time()%>&id=<%=valor%>','','scrollbars=yes,resizable=yes,width=400,height=300')">+</a>]</font>
    <% end if %>
    </td>
  </tr>
  <tr>
    <td class="celda2" valign="top" align="right">Mensaje..&nbsp;</td>
    <td class="campo" bgcolor="#FFFFFF" width="100%"><%=texto%></td>
  </tr>
<% if not isnull(objrecordset("adjunto")) and objrecordset("adjunto")<>"" then %>
  <tr>
    <td class="celda2" valign="top" align="right">Adjunto..&nbsp;</td>
    <td class="campo" bgcolor="#FFFFFF" width="100%"><a href="ftp/<%=objrecordset("adjunto")%>" target="_blank"><%=objrecordset("adjunto")%></a></td>
  </tr>
<% end if %>  
</table>

<p class="texto" align="center">
   <%'if clng(objrecordset("idgente"))=clng(iden) or permiso=1 or permiso=2 then
   if permiso=1 then%>
   <input type="button" value="BORRAR ESTE MENSAJE" onClick="location.href='foro_mensaje_borrar.asp?id=<%=valor%>';" class="boton">
   <%end if%>
   <input type="button" value="IMPRIMIR" class="boton" onClick="javascript:print();">&nbsp;<input type="button" value="CERRAR VENTANA" class="boton" onMouseUp="javascript:window.close()">
</p>
<p class="texto">Para publicar un mensaje relacionado con este:</p>
<form name="formulario" method="post" action="foro_mensaje_publicar.asp?tipo=<%=tipo%>&id=<%=valor%>">
<table border="0" cellpadding="4" cellspacing="2" width="100%" class="tabla">
   <tr>
    <td class="celda2" align="right">Asunto..&nbsp;</td>
    <td class="campo"><input type="text" name="asunto" size="60" value="Re: <%=objrecordset("asunto")%>" class="campo" maxlength="150"></td>
  </tr>
  <tr>
    <td class="celda2" valign="top" align="right">Mensaje..&nbsp;</td>
    <td class="campo"><textarea rows="10" name="texto" cols="62" class="campo"></textarea></td>
  </tr>
</table>
<p align="center">
<input type="button" value="PUBLICAR ANUNCIO" class="boton" onClick="enviar()">
</p>
</form>

</body>
</html>
