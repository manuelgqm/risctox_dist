<%@ LANGUAGE = VBScript %>
<%Response.Expires=0%>
<%
iden	 = session("id_ecogente")
tipo 	 = clng(Request("tipo"))

Const adOpenKeyset = 1
DIM objConnection
DIM objRecordset

Set OBJConnection = Server.CreateObject("ADODB.Connection")
OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"

SQLQuery = "SELECT id FROM ECO_FOROS WHERE nivel=0 AND tipo="&tipo
Set objRecordset = Server.CreateObject ("ADODB.Recordset")
set objRecordset = OBJConnection.Execute(SQLQuery)

valor = objRecordset("id")

sqlquery2 = "SELECT * FROM ECO_FOROS WHERE nivel=0 and tipo="&tipo
set objRecordset2 = OBJConnection.Execute(SQLQuery2)
%>
<html>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="Description" content="intranet ISTAS">
<meta name="Author" content="XiP multimedia">
<head>
<title>MENSAJE en el FORO</title>
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
<p class="texto">Para publicar un nuevo mensaje, escribe el Título del que trate y el texto del mensaje:</p>
<form name="formulario" method="post" action="foro_mensaje_publicar.asp?tipo=<%=tipo%>&id=<%=valor%>">
<table border="0" cellpadding="2" cellspacing="2" width="100%" class="tabla">
   <tr>
     <td class="celda2" valign="top" align="right">Título..&nbsp;</td>
     <td class="campo"  align="left"><input type="text" name="asunto" size="60" class="campo" maxlength="150"></td>
   </tr>
   <tr>
     <td class="celda2" valign="top" align="right">Mensaje..&nbsp;</td>
     <td class="campo" align="left"><textarea rows="10" name="texto" cols="62" class="campo"></textarea></td>
   </tr>
</table>
<p align="center">
<input type="button" value="PUBLICAR MENSAJE" class="boton" onclick="enviar()">&nbsp;<input type="button" value="CERRAR VENTANA" class="boton" onmouseup="javascript:window.close()">
</p>
</form>

</body>
</html>
