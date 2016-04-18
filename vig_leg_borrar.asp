<!--#include file="web_inicio.asp"-->
<%
	If cstr(request("idley")) <> "" And  cstr(request("idley")) <> "0" Then
	
		idley = clng(EliminaInyeccionSQL(request("idley")))

		'Set OBJConnection = Server.CreateObject("ADODB.Connection")
		'OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"

		orden = "DELETE FROM ECO06_VIG_LEG_LEYES WHERE idley="&idley
		Set DSql = OBJConnection.Execute(orden)
	End if
%>

<html>

<head><title>Borrar ley</title></head>


<body class="cue_fondo" onLoad="javascript:vamos();">
ok, ley borrada de la base de datos de vigilancia legislativa
</body>
</html>

<script LANGUAGE="JScript">
function vamos()
{
	//parent.frames.izquierda.location.reload();
	location.href="vig_leg_editar.asp?";

}
</script>
