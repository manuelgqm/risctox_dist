<!--#include file="web_inicio.asp"-->

<%
	'Recogemos todos los parámetros
 	idley = EliminaInyeccionSQL(request("idley"))
	
 	titulo = limpia(cstr(EliminaInyeccionSQL(request("titulo"))))
	subtitulo = limpia(cstr(EliminaInyeccionSQL(request("subtitulo"))))
	texto = limpia(cstr(EliminaInyeccionSQL(request("texto"))))
	enlaces = EliminaInyeccionSQL(request("enlaces"))
	If enlaces = "" Then enlaces = "0"
	Tipo = EliminaInyeccionSQL(request("Tipo"))
	If Tipo = "" Then Tipo = "0"
	subTipo = EliminaInyeccionSQL(request("subTipo"))
	If subTipo = "" Then subTipo = "0"
	ambito = EliminaInyeccionSQL(request("ambito"))
	If ambito = "" Then ambito = "0"
	idautonomia = EliminaInyeccionSQL(request("idautonomia"))
	If idautonomia = "" Then idautonomia = "0"
	If CStr(ambito) = "2" Then subTipo = "0"
	es_LegislacionOnline = EliminaInyeccionSQL(request("es_LegislacionOnline"))
	If es_LegislacionOnline = "" Then CStr(es_LegislacionOnline) = "0"

	if CStr(idley) <> "" And CStr(idley) <> "0" then 
		'update ley
		sql = "UPDATE ECO06_VIG_LEG_LEYES SET titulo='" & titulo & "', subtitulo='" & subtitulo & "',texto='" & texto & "', idenlace=" & enlaces & ", idtipo_ambiental=" & Tipo & ", idsubtipo_ambiental=" & subTipo & ", ambito=" & ambito & ", idautonomia=" & idautonomia & ", es_LegislacionOnline = " & es_LegislacionOnline & " WHERE idley="&idley
		Set objRecordset = Server.CreateObject ("ADODB.Recordset")
		Set objRecordset = OBJConnection.Execute(sql)
	
	Else
		'Insert nueva ley
		sql = "INSERT INTO ECO06_VIG_LEG_LEYES (titulo, subtitulo,texto, idenlace, idtipo_ambiental, idsubtipo_ambiental, ambito, idautonomia, es_LegislacionOnline) " 
		sql = sql & "VALUES ('" & titulo & "', '" & subtitulo & "','" & texto & "', " & enlaces & ", " & Tipo & ", " & subTipo & ", " & ambito & ", " & idautonomia & ", " & es_LegislacionOnline & " ) "
		Set objRecordset = Server.CreateObject ("ADODB.Recordset")
		Set objRecordset = OBJConnection.Execute(sql)

		sql2 = "SELECT MAX(idley) as maxid FROM ECO06_VIG_LEG_LEYES"        
		Set objRecordset = Server.CreateObject ("ADODB.Recordset")
		Set objRecordset = OBJConnection.Execute(sql2)
		idley = objrecordset("maxid")

	End if
	
	function limpia(texto)
		texto = replace(texto,"'","&#39;")
		limpia = texto
	end function

%>

<html>

<head>
<title>Modificar ley</title>

</head>
<link rel="stylesheet" type="text/css" href="panelcontrol.css">
<body class="cue_fondo" topmargin="20" leftmargin="20" onLoad="javascipt:vamos();">
<p class="negro">Guardando ley <%=idley%></p>
</body>
</html>

<script LANGUAGE="JScript">
function vamos()
{
	//parent.frames.izquierda.location.reload();
	location.href="vig_leg_editar.asp?idley=<%=idley%>";
}
</script>

<% 'end if %>