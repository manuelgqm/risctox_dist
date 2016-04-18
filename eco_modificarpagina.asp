<!--#include file="eco_conexion3.asp"-->
<%
		
 	id = EliminaInyeccionSQL(request("id"))
	
	titulo = cstr(EliminaInyeccionSQL(request("titulo")))
	titulo_eng = cstr(EliminaInyeccionSQL(request("titulo_eng")))
 	pagina = cstr((request("contenido")))
 	pagina_eng = cstr((request("contenido_eng")))
 	titulo = limpia(titulo)
 	titulo_eng = limpia(titulo_eng)
 	pagina = limpia(pagina)
 	pagina_eng = limpia(pagina_eng)
 	tipo = EliminaInyeccionSQL(request("tipo"))
 	destino = EliminaInyeccionSQL(request("destino"))
 	if destino="" then destino=0
 	visible = EliminaInyeccionSQL(request("visible"))
	
	set sql = Server.CreateObject ("ADODB.Recordset")
	orden = "UPDATE WEBISTAS_PAGINAS SET titulo='"&titulo&"',pagina='"&pagina&"',autor="& session( "idgente" ) &",fecha_modificacion='"&date()&"',tipo="&tipo&",destino="&destino&",visible="&visible & ", titulo_eng='" & titulo_eng & "', pagina_eng='" & pagina_eng & "' WHERE idpagina="&id
	Set sql = OBJConnection.Execute(orden)  			
	
	function limpia(texto)
		texto = replace(texto,"'","&quot;")
		limpia = texto
	end function

%>

<html>

<head>
<title>Modificar página</title>
<script LANGUAGE="JScript">
function vamos()
{
	parent.frames.izquierda.indice_arriba.location.reload();
	location.href="eco_editarpagina.asp?id=<%=id%>";
}
</script>
</head>
<link rel="stylesheet" type="text/css" href="panelcontrol.css">
<body class="cue_fondo" topmargin="20" leftmargin="20" onLoad="javascript:vamos()">
<p class="negro">Guardando página <%=id%></p>
</body>
</html>