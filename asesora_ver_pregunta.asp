
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ECOinformas:Ver Respuesta</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="Title" content="ECOinformas" />
<meta name="Author" content="XiP multimèdia" />
<meta name="description" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Subject" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Keywords" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Language" content="Spanish" />
<meta name="Revisit" content="15 days" />
<meta name="Distribution" content="Global" />
<meta name="Robots" content="All" />
<link rel="stylesheet" type="text/css" href="estructura.css"  />
</head>


<%

'----- Si es restringida y no estás identificado no puedes entrar
'if session("id_ecogente")="" then response.redirect "acceso.asp?idpagina="&idpagina
'---- ATENCIÓN: ponerlo cuando publiquemos en abierto


Const adOpenKeyset = 1
DIM objConnection	
DIM objRecordset

Set OBJConnection = Server.CreateObject("ADODB.Connection")
'OBJConnection.Open "driver={sql server};server=osiris.servidoresdns.net;database=qc507;UID=qc507;PWD=sql"
OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"

asesor     = Request("asesor")
if asesor="" then
 asesor=0
else
 asesor=cint(asesor)
end if

idconsulta          = Request("idconsulta")
estado_consulta_pri = Request("estado_consulta_pri")
act_pag             = Request("act_pag")


if idconsulta<>"" then
 orden = "SELECT tipo_consulta,puntero,fecha,texto,asunto,estado,tema_consulta,EV.desc1,E2.desc1 as estado_des from ECOINFORMAS_CONSULTAS EC "
 orden = orden & " LEFT JOIN ECOINFORMAS_VALORES EV ON EV.valor=EC.tema_consulta "
 orden = orden & " LEFT JOIN ECOINFORMAS_VALORES E2 ON E2.valor=EC.estado " 
 orden = orden & " WHERE idconsulta='"&idconsulta&"'"
 set objRecordset = objconnection.execute(orden)
 if not objRecordset.eof then
  asunto = trim(objRecordset("asunto"))
  estado_consulta = objRecordset("estado")
  estado_consulta_des = objRecordset("estado_des")  
  tema_consulta = objRecordset("tema_consulta")
  tema_consulta_des = trim(objRecordset("desc1"))
  pregunta = trim(objRecordset("texto"))
  fecha = objRecordset("fecha")
  puntero = objRecordset("puntero")
  tipo_consulta = objRecordset("tipo_consulta")
  ' Para cambiar el texto de la PREGUNTA/RESPUESTA
  if tipo_consulta=157 then
   texto_pregunta = "Pregunta"
  else
   texto_pregunta = "Respuesta"  
  end if
 end if
end if

if puntero<>"" and not isnull(puntero) then
 orden = "SELECT fecha,texto,asunto,estado,tema_consulta,EV.desc1,E2.desc1 as estado_des from ECOINFORMAS_CONSULTAS EC "
 orden = orden & " LEFT JOIN ECOINFORMAS_VALORES EV ON EV.valor=EC.tema_consulta "
 orden = orden & " LEFT JOIN ECOINFORMAS_VALORES E2 ON E2.valor=EC.estado " 
 orden = orden & " WHERE idconsulta='"&puntero&"'"
 set objRecordset = objconnection.execute(orden)
 if not objRecordset.eof then
  estado_consulta = objRecordset("estado")
  estado_consulta_des = objRecordset("estado_des")  
  tema_consulta = objRecordset("tema_consulta")
  tema_consulta_des = trim(objRecordset("desc1"))
 end if
end if

' Estado SIN ASIGNAR ponerlo en rojo...
if estado_consulta=151 then 
   color_estado="<font color=DD0000>"
end if 


%>

<body>

<p class=titulo2>
FICHA ASESORAMIENTO
</p>

<table cellspacing=0 border=0 cellpadding=0 class=tabla2 >
<tr><td width=100>&nbsp;</td><td width=400>&nbsp;</td></tr>
<tr><td align=right class=subtitulo2><b>Estado:</b>&nbsp;</td><td align=left><%=color_estado%><%=estado_consulta_des%></td></tr>
<tr><td>&nbsp;</td></tr>
<tr><td align=right class=subtitulo2><b>Tema:</b>&nbsp;</td><td align=left><%=tema_consulta_des%><br></td></tr>
<tr><td>&nbsp;</td></tr>
<tr><td align=right class=subtitulo2><b>Fecha/Hora:</b>&nbsp;</td><td align=left><%=fecha%><br></td></tr>
<tr><td>&nbsp;</td></tr>
<tr><td align=right class=subtitulo2><b>Asunto:</b>&nbsp;</td><td align=left><%=asunto%><br></td></tr>
<tr><td>&nbsp;</td></tr>
<tr><td align=right class=subtitulo2><b><%=texto_pregunta%>:</b>&nbsp;</td><td align=left><%=pregunta%><br></td></tr>
<tr><td>&nbsp;</td></tr>
</table>
<br>
<input type="button" class="boton" value="Cerrar" onclick="window.close();">
<% if asesor=0 then %>
&nbsp;&nbsp;<input type="button" class="boton" value="Repreguntar" onclick="javascript:responder('<%=idconsulta%>');">
<% end if %>


</p>


</body>
</html>

<script>
<!--
function responder(id) {
 param = 'idconsulta='+id+'&asesor=<%=asesor%>&act_pag=<%=act_pag%>&estado_consulta_pri=<%=estado_consulta_pri%>&estado_consulta=<%=estado_consulta%>';
 //alert (param);
 location.href=('asesora_respuesta.asp?'+param);
}
-->
</script>
