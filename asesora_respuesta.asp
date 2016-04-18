
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ECOinformas:Respuesta</title>
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


idconsulta = Request("idconsulta")
asesor     = Request("asesor")      ' Tipo de persona
if asesor="" then
 asesor=0
else
 asesor=cint(asesor)
end if

if asesor=2 then ' Es el coordinador. Hay que obligarle a que rellene el tema y el estado
 condi_extra = " (document.asesora_nuevo.estado_consulta.value=='') || (document.asesora_nuevo.tema_consulta.value=='')"
end if

if asesor=1 then ' Es un asesor. Hay que obligarle a que rellene la respuesta y el asunto
 condi_extra = " (document.asesora_nuevo.asunto.value=='') || (document.asesora_nuevo.pregunta.value=='')"
end if

if asesor=0 then ' Es un usuario, tendrá que rellenar el asunto, la pregunta 
 condi_extra = " (document.asesora_nuevo.estado_consulta.value=='') || (document.asesora_nuevo.asunto.value=='') || (document.asesora_nuevo.pregunta.value=='') "
end if 

estado_consulta_pri = Request("estado_consulta_pri")
tema_consulta_fil   = Request("tema_consulta_fil")
act_pag             = Request("act_pag")

if idconsulta<>"" then
 orden = "SELECT puntero,asunto,estado,tema_consulta,EV.desc1,E2.desc1 as estado_des from ECOINFORMAS_CONSULTAS EC "
 orden = orden & " LEFT JOIN ECOINFORMAS_VALORES EV ON EV.valor=EC.tema_consulta "
 orden = orden & " LEFT JOIN ECOINFORMAS_VALORES E2 ON E2.valor=EC.estado " 
 orden = orden & " WHERE idconsulta='"&idconsulta&"'"
 set objRecordset = objconnection.execute(orden)
 if not objRecordset.eof then
  asunto              = trim(objRecordset("asunto"))
  estado_consulta     = objRecordset("estado")
  estado_consulta_des = objRecordset("estado_des")  
  tema_consulta       = objRecordset("tema_consulta")
  tema_consulta_des   = trim(objRecordset("desc1"))
  puntero             = objRecordset("puntero")
  
 end if
end if

if puntero<>"" and not isnull(puntero) then
 orden = "SELECT idconsulta,fecha,texto,asunto,estado,tema_consulta,EV.desc1,E2.desc1 as estado_des from ECOINFORMAS_CONSULTAS EC "
 orden = orden & " LEFT JOIN ECOINFORMAS_VALORES EV ON EV.valor=EC.tema_consulta "
 orden = orden & " LEFT JOIN ECOINFORMAS_VALORES E2 ON E2.valor=EC.estado " 
 orden = orden & " WHERE idconsulta='"&puntero&"'"
 set objRecordset = objconnection.execute(orden)
 if not objRecordset.eof then
  estado_consulta     = objRecordset("estado")
  estado_consulta_des = objRecordset("estado_des")  
  tema_consulta       = objRecordset("tema_consulta")
  tema_consulta_des   = trim(objRecordset("desc1"))
  idconsulta          = objRecordset("idconsulta")
 end if
end if

Set UploadProgress = Server.CreateObject("Persits.UploadProgress")
PID = "PID=" & UploadProgress.CreateProgressID()
barref = "arbol_framebar.asp?to=10&" & PID


if asesor<>2 then
 onload = "document.asesora_nuevo.pregunta.focus();"
else
 onload = "document.asesora_nuevo.tema_consulta.focus();"
end if

%>

<body onload="<%=onload%>">

<form name="asesora_nuevo" action="asesora_respuesta_grabar.asp" METHOD="POST" ENCTYPE="multipart/form-data" OnSubmit="return ShowProgress();">

<% if asesor<>2 then %>
<p class="titulo2">Respuesta asesoramiento</p>
<p class="texto">Introduce la respuesta y pulsa ENVIAR.</p>
<% else %>
<p class="titulo2">Asignación asesoramiento</p>
<p class="texto">Selecciona el tema para la asignación.</p>
<% end if %>

<table cellspacing=0 border=0 cellpadding=0 class=tabla2 >

<tr><td width=100>&nbsp;</td><td width=400>&nbsp;</td></tr>

<tr><td align=right class=subtitulo2>Estado&nbsp;</td>
    <td align=left>
<%
 if asesor=1 or asesor=2 then 
  if asesor=2 and estado_consulta=151 then
   estado_consulta = 152
   'texto_coordinador = "Deja en blanco la respuesta para hacer la asignación"
  end if
 CALL valores2("015","estado_consulta",estado_consulta,"valor") 
else
%>
 <b><%=estado_consulta_des%></b>
 <input type=hidden name=estado_consulta value="<%=estado_consulta%>"> 
<% end if %>
</td></tr>
 
 
<tr><td align=right class=subtitulo2>Tema&nbsp;</td>
    <td align=left>
<% if asesor=2 then ' Sólo el coordinador puede asignar los temas... 
    CALL valores("010","tema_consulta",tema_consulta,"valor") 
   else 
    if tema_consulta<>"" then %>
     <b><%=tema_consulta_des%></b>
 <% else %>
     <b>Sin tema asignado</b> 
 <% end if %>
 
   <input type=hidden name=tema_consulta value="<%=tema_consulta%>">
   
<% end if %>
</td></tr>

<% if asesor<>2 then %>

<tr><td align=right class=subtitulo2>Asunto&nbsp;</td>
    <td align=left><input type=text size=70 class=campo name="asunto" value="RE: <%=asunto%>" maxlength=100></td></tr>

<tr><td align=right  class=subtitulo2>&nbsp;<b>Fichero adjunto</b></TD><td align=left><INPUT TYPE="FILE" NAME="myFile" SIZE="40"></TD></TR>
    
<tr><td align=right class=subtitulo2>Respuesta&nbsp;</td>
    <td align=left><textarea name="pregunta" class=campo rows=5 cols=70 ONKEYDOWN="return checkMaxLength(this, event,3000)" ONSELECT="storeSelection(this)"><%=pregunta%></textarea></td></tr>
    
<% else %>
 <tr><td align=right class=subtitulo2>
 <input type=hidden name="asunto" value="RE: <%=asunto%>">
 <input type=hidden name="myFile"> 
 <input type=hidden name=respuesta></td></tr>
<% end if %>
    
    
<tr><td >&nbsp;</td><td>&nbsp;&nbsp;</td></tr>   
<tr><td colspan=2>&nbsp;</td></tr>   
</table>

<input type="button" class="boton" value="Enviar" onclick="javascript:valida();">
</p>

<input type=hidden name=act_pag             value="<%=act_pag%>">
<input type=hidden name=estado_consulta_pri value="<%=estado_consulta_pri%>">
<input type=hidden name=tema_consulta_fil   value="<%=tema_consulta_fil%>">
<input type=hidden name=idconsulta          value="<%=idconsulta%>">

</form>
</body>
</html>

<script src="valida_textarea.js"></script>

<script>

function valida() {

if ( <%=condi_extra%> ) {
 alert ("Por favor, rellena los campos asunto, estado, tema y respuesta.");
 return (false);
 }
document.asesora_nuevo.submit();
}

function ShowProgress()
{
  strAppVersion = navigator.appVersion;
  if (document.asesora_nuevo.myFile.value != "")
  {
    if (strAppVersion.indexOf('MSIE') != -1 && strAppVersion.substr(strAppVersion.indexOf('MSIE')+5,1) > 4)
    {
      winstyle = "dialogWidth=385px; dialogHeight:190px; center:yes";
      window.showModelessDialog('<% = barref %>&b=IE',null,winstyle);
    }
    else
    {
      window.open('<% = barref %>&b=NN','','width=370,height=115', true);
    }
  }
  return true;
}

</script>

<%
FUNCTION valores (vgru,vname,vsele,vorde)
if vorde="" then
 vorde = "desc1"
end if

orden="Select * from ECOINFORMAS_VALORES WHERE grupo='"& vgru &"' ORDER BY " & vorde
Set DSQL = Server.CreateObject ("ADODB.Recordset")
dSQL.Open orden,objConnection,adOpenKeyset
seleccion = vsele
%><select name=<%=vname%> class="campo" >
  <option value="">- Selecciona de la lista -</option><%
if not(DSQL.bof and DSQL.eof) then
	dSQL.movefirst
	DO while not dSQL.eof
	    if not isnull(seleccion) then
		if cstr(seleccion)=cstr(DSQL("valor")) then
			sele ="selected"
		else
			sele=""
		end if	
	    end if	
	  %><option <%=sele%> value="<%=dSQL("valor")%>"><%=trim(dSQL("desc1"))%></option><%
    dSQL.movenext
	loop
end if
%></select><%
dSQL.close        

END FUNCTION


'
' Especial para el desplegable de ESTADOS.
'  Si es un asesor (1) no tiene que mostrar "sin asignar" ni "selecciona de la lista"
'
'
FUNCTION valores2 (vgru,vname,vsele,vorde)
if vorde="" then
 vorde = "desc1"
end if

orden="Select * from ECOINFORMAS_VALORES WHERE grupo='"& vgru &"' ORDER BY " & vorde
Set DSQL = Server.CreateObject ("ADODB.Recordset")
dSQL.Open orden,objConnection,adOpenKeyset
seleccion = vsele
%><select name=<%=vname%> class="campo" >
<% if asesor<>1 then %>
  <option value="">- Selecciona de la lista -</option>
<% end if %>

<%  
if not(DSQL.bof and DSQL.eof) then
	dSQL.movefirst
	DO while not dSQL.eof
	    if not isnull(seleccion) then
		if cstr(seleccion)=cstr(DSQL("valor")) then
			sele ="selected"
		else
			sele=""
		end if	
	    end if	
          if not (asesor=1 and cstr(dsql("valor"))="151") then 
	   %><option <%=sele%> value="<%=dSQL("valor")%>"><%=trim(dSQL("desc1"))%></option><%
	  end if 
    dSQL.movenext
	loop
end if
%></select><%
dSQL.close        

END FUNCTION

%>