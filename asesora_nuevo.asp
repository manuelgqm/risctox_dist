
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ECOinformas:Pregunta</title>
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


Set UploadProgress = Server.CreateObject("Persits.UploadProgress")
PID = "PID=" & UploadProgress.CreateProgressID()
barref = "arbol_framebar.asp?to=10&" & PID

%>

<body onload="document.asesora_nuevo.asunto.focus();">

<form name="asesora_nuevo" action="asesora_nuevo_grabar.asp" METHOD="POST" ENCTYPE="multipart/form-data" OnSubmit="return ShowProgress();">

<p class="texto">
Introduce tu pregunta y pulsa ENVIAR.<br>
Recibirás un e-mail con la notificación de la respuesta.<br>
</p>

<br>

<table cellspacing=0 border=0 cellpadding=0 class=tabla2 >
<tr><td width=100>&nbsp;</td><td width=400>&nbsp;</td></tr>
<tr><td align=right class=subtitulo2><b>Asunto</b>&nbsp;</td>
    <td align=left><input type=text size=70 class=campo name="asunto" value="<%=asunto%>" maxlength=100></td></tr>    
<tr><td>&nbsp;</td></tr>

<tr><td align=right  class=subtitulo2>&nbsp;<b>Fichero adjunto</b></TD><td align=left><INPUT TYPE="FILE" NAME="myFile" SIZE="40"></TD></TR>
    
<tr><td align=right class=subtitulo2><b>Pregunta</b>&nbsp;</td>
    <td align=left><textarea name="pregunta" class=campo rows=7 cols=70 ONKEYDOWN="return checkMaxLength(this, event,3000)" ONSELECT="storeSelection(this)"><%=pregunta%></textarea></td></tr>
<tr><td>&nbsp;</td></tr>
</table>

<br>    
    
<input type="button" class="boton" value="Enviar" onclick="javascript:valida();">
</p>

</form>
</body>
</html>

<script src="valida_textarea.js"></script>

<script>

function valida() {

if ((document.asesora_nuevo.asunto.value=='') || (document.asesora_nuevo.pregunta.value=='')) {
 alert ("Por favor, rellena los campos de asunto y pregunta.");
 return (false);
}
 
ShowProgress(); 
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