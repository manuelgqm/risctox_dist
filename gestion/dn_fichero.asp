<!--#include file="dn_fun_comunes.asp"-->
<!--#include file="../dn_funciones_texto.asp" 'necesario para el corte de texto -->
<!--#include file="dn_auten.inc"-->

<%
'++++++CODIGO PARALELO A dn_sustancia; IMAGEN es el aqui el pdf, doc, ...), estructura_molecular se llama aqui archivo ++++++

'si nos pasan id de sustancia, consultamos datos
id=EliminaInyeccionSQL(request("id"))
%>
	<!--#include file="adovbs.inc"-->
	<!--#include file="dn_conexion.asp"-->
<%
if id<>"" then

	Set UploadProgress = Server.CreateObject("Persits.UploadProgress")
	PID = "PID=" & UploadProgress.CreateProgressID()
	barref = "arbol_framebar.asp?to=10&" & PID

	'DATOS GENERALES
	
	sql3="select * from dn_alter_ficheros where id=" &id
	'response.write sqls
	set objRst3=objconn1.execute(sql3)
	
	'id=objRst3("id")
	titulo=objRst3("titulo")
	num_alternativa=objRst3("num_alternativa")
	tema=objRst3("tema")
	resumen=objRst3("resumen")
	direccion_internet=objRst3("direccion_internet")
	archivo=objRst3("archivo")
	idioma=objRst3("idioma")
	autor=objRst3("autor")
	lugar=objRst3("lugar")
	publicacion=objRst3("publicacion")
	coleccion=objRst3("coleccion")
	descripcion_fisica=objRst3("descripcion_fisica")
	numero_normalizado=objRst3("numero_normalizado")
	notas=objRst3("notas")
	soporte=objRst3("soporte")
	fecha_actualizacion=objRst3("fecha_actualizacion")
	fecha_consulta=objRst3("fecha_consulta")
	criterios_aceptacion = objRst3("criterios_aceptacion")
	alternativas_minimizacion_residuos = objRst3("alternativas_minimizacion_residuos")

	objRst3.close
	set objRst3=nothing
	

end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Istas</title>
<link rel="stylesheet" type="text/css" href="dn_estilos.css">
<script type="text/javascript" src="niftycube.js"></script>
<script type="text/javascript">
window.onload=function(){
Nifty("div#box2","big"); 
}

function ShowProgress()
{
  strAppVersion = navigator.appVersion;
  if (document.lista.myFile.value != "")
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
<script language="JavaScript" src="gen_validatorv2.js" type="text/javascript"></script>
</head>

<body>
<%flashMsgShow()%>
<div id="box2" class="centcontenido">
<!--<form name="myform" action="dn_fichero2.asp?id=<%=id%>" method="post" enctype="multipart/form-data">-->
<form name="myform" ACTION="dn_fichero2.asp?<%=PID%>&id=<%=id%>" METHOD="POST" ENCTYPE="multipart/form-data" OnSubmit="return ShowProgress();">
<fieldset>
<legend><strong>Edición de Fichero</strong></legend>

<table align="center">

<tr>
<td valign="top" align="right">Título <span class="exp">(obligatorio, max 750 c.)</span></td>
<td colspan="2"><textarea name="titulo" cols="50"><%=titulo%></textarea></td>
</tr>
<tr>
<td valign="top" align="right">Número de alternativa </td>
<td colspan="2"><input name="num_alternativa" type="text" size="50" value="<%=num_alternativa%>" /></td>
</tr>
<tr>
<td valign="top" align="right">Tema</td>
<td colspan="2">
<script type="text/javascript">

	function cambioTema(lista,sCampo){
		var campo = document.getElementById(sCampo);
		var valor = lista.options[lista.selectedIndex].value;
		
		if (valor=='Nuevo tema') {
			campo.value = '';
            campo.style.display = 'block';
		}else{
			campo.value = valor;
            campo.style.display = 'none';
		}
	}
</script>
<select name="tema_slc" onchange="cambioTema(this,'tema')">
<option>SELECCIONE UN TEMA</option>
<option value="Nuevo tema">Nuevo tema</option>
<%
sqll="SELECT DISTINCT tema  FROM dn_alter_ficheros order by tema"
Set rstt=objConn1.Execute(sqll)
if not  rstt.eof then
	do while not rstt.eof
		if tema<>"" then
			if tema<>rstt("tema") then
				marcado=""
			else
				marcado="selected"
			end if
		end if
		response.write "<option value='" &rstt("tema")& "'" &marcado&">" &corta (rstt("tema"), 95, "puntossuspensivos")& "</option>" 
	rstt.movenext
	loop
end if
rstt.Close
Set rstt = Nothing
%>
</select>    
	<input name="tema" id="tema" type="text" size="50" style="display:none" maxlength="150" value="<%=tema%>" />
</td>
</tr>
<tr>
<td valign="top" align="right">Resumen</td>
<td colspan="2"><textarea name="resumen" cols="50"><%=resumen%></textarea></td>
</tr>
<tr>
<td valign="top" align="right">Criterios de aceptación</td>
<td colspan="2"><textarea name="criterios_aceptacion" cols="50"><%=criterios_aceptacion%></textarea></td>
</tr>
<tr>
<td valign="top" align="right">Alternativas de minimizacion de residuos</td>
<td colspan="2"><textarea name="alternativas_minimizacion_residuos" cols="50"><%=alternativas_minimizacion_residuos%></textarea></td>
</tr>


<tr>
<td valign="top" align="right">Direcci&oacute;n Internet<br>
  <small>Incluir http://<br>
(ej. <strong>http://</strong>www.google.com)</small> </td>
<td colspan="2"><input name="direccion_internet" type="text" size="50" value="<%=direccion_internet%>" /></td>
</tr>

<tr><td colspan="3">
<fieldset><legend><strong>Archivo</strong></legend>
<input type="hidden" name="ficheroantiguo" value="<%=archivo%>" /> <%'mandamos el nombre antiguo por si se da el caso de que hay que eliminarlo de disco%>
<%
if archivo="" then
%> 
Archivo 
<input name="archivo" type="file">  <input name="imagen" type="hidden" value="nueva"  /><br>  
<%
else
%>
<table>
<tr><td colspan="3">Archivo<br />
<a href="estructuras/<%=archivo%>"><%=archivo%></a>
</td>
<td align="left">
<input name="imagen" type="radio" value="mantener" checked="checked" /> mantener
<br>  <input name="imagen" type="radio" value="cambiar"  /> cambiar por: <input name="archivo" type="file">
<br>  <input name="imagen" type="radio" value="eliminar" /> eliminar
</td>
</tr>
</table>
<%
end if
%>
</fieldset>
</td></tr>

<tr>
<td valign="top" align="right">Idioma</td>
<td colspan="2"><input name="idioma" type="text" value="<%=idioma%>" size="50" maxlength="100" /></td>
</tr>
<tr>
<td valign="top" align="right">Autor</td>
<td colspan="2"><input name="autor" type="text" value="<%=autor%>" size="50" maxlength="150" /></td>
</tr>
<tr>
<td valign="top" align="right">Lugar</td>
<td colspan="2"><input name="lugar" type="text" value="<%=lugar%>" size="50" maxlength="500" /></td>
</tr>
<tr>
<td valign="top" align="right">Publicación</td>
<td colspan="2"><input name="publicacion" type="text" value="<%=publicacion%>" size="50" maxlength="150" /></td>
</tr>
<tr>
<td valign="top" align="right">Colecci&oacute;n</td>
<td colspan="2"><input name="coleccion" type="text" value="<%=coleccion%>" size="50" maxlength="150" /></td>
</tr>
<tr>
<td valign="top" align="right">Descripci&oacute;n f&iacute;sica </td>
<td colspan="2"><input name="descripcion_fisica" type="text" value="<%=descripcion_fisica%>" size="50" maxlength="100" /></td>
</tr>
<tr>
<td valign="top" align="right">N&uacute;mero normalizado </td>
<td colspan="2"><input name="numero_normalizado" type="text" value="<%=numero_normalizado%>" size="50" maxlength="150" /></td>
</tr>
<tr>
<td valign="top" align="right">Notas</td>
<td colspan="2"><input name="notas" type="text" value="<%=notas%>" size="50" maxlength="300" /></td>
</tr>
<tr>
<td valign="top" align="right">Soporte</td>
<td colspan="2"><input name="soporte" type="text" value="<%=soporte%>" size="50" maxlength="100" /></td>
</tr>
<tr>
<td valign="top" align="right">Fecha actualizaci&oacute;n </td>
<td><input name="fecha_actualizacion" type="text" value="<%=fecha_actualizacion%>" size="10" maxlength="10" /></td>
<td rowspan="2"><strong><small>(Formato de fecha dd/mm/aaaa)</small></strong></td>
</tr>
<tr>
<td valign="top" align="right">Fecha de consulta </td>
<td><input name="fecha_consulta" type="text" value="<%=fecha_consulta%>" size="10" maxlength="10" /></td>
</tr>



</table>
  
    </fieldset>
  <p><input type="submit" value="Enviar" class="centcontenido"  /></p>
  </form>
  
<script language="JavaScript" type="text/javascript">
var frmvalidator = new Validator("myform");
frmvalidator.addValidation("titulo","req","El Título es obligatorio.");
frmvalidator.addValidation("num_alternativa","req","El Número de alternativa es obligatorio.");
frmvalidator.addValidation("tema","maxlen=150");
frmvalidator.addValidation("resumen","maxlen=3000");
frmvalidator.addValidation("direccion_internet","maxlen=300");
frmvalidator.addValidation("idioma","maxlen=100");
frmvalidator.addValidation("autor","maxlen=150");
</script>

</div>
</body>
</html>
<%
	cerrarconexion
%>