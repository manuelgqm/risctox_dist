<%
'pagina de proceso que asocia bloque de sustancias a grupos/usos
'muestra resultado

'si no encuentra la sustancia, avisa
'else, si no existia la asociacion, la crea

'en el caso de usos, incluye ademas si es toxico
'si no existia la asociacion, la crea; si si existia, hace update de la toxicidad, por si ha cambiado
%>

<!--#include file="dn_fun_comunes.asp"-->
<!--#include file="dn_fun_texto.asp"-->
<!--#include file="dn_auten.inc"-->
<!--#include file="adovbs.inc"-->
<!--#include file="dn_conexion.asp"-->

<%'Permite auditar las acciones%>
<!--#include file="spl_auditoria.inc.asp"-->

<%
id=EliminaInyeccionSQL(request("id"))
asociar=EliminaInyeccionSQL(request("asociar"))

correspondena=EliminaInyeccionSQL(request("correspondena"))
sustancias=EliminaInyeccionSQL(request("sustancias"))
toxico=0
if request("toxico")=1 then toxico=1 'para usos

if id="" or asociar="" then
%>
	<script>window.close();</script>
<%
else

	arrsus = Split(sustancias, chr(13))

	'recorremos el array de sustancias
	'el procedure se encarga de a/ buscar el id de la sustancia (si es que esta exista) b/ si no existia ya la asociacion, la crea
	
	For i = 0 to Ubound(arrsus)
		
		if len(arrsus(i))>1000 or arrsus(i)="" then
			resultado=resultado& "<br />-La líneas de sustancia no deben pasar de 1000 caracteres ni estar vacias. Se ha omitido la línea " &i+1& " " &arrsus(i)
		else
			'cas: num_cas, ce: num_ce_einecs, num_ce_elincs
			midato=trim(arrsus(i))
			midato=replace(midato,chr(10),"")
			midato=replace(midato,chr(13),"")
			if midato<>"" then
				if correspondena="NOMBRE" then midato=h(midato)
				'Response.Write midato& "<br>"
				set cmdsus=Server.CreateObject("ADODB.Command")
				   With cmdsus
					.ActiveConnection=objconn1
					.CommandText="dn_asociarbloque"
					.CommandType=adCmdStoredProc
					
					.Parameters.Append  .CreateParameter("@asociar", advarchar, adParamInput, 15, asociar)
					.Parameters.Append  .CreateParameter("@id", adinteger, adParamInput, , id) 'id de grupo/
					.Parameters.Append  .CreateParameter("@correspondena", advarchar, adParamInput, 10, correspondena) 
					.Parameters.Append  .CreateParameter("@sustancia", advarchar, adParamInput, 1000, midato ) 
					.Parameters.Append  .CreateParameter("@toxico", adboolean, adParamInput, , toxico ) 
					.Parameters.Append  .CreateParameter("@miresultado", adInteger, adParamoutput)			
					
					.Execute,,adexecutenorecords
					miresultado=.Parameters("@miresultado")
					'response.write "<p>" &.Parameters("@miresultado")& "</p>"
					select case miresultado
						case 0: 'todo ha ido bien
						case 1: 'no se encontro la sustancia
								resultado=resultado& "<br />-No se encontró la sustancia en línea " &i+1& " " &arrsus(i)& ". Compruebe que está correctamente escrita."
			
					end select
				   End With 
				set cmdsus=nothing		
			end if
		end if
	Next

end if
' ** AUDITORIA
spl_accion = "asociar en bloque"
spl_entidad = asociar
spl_descripcion = "corresponden a: "&correspondena&" <br> sustancias: " & sustancias
' ** AUDITORIA **
call auditaYCierraConexion(spl_accion,spl_entidad,spl_descripcion) ' accion, entidad, descripcion	

'flashMsgCreate "Las sustancias se han asociado", "OK"
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
</script>
<script language="JavaScript" src="gen_validatorv2.js" type="text/javascript"></script>
</head>

<body>
<%flashMsgShow()%>
<div id="box2" class="centcontenido">
 
<fieldset>
<legend><strong>Asociar sustancias en bloque</strong>
</legend>
<p align="left">
<strong>Ha concluido el proceso de asociación en bloque.</strong><br />
<%=resultado%>
</p>
</fieldset>

<p><input type="submit" value="Cerrar ventana" class="centcontenido" onclick="window.close()" /></p>

</div>
</body>
</html>
