<!--#include file="dn_fun_comunes.asp"-->
<!--#include file="dn_fun_texto.asp"-->
<!--#include file="dn_auten.inc"-->
<!--#include file="adovbs.inc"-->
<!--#include file="dn_conexion.asp"-->

<%
asociar=request("asociar")
idcheck=request.form("idcheck")
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

function recoge_idcheck() 
{
	var frm = window.opener.document.forms["myform"]; 
	idcheck="";
	for (i=0; i<frm.elements.length; i++)
	{
		if ((frm.elements[i].name=='idcheck') && (frm.elements[i].checked))
		{
			if (idcheck == "")
			{
				idcheck=frm.elements[i].value;
			}
			else
			{
				idcheck=idcheck + ", " + frm.elements[i].value;
			}
		}
	}

	if (idcheck == "")
	{
		alert("No se seleccionó ninguna sustancia; no se realizará ninguna asociación.")
	}
	else
	{
		var frm2 = document.forms["myform"];
		frm2.idcheck.value=idcheck;
	}
}
</script>
<script language="JavaScript" src="gen_validatorv2.js" type="text/javascript"></script>
</head>

<body onload='recoge_idcheck()'>
<%flashMsgShow()%>
<div id="box2" class="centcontenido">
<form name="myform" action="dn_asociar2.asp?asociar=<%=asociar%>" method="post" >
<input type="hidden" name="idcheck" value="<%=idcheck%>" />

<%
	select case asociar
	
		'ASOCIAR SUSTANCIAS A:
		
		case "grupo":
		
			sql3="select * from dn_risc_grupos order by nombre"
			set objRst3=objconn1.execute(sql3)
			do while not objRst3.eof 
				opciones=opciones& "<option value='" &objRst3("id")& "'>" &corta(objRst3("nombre"),130,"puntossuspensivos")& "</option>" & vbspace
			objRst3.movenext
			loop
			
			objRst3.close
			set objRst3=nothing
			
			asociartit="Seleccione el GRUPO"
			
		case "enfermedad", "enfermedad_gr":
		
			sql3="select * from dn_risc_enfermedades order by nombre"
			set objRst3=objconn1.execute(sql3)
			do while not objRst3.eof 
				opciones=opciones& "<option value='" &objRst3("id")& "'>" &acortarCadena(objRst3("nombre"),140," (...) ")& "</option>"
			objRst3.movenext
			loop
			
			objRst3.close
			set objRst3=nothing
			
			asociartit="Seleccione ENFERMEDAD"
			
		case "compania":
		
			sql3="select * from dn_risc_companias order by nombre"
			set objRst3=objconn1.execute(sql3)
			do while not objRst3.eof 
				opciones=opciones& "<option value='" &objRst3("id")& "'>" &objRst3("nombre")& "</option>"
			objRst3.movenext
			loop
			
			objRst3.close
			set objRst3=nothing
			
			asociartit="Seleccione COMPAÑÍA"

		case "sector":
		
			sql3="select * from dn_alter_sectores order by numero_cnae"
			set objRst3=objconn1.execute(sql3)
			do while not objRst3.eof 
				opciones=opciones& "<option value='" &objRst3("id")& "'>" & objRst3("numero_cnae") & " - "& acortarCadena(objRst3("nombre"),140," (...) ") &"</option>\n"
			  objRst3.movenext
			loop
			
			objRst3.close
			set objRst3=nothing
			
			asociartit="Seleccione SECTOR"
			
		case "uso", "uso_gr":
		
			sql3="select * from dn_risc_usos order by nombre"
			set objRst3=objconn1.execute(sql3)
			do while not objRst3.eof 
				opciones=opciones& "<option value='" &objRst3("id")& "'>" &objRst3("nombre")& "</option>"
			objRst3.movenext
			loop
			
			objRst3.close
			set objRst3=nothing
			
			asociartit="Seleccione USO"
			
			adicional="<br /><br /><input name='toxico' value='1' type='checkbox' /> Tóxico"
			
		'ASOCIAR FICHEROS DE ALTERNATIVAS A: (pide num_alternativa)
		
		case "fich_sustancia", "fich_grupo", "fich_sector", "fich_proceso", "fich_uso", "fich_residuo":
			
			asociartit="Fichero de alternativas"
				
	end select
cerrarconexion
%>

  
<fieldset><legend><strong><%=asociartit%></strong></legend>

<%
	select case asociar
	
		case "fich_sustancia", "fich_grupo", "fich_sector", "fich_proceso", "fich_uso", "fich_residuo": 'ficheros a
		
%>
		Número de alternativa: <input type="text" name="num_alternativa" />
<%
		case else: 'sustancias a....
			if opciones="" then
				response.write "No hay elementos a los que asociar las sustancias"
			else
%>
<select name="id"><%=opciones%></select>
<%=adicional%>

<%
		end if
	end select
%>

</fieldset>
<p><input type="submit" value="Enviar" class="centcontenido"  /></p>
</form>

</div>
</body>
</html>
