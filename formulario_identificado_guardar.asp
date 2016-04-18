<!--#include file="EliminaInyeccionSQL.asp"-->
<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	'OBJConnection.Open "driver={sql server};server=osiris.servidoresdns.net;database=qc507;UID=qc507;PWD=sql"
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"




FUNCTION unQuote(s)
  pos = Instr(s, "'")
  While pos > 0 
    s = Mid(s,1,pos) & "'" & Mid(s,pos+1)
    pos = InStr(pos+2, s, "'")
  Wend
  pos = Instr(s, """") 
  While pos > 0 
    s = Mid(s,1,pos-1) & "''" & Mid(s,pos+1)
    pos = InStr(pos+2, s, """")
  Wend
  unQuote = Trim(s)
END FUNCTION

FDT01 = EliminaInyeccionSQL(request("FDT01"))
'response.write "FDT01="&FDT01
if FDT01<>"" then FDT01=1 else FDT01=0
FDT02 = EliminaInyeccionSQL(request("FDT02"))
if FDT02<>"" then FDT02=1 else FDT02=0
FDT03 = EliminaInyeccionSQL(request("FDT03"))
if FDT03<>"" then FDT03=1 else FDT03=0
FDT04 = EliminaInyeccionSQL(request("FDT04"))
if FDT04<>"" then FDT04=1 else FDT04=0
FDT05 = EliminaInyeccionSQL(request("FDT05"))
if FDT05<>"" then FDT05=1 else FDT05=0
FDT06 = EliminaInyeccionSQL(request("FDT06"))
if FDT06<>"" then FDT06=1 else FDT06=0

orden = "UPDATE ECOINFORMAS_GENTE set FDT01='"&FDT01&"',FDT02='"&FDT02&"',FDT03='"&FDT03&"',FDT04='"&FDT04&"',FDT05='"&FDT05&"',FDT06='"&FDT06&"' WHERE idgente="&session("id_ecogente")
Set objRecordset = OBJConnection.Execute(orden)

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ECOinformas: usuario identificado</title>
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
<body>

<div id="contenedor">
	<div id="sombra_arriba"></div>
  	<div id="sombra_lateral">
		<div id="caja">
			<div id="encabezado_nuevo1">
			<table width="100%" cellpadding=0 border=0>
			<tr><td width="215" height="78" onclick="location.href='index.asp?idpagina=548'" style="cursor:hand">&nbsp;</td>
			    <td width="142" height="78" onclick="location.href='index.asp?idpagina=548'" style="cursor:hand">&nbsp;</td>
			    <td width="166" height="78" onclick="location.href='index.asp?idpagina=549'" style="cursor:hand">&nbsp;</td>
			    <td width="160" height="78" onclick="location.href='index.asp?idpagina=550'" style="cursor:hand">&nbsp;</td>
			    <td width="25"  height="78" align="center">
			    	<a href="mailto:salvira@istas.ccoo.es?subject=Contacto ECOinformas"><img src="imagenes/ico_contacto.gif" border="0" alt="Contacto"></a><br>
			    	<a href="busqueda.asp"><img src="imagenes/ico_busqueda.gif" border="0" alt="busqueda"></a><br>
			    	<a href="index.asp?idpagina=560"><img src="imagenes/ico_ayuda.gif" border="0" alt="ayuda"></a>
			    </td>
			</tr>
			</table>
			</div>
			<div id="menusup1">
			<table border="0" cellspacing="5" cellpadding="0">
            			<tr class="textmenusup"><td class=textmenusup>Formulario</td>
          		</table>
			</div>
			
			<% if session("id_ecogente")<>"" then %>
			<div class="textsubmenu" id="submenusup<% response.write (seccion) %>">
			<table width="100%"  border="0" cellspacing="4" cellpadding="0">
<%            				sql = "SELECT nombre,apellidos,sexo FROM ECOINFORMAS_GENTE WHERE idgente="&session("id_ecogente")
			   	   	set objRecordset = Server.CreateObject ("ADODB.Recordset")
		   	   	        set objRecordset = OBJConnection.Execute(sql)
		   	   	        usuario = objRecordset("nombre")&" "&objRecordset("apellidos")
		   	   	        usuario_sexo = "o"
		   	   	        if objRecordset("sexo")=75 then usuario_sexo = "a"
%>
            			<tr><td align="right">Usuari<%=usuario_sexo%> identificad<%=usuario_sexo%>:&nbsp;<%=usuario%></td></tr>
          		</table>
			</div>
       			<% end if %>
			
			<div id="texto">
			
				
				<div class="texto">
			
<br>&nbsp;				
<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td class="titulo" align="center" colspan=3>Solicitud de formación en ECOinformas</td></tr> 
<tr><td class="texto" align="justify" colspan=3>&nbsp;</td></tr>
<% if FDT01<>"0" or FDT02<>"0" or FDT03<>"0" or FDT04<>"0" or FDT05<>"0" or FDT06<>"0" then %>
<tr><td class="texto" align="justify" colspan=3>Se ha guardado tu petición para participar en las siguientes acciones formativas:</td></tr>
<tr><td>&nbsp;</td><td><strong>Código</strong></td><td><strong>Acción formativa</strong></td></tr>
<% if FDT01<>"0" then %>
<tr><td class="celda">&nbsp;</td><td class="celda">FDT01</td><td class="celda">Curso on-line sobre "Medio Ambiente, Salud y Desarrollo Sostenible"</td></tr>
<% end if %>
<% if FDT02<>"0" then %>
<tr><td class="celda">&nbsp;</td><td class="celda">FDT02</td><td class="celda">Curso on-line sobre "Medio Ambiente, Salud y Desarrollo Sostenible"</td></tr>
<% end if %>
<% if FDT03<>"0" then %>
<tr><td class="celda">&nbsp;</td><td class="celda">FDT03</td><td class="celda">Curso on-line sobre “Medio ambiente y actividades productivas: Prácticas sostenibles en agua, energía, residuos y emisiones"</td></tr>
<% end if %>
<% if FDT04<>"0" then %>
<tr><td class="celda">&nbsp;</td><td class="celda">FDT04</td><td class="celda">Curso on-line sobre “Medio ambiente y actividades productivas: Prácticas sostenibles en agua, energía, residuos y emisiones"</td></tr>
<% end if %>
<% if FDT05<>"0" then %>
<tr><td class="celda">&nbsp;</td><td class="celda">FDT05</td><td class="celda">Curso on-line sobre “Prevención del riesgo químico”</td></tr>
<% end if %>
<% if FDT06<>"0" then %>
<tr><td class="celda">&nbsp;</td><td class="celda">FDT06</td><td class="celda">Curso on-line sobre “Prevención del riesgo químico”</td></tr>
<% end if %>

<tr><td class="texto" align="center" colspan=3>&nbsp;</td></tr>
<tr><td class="texto" align="center" colspan=3>En breve nos comunicaremos contigo para informarte.</td></tr>

<% else %>
<tr><td class="texto" align="center" colspan=3>No has solicitado participar en ninguna acción formativa.</td></tr>
<% end if %>

</table>
				</div>
				<p>&nbsp;</p>
			</div>

			<map name="Map1" id="Map1">
            		<area shape="rect" coords="307,38,399,104" href="http://www.fundacion-biodiversidad.es" target="_blank" alt="Fundación Biodiversidad" />
            		<area shape="rect" coords="400,38,546,104" href="http://www.istas.ccoo.es" target="_blank" alt="Instituto Sindical de Trabajo, Ambiente y Salud" />
      			    <area shape="rect" coords="547,38,701,104" href="http://www.mtas.es/UAFSE/default.htm" target="_blank" alt="Fondo Social Europeo" />
      			</map>
			<img src="imagenes/pie1.jpg" width="708" border="0" usemap="#Map1">

    			</div>
    		</div>
		<div id="sombra_abajo"></div>
	</div>
</body>
</html>
