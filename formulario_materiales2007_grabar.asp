<!--#include file="EliminaInyeccionSQL.asp"-->
<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"




	dim campo(11,1)
	campo(1,0)="direccion" 
	campo(2,0)="localidad" 
	campo(3,0)="provincia" 
	campo(4,0)="cp" 
	campo(5,0)="SET01_2007"   
	campo(6,0)="AE01_2007"   
	campo(7,0)="AE02_2007"   
	campo(8,0)="AE03_2007"   
	campo(9,0)="AE04_2007"   
	campo(10,0)="direccion_materiales"   
	campo(11,0)="comentarios_2006" 
	
	campo(1,1)=unquote(EliminaInyeccionSQL(request(campo(1,0))))
	campo(2,1)=unquote(EliminaInyeccionSQL(request(campo(2,0))))
	campo(3,1)=EliminaInyeccionSQL(request(campo(3,0)))
	campo(4,1)=unquote(EliminaInyeccionSQL(request(campo(4,0))))
	campo(5,1)=EliminaInyeccionSQL(request(campo(5,0)))
	campo(6,1)=EliminaInyeccionSQL(request(campo(6,0)))
	campo(7,1)=EliminaInyeccionSQL(request(campo(7,0)))
	campo(8,1)=EliminaInyeccionSQL(request(campo(8,0)))
	campo(9,1)=EliminaInyeccionSQL(request(campo(9,0)))
	campo(10,1)=EliminaInyeccionSQL(request(campo(10,0)))
	campo(11,1)=unquote(mid(EliminaInyeccionSQL(request(campo(11,0))),1,1000))
	
orden = "UPDATE ECOINFORMAS_GENTE set "
for i=1 to 11
	orden = orden & campo(i,0) & "='" & campo(i,1) & "',"
next
orden = orden & "fec_hor_mod='" & now() & "',usu_mod=" & session("id_ecogente") & " WHERE idgente=" & session("id_ecogente")
Set objRecordset = OBJConnection.Execute(orden)
'response.write orden


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

FUNCTION booleano(v)
	if cstr(v)="1" then
		booleano = "sí"
	else
		booleano = "no"
	end if
END FUNCTION

FUNCTION valor(v)
	if cstr(v)<>"0" and cstr(v)<>"" then
		orden2 = "SELECT desc1 FROM ECOINFORMAS_VALORES WHERE valor="&v
		Set dSQL2 = Server.CreateObject ("ADODB.Recordset")
		dSQL2.Open orden2,objConnection,adOpenKeyset
		valor = dSQL2("desc1")
	else
		valor = "sin especificar"
	end if
END FUNCTION

FUNCTION valores (vgrupo,vname,tipo,vselected)

orden2 = "SELECT * FROM ECOINFORMAS_VALORES WHERE grupo='"&vgrupo&"' ORDER BY valor,desc1"
Set dSQL2 = Server.CreateObject ("ADODB.Recordset")
dSQL2.Open orden2,objConnection,adOpenKeyset
if tipo="1" then 				'--- tipo 1 = desplegable
%>	<select name=<%=vname%> class="campo">
	<option value="">- Selecciona de la lista -</option><%
	if not(DSQL2.bof and DSQL2.eof) then
		dSQL2.movefirst
		DO while not dSQL2.eof
		sele = ""
		if cstr(vselected)=cstr(dSQL2("valor")) then sele="SELECTED"
		  %><option <%=sele%> value="<%=dSQL2("valor")%>"><%=dSQL2("desc1")%></option><%
	        dSQL2.movenext
	        loop
	end if
%>	
	</select>
<%
else 						'--- tipo 2 = selección visible por radio
	if not(DSQL2.bof and DSQL2.eof) then
		dSQL2.movefirst
		DO while not dSQL2.eof
	%>
	<input type="radio" name="<%=vname%>" class="campo" value="<%=dSQL2("valor")%>">&nbsp;<%=dSQL2("desc1")%>&nbsp;&nbsp;
	<%      dSQL2.movenext
	        loop
	end if
end if
dSQL2.close        
END FUNCTION


%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ECOinformas: grabado formulario</title>
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
			<tr><td width="215" height="78" onClick="location.href='index.asp?idpagina=548'" style="cursor:hand">&nbsp;</td>
			    <td width="142" height="78" onClick="location.href='index.asp?idpagina=548'" style="cursor:hand">&nbsp;</td>
			    <td width="166" height="78" onClick="location.href='index.asp?idpagina=549'" style="cursor:hand">&nbsp;</td>
			    <td width="160" height="78" onClick="location.href='index.asp?idpagina=550'" style="cursor:hand">&nbsp;</td>
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
            			<tr class="textmenusup"><td class=textmenusup>Formulario de petición de materiales</td>
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
				<table width="90%" align="center">
				<tr><td class="texto">Tu información se ha grabado correctamente:</td></tr>
				</table>

<br>&nbsp;

<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td colspan="2"><table width="90%" align="center"><tr><td class="subtitulo" align="left" width="60%">Dirección de envío</td></tr> </table></td></tr>
<tr>
	<td class=texto align=right valign="middle">Direcci&oacute;n:</td>
	<td class=texto align=left><input type="text" name="direccion" size="50" maxlength="200" class="campo" value="<%=campo(1,1)%>"></td>
</tr>
<tr>
	<td class=texto align=right valign="middle">Localidad:</td>
	<td class=texto align=left><input type="text" name="localidad" size="50" maxlength="200" class="campo" value="<%=campo(2,1)%>"></td>
</tr>
<tr>
	<td class=texto align=right valign="middle">Provincia:</td>
	<td class=texto align=left><input type="text" name="provincia" size="50" maxlength="200" class="campo" value="<%=valor(EliminaInyeccionSQL(request(campo(3,0)))) %>"></td>
</tr>
<tr>
	<td class=texto align=right valign="middle">C&oacute;digo postal:</td>
	<td class=texto align=left><input type="text" name="cp" size="5" maxlength="5" class="campo" value="<%=campo(4,1)%>"></td>
</tr>
<tr>
	<td class=texto align=right valign="middle">Esta dirección corresponde a:</td>
	<td class=texto align=left><input type="text" name="cp" size="50" maxlength="200" class="campo" value="<%=valor(EliminaInyeccionSQL(request(campo(10,0)))) %>"></td>
</tr>
<tr>
	<td class=texto align=right valign="middle">Anotación:</td>
	<td class=texto align=left><textarea name="comentarios_2006" cols="48" rows="3" class="campo"><%=unquote(EliminaInyeccionSQL(request(campo(11,0)))) %></textarea></td>
</tr>
</table>
<br>&nbsp;
<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td class="texto" colspan="3" align="left">Materiales que est&aacute;s interesado/a en recibir:</td></tr>
<tr><td>&nbsp;</td><td><strong>Código</strong></td><td><strong>Nombre</strong></td></tr>
<tr><td class="celda"><input type="checkbox" name="SET01_2007" value="1" class="campo" <% if campo(5,1)="1" then response.write "checked" %>>&nbsp;</td><td class="celda">SET01</td><td class="celda">Vídeo en DVD: la participación de los trabajadores en la mejora del comportamiento ambiental de las empresas</td></tr>
<tr><td class="celda"><input type="checkbox" name="AE01_2007" value="1" class="campo" <% if campo(6,1)="1" then response.write "checked" %>>&nbsp;</td><td class="celda">AE01</td><td class="celda">Estudio: La incidencia y la aplicación de la normativa Seveso. El cumplimiento de las medidas y obligaciones que afectan a los trabajadores en este ámbito</td></tr>
<tr><td class="celda"><input type="checkbox" name="AE02_2007" value="1" class="campo" <% if campo(7,1)="1" then response.write "checked" %>>&nbsp;</td><td class="celda">AE02</td><td class="celda">Estudio: Situación del cumplimiento de la LPCIC en el último año para la obtención de la AAI en los centros de trabajo afectados</td></tr>
<tr><td class="celda"><input type="checkbox" name="AE03_2007" value="1" class="campo" <% if campo(8,1)="1" then response.write "checked" %>>&nbsp;</td><td class="celda">AE03</td><td class="celda">Estudio: Situación de la prevención de los incendios forestales y del personal que trabaja en su extinción</td></tr>
<tr><td class="celda"><input type="checkbox" name="AE04_2007" value="1" class="campo" <% if campo(9,1)="1" then response.write "checked" %>>&nbsp;</td><td class="celda">AE04</td><td class="celda">Estudio: ECO-OPINAS 2007. Actitudes, opiniones y necesidades formativas en materia ambiental de las y los trabajadores</td></tr>
</table>

				</div>
				<p>&nbsp;</p>
			<p align="center"><input type="button" class="boton" value="IR A LA PÁGINA DE INICIO" onClick="location.href='http://www.ecoinformas.com/'"></p>
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