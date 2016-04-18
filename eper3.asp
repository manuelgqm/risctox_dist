<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	'OBJConnection.Open "driver={sql server};server=osiris.servidoresdns.net;database=qc507;UID=qc507;PWD=sql"
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"
	
	
	'----- Si es restringida y no estás identificado no puedes entrar
	if session("id_ecogente")="" then response.redirect "acceso.asp?idpagina="&idpagina
	
	numeracion = "AIBBBE"
	seccion = asc(mid(numeracion,3,1))-64

	idpagina = 662	'--- página Buscador de fuentes (instalaciones/empresas) y emisiones contaminantes
	'----- Registrar la visita
	IP = Request.ServerVariables("REMOTE_ADDR")
	Set MiBrowser = Server.CreateObject("MSWC.BrowserType")
	navegador = MiBrowser.Browser
	if session("id_ecogente")<>"" then 
		usuario = session("id_ecogente")
	else
		usuario = 0
	end if
	orden = "INSERT INTO WEBISTAS_VISITAS (fecha,hora,IP,navegador,idpagina,idgente) VALUES ('"&date()&"','"&time()&"','"&IP&"','"&navegador&"',"&idpagina&","&usuario&")"
	Set objRecordset = Server.CreateObject ("ADODB.Recordset")
	Set objRecordset = OBJConnection.Execute(orden)



	titulocompleto = ""
	for i=2 to len(numeracion)
		sql = "SELECT titulo,numeracion,idpagina FROM WEBISTAS_PAGINAS WHERE numeracion='" & mid(numeracion,1,i) & "'"
		set objRecordset = Server.CreateObject ("ADODB.Recordset")
		set objRecordset = OBJConnection.Execute(sql)
		if i<>2 then titulocompleto = titulocompleto & "&nbsp;&gt;&nbsp;" 
		titulocompleto = titulocompleto & "<a href=index.asp?idpagina="&objrecordset("idpagina")&">"&objrecordset("titulo")&"</a>"
	next 
	
	FUNCTION vistaprevia(texto)
		texto = replace(texto,chr(13),"<br>")
		texto = replace(texto,"'","&#39;")
		texto = replace(texto,"<v1>","<img src=imagenes/vineta.gif>&nbsp;")
		texto = replace(texto,"<v2>","&nbsp;&nbsp;<img src=imagenes/vineta.gif>&nbsp;")
		texto = replace(texto,"<v3>","&nbsp;&nbsp;&nbsp;&nbsp;<img src=imagenes/vineta.gif>&nbsp;")
		texto = replace(texto,"<v4>","&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src=imagenes/vineta.gif>&nbsp;")
		texto = replace(texto,"<v5>","&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src=imagenes/vineta.gif>&nbsp;")
		texto = replace(texto,"<v6>","&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src=imagenes/vineta.gif>&nbsp;")
		texto = replace(texto,"<pag=","<a href=index.asp?idpagina=")
		texto = replace(texto,"</pag>","</a>")
		texto = replace(texto,"<e=","<a target=_blank href=abreenlace.asp?idenlace=")
		texto = replace(texto,"<er=","<a target=_blank href=abreenlacer.asp?idenlace=")
		texto = replace(texto,"</e>","</a>")
		texto = replace(texto,"<t>","<font class=titulo"&(seccion)&">")
		texto = replace(texto,"</t>","</font>")
		texto = replace(texto,"<st>","<font class=subtitulo"&(seccion)&">")
		texto = replace(texto,"</st>","</font>")
		texto = replace(texto,"<pd>","<table width=95% align=center cellpadding=10 cellspacing=0 class=tabla><tr><td>")
		texto = replace(texto,"</pd>","</td><td valign=top align=center><img src=pd.gif></td></tr></table>")
		vistaprevia = texto
		
	END FUNCTION

	function contenga(codigo)
		texto_sql = "("
		for n=1 to len(codigo)-1
			texto_sql = texto_sql & " ascii(substring(numeracion,"&cstr(n)&",1))=" & asc(mid(codigo,n,1)) & " AND "
		next
		texto_sql = texto_sql & " ascii(substring(numeracion,"&cstr(len(codigo))&",1))=" & asc(mid(codigo,len(codigo),1)) & ")"
		contenga = texto_sql
		'response.write texto_sql & "<br>"
	end function

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ECOinformas: Buscador de fuentes (instalaciones/empresas) y emisiones contaminantes</title>
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
			<div id="encabezado_nuevo<% response.write (seccion) %>">
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
			<div id="menusup<% response.write (seccion) %>">
			<table border="0" cellspacing="5" cellpadding="0">
            			<tr class="textmenusup">
<%              				sql = "SELECT titulo,idpagina,numeracion FROM WEBISTAS_PAGINAS WHERE visible=1 AND numeracion LIKE '"&mid(numeracion,1,3)&"%' AND len(numeracion)=4 ORDER BY numeracion"
			   	   	   	set objRecordset = Server.CreateObject ("ADODB.Recordset")
		   	   	           	set objRecordset = OBJConnection.Execute(sql)
		   	   	           	do while not objRecordset.eof
              						response.write "<td class=textmenusup>"
							if mid(numeracion,1,4)=mid(objRecordset("numeracion"),1,4) then
								response.write lcase(objRecordset("titulo"))
              						else
              							response.write "<a href=index.asp?idpagina="&objRecordset("idpagina")&" style='text-decoration:none'>"&lcase(objRecordset("titulo"))&"</a>"
              						end if
              						response.write "</td><td class=textmenusup>|</td>"
							objrecordset.movenext
 						loop %>
              			</tr>
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
            			<tr><td align="right">Usuari<%=usuario_sexo%> identificad<%=usuario_sexo%>:&nbsp;<%=usuario%>&nbsp;</td></tr>
          		</table>
			</div>
       			<% end if %>
			
			<% if 1=0 and len(numeracion)>3 then
			   sql = "SELECT titulo,idpagina,numeracion FROM WEBISTAS_PAGINAS WHERE visible=1 AND ((len(numeracion)=5 AND numeracion LIKE '"&mid(numeracion,1,4)&"%')"
			   if len(numeracion)>4 then sql = sql & " OR (len(numeracion)>4 AND numeracion LIKE '"&mid(numeracion,1,5)&"%')"
			   sql = sql & ") ORDER BY numeracion"
			   set objRecordset = Server.CreateObject ("ADODB.Recordset")
		   	   set objRecordset = OBJConnection.Execute(sql)
		   	   if not objRecordset.eof then
		   	   submenu = 1 %>
		   	   <div id="margen_izquierdo<% response.write (seccion) %>">
			<% do while not objRecordset.eof %>
			<table cellpadding="5" cellspacing="1" border=0 width="95%" align="center">
			<tr>
			<% if len(objRecordset("numeracion"))=5 then %>
			<td class="campo"><img src="imagenes/flecha.gif">&nbsp;
			<% else %>
			<td class="campo" width="<%=(len(objRecordset("numeracion"))-5)*15 %>">&nbsp;</td><td class="campo" width="100%">
			<% end if %>
			<a href="index.asp?idpagina=<%=objRecordset("idpagina")%>">
			<% if objRecordset("idpagina")=idpagina then 
				'response.write "<font style='background:#EEEEEE'>"&objRecordset("titulo")&"</font>"
				response.write "<b>"&objRecordset("titulo")&"</b>"
			   else
			   	response.write objRecordset("titulo")
			   end if %>
			</a>
			</td></tr></table>
			<% objRecordset.movenext
			   loop %>
			   </div>
			<% end if
			   end if %>

			<% if submenu=1 or cstr(idpagina)="548" then %>
			<div id="interiortext">
			<% else %>
			<div id="texto">
			<% end if %>
			
				<div class="texto">
             				<% if len(numeracion)>3 then
              				     response.write "<br><p class=campo>Est&aacute;s en: "
              				     for i=1 to len(numeracion)-3
              				   	sql = "SELECT titulo,idpagina FROM WEBISTAS_PAGINAS WHERE visible=1 AND numeracion='"&mid(numeracion,1,2+i)&"'" 
			   	   	   	set objRecordset = Server.CreateObject ("ADODB.Recordset")
		   	   	           	set objRecordset = OBJConnection.Execute(sql)
		   	   	           	response.write "<a href=index.asp?idpagina="&objRecordset("idpagina")&">"&objRecordset("titulo")&"</a>&nbsp;&gt;&nbsp;"
              				     next
              				     response.write titulo&"<a href=index.asp?idpagina=661>Buscador de fuentes (instalaciones/empresas) y emisiones contaminantes</a>&nbsp;&gt;&nbsp;<a href='eper1.asp'>Inicio</a></p>"
              				   else
              				     response.write "<p class=campo>&nbsp;</p>"
              				   end if %>
				
					<% sql = "SELECT * FROM ECO_EPER WHERE numero="&request("id")
			   		   set objRecordset = Server.CreateObject ("ADODB.Recordset")
			   		   set objRecordset = OBJConnection.Execute(sql)
			   		   dim campos(100,2)
			   		   campos(1,1)="Numero"
					   campos(2,1)="Nombre del complejo"
					   campos(3,1)="Empresa matriz"
					   campos(4,1)="Direccion"
					   campos(5,1)="Codigo postal"
					   campos(6,1)="Provincia"
					   campos(7,1)="Codigo CNAE"
					   campos(8,1)="Actividad economica principal"
					   campos(9,1)="Volumen de produccion"
					   campos(10,1)="Organismos reguladores"
					   campos(11,1)="Nº de instalaciones"
					   campos(12,1)="Nº de empleados"
					   campos(13,1)="empleados s/n"
					   campos(14,1)="Nº de horas de trabajo"
					   campos(15,1)="Facturación"
					   campos(16,1)="fact si/no"
					   campos(17,1)="Epigrafe"
					   campos(18,1)="Codigo NoseP"
					   campos(19,1)="Otros datos"
					   campos(20,1)="Nombre"
					   campos(21,1)="Sindicato"
					   campos(22,1)="mas datos"
					   campos(23,1)="Nombre 2"
					   campos(24,1)="Sindicato 2"
					   campos(25,1)="Mas datos 2"
					   campos(26,1)="Nombre 3"
					   campos(27,1)="Sindicato 3"
					   campos(28,1)="Mas datos 3"
					   campos(29,1)="Nombre 4"
					   campos(30,1)="Sindicato 4"
					   campos(31,1)="Mas datos 4"
					   campos(32,1)="Nombre 5"
					   campos(33,1)="Sindicato 5"
					   campos(34,1)="Mas datos 5"
					   campos(35,1)="ISO 9001"
					   campos(36,1)="EFQM"
					   campos(37,1)="ISO 14001"
					   campos(38,1)="EMAS"
					   campos(39,1)="GESTION DE RIESGOS LABORALES"
					   campos(40,1)="RESPONSABILIDAD SOCIAL COORPORATIVA"
					   campos(41,1)="Contaminante 1"
					   campos(42,1)="M/C/E 1:"
					   campos(43,1)="Concentracion 1"
					   campos(44,1)="Contaminante 2"
					   campos(45,1)="M/C/E 2:"
					   campos(46,1)="Concentracion 2"
					   campos(47,1)="Contaminante 3"
					   campos(48,1)="M/C/E 3:"
					   campos(49,1)="Concentracion 3"
					   campos(50,1)="Contaminante 4"
					   campos(51,1)="M/C/E 4:"
					   campos(52,1)="Concentracion 4"
					   campos(53,1)="Contaminante 5"
					   campos(54,1)="M/C/E 5:"
					   campos(55,1)="Concentracion 5"
					   campos(56,1)="Contaminante 6"
					   campos(57,1)="M/C/E 6:"
					   campos(58,1)="Concentracion 6"
					   campos(59,1)="Contaminante 7"
					   campos(60,1)="M/C/E 7:"
					   campos(61,1)="Concentracion 7"
					   campos(62,1)="Contaminante 8"
					   campos(63,1)="M/C/E 8:"
					   campos(64,1)="Concentracion 8"
					   campos(65,1)="Contaminante 9"
					   campos(66,1)="M/C/E 9:"
					   campos(67,1)="Concentracion 9"
					   campos(68,1)="Contaminante 10"
					   campos(69,1)="M/C/E 10:"
					   campos(70,1)="Concentracion 10"
					   campos(71,1)="Contaminante 11"
					   campos(72,1)="M/C/E 11:"
					   campos(73,1)="Concentracion 11"
					   campos(74,1)="Contaminante 12"
					   campos(75,1)="M/C/E 12:"
					   campos(76,1)="Concentracion 12"
					   campos(77,1)="Contaminante 1,1"
					   campos(78,1)="M/C/E 1,1:"
					   campos(79,1)="Concentracion 1,1"
					   campos(80,1)="Contaminante 2,2"
					   campos(81,1)="M/C/E 2,2:"
					   campos(82,1)="Concentracion 2,2"
					   campos(83,1)="Contaminante 3,3"
					   campos(84,1)="M/C/E 3,3:"
					   campos(85,1)="Concentracion 3,3"
					   campos(86,1)="Contaminante 4,4"
					   campos(87,1)="M/C/E 4,4:"
					   campos(88,1)="Concentracion"
					   campos(89,1)="Contaminante 5,5"
					   campos(90,1)="M/C/E 5,5"
					   campos(91,1)="Concentracion 5,5"
					   campos(92,1)="Contaminante 6,6"
					   campos(93,1)="M/C/E 6,6:"
					   campos(94,1)="Conentracion 6,6"
					   campos(95,1)="Contaminante 7,7"
					   campos(96,1)="M/C/E 7,7:"
					   campos(97,1)="Concentracion 7,7"
					   campos(98,1)="Contaminante 8,8"
					   campos(99,1)="M/C/E 8,8"
					   campos(100,1)="Concentracion 8,8"

					   campos(1,2)=1
					   campos(2,2)=1
					   campos(3,2)=1
					   campos(4,2)=1
					   campos(5,2)=1
					   campos(6,2)=1
					   campos(7,2)=1
					   campos(8,2)=1
					   campos(9,2)=1
					   campos(10,2)=1
					   campos(11,2)=1
					   campos(12,2)=1
					   campos(13,2)=1
					   campos(14,2)=1
					   campos(15,2)=1
					   campos(16,2)=1
					   campos(17,2)=1
					   campos(18,2)=1
					   campos(19,2)=1
					   campos(20,2)=0
					   campos(21,2)=0
					   campos(22,2)=0
					   campos(23,2)=0
					   campos(24,2)=0
					   campos(25,2)=0
					   campos(26,2)=0
					   campos(27,2)=0
					   campos(28,2)=0
					   campos(29,2)=0
					   campos(30,2)=0
					   campos(31,2)=0
					   campos(32,2)=0
					   campos(33,2)=0
					   campos(34,2)=0
					   campos(35,2)=0
					   campos(36,2)=0
					   campos(37,2)=0
					   campos(38,2)=0
					   campos(39,2)=0
					   campos(40,2)=0
					   campos(41,2)=1
					   campos(42,2)=1
					   campos(43,2)=1
					   campos(44,2)=1
					   campos(45,2)=1
					   campos(46,2)=1
					   campos(47,2)=1
					   campos(48,2)=1
					   campos(49,2)=1
					   campos(50,2)=1
					   campos(51,2)=1
					   campos(52,2)=1
					   campos(53,2)=1
					   campos(54,2)=1
					   campos(55,2)=1
					   campos(56,2)=1
					   campos(57,2)=1
					   campos(58,2)=1
					   campos(59,2)=1
					   campos(60,2)=1
					   campos(61,2)=1
					   campos(62,2)=1
					   campos(63,2)=1
					   campos(64,2)=1
					   campos(65,2)=1
					   campos(66,2)=1
					   campos(67,2)=1
					   campos(68,2)=1
					   campos(69,2)=1
					   campos(70,2)=1
					   campos(71,2)=1
					   campos(72,2)=1
					   campos(73,2)=1
					   campos(74,2)=1
					   campos(75,2)=1
					   campos(76,2)=1
					   campos(77,2)=1
					   campos(78,2)=1
					   campos(79,2)=1
					   campos(80,2)=1
					   campos(81,2)=1
					   campos(82,2)=1
					   campos(83,2)=1
					   campos(84,2)=1
					   campos(85,2)=1
					   campos(86,2)=1
					   campos(87,2)=1
					   campos(88,2)=1
					   campos(89,2)=1
					   campos(90,2)=1
					   campos(91,2)=1
					   campos(92,2)=1
					   campos(93,2)=1
					   campos(94,2)=1
					   campos(95,2)=1
					   campos(96,2)=1
					   campos(97,2)=1
					   campos(98,2)=1
					   campos(99,2)=1
					   campos(100,2)=1
					   
					    %>
					<table align="center" cellpadding="3" cellspacing="3">
					<tr><td class="titulo2" align="center" colspan="2">1. Descripción de la empresa</td></tr>
					<tr><td class="texto" align="right"><img src="imagenes/ico_puntito.gif" align="absmiddle"></td><td class="texto"><b><%=objRecordset("nombre del complejo")%></b></td></tr>
					<% for i=3 to 100
						campo =  campos(i,1)
						valor = objRecordset(campos(i,1))
						if valor = False then valor="No"
						if valor = True then valor="Sí"
						if isnull(valor) then valor="" %>
					<% if 1=0 and i=20 then %>
					<tr><td class="titulo2" align="center" colspan="2">2. Datos sindicales</td></tr>
					<tr><td class="texto" align="center" bgcolor="#DDDDDD" colspan="2"><b>Comité de empresa/Delegado de personal</b></td></tr>
					<% end if %>
					<% if 1=0 and i=29 then %>
					<tr><td class="texto" align="center" bgcolor="#DDDDDD" colspan="2"><b>Delegado de prevención</b></td></tr>
					<% end if %>
					<% if 1=0 and i=32 then %>
					<tr><td class="texto" align="center" bgcolor="#DDDDDD" colspan="2"><b>Delegado sindical</b></td></tr>
					<% end if %>
					<% if 1=0 and i=35 then %>
					<tr><td class="titulo2" align="center" colspan="2">3. Gestión ambiental</td></tr>
					<tr><td class="texto" align="center" bgcolor="#DDDDDD" colspan="2"><b>Gestión de la calidad</b></td></tr>
					<% end if %>
					<% if 1=0 and i=37 then %>
					<tr><td class="texto" align="center" bgcolor="#DDDDDD" colspan="2"><b>Gestión de medioambiente</b></td></tr>
					<% end if %>
					<% if i=41 then %>
					<tr><td class="titulo2" align="center" colspan="2">2. Informe EPER</td></tr>
					<tr><td class="texto" align="center" bgcolor="#DDDDDD" colspan="2"><b>Contaminantes emitidos al aire</b></td></tr>
					<% end if %>
					<% if i=77 then %>
					<tr><td class="texto" align="center" bgcolor="#DDDDDD" colspan="2"><b>Contaminantes emitidos al agua</b></td></tr>
					<% end if %>
						<% if campos(i,2)=1 then %>
					<tr><td class="texto" align="right" bgcolor="#DDDDDD"><%=campo %>:</td><td class="texto" valign=top><%=valor %></td></tr>
						<% end if
					   next %>

					</table>
					<p align="center"><input type="button" class="boton" value="imprimir" onclick="print()"></p>
					
				
				</div>
				<p>&nbsp;</p>
			</div>
			
			<map name="Map1" id="Map1">
            		<area shape="rect" coords="307,38,399,104" href="http://www.fundacion-biodiversidad.es" target="_blank" alt="Fundación Biodiversidad" />
            		<area shape="rect" coords="400,38,546,104" href="http://www.istas.ccoo.es" target="_blank" alt="Instituto Sindical de Trabajo, Ambiente y Salud" />
      			<area shape="rect" coords="547,38,701,104" href="http://www.mtas.es/UAFSE/default.htm" target="_blank" alt="Fondo Social Europeo" />
      			</map>
			<map name="Map2" id="Map2">
            		<area shape="rect" coords="300,8,392,66" href="http://www.fundacion-biodiversidad.es" target="_blank" alt="Fundación Biodiversidad" />
            		<area shape="rect" coords="393,8,539,66" href="http://www.istas.ccoo.es" target="_blank" alt="Instituto Sindical de Trabajo, Ambiente y Salud" />
      			<area shape="rect" coords="540,8,694,66" href="http://www.mtas.es/UAFSE/default.htm" target="_blank" alt="Fondo Social Europeo" />
      			</map>
      			<map name="Map3" id="Map3">
            		<area shape="rect" coords="300,18,392,80" href="http://www.fundacion-biodiversidad.es" target="_blank" alt="Fundación Biodiversidad" />
            		<area shape="rect" coords="393,18,539,80" href="http://www.istas.ccoo.es" target="_blank" alt="Instituto Sindical de Trabajo, Ambiente y Salud" />
      			<area shape="rect" coords="540,18,694,80" href="http://www.mtas.es/UAFSE/default.htm" target="_blank" alt="Fondo Social Europeo" />
      			</map>
			<img src="imagenes/pie<% response.write (seccion) %>.jpg" width="708" border="0" usemap="#Map<% response.write (seccion) %>">
      			
    			
    		</div>
    	</div>
	<div id="sombra_abajo"></div>
</div>
</body>
</html>