<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	'OBJConnection.Open "driver={sql server};server=osiris.servidoresdns.net;database=qc507;UID=qc507;PWD=sql"
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"
	
	
	'----- Si es restringida y no estás identificado no puedes entrar
	if session("id_ecogente")="" then response.redirect "acceso.asp?idpagina="&idpagina
	
	numeracion = "AIBD"
	seccion = asc(mid(numeracion,3,1))-64

	idpagina = 564	'--- página FORO (solo para registrar visitas)
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

	iden = session("id_ecogente")
	tipo = Request("tipo")
	if tipo="" then tipo="1"
	prim = Request("prim")	'-----Este es el primer mensaje del listado de esta página
	if prim="" then	
		tipop = tipo
		orden = "SELECT * FROM ECO_FOROS where nivel=0 AND tipo="&clng(tipop)
		prim = 0
	else
		orden = "SELECT * FROM ECO_FOROS where id="&clng(prim)
	end if
	Set DSql = OBJConnection.Execute(orden)
 	
	sqlquery2 = "SELECT idgente,asunto,texto as descripcion FROM ECO_FOROS WHERE tipo="&tipo&" AND nivel=0"
	Set objRecordset2 = OBJConnection.Execute(sqlquery2)
	descripcionforo = objrecordset2("descripcion")
	nombreforo = objrecordset2("asunto")


%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ECOinformas: Foro</title>
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
			<div id="encabezado_nuevo2">
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
			<div id="menusup2">
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
			<div class="textsubmenu" id="submenusup2">
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
             				<% if len(numeracion)>3 then
              				     response.write "<br><p class=campo>Est&aacute;s en: "
              				     for i=1 to len(numeracion)-3
              				   	'sql = "SELECT titulo,idpagina FROM WEBISTAS_PAGINAS WHERE visible=1 AND numeracion='"&mid(numeracion,1,2+i)&"'" 
              				   	sql = "SELECT titulo,idpagina FROM WEBISTAS_PAGINAS WHERE numeracion='"&mid(numeracion,1,2+i)&"'" 
			   	   	   	set objRecordset = Server.CreateObject ("ADODB.Recordset")
		   	   	           	set objRecordset = OBJConnection.Execute(sql)
		   	   	           	response.write "<a href=index.asp?idpagina="&objRecordset("idpagina")&">"&objRecordset("titulo")&"</a>&nbsp;&gt;&nbsp;"
              				     next
              				     response.write titulo&"Foro de experiencias</p>"
              				   else
              				     response.write "<p class=campo>&nbsp;</p>"
              				   end if %>
				<br>&nbsp;
<!------------------ FORO -->
<p>
<b>Foro:&nbsp;</b><%=nombreforo%><br>
<b>Descripción:&nbsp;</b><%=descripcionforo%><br><br>
<center><input type="button" class="boton" value="PUBLICAR UN MENSAJE" onclick="javascript:void(window.open('foro_mensaje_nuevo.asp?tipo=<%=tipo%>&actualiza=<%=time()%>', '', 'resizable=yes,scrollbars=yes,width=550,height=290'))"></center>
</p>


<table border="0" cellpadding="0" cellspacing="0">
 <% 
	menpag = 15	'------Número de mensajes por página
	
	if not(DSQL.eof and DSQL.bof) then
	primero = clng(DSQL("id")) 
	sig = clng(DSQL("sig"))
	men = 1

	Do while sig<>0 and men<=menpag
	sqlquery4 = "SELECT idmensaje FROM ECO_FOROS_LEIDOS WHERE idmensaje="&sig&" AND idgente="&iden
	set dsql1 = OBJConnection.Execute(SQLQuery4)
	if dsql1.eof or dsql1.bof then
		graf = "leido_si" '--leído
		else
		graf = "leido_no" '-- no leído
	end if
	orden = "SELECT ECO_FOROS.*,ECOINFORMAS_GENTE.nombre,ECOINFORMAS_GENTE.apellidos FROM ECO_FOROS LEFT JOIN ECOINFORMAS_GENTE ON ECO_FOROS.idgente=ECOINFORMAS_GENTE.idgente WHERE id="&sig
	Set DSql = OBJConnection.Execute(orden)
	if men<>menpag then sig=clng(DSQL("sig"))
%>
          <% if dsql("nivel")=1 then %>          
          <tr>
            <td class="campo" height="0" colspan="1">&nbsp;</td>
          </tr>
	  <% end if %>
          <tr>
            <td class="campo" valign="middle">
            <table border="0" cellpadding="0" cellspacing="0"><tr>
            	<td class="campo"><img border="0" src="imagenes/<%=graf%>.gif" width="10" height="10" name="imagen<%=men%>"></td>
            	<% if dsql("nivel")<>1 then %>
            	<td class="campo" width="<%=14*(dsql("nivel"))%>" align="right" valign="top">
            	<img border="0" src="imagenes/esquina.gif" width="8" height="10">
            	</td>
            	<% end if %>
            	<td class="campo">&nbsp;
            	<font class="campo"><a onmouseup="javascript:document['imagen<%=men%>'].src ='imagenes/leido_actual.gif';window.open('foro_mensaje.asp?an=<%=permisoforo%>&tipo=<%=tipo%>&id=<%=dsql("id")%>&actualiza=<%=time()%>', '', 'scrollbars=yes,resizable=yes,width=575,height=350');" style="text-decoration:none; cursor:hand"><%=dsql("asunto")%></a></font>
            	<font style="font-family: Verdana; font-size: 8pt; color: #555555">&nbsp;·&nbsp;(<%=DSQL("nombre")&" "&DSQL("apellidos")%>)&nbsp;·&nbsp;<%=DSQL("fecha")%></font></td>
            </tr></table>
            </td>
          </tr>
	  
          
<%
	 men = men+1
	 loop
	 Dsql.close
	 end if  
%>

      <tr><td class="ihef" colspan="1"><br>
      <% if (clng(prim)<>0) or (clng(sig)<>0) then %>
      <font class="negro">PÁGINA: <%if (clng(prim)<>0) then%> <a href="javascript:history.back(1)">ANTERIOR </a><%end if%><%if clng(sig)<>0 then%>· <a href="foro.asp?tipo=<%=tipo%>&prim=<%=sig%>&actualiza=<%=time()%>">SIGUIENTE</a><%end if%></font>
      <% end if %>
      </td></tr>
</table>
<!------------------ FIN FORO -->

				</div>
				<p>&nbsp;</p>
			</div>

			<map name="Map2" id="Map2">
            		<area shape="rect" coords="300,8,392,66" href="http://www.fundacion-biodiversidad.es" target="_blank" alt="Fundación Biodiversidad" />
            		<area shape="rect" coords="393,8,539,66" href="http://www.istas.ccoo.es" target="_blank" alt="Instituto Sindical de Trabajo, Ambiente y Salud" />
      			<area shape="rect" coords="540,8,694,66" href="http://www.mtas.es/UAFSE/default.htm" target="_blank" alt="Fondo Social Europeo" />
      			</map>
			<img src="imagenes/pie2.jpg" width="708" border="0" usemap="#Map2">

    			</div>
    		</div>
		<div id="sombra_abajo"></div>
	</div>
</body>
</html>