<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ECOinformas: Asesoramientos</title>
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



<%

        ruta_upload_fis     = "d:\xvrt\istas.net\html\Recursos\"
        ruta_upload_log     = "../../Recursos/"

 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset

        Set OBJConnection = Server.CreateObject("ADODB.Connection")
        'OBJConnection.Open "driver={sql server};server=osiris.servidoresdns.net;database=qc507;UID=qc507;PWD=sql"
        OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"

        estado_consulta   = Request("estado_consulta")
        tema_consulta_fil = Request("tema_consulta_fil")
                
	'------------------------------------------				
	' Tengo que saber qué tipo de asesor es:
	'   0 : no
	'   1 : sí
	'   2 : coordinador
	
	orden = "SELECT asesor from ECOINFORMAS_GENTE WHERE idgente='"&session("id_ecogente")&"'"
	set objRecordset = OBJConnection.Execute(orden)
	if not objRecordset.eof then
	 asesor = objRecordset("asesor")
	end if	
	if asesor="" or isnull(asesor) then
	 asesor = 0
	else
	 asesor = cint(asesor)
	end if 
	
	
	'------------------------------------------

	
	'
	'
	'
	totlin = 6
	'
	act_pag=Request("act_pag")
        IF Request.QueryString("pag")="sig" THEN
         act_pag=act_pag+1
        END IF
        IF Request.QueryString("pag")="ant" THEN
          act_pag=act_pag-1
        END IF
        IF act_pag="" THEN
          act_pag=1
        END IF
        	
	
	'
	'
	'
	
	
	
	'----- Si es restringida y no estás identificado no puedes entrar
	'if session("id_ecogente")="" then response.redirect "acceso.asp?idpagina="&idpagina
	'---- ATENCIÓN: ponerlo cuando publiquemos en abierto
	
	numeracion = "AIBAA"
	
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
		texto = replace(texto,"<t>","<font class=titulo3>")
		texto = replace(texto,"</t>","</font>")
		texto = replace(texto,"<st>","<font class=subtitulo3>")
		texto = replace(texto,"</st>","</font>")
		texto = replace(texto,"<pd>","<table width=95% align=center cellpadding=10 cellspacing=0 class=tabla><tr><td>")
		texto = replace(texto,"</pd>","</td><td valign=top align=center><img src=pd.gif></td></tr></table>")
		vistaprevia = texto
		
	END FUNCTION

%>

<SCRIPT LANGUAGE="JavaScript">
<!--

// -->
</SCRIPT>
</head>

<body>
<form name=asesora>
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
            			<tr><td align="right">Usuari<%=usuario_sexo%> identificad<%=usuario_sexo%>:&nbsp;<%=usuario%>&nbsp;</td></tr>
          		</table>
			</div>
       			<% end if %>
			

			<div id="texto">
			
				<div class="texto">
             				<% if len(numeracion)>3 then
              				     response.write "<br><p class=campo>Est&aacute;s en: "
              				     for i=1 to len(numeracion)-3
              				   	sql = "SELECT titulo,idpagina FROM WEBISTAS_PAGINAS WHERE visible=1 AND numeracion='"&mid(numeracion,1,2+i)&"'" 
			   	   	   	set objRecordset = Server.CreateObject ("ADODB.Recordset")
		   	   	           	set objRecordset = OBJConnection.Execute(sql)
		   	   	           	response.write "<a href=index.asp?idpagina="&objRecordset("idpagina")&">"&objRecordset("titulo")&"</a>&nbsp;&gt;&nbsp;"
              				     next
              				     response.write "Servicio de asesoramiento</p>"
              				   else
              				     response.write "<p class=campo>&nbsp;</p>"
              			   end if %>
				
				<p class=titulo2>Servicio de asesoramiento</p>


<!------------------------------------------------------------------------------------------------------------  JOSE ---------------------------->
				
				<% if asesor<>2 then %>
				 <p class="texto">Aquí puedes consultar los asesoramientos que hayas solicitado o realizar un 
				 <input type="button" class="boton" value="Nuevo asesoramiento" onclick="javascript:nuevo_asesoramiento();"></p>
				<% else %>
				 <p class="texto">Éstos son los asesoramientos.</p>
				<% end if %>
				
				<% if asesor=1 or asesor=2 then %>
				<p class="texto">Estado:&nbsp;<% CALL valores("015","estado_consulta",estado_consulta,"valor") %>
				<% end if %> 
                                
				<% if asesor=2 then %>
				&nbsp;&nbsp;Tema:&nbsp;<% CALL valores_tem("010","tema_consulta_fil",tema_consulta_fil,"desc1") %>
				<% else %>
				<input type=hidden name=tema_consulta_fil>
				<% end if %> 
				
				</p>
				
				<%
				
				
			        CALL cabecera
				 			  
				orden = "SELECT DISTINCT EC.idconsulta,EC.*,EV.desc1,EG.*,E2.desc1 as tema "
				orden = orden & " from ECOINFORMAS_CONSULTAS    EC  "
				orden = orden & " LEFT JOIN ECOINFORMAS_VALORES EV ON EC.estado=EV.valor "
				orden = orden & " LEFT JOIN ECOINFORMAS_GENTE   EG ON EC.usuario=EG.idgente "
				orden = orden & " LEFT JOIN ECOINFORMAS_TEM_ASE ET ON EC.tema_consulta=ET.valor "
				orden = orden & " LEFT JOIN ECOINFORMAS_VALORES E2 ON EC.tema_consulta=E2.valor "				
				
				'
				'
				' Sólo muestro las preguntas. Luego me meteré en otro bucle para ver las respuestas de cada pregunta.
 				condi = " EC.tipo_consulta=157 " 

 				' Montar la condición SQL en función del tipo de persona.				
				select case asesor
				 case 0 ' Usuario normal
				  CALL monta_condi (" EC.usuario='"&session("id_ecogente")&"' ")
				 case 1 ' Es un asesor(a)
				  CALL monta_condi (" (ET.asesor=" & session("id_ecogente") & " OR " & " EC.usuario='"&session("id_ecogente")&"') ")
				 case 2 ' Es el/la coordinador/a
				  'CALL monta_condi ("")
				end select 
				 
				' Poner el filtro del estado_consulta 
				if estado_consulta<>"" then
				 if estado_consulta<>"X" then ' 
				  CALL monta_condi(" EC.estado='"&estado_consulta&"' ")
				 else ' Caso de sin finalizar 151,152 y 153
				  CALL monta_condi(" (EC.estado=151 or EC.estado=152 or EC.estado=153) ")
				 end if 
				end if  
				  
				' Poner el filtro del tema_consulta_fil
				if tema_consulta_fil<>"" then
				  CALL monta_condi(" EC.tema_consulta='"&tema_consulta_fil&"' ")
				end if  
				  
				if condi<>"" then
				 orden = orden & " WHERE " & condi
				end if
				
				' Ordenado por idconsulta descendente, el primero será el último en verse
				orden = orden & " ORDER BY EC.idconsulta desc " 
				  
				'response.write "Orden:" & orden & "<br>"  
				  
 		   	   	set objRecordset = Server.CreateObject ("ADODB.Recordset")
			        objRecordset.Open orden, Objconnection,1  
			        
				if not objRecordset.eof then			        
			        
    				objRecordset.PageSize=totlin
    				objRecordset.AbsolutePage=act_pag
    				
    				num_preguntas = objRecordset.recordcount
    				num_paginas  = objRecordset.pagecount    
    				
                                desde_pag = (int((act_pag-1)/10))*10+1
                                hasta_pag = (int((act_pag+9)/10))*10
                                
                                if hasta_pag>num_paginas then
                                 hasta_pag = num_paginas
                                end if
                                
                                lin_act = 0

		         	do while not objRecordset.eof and lin_act<totlin
			         lin_act    = lin_act + 1
			         usuario    = trim(objRecordset("nombre"))&" "&trim(objRecordset("apellidos"))
			         estado     = trim(objRecordset("desc1"))
			         estado_id  = objRecordset("estado")			         
			         fecha      = objRecordset("fecha")
			         asunto     = trim(objRecordset("asunto"))
			         idconsulta = objRecordset("idconsulta")
			         ext        = objRecordset("fichero")
			         tema       = trim(objRecordset("tema"))
			         
			         CALL detalle
			         
			         ' Comprobar si hay respuestas a la pregunta
			         orden = "SELECT * FROM ECOINFORMAS_CONSULTAS EC "
       				 orden = orden & " LEFT JOIN ECOINFORMAS_GENTE   EG ON EC.usuario=EG.idgente "
       				 orden = orden & " WHERE EC.puntero="&idconsulta
       				 orden = orden & " ORDER BY EC.idconsulta "
			         
			         set objrecordset2 = objconnection.execute(orden)
			         do while not objrecordset2.eof 
			          usuario    = trim(objrecordset2("nombre"))&" "&trim(objrecordset2("apellidos"))			         
 			          fecha      = objrecordset2("fecha")
			          asunto     = trim(objrecordset2("asunto"))
		                  idconsulta = objRecordset2("idconsulta")
			          ext        = objRecordset2("fichero")		                  
			          CALL detalle2
			          objrecordset2.movenext
			         loop
			        
			         objRecordset.movenext
			        loop
				  
				CALL pie_pagina  
				
				else
				 CALL no_hay_datos
				end if
				  
				 %> 
				  
<!------------------------------------------------------------------------------------------------------------  JOSE ---------------------------->
				<p class=texto>&nbsp;</p>
				
				</div>
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
</form>
</body>
</html>










<%
'-----------------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------------


'-----------------------------------------------------------------------------------------------------------------------------------------

 function no_hay_datos 

%>
<br>
<br>
<table class="celda2" width="90%" align="center">
<tr><td>&nbsp;</td></tr>
<tr><td>NO HAY ASESORIAMIENTOS</td></tr>
<tr><td>&nbsp;</td></tr>
</table>

<% end function %>


<% 
'-----------------------------------------------------------------------------------------------------------------------------------------
function cabecera 

%>
<table class="tabla2" width="90%" align="center">

<tr>
 <td class="subtitulo2" align="left">&nbsp;</td>
 <% if asesor=1 or asesor=2 then %>
  <td class="subtitulo2" align="left">Usuario(a)</td>				   
 <% end if %>

 <td class="subtitulo2" align="left">Estado</td> 
 
 <% if asesor=2 then %>
  <td class="subtitulo2" align="left">Tema</td>				   
 <% end if %>
 
 <td class="subtitulo2" align="left">Fecha/hora</td>
 <td class="subtitulo2" align="left">&nbsp;</td> 
 <td class="subtitulo2" align="left">Asunto</td>
</tr>
<% end function %>



<% 
'-----------------------------------------------------------------------------------------------------------------------------------------
function detalle 

%>

<tr>
 <td class="celda2" align="left">
 <% if asesor<>0 then %>
  <input class=subtitulo type=button name=respon value="R" onclick="javascript:responder('<%=idconsulta%>');">
 <% end if %> 
 </td>
 
 <% if asesor=1 or asesor=2 then %>
  <td class="celda2c" align="left"><%=usuario%></td>				   
 <% end if %>

 <% if estado_id=151 then ' Estado SIN ASIGNAR ponerlo en rojo...%>
   <% color_estado="<font color=DD0000>"%>
 <% end if %>
  
 <td class="celda2c" align="left"><%=color_estado%><%=estado%></td>

 <% if asesor=2 then %>
  <td class="celda2c" align="left"><%=tema%>&nbsp;</td>				   
 <% end if %>
 
 <td class="celda2c" align="left"><%=fecha%></td>
 <td class="celda2c" align="left">
  <% if not isnull(ext) and ext<>"" then %>
   <img src="clip.gif" class=mano onclick="javscript:ver_adjunto('<%=idconsulta%>','<%=trim(ext)%>');" alt="Ver fichero adjunto <%=ext%>">
  <% end if %>  
  &nbsp;
 </td>
 
 <td class="celda2c" align="left">
  <a href="javascript:ver_pregunta('<%=idconsulta%>');"><%=asunto%></a>
 </td>
</tr>
<% end function %>

<% 
'-----------------------------------------------------------------------------------------------------------------------------------------
function detalle2 

%>
<tr>

 <td class="celda2d" align="left">&nbsp;</td>
 
 <% if asesor=1 or asesor=2 then %>
  <td class="celda2d" align="left">&nbsp;&nbsp;&nbsp;<%=usuario%></td>				   
 <% end if %>
 
 <td class="celda2d" align="left">&nbsp;</td>

 <% if asesor=2 then %>
  <td class="celda2d" align="left">&nbsp;</td>				   
 <% end if %>
 
 <td class="celda2d" align="left"><%=fecha%></td>
 <td class="celda2d" align="left">
  <% if not isnull(ext) and ext<>"" then %>
   <img src="clip.gif" class=mano onclick="javscript:ver_adjunto('<%=idconsulta%>','<%=trim(ext)%>');" alt="Ver fichero adjunto <%=ext%>">
  <% end if %>  
  &nbsp;
 </td>
 
 <td class="celda2d" align="left">
 <a href="javascript:ver_pregunta('<%=idconsulta%>');"><%=asunto%></a>
 </td>
</tr>
<% end function %>



<%
'-----------------------------------------------------------------------------------------------------------------------------------------
function pie_pagina

 %>

</table>
<br>
<table cellspacing=0 border=0 cellpadding=0 class=texto width="100%" height="100%">
<tr >

 <td width="25%" align=center>Asesoramientos:&nbsp;<%=num_preguntas%></b></td>
 
 <td width="50%" align=center>
 
 <% if clng(act_pag)>1 then %>
   <A HREF="javascript:paginar('ant','<%=act_pag%>');">[anterior]</A></font>&nbsp;&nbsp;
 <% else %>
  [anterior]
 <% end if %>

 <% 
 
 for x=desde_pag to hasta_pag
  negrita1="<b>"
  negrita2="</b>"
  if x=clng(act_pag) then
   negrita1="<font color=#DD0000>"
   negrita2="</font>"
  end if
  %><A HREF="javascript:paginar('','<%=x%>');"><%=negrita1%><%=x%><%=negrita2%></A>&nbsp;<%
 next
 %>

 <% if clng(act_pag)<clng(num_paginas) then %>
  <font class="tablab">
   <A HREF="javascript:paginar('sig','<%=act_pag%>');">[siguiente]</A></font>  
 <% else %>
  <!-- <font class="tabla">[SIGUIENTE]</font> -->
 <% end if %>
 </td>

 <td width="25%" align=center>Página&nbsp;<%=act_pag%>&nbsp;de&nbsp;<%=num_paginas%></td>
 
 </tr>
 </table>
 
<% end function %>

<%

'-----------------------------------------------------------------------------------------------------------------------------------------
function monta_condi (auxi) 
if condi<>"" then
 condi = condi & " and " & auxi
else
 condi = auxi
end if
end function
'
'
'-----------------------------------------------------------------------------------------------------------------------------------------
'
FUNCTION valores (vgru,vname,vsele,vorde)
if vorde="" then
 vorde = "desc1"
end if

orden="Select * from ECOINFORMAS_VALORES WHERE grupo='"& vgru &"' ORDER BY " & vorde
Set DSQL = Server.CreateObject ("ADODB.Recordset")
dSQL.Open orden,objConnection,adOpenKeyset
seleccion = vsele

if seleccion="X" then
 sele_sin="selected"
end if

%><select name=<%=vname%> class="campo" onchange="javascript:refresca();">
  <option value="">- Todos los estados -</option>
  <option <%=sele_sin%> value="X">- Sin finalizar -</option><%  
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


FUNCTION valores_tem (vgru,vname,vsele,vorde)
if vorde="" then
 vorde = "desc1"
end if

orden="Select * from ECOINFORMAS_VALORES WHERE grupo='"& vgru &"' ORDER BY " & vorde
Set DSQL = Server.CreateObject ("ADODB.Recordset")
dSQL.Open orden,objConnection,adOpenKeyset
seleccion = vsele

%><select name=<%=vname%> class="campo" onchange="javascript:refresca();">
  <option value="">- Todos los temas -</option><%
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

%>


<script>
<!--

function refresca () {
  param = 'estado_consulta='+document.asesora.estado_consulta.value+'&tema_consulta_fil='+document.asesora.tema_consulta_fil.value;
  location.href='asesora_paso1.asp?'+param;
}

function paginar (pag,act) {
  param = 'pag='+pag+'&act_pag='+act+'&estado_consulta=<%=estado_consulta%>&tema_consulta_fil=<%=tema_consulta_fil%>';
  location.href='asesora_paso1.asp?'+param;
}

function nuevo_asesoramiento() {
 window.open ('asesora_nuevo.asp','nuevo_asesora','scrollbars=1,status=0,toolbar=0,resizable=1,width=600,height=400');
}

function responder(id) {
 param = 'idconsulta='+id+'&asesor=<%=asesor%>&act_pag=<%=act_pag%>&estado_consulta_pri=<%=estado_consulta%>&tema_consulta_fil=<%=tema_consulta_fil%>';
 //alert (param);
 window.open ('asesora_respuesta.asp?'+param,'responder_asesora','scrollbars=1,status=0,toolbar=0,resizable=1,width=600,height=400');
}

function ver_pregunta(id) {
 param = 'idconsulta='+id+'&asesor=<%=asesor%>&act_pag=<%=act_pag%>&estado_consulta_pri=<%=estado_consulta%>';
 //alert (param);
 window.open ('asesora_ver_pregunta.asp?'+param,'responder_asesora','scrollbars=1,status=0,toolbar=0,resizable=1,width=600,height=400');
}

function ver_adjunto(id,ex) {
 window.open ('<%=ruta_upload_log%>ASESORA_'+id+ex,'ver_adjunto','scrollbars=1,status=0,toolbar=0,resizable=1,width=600,height=300');
}

-->
</script>
