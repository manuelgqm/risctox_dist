<!--#include virtual="config/checkAccess.asp"-->
<!--#include virtual="config/dbConnection.asp"-->
<!--#include virtual="lib/requestHelper.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
	<title>ISTAS: risctox</title>

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

	<link rel="stylesheet" type="text/css" href="../estructura.css">
	<link rel="stylesheet" type="text/css" href="../dn_estilos.css">
</head>

<body>
<div id="contenedor">
	<div id="sombra_arriba"></div>
  	<div id="sombra_lateral">
		<div id="caja">
		<!--#include file="header.asp"-->
		<div id="texto">
			
<div class="texto">
<!-- ################ CONTENIDO ###################### -->


<table width="100%" border="0">

<tr>

<td><p class=campo>Est&aacute;s en: <a href=index.asp?idpagina=550>Plataforma prevención de riesgo químico</a>&nbsp;&gt;&nbsp;<a href="dn_alternativas_portada.asp">BBDD Alternativas</a> &gt; <a href="dn_alternativas_sectores.asp">Sectores</a> &gt; Ficha de sector </p></td>

<td><input type="button" name="volver" class="boton2" value="Volver a la portada de Alternativas" onclick="window.location='dn_alternativas_portada.asp';"></td>

</tr>

</table>


<%
	id = requestClean("id")
	set rstl = objConnection2.Execute("select nombre from dn_alter_sectores where id =" &id)
	nombre=rstl("nombre")
	rstl.close
	set rstl=nothing
%>

	<!-- Datos de sustancia -->
	<p class=titulo3><%=nombre%></p>

<!-- FICHEROS DE ALTERNATIVAS 
 ficheros asociados al uso.  -->

	
<%
	set rstl=objConnection2.Execute("select dn_alter_ficheros.id, dn_alter_ficheros.titulo, dn_alter_ficheros.num_alternativa from dn_alter_ficheros_por_sectores JOIN dn_alter_ficheros ON dn_alter_ficheros_por_sectores.id_fichero=dn_alter_ficheros.id where dn_alter_ficheros.tema<>'Casos prácticos' and dn_alter_ficheros_por_sectores.id_sector =" &id & "ORDER BY dn_alter_ficheros.titulo")
	if rstl.eof then
		'response.write "No se encontraron ficheros"


	else
%>
<p class=titulo3>Documentos de alternativas</p>
	<table width="100%" class="tabla3">
		<tr>
			<td class="texto">
<%
		response.write "<ul>"
		nums_alternativas = ""

		do while not rstl.eof
			' Mostramos el documento si su número de alternativa no había salido antes

			' Lo buscamos siempre con una coma detrás

			'response.write instr(nums_alternativas, rstl(2)&",") & "<br/>"

			if (instr(nums_alternativas, rstl(2)&",") = 0) then

				response.write "<li><a href='dn_alternativas_ficha_fichero.asp?id_fichero=" &rstl(0)& "'>" &rstl(1)& "</a></li>"

				' Apuntamos el num_alternativa en la lista para no repetir

				if (nums_alternativas <> "") then

					nums_alternativas = nums_alternativas & rstl(2) & ","

				else

					nums_alternativas = rstl(2)&","

				end if

			end if


		rstl.movenext
		loop
		response.write "</ul>"


		rstl.close
		set rstl=nothing
%>
		</td>
		</tr>
	</table>
<%
	end if
%>
			

<!-- CASOS PRACTICOS asociados -->

	
<%

	' Ignoramos los que tengan numero de alternativa mostrado en documentos


	set rstl=objConnection2.Execute("select dn_alter_ficheros.id, dn_alter_ficheros.titulo, dn_alter_ficheros.num_alternativa from dn_alter_ficheros_por_sectores JOIN dn_alter_ficheros ON dn_alter_ficheros_por_sectores.id_fichero=dn_alter_ficheros.id where dn_alter_ficheros.tema='Casos prácticos' and dn_alter_ficheros_por_sectores.id_sector =" &id & " AND dn_alter_ficheros.num_alternativa NOT IN ("&nums_alternativas&"-1) ORDER BY dn_alter_ficheros.titulo")
	if rstl.eof then
		'response.write "No se encontraron Casos prácticos"


	else
%>
<p class=titulo3>Casos prácticos</p>
	<table width="100%" class="tabla3">
		<tr>
			<td class="texto">
<%
		response.write "<ul>"
		titulo_antiguo=""		

		do while not rstl.eof
			if (rstl(1) <> titulo_antiguo) then			

				response.write "<li><a href='dn_alternativas_ficha_fichero.asp?id_fichero=" &rstl(0)& "'>" &rstl(1)& "</a></li>"

				titulo_antiguo = rstl(1)

			end if
			rstl.movenext
		loop
		response.write "</ul>"
		rstl.close
		set rstl=nothing
%>
</td>
		</tr>
	</table>
<%
	end if
%>
			

<!-- ############ FIN DE CONTENIDO ################## -->

<br>
Esta página ha sido desarrollada por <a href="http://www.istas.ccoo.es/" target="_blank"><b>ISTAS</b></a> que es una Fundación de <a href="http://www.ccoo.es/" target="_blank"><font color="#FF0000"><b>CC.OO.</b></font></a><br>

				
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
			<img src="../imagenes/pie3.jpg" width="708" border="0" usemap="#Map3">
      			
    			
    		</div>
    	</div>
	<div id="sombra_abajo"></div>
</div>
</body>
</html>