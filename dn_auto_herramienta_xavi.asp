<!--#include file="dn_conexion.asp"--><!--#include file="adovbs.inc"--><!--#include file="dn_restringida.asp"--><!--#include file="dn_funciones_comunes.asp"--><!--#include file="dn_funciones_texto.asp"-->
<%maxComponentes = 20 ' Número máximo de componentes permitido por producto

' Consultamos numero de productos para el usuario
sql="select count(*) AS num_productos FROM dn_auto_productos WHERE id_ecogente="&session("id_ecogente")
set objRst=objConnection2.execute(sql)
if (objRst.eof) then
	num_productos = 0
else
	num_productos = objRst("num_productos")
end if
objRst.close()
set objRst=nothing%>

<%
	idpagina = 964	'--- Herramienta de Autoevaluación
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
 %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html lang="es" xmlns="http://www.w3.org/1999/xhtml"><head><title>ISTAS: risctox</title><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" /><meta name="Title" content="ECOinformas" /><meta name="Author" content="DABNE" /><meta name="description" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" /><meta name="Subject" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" /><meta name="Keywords" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" /><meta name="Language" content="Spanish" /><meta name="Revisit" content="15 days" /><meta name="Distribution" content="Global" /><meta name="Robots" content="All" /><link rel="stylesheet" type="text/css" href="estructura.css"><link rel="stylesheet" type="text/css" href="dn_estilos.css"><script type="text/javascript" src="dn_auto_scripts.js"></script><script type="text/javascript" src="dn_scripts.js"></script>

<script type="text/javascript" src="scripts/moo/prototype.lite.js"></script>
<script type="text/javascript" src="scripts/moo/moo.ajax.js"></script><script type="text/javascript" src="niftycube.js"></script><script type="text/javascript">window.onload=function(){Nifty("div.dn_ncc_cabecera","top"); Nifty("div.dn_ncc_pie","bottom"); }var maxComponentes = <%=maxComponentes%>;</script><script type="text/javascript">function despliegame (id){	change(id, 'muestra') 	change(id+'_producto', 'productodesp') 	change(id+'_plegar', 'muestra') }function pliegame (id){	change(id, 'oculta') 	change(id+'_producto', 'producto') 	change(id+'_plegar', 'oculta') }function change(id, newClass) {	identity=document.getElementById(id);	identity.className=newClass;}</script></head><body>
<!--#include file="dn_detecta_navegador.asp"--><div id="contenedor">	<div id="sombra_arriba"></div>  	<div id="sombra_lateral">		<div id="caja">		<!--#include file="dn_cabecera.asp"-->		<div id="texto">			<div class="texto"><p class=titulo3>Herramienta de Autoevaluación</p><p><strong>¡Bienvenid@ a la Herramienta de Autoevaluación!</strong><br/>Desde este apartado dispones de tu propio espacio donde almacenar las especificaciones de los productos que empleas. Puedes añadir nuevos productos o consultar tu lista de productos creados anteriormente, para ver sus características, evaluar su peligrosidad y comparar unos con otros.</p>
<p align="center"><strong><a href="dn_auto_portada.asp">&lt;&lt;&lt; Volver a la portada de Autoevaluación</a></strong></p><div id="main">


<% if (num_productos > 0) then %>
	<!-- ####### INICIO CESTA #################################################### -->	<div id="dn_auto_cesta">		<div id="dn_auto_cesta_cabecera" class="dn_ncc_cabecera">
			<h2 class="dn_cabecera_2"><img src="imagenes/icono_autoevalua.gif" hspace="0" vspace="0" align="absmiddle"  /> lista de productos <a href="javascript:mostrar_ayuda('Cesta');"><img src="imagenes/fam_istas/help.gif" hspace="2" vspace="2" align="absmiddle" /></a></h2>
		</div>
		<div id="dn_auto_cesta_cuerpo" class="dn_ncc_cuerpo">			<!--#include file="dn_auto_cesta_include.asp"-->		</div>

		<div id="dn_auto_evaluador_cuerpo" class="dn_ncc_cuerpo"></div>		
		<div id="dn_auto_cesta_pie" class="dn_ncc_pie">			&nbsp;		</div>	</div>	<!-- ####### FIN CESTA #################################################### -->
<% end if %>


	<!-- ####### INICIO PRODUCTO ############################################## -->	<div id="dn_auto_producto_cabecera" class="dn_ncc_cabecera">
		<h2 class="dn_cabecera_2"><img src="imagenes/icono_producto.gif" hspace="0" vspace="0" align="absmiddle"  /> nuevo producto <a href="javascript:mostrar_ayuda('NUEVO PRODUCTO');"><img src="imagenes/fam_istas/help.gif" hspace="2" vspace="2" align="absmiddle" /></a></h2>
		
	</div>
	<div id="dn_auto_producto_cuerpo" class="dn_ncc_cuerpo">
	<p>Rellena  esta ficha con la información sobre la composición y características físicas de la FDS.</p>
	<fieldset>
		<legend class="dn_cabecera_3"><img src="imagenes/fam_istas/database_add.gif" hspace="2" vspace="2" align="absmiddle" />datos del producto</legend>

	<p>Indica todos los datos de que dispongas acerca del producto y su uso.</p>
	<form id="form_prod" name="form_prod" action="dn_auto_guardar_producto.asp" method="post">
	<input type="hidden" id="num_componentes" name="num_componentes" value="0">
		<!-- DATOS DE PRODUCTO -->
		<div id="producto">
		<table border="0">
			<tr>
				<th align="left">Nombre comercial</th>
				<th align="left"><a href="javascript:frasesr('prod_frases_r');">Frases R</a> <a href="javascript:mostrar_ayuda('FRASES R PRODUCTO');"><img src="imagenes/fam_istas/help.gif" hspace="2" vspace="2" align="absmiddle" /></a></th>	
			</tr>
			<tr>
				<td>
					<input type="text" size="<%=len_auto_prod_nombre%>" maxlength="750" id="prod_nombre" name="prod_nombre"  class="campo"/>
				</td>
				<td><input type="text" size="25" maxlength="50" id="prod_frases_r" name="prod_frases_r"  class="campo"/></td>
			</tr>
			<tr>
				<th align="left" colspan="2">
					¿En qué tipo de proceso se emplea? <a href="javascript:mostrar_ayuda('TIPO DE PROCESO');"><img src="imagenes/fam_istas/help.gif" hspace="2" vspace="2" align="absmiddle" /></a>
				</th>
			</tr>
			<tr>
				<td colspan="2">
					<select name="prod_cod_proceso" class="campo">
						<% dameOpciones "cod", "nombre", "dn_auto_procesos", "cod", "", "", "", "" %>
					</select>
				</td>
			</tr>
		</table>

		<table border="0" width="100%">
			<tr>
				<th>Estado físico <a href="javascript:mostrar_ayuda('ESTADO FISICO');"><img src="imagenes/fam_istas/help.gif" hspace="2" vspace="2" align="absmiddle" /></a></th>
				<th>Presión de vapor <a href="javascript:mostrar_ayuda('PRESION DE VAPOR');"><img src="imagenes/fam_istas/help.gif" hspace="2" vspace="2" align="absmiddle" /></a></th>
			</tr>
			<tr>
				<td>
					<select name="prod_cod_estado" class="campo">
						<% dameOpciones "cod", "nombre", "dn_auto_estados", "cod", "", "", "", "" %>
					</select>
				</td>
				<td>
					<select name="prod_cod_presion" class="campo">
						<% dameOpciones "cod", "nombre", "dn_auto_presiones", "cod", "", "", "", "" %>
					</select>
				</td>
			</tr>
			<tr>
				<th>Temperatura de evaporación <a href="javascript:mostrar_ayuda('TEMPERATURA DE EVAPORACION');"><img src="imagenes/fam_istas/help.gif" hspace="2" vspace="2" align="absmiddle" /></a></th>
				<th>Inflamabilidad <a href="javascript:mostrar_ayuda('INFLAMABILIDAD');"><img src="imagenes/fam_istas/help.gif" hspace="2" vspace="2" align="absmiddle" /></a></th>
			</tr>
			<tr>
				<td>
					<select name="prod_cod_temperatura" class="campo">
						<% dameOpciones "cod", "nombre", "dn_auto_temperaturas", "cod", "", "", "", "" %>
					</select>
				</td>
				<td>
					<select name="prod_cod_inflamabilidad" class="campo">
						<% dameOpciones "cod", "nombre", "dn_auto_inflamabilidades", "cod", "", "", "", "" %>
					</select>	
				</td>
			</tr>
		</table>
		</fieldset>

		<!-- DATOS DE COMPONENTES. UN DIV PARA CADA UNO, TODOS METIDOS EN CONTENEDOR -->

		<!-- INICIO COMPONENTES -->
		<%
		for contador = 1 to maxComponentes
		%>
				<!-- INICIO COMPONENTE <%= contador %> -->
				<div id="tabla_comp_<%=contador%>" class="componente oculta">
				<br/>
				<fieldset>
					<legend class="dn_cabecera_3"><img src="imagenes/fam_istas/database_add.gif" hspace="2" vspace="2" align="absmiddle" />datos del componente <%=contador%> <a href="javascript:mostrar_ayuda('DATOS DEL COMPONENTE');"><img src="imagenes/fam_istas/help.gif" hspace="2" vspace="2" align="absmiddle" /></a></legend>

				<p>Introduce los datos que conozcas acerca de cada componente del producto. Si conoces su número identificativo o nombre, introduce los primeros caracteres y te ayudaremos buscando en nuestra base de datos.</p>

				<table border="0" width="100%">
					<tr>
						<th align="left">Tipo de número</th>
						<th align="left">Número identificativo <a href="javascript:mostrar_ayuda('NUMERO IDENTIFICATIVO');"><img src="imagenes/fam_istas/help.gif" hspace="2" vspace="2" align="absmiddle" /></a></th>
						<th align="center">Concentración <a href="javascript:mostrar_ayuda('CONCENTRACION');"><img src="imagenes/fam_istas/help.gif" hspace="2" vspace="2" align="absmiddle" /></a></th>
					</tr>

					<tr>
						<td>
							<select name="comp<%=contador%>_numero_tipo" id="comp<%=contador%>_numero_tipo" onfocus="busca_sustancia_live('comp<%=contador%>')" onblur="busca_sustancia_live_stop()" class="campo">
								<option value="cas">Número CAS</option>
								<option value="ce_einecs">Número CE EINECS</option>
								<option value="ce_elincs">Número CE ELINCS</option>
								<option value="rd">Número RD</option>
							</select>
						</td>
						<td><input type="text" size="25" maxlength="50" id="comp<%=contador%>_numero" name="comp<%=contador%>_numero" onfocus="busca_sustancia_live('comp<%=contador%>')" onblur="busca_sustancia_live_stop()"  class="campo"/> <!--<input type="button" class="boton2" id="comp<%=contador%>_buscar" name="comp<%=contador%>_buscar" onclick="busca_sustancia('comp<%=contador%>')" value="Buscar" />--></td>
						<td align="center"><input type="text" size="4" maxlength="10" name="comp<%=contador%>_cod_porcentaje"  class="campo"/> %</td>
					</tr>
				</table>

				<table border="0" width="100%">
					<tr>
						<th align="left">
							Nombre
						</th>
						<th>
							<a href="javascript:frasesr('comp<%=contador%>_frases_r');">Frases R</a> <a href="javascript:mostrar_ayuda('FRASES R COMPONENTE');"><img src="imagenes/fam_istas/help.gif" hspace="2" vspace="2" align="absmiddle" /></a>
						</th>
					</tr>
					<tr>
						<td>
							<input type="text" size="50" maxlength="750" id="comp<%=contador%>_nombre" name="comp<%=contador%>_nombre" onfocus="busca_sustancia_live('comp<%=contador%>')" onblur="busca_sustancia_live_stop()"  class="campo"/>
						</td>
						<td>
							<input type="text" size="25" maxlength="50" id="comp<%=contador%>_frases_r" name="comp<%=contador%>_frases_r"  class="campo"/>
						</td>				
					</tr>
				</table>

				<div id="busqueda_comp<%=contador%>"></div>
				</fieldset>
				<!-- FIN COMPONENTE <%= contador %> -->
			</div>
		<%
		next
		%>
	

		<!-- BOTONES AÑADIR / ELIMINAR ÚLTIMO COMPONENTE -->
		<hr/>
		<div id="botones_componentes" class="componente">
		<center><input type="button" class="boton2" name="nuevo_componente" value="Añadir otro componente" onclick="anadir_componente(<%=maxComponentes%>)" /> <input type="button" class="boton2" name="boton_eliminar_componente" id="boton_eliminar_componente" value="Eliminar componente 1" onclick="eliminar_componente()" /></center>
		</div>

			<hr/><center><input type="button" class="boton2" name="enviar" value="Guardar producto" onclick="validarProducto();" /></center>

		</div>

		<!-- FIN BOTONES AÑADIR / ELIMINAR ULTIMO COMPONENTE -->
	
</form>

<script language="JavaScript">
// Creamos inicialmente un componente
anadir_componente(<%=maxComponentes%>);
</script>



	</div>
	<div id="dn_auto_producto_pie" class="dn_ncc_pie">		&nbsp;	</div>
	<!-- ####### FIN PRODUCTO ################################################## -->

	<p>Si desea obtener más información toxicológica de los productos utilizados y conocer posibles Alternativas a los mismos, puede recurrir a nuestras bases de datos <strong><a href="http://www.istas.net/ecoinformas/web/index.asp?idpagina=575">RISCTOX</a></strong> y <strong><a href="http://www.istas.net/ecoinformas/web/index.asp?idpagina=576">Alternativas</a></strong>.</p>

	<p align="center">Esta página ha sido desarrollada por <strong><a href="http://www.istas.ccoo.es/">ISTAS</a></strong> que es una Fundación de <strong><a href="http://www.ccoo.es/">CC.OO.</a></strong><br/><br/></p>
</div>			</div></div>									<map name="Map1" id="Map1">            		<area shape="rect" coords="307,38,399,104" href="http://www.fundacion-biodiversidad.es" target="_blank" alt="Fundación Biodiversidad" />            		<area shape="rect" coords="400,38,546,104" href="http://www.istas.ccoo.es" target="_blank" alt="Instituto Sindical de Trabajo, Ambiente y Salud" />      			<area shape="rect" coords="547,38,701,104" href="http://www.mtas.es/UAFSE/default.htm" target="_blank" alt="Fondo Social Europeo" />      			</map>			<map name="Map2" id="Map2">            		<area shape="rect" coords="300,8,392,66" href="http://www.fundacion-biodiversidad.es" target="_blank" alt="Fundación Biodiversidad" />            		<area shape="rect" coords="393,8,539,66" href="http://www.istas.ccoo.es" target="_blank" alt="Instituto Sindical de Trabajo, Ambiente y Salud" />      			<area shape="rect" coords="540,8,694,66" href="http://www.mtas.es/UAFSE/default.htm" target="_blank" alt="Fondo Social Europeo" />      			</map>      			<map name="Map3" id="Map3">            		<area shape="rect" coords="300,18,392,80" href="http://www.fundacion-biodiversidad.es" target="_blank" alt="Fundación Biodiversidad" />            		<area shape="rect" coords="393,18,539,80" href="http://www.istas.ccoo.es" target="_blank" alt="Instituto Sindical de Trabajo, Ambiente y Salud" />      			<area shape="rect" coords="540,18,694,80" href="http://www.mtas.es/UAFSE/default.htm" target="_blank" alt="Fondo Social Europeo" />      			</map>			<img src="imagenes/pie3.jpg" width="708" border="0" usemap="#Map3">      			    			    		</div>    	</div>	<div id="sombra_abajo"></div></div></body></html><%cerrarconexion%>
