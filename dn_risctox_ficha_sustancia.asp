<%@ LANGUAGE="VBSCRIPT" LCID="1034" CODEPAGE="65001"%>
<!--#include file="dn_restringida.asp"-->
<!--#include file="config/dbConnection.asp"-->
<!--#include file="lib/dn_funciones_texto_utf-8.asp"-->
<!--#include file="lib/dn_funciones_comunes_utf-8.asp"-->
<!--#include file="lib/class/SubstanceClass.asp"-->
<!--#include file="lib/visitsRecorder.asp"-->
<!--#include file="lib/urlManipulations.asp"-->

<%
Response.ContentType = "text/html"
Response.AddHeader "Content-Type", "text/html;charset=UTF-8"
Response.CodePage = 65001
Response.CharSet = "UTF-8"
idpagina = 627
call recordVisit(idpagina)
errores = ""

id_sustancia = obtainSanitizedQueryParameter("id_sustancia")

set mySubstance = new SubstanceClass
mySubstance.find id_sustancia, "es", objConnection2
Set substance = mySubstance.fields
if (substance.Count = 0 ) then errores = "No se ha encontrado la sustancia indicada"

dim NEUROTOXICO_LISTS : NEUROTOXICO_LISTS = array("neurotoxico", "neurotoxico_rd", "neurotoxico_danesa", "neurotoxico_nivel")

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ISTAS: risctox</title>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<meta name="Title" content="Risctox" />
<meta name="Author" content="SPL Sistemas de Información, SL - www.spl-ssi.com" />
<meta name="description" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Subject" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Keywords" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Language" content="Spanish" />
<meta name="Revisit" content="15 days" />
<meta name="Distribution" content="Global" />
<meta name="Robots" content="All" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />

<link rel="stylesheet" type="text/css" href="estructura.css">
<link rel="stylesheet" type="text/css" href="dn_estilos.css">

<script src="scripts/prototype/prototype.js" type="text/javascript"></script>
<script src="scripts/scriptaculous/scriptaculous.js" type="text/javascript"></script>
<script type="text/javascript">
function toggle(id_objeto, id_imagen)
{
    if (Element.visible(id_objeto))
    {
      $(id_imagen).src="imagenes/desplegar.gif";
    }
    else
    {
      $(id_imagen).src="imagenes/plegar.gif";
    }
    new Effect.toggle(id_objeto,"appear");
}

function toggle_texto(id_objeto, texto)
{
    new Effect.toggle(id_objeto,"appear");
}
</script>

</head>
<body>
<div id="contenedor">
	<div id="sombra_arriba"></div>
  	<div id="sombra_lateral">
		<div id="caja">
		<!--#include file="dn_cabecera.asp"-->
		<div id="texto">

<div class="texto">
<!-- ################ CONTENIDO ###################### -->

<table width="100%" border="0">
<tr>
<td><p class=campo>Est&aacute;s en: <a href="dn_risctox_buscador.asp">bbdd risctox</a> &gt; ficha de sustancia</p></td>
<td align="right"><input type="button" name="volver" class="boton2" value="Nueva búsqueda" onClick="window.location='dn_risctox_buscador.asp';"></td>
</tr>
</table>
<br />
<div id="ficha">
	<!-- ################ Identificacion de la sustancia ###################### -->
	<table width="100%" cellpadding=5>
		<tr>
			<td>
				<a name="identificacion"></a><img src="imagenes/risctox01.gif" alt="identificación de la sustancia" width="255" height="32" />
			</td>
			<td align="right">
				<a href="#"><img src="imagenes/subir.gif" border=0 alt=subir></a>
			</td>
		</tr>
	</table>

	<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
		<!-- ################ Identificación ###################### -->

		<!-- 1.- Datos de sustancia -->
		<% ap1_identificacion() %>
	</table>

	<div style="height:3pt"></div>
		<!-- 2.1- Clasificación -->
		<% ap2_clasificacion() %>

	<br />
	<div style="height:3pt"></div>

		<!-- 2.2- Clasificación RD1272-->
		<% ap2_clasificacion_rd1272() %>

	<br />
	<div style="height:3pt"></div>

		<!-- Valores límite -->
		<% ap2_clasificacion_vl("secc-vla") %>

	<br />
</div>
<!-- fin div ficha -->

<!-- 3.- Riesgos -->
<% ap3_riesgos() %>

<!-- 4.- Normativa -->
<% ap4_normativa_ambiental() %>
<% ap4_normativa_salud_laboral() %>
<% ap4_normativa_restriccion_prohibicion() %>

<!-- 5.- Alternativas relacionadas -->
<% ap5_alternativas() %>

<!-- 6.- Sectores en los que se utiliza -->
<% ap6_sectores() %>

<!-- ############ FIN DE CONTENIDO ################## -->
<br />
<center>
<input type="button" name="imprimir" class="boton2" value="Imprimir ficha" onClick="window.print();">
<input type="button" name="enviar" class="boton2" value="Enviar ficha de sustancia" onClick="onclick=window.open('dn_recomendar.asp?id=<%=id_sustancia%>','recomendar','width=500,height=230,scrollbars=yes,resizable=yes')">
<input type="button" name="volver" class="boton2" value="Nueva búsqueda" onClick="window.location='dn_risctox_buscador.asp';">
</center>

<br>
<br>
Esta página ha sido desarrollada por <a href="http://www.istas.ccoo.es/" target="_blank"><b>ISTAS</b></a> que es una Fundación de <a href="http://www.ccoo.es/" target="_blank"><font color="#FF0000"><b>CC.OO.</b></font></a><br>


				</div>
				<p>&nbsp;</p>
			</div>


			<img src="imagenes/pie_risctox.gif" width="708" border="0">



    		</div>
    	</div>
	<div id="sombra_abajo"></div>
</div>

<!--#include file="../cookie_accept.asp" -->
</body>
</html>

<%
cerrarconexion
%>

<%
function formatHtmlGlossaryLinksString(elements, glossaryType)
	dim result : result = ""
	if not isArray(elements) then
		formatHtmlGlossaryLinksString = result
		exit function
	end if

	dim i
	dim descriptionLink
	dim element
	dim elementsLastId : elementsLastId = ubound(elements)
	for i = 0 to elementsLastId
		set element = elements(i)
		descriptionLink = getDescriptionLink(element)
		result = result & element.Item("name") & descriptionLink
		if not(i + 1 > elementsLastId) then result = result & ", "
	next

	formatHtmlGlossaryLinksString = result
end function

function getDescriptionLink(element)
	dim result : result = ""

	if element.Item("description") = "" then
		getDescriptionLink = result
		exit function
	end if

	result = " <a onclick=window.open('dn_glosario.asp?tabla=" & glossaryType & "&id=" & element.Item("item_id") & "','def','width=500,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a>"

	getDescriptionLink = result
end function

function formatHtmlCompaniesLinksString(companies)
	dim result : result = ""

	if not isArray(companies) then
		formatHtmlCompaniesLinksString = result
		exit function
	end if

	dim i
	dim companiesLastId : companiesLastId = ubound(companies)
	dim name, id, company
	for i = 0 to companiesLastId
		set company = companies(i)
		id = company.Item("item_id")
		name = company.Item("name")
		result = result & "<a onclick=window.open('dn_risctox_ficha_compania.asp?id=" & id & "','comp','width=600,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'>" &name & "</a>"
		if not(i + 1 > companiesLastId) then result = result & ", "
	next

	formatHtmlCompaniesLinksString = result
end function

function formatHtmlUnorderedList(elements)
	dim list, i
	list = ""

	if not isArray(elements) then formatHtmlUnorderedList = list

	list = list & "<ul>"
	for i = 0 to Ubound(elements)
		list = list & "<li>" & h(espaciar(elements(i))) & "</li>"
	next
	list = list & "</ul>"

	formatHtmlUnorderedList = list
end function

sub ap1_identificacion()
%>
	<tr>
		<td class="subtitulo3" align="right" valign="top">
			<a onclick=window.open('ver_definicion.asp?id=82','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a> Nombre:
		</td>
		<td class="texto" valign="middle">
			<b><%=espaciar(substance.Item("nombre"))%></b>
		</td>
	</tr>

	<%
	if ( ubound(substance.Item("sinonimos")) > -1 ) then
	%>
		<tr>
			<td class="subtitulo3" align="right" valign="top">
				<a onclick=window.open('ver_definicion.asp?id=83','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a> Sinónimos:
			</td>
			<td class="texto" valign="middle">
				<%=formatHtmlUnorderedList(substance.Item("sinonimos"))%>
			</td>
		</tr>
	<%
	end if ' hay sinonimos?
	%>

	<%
		nombre_comercial = dameNombreComercial(id_sustancia)
		if (nombre_comercial <> "") then
	%>
		<tr>
			<td class="subtitulo3" align="right" valign="top">
				Nombre comercial:
			</td>
			<td class="texto" valign="middle">
				<%=nombre_comercial%>
			</td>
		</tr>
	<% end if ' hay nombre comercial? %>

	<% if (substance.Item("num_cas") <> "") or (substance.Item("num_ce_einecs") <> "") or (substance.Item("num_ce_elincs") <> "") then %>
		<tr>
			<td class="subtitulo3" align="right" valign="top">
				Números de Identificación:
			</td>
			<td class="texto" valign="middle">
				<% if (substance.Item("num_cas") <> "") then response.write "<a onclick=window.open('ver_definicion.asp?id=84','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a> <b>CAS</b>: "&substance.Item("num_cas")&"<br/>" %>
				<% if (substance.Item("cas_alternativos") <> "") then response.write "<a onclick=window.open('ver_definicion.asp?id=84','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a> <b>CAS alternativos</b>: "&substance.Item("cas_alternativos") &"<br/>" %>
				<%
					if (substance.Item("num_ce_einecs") <> "") then
						'Sergio, si empieza por 4 y num_ce_elincs<>'' muestro el num_ce_elincs
						if (mid(substance.Item("num_ce_einecs"),1,1)="4" and substance.Item("num_ce_elincs") <> "") then
							response.write "<a onclick=window.open('ver_definicion.asp?id=85','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a> <b>CE ELINCS</b>: "&substance.Item("num_ce_elincs")&"<br/>"
						else
						response.write "<a onclick=window.open('ver_definicion.asp?id=85','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a> <b>CE EINECS</b>: "&substance.Item("num_ce_einecs")&"<br/>"
						end if
					elseif (substance.Item("num_ce_elincs") <> "") then
						response.write "<a onclick=window.open('ver_definicion.asp?id=85','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a> <b>CE ELINCS</b>: "&substance.Item("num_ce_elincs")&"<br/>"
					end if
				%>
			</td>
		</tr>
	<% end if ' hay numeros? %>

	<%
		grupos = formatHtmlGlossaryLinksString(substance.item("grupos"), "grupos")
		if (grupos <> "") then
	%>
		<tr>
			<td class="subtitulo3" align="right" valign="top">
				Grupos:
			</td>
			<td class="texto" valign="middle">
				<%=grupos%>
			</td>
		</tr>
	<% end if ' hay grupos? %>

	<%
		usos = formatHtmlGlossaryLinksString(substance.Item("aplicaciones"), "usos")
		if (usos <> "") then
	%>
		<tr>
			<td class="subtitulo3" align="right" valign="top">
				Usos:
			</td>
			<td class="texto" valign="middle">
				<%=usos%>
			</td>
		</tr>
	<% end if %>



	<%
		if (substance.Item("num_icsc") <> "") then
	%>
		<tr>
			<td class="subtitulo3" align="right" valign="top">
				 Ficha Internacional de Seguridad Química (<a onClick="window.open('ver_definicion.asp?id=<%=dame_id_definicion("INSHT")%>', 'def', 'width=300,height=200,scrollbars=yes,resizable=yes')" class="subtitulo3">INSHT</a>)
			</td>
			<td class="texto" valign="middle">
          <%

            array_icsc=split(substance.Item("num_icsc"), "@")

            for i=0 to ubound(array_icsc)
            	substance.Item("num_icsc") = cstr(array_icsc(i))
            	if len(substance.Item("num_icsc"))=4 then
            		centena_icsc = mid(substance.Item("num_icsc"),1,2)
            		icsc_max = cstr(clng(centena_icsc&"01"))
            		if icsc_max="1" then icsc_max="0"
            		icsc_min = cstr(clng(centena_icsc)+1) & "00"
            	end if

          %>

              <a href="http://www.insht.es/InshtWeb/Contenidos/Documentacion/FichasTecnicas/FISQ/Ficheros/<%=icsc_max%>a<%=icsc_min%>/nspn<%= array_icsc(i) %>.pdf" target="_blank"><%= array_icsc(i) %></a>

          <%

            next

          %>

			</td>
		</tr>
	<% end if %>



	<%
		companias = formatHtmlCompaniesLinksString(substance.Item("compañias"))
	%>

	<% if (substance.Item("nombre_ing") <> "") or (substance.Item("num_rd") <> "") or (substance.Item("formula_molecular") <> "") or (substance.Item("estructura_molecular") <> "") or (substance.Item("notas_xml") <> "") or (companias <> "") then %>
		<tr>
			<td class="subtitulo3" align="right" valign="top" width="35%">
				Más información <% plegador "secc-masinformacion", "img-masinformacion" %>
			</td>
			<td class="texto" valign="middle" id="secc-masinformacion" style="display:none">



        <% if (substance.Item("nombre_ing") <> "") then
            array_nombres_ingleses = split(substance.Item("nombre_ing"), "@")
            if (ubound(array_nombres_ingleses) > 0) then
        %>
              <b>Nombres en inglés</b>:<br/>
              <ul>
                <% for i=0 to ubound(array_nombres_ingleses) %>
                  <li><%= espaciar(array_nombres_ingleses(i)) %></li>
                <% next %>
              </ul>
        <%
            else
        %>
              <b>Nombre inglés</b>: <%= espaciar(substance.Item("nombre_ing")) %><br/>
        <%
            end if
           end if %>

				<% if (substance.Item("num_rd") <> "") then response.write "<a onclick=window.open('ver_definicion.asp?id=86','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a> <b>Nº &iacute;ndice</b>: "&substance.Item("num_rd")&"<br/>" %>
				<% if (substance.Item("formula_molecular") <> "") then response.write "<b>Fórmula molecular</b>: "&substance.Item("formula_molecular")&"<br/>" %>
				<% if (substance.Item("estructura_molecular") <> "") then response.write "<b>Estructura molecular</b>:<br /><img src='../gestion/estructuras/"&substance.Item("estructura_molecular")&"' /><br/>" %>

				<% if (substance.Item("notas_xml") <> "") then %>
          <a onClick="window.open('ver_definicion.asp?id=<%=dame_id_definicion("ECB")%>', 'def', 'width=300,height=200,scrollbars=yes,resizable=yes')" style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>
          <b>Notas ECB</b>: <%= espaciar(substance.Item("notas_xml")) %> <br />
        <% end if %>

        <% if (companias <> "") then %>
          <b>Compañías distribuidoras</b>: <%= companias %>
        <% end if %>
			</td>
		</tr>
	<% end if
%>
	<tr>
		<td valign="top" colspan="2">
			<!-- Lista negra -->

			<% ap2_clasificacion_lista_negra(mySubstance) %>
		</td>
	</tr>
<%
end sub ' ap1_identificacion

' ###################################################################################

sub ap2_clasificacion()
	' Solo mostramos este apartado si hay información para él
	if ((substance.Item("simbolos") <> "") or (substance.Item("clasificacion_1") <> "") or (substance.Item("clasificacion_2") <> "") or (substance.Item("clasificacion_3") <> "") or (substance.Item("clasificacion_4") <> "") or (substance.Item("clasificacion_5") <> "") or (substance.Item("clasificacion_6") <> "") or (substance.Item("clasificacion_7") <> "") or (substance.Item("clasificacion_8") <> "") or (substance.Item("clasificacion_9") <> "") or (substance.Item("clasificacion_10") <> "") or (substance.Item("clasificacion_11") <> "") or (substance.Item("clasificacion_12") <> "") or (substance.Item("clasificacion_13") <> "") or (substance.Item("clasificacion_14") <> "") or (substance.Item("clasificacion_15") <> "") or (substance.Item("frases_r_danesa") <> "") or (substance.Item("notas_rd_363") <> "") or (substance.Item("conc_1") <> "") or (substance.Item("eti_conc_1") <> "") or (substance.Item("conc_2") <> "") or (substance.Item("eti_conc_2") <> "") or (substance.Item("conc_3") <> "") or (substance.Item("eti_conc_3") <> "") or (substance.Item("conc_4") <> "") or (substance.Item("eti_conc_4") <> "") or (substance.Item("conc_5") <> "") or (substance.Item("eti_conc_5") <> "") or (substance.Item("conc_6") <> "") or (substance.Item("eti_conc_6") <> "") or (substance.Item("conc_7") <> "") or (substance.Item("eti_conc_7") <> "") or (substance.Item("conc_8") <> "") or (substance.Item("eti_conc_8") <> "") or (substance.Item("conc_9") <> "") or (substance.Item("eti_conc_9") <> "") or (substance.Item("conc_10") <> "") or (substance.Item("eti_conc_10") <> "") or (substance.Item("conc_11") <> "") or (substance.Item("eti_conc_11") <> "") or (substance.Item("conc_12") <> "") or (substance.Item("eti_conc_12") <> "") or (substance.Item("conc_13") <> "") or (substance.Item("eti_conc_13") <> "") or (substance.Item("conc_14") <> "") or (substance.Item("eti_conc_14") <> "") or (substance.Item("conc_15") <> "") or (substance.Item("eti_conc_15") <> "") ) then

%>
	<!-- ################ Clasificación ###################### -->
	<table id="tabla_clasificacionm" class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
  <tr>
		<td class="celdaabajo" colspan="2" align="center">
			<table cellpadding=0 cellspacing=0 width="100%" border="0"><tr><td width="100%" class="titulo3" align="left"><a onclick=window.open('ver_definicion.asp?id=87','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> CLASIFICACIÓN (RD 363/1995)
			<a href="javascript:toggle('secc-clasificacion-363', 'img-mas_clasificacion-363');"><img src="imagenes/desplegar.gif" align="absmiddle" id="img-mas_clasificacion-363" alt="Pulse para desplegar la información" title="Pulse para desplegar la información" /></a>
			</td></tr></table>
		</td>
	</tr>
	<!-- Simbolos y frases R -->
	<tr><td>
		<table id="secc-clasificacion-363" style="display:none">
			<tr>
				<td valign="top">
					<% ap2_clasificacion_simbolos() %>
				</td>
				<td valign="top">
					<% ap2_clasificacion_frases_r(substance) %>
					<%

		        if mySubstance.hasFrasesRdanesa() then
		          ap2_clasificacion_frases_r_danesa(substance)
		        end if
		      %>
					<% ap2_clasificacion_frases_s() %>
					<% ap2_clasificacion_notas() %>
					<% ap2_clasificacion_etiquetado() %>
				</td>
			</tr>


		</table>
		</td>
		</tr>
	</table>
<%
	end if
end sub ' ap2_clasificacion

sub ap2_clasificacion_rd1272()
	' Solo mostramos este apartado si hay información para él
	if ((substance.Item("simbolos_rd1272") <> "") or (substance.Item("clasificacion_rd1272_1") <> "") or (substance.Item("clasificacion_rd1272_2") <> "") or (substance.Item("clasificacion_rd1272_3") <> "") or (substance.Item("clasificacion_rd1272_4") <> "") or (substance.Item("clasificacion_rd1272_5") <> "") or (substance.Item("clasificacion_rd1272_6") <> "") or (substance.Item("clasificacion_rd1272_7") <> "") or (substance.Item("clasificacion_rd1272_8") <> "") or (substance.Item("clasificacion_rd1272_9") <> "") or (substance.Item("clasificacion_rd1272_10") <> "") or (substance.Item("clasificacion_rd1272_11") <> "") or (substance.Item("clasificacion_rd1272_12") <> "") or (substance.Item("clasificacion_rd1272_13") <> "") or (substance.Item("clasificacion_rd1272_14") <> "") or (substance.Item("clasificacion_rd1272_15") <> "") or (substance.Item("conc_rd1272_1") <> "") or (substance.Item("eti_conc_rd1272_1") <> "") or (substance.Item("conc_rd1272_2") <> "") or (substance.Item("eti_conc_rd1272_2") <> "") or (substance.Item("conc_rd1272_3") <> "") or (substance.Item("eti_conc_rd1272_3") <> "") or (substance.Item("conc_rd1272_4") <> "") or (substance.Item("eti_conc_rd1272_4") <> "") or (substance.Item("conc_rd1272_5") <> "") or (substance.Item("eti_conc_rd1272_5") <> "") or (substance.Item("conc_rd1272_6") <> "") or (substance.Item("eti_conc_rd1272_6") <> "") or (substance.Item("conc_rd1272_7") <> "") or (substance.Item("eti_conc_rd1272_7") <> "") or (substance.Item("conc_rd1272_8") <> "") or (substance.Item("eti_conc_rd1272_8") <> "") or (substance.Item("conc_rd1272_9") <> "") or (substance.Item("eti_conc_rd1272_9") <> "") or (substance.Item("conc_rd1272_10") <> "") or (substance.Item("eti_conc_rd1272_10") <> "") or (substance.Item("conc_rd1272_11") <> "") or (substance.Item("eti_conc_rd1272_11") <> "") or (substance.Item("conc_rd1272_12") <> "") or (substance.Item("eti_conc_rd1272_12") <> "") or (substance.Item("conc_rd1272_13") <> "") or (substance.Item("eti_conc_rd1272_13") <> "") or (substance.Item("conc_rd1272_14") <> "") or (substance.Item("eti_conc_rd1272_14") <> "") or (substance.Item("conc_rd1272_15") <> "") or (substance.Item("eti_conc_rd1272_15") <> "") ) then

%>
	<!-- ################ Clasificación ###################### -->
	<table id="tabla_clasificacionm_rd1272" class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
  <tr>
		<td class="celdaabajo" colspan="2" align="center">
			<table cellpadding=0 cellspacing=0 width="100%" border="0"><tr><td width="100%" class="titulo3" align="left"><a onclick=window.open('ver_definicion.asp?id=280','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> CLASIFICACIÓN Y ETIQUETADO (Reglamento 1272/2008)
			<a href="javascript:toggle('secc-clasificacion-rd1272', 'img-mas_clasificacion-rd1272');"><img src="imagenes/desplegar.gif" align="absmiddle" id="img-mas_clasificacion-rd1272" alt="Pulse para desplegar la información" title="Pulse para desplegar la información" /></a>
			</td></tr></table>
		</td>
	</tr>
	<!-- Simbolos y frases H -->
	<tr><td>
		<table id="secc-clasificacion-rd1272" style="display:none">
			<tr>
				<td valign="top">
					<% ap2_clasificacion_simbolos_rd1272() %>
				</td>
				<td valign="top">
					<% ap2_clasificacion_frases_h() %>
					<br>
					<% ap2_clasificacion_notas_rd1272() %>
					<br>
					<% ap2_clasificacion_etiquetado_rd1272() %>
				</td>
			</tr>

		</table>
	</td></tr>
	</table>
<%
	end if
end sub ' ap2_clasificacion

' ##################################################################################

sub ap2_clasificacion_simbolos()
	if (substance.Item("simbolos") <> "") then
%>
		<p id="ap2_clasificacion_simbolos_titulo" class="ficha_titulo_2">Símbolos</p>
		<p id="ap2_clasificacion_simbolos_cuerpo" class="texto" align="center">
<%
		' Tiene símbolos, muestro cada uno
		substance.Item("simbolos") = replace(substance.Item("simbolos"), ",", ";")
		array_simbolos = split(substance.Item("simbolos"), ";")
		for i=0 to ubound(array_simbolos)
			simbolo = trim(array_simbolos(i))
			imagen = imagen_simbolo(simbolo)
			descripcion = describe_simbolo(simbolo)
      if (trim(simbolo) <> "") then
%>
			<img src="imagenes/pictogramas/<%= imagen %>" title="<%= simbolo %>; <%= descripcion %>" width="75px" /><br/>
			<b><%= simbolo %></b>; <%= descripcion %>
			<br/>
<%
      end if
		next
%>
		</p>
<%
	end if
end sub ' ap2_clasificacion_simbolos
' ##################################################################################

sub ap2_clasificacion_simbolos_rd1272()
	if (substance.Item("simbolos_rd1272") <> "") then
%>
		<p id="ap2_clasificacion_simbolos_titulo" class="ficha_titulo_2">Pictogramas y palabras de advertencia</p>
		<p id="ap2_clasificacion_simbolos_cuerpo" class="texto" align="center">
<%
		' Tiene símbolos, muestro cada uno
		substance.Item("simbolos") = replace(substance.Item("simbolos_rd1272"), ",", ";")
		array_simbolos = split(substance.Item("simbolos"), ";")
		for i=0 to ubound(array_simbolos)
			simbolo = trim(array_simbolos(i))
			imagen = ""
			descripcion = ""
			if (left(simbolo,3) = "GHS") then
				imagen = imagen_simbolo(simbolo)
				descripcion = describe_simbolo(simbolo)
			else ' Peligro
				descripcion = "<b style='background-color:red;color:#FFF;'>"+simbolo+"</b>"
			end if
			if (imagen<>"") then
%>
			<img src="imagenes/pictogramas/<%= imagen %>" title="<%= simbolo %>; <%= descripcion %>" width="75px" /><br/>
<%
			end if
%>
			<%= descripcion %>
			<br/><br/>
<%
		next
%>
		</p>
<%
	end if
end sub ' ap2_clasificacion_simbolos_rd1272

' ##################################################################################

sub ap2_clasificacion_frases_r(substance)
	' Muestra las frases R segun clasificacion_1 hasta clasificacion_15
	' No incluye las frases R danesas

	' Montamos frases R

	if (substance.Item("frasesR") <> "") then
%>
		<p id="ap2_clasificacion_frases_r_titulo" class="ficha_titulo_2" style="margin-bottom: -10px;"><a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("Frases R")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> Frases R</p>
<%
		bucle_frases "r", substance.Item("frasesR")
	end if
end sub

' ##################################################################################

sub ap2_clasificacion_frases_h()
	' Muestra las frases H segun clasificacion_rd1272_1 hasta clasificacion_rd1272_15

	' Montamos frases H
	frases_h=monta_frases("H", substance.Item("clasificacion_rd1272_1"), substance.Item("clasificacion_rd1272_2"), substance.Item("clasificacion_rd1272_3"), substance.Item("clasificacion_rd1272_4"), substance.Item("clasificacion_rd1272_5"), substance.Item("clasificacion_rd1272_6"), substance.Item("clasificacion_rd1272_7"), substance.Item("clasificacion_rd1272_8"), substance.Item("clasificacion_rd1272_9"), substance.Item("clasificacion_rd1272_10"), substance.Item("clasificacion_rd1272_11"), substance.Item("clasificacion_rd1272_12"), substance.Item("clasificacion_rd1272_13"), substance.Item("clasificacion_rd1272_14"), substance.Item("clasificacion_rd1272_15"))

	if (frases_h <> "") then
%>
		<p id="ap2_clasificacion_frases_r_titulo" class="ficha_titulo_2" style="margin-bottom: -10px;"><a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("Frases H")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> Frases H</p>
<%
'		bucle_frases "h", frases_h
		muestra_clasificacion 1, substance.Item("clasificacion_rd1272_1")
		muestra_clasificacion 2, substance.Item("clasificacion_rd1272_2")
		muestra_clasificacion 3, substance.Item("clasificacion_rd1272_3")
		muestra_clasificacion 4, substance.Item("clasificacion_rd1272_4")
		muestra_clasificacion 5, substance.Item("clasificacion_rd1272_5")
		muestra_clasificacion 6, substance.Item("clasificacion_rd1272_6")
		muestra_clasificacion 7, substance.Item("clasificacion_rd1272_7")
		muestra_clasificacion 8, substance.Item("clasificacion_rd1272_8")
		muestra_clasificacion 9, substance.Item("clasificacion_rd1272_9")
		muestra_clasificacion 10, substance.Item("clasificacion_rd1272_10")
		muestra_clasificacion 11, substance.Item("clasificacion_rd1272_11")
		muestra_clasificacion 12, substance.Item("clasificacion_rd1272_12")
		muestra_clasificacion 13, substance.Item("clasificacion_rd1272_13")
		muestra_clasificacion 14, substance.Item("clasificacion_rd1272_14")
		muestra_clasificacion 15, substance.Item("clasificacion_rd1272_15")
	end if
	' 23/06/2014 - SPL - Por indicación de Tatiana se pone esta definición.
	if (trim(substance.Item("clasificacion_rd1272_1"))="Expl., ****;") then
		%>
		<p><b>Explosiva</b>: Peligros físicos que deben confirmarse mediante ensayos</p>
		<%
	end if

end sub

' ##################################################################################

sub muestra_clasificacion(numero, clasificacion)
	if (len(trim(clasificacion))>0) then
		' El formato de la clasificacion es Código - Categoria: Frase H
		array_clasificacion = split(clasificacion, ";")
		clas_cat_peligro = trim(array_clasificacion(0))
		if ubound(array_clasificacion)>0 then
			frase = trim(array_clasificacion(1))
		else frase = trim(array_clasificacion(0))
		end if
%>
	    <blockquote style="margin-left: 10px; margin-bottom: -20px;">
<%
			descripcion = describe_frase("h", replace (frase, "*", ""))
			' Para ver definición de los *
 			frase = buscaDefinicionAsteriscos(frase)

 			' Las frases H??? son Gases a presión. Cambio solicitado por Tatiana en abril 2012
 			if (frase="H???") then
%>
	        <b>Gases a presi&oacute;n </b>
<%
 			else
%>
	        <b><%=frase%></b>: <%= descripcion %>
					<% if ubound(array_clasificacion)>0 then %>
						<a href="javascript:toggle('<%= "secc-categpeligro-"+CStr(numero) %>', '<%= "img-fraseh-"+CStr(numero) %>');"><img src="imagenes/desplegar.gif" align="absmiddle" id="<%= "img-fraseh-"+CStr(numero) %>" alt="Pulse para ver el etiquetado" title="Pulse para ver el etiquetado" /></a>
					<% end if %>
	        <br/>
    		<blockquote style="margin-left: 30px; margin-top: 12px; display:none" id="secc-categpeligro-<%=numero%>">
<%
				muestra_frase_clasificacion_rd1272 clas_cat_peligro
%>
		    </blockquote>
<%
			end if
%>
    	</blockquote>

	    <br clear="all" />
<%
	end if
end sub


function buscaDefinicionAsteriscos(cadena)
	' Para ver definición de los *
	if (InStr(cadena,"****")>0) then ' Si hay 4*
		cadena = replace(cadena, "****", "<a onclick=window.open('ver_definicion.asp?id=" + CStr(dame_id_definicion("****")) + "','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'>****</a>")
	else
		if (InStr(cadena,"***")>0) then ' Si hay 3*
			cadena = replace(cadena, "***", "<a onclick=window.open('ver_definicion.asp?id=" + CStr(dame_id_definicion("***")) + "','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'>***</a>")
		else
			if (InStr(cadena,"**")>0) then ' Si hay 2*
				cadena = replace(cadena, "**", "<a onclick=window.open('ver_definicion.asp?id=" + CStr(dame_id_definicion("**")) + "','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'>**</a>")
			else
				if (InStr(cadena,"*")>0) then ' Si hay 1*
					cadena = replace(cadena, "*", "<a onclick=window.open('ver_definicion.asp?id=" + CStr(dame_id_definicion("*")) + "','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'>*</a>")
				end if
			end if
		end if
	end if
	buscaDefinicionAsteriscos = cadena
end function


' ##################################################################################

sub bucle_frases(tipo, byval frases)
		' Pasandole las frases R o H separadas por comas, muestra cada una junto a su descripción
		array_frases = split(frases, ",")
%>
    <blockquote style="margin-left: 10px; margin-bottom: -20px;">
<%
    ' Apuntamos las que hemos mostrado por si hay repetidas
    frases_mostradas = ";"

		for i=0 to ubound(array_frases)
			frase = trim(array_frases(i))
      if(instr(frases_mostradas,frase+";") = 0) then
  			descripcion = describe_frase(tipo, frase)
%>
        <b><%=frase%></b>: <%= descripcion %><br/>
<%
        frases_mostradas = frases_mostradas + frase + ";"
      end if
		next
%>

    </blockquote>

    <br clear="all" />
<%
end sub

' ##################################################################################

sub bucle_frases_s(byval frases_s)
		' Pasandole las frases S separadas por guión, muestra cada una junto a su descripción
		frases_s = replace (frases_s, "S: ", "")
		array_frases_s = split(frases_s, "-")
%>
    <blockquote style="margin-left: 10px; margin-top: -12px; display:none" id="secc-frasess">
<%
		for i=0 to ubound(array_frases_s)
			frase = trim(array_frases_s(i))
			descripcion = describe_frase_s("S"&frase)
%>
			  <b>S<%=frase%></b>:
        <%= descripcion %><br />
<%
		next
%>
    </blockquote>
<%
end sub
' ##################################################################################

sub bucle_categorias_peligro_rd1272(byval frases)
		' Pasandole las frases separadas por guión, muestra cada una junto a su descripción
		array_frases = split(frases, ";")
response.write frases
%>
    <blockquote style="margin-left: 10px; margin-top: -12px; display:none" id="secc-categpeligro">
<%
		for i=0 to ubound(array_frases)
			frase = trim(array_frases(i))
			muestra_frase_clasificacion_rd1272 frase
		next
%>
    </blockquote>
<%
end sub



sub muestra_frase_clasificacion_rd1272(frase)
			if not(trim(frase)="") then
				arrFrase = split(frase, ",")
				descripcion = describe_categoria_peligro(arrFrase(0))
				frase = arrFrase(0)
				if (ubound(arrFrase)>0)then
					categoria = "Cat. " + arrFrase(1)
				else
					categoria = ""
				end if
%>
				  <b><%=frase%> (<%=buscaDefinicionAsteriscos(categoria)%>)</b>:
		        <%= descripcion %><br />
<%
			end if

end sub


' ##################################################################################

sub ap2_clasificacion_frases_r_danesa(substance)
	' Muestra las frases R danesas

	if (substance.Item("frases_r_danesa") <> "") then
%>
	<p id="ap2_clasificacion_frases_r_danesa_titulo" class="ficha_titulo_2"><a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("Frases R según la lista danesa de la EPA")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> Frases R según clasificación de la EPA danesa</p>
<%
		bucle_frases "r", joinFrasesRDanesa(substance.Item("frases_r_danesa"))
	end if
end sub


' ##################################################################################

sub ap2_clasificacion_frases_s
	' Muestra las frases S

	if (substance.Item("frases_s") <> "") then
		' Eliminamos los paréntesis de las frases S
		frases_s = replace (substance.Item("frases_s"), "(", "")
		frases_s = replace (frases_s, ")", "")

%>
	<p id="ap2_clasificacion_frases_s_titulo" class="ficha_titulo_2" style="margin-top: 14px;"><a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("Frases S")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> Frases S <% plegador "secc-frasess", "img-frasess" %></p>

		<% bucle_frases_s frases_s%>

<%
	end if
end sub

' ##################################################################################

sub ap2_clasificacion_notas()
	if (substance.Item("notas_rd_363") <> "") then

		' Dividimos las notas, separadas por puntos, en un array
		array_notas = split(substance.Item("notas_rd_363"), ".")
%>
	<p id="ap2_clasificacion_notas_titulo" class="ficha_titulo_2">Notas <% plegador "secc-notas", "img-notas" %></p>
	<p class="texto" >
		<blockquote id="secc-notas" style="display:none">
<%
		for i=0 to ubound(array_notas)
			nota = trim(array_notas(i))
			id_nota = dame_id_definicion(nota)
			if nota<>"" then
%>

			<b><%=nota%></b> <a onclick=window.open('ver_definicion.asp?id=<%=id_nota%>','def','width=600,height=400,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a><br />
<%
			end if
		next
%>
		</blockquote>
	</p>
<%
	end if
end sub

' ##################################################################################

sub ap2_clasificacion_notas_rd1272()
	dim notas_rd1272LastId : notas_rd1272LastId = Ubound( substance.Item("notas_rd1272") )
	if notas_rd1272LastId < 0 then
		exit sub
	end if
	%>
	<p id="ap2_clasificacion_notas_titulo" class="ficha_titulo_2">Notas <% plegador "secc-notas-rd1272", "img-notas-rd1272" %></p>
	<p class="texto" >
		<blockquote id="secc-notas-rd1272" style="display:none">
		<%
		for i = 0 to notas_rd1272LastId
			set nota = substance.Item("notas_rd1272")(i)
		%>
			<b><%=nota.item("key")%></b>&nbsp;
			<% if nota.item("id")<>""then %>
			<a onclick=window.open('ver_definicion.asp?id=<%=nota.item("id")%>','def','width=600,height=400,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a><br />
			<%end if%>
		<% next %>
		</blockquote>
	</p>
	<%
end sub

' ##################################################################################

sub ap2_clasificacion_etiquetado()
	' Muestra el etiquetado

	if ((substance.Item("conc_1") <> "") or (substance.Item("eti_conc_1") <> "") or (substance.Item("conc_2") <> "") or (substance.Item("eti_conc_2") <> "") or (substance.Item("conc_3") <> "") or (substance.Item("eti_conc_3") <> "") or (substance.Item("conc_4") <> "") or (substance.Item("eti_conc_4") <> "") or (substance.Item("conc_5") <> "") or (substance.Item("eti_conc_5") <> "") or (substance.Item("conc_6") <> "") or (substance.Item("eti_conc_6") <> "") or (substance.Item("conc_7") <> "") or (substance.Item("eti_conc_7") <> "") or (substance.Item("conc_8") <> "") or (substance.Item("eti_conc_8") <> "") or (substance.Item("conc_9") <> "") or (substance.Item("eti_conc_9") <> "") or (substance.Item("conc_10") <> "") or (substance.Item("eti_conc_10") <> "") or (substance.Item("conc_11") <> "") or (substance.Item("eti_conc_11") <> "") or (substance.Item("conc_12") <> "") or (substance.Item("eti_conc_12") <> "") or (substance.Item("conc_13") <> "") or (substance.Item("eti_conc_13") <> "") or (substance.Item("conc_14") <> "") or (substance.Item("eti_conc_14") <> "") or (substance.Item("conc_15") <> "") or (substance.Item("eti_conc_15") <> "")) then

%>
	<span id="ap2_clasificacion_etiquetado_titulo" class="ficha_titulo_2"><a onclick=window.open('ver_definicion.asp?id=88','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> Etiquetado <% plegador "secc-etiquetado", "img-etiquetado" %></span>


  <fieldset id="secc-etiquetado" style="display:none; margin: 15px 45px;">
	<table cellspacing="0" cellpadding="3" width="100%" align="center">
		<tr>
			<th class="subtitulo3 celdaabajo">Concentración</th><th class="subtitulo3 celdaabajo">Etiquetado</th>
		</tr>
<%
	ap2_clasificacion_etiquetado_fila	"r", substance.Item("conc_1"), substance.Item("eti_conc_1")
	ap2_clasificacion_etiquetado_fila	"r", substance.Item("conc_2"), substance.Item("eti_conc_2")
	ap2_clasificacion_etiquetado_fila	"r", substance.Item("conc_3"), substance.Item("eti_conc_3")
	ap2_clasificacion_etiquetado_fila	"r", substance.Item("conc_4"), substance.Item("eti_conc_4")
	ap2_clasificacion_etiquetado_fila	"r", substance.Item("conc_5"), substance.Item("eti_conc_5")
	ap2_clasificacion_etiquetado_fila	"r", substance.Item("conc_6"), substance.Item("eti_conc_6")
	ap2_clasificacion_etiquetado_fila	"r", substance.Item("conc_7"), substance.Item("eti_conc_7")
	ap2_clasificacion_etiquetado_fila	"r", substance.Item("conc_8"), substance.Item("eti_conc_8")
	ap2_clasificacion_etiquetado_fila	"r", substance.Item("conc_9"), substance.Item("eti_conc_9")
	ap2_clasificacion_etiquetado_fila	"r", substance.Item("conc_10"), substance.Item("eti_conc_10")
	ap2_clasificacion_etiquetado_fila	"r", substance.Item("conc_11"), substance.Item("eti_conc_11")
	ap2_clasificacion_etiquetado_fila	"r", substance.Item("conc_12"), substance.Item("eti_conc_12")
	ap2_clasificacion_etiquetado_fila	"r", substance.Item("conc_13"), substance.Item("eti_conc_13")
	ap2_clasificacion_etiquetado_fila	"r", substance.Item("conc_14"), substance.Item("eti_conc_14")
	ap2_clasificacion_etiquetado_fila	"r", substance.Item("conc_15"), substance.Item("eti_conc_15")
%>
	</table>
  </fieldset>

<%
	end if
end sub


' ##################################################################################

sub ap2_clasificacion_etiquetado_rd1272()
	' Muestra el etiquetado

	if ((substance.Item("conc_rd1272_1") <> "") or (substance.Item("eti_conc_rd1272_1") <> "") or (substance.Item("conc_rd1272_2") <> "") or (substance.Item("eti_conc_rd1272_2") <> "") or (substance.Item("conc_rd1272_3") <> "") or (substance.Item("eti_conc_rd1272_3") <> "") or (substance.Item("conc_rd1272_4") <> "") or (substance.Item("eti_conc_rd1272_4") <> "") or (substance.Item("conc_rd1272_5") <> "") or (substance.Item("eti_conc_rd1272_5") <> "") or (substance.Item("conc_rd1272_6") <> "") or (substance.Item("eti_conc_rd1272_6") <> "") or (substance.Item("conc_rd1272_7") <> "") or (substance.Item("eti_conc_rd1272_7") <> "") or (substance.Item("conc_rd1272_8") <> "") or (substance.Item("eti_conc_rd1272_8") <> "") or (substance.Item("conc_rd1272_9") <> "") or (substance.Item("eti_conc_rd1272_9") <> "") or (substance.Item("conc_rd1272_10") <> "") or (substance.Item("eti_conc_rd1272_10") <> "") or (substance.Item("conc_rd1272_11") <> "") or (substance.Item("eti_conc_rd1272_11") <> "") or (substance.Item("conc_rd1272_12") <> "") or (substance.Item("eti_conc_rd1272_12") <> "") or (substance.Item("conc_rd1272_13") <> "") or (substance.Item("eti_conc_rd1272_13") <> "") or (substance.Item("conc_rd1272_14") <> "") or (substance.Item("eti_conc_rd1272_14") <> "") or (substance.Item("conc_rd1272_15") <> "") or (substance.Item("eti_conc_rd1272_15") <> "")) then

%>
	<span id="ap2_clasificacion_etiquetado_titulo" class="ficha_titulo_2"><a onclick=window.open('ver_definicion.asp?id=279','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> Etiquetado <% plegador "secc-etiquetado_rd1272", "img-etiquetado-rd1272" %></span>


  <fieldset id="secc-etiquetado_rd1272" style="display:none; margin: 15px 45px;">
<%
	if (substance.Item("conc_rd1272_1") & substance.Item("conc_rd1272_2"))<>"" then
		if (substance.Item("conc_rd1272_1"))="" then
			if substance.Item("eti_conc_rd1272_1")<>"" then
%>
			Factor <%= substance.Item("eti_conc_rd1272_1") %>
<%
			end if
		end if

%>
	<table cellspacing="0" cellpadding="3" width="100%" align="center">
		<tr>
			<th class="subtitulo3 celdaabajo">Concentración</th><th class="subtitulo3 celdaabajo">Etiquetado</th>
		</tr>
<%
	ap2_clasificacion_etiquetado_fila	"h", substance.Item("conc_rd1272_1"), substance.Item("eti_conc_rd1272_1")
	ap2_clasificacion_etiquetado_fila	"h", substance.Item("conc_rd1272_2"), substance.Item("eti_conc_rd1272_2")
	ap2_clasificacion_etiquetado_fila	"h", substance.Item("conc_rd1272_3"), substance.Item("eti_conc_rd1272_3")
	ap2_clasificacion_etiquetado_fila	"h", substance.Item("conc_rd1272_4"), substance.Item("eti_conc_rd1272_4")
	ap2_clasificacion_etiquetado_fila	"h", substance.Item("conc_rd1272_5"), substance.Item("eti_conc_rd1272_5")
	ap2_clasificacion_etiquetado_fila	"h", substance.Item("conc_rd1272_6"), substance.Item("eti_conc_rd1272_6")
	ap2_clasificacion_etiquetado_fila	"h", substance.Item("conc_rd1272_7"), substance.Item("eti_conc_rd1272_7")
	ap2_clasificacion_etiquetado_fila	"h", substance.Item("conc_rd1272_8"), substance.Item("eti_conc_rd1272_8")
	ap2_clasificacion_etiquetado_fila	"h", substance.Item("conc_rd1272_9"), substance.Item("eti_conc_rd1272_9")
	ap2_clasificacion_etiquetado_fila	"h", substance.Item("conc_rd1272_10"), substance.Item("eti_conc_rd1272_10")
	ap2_clasificacion_etiquetado_fila	"h", substance.Item("conc_rd1272_11"), substance.Item("eti_conc_rd1272_11")
	ap2_clasificacion_etiquetado_fila	"h", substance.Item("conc_rd1272_12"), substance.Item("eti_conc_rd1272_12")
	ap2_clasificacion_etiquetado_fila	"h", substance.Item("conc_rd1272_13"), substance.Item("eti_conc_rd1272_13")
	ap2_clasificacion_etiquetado_fila	"h", substance.Item("conc_rd1272_14"), substance.Item("eti_conc_rd1272_14")
	ap2_clasificacion_etiquetado_fila	"h", substance.Item("conc_rd1272_15"), substance.Item("eti_conc_rd1272_15")
%>
	</table>
<%
	else
		if substance.Item("eti_conc_rd1272_1")<>"" then
%>
			Factor <%= substance.Item("eti_conc_rd1272_1") %>
<%
		end if
	end if
%>
  </fieldset>

<%
	end if
end sub


' ##################################################################################
sub ap2_clasificacion_etiquetado_fila(tipo_frase, byval c, byval e)
  if (not isnull(c) and not isnull(e)) then
	  c = replace (c, ":", "")
	  c = replace (c, "<", "&lt;")
	  c = replace (c, ">", "&gt;")

	  if (c <> "") and (e <> "") then
%>
			<tr>
				<td class="celdaabajo"><%= h(c) %></td><td class="celdaabajo"><%= h(e) %> <a onClick="window.open('busca_frases_<%=tipo_frase%>.asp?id=<%=e%>','fr','width=300,height=200,scrollbars=yes,resizable=yes')" style="text-decoration:none;cursor:pointer"><img src="imagenes/ayuda.gif" border="0" align="absmiddle" alt="busca Frases R"></a></td>
			</tr>
<%
	  end if
  else
  	if (not isnull(e) and e = "*") then
%>
			<tr>
				<td class="celdaabajo" colspan="2">
				Esta entrada tiene límites de concentración específicos para la toxicidad aguda conforme al RD 363/1995 que no pueden «hacerse corresponder» con los límites de concentración con arreglo al Reglamento CLP (como referencia, ver etiquetado del apartado de clasificación (RD 363/1995) de la sustancia).
				</td>
			</tr>
<%
  	end if
  end if
end sub



' ****************
' INICIO DE LISTAS RELACIONADAS
' ****************



' ##################################################################################
' VALORES LÍMITE
sub ap2_clasificacion_vl(id_cajetilla)
	if ((substance.Item("estado_1") <> "") or (substance.Item("vla_ed_ppm_1") <> "") or (substance.Item("vla_ed_mg_m3_1") <> "") or (substance.Item("vla_ec_ppm_1") <> "") or (substance.Item("vla_ec_mg_m3_1") <> "") or (substance.Item("notas_vla_1") <> "") or (substance.Item("estado_2") <> "") or (substance.Item("vla_ed_ppm_2") <> "") or (substance.Item("vla_ed_mg_m3_2") <> "") or (substance.Item("vla_ec_ppm_2") <> "") or (substance.Item("vla_ec_mg_m3_2") <> "") or (substance.Item("notas_vla_2") <> "") or (substance.Item("estado_3") <> "") or (substance.Item("vla_ed_ppm_3") <> "") or (substance.Item("vla_ed_mg_m3_3") <> "") or (substance.Item("vla_ec_ppm_3") <> "") or (substance.Item("vla_ec_mg_m3_3") <> "") or (substance.Item("notas_vla_3") <> "") or (substance.Item("estado_4") <> "") or (substance.Item("vla_ed_ppm_4") <> "") or (substance.Item("vla_ed_mg_m3_4") <> "") or (substance.Item("vla_ec_ppm_4") <> "") or (substance.Item("vla_ec_mg_m3_4") <> "") or (substance.Item("notas_vla_4") <> "") or (substance.Item("estado_5") <> "") or (substance.Item("vla_ed_ppm_5") <> "") or (substance.Item("vla_ed_mg_m3_5") <> "") or (substance.Item("vla_ec_ppm_5") <> "") or (substance.Item("vla_ec_mg_m3_5") <> "") or (substance.Item("notas_vla_5") <> "") or (substance.Item("estado_6") <> "") or (substance.Item("vla_ed_ppm_6") <> "") or (substance.Item("vla_ed_mg_m3_6") <> "") or (substance.Item("vla_ec_ppm_6")  <> "") or (substance.Item("vla_ec_mg_m3_6") <> "") or (substance.Item("notas_vla_6") <> "") or (substance.Item("ib_1") <> "") or  (substance.Item("vlb_1") <> "") or (substance.Item("momento_1") <> "") or (substance.Item("notas_vlb_1") <> "") or (substance.Item("ib_2") <> "") or  (substance.Item("vlb_2") <> "") or (substance.Item("momento_2") <> "") or (substance.Item("notas_vlb_2") <> "") or (substance.Item("ib_3") <> "") or  (substance.Item("vlb_3") <> "") or (substance.Item("momento_3") <> "") or (substance.Item("notas_vlb_3") <> "") or (substance.Item("ib_4") <> "") or  (substance.Item("vlb_4") <> "") or (substance.Item("momento_4") <> "") or (substance.Item("notas_vlb_4") <> "") or (substance.Item("ib_5") <> "") or  (substance.Item("vlb_5") <> "") or (substance.Item("momento_5") <> "") or (substance.Item("notas_vlb_5") <> "") or (substance.Item("ib_6") <> "") or  (substance.Item("vlb_6") <> "") or (substance.Item("momento_6") <> "") or (substance.Item("notas_vlb_6") <> "")) then

%>

	<table id="tabla_valores_limite" class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
  	<tr>
		<td class="celdaabajo" colspan="2" align="center">
			<table cellpadding=0 cellspacing=0 width="100%" border="0"><tr><td width="100%" class="titulo3" align="left"> VALORES L&Iacute;MITE DE EXPOSICI&Oacute;N PROFESIONAL
			</td></tr></table>
		</td>
	</tr>

		<tr>
			<td valign="top" width="50%">
			<!-- VLA -->

<%
		' VLA
		if ((substance.Item("estado_1") <> "") or (substance.Item("vla_ed_ppm_1") <> "") or (substance.Item("vla_ed_mg_m3_1") <> "") or (substance.Item("vla_ec_ppm_1") <> "") or (substance.Item("vla_ec_mg_m3_1") <> "") or (substance.Item("notas_vla_1") <> "") or (substance.Item("estado_2") <> "") or (substance.Item("vla_ed_ppm_2") <> "") or (substance.Item("vla_ed_mg_m3_2") <> "") or (substance.Item("vla_ec_ppm_2") <> "") or (substance.Item("vla_ec_mg_m3_2") <> "") or (substance.Item("notas_vla_2") <> "") or (substance.Item("estado_3") <> "") or (substance.Item("vla_ed_ppm_3") <> "") or (substance.Item("vla_ed_mg_m3_3") <> "") or (substance.Item("vla_ec_ppm_3") <> "") or (substance.Item("vla_ec_mg_m3_3") <> "") or (substance.Item("notas_vla_3") <> "") or (substance.Item("estado_4") <> "") or (substance.Item("vla_ed_ppm_4") <> "") or (substance.Item("vla_ed_mg_m3_4") <> "") or (substance.Item("vla_ec_ppm_4") <> "") or (substance.Item("vla_ec_mg_m3_4") <> "") or (substance.Item("notas_vla_4") <> "") or (substance.Item("estado_5") <> "") or (substance.Item("vla_ed_ppm_5") <> "") or (substance.Item("vla_ed_mg_m3_5") <> "") or (substance.Item("vla_ec_ppm_5") <> "") or (substance.Item("vla_ec_mg_m3_5") <> "") or (substance.Item("notas_vla_5") <> "") or (substance.Item("estado_6") <> "") or (substance.Item("vla_ed_ppm_6") <> "") or (substance.Item("vla_ed_mg_m3_6") <> "") or (substance.Item("vla_ec_ppm_6")  <> "") or (substance.Item("vla_ec_mg_m3_6") <> "") or (substance.Item("notas_vla_6") <> "")) then
%>
	<span id="ap2_clasificacion_vla_titulo" class="ficha_titulo_1"><a href="index.asp?idpagina=616"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>Valores Límite Ambientales <% plegador "secc-vla"+id_cajetilla, "img-vla"+id_cajetilla %></span>
	<fieldset id="secc-vla<%=id_cajetilla%>" style="display:none">
	<table border="0" width="100%" cellspacing="0" cellpadding="3">
		<tr>
			<% if ap2_clasificacion_vl_a_hay_columna_estado then %>
				<td class="subtitulo3 celdaabajo">Estado</td>
			<% end if %>

			<% if (ap2_clasificacion_vl_a_hay_columna_vla_ed_ppm or  ap2_clasificacion_vl_a_hay_columna_vla_ed_mg_m3) then %>
				<td class="subtitulo3 celdaabajo"><a onclick=window.open('ver_definicion.asp?id=230','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> VLA-ED</td>
			<% end if %>

			<% if (ap2_clasificacion_vl_a_hay_columna_vla_ec_ppm or  ap2_clasificacion_vl_a_hay_columna_vla_ec_mg_m3) then %>
				<td class="subtitulo3 celdaabajo"><a onclick=window.open('ver_definicion.asp?id=229','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> VLA-EC</td>
			<% end if %>

			<% if (ap2_clasificacion_vl_a_hay_columna_notas_vla) then %>
				<td class="subtitulo3 celdaabajo" width="25%">Notas</td>
			<% end if %>
		</tr>
<%
		ap2_clasificacion_vl_a substance.Item("estado_1"), substance.Item("vla_ed_ppm_1"), substance.Item("vla_ed_mg_m3_1"), substance.Item("vla_ec_ppm_1"), substance.Item("vla_ec_mg_m3_1"), substance.Item("notas_vla_1")
		ap2_clasificacion_vl_a substance.Item("estado_2"), substance.Item("vla_ed_ppm_2"), substance.Item("vla_ed_mg_m3_2"), substance.Item("vla_ec_ppm_2"), substance.Item("vla_ec_mg_m3_2"), substance.Item("notas_vla_2")
		ap2_clasificacion_vl_a substance.Item("estado_3"), substance.Item("vla_ed_ppm_3"), substance.Item("vla_ed_mg_m3_3"), substance.Item("vla_ec_ppm_3"), substance.Item("vla_ec_mg_m3_3"), substance.Item("notas_vla_3")
		ap2_clasificacion_vl_a substance.Item("estado_4"), substance.Item("vla_ed_ppm_4"), substance.Item("vla_ed_mg_m3_4"), substance.Item("vla_ec_ppm_4"), substance.Item("vla_ec_mg_m3_4"), substance.Item("notas_vla_4")
		ap2_clasificacion_vl_a substance.Item("estado_5"), substance.Item("vla_ed_ppm_5"), substance.Item("vla_ed_mg_m3_5"), substance.Item("vla_ec_ppm_5"), substance.Item("vla_ec_mg_m3_5"), substance.Item("notas_vla_5")
		ap2_clasificacion_vl_a substance.Item("estado_6"), substance.Item("vla_ed_ppm_6"), substance.Item("vla_ed_mg_m3_6"), substance.Item("vla_ec_ppm_6"), substance.Item("vla_ec_mg_m3_6"), substance.Item("notas_vla_6")
%>
	</table>
	</fieldset>
<%
		end if

%>
		<!-- Fin VLA -->
		</td>
		<td valign="top">
		<!-- VLB -->
<%

		' VLB
		if ((substance.Item("ib_1") <> "") or (substance.Item("vlb_1") <> "") or (substance.Item("momento_1") <> "") or (substance.Item("notas_vlb_1") <> "") or (substance.Item("ib_2") <> "") or (substance.Item("vlb_2") <> "") or (substance.Item("momento_2") <> "") or (substance.Item("notas_vlb_2") <> "") or (substance.Item("ib_3") <> "") or (substance.Item("vlb_3") <> "") or (substance.Item("momento_3") <> "") or (substance.Item("notas_vlb_3") <> "") or (substance.Item("ib_4") <> "") or (substance.Item("vlb_4") <> "") or (substance.Item("momento_4") <> "") or (substance.Item("notas_vlb_4") <> "") or (substance.Item("ib_5") <> "") or (substance.Item("vlb_5") <> "") or (substance.Item("momento_5") <> "") or (notas_vlb_51 <> "") or (substance.Item("ib_6") <> "") or (substance.Item("vlb_6") <> "") or (substance.Item("momento_6") <> "") or (substance.Item("notas_vlb_6") <> "")) then
%>

		<p id="ap2_clasificacion_vlb_titulo" class="ficha_titulo_1"><a href="index.asp?idpagina=616"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>Valores Límite Biológicos <% plegador "secc-vlb"+id_cajetilla, "img-vlb"+id_cajetilla %></p>
		<fieldset id="secc-vlb<%=id_cajetilla%>" style="display:none">
		<table width="100%" cellspacing="0" cellpadding="3">
			<tr>
			<% if ap2_clasificacion_vl_b_hay_columna_ib then %>
				<td class="subtitulo3 celdaabajo">Indicador</th>
			<% end if %>

			<% if ap2_clasificacion_vl_b_hay_columna_vlb then %>
				<td class="subtitulo3 celdaabajo">Valor límite</th>
			<% end if %>

			<% if ap2_clasificacion_vl_b_hay_columna_momento then %>
				<td class="subtitulo3 celdaabajo">Momento de muestreo</th>
			<% end if %>

			<% if ap2_clasificacion_vl_b_hay_columna_notas_vlb then %>
				<td class="subtitulo3 celdaabajo">Notas</th>
			<% end if %>
			</tr>
<%
			ap2_clasificacion_vl_b substance.Item("ib_1"), substance.Item("vlb_1"), substance.Item("momento_1"), substance.Item("notas_vlb_1")
			ap2_clasificacion_vl_b substance.Item("ib_2"), substance.Item("vlb_2"), substance.Item("momento_2"), substance.Item("notas_vlb_2")
			ap2_clasificacion_vl_b substance.Item("ib_3"), substance.Item("vlb_3"), substance.Item("momento_3"), substance.Item("notas_vlb_3")
			ap2_clasificacion_vl_b substance.Item("ib_4"), substance.Item("vlb_4"), substance.Item("momento_4"), substance.Item("notas_vlb_4")
			ap2_clasificacion_vl_b substance.Item("ib_5"), substance.Item("vlb_5"), substance.Item("momento_5"), substance.Item("notas_vlb_5")
			ap2_clasificacion_vl_b substance.Item("ib_6"), substance.Item("vlb_6"), substance.Item("momento_6"), substance.Item("notas_vlb_6")
%>
		</table>
		</fieldset>
<%
		end if
%>
		<!-- Fin VLB -->
		</td>
	</tr>
	</table>

<%
	end if
end sub

' ##################################################################################

sub ap2_clasificacion_vl_a(estado, vla_ed_ppm, vla_ed_mg_m3, vla_ec_ppm, vla_ec_mg_m3, notas_vla)
	' Mostramos una fila si hay algún dato
	if (trim(estado&vla_ed_ppm&vla_ed_mg_m3&vla_ec_ppm&vla_ec_mg_m3&notas_vla) <> "") then
%>
		<tr>
			<% if ap2_clasificacion_vl_a_hay_columna_estado then %>
				<td style="	border-bottom-color: #DDDDDD;	border-bottom-style: solid;	border-bottom-width: 1px;" valign="middle"><%= estado %></td>
			<% end if %>

			<% if (ap2_clasificacion_vl_a_hay_columna_vla_ed_ppm or  ap2_clasificacion_vl_a_hay_columna_vla_ed_mg_m3) then %>
				<td style="	border-bottom-color: #DDDDDD;	border-bottom-style: solid;	border-bottom-width: 1px;" valign="middle">
					<% if (vla_ed_ppm <> "") then response.write vla_ed_ppm & " ppm<br />" end if %>
					<% if (vla_ed_mg_m3 <> "") then response.write vla_ed_mg_m3 & " mg/m3" end if %>
				</td>
			<% end if %>

			<% if (ap2_clasificacion_vl_a_hay_columna_vla_ec_ppm or  ap2_clasificacion_vl_a_hay_columna_vla_ec_mg_m3) then %>
			<td style="	border-bottom-color: #DDDDDD;	border-bottom-style: solid;	border-bottom-width: 1px;" valign="middle">
				<% if (vla_ec_ppm <> "") then response.write vla_ec_ppm & " ppm<br />" end if %>
				<% if (vla_ec_mg_m3 <> "") then response.write vla_ec_mg_m3 & " mg/m3" end if %>
			</td>
			<% end if %>

			<% if (ap2_clasificacion_vl_a_hay_columna_notas_vla) then %>
				<td style="	border-bottom-color: #DDDDDD;	border-bottom-style: solid;	border-bottom-width: 1px;" valign="middle"><% notas_con_ayuda notas_vla, "VLA" %></td>
			<% end if %>
		</tr>
<%
	end if
end sub

' ##################################################################################

sub ap2_clasificacion_vl_b(ib, vlb, momento, notas_vlb)
	' Pinta una fila si hay algún dato
		if (trim(replace( ib&vlb&momento&notas_vlb, ",", "") ) <> "") then

%>
			<tr>
			<% if ap2_clasificacion_vl_b_hay_columna_ib then %>
				<td style="	border-bottom-color: #DDDDDD;	border-bottom-style: solid;	border-bottom-width: 1px;" valign="middle"><%=ib%></td>
			<% end if %>

			<% if ap2_clasificacion_vl_b_hay_columna_vlb then %>
				<td style="	border-bottom-color: #DDDDDD;	border-bottom-style: solid;	border-bottom-width: 1px;" valign="middle"><%=vlb%></td>
			<% end if %>

			<% if ap2_clasificacion_vl_b_hay_columna_momento then %>
				<td style="	border-bottom-color: #DDDDDD;	border-bottom-style: solid;	border-bottom-width: 1px;" valign="middle"><%=parche_definicion(momento, "MomentoVLBInicio")%><a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion(parche_definicion(momento, "MomentoVLB"))%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><%= parche_definicion(momento, "MomentoVLB") %></a>
				</td>
			<% end if %>

			<% if ap2_clasificacion_vl_b_hay_columna_notas_vlb then %>
				<td style="	border-bottom-color: #DDDDDD;	border-bottom-style: solid;	border-bottom-width: 1px;" valign="middle"><% notas_con_ayuda notas_vlb, "VLB" %></td>
			<% end if %>
			</tr>
<%
	end if
end sub

' ##################################################################################

function ap2_clasificacion_vl_a_hay_columna_estado()
	valores = substance.Item("estado_1") & substance.Item("estado_2") & substance.Item("estado_3") & substance.Item("estado_4") & substance.Item("estado_5") & substance.Item("estado_6")
	if (valores <> "") then
		ap2_clasificacion_vl_a_hay_columna_estado = true
	else
		ap2_clasificacion_vl_a_hay_columna_estado = false
	end if
end function

' ##################################################################################

function ap2_clasificacion_vl_a_hay_columna_vla_ed_ppm()
	valores = substance.Item("vla_ed_ppm_1") & substance.Item("vla_ed_ppm_2") & substance.Item("vla_ed_ppm_3") & substance.Item("vla_ed_ppm_4") & substance.Item("vla_ed_ppm_5") & substance.Item("vla_ed_ppm_6")
	if (valores <> "") then
		ap2_clasificacion_vl_a_hay_columna_vla_ed_ppm = true
	else
		ap2_clasificacion_vl_a_hay_columna_vla_ed_ppm = false
	end if
end function

' ##################################################################################

function ap2_clasificacion_vl_a_hay_columna_vla_ed_mg_m3()
	valores = substance.Item("vla_ed_mg_m3_1") & substance.Item("vla_ed_mg_m3_2") & substance.Item("vla_ed_mg_m3_3") & substance.Item("vla_ed_mg_m3_4") & substance.Item("vla_ed_mg_m3_5") & substance.Item("vla_ed_mg_m3_6")
	if (valores <> "") then
		ap2_clasificacion_vl_a_hay_columna_vla_ed_mg_m3 = true
	else
		ap2_clasificacion_vl_a_hay_columna_vla_ed_mg_m3 = false
	end if
end function

' ##################################################################################

function ap2_clasificacion_vl_a_hay_columna_vla_ec_ppm()
	valores = substance.Item("vla_ec_ppm_1") & substance.Item("vla_ec_ppm_2") & substance.Item("vla_ec_ppm_3") & substance.Item("vla_ec_ppm_4") & substance.Item("vla_ec_ppm_5") & substance.Item("vla_ec_ppm_6")
	if (valores <> "") then
		ap2_clasificacion_vl_a_hay_columna_vla_ec_ppm = true
	else
		ap2_clasificacion_vl_a_hay_columna_vla_ec_ppm = false
	end if
end function

' ##################################################################################

function ap2_clasificacion_vl_a_hay_columna_vla_ec_mg_m3()
	valores = substance.Item("vla_ec_mg_m3_1") & substance.Item("vla_ec_mg_m3_2") & substance.Item("vla_ec_mg_m3_3") & substance.Item("vla_ec_mg_m3_4") & substance.Item("vla_ec_mg_m3_5") & substance.Item("vla_ec_mg_m3_6")
	if (valores <> "") then
		ap2_clasificacion_vl_a_hay_columna_vla_ec_mg_m3 = true
	else
		ap2_clasificacion_vl_a_hay_columna_vla_ec_mg_m3 = false
	end if
end function

' ##################################################################################

function ap2_clasificacion_vl_a_hay_columna_notas_vla()
	valores = substance.Item("notas_vla_1") & substance.Item("notas_vla_2") & substance.Item("notas_vla_3") & substance.Item("notas_vla_4") & substance.Item("notas_vla_5") & substance.Item("notas_vla_6")
	if (valores <> "") then
		ap2_clasificacion_vl_a_hay_columna_notas_vla = true
	else
		ap2_clasificacion_vl_a_hay_columna_notas_vla = false
	end if
end function

' ##################################################################################

function ap2_clasificacion_vl_b_hay_columna_ib()
	valores = substance.Item("ib_1") & substance.Item("ib_2") & substance.Item("ib_3") & substance.Item("ib_4") & substance.Item("ib_5") & substance.Item("ib_6")
	if (valores <> "") then
		ap2_clasificacion_vl_b_hay_columna_ib = true
	else
		ap2_clasificacion_vl_b_hay_columna_ib = false
	end if
end function

' ##################################################################################

function ap2_clasificacion_vl_b_hay_columna_vlb()
	valores = substance.Item("vlb_1") & substance.Item("vlb_2") & substance.Item("vlb_3") & substance.Item("vlb_4") & substance.Item("vlb_5") & substance.Item("vlb_6")
	if (valores <> "") then
		ap2_clasificacion_vl_b_hay_columna_vlb = true
	else
		ap2_clasificacion_vl_b_hay_columna_vlb = false
	end if
end function

' ##################################################################################

function ap2_clasificacion_vl_b_hay_columna_momento()
	valores = substance.Item("momento_1") & substance.Item("momento_2") & substance.Item("momento_3") & substance.Item("momento_4") & substance.Item("momento_5") & substance.Item("momento_6")
	if (valores <> "") then
		ap2_clasificacion_vl_b_hay_columna_momento = true
	else
		ap2_clasificacion_vl_b_hay_columna_momento = false
	end if
end function

' ##################################################################################

function ap2_clasificacion_vl_b_hay_columna_notas_vlb()
	valores = substance.Item("notas_vlb_1") & substance.Item("notas_vlb_2") & substance.Item("notas_vlb_3") & substance.Item("notas_vlb_4") & substance.Item("notas_vlb_5") & substance.Item("notas_vlb_6")
	if (valores <> "") then
		ap2_clasificacion_vl_b_hay_columna_notas_vlb = true
	else
		ap2_clasificacion_vl_b_hay_columna_notas_vlb = false
	end if
end function

' ##################################################################################

sub notas_con_ayuda(byval notas, byval tipo)

	' Para buscar la definición hay ocasiones en las que hay que aplicar un parche.

	array_notas = split(notas, ",")
	cadena_notas = ""
	for i=0 to ubound(array_notas)
		nota = trim(array_notas(i))
		id_nota = dame_id_definicion(parche_definicion(nota, tipo))
		if (nota <> "") then
			if (cadena_notas = "") then
				cadena_notas = "<a onclick=window.open('ver_definicion.asp?id="&id_nota&"','def','width=600,height=400,scrollbars=yes,resizable=yes') style='cursor:pointer'>"&nota&"</a>"
			else
				cadena_notas = cadena_notas & ", <a onclick=window.open('ver_definicion.asp?id="&id_nota&"','def','width=600,height=400,scrollbars=yes,resizable=yes') style='cursor:pointer'>"&nota&"</a>"
			end if
		end if
	next
	response.write cadena_notas
end sub

' ##################################################################################

sub ap2_clasificacion_lista_negra(mySubstance)
	' Muestra el etiquetado

	if mySubstance.hasListaNegraClassifications() then

    ' Esta en lista negra. Aprovechamos para marcarle el bit correspondiente para que aparezca en el listado de lista negra
    sqlListaNegra="UPDATE dn_risc_sustancias SET negra=1 WHERE id="&id_sustancia
    objConnection2.execute(sqlListaNegra),,adexecutenorecords

%>
		<p id="ap2_clasificacion_lista_negra_titulo" class="subtitulo3">&nbsp;<img src="imagenes/icono_atencion_20.png" align="absmiddle" /> <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("Lista negra")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> Sustancia incluida en la Lista negra de ISTAS <% plegador "secc-listanegra", "img-listanegra" %></p>
		<p id="secc-listanegra" class="texto" style="display:none">Esta sustancia está incluida en la Lista negra de ISTAS por los siguientes motivos: <%=arraySerialize(mySubstance.fields.item("listaNegraClassifications"))%></p>

<%
	end if
end sub

' ###################################################################################

sub ap3_riesgos()
	' SALUD

	'Sergio
	sql = "select comentarios from dn_risc_sustancias_salud where id_sustancia="&id_sustancia
	set objRstq=objConnection2.execute(sql)
	if(not objRstq.eof) then
		comentarios_sl = objrstq("comentarios")
	end if


	if (mySubstance.inList("cancer_rd") or mySubstance.inList("cancer_danesa") or mySubstance.inList("cancer_iarc") or mySubstance.inList("cancer_otras") or mySubstance.inList("cancer_mama") or mySubstance.inList("de") or mySubstance.inNeurotoxicosLists() or substance.Item("efecto_neurotoxico")="OTOTÓXICO" or mySubstance.inList("sensibilizante") or mySubstance.inList("sensibilizante_reach") or mySubstance.inList("sensibilizante_danesa") or mySubstance.inList("tpr") or mySubstance.inList("tpr_danesa") or mySubstance.inList("eepp") or mySubstance.inList("mutageno_rd") or mySubstance.inList("mutageno_danesa") or mySubstance.inList("salud") or mySubstance.inList("prohibidas_embarazadas") or mySubstance.inList("prohibidas_lactantes") or comentarios_sl <> "") then
%>

		<!-- ################ Riesgos para la salud ###################### -->
		<br />
		<div id="ficha">
		<table width="100%" cellpadding=5>
			<tr>
				<td>
					<a name="identificacion"></a><img src="imagenes/risctox02.gif" alt="Riesgos específicos para la salud" />
				</td>
				<td align="right">
					<a href="#"><img src="imagenes/subir.gif" border=0 alt=subir></a>
				</td>
			</tr>
		</table>

<%
		if (mySubstance.inList("cancer_rd") or mySubstance.inList("cancer_danesa") or mySubstance.inList("cancer_iarc") or mySubstance.inList("cancer_otras") or mySubstance.inList("cancer_mama")) then ap3_riesgos_tabla("Cancerígeno") end if
		if (mySubstance.inList("mutageno_rd") or mySubstance.inList("mutageno_danesa") ) then ap3_riesgos_tabla("Mutágeno") end if

		if mySubstance.inList("de") then ap3_riesgos_tabla("Disruptor endocrino") end if
		if mySubstance.inNeurotoxicosLists() or substance.Item("efecto_neurotoxico")="OTOTÓXICO" then ap3_riesgos_tabla("Neurotóxico") end if
		if mySubstance.inList("sensibilizante") or mySubstance.inList("sensibilizante_danesa") or mySubstance.inList("sensibilizante_reach") then ap3_riesgos_tabla("Sensibilizante") end if
		if mySubstance.inList("tpr") or mySubstance.inList("tpr_danesa") then ap3_riesgos_tabla("Tóxico para la reproducción") end if
		if mySubstance.inList("eepp") then ap3_riesgos_enfermedades() end if
    	if mySubstance.inList("salud") then ap7_salud() end if
%>

		<%

			if comentarios_sl <> "" then
		%>
			<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr>
					<td class="celdaabajo" colspan="2" align="center">
						<table cellpadding=0 cellspacing=0 width="100%" border="0">
							<tr>
								<td width="100%" class="titulo3" align="left">

							Más información en salud laboral
							<a href="javascript:toggle('secc-mas_informacion_salud_laboral', 'img-mas_informacion_salud_laboral');"><img src="imagenes/desplegar.gif" align="absmiddle" id="img-mas_informacion_salud_laboral" alt="Pulse para desplegar la información" title="Pulse para desplegar la información" /></a>
			        			</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>

					<td id="secc-mas_informacion_salud_laboral" style="display:none">


						<ul>
							<li>
							<%= comentarios_sl %>
							</li>
						</ul>

					</td>
				</tr>
			</table>
			<br />
		<%
			end if
		%>

		</div>
		<!-- ################ Fin Riesgos para la salud ################## -->
<%
	end if ' salud
%>

<% ' MEDIO AMBIENTE %>
<%
if (mySubstance.inList("tpb") or mySubstance.inList("directiva_aguas") or mySubstance.inList("alemana") or mySubstance.inList("sustancias_prioritarias")  or mySubstance.inList("ozono") or mySubstance.inList("clima") or mySubstance.inList("aire") or mySubstance.inList("cop") or substance.Item("comentarios_medio_ambiente") <>"" or mySubstance.inList("suelos")) then %>

		<!-- ################ Riesgos para el medio ambiente ###################### -->
		<br />
		<div id="ficha">
		<table width="100%" cellpadding=5>
			<tr>
				<td>
                        <a name="identificacion"></a><img src="imagenes/risctox03.gif" alt="Riesgos específicos para el medio ambiente" />

				</td>
				<td align="right">
					<a href="#"><img src="imagenes/subir.gif" border=0 alt=subir></a>
				</td>
			</tr>
		</table>
<%
		if mySubstance.inList("tpb") then
			ap3_riesgos_tabla("Tóxica, Persistente y Bioacumulativa")
		end if
		' SPL (16/06/20014)
'		if num_cas="87-68-3" or num_cas="133-49-3" or num_cas="75-74-1" then
		if mySubstance.inMpmbList() then
			ap3_riesgos_tabla("mPmB")
		end if
		if (mySubstance.inList("directiva_aguas") or mySubstance.inList("alemana")) then ap3_riesgos_tabla("Tóxica para el agua") end if
		if (mySubstance.inList("suelos")) then ap3_riesgos_tabla("Contaminante de suelos") end if
		if (mySubstance.inList("ozono") or mySubstance.inList("clima") or mySubstance.inList("aire")) then ap3_riesgos_tabla("Contaminante del aire") end if

		if (mySubstance.inList("cop")) then ap3_riesgos_tabla("Contaminante Orgánico Persistente (COP)") end if
%>

		<%
		if (substance.Item("comentarios_medio_ambiente") <>"") then
		%>
			<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr>
					<td class="celdaabajo" colspan="2" align="center">
						<table cellpadding=0 cellspacing=0 width="100%" border="0">
							<tr>
								<td width="100%" class="titulo3" align="left">

							Más información en medio ambiente
							<a href="javascript:toggle('secc-mas_informacion_medio_ambiente', 'img-mas_informacion_medio_ambiente');"><img src="imagenes/desplegar.gif" align="absmiddle" id="img-mas_informacion_medio_ambiente" alt="Pulse para desplegar la información" title="Pulse para desplegar la información" /></a>
			        			</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>

					<td id="secc-mas_informacion_medio_ambiente" style="display:none">


						<ul>
							<li>
							<%=substance.Item("comentarios_medio_ambiente") %>
							</li>
						</ul>

					</td>
				</tr>
			</table>
			<br />

		<%
		end if
		%>
		</div>
		<!-- ################ Fin Riesgos para el medio ambiente ################## -->
<%
	end if ' medio ambiente
end sub ' ap3_riesgos


' ###################################################################################

sub ap3_riesgos_tabla(byval tipo)

	' Muestra la tabla de riesgos con sus datos, dependiendo del tipo

%>
	<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
   	<tr>
			<td class="celdaabajo" colspan="2" align="center">
				<table cellpadding=0 cellspacing=0 width="100%" border="0"><tr><td width="100%" class="titulo3" align="left"><% ap3_riesgos_tabla_ayuda(tipo) %><%=tipo%>

        <% if ((tipo <> "COV") and (tipo <> "Vertidos") and (tipo <> "IPPC (PRTR Agua)") and (tipo <> "IPPC (PRTR Aire)") and (tipo <> "IPPC (PRTR Suelo)") and (tipo <> "Residuos Peligrosos") and (tipo <> "Accidentes Graves") and (tipo <> "Emisiones Atmosféricas") ) then %>

        <% plegador "secc-"&tipo, "img-"&tipo %>

        <% end if %>

        </td></tr></table>
			</td>
		</tr>
		<tr>
			<td id="secc-<%= aplana(tipo) %>" style="display:none">
			<% ap3_riesgos_tabla_contenidos(tipo) %>
			</td>
		</tr>
	</table>
	<br />
<%
end sub

' ###################################################################################

sub ap3_riesgos_tabla_ayuda(tipo)

	select case tipo
		case "Cancerígeno":
%>
			<a href="index.asp?idpagina=607"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>
<%
		case "Mutágeno":
%>
			<a href="index.asp?idpagina=607"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>
<%
		case "Disruptor endocrino":
%>
			<a href="index.asp?idpagina=610"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>
<%
		case "Neurotóxico":
%>
			<a href="index.asp?idpagina=611"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>
<%
		case "Sensibilizante":
%>
			<a href="index.asp?idpagina=612"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>
<%
		case "Tóxico para la reproducción":
%>
			<a href="index.asp?idpagina=609"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>
<%
		case "Tóxica, Persistente y Bioacumulativa":
%>
			<a href="index.asp?idpagina=613"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>
        <%
		case "mPmB":
%>
			<a href="index.asp?idpagina=613"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>
<%
		case "Tóxica para el agua":
%>
			<a href="index.asp?idpagina=614"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>

       <%
		case "Contaminante de suelos":
%>
			<a href="index.asp?idpagina=622"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>

<%
		case "Contaminante Orgánico Persistente (COP)":
%>
			<a href="index.asp?idpagina=1185"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>
<%
		case "Contaminante del aire":
%>
			<a href="index.asp?idpagina=615"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>
<%
		case "Residuos Peligrosos":
%>
			<a href="index.asp?idpagina=618"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>
<%
		case "Vertidos":
%>
			<a href="index.asp?idpagina=619"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>
<%
		case "Accidentes Graves":
%>
			<a href="index.asp?idpagina=623"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>
<%
		case "COV":
%>
			<a href="index.asp?idpagina=621"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>
<%
		case "IPPC (PRTR Agua)":
%>
			<a href="index.asp?idpagina=622"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>
<%
		case "IPPC (PRTR Aire)":
%>
			<a href="index.asp?idpagina=622"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>
<%
		case "IPPC (PRTR Suelo)":
%>
			<a href="index.asp?idpagina=622"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>
<%
		case "Emisiones Atmosféricas":
%>
			<a href="index.asp?idpagina=620"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>
<%
		case "Prohibida para trabajadoras embarazadas":
%>
			<a href="index.asp?idpagina=1188"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>
<%
		case "Prohibida para trabajadoras lactantes":
%>
			<a href="index.asp?idpagina=1188"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>
<%
		case "Sustancia candidata REACH":
%>
			<a href="index.asp?idpagina=1194"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>
<%
		case "Sustancia REACH sujeta a autorización":
%>
			<a href="index.asp?idpagina=1194"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>
<%
		case "Sustancia biocida autorizada":
%>
			<a href="index.asp?idpagina=1192"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>
<%
		case "Sustancia biocida prohibida":
%>
			<a href="index.asp?idpagina=1192"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>
<%
		case "Sustancia pesticida autorizada":
%>
			<a href="index.asp?idpagina=1191"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>
<%
		case "Sustancia pesticida prohibida":
%>
			<a href="index.asp?idpagina=1191"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>
<%
		case "Sustancia bajo evaluación. CoRAP":
%>
			<a href="index.asp?idpagina=1194"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a>
<%
	end select

end sub

' ###################################################################################

sub ap3_riesgos_tabla_contenidos(tipo)

	select case tipo

	case "Accidente Grave"
	%>

    Accidente Grave


<%
	case "Contaminante de suelos"
	%>

    Según <a href="http://www.istas.net/web/abreenlace.asp?idenlace=2940" target="_blank">Real Decreto 9/2005</a>


<%


    case "Contaminante Orgánico Persistente (COP)":

%>

    <fieldset>

      <legend class="subtitulo3"><strong>Según Convenio de Estocolmo</strong></legend>

      <ul>

<%

      if isNull(substance.Item("cop")) then

        substance.Item("cop") = ""

      end if



      array_anexos = split(substance.Item("cop"), ";")

      for i=0 to ubound(array_anexos)

%>

        <li><%=get_definition("COP Anexo "&trim(array_anexos(i)), "es")%></li>

<%

      next

%>
		<%
	  	if (trim(substance.Item("enlace_cop")) <> "") then
			response.write "<li><a href='"&substance.Item("enlace_cop")&"' target='_blank'>Más información</a></li>"
		end if
	  %>

      </ul>


    </fieldset>

<%
		case "Cancerígeno":

				' Real Decreto ---------------------------------------------------------------
				if (mySubstance.inList("cancer_rd")) then
%>
					<fieldset>
						<legend class="subtitulo3"><strong>Según R. 1272/2008</strong></legend>
						<blockquote>
<%
				nivel_cancerigeno_rd = dame_nivel_cancerigeno_rd()
				' Tatiana - 01/8/2012 - Las categorías sustituir 1 por 1A, 2 por 1B y 3 por 2.
				nivel_cancerigeno_rd_txt = replace(nivel_cancerigeno_rd, "1", "1A")
				nivel_cancerigeno_rd_txt = replace(nivel_cancerigeno_rd_txt, "2", "1B")
				nivel_cancerigeno_rd_txt = replace(nivel_cancerigeno_rd_txt, "3", "2")

				if (nivel_cancerigeno_rd <> "") then
							response.write "<strong>Nivel cancerígeno:</strong> "&nivel_cancerigeno_rd_txt
%>
					 		<a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("C"&nivel_cancerigeno_rd_txt)%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>
<%
				end if
%>

<%
				if (trim(substance.Item("notas_cancer_rd")) <> "") then
%>
							<br/><strong>Notas:</strong> <%=substance.Item("notas_cancer_rd")%>
<%
				end if
%>
						</blockquote>
					</fieldset>
<%
				end if



				' Lista danesa ---------------------------------------------------------------
				if (mySubstance.inList("cancer_danesa")) then
		%>
					<fieldset>
						<legend class="subtitulo3"><strong>Según <% plegador_texto "frases_r_danesa_cancer", "frases R", "subtitulo3" %> en la clasificación de la EPA danesa <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("EPA Danesa")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>.</strong></legend>
						<blockquote>
		<%
				nivel_cancerigeno_danesa = dame_nivel_cancerigeno_danesa()
				if (nivel_cancerigeno_danesa <> "") then
					response.write "<strong>Nivel cancerígeno:</strong> "&nivel_cancerigeno_danesa
		%>

					 <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("RDC"&nivel_cancerigeno_danesa)%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>
		<%
				end if
		%>

		<%
					if (trim(substance.Item("notas_cancer_rd")) <> "") then
		%>
						<br/><strong>Notas:</strong> <%=substance.Item("notas_cancer_rd")%>
		<%
					end if
		%>
		        <div id="frases_r_danesa_cancer" style="display:none"><br />
		        <% ap2_clasificacion_frases_r_danesa(substance) %>
		        </div>

						</blockquote>
					</fieldset>
		<%
				end if



				' IARC -----------------------------------------------------------------------
				if (mySubstance.inList("cancer_iarc")) then
		%>
					<fieldset>
						<legend class="subtitulo3"><strong>Según IARC <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("IARC")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a></strong></legend>
		<%
						if (trim(substance.Item("grupo_iarc")) <> "") or (trim(substance.Item("volumen_iarc")) <> "") or (trim(substance.Item("notas_iarc")) <> "") then
		%>
							<blockquote>
							<table>
		<%
							if (trim(substance.Item("grupo_iarc")) <> "") then
		%>
								<tr><td class="subtitulo3">Grupo:</td><td><%=trim(replace(ucase(substance.Item("grupo_iarc")), "GRUPO", ""))%> <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion(trim(substance.Item("grupo_iarc")))%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a></td></tr>
		<%
							end if

							if (trim(substance.Item("volumen_iarc")) <> "") then
		%>
								<tr><td class="subtitulo3">Volumen:</td><td><%=substance.Item("volumen_iarc")%></td></tr>
		<%
							end if
							if (trim(substance.Item("notas_iarc")) <> "") then
		%>
								<tr><td class="subtitulo3">Notas:</td><td><%=substance.Item("notas_iarc")%></td></tr>
		<%
							end if
		%>
							</table>
							</blockquote>
		<%
						end if
		%>
					</fieldset>
		<%
				end if

				' Otras fuentes
				if (mySubstance.inList("cancer_otras")) then

		%>

		    <fieldset>
				  <legend class="subtitulo3"><strong>Según otras fuentes</strong></legend>

		<%

		      if (isNull(substance.Item("categoria_cancer_otras"))) then

		        substance.Item("categoria_cancer_otras") = ""

		      end if



		      if (isNull(substance.Item("fuente"))) then

		        substance.Item("fuente") = ""

		      end if


					array_categorias=split(substance.Item("categoria_cancer_otras"), ",")
					array_fuentes=split(substance.Item("fuente"), ",")

					' Damos por hecho que hay el mismo numero de categorias y fuentes y que coinciden en orden
					for i=0 to ubound(array_fuentes)
		%>
					<fieldset>
						<legend class="subtitulo3"><strong>Según <%=trim(array_fuentes(i))%> <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion(trim(array_fuentes(i)))%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a></strong></legend>
						<blockquote>
						<table>
							<tr><td class="subtitulo3"><%=trim(array_categorias(i))%>:</td><td><%= get_definition(trim(array_categorias(i)), "es") %></td></tr>
						</table>
						</blockquote>
					</fieldset>
		<%
					next

		%>

		    </fieldset>

		<%
				end if



		    ' Cancer mama

		    if (mySubstance.inList("cancer_mama")) then

		      if (isNull(substance.Item("cancer_mama_fuente"))) then

		        substance.Item("cancer_mama_fuente") = ""

		      end if

		%>

					<fieldset>
						<legend class="subtitulo3"><strong>Según SSI (cáncer de mama) <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("SSI")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a></strong></legend>
						<blockquote>
						<table>
							<tr><td class="subtitulo3"><strong>Fuente:</strong><br /><a href="<%= substance.Item("cancer_mama_fuente") %>" target="_blank"><%= replace(substance.Item("cancer_mama_fuente"), "http://", "") %></a></td></tr>
						</table>
						</blockquote>
					</fieldset>

		<%

		    end if

		case "Mutágeno":
      ' MUTAGENO RD -------------------------------------------------------------
      if (mySubstance.inList("mutageno_rd")) then
%>
			<fieldset>
				<legend class="subtitulo3"><strong>Según R. 1272/2008</strong></legend>
				<blockquote>
				<%
					nivel_mutageno_rd = dame_nivel_mutageno_rd()
					' Tatiana - 01/8/2012 - Las categorías sustituir 1 por 1A, 2 por 1B y 3 por 2.
					nivel_mutageno_rd_txt = replace(nivel_mutageno_rd, "1", "1A")
					nivel_mutageno_rd_txt = replace(nivel_mutageno_rd_txt, "2", "1B")
					nivel_mutageno_rd_txt = replace(nivel_mutageno_rd_txt, "3", "2")

					if (nivel_mutageno_rd <> "") then
					response.write "<br /><strong>Nivel mutágeno:</strong> "&nivel_mutageno_rd_txt
				%>
					 <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("M"&nivel_mutageno_rd_txt)%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>
				<%
					end if
				%>
				</blockquote>
			</fieldset>
<%
      end if


      ' MUTAGENO DANESA -------------------------------------------------------------
      if (mySubstance.inList("mutageno_danesa")) then
%>
			<fieldset>
				<legend class="subtitulo3"><strong>Según <% plegador_texto "frases_r_danesa_mutageno", "frases R", "subtitulo3" %> en la clasificación de la EPA danesa <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("EPA Danesa")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>.</strong></legend>
				<blockquote>
				<%
					nivel_mutageno_danesa = dame_nivel_mutageno_danesa()
					if (nivel_mutageno_danesa <> "") then
					response.write "<br /><strong>Nivel mutágeno:</strong> "&nivel_mutageno_danesa
				%>
					 <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("RDM"&nivel_mutageno_danesa)%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>
				<%
					end if
				%>

        <div id="frases_r_danesa_mutageno" style="display:none"><br />
        <% ap2_clasificacion_frases_r_danesa(substance) %>
        </div>

				</blockquote>
			</fieldset>
<%
      end if




		case "Disruptor endocrino":
%>
			<blockquote>
			<table>
			<% if (substance.Item("nivel_disruptor") <> "") then %>
				<tr>
					<td class="subtitulo3" valign="top">Fuente:</td>
					<td>
					<%
					array_niveles=split(substance.Item("nivel_disruptor"), ",")
					for i=0 to ubound(array_niveles)
						nivel=get_definition(trim(array_niveles(i)), "es")
						response.write nivel&"<br /><br />"
					next
					%>
					</td>
				</tr>
			<% end if %>
			</table>
			</blockquote>
<%
		case "Neurotóxico":


        if mySubstance.inList("neurotoxico_rd") or mySubstance.inList("neurotoxico_danesa") then
          ' Añadimos SNC a efecto neurotoxico si no existía ya
          if (substance.Item("efecto_neurotoxico") = "") or (IsNull(substance.Item("efecto_neurotoxico"))) then
            substance.Item("efecto_neurotoxico")="SNC"
          else
            if (not (inStr(substance.Item("efecto_neurotoxico"), "SNC") > 0)) then
              substance.Item("efecto_neurotoxico") = substance.Item("efecto_neurotoxico") & "/SNC"
            end if
          end if
        end if

        if mySubstance.inList("neurotoxico_rd") then
          if (substance.Item("fuente_neurotoxico") = "") or (IsNull(substance.Item("fuente_neurotoxico"))) then
            substance.Item("fuente_neurotoxico") = "363"
          else
            substance.Item("fuente_neurotoxico") = substance.Item("fuente_neurotoxico") & ",363"
          end if
        end if

        if mySubstance.inList("neurotoxico_danesa") then
          if (substance.Item("fuente_neurotoxico") = "") or (IsNull(substance.Item("fuente_neurotoxico"))) then
            substance.Item("fuente_neurotoxico") = "EPA-R67"
          else
            substance.Item("fuente_neurotoxico") = substance.Item("fuente_neurotoxico") & ",EPA-R67"
          end if
        end if
      %>


      <% if ((substance.Item("efecto_neurotoxico") <> "") or (substance.Item("nivel_neurotoxico") <> "") or (substance.Item("fuente_neurotoxico") <> "")) then %>
			<blockquote>
			<table>
			<%	if (substance.Item("efecto_neurotoxico") <> "") then %>
				<tr>
					<td class="subtitulo3" valign="top">Efecto:</td>
					<td>
						<%
							' Separamos el efecto neurotoxico por "/". Ejemplo: "SNC/NEUROTOXICO/OTOTOXICO" se convierte en 3 definiciones, cada una con su ayuda.
							array_neurotoxico = split(substance.Item("efecto_neurotoxico"), "/")
							for i=0 to ubound(array_neurotoxico)
								efecto = trim(array_neurotoxico(i))
                efecto = ucase(efecto)
						%>

						<%= efecto %> <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion(efecto)%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>

						<%
							next
						%>
					</td>
				</tr>
			<% end if %>
			<% if (substance.Item("nivel_neurotoxico") <> "") then %>
				<tr>
					<td class="subtitulo3" valign="top">Nivel:</td><td><%=substance.Item("nivel_neurotoxico")%>

					 <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("Nivel "&substance.Item("nivel_neurotoxico"))%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>

					</td>
			</tr>
			<% end if %>
			<% if (substance.Item("fuente_neurotoxico") <> "") then %>
				<tr>
					<td class="subtitulo3" valign="top">Fuente:</td>
					<td>
					<%
					array_fuentes=split(substance.Item("fuente_neurotoxico"), ",")
					for i=0 to ubound(array_fuentes)
          				response.write get_definition(trim(array_fuentes(i)), "es")


			%>
          <%
            		if (i < ubound(array_fuentes)) then
              			response.write "<br><br> "
           			 end if
					next
					%>
					</td></tr>
			<% end if %>
			</table>
			</blockquote>
      <% end if %>
<%

		case "Sensibilizante":
		      response.write "<ul>"
					' Indicamos si es por lista RD o por lista danesa
		      if mySubstance.inList("sensibilizante") then
		        response.write "<li class='subtitulo3'>Sensibilizante según R. 1272/2008</li>"
		      end if

			  if mySubstance.inList("sensibilizante_reach") then
		        response.write "<li class='subtitulo3'>Alérgeno REACH &nbsp;<a href='http://www.istas.net/web/abreenlace.asp?idenlace=6340' target='_blank'>Ver documento</a></li>"
		      end if

		      if mySubstance.inList("sensibilizante_danesa") then
		      %>
		        <li class='subtitulo3'>Sensibilizante según <% plegador_texto "frases_r_danesa_sensibilizante", "frases R", "subtitulo3" %> en la clasificación de la EPA danesa <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("EPA Danesa")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a></li>
		      <%


		      %>
		        <div id="frases_r_danesa_sensibilizante" style="display:none"><br />
		        <blockquote>
		        <% ap2_clasificacion_frases_r_danesa(substance) %>
		        </blockquote>
		        </div>
		      <%
			  end if
			  response.write "</ul>"


		case "Tóxico para la reproducción":
	      ' TPR SEGUN RD -------------------------------------------------------------
	      if (mySubstance.inList("tpr")) then
	%>
	    			<fieldset>
	  				<legend class="subtitulo3"><strong>Según R. 1272/2008</strong></legend>
	<%
	  			nivel_reproduccion_rd = dame_nivel_reproduccion_rd()
				' Tatiana - 01/8/2012 - Las categorías sustituir 1 por 1A, 2 por 1B y 3 por 2.
				nivel_reproduccion_rd_txt = replace(nivel_reproduccion_rd, "1", "1A")
				nivel_reproduccion_rd_txt = replace(nivel_reproduccion_rd_txt, "2", "1B")
				nivel_reproduccion_rd_txt = replace(nivel_reproduccion_rd_txt, "3", "2")
	  			if (nivel_reproduccion_rd <> "") then
				  %>
	  				<blockquote>
	  					<strong>Categoría:</strong> <%=nivel_reproduccion_rd_txt%>
	  				  <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("TR"&nivel_reproduccion_rd_txt)%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>
	  					</blockquote>
	  			<%
	  			end if
	%>
	          </fieldset>
	<%
	      end if


	      ' TPR SEGUN LISTA DANESA ---------------------------------------------------
	      if (mySubstance.inList("tpr_danesa")) then
	%>
	    			<fieldset>
	  				<legend class="subtitulo3"><strong>Según <% plegador_texto "frases_r_danesa_tpr", "frases R", "subtitulo3" %> en la clasificación de la EPA danesa <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("EPA Danesa")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a></strong></legend>
	<%
	  			nivel_reproduccion_danesa = dame_nivel_reproduccion_danesa()
	  			if (nivel_reproduccion_danesa <> "") then
				  %>
	  				<blockquote>
	  					<strong>Categoría:</strong> <%=nivel_reproduccion_danesa%>
	  				  <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("RDR"&nivel_reproduccion_danesa)%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>
	  					</blockquote>
	  			<%
	  			end if
	%>
	        <div id="frases_r_danesa_tpr" style="display:none"><br />
	        <blockquote>
	        <% ap2_clasificacion_frases_r_danesa(substance) %>
	        </blockquote>
	        </div>
	          </fieldset>
	<%
	      end if

	case "Prohibida para trabajadoras embarazadas":

      if (mySubstance.inList("prohibidas_embarazadas")) then
%>
  				<blockquote>
  					<strong>Fuente:</strong> Real Decreto 298/2009
				</blockquote>
<%
      end if

	case "Prohibida para trabajadoras lactantes":

      if (mySubstance.inList("prohibidas_lactantes")) then
%>
  				<blockquote>
  					<strong>Fuente:</strong> Real Decreto 298/2009
				</blockquote>
<%
      end if


		case "Tóxica, Persistente y Bioacumulativa":
%>
			<blockquote>
			<table>
				<tr>
					<td class="subtitulo3">Más información (en inglés):</td>
					<td><a href="<%= substance.Item("enlace_tpb") %>"><%= corta(substance.Item("anchor_tpb"), 70, "puntossuspensivos") %></a></td>
				</tr>
				<tr>
					<td class="subtitulo3" valign="top">Fuente/s:</td>
					<td class="subtitulo3"><%
						if trim(substance.Item("fuentes_tpb")) <> "" then
							array_tpb = split(substance.Item("fuentes_tpb"),",")
							for i=0 to ubound(array_tpb)
								response.write "<li>" & get_definition(trim(array_tpb(i)), "es")&"</li>"
							next
						end if
						if trim(substance.Item("fuente_tpb")) <> "" then
							array_tpb = split(substance.Item("fuente_tpb"),",")
							for i=0 to ubound(array_tpb)
								response.write "<li>" & get_definition(trim(array_tpb(i)), "es")&"</li>"
							next
						end if

					%>
					 </td>
				</tr>
			</table>
			</blockquote>
<%
		case "mPmB":
%>
			<blockquote>
			<table>
				<tr>
					<td class="subtitulo3"><%= get_definition("REACH") %></td>

				</tr>

			</table>
			</blockquote>
        			</blockquote>
<%
		case "Sustancia restringida":
%>
			<blockquote>
			<table>
				<tr>
					<td class="subtitulo3">
	                    <a href="#" onClick="window.open('dn_mas_informacion.asp?listado=restringidas&id=<%=id_sustancia%>','Informacion','width=500,height=230,scrollbars=yes,resizable=yes')">Más información</a>
                    </td>

				</tr>

			</table>
			</blockquote>


       <%
		case "Sustancia prohibida":
%>
			<blockquote>
            			<table>
				<tr>
                	<td class="subtitulo3">
                    	<a href="#" onClick="window.open('dn_mas_informacion.asp?listado=prohibidas&id=<%=id_sustancia%>','Informacion','width=500,height=230,scrollbars=yes,resizable=yes')">Más información</a>
                    </td>
				</tr>
			</table>
			</blockquote>


<%
		case "Tóxica para el agua":
			response.write "<table>"
			if (substance.Item("directiva_aguas") or mySubstance.inList("directiva_aguas")) then
%>
				<tr>
					<td class="subtitulo3" colspan="2">· Según <a href="http://www.istas.net/web/abreenlace.asp?idenlace=2227" target="_blank">Directiva de aguas</a>, y sus posteriores <a href="http://www.istas.net/web/abreenlace.asp?idenlace=6323">modificaciones</a></td>
				</tr>
<%
			end if

			if (mySubstance.inList("sustancias_prioritarias")) then
%>
				<tr>
					<td class="subtitulo3" colspan="2">· Posible sustancia prioritaria según la <a href="http://www.istas.net/web/abreenlace.asp?idenlace=2227" target="_blank">Directiva de aguas</a>, y sus posteriores <a href="http://www.istas.net/web/abreenlace.asp?idenlace=6323" target="_blank">modificaciones</a></td>
				</tr>
<%
			end if

			if (trim(substance.Item("clasif_mma")) <> "") and (trim(substance.Item("clasif_mma"))<>"nwg")then
%>
				<tr>
					<td class="subtitulo3" colspan="2">
						· Según <a href="http://www.istas.net/risctox/abreenlace.asp?idenlace=2226" target="_blank">Ministerio de Medio Ambiente de Alemania</a>
					</td>
				</tr>
				<tr>
					<td>&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<td><strong>Clasificación</strong>: <%=substance.Item("clasif_mma")%>
					 	<a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion(parche_definicion(substance.Item("clasif_mma"), "MMA"))%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>
					</td>
				</tr>
<%
			end if
			if (substance.Item("sustancia_prioritaria")=1)then
%>
				<tr>
					<td class="subtitulo3">Posible sustancia prioritaria </td><td></td>
				</tr>
<%
			end if
			response.write "</table>"


			case "Contaminante del aire":
%>
				<table>
<%
				if (substance.Item("dano_calidad_aire") or mySubstance.inList("aire")) then
%>
					<tr>
						<td class="subtitulo3">Calidad del aire:</td>
						<td>Sustancia incluida en la <a href="abreenlace.asp?idenlace=2234" target="_blank">Directiva 96/62/CE</a> de 27 de septiembre sobre evaluación y gestión de la calidad del aire ambiente</td>
					</tr>
<%
				end if
%>
<%
				if (substance.Item("dano_ozono")) then
%>
					<tr>
						<td class="subtitulo3">Capa de ozono:</td>
						<td>Sustancia que agota la capa de ozono, según <a href="abreenlace.asp?idenlace=2229" target="_blank">Reglamento (CE) 2037/2000</a> del Parlamento Europeo y del Consejo, de 29 de junio de 2000</td>
					</tr>
<%
				end if
%>
<%
				if (substance.Item("dano_cambio_clima")) then
%>
					<tr>
						<td class="subtitulo3">Cambio climático:</td>
						<td>Sustancia incluida en el listado del <a href="abreenlace.asp?idenlace=2230" target="_blank">Protocolo de Kyoto</a></td>
					</tr>
<%
				end if
%>



				</table>
<%
		case "Sustancia candidata REACH":
%>
			<blockquote>
            			<table>
				<tr>
                	<td class="subtitulo3">
                    	Fuente: <a href="https://echa.europa.eu/es/candidate-list-table" target="_blank">Agencia Europea de sustancias y mezclas químicas (ECHA)</a>
                    </td>
				</tr>
			</table>
			</blockquote>


<%
		case "Sustancia REACH sujeta a autorización":
%>
			<blockquote>
            			<table>
				<tr>
                	<td class="subtitulo3">
                    	Fuente: <a href="https://echa.europa.eu/es/authorisation-list" target="_blank">Agencia Europea de sustancias y mezclas químicas (ECHA)</a>
                    </td>
				</tr>
			</table>
			</blockquote>


<%
		case "Sustancia biocida prohibida":
%>
			<blockquote>
            			<table>
				<tr>
                	<td class="subtitulo3">
                    	<a href="#" onClick="window.open('dn_mas_informacion.asp?listado=biocidas_prohibidas&id=<%=id_sustancia%>','Informacion','width=500,height=230,scrollbars=yes,resizable=yes')">Más información</a>
                    </td>
				</tr>
			</table>
			</blockquote>


<%
		case "Sustancia biocida autorizada":
%>
			<blockquote>
            			<table>
				<tr>
                	<td class="subtitulo3">
                    	<a href="#" onClick="window.open('dn_mas_informacion.asp?listado=biocidas_autorizadas&id=<%=id_sustancia%>','Informacion','width=500,height=230,scrollbars=yes,resizable=yes')">Más información</a>
                    </td>
				</tr>
			</table>
			</blockquote>


<%
		case "Sustancia pesticida prohibida":
%>
			<blockquote>
            			<table>
				<tr>
                	<td class="subtitulo3">
                    	<a href="#" onClick="window.open('dn_mas_informacion.asp?listado=pesticidas_prohibidas&id=<%=id_sustancia%>','Informacion','width=500,height=230,scrollbars=yes,resizable=yes')">Más información</a>
                    </td>
				</tr>
			</table>
			</blockquote>


<%
		case "Sustancia pesticida autorizada":
%>
			<blockquote>
            			<table>
				<tr>
                	<td class="subtitulo3">
                    	<a href="#" onClick="window.open('dn_mas_informacion.asp?listado=pesticidas_autorizadas&id=<%=id_sustancia%>','Informacion','width=500,height=230,scrollbars=yes,resizable=yes')">Más información</a>
                    </td>
				</tr>
			</table>
			</blockquote>
<%
		case "Sustancia bajo evaluación. CoRAP":
%>
			<blockquote>
				<table>
				<tr>
					<td class="subtitulo3">
							<a href="#" onClick="window.open('dn_mas_informacion.asp?listado=corap&id=<%=id_sustancia%>','Informacion','width=500,height=230,scrollbars=yes,resizable=yes')">Mas información</a>
						</td>
				</tr>
				<tr>
					<td class="subtitulo3">
						Fuente: <a href="http://echa.europa.eu/es/information-on-chemicals/evaluation/community-rolling-action-plan/corap-table" target="_blank">European Chemicals Agency (ECHA)</a>
					</td>
				</tr>
			</table>
			</blockquote>
<%
	end select
end sub

' ###################################################################################

sub ap3_riesgos_enfermedades()

	' Se agrupan por listado, cada listado en una ficha blanca y dentro cada enfermedad
	sql_enf = "select distinct enf.id, enf.listado, enf.nombre, enf.sintomas, enf.actividades FROM dn_risc_enfermedades AS enf LEFT OUTER JOIN dn_risc_grupos_por_enfermedades AS gpe ON enf.id = gpe.id_enfermedad LEFT OUTER JOIN dn_risc_sustancias_por_grupos AS spg ON gpe.id_grupo = spg.id_grupo LEFT OUTER JOIN dn_risc_sustancias_por_enfermedades AS spe ON spe.id_enfermedad = enf.id WHERE spg.id_sustancia="&id_sustancia&" OR spe.id_sustancia="&id_sustancia&" ORDER BY enf.listado, enf.nombre"
	set objRstEnf=objConnection2.execute(sql_enf)
	if (not objRstEnf.eof) then
		listado_antiguo = ""
		do while (not objRstEnf.eof)
			' Para mostrar agrupados por listado, solo escribimos la cabecera si el listado es nuevo
			if (listado_antiguo <> objRstEnf("listado")) then

				' Si el listado antiguo no es vacío, es que ya habiamos abierto antes uno así que primero cerramos el anterior
				if (listado_antiguo <> "") then
%>
			</td>
		</tr>
	</table>
	<br />
<%
				end if
%>
	<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
   	<tr>
			<td class="celdaabajo" colspan="2" align="center">
				<table cellpadding=0 cellspacing=0 width="100%" border="0"><tr><td width="100%" class="titulo3" align="left"><a href="index.asp?idpagina=617"><img src="imagenes/ayuda.gif" align="absmiddle" border="0" /></a> <%=objRstEnf("listado")%>  <% plegador "secc-enf"&objRstEnf("listado"), "img-enf"&objRstEnf("listado") %></td></tr></table>
			</td>
		</tr>
		<tr id="secc-enf<%= aplana(objRstEnf("listado")) %>" style="display:none">
			<td>
<%
				listado_antiguo = objRstEnf("listado")
			end if
				if objRstEnf("nombre")<>"" then
%>
				<fieldset style="padding:10px;">
				<!-- Tabla enfermedad -->
				<table cellspacing=1 cellpadding=1 border=0>
					<tr>
						<td class="subtitulo3" colspan=2><%=objRstEnf("nombre")%></td>
					</tr>
				<%
					if (objRstEnf("sintomas") <> "") then
				%>
					<tr>
						<td class="subtitulo3" align="right" valign="top" width='10%' nowrap style='padding-top:10px'>Síntomas:</td><td align="left" style'padding-top:10px'><%=replace(objRstEnf("sintomas"), vbcrlf, "<br>")%></td>
					</tr>
				<%
					end if
				%>
				<%
					if (objRstEnf("actividades") <> "") then
				%>
					<tr>
						<td class="subtitulo3" align="right" valign="top" width="10%" nowrap style='padding-top:10px'>Actividades:</td><td align="left"  style='padding-top:10px'><%=replace(objRstEnf("actividades"), vbcrlf, "<br>")%></td>
					</tr>
				<%
					end if
				%>
				</table>
				<!-- Fin tabla enfermedad -->
                </fieldset>
                <br />

<%
			end if
			objRstEnf.movenext
		loop
		' Tras el bucle siempre cerramos la tabla
%>
			</td>
		</tr>
	</table>
	<br />
<%
	end if
	objRstEnf.close()
	set objRstEnf=nothing
end sub

' ###################################################################################

sub ap4_normativa_ambiental()
	if mySubstance.inList("cov") or mySubstance.inList("residuos") or mySubstance.inList("vertidos") or mySubstance.inList("lpcic")  or mySubstance.inList("accidentes") or mySubstance.inList("emisiones") then
%>

		<!-- ################ Normativa ambiental ###################### -->
		<br />
		<div id="ficha">
		<table width="100%" cellpadding=5>
			<tr>
				<td>
					<a name="identificacion"></a><img src="imagenes/risctox05.gif" alt="Normativa ambiental que le afecta" />
				</td>
				<td align="right">
					<a href="#"><img src="imagenes/subir.gif" border=0 alt=subir></a>
				</td>
			</tr>
		</table>

<%
' Para dividir los 7 posibles apartados en dos columnas, primero calculamos cuántos hay en total.
total = 0

if mySubstance.inList("cov") then total = total +1 end if
if mySubstance.inList("vertidos") then total = total +1 end if
if mySubstance.inList("lpcic-agua") then total = total +1 end if
if mySubstance.inList("lpcic-aire") then total = total +1 end if
if mySubstance.inList("lpcic-suelo") then total = total +1 end if
if mySubstance.inList("residuos") then total = total +1 end if
if mySubstance.inList("accidentes") then total = total +1 end if
if mySubstance.inList("emisiones") then total = total +1 end if

mitad = round(total / 2)
' Ajustamos la mitad para arriba si es impar
if ((mitad * 2) < total) then
	mitad = mitad + 1
end if
%>

		<table border="0" width="100%">
			<tr>
				<td valign="top" width="50%">
<%
' Contaremos cuantos llevamos para ver en qué momento hay que poner la división de columnas
llevo = 0
%>

<%
		if mySubstance.inList("cov") then
			ap3_riesgos_tabla("COV")
			llevo = llevo +1
			if llevo >= mitad then
				response.write "</td><td valign='top' width='50%'>"
				llevo = 0 ' Lo reseteo para que no vuelva a dividir
			end if
		end if

		if mySubstance.inList("vertidos") then
			ap3_riesgos_tabla("Vertidos")
			llevo = llevo +1
			if llevo >= mitad then
				response.write "</td><td valign='top' width='50%'>"
				llevo = 0 ' Lo reseteo para que no vuelva a dividir
			end if
		end if

		if mySubstance.inList("lpcic-agua") then
			ap3_riesgos_tabla("IPPC (PRTR Agua)")
			llevo = llevo +1
			if llevo >= mitad then
				response.write "</td><td valign='top' width='50%'>"
				llevo = 0 ' Lo reseteo para que no vuelva a dividir
			end if
		end if

		if mySubstance.inList("lpcic-aire") then
			ap3_riesgos_tabla("IPPC (PRTR Aire)")
			llevo = llevo +1
			if llevo >= mitad then
				response.write "</td><td valign='top' width='50%'>"
				llevo = 0 ' Lo reseteo para que no vuelva a dividir
			end if
		end if

		if mySubstance.inList("lpcic-suelo") then
			ap3_riesgos_tabla("IPPC (PRTR Suelo)")
			llevo = llevo +1
			if llevo >= mitad then
				response.write "</td><td valign='top' width='50%'>"
				llevo = 0 ' Lo reseteo para que no vuelva a dividir
			end if
		end if

		if mySubstance.inList("residuos") then
			ap3_riesgos_tabla("Residuos Peligrosos")
			llevo = llevo +1
			if llevo >= mitad then
				response.write "</td><td valign='top' width='50%'>"
				llevo = 0 ' Lo reseteo para que no vuelva a dividir
			end if
		end if

		if mySubstance.inList("accidentes") then
			ap3_riesgos_tabla("Accidentes Graves")
			llevo = llevo +1
			if llevo >= mitad then
				response.write "</td><td valign='top' width='50%'>"
				llevo = 0 ' Lo reseteo para que no vuelva a dividir
			end if
		end if

		if mySubstance.inList("emisiones") then
			ap3_riesgos_tabla("Emisiones Atmosféricas")
			llevo = llevo +1
			if llevo >= mitad then
				response.write "</td><td valign='top' width='50%'>"
				llevo = 0 ' Lo reseteo para que no vuelva a dividir
			end if
		end if
	%>
				</td>
			</tr>
		</table>
		</div>
		<!-- ################ Fin Normativa ambiental ################## -->
<%
	end if
end sub ' ap4_normativa_ambiental

sub ap4_normativa_salud_laboral()

end sub ' ap4_normativa_salud_laboral



sub ap4_normativa_restriccion_prohibicion()
	if mySubstance.inList("prohibidas") or mySubstance.inList("restringidas") or mySubstance.inList("candidatas_reach") or mySubstance.inList("autorizacion_reach") or mySubstance.inList("biocidas_autorizadas") or mySubstance.inList("biocidas_prohibidas") or mySubstance.inList("pesticidas_autorizadas") or mySubstance.inList("pesticidas_prohibidas") or mySubstance.inList("prohibidas_embarazadas") or mySubstance.inList("prohibidas_lactantes") or mySubstance.inList("corap") then
%>

		<!-- ################ Normativa salud laboral ###################### -->
		<br />
		<div id="ficha">
		<table width="100%" cellpadding=5>
			<tr>
				<td>
					<a name="identificacion"></a><img src="imagenes/risctox04-restricciones.gif" alt="Normativa sobre restricción/prohibición de sustancias" />
				</td>
				<td align="right">
					<a href="#"><img src="imagenes/subir.gif" border=0 alt=subir></a>
				</td>
			</tr>
		</table>

		<table border="0" width="100%">
			<tr>
				<td valign="top">
<%
		if mySubstance.inList("prohibidas") then
			ap3_riesgos_tabla("Sustancia prohibida")
		end if

		if mySubstance.inList("restringidas") then
			ap3_riesgos_tabla("Sustancia restringida")
		end if

		if mySubstance.inList("prohibidas_embarazadas") then ap3_riesgos_tabla("Prohibida para trabajadoras embarazadas") end if

		if mySubstance.inList("prohibidas_lactantes") then ap3_riesgos_tabla("Prohibida para trabajadoras lactantes") end if

		if mySubstance.inList("candidatas_reach") then
			ap3_riesgos_tabla("Sustancia candidata REACH")
		end if
		if mySubstance.inList("autorizacion_reach") then
			ap3_riesgos_tabla("Sustancia REACH sujeta a autorización")
		end if
		if mySubstance.inList("biocidas_autorizadas") then
			ap3_riesgos_tabla("Sustancia biocida autorizada")
		end if
		if mySubstance.inList("biocidas_prohibidas") then
			ap3_riesgos_tabla("Sustancia biocida prohibida")
		end if
		if mySubstance.inList("pesticidas_autorizadas") then
			ap3_riesgos_tabla("Sustancia pesticida autorizada")
		end if
		if mySubstance.inList("pesticidas_prohibidas") then
			ap3_riesgos_tabla("Sustancia pesticida prohibida")
		end if
		if mySubstance.inList("corap") then
			ap3_riesgos_tabla("Sustancia bajo evaluación. CoRAP")
		end if

%>
				</td>
			</tr>
		</table>
		</div>
		<!-- ################ Fin Normativa salud laboral ################## -->
<%
	end if
end sub ' ap4_normativa_restriccion_prohibicion



' ##################################################################################
sub ap5_alternativas()

	sql="SELECT DISTINCT f.id AS id_fichero, f.titulo FROM dn_alter_ficheros AS f LEFT OUTER JOIN dn_alter_ficheros_por_sustancias AS fps ON f.id = fps.id_fichero LEFT OUTER JOIN dn_alter_ficheros_por_grupos AS fpg ON f.id = fpg.id_fichero LEFT OUTER JOIN dn_risc_grupos AS g ON fpg.id_grupo = g.id LEFT OUTER JOIN dn_risc_sustancias_por_grupos AS spg ON g.id = spg.id_grupo WHERE fps.id_sustancia="&id_sustancia&" OR spg.id_sustancia = "& id_sustancia&" ORDER BY titulo"

	set objRst=objConnection2.execute(sql)
	if (not objRst.eof) then
%>
	<!-- Alternativas -->
	<br />
	<div id="ficha">
	<table width="100%" cellpadding=5>
		<tr>
			<td>
				<a name="identificacion"></a><img src="imagenes/risctox08.gif" alt="Alternativas" />
			</td>
			<td align="right">
				<a href="#"><img src="imagenes/subir.gif" border=0 alt=subir></a>
			</td>
		</tr>
	</table>
	<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
   	<tr>
			<td class="celdaabajo" colspan="2" align="center">
				<table cellpadding=0 cellspacing=0 width="100%" border="0"><tr><td width="100%" class="titulo3" align="left">Alternativas <% plegador "secc-alternativas", "img-alternativas" %></td></tr></table>
			</td>
		</tr>
		<tr id="secc-alternativas" style="display:none">
			<td>
				<ul>
<%
	' Mostramos los ficheros, comprobando que no haya titulos repetidos. Como vienen ordenados por título, basta comparar con el título anterior
	titulo_antiguo = ""
	do while (not objRst.eof)
		id_fichero=objRst("id_fichero")
		titulo=objRst("titulo")
		if (titulo <> titulo_antiguo) then
%>
					<li><a href="dn_alternativas_ficha_fichero.asp?id_fichero=<%=id_fichero%>"><%=titulo%></a></li>
<%
			titulo_antiguo = titulo
		end if
		objRst.movenext
	loop
%>
				</ul>
			</td>
		</tr>
	</table>
	<br />
	</div>
	<!-- Fin alternativas -->
<%
	end if
	objRst.close()
	set objRst = nothing
end sub

' ##################################################################################
sub ap6_sectores()

	sql="SELECT DISTINCT s.numero_cnae AS codigo, s.nombre AS nombre, s.id AS id_sector FROM dn_alter_sectores AS s LEFT OUTER JOIN dn_risc_sustancias_por_sectores AS sps ON s.id = sps.id_sector WHERE sps.id_sustancia="&id_sustancia&" ORDER BY numero_cnae"

	set objRst=objConnection2.execute(sql)
	if (not objRst.eof) then
%>
	<!-- Sectores -->
	<br />
	<div id="ficha">
	<table width="100%" cellpadding=5>
		<tr>
			<td>
				<a name="identificacion"></a><img src="imagenes/risctox07.gif" alt="Sectores" />
			</td>
			<td align="right">
				<a href="#"><img src="imagenes/subir.gif" border=0 alt=subir></a>
			</td>
		</tr>
	</table>
	<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
   	<tr>
			<td class="celdaabajo" colspan="2" align="center">
				<table cellpadding=0 cellspacing=0 width="100%" border="0"><tr><td width="100%" class="titulo3" align="left">Sectores donde se encuentra esta sustancia <% plegador "secc-sectores", "img-sectores" %></td></tr></table>
			</td>
		</tr>
		<tr id="secc-sectores" style="display:none">
			<td>
				<ul>
<%
	' Mostramos los sectores, comprobando que no haya codigos repetidos. Como vienen ordenados por código, basta comparar con el código anterior
	codigo_antiguo = ""
	do while (not objRst.eof)
		id_sector=objRst("id_sector")
		codigo=objRst("codigo")
		nombre=objRst("nombre")
		if (codigo <> codigo_antiguo) then
      ' Si no tiene documentos asociados, mostraremos solo el texto sin enlace.
      sqlDocs="SELECT COUNT(*) AS num FROM dn_alter_ficheros_por_sectores WHERE id_sector="&id_sector
      set objRstDocs = objConnection2.execute(sqlDocs)
      if objRstDocs("num") > 0 then
%>
					<li><a href="dn_alternativas_ficha_sector.asp?id=<%=id_sector%>"><%=codigo%> - <%=nombre%></a></li>
<%
      else
%>
					<li><%=codigo%> - <%=nombre%></li>
<%
      end if

			codigo_antiguo = codigo
		end if
		objRst.movenext
	loop
%>
				</ul>
			</td>
		</tr>
	</table>
	<br />
	</div>
	<!-- Fin sectores -->
<%
	end if
	objRst.close()
	set objRst = nothing
end sub

' #############################################################################################

sub ap7_salud()

	sql="SELECT cardiocirculatorio, rinyon, respiratorio, reproductivo, piel_sentidos, neuro_toxicos, musculo_esqueletico, sistema_inmunitario, higado_gastrointestinal, sistema_endocrino, embrion, cancer, comentarios FROM dn_risc_sustancias_salud WHERE id_sustancia="&id_sustancia&" AND (cardiocirculatorio=1 OR rinyon=1 OR respiratorio=1 OR reproductivo=1 OR piel_sentidos=1 OR neuro_toxicos=1 OR musculo_esqueletico=1 OR sistema_inmunitario=1 OR higado_gastrointestinal=1 OR sistema_endocrino=1 OR embrion=1 OR cancer=1)"

	set objRst=objConnection2.execute(sql)
	if (not objRst.eof) then
%>
	<!-- Efectos para la salud -->
	<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
   	<tr>
			<td class="celdaabajo" colspan="2" align="center">
				<table cellpadding=0 cellspacing=0 width="100%" border="0"><tr><td width="100%" class="titulo3" align="left">Otras alteraciones para la salud y sistemas y órganos afectados <% plegador "secc-salud", "img-salud" %></td></tr></table>
			</td>
		</tr>
		<tr id="secc-salud" style="display:none">
			<td>
      <table border="0" width="100%">
        <tr>
<%
	' Mostramos los efectos
	if (not objRst.eof) then
    cardiocirculatorio = objRst("cardiocirculatorio")
    respiratorio = objRst("respiratorio")
    reproductivo = objRst("reproductivo")
    musculo_esqueletico = objRst("musculo_esqueletico")
    sistema_inmunitario = objRst("sistema_inmunitario")
    higado_gastrointestinal = objRst("higado_gastrointestinal")
    sistema_endocrino = objRst("sistema_endocrino")

    embrion = objRst("embrion")
    cancer = objRst("cancer")
    rinyon = objRst("rinyon")
    piel_sentidos = objRst("piel_sentidos")
    neuro_toxicos = objRst("neuro_toxicos")

	comentarios_sl = objrst("comentarios")

    if (cardiocirculatorio OR respiratorio OR reproductivo OR musculo_esqueletico OR sistema_inmunitario OR higado_gastrointestinal OR sistema_endocrino) then
%>
        <td valign="top">
        <strong>- Sistemas a los que afecta:</strong><br/>
        <ul>
<%
          if (cardiocirculatorio) then response.write "<li>Cardiocirculatorio</li>" end if
          if (respiratorio) then response.write "<li>Respiratorio</li>" end if
          if (reproductivo) then response.write "<li>Reproductivo</li>" end if
          if (musculo_esqueletico) then response.write "<li>Musculoesquelético</li>" end if
          if (sistema_inmunitario) then response.write "<li>Inmunitario</li>" end if
          if (higado_gastrointestinal) then response.write "<li>Gastrointestinal - Hígado</li>" end if
          if (sistema_endocrino) then response.write "<li>Endocrino</li>" end if
%>
        </ul>
        </td>
<%
    end if

    if (embrion OR cancer OR rinyon OR piel_sentidos OR neuro_toxicos) then
%>
        <td valign="top">
        <strong>- Otros efectos:</strong><br />
        <ul>
<%
          if (embrion) then response.write "<li>Daños en el embrión</li>" end if
          if (cancer) then response.write "<li>Cáncer</li>" end if
          if (rinyon) then response.write "<li>Daños en el riñón</li>" end if
          if (piel_sentidos) then response.write "<li>Piel y mucosas</li>" end if
          if (neuro_toxicos) then response.write "<li>Efectos neurotóxicos</li>" end if
%>
        </ul>
        </td>
<%
    end if
	end if
%>
        </tr>
      </table>
			</td>
		</tr>
	</table>
  <br />
	<!-- Fin salud -->
<%
	end if
	objRst.close()
	set objRst = nothing
end sub

' #############################################################################################
' Obtiene el nivel cancerígeno de los campos de clasificación
function dame_nivel_cancerigeno_rd()
	' Juntamos todas las clasificaciones
	clasificacion_rd = substance.Item("clasificacion_1") & substance.Item("clasificacion_2") & substance.Item("clasificacion_3") & substance.Item("clasificacion_4") & substance.Item("clasificacion_5") & substance.Item("clasificacion_6") & substance.Item("clasificacion_7") & substance.Item("clasificacion_8") & substance.Item("clasificacion_9") & substance.Item("clasificacion_10") & substance.Item("clasificacion_11") & substance.Item("clasificacion_12") & substance.Item("clasificacion_13") & substance.Item("clasificacion_14") & substance.Item("clasificacion_15")

	' Sustituimos "Carc. Cat." por "Carc.Cat." para unificar
	clasificacion_rd = replace(clasificacion_rd, "Carc. Cat.", "Carc.Cat.")

	' Quitamos los espacios en blanco
	clasificacion_rd = replace(clasificacion_rd, " ", "")

	' Buscamos la primera aparicion de "Carc.Cat."
	posicion = instr(1,clasificacion_rd, "Carc.Cat.")

	' Sacamos el nivel como el caracter que hay justo detrás de la primera aparición de la subcadena

	if (posicion > 0) then
		dame_nivel_cancerigeno_rd = mid(clasificacion_rd, posicion+9, 1)
	else
		dame_nivel_cancerigeno_rd = ""
	end if
end function

' #############################################################################################

function dame_nivel_cancerigeno_danesa()
	' Buscamos la primera aparicion de "Carc"
	posicion = instr(1,substance.Item("frases_r_danesa"), "Carc")

	' Sacamos el nivel como el caracter que hay justo detrás de la primera aparición de la subcadena
	if (posicion > 0) then
		dame_nivel_cancerigeno_danesa = mid(substance.Item("frases_r_danesa"), posicion+4, 1)
	else
		dame_nivel_cancerigeno_danesa = ""
	end if
end function

' #############################################################################################

function dame_nivel_mutageno_rd()
	' Juntamos todas las clasificaciones
	clasificacion_rd = substance.Item("clasificacion_1") & substance.Item("clasificacion_2") & substance.Item("clasificacion_3") & substance.Item("clasificacion_4") & substance.Item("clasificacion_5") & substance.Item("clasificacion_6") & substance.Item("clasificacion_7") & substance.Item("clasificacion_8") & substance.Item("clasificacion_9") & substance.Item("clasificacion_10") & substance.Item("clasificacion_11") & substance.Item("clasificacion_12") & substance.Item("clasificacion_13") & substance.Item("clasificacion_14") & substance.Item("clasificacion_15")

	' Sustituimos "Muta. Cat." por "Muta.Cat." para unificar
	clasificacion_rd = replace(clasificacion_rd, "Muta. Cat.", "Muta.Cat.")

	' Quitamos los espacios en blanco
	clasificacion_rd = replace(clasificacion_rd, " ", "")

	'response.write "["&clasificacion_rd&"]"

	' Buscamos la primera aparicion de "Muta.Cat."
	posicion = instr(1,clasificacion_rd, "Muta.Cat.")

	' Sacamos el nivel como el caracter que hay justo detrás de la primera aparición de la subcadena
	if (posicion > 0) then
		dame_nivel_mutageno_rd = mid(clasificacion_rd, posicion+9, 1)
	else
		dame_nivel_mutageno_rd = ""
	end if
end function

' #############################################################################################

function dame_nivel_mutageno_danesa()
	' Buscamos la primera aparicion de "Mut"
	posicion = instr(1,substance.Item("frases_r_danesa"), "Mut")

	' Sacamos el nivel como el caracter que hay justo detrás de la primera aparición de la subcadena
	if (posicion > 0) then
		dame_nivel_mutageno_danesa = mid(substance.Item("frases_r_danesa"), posicion+3, 1)
	else
		dame_nivel_mutageno_danesa = ""
	end if
end function

' #############################################################################################

function dame_nivel_reproduccion_rd()
	' Juntamos todas las clasificaciones
	clasificacion_rd = substance.Item("clasificacion_1") & substance.Item("clasificacion_2") & substance.Item("clasificacion_3") & substance.Item("clasificacion_4") & substance.Item("clasificacion_5") & substance.Item("clasificacion_6") & substance.Item("clasificacion_7") & substance.Item("clasificacion_8") & substance.Item("clasificacion_9") & substance.Item("clasificacion_10") & substance.Item("clasificacion_11") & substance.Item("clasificacion_12") & substance.Item("clasificacion_13") & substance.Item("clasificacion_14") & substance.Item("clasificacion_15")

	' Sustituimos "Repr. Cat." por "Repr.Cat." para unificar
	clasificacion_rd = replace(clasificacion_rd, "Repr. Cat.", "Repr.Cat.")

	' Quitamos los espacios en blanco
	clasificacion_rd = replace(clasificacion_rd, " ", "")

	'response.write "["&clasificacion_rd&"]"

	' Buscamos la primera aparicion de "Repr.Cat."
	posicion = instr(1,clasificacion_rd, "Repr.Cat.")

	' Sacamos el nivel como el caracter que hay justo detrás de la primera aparición de la subcadena
	if (posicion > 0) then
		dame_nivel_reproduccion_rd = mid(clasificacion_rd, posicion+9, 1)
	else
		dame_nivel_reproduccion_rd = ""
	end if
end function

' #############################################################################################

function dame_nivel_reproduccion_danesa()
	' Buscamos la primera aparicion de "Repr.Cat."
	posicion = instr(1,sus_frases_r_danesa, "Repr.Cat.")

	' Sacamos el nivel como el caracter que hay justo detrás de la primera aparición de la subcadena
	if (posicion > 0) then
		dame_nivel_reproduccion_danesa = mid(clasificacion_rd, posicion+9, 1)
	else
		dame_nivel_reproduccion_danesa = ""
	end if
end function


' #############################################################################################

sub plegador(byval id_bloque, byval id_imagen)
  ' Pinta el HTML necesario para las llamadas a mostrar/ocultar el objeto, y a cambiar la imagen
  id_bloque=aplana(id_bloque)
  id_imagen=aplana(id_imagen)
%>
  <a href="javascript:toggle('<%= id_bloque %>', '<%= id_imagen %>');"><img src="imagenes/desplegar.gif" align="absmiddle" id="<%= id_imagen %>" alt="Pulse para desplegar la información" title="Pulse para desplegar la información" /></a>
<%
end sub

' #############################################################################################

sub plegador_texto(byval id_bloque, byval texto, byval clase)
  ' Pinta el HTML necesario para las llamadas a mostrar/ocultar el objeto
  ' Solo se emplea para el plegador de frases R danesas, en caso de que no se hayan mostrado ya.
  id_bloque=aplana(id_bloque)

  if (mySubstance.hasFrasesRdanesa()) then
%>
  <%=texto%>
<%
  else
%>
  <a href="javascript:toggle_texto('<%= id_bloque %>');" class="<%=clase%>"><%=texto%></a>
<%
  end if
end sub

' #############################################################################################

function aplana(byval cadena)
  cadena = quitartildes(cadena)
  cadena = replace(cadena, " ", "")
  aplana = cadena
end function
%>
