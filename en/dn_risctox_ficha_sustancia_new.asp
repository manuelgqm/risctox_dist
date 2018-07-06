<%@ LANGUAGE="VBSCRIPT" LCID="1034" CODEPAGE="65001"%>
<!--#include file="../dn_restringida.asp"-->
<!--#include file="../config/dbConnection.asp"-->
<!--#include file="../lib/dn_funciones_texto_utf-8.asp"-->
<!--#include file="../lib/dn_funciones_comunes_utf-8.asp"-->
<!--#include file="../lib/class/SubstanceInternationalClass.asp"-->
<!--#include file="../lib/visitsRecorder.asp"-->
<!--#include file="../lib/urlManipulations.asp"-->

<%
Response.ContentType = "text/html"
Response.AddHeader "Content-Type", "text/html;charset=UTF-8"
Response.CodePage = 65001
Response.CharSet = "UTF-8"

dim idpagina : idpagina = 627
call recordVisit(idpagina)

dim LANG : LANG = "en"
Dim CARCINOGENIC_LISTS : CARCINOGENIC_LISTS = Array( _
  "cancer_rd", _
  "cancer_danesa", _
  "cancer_iarc", _
  "cancer_otras", _
  "cancer_mama" _
)
dim MUTAGENIC_LISTS: MUTAGENIC_LISTS = Array( _
  "mutageno_rd", _
  "mutageno_danesa" _
)
dim NEUROTOXIC_LISTS : NEUROTOXIC_LISTS = array( _
  "neurotoxico", _
  "neurotoxico_rd", _
  "neurotoxico_danesa", _
  "neurotoxico_nivel" _
)
dim SENTITISER_LISTS : SENTITISER_LISTS = Array( _
  "sensibilizante", _
  "sensibilizante_danesa", _
  "sensibilizante_reach" _
)
dim TPR_LISTS : TPR_LISTS = Array("tpr", "tpr_danesa")
dim HEALTH_EFFECTS_LISTS : HEALTH_EFFECTS_LISTS = array( _
  "cancer_rd", _
  "cancer_danesa", _
  "cancer_iarc", _
  "cancer_otras", _
  "cancer_mama", _
  "de", _
  "sensibilizante", _
  "sensibilizante_reach", _
  "sensibilizante_danesa", _
  "tpr", _
  "tpr_danesa", _
  "eepp", _
  "mutageno_rd", _
  "mutageno_danesa", _
  "salud", _
  "prohibidas_embarazadas", _
  "prohibidas_lactantes", _
  "neurotoxico", _
  "neurotoxico_rd", _
  "neurotoxico_danesa", _
  "neurotoxico_nivel" _
)

dim id_sustancia : id_sustancia = obtainSanitizedQueryParameter("id_sustancia")
dim substance : set substance = (new SubstanceClassInternational)(id_sustancia, LANG, objConnection2)
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="en" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>RISCTOX: Toxic and hazardous substances database</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="Title" content="Risctox" />
<meta name="Author" content="SPL Sistemas de Informaci&oacute;n - www.spl-ssi.com" />
<meta name="description" content="RISCTOX: Toxic and hazardous substances database" />
<meta name="Subject" content="RISCTOX: Toxic and hazardous substances database" />
<meta name="Keywords" content="RISCTOX: Toxic and hazardous substances database" />
<meta name="Language" content="English" />
<meta name="Revisit" content="15 days" />
<meta name="Distribution" content="Global" />
<meta name="Robots" content="All" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />

<link rel="stylesheet" type="text/css" href="../estructura.css">
<link rel="stylesheet" type="text/css" href="../dn_estilos.css">
<link rel="stylesheet" type="text/css" href="css/en.css">

<script type="text/javascript" src="https://ajax.googleapis.com/ajax/libs/prototype/1.7.0.0/prototype.js"></script>
<script type="text/javascript" src="https://ajax.googleapis.com/ajax/libs/scriptaculous/1.9.0/scriptaculous.js"></script>

<script type="text/javascript">
function toggle(id_objeto, id_imagen)
{
    if (Element.visible(id_objeto))
    {
      $(id_imagen).src="../imagenes/desplegar.gif";
    }
    else
    {
      $(id_imagen).src="../imagenes/plegar.gif";
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
          <table width="100%" border="0">
            <tr>
              <td>
                <a href="http://www.etuc.org/a/6023" target="_blank"><b>Trade Union priority list for REACH authorization</b></a>
              </td>
            	<td align="right">
                <input type="button" name="volver" class="boton2" value="New search" onClick="window.location='dn_risctox_buscador.asp';">
              </td>
            </tr>
            <tr>
            	<td>
                <p class=campo>You are in: <a href="dn_risctox_buscador.asp">Risctox</a> &gt; Substance card</p>
              </td>
            	<td align="right"></td>
            </tr>
          </table>
          <div id="ficha">
          	<!-- ################ Identificacion de la sustancia ###################### -->
          	<table width="100%" cellpadding=5>
          		<tr>
          			<td>
          				<a name="identificacion"></a><img src="imagenes/risctox01.gif" alt="Substance identification" width="255" height="32" />
          			</td>
          			<td align="right">
          				<a href="#"><img src="imagenes/subir.gif" border=0 alt=subir></a>
          			</td>
          		</tr>
          	</table>

          	<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
          		<% ap1_identificacion(LANG) %>
          	</table>

          	<div style="height:3pt"></div>
          		<%
                ' ap2_clasificacion()
              %>
          	<br />
          	<div style="height:3pt"></div>
          	 <%
             ap2_clasificacion_rd1272()
             %>

          	<br />
          </div>
          <%
          ap3_riesgos()
          ap4_normativa_ambiental()
          ap4_normativa_restriccion_prohibicion()
          'ap5_alternativas()
          'ap6_sectores()
          %>

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

<!--#include file="../../cookie_accept.asp" -->
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

function formatHtmlIcscList(icsc_nums)
  dim i, result
  result = ""

  for i = 0 to ubound(icsc_nums)
    icsc_num = icsc_nums(i).item("id")
    result = result & "<a href='http://www.ilo.org/dyn/icsc/showcard.display?p_lang=en&p_card_id='" & icsc_num & " target='_blank'>" & icsc_num & "</a>"
  next

  formatHtmlIcscList = result
end function

sub ap1_identificacion(LANG)
  Dim nombre_field_name : nombre_field_name = "nombre"

  if LANG = "en" then
    nombre_field_name = "nombre_ing"
  end if

  nombres = split(espaciar(substance.identification.item(nombre_field_name)), "@")
	nombre = nombres(0)

	sinonimos=""
	if UBound(nombres) > 0 then
		sinonimos = "<ul>"
		For i = LBound(nombres) + 1 To UBound(nombres)
			sinonimos = sinonimos & "<li>" & h(espaciar(nombres(i))) & "</li>"
		Next
		sinonimos = sinonimos & "</ul>"
	end if
%>
	<tr>
		<td class="subtitulo3" align="right" valign="top">
			<a onclick=window.open('ver_definicion.asp?id=82','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a>&nbsp;<span id="name.label">Chemical name:</a>
		</td>
		<td class="texto" valign="middle">
			<b><span id="name.value"><%=nombre%></div></b>
		</td>
	</tr>

	<%
	if (sinonimos<>"") then
	%>
		<tr>
			<td class="subtitulo3" align="right" valign="top">
				<a onclick=window.open('ver_definicion.asp?id=83','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a>&nbsp;<span id="synonyms.label">Synonyms:</span>
			</td>
			<td class="texto" valign="middle">
				<span id="synonyms.value"><%=sinonimos%></span>
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
				<span id="trade_name.label">Trade name</span>:
			</td>
			<td class="texto" valign="middle">
				<span id="trade_name.value"><%=nombre_comercial%></span>
			</td>
		</tr>
	<% end if ' hay nombre comercial? %>

	<% if (substance.identification.Item("num_cas") <> "") or (substance.identification.Item("num_ce_einecs") <> "") or (substance.identification.Item("num_ce_elincs") <> "") or not is_empty(substance.identification.item("cas_num_alternatives")) then %>
		<tr>
			<td class="subtitulo3" align="right" valign="top">
				<span id="identification_numbers.label">Identification numbers:</span>
			</td>
			<td class="texto" valign="middle">
				<%
        if not is_empty(substance.identification.item("num_cas")) then
          response.write "<a onclick=window.open('ver_definicion.asp?id=84','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a> <b><span id='cas_num.label'>CAS</span></b>:&nbsp;<span id='cas_num.value'>" & substance.identification.item("num_cas") & "</span><br/>"
        end if
        if not is_empty(substance.identification.item("cas_num_alternatives")) then
          response.write _
            "<a onclick=window.open('ver_definicion.asp?id=84', 'def', 'width=300, height=200, scrollbars=yes, resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a>" &_
          "<b>&nbsp;<span id='cas_num_alternatives.label'>Alternative CAS</span></b>:&nbsp;" &_
          "<span id='cas_num_alternatives.value'>" &_
            substance.identification.item("cas_num_alternatives") &_
          "</span><br/>"
        end if
					if (substance.identification.item("num_ce_einecs") <> "") then
						'Sergio, si empieza por 4 y num_ce_elincs<>'' muestro el num_ce_elincs
						if (mid(num_ce_einecs, 1, 1) = "4" and substance.identification.item("num_ce_elincs") <> "") then
							response.write "<a onclick=window.open('ver_definicion.asp?id=85','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a> <b><span id='ec_elincs_num.label'>EC ELINCS</span></b>:&nbsp;<span id='ec_elincs_num.value'>" & substance.identification.item("num_ce_elincs") & "</span><br/>"
						else
						response.write "<a onclick=window.open('ver_definicion.asp?id=85','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a> <b><span id='ec_einecs_num.label'>EC EINECS</span></b>:&nbsp;<span id='ec_einecs_num.value'>" & substance.identification.item("num_ce_einecs") & "</span><br/>"
						end if
					elseif substance.identification.item("num_ce_elincs") <> "" then
						response.write "<a onclick=window.open('ver_definicion.asp?id=85','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a> <b><span id='ec_elincs_num.label'>EC ELINCS</span></b>:&nbsp;<span id='ec_elincs_num.value'>" & substance.identification.item("num_ce_elincs") & "</span><br/>"
					end if
				%>
			</td>
		</tr>
	<% end if ' hay numeros? %>

	<%
		grupos = formatHtmlGlossaryLinksString(substance.identification.item("grupos"), "grupos")
		if (grupos <> "") then
	%>
		<tr>
			<td class="subtitulo3" align="right" valign="top">
				<span id="groups.label">Groups</span>:
			</td>
			<td class="texto" valign="middle">
				<span id="groups.value"><%=grupos%></span>
			</td>
		</tr>
	<% end if ' hay grupos? %>

	<%
		usos = formatHtmlGlossaryLinksString(substance.identification.item("applications"), "usos")
		if (usos <> "") then
	%>
		<tr>
			<td class="subtitulo3" align="right" valign="top">
				<span id="uses.label">Uses</span>:
			</td>
			<td class="texto" valign="middle">
				<span id="uses.value"><%=usos%></span>
			</td>
		</tr>
	<% end if %>

	<%
		if UBound(substance.identification.item("icsc_nums")) >= 0 then
	%>
		<tr>
			<td class="subtitulo3" align="right" valign="top">
				 <span id="icsc_nums.label">
           International Chemical Safety Card (<a onClick="window.open('ver_definicion.asp?id=<%=dame_id_definicion("ICSC")%>', 'def', 'width=300,height=200,scrollbars=yes,resizable=yes')" class="subtitulo3">ICSC</a>)
         <span>
			</td>
			<td class="texto" valign="middle">
        <span id="icsc_nums.value">
        <%
        icsc_nums_list = formatHtmlIcscList(substance.identification.item("icsc_nums"))
        response.write icsc_nums_list
        %>
        </span>
			</td>
		</tr>
	<% end if %>

	<%
  if (substance.identification.item("nombre_ing") <> "") or (substance.identification.item("num_rd") <> "") or (substance.identification.item("formula_molecular") <> "") or (substance.identification.item("estructura_molecular") <> "") or (substance.identification.item("notas_xml") <> "") or not is_empty(substance.identification.item("compañias")) then
  %>
		<tr>
			<td class="subtitulo3" align="right" valign="top" width="35%">
				<span id="additional_information.label">Additional information</span>&nbsp;<% plegador "secc-masinformacion", "img-masinformacion" %>
			</td>
			<td class="texto" valign="middle" id="secc-masinformacion" style="display:none">
				<% if (substance.identification.item("num_rd") <> "") then %>
          <a onclick = window.open('ver_definicion.asp?id=86','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'>
          <img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a>
          &nbsp;<b><span id="rd_num.label">Index No</span></b>:
          &nbsp;<span id="rd_num.value"><%= substance.identification.item("num_rd") %></span><br/>
        <% end if %>
				<% if (substance.identification.item("molecular_formula") <> "") then %>
          <b><span id="molecular_formula.label">Molecular formula</span></b>:
          <span id="molecular_formula.value"><%= substance.identification.item("molecular_formula") %><br/>
        <% end if %>
				<% if (substance.identification.item("notas_xml") <> "") then %>
          <a onClick="window.open('ver_definicion.asp?id=<%=dame_id_definicion("ECB")%>', 'def', 'width=300,height=200,scrollbars=yes,resizable=yes')" style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>
          <b>CLP Notes</b>: <%= espaciar(substance.identification.item("notas_xml")) %> <br />
        <% end if %>

        <% if not is_empty(substance.identification.item("compañias")) then %>
          <b><span id="companies.label">Distribution companies</span></b>:
          <span id="companies.value">
            <%= formatHtmlCompaniesLinksString(substance.identification.item("compañias")) %>
          </span>
        <% end if %>
			</td>
		</tr>
	<% end if
%>
	<tr>
		<td valign="top" colspan="2">
			<% concern_trade_union_list(mySubstance) %>
		</td>
	</tr>
<%
end sub ' ap1_identificacion

' ###################################################################################

sub ap2_clasificacion()
	' Solo mostramos este apartado si hay información para él
	if ((substance.classification.item("simbolos") <> "") or (substance.classification.item("clasificacion_1") <> "") or (substance.classification.item("clasificacion_2") <> "") or (substance.classification.item("clasificacion_3") <> "") or (substance.classification.item("clasificacion_4") <> "") or (substance.classification.item("clasificacion_5") <> "") or (substance.classification.item("clasificacion_6") <> "") or (substance.classification.item("clasificacion_7") <> "") or (substance.classification.item("clasificacion_8") <> "") or (substance.classification.item("clasificacion_9") <> "") or (substance.classification.item("clasificacion_10") <> "") or (substance.classification.item("clasificacion_11") <> "") or (substance.classification.item("clasificacion_12") <> "") or (substance.classification.item("clasificacion_13") <> "") or (substance.classification.item("clasificacion_14") <> "") or (substance.classification.item("clasificacion_15") <> "") or (substance.classification.item("frases_r_danesa") <> "") or (substance.classification.item("notas_rd_363") <> "") or (substance.classification.item("conc_1") <> "") or (substance.classification.item("eti_conc_1") <> "") or (substance.classification.item("conc_2") <> "") or (substance.classification.item("eti_conc_2") <> "") or (substance.classification.item("conc_3") <> "") or (substance.classification.item("eti_conc_3") <> "") or (substance.classification.item("conc_4") <> "") or (substance.classification.item("eti_conc_4") <> "") or (substance.classification.item("conc_5") <> "") or (substance.classification.item("eti_conc_5") <> "") or (substance.classification.item("conc_6") <> "") or (substance.classification.item("eti_conc_6") <> "") or (substance.classification.item("conc_7") <> "") or (substance.classification.item("eti_conc_7") <> "") or (substance.classification.item("conc_8") <> "") or (substance.classification.item("eti_conc_8") <> "") or (substance.classification.item("conc_9") <> "") or (substance.classification.item("eti_conc_9") <> "") or (substance.classification.item("conc_10") <> "") or (substance.classification.item("eti_conc_10") <> "") or (substance.classification.item("conc_11") <> "") or (substance.classification.item("eti_conc_11") <> "") or (substance.classification.item("conc_12") <> "") or (substance.classification.item("eti_conc_12") <> "") or (substance.classification.item("conc_13") <> "") or (substance.classification.item("eti_conc_13") <> "") or (substance.classification.item("conc_14") <> "") or (substance.classification.item("eti_conc_14") <> "") or (substance.classification.item("conc_15") <> "") or (substance.classification.item("eti_conc_15") <> "") ) then

%>
	<!-- ################ Clasificación ###################### -->
	<table id="tabla_clasificacionm" class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
  <tr>
		<td class="celdaabajo" colspan="2" align="center">
			<table cellpadding=0 cellspacing=0 width="100%" border="0"><tr><td width="100%" class="titulo3" align="left"><a onclick=window.open('ver_definicion.asp?id=87','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> CLASIFICACIÓN (RD 363/1995)
			<a href="javascript:toggle('secc-clasificacion-363', 'img-mas_clasificacion-363');"><img src="../imagenes/desplegar.gif" align="absmiddle" id="img-mas_clasificacion-363" alt="Click for more information" title="Click for more information" /></a>
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

		        if substance.hasFrasesRdanesa() then
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
end sub

sub ap2_clasificacion_rd1272()
  if is_empty(substance.classification.item("pictogramasRd1272")) and is_empty(substance.classification.item("frasesH")) and is_empty(substance.classification.item("concentracionEtiquetadoRd1272")) then
    exit Sub
  end if

%>
	<table id="tabla_clasificacionm_rd1272" class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
    <tr><td class="celdaabajo" colspan="2" align="center">
			<table cellpadding=0 cellspacing=0 width="100%" border="0"><tr><td width="100%" class="titulo3" align="left"><a onclick=window.open('ver_definicion.asp?id=280','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>
      &nbsp;<span id="1272_classification.label">CLASSIFICATION AND LABELLING (Regulation 1272/2008)</span>
			<a href="javascript:toggle('secc-clasificacion-rd1272', 'img-mas_clasificacion-rd1272');"><img src="../imagenes/desplegar.gif" align="absmiddle" id="img-mas_clasificacion-rd1272" alt="Click for more information" title="Click for more information" /></a>
			</td></tr></table>
		</td></tr>

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
end sub ' ap2_clasificacion

' ##################################################################################

sub ap2_clasificacion_simbolos()
	if (substance.classification.item("simbolos") <> "") then
%>
		<p id="ap2_clasificacion_simbolos_titulo" class="ficha_titulo_2">Símbolos</p>
		<p id="ap2_clasificacion_simbolos_cuerpo" class="texto" align="center">
<%
		' Tiene símbolos, muestro cada uno
		substance.classification.item("simbolos") = replace(substance.classification.item("simbolos"), ",", ";")
		array_simbolos = split(substance.classification.item("simbolos"), ";")
		for i=0 to ubound(array_simbolos)
			simbolo = trim(array_simbolos(i))
			imagen = imagen_simbolo(simbolo)
			descripcion = get_symbol_description(simbolo, lang)
      if (trim(simbolo) <> "") then
%>
			<img src="../imagenes/pictogramas/<%= imagen %>" title="<%= simbolo %>; <%= descripcion %>" width="75px" /><br/>
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
	if is_empty(substance.classification.item("pictogramasRd1272")) then
    exit sub
  end if

  simbolos = replace(substance.classification.Item("simbolos_rd1272"), ",", ";")
	array_simbolos = split(simbolos, ";")
  symbols_html = ""
	for i=0 to ubound(array_simbolos)
		simbolo = trim(array_simbolos(i))
		imagen = ""
		descripcion = ""
		if left(simbolo, 3) = "GHS" then
			imagen = imagen_simbolo(simbolo)
			descripcion = get_symbol_description(simbolo, lang)
		else ' Peligro
			descripcion = "<b style='background-color:red;color:#FFF;'>" + traduceSimbolo(simbolo) + "</b>"
		end if
		if imagen <> "" then
      symbols_html = symbols_html & "<img src='../imagenes/pictogramas/" & imagen  & "' title='" & simbolo & "; " & descripcion & "'' width='75px' /><br/>"
		end if
    symbols_html = symbols_html & descripcion & "<br/><br/>"
	next

  %>
  <p id="ap2_clasificacion_simbolos_titulo" class="ficha_titulo_2">Pictograms and signal words</p>
	<p id="ap2_clasificacion_simbolos_cuerpo" class="texto" align="center">
    <span id="rd1272_symbols"><%= symbols_html %></span>
	</p>
  <%
end sub ' ap2_clasificacion_simbolos_rd1272

' ##################################################################################

sub ap2_clasificacion_frases_r(substance)
	' Muestra las frases R segun clasificacion_1 hasta clasificacion_15
	' No incluye las frases R danesas

	' Montamos frases R

	if (substance.classification.item("frasesR") <> "") then
%>
		<p id="ap2_clasificacion_frases_r_titulo" class="ficha_titulo_2" style="margin-bottom: -10px;"><a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("Frases R")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> Frases R</p>
<%
		bucle_frases "r", substance.classification.item("frasesR")
	end if
end sub

' ##################################################################################

sub ap2_clasificacion_frases_h()
	if is_empty(substance.classification.item("frasesH")) then _
    exit sub
  %>
	<p id="ap2_clasificacion_frases_r_titulo" class="ficha_titulo_2" style="margin-bottom: -10px;"><a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("Frases H")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> H-phrases</p>
  <span id="H_phrases">
  <%
	muestra_clasificacion 1, substance.classification.item("clasificacion_rd1272_1")
	muestra_clasificacion 2, substance.classification.Item("clasificacion_rd1272_2")
	muestra_clasificacion 3, substance.classification.Item("clasificacion_rd1272_3")
	muestra_clasificacion 4, substance.classification.Item("clasificacion_rd1272_4")
	muestra_clasificacion 5, substance.classification.Item("clasificacion_rd1272_5")
	muestra_clasificacion 6, substance.classification.Item("clasificacion_rd1272_6")
	muestra_clasificacion 7, substance.classification.Item("clasificacion_rd1272_7")
	muestra_clasificacion 8, substance.classification.Item("clasificacion_rd1272_8")
	muestra_clasificacion 9, substance.classification.Item("clasificacion_rd1272_9")
	muestra_clasificacion 10, substance.classification.Item("clasificacion_rd1272_10")
	muestra_clasificacion 11, substance.classification.Item("clasificacion_rd1272_11")
	muestra_clasificacion 12, substance.classification.Item("clasificacion_rd1272_12")
	muestra_clasificacion 13, substance.classification.Item("clasificacion_rd1272_13")
	muestra_clasificacion 14, substance.classification.Item("clasificacion_rd1272_14")
	muestra_clasificacion 15, substance.classification.Item("clasificacion_rd1272_15")

	' 23/06/2014 - SPL - Por indicación de Tatiana se pone esta definición.
	if (trim(substance.classification.Item("clasificacion_rd1272_1"))="Expl., ****;") then
		%>
		<p><b>Explosive</b>: Physical hazards that need to be confirmed by testing</p>
		<%
	end if
  %></span><%

end sub

' ##################################################################################

sub muestra_clasificacion(numero, clasificacion)
	if is_empty(clasificacion) then
    exit sub
  end if

	array_clasificacion = split(clasificacion, ";")
	clas_cat_peligro = trim(array_clasificacion(0))
	if ubound(array_clasificacion)>0 then
		frase = trim(array_clasificacion(1))
	end if
  descripcion = describe_frase_international("h", replace(frase, "*", ""), lang)
  frase = buscaDefinicionAsteriscos(frase)

  %>
  <blockquote style="margin-left: 10px; margin-bottom: -20px;">
  <%
  ' Las frases H??? son Gases a presión. Cambio solicitado por Tatiana en abril 2012
	if (frase = "H???") then
    %><b>Gases under pressure</b><%
	else
  %>
  <b><%= frase %></b>: <%= descripcion %>
  <br/>
	<blockquote style="margin-left: 30px; margin-top: 12px;" id="secc-categpeligro-<%=numero%>">
    <% muestra_frase_clasificacion_rd1272 clas_cat_peligro %>
  </blockquote>
  <% end if %>
	</blockquote>

  <br clear="all" />
<%
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
  if is_empty(frase) then
    exit sub
  end if
	arrFrase = split(frase, ",")
	frase_descripcion = describe_categoria_peligro_international(arrFrase(0), lang)
	frase = frase_descripcion(0)
  descripcion = frase_descripcion(1)
	if (ubound(arrFrase)>0)then
		categoria = "Cat. " + arrFrase(1)
	else
		categoria = ""
	end if

  response.write "<b>" & frase & " (" & buscaDefinicionAsteriscos(categoria) & ")</b>:&nbsp;"
  response.write descripcion & "<br />"

end sub


' ##################################################################################

sub ap2_clasificacion_frases_r_danesa(substance)
	' Muestra las frases R danesas

	if (substance.classification.item("frases_r_danesa") <> "") then
%>
	<p id="ap2_clasificacion_frases_r_danesa_titulo" class="ficha_titulo_2"><a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("Frases R según la lista danesa de la EPA")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> Frases R según clasificación de la EPA danesa</p>
<%
		bucle_frases "r", joinFrasesRDanesa(substance.classification.item("frases_r_danesa"))
	end if
end sub


' ##################################################################################

sub ap2_clasificacion_frases_s
	' Muestra las frases S

	if (substance.classification.item("frases_s") <> "") then
		' Eliminamos los paréntesis de las frases S
		frases_s = replace (substance.classification.item("frases_s"), "(", "")
		frases_s = replace (frases_s, ")", "")

%>
	<p id="ap2_clasificacion_frases_s_titulo" class="ficha_titulo_2" style="margin-top: 14px;"><a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("Frases S")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> Frases S <% plegador "secc-frasess", "img-frasess" %></p>

		<% bucle_frases_s frases_s%>

<%
	end if
end sub

' ##################################################################################

sub ap2_clasificacion_notas()
	if (substance.classification.item("notas_rd_363") <> "") then

		' Dividimos las notas, separadas por puntos, en un array
		array_notas = split(substance.classification.item("notas_rd_363"), ".")
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
	if is_empty(substance.classification.item("notas_rd1272")) then _
		exit sub
	%>
	<p id="ap2_clasificacion_notas_titulo" class="ficha_titulo_2">Notes&nbsp;
	<p class="texto" >
		<blockquote id="secc-notas-rd1272">
    <span id="rd1272_notes">
		<%
		for i = 0 to ubound(substance.classification.item("notas_rd1272"))
			set nota = substance.classification.Item("notas_rd1272")(i)
		%>
			<b><%= replace(nota.item("key"), "Nota", "Note") %></b>&nbsp;
			<% if nota.item("id")<>""then %>
			<a onclick=window.open('ver_definicion.asp?id=<%=nota.item("id")%>','def','width=600,height=400,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a><br />
			<%end if%>
		<% next %>
    </span>
		</blockquote>
	</p>
	<%
end sub

' ##################################################################################

sub ap2_clasificacion_etiquetado()
	' Muestra el etiquetado

	if ((substance.classification.item("conc_1") <> "") or (substance.classification.item("eti_conc_1") <> "") or (substance.classification.item("conc_2") <> "") or (substance.classification.item("eti_conc_2") <> "") or (substance.classification.item("conc_3") <> "") or (substance.classification.item("eti_conc_3") <> "") or (substance.classification.item("conc_4") <> "") or (substance.classification.item("eti_conc_4") <> "") or (substance.classification.item("conc_5") <> "") or (substance.classification.item("eti_conc_5") <> "") or (substance.classification.item("conc_6") <> "") or (substance.classification.item("eti_conc_6") <> "") or (substance.classification.item("conc_7") <> "") or (substance.classification.item("eti_conc_7") <> "") or (substance.classification.item("conc_8") <> "") or (substance.classification.item("eti_conc_8") <> "") or (substance.classification.item("conc_9") <> "") or (substance.classification.item("eti_conc_9") <> "") or (substance.classification.item("conc_10") <> "") or (substance.classification.item("eti_conc_10") <> "") or (substance.classification.item("conc_11") <> "") or (substance.classification.item("eti_conc_11") <> "") or (substance.classification.item("conc_12") <> "") or (substance.classification.item("eti_conc_12") <> "") or (substance.classification.item("conc_13") <> "") or (substance.classification.item("eti_conc_13") <> "") or (substance.classification.item("conc_14") <> "") or (substance.classification.item("eti_conc_14") <> "") or (substance.classification.item("conc_15") <> "") or (substance.classification.item("eti_conc_15") <> "")) then

%>
	<span id="ap2_clasificacion_etiquetado_titulo" class="ficha_titulo_2"><a onclick=window.open('ver_definicion.asp?id=88','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> Etiquetado <% plegador "secc-etiquetado", "img-etiquetado" %></span>


  <fieldset id="secc-etiquetado" style="display:none; margin: 15px 45px;">
	<table cellspacing="0" cellpadding="3" width="100%" align="center">
		<tr>
			<th class="subtitulo3 celdaabajo">Concentración</th><th class="subtitulo3 celdaabajo">Etiquetado</th>
		</tr>
<%
	ap2_clasificacion_etiquetado_fila	"r", substance.classification.item("conc_1"), substance.classification.item("eti_conc_1")
	ap2_clasificacion_etiquetado_fila	"r", substance.classification.item("conc_2"), substance.classification.item("eti_conc_2")
	ap2_clasificacion_etiquetado_fila	"r", substance.classification.item("conc_3"), substance.classification.item("eti_conc_3")
	ap2_clasificacion_etiquetado_fila	"r", substance.classification.item("conc_4"), substance.classification.item("eti_conc_4")
	ap2_clasificacion_etiquetado_fila	"r", substance.classification.item("conc_5"), substance.classification.item("eti_conc_5")
	ap2_clasificacion_etiquetado_fila	"r", substance.classification.item("conc_6"), substance.classification.item("eti_conc_6")
	ap2_clasificacion_etiquetado_fila	"r", substance.classification.item("conc_7"), substance.classification.item("eti_conc_7")
	ap2_clasificacion_etiquetado_fila	"r", substance.classification.item("conc_8"), substance.classification.item("eti_conc_8")
	ap2_clasificacion_etiquetado_fila	"r", substance.classification.item("conc_9"), substance.classification.item("eti_conc_9")
	ap2_clasificacion_etiquetado_fila	"r", substance.classification.item("conc_10"), substance.classification.item("eti_conc_10")
	ap2_clasificacion_etiquetado_fila	"r", substance.classification.item("conc_11"), substance.classification.item("eti_conc_11")
	ap2_clasificacion_etiquetado_fila	"r", substance.classification.item("conc_12"), substance.classification.item("eti_conc_12")
	ap2_clasificacion_etiquetado_fila	"r", substance.classification.item("conc_13"), substance.classification.item("eti_conc_13")
	ap2_clasificacion_etiquetado_fila	"r", substance.classification.item("conc_14"), substance.classification.item("eti_conc_14")
	ap2_clasificacion_etiquetado_fila	"r", substance.classification.item("conc_15"), substance.classification.item("eti_conc_15")
%>
	</table>
  </fieldset>

<%
	end if
end sub


' ##################################################################################

sub ap2_clasificacion_etiquetado_rd1272()

  if is_empty(substance.classification.item("concentracionEtiquetadoRd1272")) then _
    exit sub

  %>
	<span id="ap2_clasificacion_etiquetado_titulo" class="ficha_titulo_2"><a onclick=window.open('ver_definicion.asp?id=279','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>&nbsp;Labeling&nbsp;</span>

  <fieldset id="secc-etiquetado_rd1272" style="margin: 15px 45px;">
  <span id="rd1272_labeling">
   <%
  	if not is_empty(substance.classification.Item("conc_rd1272_1")) or not is_empty(substance.classification.Item("conc_rd1272_2")) then
  		if (substance.classification.Item("conc_rd1272_1")) = "" then
  			if substance.classification.Item("eti_conc_rd1272_1") <> "" then
  			   response.write "Factor " & substance.classification.Item("eti_conc_rd1272_1")
  			end if
  		end if

      %>
    	<table cellspacing="0" cellpadding="3" width="100%" align="center">
    		<tr>
    			<th class="subtitulo3 celdaabajo">Concentration</th><th class="subtitulo3 celdaabajo">Labeling</th>
    		</tr>
      <%
    	ap2_clasificacion_etiquetado_fila	"h", substance.classification.Item("conc_rd1272_1"), substance.classification.Item("eti_conc_rd1272_1")
    	ap2_clasificacion_etiquetado_fila	"h", substance.classification.Item("conc_rd1272_2"), substance.classification.Item("eti_conc_rd1272_2")
    	ap2_clasificacion_etiquetado_fila	"h", substance.classification.Item("conc_rd1272_3"), substance.classification.Item("eti_conc_rd1272_3")
    	ap2_clasificacion_etiquetado_fila	"h", substance.classification.Item("conc_rd1272_4"), substance.classification.Item("eti_conc_rd1272_4")
    	ap2_clasificacion_etiquetado_fila	"h", substance.classification.Item("conc_rd1272_5"), substance.classification.Item("eti_conc_rd1272_5")
    	ap2_clasificacion_etiquetado_fila	"h", substance.classification.Item("conc_rd1272_6"), substance.classification.Item("eti_conc_rd1272_6")
    	ap2_clasificacion_etiquetado_fila	"h", substance.classification.Item("conc_rd1272_7"), substance.classification.Item("eti_conc_rd1272_7")
    	ap2_clasificacion_etiquetado_fila	"h", substance.classification.Item("conc_rd1272_8"), substance.classification.Item("eti_conc_rd1272_8")
    	ap2_clasificacion_etiquetado_fila	"h", substance.classification.Item("conc_rd1272_9"), substance.classification.Item("eti_conc_rd1272_9")
    	ap2_clasificacion_etiquetado_fila	"h", substance.classification.Item("conc_rd1272_10"), substance.classification.Item("eti_conc_rd1272_10")
    	ap2_clasificacion_etiquetado_fila	"h", substance.classification.Item("conc_rd1272_11"), substance.classification.Item("eti_conc_rd1272_11")
    	ap2_clasificacion_etiquetado_fila	"h", substance.classification.Item("conc_rd1272_12"), substance.classification.Item("eti_conc_rd1272_12")
    	ap2_clasificacion_etiquetado_fila	"h", substance.classification.Item("conc_rd1272_13"), substance.classification.Item("eti_conc_rd1272_13")
    	ap2_clasificacion_etiquetado_fila	"h", substance.classification.Item("conc_rd1272_14"), substance.classification.Item("eti_conc_rd1272_14")
    	ap2_clasificacion_etiquetado_fila	"h", substance.classification.Item("conc_rd1272_15"), substance.classification.Item("eti_conc_rd1272_15")
      %>
  	</table>
    <%
  	else
  		if substance.classification.Item("eti_conc_rd1272_1")<>"" then
  			response.write "Factor " & substance.classification.Item("eti_conc_rd1272_1")
  		end if
  	end if
    %>
  </span>
  </fieldset>
  <%
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
				<td class="celdaabajo"><%= h(c) %></td><td class="celdaabajo"><%= h(traduceEtiquetado(e)) %> </td>
			</tr>
      <%
	  end if

    exit sub
  end if

	if e = "*" then
    %>
		<tr>
			<td class="celdaabajo" colspan="2">
			This entry has specific concentration limits for acute toxicity according to RD 363/1995 which can not "be matched" with the concentration limits under the CLP Regulation (by reference, see the section on classification labeling (RD 363/1995) of the substance).
			</td>
		</tr>
    <%
	end if
end sub

' ****************
' INICIO DE LISTAS RELACIONADAS
' ****************

sub concern_trade_union_list(mySubstance)
  concern_lists = array("cancer_rd", "cancer_danesa", "cancer_iarc_excepto_grupo_3", "cancer_otras", "de", "neurotoxico", "tpb", "sensibilizante", "sensibilizante_danesa", "sensibilizante_reach", "tpr", "tpr_danesa", "mutageno_rd", "mutageno_danesa", "cancer_mama", "cop")
  frasesR = array("R53", "R50-53", "R51-53", "R52-53", "R58")
  if substance.inLists(concern_lists) or anyElementInArray(frasesR, substance.identification.item("frasesR")) then

    sqlListaNegra="UPDATE dn_risc_sustancias SET negra=1 WHERE id=" & id_sustancia
    objConnection2.execute(sqlListaNegra),,adexecutenorecords

		razones = ""
    Dim carcinogenic_trade_union_lists : carcinogenic_trade_union_lists = Array( _
      "cancer_rd", _
      "cancer_danesa", _
      "cancer_iarc_excepto_grupo_3", _
      "cancer_otras", _
      "cancer_mama" _
    )
		if substance.inLists(carcinogenic_trade_union_lists) then
			razones = razones & ", carcinogenic"
		end if

		if substance.inLists("cop") then
			razones = razones & ", POP"
		end if

    if substance.inLists(MUTAGENIC_LISTS) then
			razones = razones & ", mutagenic"
		end if

		if substance.inList("de") then
			razones = razones & ", endocrine disrupter"
		end if

		if substance.inLists(NEUROTOXIC_LISTS) then
			razones = razones & ", neurotoxic"
		end if

    if substance.inLists(SENTITISER_LISTS) then
			razones = razones & ", sensitizer"
		end if

		if substance.inLists(TPR_LISTS) then
			razones = razones & ", toxic for reproduction"
		end if

    if stringContains(substance.classification.item("frasesR"), "R58") or stringContains(substance.classification.item("frasesR"),"R33") then
			razones = razones & ", bioaccumulative"
		end if

    if stringContains(substance.classification.item("frasesR"), "R58") then
			razones = razones & ", may cause long term adverse effects on the environment"
		end if

		if substance.inList("tpb") then
			razones = razones & ", toxic, persistent and bioaccumulative"
		end if

    cas_nums = array("87-68-3", "133-49-3", "75-74-1")
    if anyElementInArray(cas_nums, Array(substance.identification.item("cas_num"))) then
			razones = razones & ", very persistent and very bioaccumulative"
		end if

    r_phrases_aquatic_environment = Array("R53", "R50-53", "R51-53", "R52-53")
    frases_r_list = split(substance.classification.item("frasesR"), ", ")
    if anyElementInArray(r_phrases_aquatic_environment, frases_r_list) then
			razones = razones & ", may cause long term adverse effects in the aquatic environment"
		end if
		' Quitamos, si existe, el espacio y coma y despu�s convertimos el primer caracter en may�scula
		if (Len(razones)>0) then
			razones = Right(razones,Len(razones)-2)
			razones = UCase(Left(razones,1)) + Right(razones,Len(razones)-1)
		end if
%>
		<p id="concern_trade_union_list_title" class="subtitulo3">&nbsp;
			<img src="../imagenes/icono_atencion_20.png" align="absmiddle" />
			<a onclick="window.open('ver_definicion.asp?id=<%=dame_id_definicion("Lista negra")%>','def','width=300,height=200,scrollbars=yes,resizable=yes')" style="cursor:pointer"><img src="imagenes/ayuda.gif" width="14" height="14" align="absmiddle" border="0" /></a>&nbsp;Substance included in the List of Substances of concern for Trade Unions<% plegador "secc-concern_trade_union_list", "img-listanegra" %>
		</p>
		<p id="secc-concern_trade_union_list" class="texto" style="display:none">
			<span id="concern_trade_union_reasons.label">This substance is included in the List of Substances of concern for Trade Unions for the following reasons:</span><br/>
      <span id="concern_trade_union_reasons.value"><%=razones%>
		</p>

<%
	end if
end sub

' ###################################################################################

sub ap3_riesgos()
	sql = "select comentarios from dn_risc_sustancias_salud where id_sustancia=" & id_sustancia
	set objRstq = objConnection2.execute(sql)
  if substance.inLists(HEALTH_EFFECTS_LISTS) or _
    not is_empty(substance.health_effects.item("efecto_neurotoxico")) or _
    not is_empty(substance.health_effects.item("comentarios")) then
%>

		<!-- ################ Riesgos para la salud ###################### -->
    <br />
		<div id="ficha">
		<table width="100%" cellpadding=5>
			<tr>
				<td>
					<a name="identificacion"></a><img src="imagenes/risctox02.gif" alt="Health effects" />
				</td>
				<td align="right">
					<a href="#"><img src="../imagenes/subir.gif" border=0 alt=subir></a>
				</td>
			</tr>
		</table>

<%
		if substance.inLists(CARCINOGENIC_LISTS) then
      ap3_riesgos_tabla("Cancerígeno")
    end if

		if substance.inLists(MUTAGENIC_LISTS) then
      ap3_riesgos_tabla("Mutágeno")
    end if

		if substance.inList("de") then
      ap3_riesgos_tabla("Disruptor endocrino")
    end if

		if substance.inLists(NEUROTOXIC_LISTS) or _
      not is_empty(substance.health_effects.item("efecto_neurotoxico")) then
      ap3_riesgos_tabla("Neurotóxico")
    end if

    if substance.inLists(SENTITISER_LISTS) then
      ap3_riesgos_tabla("Sensibilizante")
    end if

    if substance.inLists(TPR_LISTS) then
      ap3_riesgos_tabla("Tóxico para la reproducción")
    end if

    if substance.inList("eepp") then
      ap3_riesgos_enfermedades()
    end if

  	if substance.inList("salud") then
      ap7_salud()
    end if


		if not is_empty(substance.health_effects.item("comentarios")) then
		  %>
			<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr>
					<td class="celdaabajo" colspan="2" align="center">
						<table cellpadding=0 cellspacing=0 width="100%" border="0">
							<tr>
								<td width="100%" class="titulo3" align="left">
                  More information on occupational health
    							<a href="javascript:toggle('secc-mas_informacion_salud_laboral', 'img-mas_informacion_salud_laboral');"><img src="../imagenes/desplegar.gif" align="absmiddle" id="img-mas_informacion_salud_laboral" alt="Click for more information" title="Click for more information" /></a>
	        			</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td id="secc-mas_informacion_salud_laboral" style="display:none">
						<ul>
							<li>
							<%= substance.health_effects.item("comentarios") %>
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
  <%
	end if ' salud
  Dim environment_effects_lists : environment_effects_lists = Array( _
    "tpb", _
    "directiva_aguas", _
    "alemana", _
    "sustancias_prioritarias", _
    "ozono", _
    "clima", _
    "aire", _
    "cop", _
    "suelos" _
  )
  if substance.inLists(environment_effects_lists) _
    or not is_empty(substance.health_effects.item("comentarios_medio_ambiente")) _
  then
    %>
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
  	if substance.inList("tpb") then
  		ap3_riesgos_tabla("Tóxica, Persistente y Bioacumulativa")
  	end if
  	if not is_empty(substance.environment_effects.item("mpmb")) then
  		ap3_riesgos_tabla("mPmB")
  	end if
  	if substance.inList("directiva_aguas") or substance.inList("alemana") then
      ap3_riesgos_tabla("Tóxica para el agua")
    end if
  	if substance.inList("suelos") then
      ap3_riesgos_tabla("Contaminante de suelos")
    end if
  	if substance.inList("ozono") or substance.inList("clima") or substance.inList("aire") then
      ap3_riesgos_tabla("Contaminante del aire")
    end if
  	if substance.inList("cop") then
      ap3_riesgos_tabla("Contaminante Orgánico Persistente (COP)")
    end if
  	if not is_empty(substance.environment_effects.Item("comentarios_medio_ambiente")) then
  		%>
  		<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
  			<tr>
  				<td class="celdaabajo" colspan="2" align="center">
  					<table cellpadding=0 cellspacing=0 width="100%" border="0">
  						<tr>
  							<td width="100%" class="titulo3" align="left">
      						Más información en medio ambiente
      						<a href="javascript:toggle('secc-mas_informacion_medio_ambiente', 'img-mas_informacion_medio_ambiente');"><img src="../imagenes/desplegar.gif" align="absmiddle" id="img-mas_informacion_medio_ambiente" alt="Click for more information" title="Click for more information" /></a>
          			</td>
  						</tr>
  					</table>
  				</td>
  			</tr>
  			<tr>
  				<td id="secc-mas_informacion_medio_ambiente" style="display:none">
  					<ul>
  						<li>
  						<%= substance.environment_effects.Item("comentarios_medio_ambiente") %>
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
    <%
	end if ' medio ambiente
end sub ' ap3_riesgos

sub ap3_riesgos_tabla(byval tipo)
  %>
	<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
   	<tr>
			<td class="celdaabajo" colspan="2" align="center">
				<table cellpadding=0 cellspacing=0 width="100%" border="0">
          <tr>
            <td width="100%" class="titulo3" align="left">
              <% ap3_riesgos_tabla_ayuda(tipo) %><%= traduceRiesgo(tipo) %>
              <% if ((tipo <> "COV") and (tipo <> "Vertidos") and (tipo <> "IPPC (PRTR Agua)") and (tipo <> "IPPC (PRTR Aire)") and (tipo <> "IPPC (PRTR Suelo)") and (tipo <> "Residuos Peligrosos") and (tipo <> "Accidentes Graves") and (tipo <> "Emisiones Atmosféricas") ) then %>
                <% plegador "secc-"&tipo, "img-"&tipo %>
              <% end if %>
            </td>
          </tr>
        </table>
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
    Response.Write "SEVESO (major-accidents)"

	case "Contaminante de suelos"
    Response.Write "According to <a href='http://www.istas.net/web/abreenlace.asp?idenlace=2940' target='_blank'>Spanish RD 9/2005</a>"

  case "Contaminante Orgánico Persistente (COP)":
    %>
    <fieldset>
      <legend class="subtitulo3"><strong>According to Stockholm Convention</strong></legend>
      <ul>
        <%
        if isNull(substance.environment_effects.item("cop")) then
          substance.environment_effects.item("cop") = ""
        end if

        array_anexos = split(substance.environment_effects.item("cop"), ";")
        for i=0 to ubound(array_anexos)
          Response.Write "<li>" & get_definition("COP Anexo " & trim(array_anexos(i)), "en") & "</li>"
        next

        if (trim(substance.environment_effects.item("enlace_cop")) <> "") then
			   response.write "<li><a href='" & substance.environment_effects.item("enlace_cop") & "' target='_blank'>Más información</a></li>"
        end if
	      %>
      </ul>
    </fieldset>

    <%
    case "Cancerígeno":
		  if (substance.inList("cancer_rd")) then
        %>
  			<fieldset>
  				<legend class="subtitulo3"><strong><span id="carcinogen_rd1272.label">According to R. 1272/2008</span></strong></legend>
  				<blockquote>
          <span id="carcinogen_rd1272.value">
          <%
  				nivel_cancerigeno_rd = dame_nivel_cancerigeno_rd()
  				nivel_cancerigeno_rd_txt = replace(nivel_cancerigeno_rd, "1", "1A")
  				nivel_cancerigeno_rd_txt = replace(nivel_cancerigeno_rd_txt, "2", "1B")
  				nivel_cancerigeno_rd_txt = replace(nivel_cancerigeno_rd_txt, "3", "2")

  				if (nivel_cancerigeno_rd <> "") then
				    response.write "<strong>Carcinogen level:</strong> " & nivel_cancerigeno_rd_txt
            Response.Write "&nbsp;<a onclick=window.open('ver_definicion.asp?id=" & dame_id_definicion("C") & nivel_cancerigeno_rd_txt & "','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>"
  				end if

          if not is_empty(substance.health_effects.Item("notas_cancer_rd")) then
            Response.Write "<br/><strong>Notas:</strong> " & substance.health_effects.Item("notas_cancer_rd")
  				end if
          %>
          </span>
  				</blockquote>
  			</fieldset>
        <%
			end if

				' Lista danesa ---------------------------------------------------------------
				if (substance.inList("cancer_danesa")) then
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
					if (trim(substance.health_effects.Item("notas_cancer_rd")) <> "") then
		%>
						<br/><strong>Notas:</strong> <%=substance.health_effects.Item("notas_cancer_rd")%>
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

        if (substance.inList("cancer_iarc")) then
		      %>
					<fieldset>
						<legend class="subtitulo3"><strong><span id="carcinogen_iarc">According to IARC </span><a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("IARC")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a></strong></legend>
	          <%
						if not is_empty(substance.health_effects.item("grupo_iarc")) or _
              not is_empty(substance.health_effects.Item("volumen_iarc")) or _
              not is_empty(substance.health_effects.Item("notas_iarc")) then
	          %>
							<blockquote>
							<table>
              <%
							if not is_empty(substance.health_effects.Item("grupo_iarc")) then
	            %>
								<tr>
                  <td class="subtitulo3">Group:</td>
                  <td>
                      <span id="carcinogen_iarc_group">
                        <%= substance.health_effects.Item("grupo_iarc")%>
                      </span>
                      &nbsp;<a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion(trim(substance.health_effects.Item("grupo_iarc")))%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>
                  </td>
                </tr>
              <%
							end if

							if not is_empty(substance.health_effects.Item("volumen_iarc")) then
              %>
								<tr>
                  <td class="subtitulo3">Volume:</td>
                  <td><span id="carcinogen_iarc_volume"><%= substance.health_effects.Item("volumen_iarc") %></span></td>
                </tr>
              <%
							end if

							if not is_empty(substance.health_effects.Item("notas_iarc")) then
	            %>
								<tr>
                  <td class="subtitulo3">Notes:</td>
                  <td><span id="carcinogen_iarc_notes"><%= substance.health_effects.Item("notas_iarc") %></span></td>
                </tr>
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

				' ntes
				if substance.inList("cancer_otras") then
		    %>
		    <fieldset>
				  <legend class="subtitulo3"><strong>According to other sources</strong></legend>
          <%
		      if is_empty(substance.health_effects.Item("categoria_cancer_otras")) then
		        substance.health_effects.Item("categoria_cancer_otras") = ""
		      end if

		      if is_empty(substance.health_effects.Item("fuente")) then
		        substance.health_effects.Item("fuente") = ""
		      end if

					array_categorias=split(substance.health_effects.Item("categoria_cancer_otras"), ",")
					array_fuentes=split(substance.health_effects.Item("fuente"), ",")

					' Damos por hecho que hay el mismo numero de categorias y fuentes y que coinciden en orden
          Dim fuente : fuente = ""
          Dim categoria : categoria = ""
					for i = 0 to ubound(array_fuentes)
            fuente = trim(array_fuentes(i))
            categoria = trim(array_categorias(i))
	        %>
					<fieldset>
						<legend class="subtitulo3">
              <strong>
                According to <%= fuente %> <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion(fuente)%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>
              </strong>
            </legend>
						<blockquote>
						<table>
							<tr>
                <td class="subtitulo3">
                  <span id="carcinogen_other_sources_category"><%= categoria %></span>:
                </td>
                <td>
                  <span id="carcinogen_other_sources_definition"><%= get_definition(categoria, "en") %>
                </td>
              </tr>
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

		    if (substance.inList("cancer_mama")) then

		      if (isNull(substance.health_effects.Item("cancer_mama_fuente"))) then

		        substance.health_effects.Item("cancer_mama_fuente") = ""

		      end if

		%>

					<fieldset>
						<legend class="subtitulo3"><strong>Según SSI (cáncer de mama) <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("SSI")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a></strong></legend>
						<blockquote>
						<table>
							<tr><td class="subtitulo3"><strong>Fuente:</strong><br /><a href="<%= substance.health_effects.Item("cancer_mama_fuente") %>" target="_blank"><%= replace(substance.health_effects.Item("cancer_mama_fuente"), "http://", "") %></a></td></tr>
						</table>
						</blockquote>
					</fieldset>

		<%

		    end if

		case "Mutágeno":
      ' MUTAGENO RD -------------------------------------------------------------
      if (substance.inList("mutageno_rd")) then
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
      if (substance.inList("mutageno_danesa")) then
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
			<% if not is_empty(substance.health_effects.item("nivel_disruptor")) then %>
				<tr>
					<td class="subtitulo3" valign="top">Source:</td>
					<td>
					<%
					endocrine_disrupter_levels = substance.health_effects.item("nivel_disruptor")
					for i = 0 to ubound(endocrine_disrupter_levels)
            set endocrine_disrupter_level = endocrine_disrupter_levels(i)
						response.write endocrine_disrupter_level.item("description") & "<br /><br />"
					next
					%>
					</td>
				</tr>
			<% end if %>
			</table>
			</blockquote>
<%
		case "Neurotóxico":
      if not is_empty(substance.health_effects.Item("efecto_neurotoxico")) or _
        not is_empty(substance.health_effects.Item("nivel_neurotoxico")) or _
        not is_empty(substance.health_effects.Item("fuente_neurotoxico")) then
      %>
			<blockquote>
			<table>
			<%	if not is_empty(substance.health_effects.Item("efecto_neurotoxico")) then %>
				<tr>
					<td class="subtitulo3" valign="top">Effect:</td>
					<td>
						<%
						neurotoxic_effects = substance.health_effects.Item("efecto_neurotoxico")
						for i = 0 to ubound(neurotoxic_effects)
              set neurotoxic_effect = neurotoxic_effects(i)
			        %>
			        <%= neurotoxic_effect.item("key") %> <a onclick=window.open('ver_definicion.asp?id=<%= neurotoxic_effect.item("id")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>
			        <%
						next
						%>
					</td>
				</tr>
        <%
			end if
			if not is_empty(substance.health_effects.Item("nivel_neurotoxico")) then
        dim neurotoxic_level
        set neurotoxic_level = substance.health_effects.Item("nivel_neurotoxico")(0)
        %>
				<tr>
					<td class="subtitulo3" valign="top">Level:</td><td><%= replace(neurotoxic_level.item("key"), "Level ", "") %>
					 <a onclick=window.open('ver_definicion.asp?id=<%= neurotoxic_level.item("id")%> ','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>
					</td>
        </tr>
        <%
      end if
        if not is_empty(substance.health_effects.Item("fuente_neurotoxico")) then
        %>
				<tr>
					<td class="subtitulo3" valign="top">Source:</td>
					<td>
					<%
          dim neurotoxic_sources
          neurotoxic_sources = substance.health_effects.item("fuente_neurotoxico")
					for i = 0 to ubound(neurotoxic_sources)
            response.write neurotoxic_sources(i).item("description")
        		if i < ubound(neurotoxic_sources) then
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
		      if substance.inList("sensibilizante") then
		        response.write "<li class='subtitulo3'>Sensitizer according to Regulation 1272/2008</li>"
		      end if

			  if substance.inList("sensibilizante_reach") then
		        response.write "<li class='subtitulo3'>REACH allergen &nbsp;<a href='http://www.istas.net/web/abreenlace.asp?idenlace=6340' target='_blank'>Ver documento</a></li>"
		      end if

		      if substance.inList("sensibilizante_danesa") then
		      %>
		        <li class='subtitulo3'>Sensitiser according to Danish EPA's<a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("EPA Danesa")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>  <% plegador_texto "frases_r_danesa_sensibilizante", "R phrases", "subtitulo3" %></li>
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
	      if (substance.inList("tpr")) then
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
	      if (substance.inList("tpr_danesa")) then
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

      if (substance.inList("prohibidas_embarazadas")) then
%>
  				<blockquote>
  					<strong>Fuente:</strong> Real Decreto 298/2009
				</blockquote>
<%
      end if

	case "Prohibida para trabajadoras lactantes":

      if (substance.inList("prohibidas_lactantes")) then
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
					<td><a href="<%= substance.health_effects.Item("enlace_tpb") %>"><%= corta(substance.health_effects.Item("anchor_tpb"), 70, "puntossuspensivos") %></a></td>
				</tr>
				<tr>
					<td class="subtitulo3" valign="top">Fuente/s:</td>
					<td class="subtitulo3"><%
						if trim(substance.health_effects.Item("fuentes_tpb")) <> "" then
							array_tpb = split(substance.health_effects.Item("fuentes_tpb"),",")
							for i=0 to ubound(array_tpb)
								response.write "<li>" & get_definition(trim(array_tpb(i)), "en")&"</li>"
							next
						end if
						if trim(substance.health_effects.Item("fuente_tpb")) <> "" then
							array_tpb = split(substance.health_effects.Item("fuente_tpb"),",")
							for i=0 to ubound(array_tpb)
								response.write "<li>" & get_definition(trim(array_tpb(i)), "en")&"</li>"
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
					<td class="subtitulo3"><%= get_definition("REACH", "en")%></td>

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
	                    <a href="#" onClick="window.open('dn_mas_informacion.asp?listado=restringidas&id=<%=id_sustancia%>','Informacion','width=500,height=230,scrollbars=yes,resizable=yes')">More information</a>
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
              <a href="#" onClick="window.open('dn_mas_informacion.asp?listado=prohibidas&id=<%=id_sustancia%>','Informacion','width=500,height=230,scrollbars=yes,resizable=yes')">More information</a>
            </td>
			</table>
			</blockquote>


<%
		case "Tóxica para el agua":
			response.write "<table>"
			if (substance.environment_effects.item("directiva_aguas") or substance.inList("directiva_aguas")) then
      %>
      <tr>
        <td class="subtitulo3" colspan="2">According to <a href="http://ec.europa.eu/environment/water/water-framework/index_en.html" target="_blank">Water Directive</a>, and subsequents <a href="http://www.istas.net/web/abreenlace.asp?idenlace=6323">amendments</a></td>
      </tr>
<%
			end if

			if (substance.inList("sustancias_prioritarias")) then
        %>
        <tr>
          <td class="subtitulo3" colspan="2">Possible priority substance according to <a href="http://ec.europa.eu/environment/water/water-framework/index_en.html" target="_blank">Water Directive</a>, and subsequents <a href="http://www.istas.net/web/abreenlace.asp?idenlace=6323">amendments</a></td>
        </tr>
        <%
			end if

			if not is_empty(substance.environment_effects.item("clasif_mma")) then
        %>
        <tr>
          <td class="subtitulo3" colspan="2">
            According to <a href="http://www.istas.net/risctox/abreenlace.asp?idenlace=2226" target="_blank">Ministry of Environment of Germany</a>
          </td>
        </tr>
        <tr>
          <td>&nbsp;&nbsp;&nbsp;</td>
          <td>
            <strong>Classification</strong>:
              <%=substance.environment_effects.item("clasif_mma")(0).item("key")%>
              <%=substance.environment_effects.item("clasif_mma")(0).item("description")%>
            <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion(parche_definicion(clasif_mma, "MMA"))%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>
          </td>
        </tr>
        <%
			end if
			if substance.environment_effects.item("sustancia_prioritaria") = 1 then
        %>
        <tr>
					<td class="subtitulo3">Possible priority substance </td><td></td>
				</tr>
        <%
			end if
			response.write "</table>"

			case "Contaminante del aire":
      %>
				<table>
        <%
				if (substance.environment_effects.item("dano_calidad_aire") or substance.inList("aire")) then
        %>
          <tr>
            <td class="subtitulo3">Air quality:</td>
            <td>Substance included in the <a href="http://eur-lex.europa.eu/LexUriServ/LexUriServ.do?uri=OJ:L:2008:152:0001:0044:EN:PDF" target="_blank">Directive 2008/50/EC</a> of 21 May 2008 on ambient air quality and cleaner air for Europe </td>
          </tr>
        <%
				end if
				if (substance.environment_effects.item("dano_ozono")) then
          %>
          <tr>
            <td class="subtitulo3">Ozone layer:</td>
            <td>A substance that deplete the ozone layer, according to <a href="abreenlace.asp?idenlace=2229" target="_blank">Regulation (EC) No 2037/2000</a> of the European Parliament and of the Council of 29 June 2000</td>
          </tr>
          <%
				end if
        %>
        <%
				if (substance.environment_effects.item("dano_cambio_clima")) then
        %>
        <tr>
          <td class="subtitulo3">Climate Change:</td>
          <td>Substance listed in the list of the <a href="abreenlace.asp?idenlace=2230" target="_blank">Kyoto Protocol</a></td>
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
            Source: <a href="https://echa.europa.eu/en/candidate-list-table" target="_blank">European Chemicals Agency (ECHA)</a>
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
            Source: <a href="https://echa.europa.eu/en/candidate-list-table" target="_blank">European Chemicals Agency (ECHA)</a>
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
            <a href="#" onClick="window.open('dn_mas_informacion.asp?listado=biocidas_prohibidas&id=<%=id_sustancia%>','Informacion','width=500,height=230,scrollbars=yes,resizable=yes')">More information</a>
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
            <a href="#" onClick="window.open('dn_mas_informacion.asp?listado=biocidas_autorizadas&id=<%=id_sustancia%>','Informacion','width=500,height=230,scrollbars=yes,resizable=yes')">More information</a>
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
              <a href="#" onClick="window.open('dn_mas_informacion.asp?listado=pesticidas_prohibidas&id=<%=id_sustancia%>','Informacion','width=500,height=230,scrollbars=yes,resizable=yes')">More information</a>
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
              <a href="#" onClick="window.open('dn_mas_informacion.asp?listado=pesticidas_autorizadas&id=<%=id_sustancia%>','Informacion','width=500,height=230,scrollbars=yes,resizable=yes')">More information</a>
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
            <a href="#" onClick="window.open('dn_mas_informacion.asp?listado=corap&id=<%=id_sustancia%>','Informacion','width=500,height=230,scrollbars=yes,resizable=yes')">More information</a>
          </td>
				</tr>
				<tr>
          <td class="subtitulo3">
						Source: <a href="http://echa.europa.eu/es/information-on-chemicals/evaluation/community-rolling-action-plan/corap-table" target="_blank">European Chemicals Agency (ECHA)</a>
					</td>
				</tr>
			</table>
			</blockquote>
<%
	end select
end sub

sub ap3_riesgos_enfermedades()

	' Se agrupan por listado, cada listado en una ficha blanca y dentro cada enfermedad
	sql_enf = "select distinct enf.id, enf.listado_ing, enf.nombre, enf.nombre_ing, enf.sintomas_ing, enf.actividades_ing FROM dn_risc_enfermedades AS enf LEFT OUTER JOIN dn_risc_grupos_por_enfermedades AS gpe ON enf.id = gpe.id_enfermedad LEFT OUTER JOIN dn_risc_sustancias_por_grupos AS spg ON gpe.id_grupo = spg.id_grupo LEFT OUTER JOIN dn_risc_sustancias_por_enfermedades AS spe ON spe.id_enfermedad = enf.id WHERE spg.id_sustancia="&id_sustancia&" OR spe.id_sustancia="&id_sustancia&" ORDER BY enf.listado_ing, enf.nombre_ing"
	'response.write "<br />"&sql_enf
	set objRstEnf=objConnection2.execute(sql_enf)
	if (not objRstEnf.eof) then
		listado_antiguo = ""
		do while (not objRstEnf.eof)
			' Para mostrar agrupados por listado, solo escribimos la cabecera si el listado es nuevo
			if (listado_antiguo <> objRstEnf("listado_ing")) then

				' Si el listado antiguo no es vacï¿½o, es que ya habiamos abierto antes uno asï¿½ que primero cerramos el anterior
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
				<table cellpadding=0 cellspacing=0 width="100%" border="0"><tr><td width="100%" class="titulo3" align="left"><a href="index.asp?idpagina=617"><img src="../imagenes/ayuda.gif" align="absmiddle" border="0" /></a> <%=objRstEnf("listado_ing")%>  <% plegador "secc-enf"&objRstEnf("listado_ing"), "img-enf"&objRstEnf("listado_ing") %></td></tr></table>
			</td>
		</tr>
		<tr id="secc-enf<%= aplana(objRstEnf("listado_ing")) %>" style="display:none">
			<td>
<%
				listado_antiguo = objRstEnf("listado_ing")
			end if
				if objRstEnf("nombre_ing")<>"" then
%>
				<fieldset style="padding:10px;">
				<!-- Tabla enfermedad -->
				<table cellspacing=1 cellpadding=1 border=0>
					<tr>
						<td class="subtitulo3" colspan=2><%=objRstEnf("nombre_ing")%></td>
					</tr>
				<%
					if (objRstEnf("sintomas_ing") <> "") then
				%>
					<tr>
						<td class="subtitulo3" align="right" valign="top" width='10%' nowrap style='padding-top:10px'>Symptoms:</td><td align="left" style'padding-top:10px'><%=replace(objRstEnf("sintomas_ing"), vbcrlf, "<br>")%></td>
					</tr>
				<%
					end if
				%>
				<%
					if (objRstEnf("actividades_ing") <> "") then
				%>
					<tr>
						<td class="subtitulo3" align="right" valign="top" width="10%" nowrap style='padding-top:10px'>Activities:</td><td align="left"  style='padding-top:10px'><%=replace(objRstEnf("actividades_ing"), vbcrlf, "<br>")%></td>
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

sub ap4_normativa_ambiental()
	if LANG = "en" and not substance.inList("cov") then
    exit sub
  end if
  if not (substance.inList("residuos") or substance.inList("vertidos") or substance.inList("lpcic")  or substance.inList("accidentes") or substance.inList("emisiones")) then
    exit sub
  end if
  %>

	<!-- ################ Normativa ambiental ###################### -->
	<br />
	<div id="ficha">
    <table width="100%" cellpadding=5>
      <tr>
        <td>
          <a name="identificacion"></a><img src="imagenes/risctox05.gif" alt="Environmental regulations" />
        </td>
        <td align="right">
          <a href="#"><img src="../imagenes/subir.gif" border=0 alt=subir></a>
        </td>
      </tr>
    </table>

  <%
  ' Para dividir los 7 posibles apartados en dos columnas, primero calculamos cuántos hay en total.
  total = 0

  if substance.inList("cov") then total = total +1 end if
  if substance.inList("vertidos") and LANG = "en" then total = total +1 end if
  if substance.inList("lpcic-agua") then total = total +1 end if
  if substance.inList("lpcic-aire") then total = total +1 end if
  if substance.inList("lpcic-suelo") then total = total +1 end if
  if substance.inList("residuos") and LANG = "en" then total = total +1 end if
  if substance.inList("accidentes") then total = total +1 end if
  if substance.inList("emisiones") then total = total +1 end if

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
  		if substance.inList("cov") then
  			ap3_riesgos_tabla("COV")
  			llevo = llevo +1
  			if llevo >= mitad then
  				response.write "</td><td valign='top' width='50%'>"
  				llevo = 0 ' Lo reseteo para que no vuelva a dividir
  			end if
  		end if

  		if substance.inList("vertidos") then
  			ap3_riesgos_tabla("Vertidos")
  			llevo = llevo +1
  			if llevo >= mitad then
  				response.write "</td><td valign='top' width='50%'>"
  				llevo = 0 ' Lo reseteo para que no vuelva a dividir
  			end if
  		end if

  		if substance.inList("lpcic-agua") then
  			ap3_riesgos_tabla("IPPC (PRTR Agua)")
  			llevo = llevo +1
  			if llevo >= mitad then
  				response.write "</td><td valign='top' width='50%'>"
  				llevo = 0 ' Lo reseteo para que no vuelva a dividir
  			end if
  		end if

  		if substance.inList("lpcic-aire") then
  			ap3_riesgos_tabla("IPPC (PRTR Aire)")
  			llevo = llevo +1
  			if llevo >= mitad then
  				response.write "</td><td valign='top' width='50%'>"
  				llevo = 0 ' Lo reseteo para que no vuelva a dividir
  			end if
  		end if

  		if substance.inList("lpcic-suelo") then
  			ap3_riesgos_tabla("IPPC (PRTR Suelo)")
  			llevo = llevo +1
  			if llevo >= mitad then
  				response.write "</td><td valign='top' width='50%'>"
  				llevo = 0 ' Lo reseteo para que no vuelva a dividir
  			end if
  		end if

  		if substance.inList("residuos") then
  			ap3_riesgos_tabla("Residuos Peligrosos")
  			llevo = llevo +1
  			if llevo >= mitad then
  				response.write "</td><td valign='top' width='50%'>"
  				llevo = 0 ' Lo reseteo para que no vuelva a dividir
  			end if
  		end if

  		if substance.inList("accidentes") then
  			ap3_riesgos_tabla("Accidentes Graves")
  			llevo = llevo +1
  			if llevo >= mitad then
  				response.write "</td><td valign='top' width='50%'>"
  				llevo = 0 ' Lo reseteo para que no vuelva a dividir
  			end if
  		end if

  		if substance.inList("emisiones") then
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
<%
end sub

sub ap4_normativa_restriccion_prohibicion()
	if not(substance.inList("prohibidas") or substance.inList("restringidas") or substance.inList("candidatas_reach") or substance.inList("autorizacion_reach") or substance.inList("biocidas_autorizadas") or substance.inList("biocidas_prohibidas") or substance.inList("pesticidas_autorizadas") or substance.inList("pesticidas_prohibidas") or substance.inList("prohibidas_embarazadas") or substance.inList("prohibidas_lactantes") or substance.inList("corap")) then
    exit sub
  end if
  %>
		<br />
		<div id="ficha">
      <table width="100%" cellpadding=5>
  			<tr>
  				<td>
  					<a name="identificacion"></a><img src="imagenes/risctox04-restricciones.gif" alt="Regulations on restriction / prohibition of substances" />
  				</td>
  				<td align="right">
  					<a href="#"><img src="../imagenes/subir.gif" border=0 alt=subir></a>
  				</td>
  			</tr>
  		</table>

		<table border="0" width="100%">
			<tr>
				<td valign="top">
          <span id="regulations">
<%
		if substance.inList("prohibidas") then
			ap3_riesgos_tabla("Sustancia prohibida")
		end if

		if substance.inList("restringidas") then
			ap3_riesgos_tabla("Sustancia restringida")
		end if

		if substance.inList("prohibidas_embarazadas") and LANG = "es" then ap3_riesgos_tabla("Prohibida para trabajadoras embarazadas") end if

		if substance.inList("prohibidas_lactantes") and LANG = "es" then ap3_riesgos_tabla("Prohibida para trabajadoras lactantes") end if

		if substance.inList("candidatas_reach") then
			ap3_riesgos_tabla("Sustancia candidata REACH")
		end if
		if substance.inList("autorizacion_reach") then
			ap3_riesgos_tabla("Sustancia REACH sujeta a autorización")
		end if
		if substance.inList("biocidas_autorizadas") then
			ap3_riesgos_tabla("Sustancia biocida autorizada")
		end if
		if substance.inList("biocidas_prohibidas") then
			ap3_riesgos_tabla("Sustancia biocida prohibida")
		end if
		if substance.inList("pesticidas_autorizadas") then
			ap3_riesgos_tabla("Sustancia pesticida autorizada")
		end if
		if substance.inList("pesticidas_prohibidas") then
			ap3_riesgos_tabla("Sustancia pesticida prohibida")
		end if
		if substance.inList("corap") then
			ap3_riesgos_tabla("Sustancia bajo evaluación. CoRAP")
		end if

        %>
      </span>
				</td>
			</tr>
		</table>
		</div>
<%
end sub

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
  if not substance.has_health_effects() then
    exit sub
  end if
  %>
	<!-- Efectos para la salud -->
	<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
   	<tr>
			<td class="celdaabajo" colspan="2" align="center">
				<table cellpadding=0 cellspacing=0 width="100%" border="0"><tr><td width="100%" class="titulo3" align="left">Other health adverse effects and affected organs <% plegador "secc-salud", "img-salud" %></td></tr></table>
			</td>
		</tr>
		<tr id="secc-salud" style="display:none">
			<td>
      <table border="0" width="100%">
        <tr>
        <%
        cardiocirculatorio = substance.health_effects.item("cardiocirculatorio")
        respiratorio = substance.health_effects.item("respiratorio")
        reproductivo = substance.health_effects.item("reproductivo")
        musculo_esqueletico = substance.health_effects.item("musculo_esqueletico")
        higado_gastrointestinal = substance.health_effects.item("higado_gastrointestinal")
        sistema_endocrino = substance.health_effects.item("sistema_endocrino")
        embrion = substance.health_effects.item("embrion")
        cancer = substance.health_effects.item("cancer")
        rinyon = substance.health_effects.item("rinyon")
        piel_sentidos = substance.health_effects.item("piel_sentidos")
        neuro_toxicos = substance.health_effects.item("neuro_toxicos")
        comentarios_sl = substance.health_effects.item("comentarios")

        if _
          cardiocirculatorio OR _
          respiratorio OR _
          reproductivo OR _
          musculo_esqueletico OR _
          sistema_inmunitario OR _
          higado_gastrointestinal OR _
          sistema_endocrino _
        then
          %>
          <td valign="top">
            <strong>- Affected systems:</strong><br/>
            <ul>
            <%
            if cardiocirculatorio then
              response.write "<li>Cardiovascular</li>"
            end if
            if respiratorio then
              response.write "<li>Respiratory</li>"
            end if
            if reproductivo then
              response.write "<li>Reproductive</li>"
            end if
            if musculo_esqueletico then
              response.write "<li>Musculoskeletal</li>"
            end if
            if sistema_inmunitario then
              response.write "<li>Immune</li>"
            end if
            if higado_gastrointestinal then
              response.write "<li>Gastrointestinal - liver</li>"
            end if
            if sistema_endocrino then
              response.write "<li>Endocrine</li>"
            end if
            %>
            </ul>
          </td>
          <%
        end if

        if embrion OR cancer OR rinyon OR piel_sentidos OR neuro_toxicos then
          %>
          <td valign="top">
            <strong>- Other effects:</strong><br />
            <ul>
              <%
              if embrion then
                response.write "<li>Damage to the embryo</li>"
              end if
              if cancer then
                response.write "<li>Cancer</li>"
              end if
              if rinyon then
                response.write "<li>Kidney damage</li>"
              end if
              if piel_sentidos then
                response.write "<li>Skin and mucous</li>"
              end if
              if neuro_toxicos then
                response.write "<li>Neurotoxic Effects</li>"
              end if
              %>
            </ul>
          </td>
          <%
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
end sub

' #############################################################################################
' Obtiene el nivel cancerígeno de los campos de clasificación
function dame_nivel_cancerigeno_rd()
	' Juntamos todas las clasificaciones
	clasificacion_rd = substance.classification.Item("clasificacion_1") & substance.classification.Item("clasificacion_2") & substance.classification.Item("clasificacion_3") & substance.classification.Item("clasificacion_4") & substance.classification.Item("clasificacion_5") & substance.classification.Item("clasificacion_6") & substance.classification.Item("clasificacion_7") & substance.classification.Item("clasificacion_8") & substance.classification.Item("clasificacion_9") & substance.classification.Item("clasificacion_10") & substance.classification.Item("clasificacion_11") & substance.classification.Item("clasificacion_12") & substance.classification.Item("clasificacion_13") & substance.classification.Item("clasificacion_14") & substance.classification.Item("clasificacion_15")

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
	posicion = instr(1,substance.classification.item("frases_r_danesa"), "Carc")

	' Sacamos el nivel como el caracter que hay justo detrás de la primera aparición de la subcadena
	if (posicion > 0) then
		dame_nivel_cancerigeno_danesa = mid(substance.classification.item("frases_r_danesa"), posicion+4, 1)
	else
		dame_nivel_cancerigeno_danesa = ""
	end if
end function

' #############################################################################################

function dame_nivel_mutageno_rd()
	' Juntamos todas las clasificaciones
	clasificacion_rd = substance.classification.item("clasificacion_1") & substance.classification.item("clasificacion_2") & substance.classification.item("clasificacion_3") & substance.classification.item("clasificacion_4") & substance.classification.item("clasificacion_5") & substance.classification.item("clasificacion_6") & substance.classification.item("clasificacion_7") & substance.classification.item("clasificacion_8") & substance.classification.item("clasificacion_9") & substance.classification.item("clasificacion_10") & substance.classification.item("clasificacion_11") & substance.classification.item("clasificacion_12") & substance.classification.item("clasificacion_13") & substance.classification.item("clasificacion_14") & substance.classification.item("clasificacion_15")

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
	posicion = instr(1,substance.classification.item("frases_r_danesa"), "Mut")

	' Sacamos el nivel como el caracter que hay justo detrás de la primera aparición de la subcadena
	if (posicion > 0) then
		dame_nivel_mutageno_danesa = mid(substance.classification.item("frases_r_danesa"), posicion+3, 1)
	else
		dame_nivel_mutageno_danesa = ""
	end if
end function

' #############################################################################################

function dame_nivel_reproduccion_rd()
	' Juntamos todas las clasificaciones
	clasificacion_rd = substance.classification.item("clasificacion_1") & substance.classification.item("clasificacion_2") & substance.classification.item("clasificacion_3") & substance.classification.item("clasificacion_4") & substance.classification.item("clasificacion_5") & substance.classification.item("clasificacion_6") & substance.classification.item("clasificacion_7") & substance.classification.item("clasificacion_8") & substance.classification.item("clasificacion_9") & substance.classification.item("clasificacion_10") & substance.classification.item("clasificacion_11") & substance.classification.item("clasificacion_12") & substance.classification.item("clasificacion_13") & substance.classification.item("clasificacion_14") & substance.classification.item("clasificacion_15")

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
  <a href="javascript:toggle('<%= id_bloque %>', '<%= id_imagen %>');"
     class="toggler"><img src="../imagenes/desplegar.gif" align="absmiddle" id="<%= id_imagen %>" alt="Click for more information" title="Click for more information" /></a>
<%
end sub

' #############################################################################################

sub plegador_texto(byval id_bloque, byval texto, byval clase)
  ' Pinta el HTML necesario para las llamadas a mostrar/ocultar el objeto
  ' Solo se emplea para el plegador de frases R danesas, en caso de que no se hayan mostrado ya.
  id_bloque=aplana(id_bloque)

  if (substance.hasFrasesRdanesa()) then
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

function traduceSimbolo(s)
	s = replace(s, "Peligro","Danger")
	s = replace(s, "Atención","Warning")
	traduceSimbolo = s
end function

function traduceEtiquetado(s)
	s = replace(s, "Expl. inest.","Unst. Expl")
	s = replace(s, "Expl.","Expl.")
	s = replace(s, "Gas infl.","Flam. Gas")
	s = replace(s, "Aerosol infl.","")
	s = replace(s, "Gas comb.","Ox. Gas")
	s = replace(s, "Gas a pres.","Press. Gas")
	s = replace(s, "Líq. infl.","Flam. Liq.")
	s = replace(s, "Sól. infl.","Flam. Sol.")
	s = replace(s, "Autorreact.","Self-react.")
	s = replace(s, "Líq. pir.","Pyr. Liq.")
	s = replace(s, "Sól. pir.","Pyr. Sol.")
	s = replace(s, "Calent. esp.","Self-heat.")
	s = replace(s, "Reac. agua","Water-react.")
	s = replace(s, "Líq. comb.","Ox. Liq.")
	s = replace(s, "Sól. comb.// Ojo","Ox. Sol.")
	s = replace(s, "Sól. comb.// ","Oxid. Sol.")
	s = replace(s, "Peróx. org.","Org. Perox.")
	s = replace(s, "Corr. met.","Met. Corr.")
	s = replace(s, "Tox. ag.","Acute Tox.")
	s = replace(s, "Corr. cut.","Skin Corr.")
	s = replace(s, "Irrit. cut.","Skin Irrit.")
	s = replace(s, "Les. oc.","Eye Dam.")
	s = replace(s, "Irrit. oc.","Eye Irrit.")
	s = replace(s, "Sens. resp.","Resp. Sens.")
	s = replace(s, "Sens. cut.","Skin Sens.")
	s = replace(s, "Muta.","Muta.")
	s = replace(s, "Carc.","Carc.")
	s = replace(s, "Repr.","Repr.")
	s = replace(s, "Lact.","Lact.")
	s = replace(s, "STOT única","STOT SE")
	s = replace(s, "STOT repe.","STOT RE")
	s = replace(s, "Tox. asp.","Asp. Tox.")
	s = replace(s, "Acuático agudo.","Aquatic Acute")
	s = replace(s, "Acuático crónico.","Aquatic Chronic")
	s = replace(s, "Ozono","Ozone")
	traduceEtiquetado = s
end function

function traduceRiesgo(riesgo)
	s = riesgo
	s = replace(s, "Cancerígeno","Carcinogen")
	s = replace(s, "Mutágeno","Mutagen")
	s = replace(s, "Disruptor endocrino","Endocrine disrupter")
	s = replace(s, "Neurotóxico","Neurotoxic")
	s = replace(s, "Sensibilizante","Sensitiser")
	s = replace(s, "Sensibilizante para REACH","REACH Sensitiser")
	s = replace(s, "Tóxico para la reproducción","Toxic for reproduction")
	s = replace(s, "mPmB","vPvB")
	s = replace(s, "Tóxica para el agua","Toxic for water")
	s = replace(s, "Contaminante de suelos","Soil pollutants")
	s = replace(s, "Contaminante del aire","Air pollutant")
	s = replace(s, "Contaminante Orgánico Persistente (COP)","Persistent Organic Pollutant (POP)")
	s = replace(s, "Residuos Peligrosos","Hazardous waste")
	s = replace(s, "Vertidos","Spill")
	s = replace(s, "Accidentes Graves","SEVESO (major-accidents)")
	s = replace(s, "COV","VOC")
	s = replace(s, "IPPC (PRTR Agua)","PRTR (water)")
	s = replace(s, "IPPC (PRTR Aire)","PRTR (air)")
	s = replace(s, "IPPC (PRTR Suelo)","PRTR (soil)")
	s = replace(s, "Emisiones Atmosféricas","Atmospheric emissions")
	s = replace(s, "Prohibida para trabajadoras embarazadas","Prohibited for pregnant workers")
	s = replace(s, "Prohibida para trabajadoras lactantes","Prohibited for nursing workers")
	s = replace(s, "Sustancia candidata REACH","REACH candidate list substance")
	s = replace(s, "Sustancia REACH sujeta a autorización","Substance under REACH authorisation")
	s = replace(s, "Sustancia biocida autorizada","Authorised biocide substance")
	s = replace(s, "Sustancia biocida prohibida","Banned biocide substance")
	s = replace(s, "Sustancia pesticida autorizada","Authorised pesticide substance")
	s = replace(s, "Sustancia pesticida prohibida","Banned pesticide substance")
	s = replace(s, "Sustancia restringida","Restricted substance")
	s = replace(s, "Sustancia prohibida","Banned substance")
	s = replace(s, "Tóxica, Persistente y Bioacumulativa","Persistent, Bioaccumulative and Toxic")
	s = replace(s, "Sustancia bajo evaluación. CoRAP","Substance under CoRAP evaluation")

	traduceRiesgo = s
end function
%>
