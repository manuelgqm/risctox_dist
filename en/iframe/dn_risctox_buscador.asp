<%
'+++++++ XIP +++++
	'----- Si es restringida y no est�s identificado no puedes entrar
	'if session("Id_Ecogente")="" then response.redirect "acceso.asp?idpagina="&idpagina
	id_ecogente = session("id_ecogente")
	'---- ATENCI�N: ponerlo cuando publiquemos en abierto
	'id_ecogente = 1

	idpagina = 626	'--- p�gina buscador, s�lo para registrar estad�sticas

FUNCTION vistaprevia(texto)
		texto = replace(texto,chr(13),"<br>")
		texto = replace(texto,"'","&#39;")
		texto = replace(texto,"<v1>","<img src=../imagenes/vineta.gif>&nbsp;")
		texto = replace(texto,"<v2>","&nbsp;&nbsp;<img src=../imagenes/vineta.gif>&nbsp;")
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

	FUNCTION lista(x)
		response.write "<a href='index.asp?idpagina="&x&"'><img src='imagenes/ayuda_grey.gif' width=14 height=14 border=0 align=absmiddle></a>&nbsp;"
	END FUNCTION

'+++fin xip +++++
%>

<!--#include file="../../adovbs.inc"-->
<!--#include file="../../dn_conexion.asp"-->
<!--#include file="../../dn_funciones_comunes.asp"-->
<!--#include file="../../dn_funciones_texto.asp"-->

<!--#include file="../../dn_restringida.asp"-->

<%
'si busc est� vacio, mostramos formulario; si es 1, han dado a "buscar"; si es dos, han dado a paginaci�n
busc = request.form("busc")
busc = EliminaInyeccionSQL(busc)

	filtro=0 'este filtro diferencia a este buscador del de sustancias: si esta a true, muestra solo las que son toxicas y tienen alternativas
%>
	<!--#include file="../dn_buscador_sustancias.asp"-->

<%
	sql = "SELECT numeracion FROM WEBISTAS_PAGINAS WHERE idpagina="&idpagina
	Set objR = Server.CreateObject ("ADODB.Recordset")
	set objR = OBJConnection.Execute(sql)
	numeracion = objR("numeracion")

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

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>RISCTOX: Toxic and hazardous substances database</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="Title" content="Risctox" />
<meta name="Author" content="SPL Sistemas de Información - www.spl-ssi.com" />
<meta name="description" content="RISCTOX: Toxic and hazardous substances database" />
<meta name="Subject" content="RISCTOX: Toxic and hazardous substances database" />
<meta name="Keywords" content="RISCTOX: Toxic and hazardous substances database" />
<meta name="Language" content="English" />
<meta name="Revisit" content="15 days" />
<meta name="Distribution" content="Global" />
<meta name="Robots" content="All" />

<link rel="stylesheet" type="text/css" href="css/estructura.css">
<link rel="stylesheet" type="text/css" href="../../jqueryUI/css/smoothness/jquery-ui-1.10.3.custom.min.css">

<script type="text/javascript" src="http://istas.net/_framework/jquery/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="../..//jqueryUI/js/jquery-ui-1.10.3.custom.min.js"></script>

<script type="text/javascript">
	$( document ).ready( function() {
		$( "#dialog1" ).dialog({
			resizable: false,
			height: 200,
			modal: true,
			autoOpen: false,
			buttons: {
				"Go": function() { $( this ).dialog( "close" ); window.location.replace( 'dn_risctox_lista.asp?f=cym' ) },
				Cancel: function() { $( this ).dialog( "close" )	}
			}
		});

		$( "#dialog2" ).dialog({
			resizable: false,
			height: 220,
			modal: true,
			autoOpen: false,
			buttons: {
				"Go": function() { $( this ).dialog( "close" ); window.location.replace( 'dn_risctox_lista.asp?f=negra' ) },
				Cancel: function() { $( this ).dialog( "close" )	}
			}
		});

		$( "#regulation" ).click( function(ev) {
			ev.preventDefault();
			$( "#dialog1" ).dialog( "open" );
		} );
		$( "#substancesTU" ).click( function( ev ) {
			ev.preventDefault();
			$( "#dialog2" ).dialog( "open" );
		} );
	} );
function cambiapag(paginadest)
{
	var frm = document.forms["myform"];
	frm.busc.value=2;
	frm.pag.value=paginadest;
	frm.submit();
}

function primerapag()
{
	var frm = document.forms["myform"];



	if ((frm.nombre.value == "") && (frm.numero.value == "")&& (frm.cas_alternativo.value == ""))

	{

		alert("Please enter your search criteria");

	}

	else

	{

		frm.busc.value=1;
		frm.pag.value=1;
		frm.submit();

	}
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
                <td><a href="http://www.etuc.org/a/6023" target="_blank" style="font-size: 0.9em;">Trade Union priority list for REACH authorization</a></td>
                <td align='right'><input type="button" name="volver" class="boton" value="back to homepage" onClick="window.location='./index.asp';"></td>
                </tr>
                </table>
<p class=titulo3>RISCTOX: Toxic and hazardous substances database </p>


<form action="dn_risctox_buscador2.asp?busc=1" method="post" name="myform"  onSubmit="primerapag();" >
 <input type="hidden" name='busc' value='<%=busc%>' />
 <input type="hidden" name='pag' value='<%=pag%>' />
 <input type="hidden" name='hr' value='<%=hr%>' />
 <input type="hidden" name='arr' value='<%=arr%>' />
 <input type="hidden" name='ordenacion' value='<%=ordenacion%>' />
 <input type="hidden" name='nregs' value='<%=nregs%>' />
<table class="tabla3" width="95%" align="center">
<tr><td colspan="3" class="subtitulo3">Substance search</td></td></tr>
	<tr>
		<td align="right"><strong>Name</strong></td>
		<td><input type="text" name="nombre" value="<%=nombre%>" />
		<select name="tipobus">
		<option value="exacto" <%if tipobus="exacto" then response.write "selected"%>>exact name</option>

		<option value="parte" <%if tipobus="parte" then response.write "selected"%>>part of the name</option>

		</select></td>
	</tr>
	<tr>
		<td align="right"><strong>CAS/EC/Index No</strong></td>
		<td><input type="text" name="numero" value="<%=numero%>" /></td>
	</tr>
	<tr>
		<td colspan="2" align="center"><input type="submit" value="Search"/> <input type="reset" value="Erase" /></td>
	</tr>
</table>

</form>
<br /><br />
<table width="95%" align="center" class="tabla3" border="0">

        <tr>
          <td colspan="2" align="center">
          <%
            ' Para centrar la lista negra depender� del navegador
            if (navegador() = "FF") then
              w="55%"
            else
              w="55%"
            end if
          %>

            <table border="0" width="<%=w%>">
              <tr>
                <td class="subtitulo3" align="left">
                  <img src="imagenes/ico_lista_negra.gif" alt="ISTAS's black list" align="absmiddle">&nbsp;<a onclick=window.open('ver_definicion.asp?id=195','def','width=300,height=200,scrollbars=yes,resizable=yes') ><img src='imagenes/ayuda_grey.gif' width=14 height=14 border='0' style='cursor:pointer' /></a> <a href="dn_risctox_lista.asp?f=negra" class="subtitulo3" id="substancesTU">Substances of concern for Trade Unions</a>
                </td>
            </table>
          </td>
        </tr>

				<tr><td valign="top" width="45%">
					<table width="100%" align="center" border="0">
					<tr><td class="subtitulo3"><img src="imagenes/ico_danos_sl.gif" alt="Health effects" align="absmiddle">&nbsp;Health effects</td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(607)%>Carcinogens and mutagens:</li><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="dn_risctox_lista.asp?f=cym" id="regulation">According to Regulation 1272/2008</a><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="dn_risctox_lista.asp?f=cym2">According to IARC</a><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="dn_risctox_lista.asp?f=cym3">According to other sources</a><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="dn_risctox_lista.asp?f=mama">According to SSI (breast cancer)</a><br><br>
					</td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(609)%><a href="dn_risctox_lista.asp?f=tpr">Reproductive toxicants</a></li>

					</td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(610)%><a href="dn_risctox_lista.asp?f=dis">Endocrine disrupters</a></li></td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(611)%><a href="dn_risctox_lista.asp?f=neu">Neurotoxicants</a><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					&nbsp;&nbsp;&nbsp;<% call lista(1190)%><a href="dn_risctox_lista.asp?f=oto">Ototoxicants</a></li><br>


					</td></tr>

					<tr><td class="texto"><li class=vineta_risctox><% call lista(612)%><a href="dn_risctox_lista.asp?f=sen">Sensitisers</a><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					&nbsp;&nbsp;&nbsp;<a href="dn_risctox_lista.asp?f=senreach">REACH allergens</a></li><br></td></tr>
					<tr><td class="texto">&nbsp;</td></tr>
					</table>
				</td><td valign="top" width="45%">
					<table width="100%" align="center">

					<tr><td class="subtitulo3"><img src="imagenes/ico_danos_ma.gif" alt="Environmental effects" align="absmiddle">&nbsp;Environmental effects</td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(613)%><a href="dn_risctox_lista.asp?f=pyb">Persistent, Bioaccumulative and Toxics</a></li></td></tr>
                    <tr><td class="texto"><li class=vineta_risctox><% call lista(613)%><a href="dn_risctox_lista.asp?f=mpmb">vPvB</a></li></td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(614)%>Aquatic toxicity:</li><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="dn_risctox_lista.asp?f=tac">Water Frame Directive</a><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="dn_risctox_lista.asp?f=tac2">German water pollutants</a><br>
					</td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(615)%>Atmospheric pollutants:</li><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="dn_risctox_lista.asp?f=dat">Ozone-depleting substances</a><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="dn_risctox_lista.asp?f=dat2">Greenhouse gases</a><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="dn_risctox_lista.asp?f=dat3">Air pollutants</a>
					</td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(1195)%><a href='dn_risctox_lista.asp?f=cos'>Soil pollutants</a></li><br>

					</td>
					</tr>

					<tr><td class="texto"><li class=vineta_risctox><% call lista(1185)%><a href="dn_risctox_lista.asp?f=cop">Persistent Organic Pollutants (POPs)</a></li></td></tr>

					<tr><td class="texto">&nbsp;</td></tr>
					</table>
				</td></tr>
				</table>
<br />		<br />
<table class="tabla3" width="95%" align="center">
				<tr>
				<td valign="top" width="45%">
					<table width="100%" align="center">
					<tr><td class="subtitulo3"><img src="imagenes/ico_normativa.gif" alt="Occupational Health and Safety Regulations" align="absmiddle">&nbsp;Occupational Health and Safety Regulations</td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(617)%><a href="dn_risctox_lista.asp?f=enf">Occupational diseases</a></li></td></tr>
					<tr><td class="texto">&nbsp;</td></tr>
					</table>
				</td>
				<td valign="top" width="45%">
					<table width="100%" align="center">
					<tr><td class="subtitulo3"><img src="imagenes/ico_normativa.gif" alt="Environmental regulations" align="absmiddle">&nbsp;Environmental regulations</td></tr>
					<tr><td class="texto"><li class="vineta_risctox"><% call lista(621)%><a href="dn_risctox_lista.asp?f=cov">VOCs</a></li></td></tr>
					<tr><td class="texto">&nbsp;</td></tr>
					</table>
				</td></tr>
				</table>
<br><br>
			<table class="tabla3" width="95%" align="center">
				<tr>
					<td valign="top" width="90%" colspan="2">
						<table width="100%" align="center">
						<tr><td class="subtitulo3"><img src="imagenes/ico_normativa.gif" alt="Regulations on restriction / prohibition of substances" align="absmiddle">&nbsp;Regulations on restriction / prohibition of substances</td></tr>
						<tr><td class="texto"><li class=vineta_risctox><a href="dn_risctox_lista.asp?f=pro">Banned substances</a></li></td></tr>
						<tr><td class="texto"><li class=vineta_risctox><% call lista(1193)%><a href="dn_risctox_lista.asp?f=rest">Restricted substances under REACH</a></li></td></tr>
						<tr><td class="texto"><li class=vineta_risctox><% call lista(1194)%><a href="dn_risctox_lista.asp?f=candidatas_reach">REACH Candidate list</a></li></td></tr>
						<tr><td class="texto"><li class=vineta_risctox><% call lista(1194)%><a href="dn_risctox_lista.asp?f=autorizacion_reach">REACH Authorisation list</a></li></td></tr>
						<tr><td class="texto"><li class=vineta_risctox><% call lista(1192)%><a href="dn_risctox_lista.asp?f=biocidas_prohibidas">Banned biocides</a></li></td></tr>
						<tr><td class="texto"><li class=vineta_risctox><% call lista(1192)%><a href="dn_risctox_lista.asp?f=biocidas_autorizadas">Authorised biocides</a></li></td></tr>
						<tr><td class="texto"><li class=vineta_risctox><% call lista(1191)%><a href="dn_risctox_lista.asp?f=pesticidas_prohibidas">Banned pesticides</a></li></td></tr>
						<tr><td class="texto"><li class=vineta_risctox><% call lista(1191)%><a href="dn_risctox_lista.asp?f=pesticidas_autorizadas">Authorised pesticides</a></li></td></tr>
						<tr><td class="texto">&nbsp;</td></tr>
						</table>
					</td>
				</tr>
			</table>

				<div style="text-align: center; margin-bottom: 20px; margin-top: 30px; width: 340px; margin-left: auto; margin-right: auto;">
					<div style="text-align: right; float: left; margin-right: 5px;">
						This database has been developed by<br />
						<a href="http://www.istas.ccoo.es/" target="_blank" style="font-weight: bold">ISTAS</a> -
						<a href="http://www.ccoo.es/" target="_blank" style="font-weight: bold">CC.OO.</a><br />
						in cooperation with <a href="http://www.etui.org/" target="_blank" style="font-weight: bold">ETUI</a>
					</div>
					<div style="float: right; border:">
						<a href="http://www.istas.ccoo.es/" target="_blank">
							<img src="imagenes/logo_istas_w80.jpg"  />
						</a>
					</div>
					<div style="clear: both;"></div>
				</div>

		  </div>
			</div>


			<img src="imagenes/pie_risctox.gif" width="708" border="0">


    		</div>
    	</div>
	<p class="texto" style="padding-left: 5px; padding-right: 5px;color:#999; font-size: 9pt;">
		This web has been developed by <a href="http://www.spl-ssi.com" style="color:#999;" target="_blank">SPL Sistemas de Informaci&oacute;n</a>
	</p>

</div>

<script type="text/javascript">
var gaJsHost = (("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));
</script>
<script type="text/javascript">
var pageTracker = _gat._getTracker("UA-6837368-1");
pageTracker._trackPageview();
</script>

</body>
	<div id="dialog1" title="confirmation">
		<p style="text-align: left;">
			<span class="ui-icon ui-icon-alert" style="float:left; margin:0 7px 20px 0;"></span>
			This list can take time to arise due to its complexity.
		</p>
	</div>
	<div id="dialog2" title="confirmation">
		<p style="text-align: left;">
			<span class="ui-icon ui-icon-alert" style="float:left; margin:0 7px 20px 0;"></span>
			The Substances of concern for Trade Unions list may take several seconds to be generated due to its large size.
		</p>
	</div>
</html>

<% cerrarconexion %>