<%
'+++++++ XIP +++++
	'----- Si es restringida y no estás identificado no puedes entrar
	'if session("Id_Ecogente")="" then response.redirect "acceso.asp?idpagina="&idpagina
	id_ecogente = session("id_ecogente")
	'---- ATENCIÓN: ponerlo cuando publiquemos en abierto
	'id_ecogente = 1

	idpagina = 626	'--- página buscador, sólo para registrar estadísticas

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

	FUNCTION lista(x)
		response.write "<a href='index.asp?idpagina="&x&"'><img src='imagenes/ayuda.gif' width=14 height=14 border=0 align=absmiddle></a>&nbsp;"
	END FUNCTION

'+++fin xip +++++
%>
<!--#include file="adovbs.inc"-->
<!--#include file="dn_conexion.asp"-->
<!--#include file="dn_funciones_comunes.asp"-->
<!--#include file="dn_funciones_texto.asp"-->

<!--#include file="dn_restringida.asp"-->

<%
'si busc está vacio, mostramos formulario; si es 1, han dado a "buscar"; si es dos, han dado a paginación
busc = request.form("busc")
busc = EliminaInyeccionSQL(busc)

	filtro=0 'este filtro diferencia a este buscador del de sustancias: si esta a true, muestra solo las que son toxicas y tienen alternativas
%>
	<!--#include file="dn_buscador_sustancias.asp"-->

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

<link rel="stylesheet" type="text/css" href="estructura.css">
<script type="text/javascript">
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

		alert("Por favor, introduzca sus criterios de búsqueda");

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

<% if 1=0 then %><p class=campo>Est&aacute;s en: <a href=index.asp?idpagina=550>prevención riesgo químico</a> &gt; bbdd risctox</p><% end if %>
<table width="100%" border="0">
                <tr>
                <td></td>
                <td align='right'><input type="button" name="volver" class="boton" value="Volver a la portada de Risctox" onClick="window.location='./index.asp';"></td>
                </tr>
                </table>
<p class=titulo3>Base de datos de sustancias t&oacute;xicas y peligrosas RISCTOX </p>

<%

%>

<form action="dn_risctox_buscador2.asp?busc=1" method="post" name="myform"  onSubmit="primerapag();" >
 <input type="hidden" name='busc' value='<%=busc%>' />
 <input type="hidden" name='pag' value='<%=pag%>' />
 <input type="hidden" name='hr' value='<%=hr%>' />
 <input type="hidden" name='arr' value='<%=arr%>' />
 <input type="hidden" name='ordenacion' value='<%=ordenacion%>' />
 <input type="hidden" name='nregs' value='<%=nregs%>' />
<table class="tabla3" width="95%" align="center">
<tr><td colspan="3" class="subtitulo3">Buscador de sustancias</td></td></tr>
	<tr>
		<td align="right"><strong>Nombre</strong></td>
		<td><input type="text" name="nombre" value="<%=nombre%>" />
		<select name="tipobus">
		<option value="exacto" <%if tipobus="exacto" then response.write "selected"%>>nombre exacto</option>

		<option value="parte" <%if tipobus="parte" then response.write "selected"%>>parte del nombre</option>

		</select></td>
	</tr>
	<tr>
		<td align="right"><strong>Número CAS/CE/RD</strong></td>
		<td><input type="text" name="numero" value="<%=numero%>" /></td>
	</tr>
	<tr>
		<td colspan="2" align="center"><input type="submit" value="Buscar"/> <input type="reset" value="Borrar" /></td>
	</tr>
</table>

</form>
<br /><br />
<table width="95%" align="center" class="tabla3" border="0">

        <tr>
          <td colspan="2" align="center">
          <%
            ' Para centrar la lista negra dependerá del navegador
            if (navegador() = "FF") then
              w="45%"
            else
              w="45%"
            end if
          %>

            <table border="0" width="<%=w%>">
              <tr>
                <td class="subtitulo3" align="left">
                  <img src="imagenes/ico_lista_negra.gif" alt="Lista negra de ISTAS" align="absmiddle">&nbsp;<a onclick=window.open('ver_definicion.asp?id=195','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> <a href="dn_risctox_negra.asp" class="subtitulo3" onClick="alert('La lista negra de ISTAS puede tardar unos segundos en generarse debido a su gran tamaño.\n\nPor favor, espere.');">Lista negra de ISTAS</a>
                </td>
              <!--
              <tr>
                <td class="texto">
                  <li class=vineta_risctox><a onclick=window.open('ver_definicion.asp?id=195','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> <a href="dn_risctox_negra.asp">Lista negra de ISTAS</a></li>
                </td>
              </tr>
              -->
            </table>
          </td>
        </tr>

				<tr><td valign="top" width="45%">
					<table width="100%" align="center" border="0">
					<tr><td class="subtitulo3"><img src="imagenes/ico_danos_sl.gif" alt="Riesgos específicos para la salud" align="absmiddle">&nbsp;Riesgos específicos para la salud</td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(607)%>Cancerígenos y mutágenos:</li><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="dn_risctox_cym.asp" onClick="alert('Esta lista puede tardar unos segundos en generarse debido a su complejidad.\n\nPor favor, espere.');">Según R. 1272/2008</a><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="dn_risctox_cym2.asp">Según IARC</a><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="dn_risctox_cym3.asp">Según otras fuentes</a><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="dn_risctox_mama.asp">Según SSI (cáncer de mama)</a><br><br>
					</td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(609)%><a href="dn_risctox_tpr.asp">Tóxicos para la reproducción</a></li>

					</td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(610)%><a href="dn_risctox_dis.asp">Disruptores endocrinos</a></li></td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(611)%><a href="dn_risctox_neu.asp">Neurotóxicos</a><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					&nbsp;&nbsp;&nbsp;<% call lista(1190)%><a href="dn_risctox_oto.asp">Ototóxicos</a></li><br>


					</td></tr>

					<tr><td class="texto"><li class=vineta_risctox><% call lista(612)%><a href="dn_risctox_sen.asp">Sensibilizantes</a><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					&nbsp;&nbsp;&nbsp;<a href="dn_risctox_sen_reach.asp">Alérgenos REACH</a></li><br></td></tr>
					<tr><td class="texto">&nbsp;</td></tr>
					</table>
				</td><td valign="top" width="45%">
					<table width="100%" align="center">

					<tr><td class="subtitulo3"><img src="imagenes/ico_danos_ma.gif" alt="Riesgos específicos medioambiente" align="absmiddle">&nbsp;Riesgos específicos medioambiente</td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(613)%><a href="dn_risctox_pyb.asp">Tóxicas, persistentes y bioacumulativas</a></li></td></tr>
                    <tr><td class="texto"><li class=vineta_risctox><% call lista(613)%><a href="dn_risctox_mpmb.asp">mPmB</a></li></td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(614)%>Toxicidad acuática:</li><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="dn_risctox_tac.asp">Directiva de aguas</a><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="dn_risctox_tac2.asp">Peligrosas agua Alemania</a><br>
					</td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(615)%>Daño a la atmósfera:</li><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="dn_risctox_dat.asp">Capa de Ozono</a><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="dn_risctox_dat2.asp">Cambio climático</a><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="dn_risctox_dat3.asp">Calidad del aire</a>
					</td></tr>
					<tr><td class="texto"><li class=vineta_risctox>Contaminantes de suelos:</li><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<% call lista(1195)%><a href='dn_risctox_co_suelos.asp'>Según RD 9/2005</a><br>
					</td>
					</tr>

					<tr><td class="texto"><li class=vineta_risctox><% call lista(1185)%><a href="dn_risctox_cop.asp">Contaminantes&nbsp;Orgánicos&nbsp;Persistentes&nbsp;(COP's)</a></li></td></tr>

					<tr><td class="texto">&nbsp;</td></tr>
					</table>
				</td></tr>
				</table>
<br />		<br />
<table class="tabla3" width="95%" align="center">
				<tr>
				<td valign="top" width="45%">
					<table width="100%" align="center">
					<tr><td class="subtitulo3"><img src="imagenes/ico_normativa.gif" alt="Normativa sobre salud laboral" align="absmiddle">&nbsp;Normativa sobre salud laboral</td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(616)%>Límites de exposición profesional:</li>
					<table cellpadding=0 cellspacing=0 border=0><tr><td class="campo">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td class="campo"><a href="dn_risctox_vl1.asp">Valores Límite Ambientales</a></td></tr></table>
					<table cellpadding=0 cellspacing=0 border=0><tr><td class="campo">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td class="campo"><a href="dn_risctox_vl2.asp">Valores Límite Ambientales Cancerígenos</a></td></tr></table>
					<table cellpadding=0 cellspacing=0 border=0><tr><td class="campo">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td class="campo"><a href="dn_risctox_vl3.asp">Valores Límite Biológicos</a></td></tr></table>
					</td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(617)%><a href="dn_risctox_enf.asp">Enfermedades profesionales</a></li></td></tr>
					<tr><td class="texto">&nbsp;</td></tr>
					</table>
				</td>
				<td valign="top" width="45%">
					<table width="100%" align="center">
					<tr><td class="subtitulo3"><img src="imagenes/ico_normativa.gif" alt="Normativa ambiental" align="absmiddle">&nbsp;Normativa ambiental</td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(618)%><a href="dn_risctox_res.asp">Residuos peligrosos</a></li></td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(619)%><a href="dn_risctox_ver.asp">Vertidos</a></li></td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(620)%><a href="dn_risctox_emi.asp">Emisiones</a></li></td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(621)%><a href="dn_risctox_cov.asp">COV</a></li></td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(622)%>IPPC:</a></li>
                    <table cellpadding=0 cellspacing=0 border=0><tr><td class="campo">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td class="campo"><a href="dn_risctox_lpc.asp?filtro=ep1">PRTR (Agua)</a></td></tr></table>
					<table cellpadding=0 cellspacing=0 border=0><tr><td class="campo">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td class="campo"><a href="dn_risctox_lpc.asp?filtro=ep2">PRTR (Aire)</a></td></tr></table>
					<table cellpadding=0 cellspacing=0 border=0><tr><td class="campo">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td class="campo"><a href="dn_risctox_lpc.asp?filtro=ep3">PRTR (Suelo)</a></td></tr></table>
                    </td>
                    </tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(623)%><a href="dn_risctox_acm.asp">Accidentes graves</a></li></td></tr>
					<tr><td class="texto">&nbsp;</td></tr>
					</table>
				</td></tr>
				</table>

<br />		<br />
			<table class="tabla3" width="95%" align="center">
				<tr>
					<td valign="top" width="90%" colspan="2">
						<table width="100%" align="center">
						<tr><td class="subtitulo3"><img src="imagenes/ico_normativa.gif" alt="Normativa sobre restricción/prohibición de sustancias" align="absmiddle">&nbsp;Normativa sobre restricci&oacute;n / prohibici&oacute;n de sustancias</td></tr>
<%
'<!-- Tras conversación mantenida con Tatiana el 30/11/2010, se elimina esta lista ya que todas las sustancias están en restringidas -->
'<!--						<tr><td class="texto"><li class=vineta_risctox><a href="./dn_risctox_prohibidas.asp">Sustancias prohibidas</a></li></td></tr>-->
%>
						<tr><td class="texto"><li class=vineta_risctox><% call lista(1193)%><a href="./dn_risctox_restringidas.asp">Sustancias restringidas REACH</a></li></td></tr>
						<tr><td class="texto"><li class=vineta_risctox><% call lista(1194)%><a href="./spl_risctox_candidatas_reach.asp">Sustancias candidatas REACH</a></li></td></tr>
						<tr><td class="texto"><li class=vineta_risctox><% call lista(1194)%><a href="./spl_risctox_autorizacion_reach.asp">Sustancias sujetas a autorización de REACH</a></li></td></tr>
						<tr><td class="texto"><li class=vineta_risctox><% call lista(1188)%><a href="spl_risctox_prohibidas_embarazadas.asp">Prohibidas para trabajadoras embarazadas</a></li></td></tr>
						<tr><td class="texto"><li class=vineta_risctox><% call lista(1188)%><a href="spl_risctox_prohibidas_lactantes.asp">Prohibidas para trabajadoras lactantes</a></li></td></tr>
						<tr><td class="texto"><li class=vineta_risctox><% call lista(1192)%><a href="./spl_risctox_biocidas_prohibidas.asp">Sustancias Biocidas prohibidas</a></li></td></tr>
						<tr><td class="texto"><li class=vineta_risctox><% call lista(1192)%><a href="./spl_risctox_biocidas_autorizadas.asp">Sustancias Biocidas autorizadas</a></li></td></tr>
						<tr><td class="texto"><li class=vineta_risctox><% call lista(1191)%><a href="./spl_risctox_pesticidas_autorizadas.asp">Sustancias Pesticidas autorizadas</a></li></td></tr>
						<tr><td class="texto"><li class=vineta_risctox><% call lista(1191)%><a href="./spl_risctox_pesticidas_prohibidas.asp">Sustancias Pesticidas prohibidas</a></li></td></tr>
						<tr><td class="texto"><li class=vineta_risctox><% call lista(1194)%><a href="ist_risctox_corap.asp?f=corap">Sustancias bajo evaluación. CoRAP</a></li></td></tr>
						<tr><td class="texto">&nbsp;</td></tr>
						</table>
					</td>
				</tr>
			</table>

<!-- ############ FIN DE CONTENIDO ################## -->


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

</script>

</body>
</html>

<%
cerrarconexion
%>


