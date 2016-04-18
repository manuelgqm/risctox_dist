<!--#include file="adovbs.inc"--><!--#include file="dn_conexion.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd"><html><head>	<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">	<title>Istas</title>	<link rel="stylesheet" type="text/css" href="dn_estilos.css">	<link rel="stylesheet" type="text/css" href="dn_estilosmenu.css">	<script type="text/javascript" src="niftycube.js"></script>	<script type="text/javascript">	window.onload=function(){	Nifty("ul#split h3","top");	Nifty("ul#split div","bottom same-height");	}	</script></head><body><!--#include file="dn_menu.asp"-->	<h1>Exportador de Risctox</h1>
	<p>Escoge el listado de sustancias que quieres exportar en formato Access.</p>
	<p>Ten en cuenta que el proceso es largo y puede durar varias horas. Aproximadamente, unos 5 minutos cada 1.000 sustancias. No interrumpas el proceso de exportaci�n una vez comenzado: puedes usar otro navegador mientras tanto.</p>

	<table width="90%" align="center" class="tabla3">
		<tr><td valign="top" width="45%">
			<table width="100%" align="center">
			<tr><td class="subtitulo3"><h3>Riesgos espec�ficos para la salud</h3></td></tr>
			<tr>				<td class="texto">					<li class=vineta_risctox><a href="dn_access2.asp?listado=cancerigenos_all">Cancer�genos</a>:</li>					<div style="padding-left: 10px">
						<a href="dn_access2.asp?listado=cancerigenos_rd">Seg�n RD 363/1995</a><br>
						<a href="dn_access2.asp?listado=cym2">Seg�n IARC</a><br>
						<a href="dn_access2.asp?listado=cym3">Seg�n otras fuentes</a><br>
						<a href="dn_access2.asp?listado=mama">Seg�n SSI (c�ncer de mama)</a>					</div>
				</td>			</tr>			<tr>
				<td class="texto"><a href="dn_access2.asp?listado=mutagenos_all"><li class="vineta_risctox">Mut�genos</a></td>
			</tr>
			<tr><td class="texto"><li class=vineta_risctox><a href="dn_access2.asp?listado=tpr">T�xicos para la reproducci�n</a></li></td></tr>
			<tr><td class="texto"><li class=vineta_risctox><a href="dn_access2.asp?listado=dis">Disruptores endocrinos</a></li></td></tr>
			<tr><td class="texto"><li class=vineta_risctox><a href="dn_access2.asp?listado=neu">Neurot�xicos</a></li></td></tr>
			<tr><td class="texto"><li class=vineta_risctox><a href="dn_access2.asp?listado=sen">Sensibilizantes</a></li></td></tr>
			<tr><td class="texto"><li class=vineta_risctox><a href="dn_access2.asp?listado=negra">Lista negra de ISTAS</a></li></td></tr>
			<tr><td class="texto">&nbsp;</td></tr>
			</table>
		</td><td valign="top" width="45%">
			<table width="100%" align="center">
			<tr><td class="subtitulo3"><h3>Riesgos espec�ficos medioambiente</h3></td></tr>
			<tr><td class="texto"><li class=vineta_risctox><a href="dn_access2.asp?listado=pyb">T�xicas, persistentes y bioacumulativas</a></li></td></tr>
			<tr><td class="texto"><li class=vineta_risctox>Toxicidad acu�tica:</li><br>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="dn_access2.asp?listado=tac">Directiva de aguas</a><br>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="dn_access2.asp?listado=tac2">Peligrosas agua Alemania</a><br>
			</td></tr>
			<tr><td class="texto"><li class=vineta_risctox>Da�o a la atm�sfera:</li><br>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="dn_access2.asp?listado=dat">Capa de Ozono</a><br>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="dn_access2.asp?listado=dat2">Cambio clim�tico</a><br>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="dn_access2.asp?listado=dat3">Calidad del aire</a>
			</td></tr>
			<tr><td class="texto"><li class=vineta_risctox><a href="dn_access2.asp?listado=cop">Sustancias COP</a></li></td></tr>
			<tr><td class="texto">&nbsp;</td></tr>
			</table>
		</td></tr>
	</table>
<br /><br />			
	<table class="tabla3" width="90%" align="center">
		<tr>
		<td valign="top" width="45%">
			<table width="100%" align="center">
			<tr><td class="subtitulo3"><h3>Normativa sobre salud laboral</h3></td></tr>
			<tr><td class="texto"><li class=vineta_risctox>L�mites de exposici�n profesional:</li>
			<table cellpadding=0 cellspacing=0 border=0><tr><td class="campo">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td class="campo"><a href="dn_access2.asp?listado=vl1">Valores L�mite Ambientales</a></td></tr></table>
			<table cellpadding=0 cellspacing=0 border=0><tr><td class="campo">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td class="campo"><a href="dn_access2.asp?listado=vl2">Valores L�mite Ambientales Cancer�genos</a></td></tr></table>
			<table cellpadding=0 cellspacing=0 border=0><tr><td class="campo">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td class="campo"><a href="dn_access2.asp?listado=vl3">Valores L�mite Biol�gicos</a></td></tr></table>
			</td></tr>
			<tr><td class="texto"><li class=vineta_risctox><a href="dn_access2.asp?listado=enf">Enfermedades profesionales</a></li></td></tr>
			<tr><td class="texto">&nbsp;</td></tr>
			<tr><td class="texto">&nbsp;</td></tr>
			<tr><td class="texto">&nbsp;</td></tr>
			</table>
		</td>
		<td valign="top" width="45%">
			<table width="100%" align="center">
			<tr><td class="subtitulo3"><h3>Normativa ambiental</h3></td></tr>
			<tr><td class="texto"><li class=vineta_risctox><a href="dn_access2.asp?listado=res">Residuos peligrosos</a></li></td></tr>
			<tr><td class="texto"><li class=vineta_risctox><a href="dn_access2.asp?listado=ver">Vertidos</a></li></td></tr>
			<tr><td class="texto"><li class=vineta_risctox><a href="dn_access2.asp?listado=emi">Emisiones</a></li></td></tr>
			<tr><td class="texto"><li class=vineta_risctox><a href="dn_access2.asp?listado=cov">COV</a></li></td></tr>
			<tr><td class="texto"><li class=vineta_risctox><a href="dn_access2.asp?listado=lpc">LPCIC</a></li></td></tr>
			<tr><td class="texto"><li class=vineta_risctox><a href="dn_access2.asp?listado=acm">Accidentes graves</a></li></td></tr>
			<tr><td class="texto">&nbsp;</td></tr>
			</table>
		</td></tr>		<tr>
			<td valign="top" width="45%">				<h3>Normativa sobre restricci�n / prohibici�n de sustancias</h3>				<ul style="line-height: 18px">
					<li class="vineta_risctox" style="font-size: 8pt"><a href="dn_access2.asp?listado=rest">Sustancias restringidas REACH</a></li>
					<li class="vineta_risctox" style="font-size: 8pt"><a href="dn_access2.asp?listado=candidatas_reach">Sustancias candidatas REACH</a></li>
					<li class="vineta_risctox" style="font-size: 8pt"><a href="dn_access2.asp?listado=autorizacion_reach">Sustancias sujetas a autorizaci�n de REACH</a></li>
					<li class="vineta_risctox" style="font-size: 8pt"><a href="dn_access2.asp?listado=pro_emb">Prohibidas para trabajadoras embarazadas</a></li>
					<li class="vineta_risctox" style="font-size: 8pt"><a href="dn_access2.asp?listado=pro_lac">Prohibidas para trabajadoras lactantes</a></li>
					<li class="vineta_risctox" style="font-size: 8pt"><a href="dn_access2.asp?listado=biocidas_prohibidas">Sustancias Biocidas prohibidas</a></li>
					<li class="vineta_risctox" style="font-size: 8pt"><a href="dn_access2.asp?listado=biocidas_autorizadas">Sustancias Biocidas autorizadas</a></li>
					<li class="vineta_risctox" style="font-size: 8pt"><a href="dn_access2.asp?listado=pesticidas_autorizadas">Sustancias Pesticidas autorizadas</a></li>
					<li class="vineta_risctox" style="font-size: 8pt"><a href="dn_access2.asp?listado=pesticidas_prohibidas">Sustancias Pesticidas prohibidas</a></li>
				</ul>			</td>
			<td valign="top" width="45%">			</td>
		</tr>
	</table>

	<p>Tambi�n puedes <a href="dn_access2.asp?listado=todas">exportar todas las sustancias</a>, o <a href="./estructuras/risctox.mdb">volver a descargar el �ltimo access</a> sin necesidad de volver a generarlo.</p>

</body></html>

<% cerrarconexion %>