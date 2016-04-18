<div class="menu1">
<ul>
<li><a class="menu1two" href="dn_portada.asp">Portada</a></li>
<li><a class="menu1three" href="#">Risctox<!--[if IE 7]><!--></a><!--<![endif]-->
	<table><tr><td>
	<ul>
	<li><a href="dn_sustancias.asp">Sustancias</a></li>
	<li><a href="dn_grupos.asp">Grupos</a></li>	
	<li><a href="dn_companias.asp">Compa&ntilde;&iacute;as</a></li>
	<li><a href="dn_enfermedades.asp">Enfermedades</a></li>
	<li><a href="dn_access.asp">Exportador</a></li>
	<li><a href="dn_icsc.asp">ICSC</a></li>
	</ul>
	</td></tr></table>
<!--[if lte IE 6]></a><![endif]-->
</li>
<li><a class="menu1six" href="#">Alternativas<!--[if IE 7]><!--></a><!--<![endif]-->
	<table><tr><td>
	<ul>
	<li><a href="dn_ficheros.asp">Ficheros</a></li>
	<li><a href="dn_sectores.asp">Sectores</a></li>
	<li><a href="dn_procesos.asp">Procesos</a></li>
	<li><a href="dn_usos.asp">Usos</a></li>
  <li><a href="dn_enlaces.asp">Enlaces</a></li>
  <li><a href="dn_residuos.asp">Residuos</a></li>
	</ul>
	</td></tr></table>
<!--[if lte IE 6]></a><![endif]-->
</li>
<li><a class="menu1two" href="spl_listado_auditoria.asp">Auditor&iacute;a</a></li>


</ul>

</div>

<div align="right">
<strong>
<%
if (session("modo")="pruebas") then
%>
(EN PRUEBAS)
<%
elseif (session("modo")="produccion") then
%>
<font color="red"><blink>EN PRODUCCIÓN</blink></font>
<%
end if
%></strong>
</div>
