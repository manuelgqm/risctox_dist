<!--#include file="../dn_restringida.asp"-->
<!--#include file="../dn_conexion.asp"-->
<!--#include file="../dn_funciones_comunes.asp"-->
<!--#include file="../dn_funciones_texto.asp"-->

<%
on error resume next

' Borde para ver las tablas u ocultarlas
'borde=" border='1'"
borde=""

' Inicialmente no hay errores...
errores = ""

' Cogemos el id de la sustancia elegida y traemos sus datos
id_sustancia = EliminaInyeccionSQL(request("id_sustancia"))

'SERGIO
response.redirect("http://www.istas.net/risctox/dn_risctox_ficha_sustancia.asp?id_sustancia="&id_sustancia)
response.End()


sql="SELECT * FROM dn_risc_sustancias FULL OUTER JOIN dn_risc_sustancias_vl ON dn_risc_sustancias.id = dn_risc_sustancias_vl.id_sustancia FULL OUTER JOIN dn_risc_sustancias_iarc ON dn_risc_sustancias.id = dn_risc_sustancias_iarc.id_sustancia FULL OUTER JOIN dn_risc_sustancias_cancer_otras ON dn_risc_sustancias.id = dn_risc_sustancias_cancer_otras.id_sustancia FULL OUTER JOIN dn_risc_sustancias_neuro_disruptor ON dn_risc_sustancias.id = dn_risc_sustancias_neuro_disruptor.id_sustancia FULL OUTER JOIN dn_risc_sustancias_ambiente ON dn_risc_sustancias.id = dn_risc_sustancias_ambiente.id_sustancia FULL OUTER JOIN dn_risc_sustancias_mama_cop ON dn_risc_sustancias.id = dn_risc_sustancias_mama_cop.id_sustancia WHERE dn_risc_sustancias.id="&id_sustancia
set objRst=objConnection2.execute(sql)
if(objRst.eof) then
	errores="No se ha encontrado la sustancia indicada"
else
	' dn_risc_sustancias
	nombre = objRst("nombre")
	nombre_ing = elimina_repes(objRst("nombre_ing"), "@")
	num_rd = objRst("num_rd")
	num_ce_einecs = objRst("num_ce_einecs")
	num_ce_elincs = objRst("num_ce_elincs")
	num_cas = objRst("num_cas")
	num_onu = objRst("num_onu")
	num_icsc = objRst("num_icsc")
	formula_molecular = objRst("formula_molecular")
	estructura_molecular = objRst("estructura_molecular")
	simbolos = objRst("simbolos")
	clasificacion_1 = trim(objRst("clasificacion_1"))
	clasificacion_2 = trim(objRst("clasificacion_2"))
	clasificacion_3 = trim(objRst("clasificacion_3"))
	clasificacion_4 = trim(objRst("clasificacion_4"))
	clasificacion_5 = trim(objRst("clasificacion_5"))
	clasificacion_6 = trim(objRst("clasificacion_6"))
	clasificacion_7 = trim(objRst("clasificacion_7"))
	clasificacion_8 = trim(objRst("clasificacion_8"))
	clasificacion_9 = trim(objRst("clasificacion_9"))
	clasificacion_10 = trim(objRst("clasificacion_10"))
	clasificacion_11 = trim(objRst("clasificacion_11"))
	clasificacion_12 = trim(objRst("clasificacion_12"))
	clasificacion_13 = trim(objRst("clasificacion_13"))
	clasificacion_14 = trim(objRst("clasificacion_14"))
	clasificacion_15 = trim(objRst("clasificacion_15"))
	frases_s = trim(objRst("frases_s"))
	conc_1 = objRst("conc_1")
	eti_conc_1 = objRst("eti_conc_1")
	conc_2 = objRst("conc_2")
	eti_conc_2 = objRst("eti_conc_2")
	conc_3 = objRst("conc_3")
	eti_conc_3 = objRst("eti_conc_3")
	conc_4 = objRst("conc_4")
	eti_conc_4 = objRst("eti_conc_4")
	conc_5 = objRst("conc_5")
	eti_conc_5 = objRst("eti_conc_5")
	conc_6 = objRst("conc_6")
	eti_conc_6 = objRst("eti_conc_6")
	conc_7 = objRst("conc_7")
	eti_conc_7 = objRst("eti_conc_7")
	conc_8 = objRst("conc_8")
	eti_conc_8 = objRst("eti_conc_8")
	conc_9 = objRst("conc_9")
	eti_conc_9 = objRst("eti_conc_9")
	conc_10 = objRst("conc_10")
	eti_conc_10 = objRst("eti_conc_10")
	conc_11 = objRst("conc_11")
	eti_conc_11 = objRst("eti_conc_11")
	conc_12 = objRst("conc_12")
	eti_conc_12 = objRst("eti_conc_12")
	conc_13 = objRst("conc_13")
	eti_conc_13 = objRst("eti_conc_13")
	conc_14 = objRst("conc_14")
	eti_conc_14 = objRst("eti_conc_14")
	conc_15 = objRst("conc_15")
	eti_conc_15 = objRst("eti_conc_15")
	notas_rd_363 = objRst("notas_rd_363")
	notas_xml = replace(objRst("notas_xml"), "@", "@ ")
	frases_r_danesa = trim(objRst("frases_r_danesa"))

	' dn_risc_sustancias_vl
	estado_1 = objRst("estado_1")
	vla_ed_ppm_1 = objRst("vla_ed_ppm_1")
	vla_ed_mg_m3_1 = objRst("vla_ed_mg_m3_1")
	vla_ec_ppm_1 = objRst("vla_ec_ppm_1")
	vla_ec_mg_m3_1 = objRst("vla_ec_mg_m3_1")
	notas_vla_1 = objRst("notas_vla_1")
	' Parche: quitar "VLB" en notas VLA
	if (not isnull(notas_vla_1)) then
		notas_vla_1 = replace(notas_vla_1, "VLB", "")
	end if

	estado_2 = objRst("estado_2")
	vla_ed_ppm_2 = objRst("vla_ed_ppm_2")
	vla_ed_mg_m3_2 = objRst("vla_ed_mg_m3_2")
	vla_ec_ppm_2 = objRst("vla_ec_ppm_2")
	vla_ec_mg_m3_2 = objRst("vla_ec_mg_m3_2")
	notas_vla_2 = objRst("notas_vla_2")
	' Parche: quitar "VLB" en notas VLA
	if (not isnull(notas_vla_2)) then
		notas_vla_2 = replace(notas_vla_2, "VLB", "")
	end if

	estado_3 = objRst("estado_3")
	vla_ed_ppm_3 = objRst("vla_ed_ppm_3")
	vla_ed_mg_m3_3 = objRst("vla_ed_mg_m3_3")
	vla_ec_ppm_3 = objRst("vla_ec_ppm_3")
	vla_ec_mg_m3_3 = objRst("vla_ec_mg_m3_3")
	notas_vla_3 = objRst("notas_vla_3")
	' Parche: quitar "VLB" en notas VLA
	if (not isnull(notas_vla_3)) then
		notas_vla_3 = replace(notas_vla_3, "VLB", "")
	end if

	estado_4 = objRst("estado_4")
	vla_ed_ppm_4 = objRst("vla_ed_ppm_4")
	vla_ed_mg_m3_4 = objRst("vla_ed_mg_m3_4")
	vla_ec_ppm_4 = objRst("vla_ec_ppm_4")
	vla_ec_mg_m3_4 = objRst("vla_ec_mg_m3_4")
	notas_vla_4 = objRst("notas_vla_4")
	' Parche: quitar "VLB" en notas VLA
	if (not isnull(notas_vla_4)) then
		notas_vla_4 = replace(notas_vla_4, "VLB", "")
	end if

	estado_5 = objRst("estado_5")
	vla_ed_ppm_5 = objRst("vla_ed_ppm_5")
	vla_ed_mg_m3_5 = objRst("vla_ed_mg_m3_5")
	vla_ec_ppm_5 = objRst("vla_ec_ppm_5")
	vla_ec_mg_m3_5 = objRst("vla_ec_mg_m3_5")
	notas_vla_5 = objRst("notas_vla_5")
	' Parche: quitar "VLB" en notas VLA
	if (not isnull(notas_vla_5)) then
		notas_vla_5 = replace(notas_vla_5, "VLB", "")
	end if

	estado_6 = objRst("estado_6")
	vla_ed_ppm_6 = objRst("vla_ed_ppm_6")
	vla_ed_mg_m3_6 = objRst("vla_ed_mg_m3_6")
	vla_ec_ppm_6 = objRst("vla_ec_ppm_6")
	vla_ec_mg_m3_6 = objRst("vla_ec_mg_m3_6")
	notas_vla_6 = objRst("notas_vla_6")
	' Parche: quitar "VLB" en notas VLA
	if (not isnull(notas_vla_6)) then
		notas_vla_6 = replace(notas_vla_6, "VLB", "")
	end if

	ib_1 = objRst("ib_1")
	vlb_1 = objRst("vlb_1")
	momento_1 = objRst("momento_1")
	notas_vlb_1 = objRst("notas_vlb_1")

	ib_2 = objRst("ib_2")
	vlb_2 = objRst("vlb_2")
	momento_2 = objRst("momento_2")
	notas_vlb_2 = objRst("notas_vlb_2")

	ib_3 = objRst("ib_3")
	vlb_3 = objRst("vlb_3")
	momento_3 = objRst("momento_3")
	notas_vlb_3 = objRst("notas_vlb_3")

	ib_4 = objRst("ib_4")
	vlb_4 = objRst("vlb_4")
	momento_4 = objRst("momento_4")
	notas_vlb_4 = objRst("notas_vlb_4")

	ib_5 = objRst("ib_5")
	vlb_5 = objRst("vlb_5")
	momento_5 = objRst("momento_5")
	notas_vlb_5 = objRst("notas_vlb_5")

	ib_6 = objRst("ib_6")
	vlb_6 = objRst("vlb_6")
	momento_6 = objRst("momento_6")
	notas_vlb_6 = objRst("notas_vlb_6")

	' Cancer
	notas_cancer_rd = objRst("notas_cancer_rd")
	' Parche: quitar las que diga "v�ase Tabla 3"
	notas_cancer_rd = replace(notas_cancer_rd, "v�ase Tabla 3", "")

	grupo_iarc = objRst("grupo_iarc")
	volumen_iarc = objRst("volumen_iarc")
	notas_iarc = objRst("notas_iarc")
	categoria_cancer_otras = objRst("categoria_cancer_otras")
	fuente = objRst("fuente")

	' Disruptor endocrino
	nivel_disruptor = objRst("nivel_disruptor")

	' Neurot�xico
	efecto_neurotoxico=objRst("efecto_neurotoxico")
	nivel_neurotoxico=objRst("nivel_neurotoxico")
	fuente_neurotoxico=objRst("fuente_neurotoxico")

	' TPB
	enlace_tpb = objRst("enlace_tpb")
	anchor_tpb = objRst("anchor_tpb")

	' T�xica para el agua
	directiva_aguas = objRst("directiva_aguas")
	clasif_mma = objRst("clasif_mma")

	' Contaminante del aire
	dano_calidad_aire = objRst("dano_calidad_aire")
	dano_ozono = objRst("dano_ozono")
	dano_cambio_clima = objRst("dano_cambio_clima")

	' Cancer Mama
	cancer_mama = objRst("cancer_mama")
	cancer_mama_fuente = objRst("cancer_mama_fuente")

  ' COP
  cop = objRst("cop")
end if
objRst.close()
set objRst=nothing

' Sinonimos
sinonimos = dameSinonimos(id_sustancia)

' Comprobamos si est� en cada lista, para no tener que buscar varias veces
esta_en_lista_cancer_rd = esta_en_lista ("cancer_rd", id_sustancia)
esta_en_lista_cancer_danesa = esta_en_lista ("cancer_danesa", id_sustancia)
esta_en_lista_mutageno_rd = esta_en_lista ("mutageno_rd", id_sustancia)
esta_en_lista_mutageno_danesa = esta_en_lista ("mutageno_danesa", id_sustancia)
esta_en_lista_cancer_iarc = esta_en_lista ("cancer_iarc", id_sustancia)
esta_en_lista_cancer_iarc_excepto_grupo_3 = esta_en_lista ("cancer_iarc_excepto_grupo_3", id_sustancia)
esta_en_lista_cancer_otras = esta_en_lista ("cancer_otras", id_sustancia)
esta_en_lista_cancer_mama = esta_en_lista ("cancer_mama", id_sustancia)
esta_en_lista_tpr = esta_en_lista ("tpr", id_sustancia)
esta_en_lista_tpr_danesa = esta_en_lista ("tpr_danesa", id_sustancia)
esta_en_lista_de = esta_en_lista ("de", id_sustancia)
esta_en_lista_neurotoxico_rd = esta_en_lista ("neurotoxico_rd", id_sustancia)
esta_en_lista_neurotoxico_danesa = esta_en_lista ("neurotoxico_danesa", id_sustancia)
esta_en_lista_neurotoxico_nivel = esta_en_lista ("neurotoxico_nivel", id_sustancia)
esta_en_lista_neurotoxico = esta_en_lista_neurotoxico_rd OR esta_en_lista_neurotoxico_danesa OR esta_en_lista_neurotoxico_nivel
esta_en_lista_sensibilizante = esta_en_lista ("sensibilizante", id_sustancia)
esta_en_lista_sensibilizante_danesa = esta_en_lista ("sensibilizante_danesa", id_sustancia)
esta_en_lista_eepp = esta_en_lista ("eepp", id_sustancia)
esta_en_lista_tpb = esta_en_lista ("tpb", id_sustancia)
esta_en_lista_directiva_aguas = esta_en_lista ("directiva_aguas", id_sustancia)
esta_en_lista_alemana = esta_en_lista ("alemana", id_sustancia)
esta_en_lista_aire = esta_en_lista ("aire", id_sustancia)
esta_en_lista_ozono = esta_en_lista ("ozono", id_sustancia)
esta_en_lista_clima = esta_en_lista ("clima", id_sustancia)
esta_en_lista_aire = esta_en_lista ("aire", id_sustancia)
esta_en_lista_cov = esta_en_lista ("cov", id_sustancia)
esta_en_lista_vertidos = esta_en_lista ("vertidos", id_sustancia)
esta_en_lista_lpcic = esta_en_lista ("lpcic", id_sustancia)
esta_en_lista_lpcic_agua = esta_en_lista ("lpcic-agua", id_sustancia)
esta_en_lista_lpcic_aire = esta_en_lista ("lpcic-aire", id_sustancia)
esta_en_lista_residuos = esta_en_lista ("residuos", id_sustancia)
esta_en_lista_accidentes = esta_en_lista ("accidentes", id_sustancia)
esta_en_lista_emisiones = esta_en_lista ("emisiones", id_sustancia)
esta_en_lista_salud = esta_en_lista ("salud", id_sustancia)
esta_en_lista_cop = esta_en_lista ("cop", id_sustancia)
'response.write "esta en lista cop? "&esta_en_lista_cop

' Condiciones para mostrar las frases R danesas en Clasificacion
' Se mostrar�n si existen las frases R danesas y NO existen las de RD

' Montamos frases R
frases_r=trim(monta_frases_r(clasificacion_1, clasificacion_2, clasificacion_3, clasificacion_4, clasificacion_5, clasificacion_6, clasificacion_7, clasificacion_8, clasificacion_9, clasificacion_10, clasificacion_11, clasificacion_12, clasificacion_13, clasificacion_14, clasificacion_15))

'if ((not esta_en_lista_cancer_rd) and (not esta_en_lista_sensibilizante_danesa) or (frases_r = "")) then
if (frases_r = "") and (frases_r_danesa <> "") then
  frases_r_danesa_mostradas=true
else
  frases_r_danesa_mostradas=false
end if
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ECOinformas: Plataforma prevenci�n de riesgo qu�mico</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="Title" content="ECOinformas" />
<meta name="Author" content="Dabne" />
<meta name="description" content="Informaci�n, formaci�n y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Subject" content="Informaci�n, formaci�n y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Keywords" content="Informaci�n, formaci�n y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Language" content="Spanish" />
<meta name="Revisit" content="15 days" />
<meta name="Distribution" content="Global" />
<meta name="Robots" content="All" />

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
		<!--#include file="../dn_cabecera.asp"-->
		<div id="texto">
			
<div class="texto">
<!-- ################ CONTENIDO ###################### -->

<table width="100%" border="0">
<tr>
<td><p class=campo>Est&aacute;s en: <a href="index.asp?idpagina=550">prevenci�n riesgo qu�mico</a> &gt; <a href="dn_risctox_buscador.asp">bbdd risctox</a> &gt; ficha de sustancia</p></td>
<td align="right"><input type="button" name="volver" class="boton2" value="Nueva b�squeda" onclick="window.location='dn_risctox_buscador.asp';"></td>
</tr>
</table>

<p class=titulo3>RISCTOX: Ficha de la sustancia</p>

<div id="ficha">
	<!-- ################ Identificacion de la sustancia ###################### -->
	<table width="100%" cellpadding=5>
		<tr>
			<td>
				<a name="identificacion"></a><img src="imagenes/risctox01.gif" alt="identificaci�n de la sustancia" width="255" height="32" />
			</td>
			<td align="right">
				<a href="#"><img src="imagenes/subir.gif" border=0 alt=subir></a>
			</td>
		</tr>
	</table>

	<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
		<!-- ################ Identificaci�n ###################### -->
   	<tr>
			<td class="celdaabajo" colspan="2" align="center">
				<table cellpadding=0 cellspacing=0 width="100%" border="0"><tr><td width="100%" class="titulo3" align="left">SUSTANCIA</td></tr></table>
			</td>
		</tr>
		<!-- 1.- Datos de sustancia -->
		<% ap1_identificacion() %>
	</table>

	<div style="height:3pt"></div>

		<!-- 2.- Clasificaci�n -->
		<% ap2_clasificacion() %>

	<br />
</div>
<!-- fin div ficha -->

<!-- 3.- Riesgos -->
<% ap3_riesgos() %>

<!-- 4.- Normativa ambiental -->
<% ap4_normativa_ambiental() %>

<!-- 5.- Alternativas relacionadas -->
<% ap5_alternativas() %>

<!-- 6.- Sectores en los que se utiliza -->
<% ap6_sectores() %>

<!-- ############ FIN DE CONTENIDO ################## -->
<br />
<center>
<input type="button" name="imprimir" class="boton2" value="Imprimir ficha" onclick="window.print();"> 
<input type="button" name="enviar" class="boton2" value="Enviar ficha de sustancia" onclick="onclick=window.open('dn_recomendar.asp?id=<%=id_sustancia%>','recomendar','width=500,height=230,scrollbars=yes,resizable=yes')"> 
<input type="button" name="volver" class="boton2" value="Nueva b�squeda" onclick="window.location='dn_risctox_buscador.asp';">
</center>

<br>
<br>
Esta p�gina ha sido desarrollada por <a href="http://www.istas.ccoo.es/" target="_blank"><b>ISTAS</b></a> que es una Fundaci�n de <a href="http://www.ccoo.es/" target="_blank"><font color="#FF0000"><b>CC.OO.</b></font></a><br>

				
				</div>
				<p>&nbsp;</p>
			</div>
			
			
			<map name="Map1" id="Map1">
            		<area shape="rect" coords="307,38,399,104" href="http://www.fundacion-biodiversidad.es" target="_blank" alt="Fundaci�n Biodiversidad" />
            		<area shape="rect" coords="400,38,546,104" href="http://www.istas.ccoo.es" target="_blank" alt="Instituto Sindical de Trabajo, Ambiente y Salud" />
      			<area shape="rect" coords="547,38,701,104" href="http://www.mtas.es/UAFSE/default.htm" target="_blank" alt="Fondo Social Europeo" />
      			</map>

			<map name="Map2" id="Map2">
            		<area shape="rect" coords="300,8,392,66" href="http://www.fundacion-biodiversidad.es" target="_blank" alt="Fundaci�n Biodiversidad" />
            		<area shape="rect" coords="393,8,539,66" href="http://www.istas.ccoo.es" target="_blank" alt="Instituto Sindical de Trabajo, Ambiente y Salud" />
      			<area shape="rect" coords="540,8,694,66" href="http://www.mtas.es/UAFSE/default.htm" target="_blank" alt="Fondo Social Europeo" />
      			</map>
      			<map name="Map3" id="Map3">
            		<area shape="rect" coords="300,18,392,80" href="http://www.fundacion-biodiversidad.es" target="_blank" alt="Fundaci�n Biodiversidad" />
            		<area shape="rect" coords="393,18,539,80" href="http://www.istas.ccoo.es" target="_blank" alt="Instituto Sindical de Trabajo, Ambiente y Salud" />
      			<area shape="rect" coords="540,18,694,80" href="http://www.mtas.es/UAFSE/default.htm" target="_blank" alt="Fondo Social Europeo" />

      			</map>
			<img src="imagenes/pie3.jpg" width="708" border="0" usemap="#Map3">
      			
    			
    		</div>
    	</div>
	<div id="sombra_abajo"></div>
</div>
</body>
</html>

<%
cerrarconexion
%>




<%
' ##########################################################################
function dameGrupos(byval id_sustancia)
	' Devuelve lista de grupos para la sustancia indicada

	lista = ""

	sql="SELECT dn_risc_grupos.id AS id_grupo, nombre, descripcion FROM dn_risc_sustancias_por_grupos INNER JOIN dn_risc_grupos ON dn_risc_sustancias_por_grupos.id_grupo = dn_risc_grupos.id WHERE id_sustancia="&id_sustancia&" ORDER BY nombre"
	set objRst=objConnection2.execute(sql)
	if (not objRst.eof) then
		do while (not objRst.eof)
      id_grupo = objRst("id_grupo")
      nombre = objRst("nombre")
      descripcion = objRst("descripcion")
      if (descripcion <> "") then
        ' Montamos enlace para abrir ventana emergente de descripci�n
        enlace_descripcion = " <a onclick=window.open('dn_glosario.asp?tabla=grupos&id="&id_grupo&"','def','width=500,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a>"
      else
        ' No hay descripci�n
        enlace_descripcion = ""
      end if

			if (lista = "") then
				lista = objRst("nombre")&enlace_descripcion
			else
				lista = lista&", "&objRst("nombre")&enlace_descripcion
			end if

			objRst.movenext
		loop
	end if
	objRst.close()
	set objRst=nothing

	dameGrupos = lista
end function

' ##########################################################################
function dameUsos(byval id_sustancia)
	' Devuelve lista de usos para la sustancia indicada

	lista = ""

  sql="SELECT DISTINCT u.id AS id_uso, u.nombre AS nombre_uso, u.descripcion AS descripcion_uso FROM dn_risc_usos AS u LEFT OUTER JOIN dn_risc_grupos_por_usos AS gpu ON u.id = gpu.id_uso LEFT OUTER JOIN dn_risc_sustancias_por_grupos AS spg ON gpu.id_grupo = spg.id_grupo LEFT OUTER JOIN dn_risc_sustancias_por_usos AS spu ON spu.id_uso = u.id WHERE spg.id_sustancia="&id_sustancia&" OR spu.id_sustancia="&id_sustancia&" ORDER BY u.nombre"
  'response.write sql

	set objRst=objConnection2.execute(sql)

	if (not objRst.eof) then

		do while (not objRst.eof)

      id_uso = objRst("id_uso")
      nombre_uso = objRst("nombre_uso")
      descripcion = objRst("descripcion_uso")

      if (descripcion <> "") then
        ' Montamos enlace para abrir ventana emergente de descripci�n
        enlace_descripcion = " <a onclick=window.open('dn_glosario.asp?tabla=usos&id="&id_uso&"','def','width=500,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>"&nombre_uso&"</a>"
      else
        ' No hay descripci�n
        enlace_descripcion = nombre_uso
      end if

			if (lista = "") then
				lista = enlace_descripcion
			else
				lista = lista&", "&enlace_descripcion
			end if

			objRst.movenext
		loop
	end if
	objRst.close()
	set objRst=nothing

	dameUsos = lista
end function

' ##########################################################################
function dameCompanias(byval id_sustancia)
	' Devuelve lista de compa��as para la sustancia indicada

	lista = ""

	sql="SELECT dn_risc_companias.id as idcomp, nombre FROM dn_risc_sustancias_por_companias INNER JOIN dn_risc_companias ON dn_risc_sustancias_por_companias.id_compania = dn_risc_companias.id WHERE id_sustancia="&id_sustancia&" ORDER BY nombre"
	set objRst=objConnection2.execute(sql)
	if (not objRst.eof) then
		do while (not objRst.eof)
			if (lista = "") then
				lista = "<a onclick=window.open('dn_risctox_ficha_compania.asp?id="&objRst("idcomp")&"','comp','width=600,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>"&objRst("nombre")&"</a>"
			else
				lista = lista&", <a onclick=window.open('dn_risctox_ficha_compania.asp?id="&objRst("idcomp")&"','comp','width=600,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>"&objRst("nombre")&"</a>"
			end if

			objRst.movenext
		loop
	end if
	objRst.close()
	set objRst=nothing

	dameCompanias = lista
end function

' ##########################################################################
'' Deprecated. Se usa la definida en dn_funciones_comunes.asp

''function dameNombreComercial (byval id_sustancia)
''	nombre_comercial = ""
''	sql = "SELECT nombre FROM dn_risc_nombres_comerciales WHERE ''id_sustancia="&id_sustancia
''	set objRst=objConnection2.execute(sql)
''	if (not objRst.eof) then
''		nombre_comercial = objRst("nombre")
''	end if
''	objRst.close()
''	set objRst=nothing
''	dameNombreComercial = nombre_comercial
''end function

' ##########################################################################

sub ap1_identificacion()
%>
	<tr>
		<td class="subtitulo3" align="right" valign="top">
			<a onclick=window.open('ver_definicion.asp?id=82','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a> Nombre:
		</td>
		<td class="texto" valign="middle">
			<b><%=espaciar(nombre)%></b>
		</td>
	</tr>

	<%
	if (sinonimos<>"") then
	%>
		<tr>
			<td class="subtitulo3" align="right" valign="top">
				<a onclick=window.open('ver_definicion.asp?id=83','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a> Sin�nimos:
			</td>
			<td class="texto" valign="middle">
				<%=sinonimos%>
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

	<% if (num_cas <> "") or (num_ce_einecs <> "") or (num_ce_elincs <> "") then %>
		<tr>
			<td class="subtitulo3" align="right" valign="top">
				N�meros de Identificaci�n:
			</td>
			<td class="texto" valign="middle">
				<% if (num_cas <> "") then response.write "<a onclick=window.open('ver_definicion.asp?id=84','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a> <b>CAS</b>: "&num_cas&"<br/>" %>
				<%
					if (num_ce_einecs <> "") then
						response.write "<a onclick=window.open('ver_definicion.asp?id=85','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a> <b>CE EINECS</b>: "&num_ce_einecs&"<br/>"
					elseif (num_ce_elincs <> "") then
						response.write "<a onclick=window.open('ver_definicion.asp?id=85','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a> <b>CE ELINCS</b>: "&num_ce_elincs&"<br/>"
					end if
				%>
			</td>
		</tr>
	<% end if ' hay numeros? %>

	<%
		grupos = dameGrupos(id_sustancia)
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
		usos = dameUsos(id_sustancia)
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
		if (num_icsc <> "") then
	%>
		<tr>
			<td class="subtitulo3" align="right" valign="top">
				 Ficha Internacional de Seguridad Qu�mica (<a onclick="window.open('ver_definicion.asp?id=<%=dame_id_definicion("INSHT")%>', 'def', 'width=300,height=200,scrollbars=yes,resizable=yes')" class="subtitulo3">INSHT</a>)
			</td>
			<td class="texto" valign="middle">
          <% 
            array_icsc=split(num_icsc, "@")
            for i=0 to ubound(array_icsc)
          %>
              <a href="http://www.mtas.es/insht/ipcsnspn/nspn<%= array_icsc(i) %>.htm" target="_blank"><%= array_icsc(i) %></a> 
          <%
            next
          %>
			</td>
		</tr>
	<% end if %>


	<%
		companias = dameCompanias(id_sustancia)
		'if (companias <> "") then
	%>
    <!--
		<tr>
			<td class="subtitulo3" align="right" valign="top" width="50%">
				Compa��as productoras/distribuidoras:
			</td>
			<td class="texto" valign="middle">
				<%=companias%>
			</td>
		</tr>
    -->
	<% 'end if ' hay companias? %>

	<% if (nombre_ing <> "") or (num_rd <> "") or (formula_molecular <> "") or (estructura_molecular <> "") or (notas_xml <> "") or (companias <> "") then %>
		<tr>
			<td class="subtitulo3" align="right" valign="top" width="35%">
				M�s informaci�n <% plegador "secc-masinformacion", "img-masinformacion" %>
			</td>
			<td class="texto" valign="middle" id="secc-masinformacion" style="display:none">

        <% if (nombre_ing <> "") then 
            array_nombres_ingleses = split(nombre_ing, "@")
            if (ubound(array_nombres_ingleses) > 0) then
        %>
              <b>Nombres en ingl�s</b>:<br/>
              <ul>
                <% for i=0 to ubound(array_nombres_ingleses) %>
                  <li><%= espaciar(array_nombres_ingleses(i)) %></li>
                <% next %>
              </ul>
        <%
            else
        %>
              <b>Nombre ingl�s</b>: <%= espaciar(nombre_ing) %><br/>
        <%
            end if  
           end if %>

				<% if (num_rd <> "") then response.write "<a onclick=window.open('ver_definicion.asp?id=86','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a> <b>RD</b>: "&num_rd&"<br/>" %>
				<% if (formula_molecular <> "") then response.write "<b>F�rmula molecular</b>: "&formula_molecular&"<br/>" %>
				<% if (estructura_molecular <> "") then response.write "<b>Estructura molecular</b>:<br /><img src='../gestion/estructuras/"&estructura_molecular&"' /><br/>" %>

				<% if (notas_xml <> "") then %>
          <a onclick="window.open('ver_definicion.asp?id=<%=dame_id_definicion("ECB")%>', 'def', 'width=300,height=200,scrollbars=yes,resizable=yes')" style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> 
          <b>Notas ECB</b>: <%= espaciar(notas_xml) %> <br />
        <% end if %>

        <% if (companias <> "") then %>
          <b>Compa��as distribuidoras</b>: <%= companias %>
        <% end if %>
			</td>
		</tr>
	<% end if 
end sub ' ap1_identificacion

' ###################################################################################

sub ap2_clasificacion()
	' Solo mostramos este apartado si hay informaci�n para �l
	if ((simbolos <> "") or (clasificacion_1 <> "") or (clasificacion_2 <> "") or (clasificacion_3 <> "") or (clasificacion_4 <> "") or (clasificacion_5 <> "") or (clasificacion_6 <> "") or (clasificacion_7 <> "") or (clasificacion_8 <> "") or (clasificacion_9 <> "") or (clasificacion_10 <> "") or (clasificacion_11 <> "") or (clasificacion_12 <> "") or (clasificacion_13 <> "") or (clasificacion_14 <> "") or (clasificacion_15 <> "") or (frases_r_danesa <> "") or (notas_rd_363 <> "") or (conc_1 <> "") or (eti_conc_1 <> "") or (conc_2 <> "") or (eti_conc_2 <> "") or (conc_3 <> "") or (eti_conc_3 <> "") or (conc_4 <> "") or (eti_conc_4 <> "") or (conc_5 <> "") or (eti_conc_5 <> "") or (conc_6 <> "") or (eti_conc_6 <> "") or (conc_7 <> "") or (eti_conc_7 <> "") or (conc_8 <> "") or (eti_conc_8 <> "") or (conc_9 <> "") or (eti_conc_9 <> "") or (conc_10 <> "") or (eti_conc_10 <> "") or (conc_11 <> "") or (eti_conc_11 <> "") or (conc_12 <> "") or (eti_conc_12 <> "") or (conc_13 <> "") or (eti_conc_13 <> "") or (conc_14 <> "") or (eti_conc_14 <> "") or (conc_15 <> "") or (eti_conc_15 <> "") or (estado_1 <> "") or (vla_ed_ppm_1 <> "") or (vla_ed_mg_m3_1 <> "") or (vla_ec_ppm_1 <> "") or (vla_ec_mg_m3_1 <> "") or (notas_vla_1 <> "") or (estado_2 <> "") or (vla_ed_ppm_2 <> "") or (vla_ed_mg_m3_2 <> "") or (vla_ec_ppm_2 <> "") or (vla_ec_mg_m3_2 <> "") or (notas_vla_2 <> "") or (estado_3 <> "") or (vla_ed_ppm_3 <> "") or (vla_ed_mg_m3_3 <> "") or (vla_ec_ppm_3 <> "") or (vla_ec_mg_m3_3 <> "") or (notas_vla_3 <> "") or (estado_4 <> "") or (vla_ed_ppm_4 <> "") or (vla_ed_mg_m3_4 <> "") or (vla_ec_ppm_4 <> "") or (vla_ec_mg_m3_4 <> "") or (notas_vla_4 <> "") or (estado_5 <> "") or (vla_ed_ppm_5 <> "") or (vla_ed_mg_m3_5 <> "") or (vla_ec_ppm_5 <> "") or (vla_ec_mg_m3_5 <> "") or (notas_vla_5 <> "") or (estado_6 <> "") or (vla_ed_ppm_6 <> "") or (vla_ed_mg_m3_6 <> "") or (vla_ec_ppm_6  <> "") or (vla_ec_mg_m3_6 <> "") or (notas_vla_6 <> "") or (ib_1 <> "") or  (vlb_1 <> "") or (momento_1 <> "") or (notas_vlb_1 <> "") or (ib_2 <> "") or  (vlb_2 <> "") or (momento_2 <> "") or (notas_vlb_2 <> "") or (ib_3 <> "") or  (vlb_3 <> "") or (momento_3 <> "") or (notas_vlb_3 <> "") or (ib_4 <> "") or  (vlb_4 <> "") or (momento_4 <> "") or (notas_vlb_4 <> "") or (ib_5 <> "") or  (vlb_5 <> "") or (momento_5 <> "") or (notas_vlb_5 <> "") or (ib_6 <> "") or  (vlb_6 <> "") or (momento_6 <> "") or (notas_vlb_6 <> "") or esta_en_lista_cancer_rd or esta_en_lista_cancer_iarc_excepto_grupo_3 or esta_en_lista_cancer_otras or esta_en_lista_de or esta_en_lista_neurotoxico or  esta_en_lista_tpb or esta_en_lista_sensibilizante or esta_en_lista_tpr or esta_en_lista_cancer_mama or esta_en_lista_cop) then

%>
	<!-- ################ Clasificaci�n ###################### -->
	<table id="tabla_clasificacionm" class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
  <tr>
		<td class="celdaabajo" colspan="2" align="center">
			<table cellpadding=0 cellspacing=0 width="100%" border="0"><tr><td width="100%" class="titulo3" align="left"><a onclick=window.open('ver_definicion.asp?id=87','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> CLASIFICACI�N</td></tr></table>
		</td>
	</tr>
	<!-- Simbolos y frases R -->
	<tr>
		<td valign="top">
			<% ap2_clasificacion_simbolos() %>
		</td>
		<td valign="top">
			<% ap2_clasificacion_frases_r() %>
			<% 
        if frases_r_danesa_mostradas then
          ap2_clasificacion_frases_r_danesa()
        end if
      %>
			<% ap2_clasificacion_frases_s() %>
			<% ap2_clasificacion_notas() %>
			<% ap2_clasificacion_etiquetado() %>
		</td>
	</tr>

	<tr>
		<td colspan="2">
      <!-- Valores l�mite -->
      <% ap2_clasificacion_vl() %>
    </td>
	</tr>

	<tr>
		<td valign="top" colspan="2">
			<!-- Lista negra -->
			<% ap2_clasificacion_lista_negra() %>
		</td>
	</tr>
	</table>
<%
	end if
end sub ' ap2_clasificacion

' ##################################################################################

sub ap2_clasificacion_simbolos()
	if (simbolos <> "") then
%>
		<p id="ap2_clasificacion_simbolos_titulo" class="ficha_titulo_2">S�mbolos</p>
		<p id="ap2_clasificacion_simbolos_cuerpo" class="texto" align="center">		
<%
		' Tiene s�mbolos, muestro cada uno
		simbolos = replace(simbolos, ",", ";")
		array_simbolos = split(simbolos, ";")
		for i=0 to ubound(array_simbolos)
			simbolo = trim(array_simbolos(i))
			imagen = imagen_simbolo(simbolo)
			descripcion = describe_simbolo(simbolo)
      if (simbolo <> "") then
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

sub ap2_clasificacion_frases_r()
	' Muestra las frases R segun clasificacion_1 hasta clasificacion_15
	' No incluye las frases R danesas

	' Montamos frases R
	frases_r=monta_frases_r(clasificacion_1, clasificacion_2, clasificacion_3, clasificacion_4, clasificacion_5, clasificacion_6, clasificacion_7, clasificacion_8, clasificacion_9, clasificacion_10, clasificacion_11, clasificacion_12, clasificacion_13, clasificacion_14, clasificacion_15)

	if (frases_r <> "") then
%>
		<p id="ap2_clasificacion_frases_r_titulo" class="ficha_titulo_2" style="margin-bottom: -10px;"><a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("Frases R")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> Frases R</p>
<%
		bucle_frases_r(frases_r)
	end if
end sub

' ##################################################################################

sub bucle_frases_r(byval frases_r)
		' Pasandole las frases R separadas por comas, muestra cada una junto a su descripci�n
		array_frases_r = split(frases_r, ",")
%>
    <blockquote style="margin-left: 10px; margin-bottom: -20px;">
<%
    ' Apuntamos las que hemos mostrado por si hay repetidas
    frases_mostradas = ";" 

		for i=0 to ubound(array_frases_r)
			frase = trim(array_frases_r(i))
      if(instr(frases_mostradas,frase+";") = 0) then
  			descripcion = describe_frase_r(frase)
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
		' Pasandole las frases S separadas por gui�n, muestra cada una junto a su descripci�n
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

sub ap2_clasificacion_frases_r_danesa()
	' Muestra las frases R danesas

	' Montamos frases R
	frases_r = monta_frases_r_danesa(frases_r_danesa)

	if (frases_r <> "") then
%>
	<p id="ap2_clasificacion_frases_r_danesa_titulo" class="ficha_titulo_2"><a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("Frases R seg�n la lista danesa")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> Frases R seg�n clasificaci�n de la EPA danesa</p>
<%
		bucle_frases_r(frases_r)
	end if
end sub

' ##################################################################################

sub ap2_clasificacion_frases_s
	' Muestra las frases S

	if (frases_s <> "") then
		' Eliminamos los par�ntesis de las frases S
		frases_s = replace (frases_s, "(", "")
		frases_s = replace (frases_s, ")", "")

%>
	<p id="ap2_clasificacion_frases_s_titulo" class="ficha_titulo_2" style="margin-top: 14px;"><a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("Frases S")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> Frases S <% plegador "secc-frasess", "img-frasess" %></p>
		<!-- <%= frases_s %> <a onclick="window.open('busca_frases_s.asp?id=<%= frases_s %>', 'fr','width=300,height=200,scrollbars=yes,resizable=yes')" style="text-decoration:none; cursor:hand;"><img src="imagenes/ayuda.gif" border="0" align="absmiddle" alt="busca Frases S"></a> -->

		<% bucle_frases_s frases_s%>

<%
	end if
end sub

' ##################################################################################

sub ap2_clasificacion_notas()
	if (notas_rd_363 <> "") then

		' Dividimos las notas, separadas por puntos, en un array
		array_notas = split(notas_rd_363, ".")
%>
	<p id="ap2_clasificacion_notas_titulo" class="ficha_titulo_2">Notas <% plegador "secc-notas", "img-notas" %></p>
	<p class="texto" >
		<blockquote id="secc-notas" style="display:none">
<%
		for i=0 to ubound(array_notas)
			nota = trim(array_notas(i))
			id_nota = dame_id_definicion(nota)
%>
			<b><%=nota%></b> <a onclick=window.open('ver_definicion.asp?id=<%=id_nota%>','def','width=600,height=400,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a><br />
<%
		next
%>
		</blockquote>
	</p>
<%
	end if
end sub

' ##################################################################################

sub ap2_clasificacion_etiquetado()
	' Muestra el etiquetado

	if ((conc_1 <> "") or (eti_conc_1 <> "") or (conc_2 <> "") or (eti_conc_2 <> "") or (conc_3 <> "") or (eti_conc_3 <> "") or (conc_4 <> "") or (eti_conc_4 <> "") or (conc_5 <> "") or (eti_conc_5 <> "") or (conc_6 <> "") or (eti_conc_6 <> "") or (conc_7 <> "") or (eti_conc_7 <> "") or (conc_8 <> "") or (eti_conc_8 <> "") or (conc_9 <> "") or (eti_conc_9 <> "") or (conc_10 <> "") or (eti_conc_10 <> "") or (conc_11 <> "") or (eti_conc_11 <> "") or (conc_12 <> "") or (eti_conc_12 <> "") or (conc_13 <> "") or (eti_conc_13 <> "") or (conc_14 <> "") or (eti_conc_14 <> "") or (conc_15 <> "") or (eti_conc_15 <> "")) then

%>
	<span id="ap2_clasificacion_etiquetado_titulo" class="ficha_titulo_2"><a onclick=window.open('ver_definicion.asp?id=88','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> Etiquetado <% plegador "secc-etiquetado", "img-etiquetado" %></span>
  <fieldset id="secc-etiquetado" style="display:none; margin: 15px 45px;">
	<table cellspacing="0" cellpadding="3" width="100%" align="center">
		<tr>
			<th class="subtitulo3 celdaabajo">Concentraci�n</th><th class="subtitulo3 celdaabajo">Etiquetado</th>
		</tr>
<%
	ap2_clasificacion_etiquetado_fila	conc_1, eti_conc_1
	ap2_clasificacion_etiquetado_fila	conc_2, eti_conc_2
	ap2_clasificacion_etiquetado_fila	conc_3, eti_conc_3
	ap2_clasificacion_etiquetado_fila	conc_4, eti_conc_4
	ap2_clasificacion_etiquetado_fila	conc_5, eti_conc_5
	ap2_clasificacion_etiquetado_fila	conc_6, eti_conc_6
	ap2_clasificacion_etiquetado_fila	conc_7, eti_conc_7
	ap2_clasificacion_etiquetado_fila	conc_8, eti_conc_8
	ap2_clasificacion_etiquetado_fila	conc_9, eti_conc_9
	ap2_clasificacion_etiquetado_fila	conc_10, eti_conc_10
	ap2_clasificacion_etiquetado_fila	conc_11, eti_conc_11
	ap2_clasificacion_etiquetado_fila	conc_12, eti_conc_12
	ap2_clasificacion_etiquetado_fila	conc_13, eti_conc_13
	ap2_clasificacion_etiquetado_fila	conc_14, eti_conc_14
	ap2_clasificacion_etiquetado_fila	conc_15, eti_conc_15
%>
	</table>
  </fieldset>

<%
	end if
end sub

' ##################################################################################
sub ap2_clasificacion_etiquetado_fila(byval c, byval e)
	c = replace (c, ":", "")
  c = replace (c, "<", "&lt;")
  c = replace (c, ">", "&gt;")

	if (c <> "") and (e <> "") then
%>
			<tr>
				<td class="celdaabajo"><%= h(c) %></td><td class="celdaabajo"><%= h(e) %> <a onclick="window.open('busca_frases_r.asp?id=<%=e%>','fr','width=300,height=200,scrollbars=yes,resizable=yes')" style="text-decoration:none;cursor:hand"><img src="imagenes/ayuda.gif" border="0" align="absmiddle" alt="busca Frases R"></a></td>
			</tr>
<%
	end if
end sub

' ##################################################################################

sub ap2_clasificacion_vl()
	if ((estado_1 <> "") or (vla_ed_ppm_1 <> "") or (vla_ed_mg_m3_1 <> "") or (vla_ec_ppm_1 <> "") or (vla_ec_mg_m3_1 <> "") or (notas_vla_1 <> "") or (estado_2 <> "") or (vla_ed_ppm_2 <> "") or (vla_ed_mg_m3_2 <> "") or (vla_ec_ppm_2 <> "") or (vla_ec_mg_m3_2 <> "") or (notas_vla_2 <> "") or (estado_3 <> "") or (vla_ed_ppm_3 <> "") or (vla_ed_mg_m3_3 <> "") or (vla_ec_ppm_3 <> "") or (vla_ec_mg_m3_3 <> "") or (notas_vla_3 <> "") or (estado_4 <> "") or (vla_ed_ppm_4 <> "") or (vla_ed_mg_m3_4 <> "") or (vla_ec_ppm_4 <> "") or (vla_ec_mg_m3_4 <> "") or (notas_vla_4 <> "") or (estado_5 <> "") or (vla_ed_ppm_5 <> "") or (vla_ed_mg_m3_5 <> "") or (vla_ec_ppm_5 <> "") or (vla_ec_mg_m3_5 <> "") or (notas_vla_5 <> "") or (estado_6 <> "") or (vla_ed_ppm_6 <> "") or (vla_ed_mg_m3_6 <> "") or (vla_ec_ppm_6  <> "") or (vla_ec_mg_m3_6 <> "") or (notas_vla_6 <> "") or (ib_1 <> "") or  (vlb_1 <> "") or (momento_1 <> "") or (notas_vlb_1 <> "") or (ib_2 <> "") or  (vlb_2 <> "") or (momento_2 <> "") or (notas_vlb_2 <> "") or (ib_3 <> "") or  (vlb_3 <> "") or (momento_3 <> "") or (notas_vlb_3 <> "") or (ib_4 <> "") or  (vlb_4 <> "") or (momento_4 <> "") or (notas_vlb_4 <> "") or (ib_5 <> "") or  (vlb_5 <> "") or (momento_5 <> "") or (notas_vlb_5 <> "") or (ib_6 <> "") or  (vlb_6 <> "") or (momento_6 <> "") or (notas_vlb_6 <> "")) then

%>
	<table border="0" width="100%" cellpadding="0" cellspacing="0">
		<tr>
			<td valign="top" width="50%">
			<!-- VLA -->
	
<%
		' VLA
		if ((estado_1 <> "") or (vla_ed_ppm_1 <> "") or (vla_ed_mg_m3_1 <> "") or (vla_ec_ppm_1 <> "") or (vla_ec_mg_m3_1 <> "") or (notas_vla_1 <> "") or (estado_2 <> "") or (vla_ed_ppm_2 <> "") or (vla_ed_mg_m3_2 <> "") or (vla_ec_ppm_2 <> "") or (vla_ec_mg_m3_2 <> "") or (notas_vla_2 <> "") or (estado_3 <> "") or (vla_ed_ppm_3 <> "") or (vla_ed_mg_m3_3 <> "") or (vla_ec_ppm_3 <> "") or (vla_ec_mg_m3_3 <> "") or (notas_vla_3 <> "") or (estado_4 <> "") or (vla_ed_ppm_4 <> "") or (vla_ed_mg_m3_4 <> "") or (vla_ec_ppm_4 <> "") or (vla_ec_mg_m3_4 <> "") or (notas_vla_4 <> "") or (estado_5 <> "") or (vla_ed_ppm_5 <> "") or (vla_ed_mg_m3_5 <> "") or (vla_ec_ppm_5 <> "") or (vla_ec_mg_m3_5 <> "") or (notas_vla_5 <> "") or (estado_6 <> "") or (vla_ed_ppm_6 <> "") or (vla_ed_mg_m3_6 <> "") or (vla_ec_ppm_6  <> "") or (vla_ec_mg_m3_6 <> "") or (notas_vla_6 <> "")) then
%>
	<span id="ap2_clasificacion_vla_titulo" class="ficha_titulo_1"><a href="index.asp?idpagina=616"><img src="http://www.istas.net/ecoinformas/web/imagenes/ayuda.gif" align="absmiddle" border="0" /></a>Valores L�mite Ambientales <% plegador "secc-vla", "img-vla" %></span>
	<fieldset id="secc-vla" style="display:none">
	<table border="0" width="100%" cellspacing="0" cellpadding="3">
		<tr>
			<% if ap2_clasificacion_vl_a_hay_columna_estado then %>
				<td class="subtitulo3 celdaabajo">Estado</td>
			<% end if %>

			<% if (ap2_clasificacion_vl_a_hay_columna_vla_ed_ppm or  ap2_clasificacion_vl_a_hay_columna_vla_ed_mg_m3) then %>
				<td class="subtitulo3 celdaabajo">VLA-ED</td>
			<% end if %>

			<% if (ap2_clasificacion_vl_a_hay_columna_vla_ec_ppm or  ap2_clasificacion_vl_a_hay_columna_vla_ec_mg_m3) then %>
				<td class="subtitulo3 celdaabajo">VLA-EC</td>
			<% end if %>

			<% if (ap2_clasificacion_vl_a_hay_columna_notas_vla) then %>
				<td class="subtitulo3 celdaabajo" width="25%">Notas</td>
			<% end if %>
		</tr>
<%
		ap2_clasificacion_vl_a estado_1, vla_ed_ppm_1, vla_ed_mg_m3_1, vla_ec_ppm_1, vla_ec_mg_m3_1, notas_vla_1
		ap2_clasificacion_vl_a estado_2, vla_ed_ppm_2, vla_ed_mg_m3_2, vla_ec_ppm_2, vla_ec_mg_m3_2, notas_vla_2
		ap2_clasificacion_vl_a estado_3, vla_ed_ppm_3, vla_ed_mg_m3_3, vla_ec_ppm_3, vla_ec_mg_m3_3, notas_vla_3
		ap2_clasificacion_vl_a estado_4, vla_ed_ppm_4, vla_ed_mg_m3_4, vla_ec_ppm_4, vla_ec_mg_m3_4, notas_vla_4
		ap2_clasificacion_vl_a estado_5, vla_ed_ppm_5, vla_ed_mg_m3_5, vla_ec_ppm_5, vla_ec_mg_m3_5, notas_vla_5
		ap2_clasificacion_vl_a estado_6, vla_ed_ppm_6, vla_ed_mg_m3_6, vla_ec_ppm_6, vla_ec_mg_m3_6, notas_vla_6
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
		if ((ib_1 <> "") or (vlb_1 <> "") or (momento_1 <> "") or (notas_vlb_1 <> "") or (ib_2 <> "") or (vlb_2 <> "") or (momento_2 <> "") or (notas_vlb_2 <> "") or (ib_3 <> "") or (vlb_3 <> "") or (momento_3 <> "") or (notas_vlb_3 <> "") or (ib_4 <> "") or (vlb_4 <> "") or (momento_4 <> "") or (notas_vlb_4 <> "") or (ib_5 <> "") or (vlb_5 <> "") or (momento_5 <> "") or (notas_vlb_51 <> "") or (ib_6 <> "") or (vlb_6 <> "") or (momento_6 <> "") or (notas_vlb_6 <> "")) then
%>

		<p id="ap2_clasificacion_vlb_titulo" class="ficha_titulo_1"><a href="index.asp?idpagina=616"><img src="http://www.istas.net/ecoinformas/web/imagenes/ayuda.gif" align="absmiddle" border="0" /></a>Valores L�mite Biol�gicos <% plegador "secc-vlb", "img-vlb" %></p>
		<fieldset id="secc-vlb" style="display:none">
		<table width="100%" cellspacing="0" cellpadding="3">
			<tr>
			<% if ap2_clasificacion_vl_b_hay_columna_ib then %>
				<td class="subtitulo3 celdaabajo">Indicador</th>
			<% end if %>

			<% if ap2_clasificacion_vl_b_hay_columna_vlb then %>
				<td class="subtitulo3 celdaabajo">Valor l�mite</th>
			<% end if %>

			<% if ap2_clasificacion_vl_b_hay_columna_momento then %>
				<td class="subtitulo3 celdaabajo">Momento de muestreo</th>
			<% end if %>

			<% if ap2_clasificacion_vl_b_hay_columna_notas_vlb then %>
				<td class="subtitulo3 celdaabajo">Notas</th>				
			<% end if %>
			</tr>
<%
			ap2_clasificacion_vl_b ib_1, vlb_1, momento_1, notas_vlb_1
			ap2_clasificacion_vl_b ib_2, vlb_2, momento_2, notas_vlb_2
			ap2_clasificacion_vl_b ib_3, vlb_3, momento_3, notas_vlb_3
			ap2_clasificacion_vl_b ib_4, vlb_4, momento_4, notas_vlb_4
			ap2_clasificacion_vl_b ib_5, vlb_5, momento_5, notas_vlb_5
			ap2_clasificacion_vl_b ib_6, vlb_6, momento_6, notas_vlb_6
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
	' Mostramos una fila si hay alg�n dato
	if (estado&vla_ed_ppm&vla_ed_mg_m3&vla_ec_ppm&vla_ec_mg_m3&notas_vla <> "") then
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
	' Pinta una fila si hay alg�n dato
	if (ib&vlb&momento&notas_vlb <> "") then
%>
			<tr>
			<% if ap2_clasificacion_vl_b_hay_columna_ib then %>
				<td style="	border-bottom-color: #DDDDDD;	border-bottom-style: solid;	border-bottom-width: 1px;" valign="middle"><%=ib%></td>
			<% end if %>

			<% if ap2_clasificacion_vl_b_hay_columna_vlb then %>
				<td style="	border-bottom-color: #DDDDDD;	border-bottom-style: solid;	border-bottom-width: 1px;" valign="middle"><%=vlb%></td>
			<% end if %>

			<% if ap2_clasificacion_vl_b_hay_columna_momento then %>
				<td style="	border-bottom-color: #DDDDDD;	border-bottom-style: solid;	border-bottom-width: 1px;" valign="middle"><%=parche_definicion(momento, "MomentoVLBInicio")%><a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion(parche_definicion(momento, "MomentoVLB"))%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><%= parche_definicion(momento, "MomentoVLB") %></a>
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
	valores = estado_1 & estado_2 & estado_3 & estado_4 & estado_5 & estado_6
	if (valores <> "") then
		ap2_clasificacion_vl_a_hay_columna_estado = true
	else
		ap2_clasificacion_vl_a_hay_columna_estado = false
	end if
end function

' ##################################################################################

function ap2_clasificacion_vl_a_hay_columna_vla_ed_ppm()
	valores = vla_ed_ppm_1 & vla_ed_ppm_2 & vla_ed_ppm_3 & vla_ed_ppm_4 & vla_ed_ppm_5 & vla_ed_ppm_6
	if (valores <> "") then
		ap2_clasificacion_vl_a_hay_columna_vla_ed_ppm = true
	else
		ap2_clasificacion_vl_a_hay_columna_vla_ed_ppm = false
	end if
end function

' ##################################################################################

function ap2_clasificacion_vl_a_hay_columna_vla_ed_mg_m3()
	valores = vla_ed_mg_m3_1 & vla_ed_mg_m3_2 & vla_ed_mg_m3_3 & vla_ed_mg_m3_4 & vla_ed_mg_m3_5 & vla_ed_mg_m3_6
	if (valores <> "") then
		ap2_clasificacion_vl_a_hay_columna_vla_ed_mg_m3 = true
	else
		ap2_clasificacion_vl_a_hay_columna_vla_ed_mg_m3 = false
	end if
end function

' ##################################################################################

function ap2_clasificacion_vl_a_hay_columna_vla_ec_ppm()
	valores = vla_ec_ppm_1 & vla_ec_ppm_2 & vla_ec_ppm_3 & vla_ec_ppm_4 & vla_ec_ppm_5 & vla_ec_ppm_6
	if (valores <> "") then
		ap2_clasificacion_vl_a_hay_columna_vla_ec_ppm = true
	else
		ap2_clasificacion_vl_a_hay_columna_vla_ec_ppm = false
	end if
end function

' ##################################################################################

function ap2_clasificacion_vl_a_hay_columna_vla_ec_mg_m3()
	valores = vla_ec_mg_m3_1 & vla_ec_mg_m3_2 & vla_ec_mg_m3_3 & vla_ec_mg_m3_4 & vla_ec_mg_m3_5 & vla_ec_mg_m3_6
	if (valores <> "") then
		ap2_clasificacion_vl_a_hay_columna_vla_ec_mg_m3 = true
	else
		ap2_clasificacion_vl_a_hay_columna_vla_ec_mg_m3 = false
	end if
end function

' ##################################################################################

function ap2_clasificacion_vl_a_hay_columna_notas_vla()
	valores = notas_vla_1 & notas_vla_2 & notas_vla_3 & notas_vla_4 & notas_vla_5 & notas_vla_6
	if (valores <> "") then
		ap2_clasificacion_vl_a_hay_columna_notas_vla = true
	else
		ap2_clasificacion_vl_a_hay_columna_notas_vla = false
	end if
end function

' ##################################################################################

function ap2_clasificacion_vl_b_hay_columna_ib()
	valores = ib_1 & ib_2 & ib_3 & ib_4 & ib_5 & ib_6
	if (valores <> "") then
		ap2_clasificacion_vl_b_hay_columna_ib = true
	else
		ap2_clasificacion_vl_b_hay_columna_ib = false
	end if
end function

' ##################################################################################

function ap2_clasificacion_vl_b_hay_columna_vlb()
	valores = vlb_1 & vlb_2 & vlb_3 & vlb_4 & vlb_5 & vlb_6
	if (valores <> "") then
		ap2_clasificacion_vl_b_hay_columna_vlb = true
	else
		ap2_clasificacion_vl_b_hay_columna_vlb = false
	end if
end function

' ##################################################################################

function ap2_clasificacion_vl_b_hay_columna_momento()
	valores = momento_1 & momento_2 & momento_3 & momento_4 & momento_5 & momento_6
	if (valores <> "") then
		ap2_clasificacion_vl_b_hay_columna_momento = true
	else
		ap2_clasificacion_vl_b_hay_columna_momento = false
	end if
end function

' ##################################################################################

function ap2_clasificacion_vl_b_hay_columna_momento()
	valores = momento_1 & momento_2 & momento_3 & momento_4 & momento_5 & momento_6
	if (valores <> "") then
		ap2_clasificacion_vl_b_hay_columna_momento = true
	else
		ap2_clasificacion_vl_b_hay_columna_momento = false
	end if
end function

' ##################################################################################

function ap2_clasificacion_vl_b_hay_columna_notas_vlb()
	valores = notas_vlb_1 & notas_vlb_2 & notas_vlb_3 & notas_vlb_4 & notas_vlb_5 & notas_vlb_6
	if (valores <> "") then
		ap2_clasificacion_vl_b_hay_columna_notas_vlb = true
	else
		ap2_clasificacion_vl_b_hay_columna_notas_vlb = false
	end if
end function

' ##################################################################################

sub notas_con_ayuda(byval notas, byval tipo)

	' Para buscar la definici�n hay ocasiones en las que hay que aplicar un parche.

	array_notas = split(notas, ",")
	cadena_notas = ""				
	for i=0 to ubound(array_notas)
		nota = trim(array_notas(i))
		id_nota = dame_id_definicion(parche_definicion(nota, tipo))
		if (nota <> "") then
			if (cadena_notas = "") then
				cadena_notas = "<a onclick=window.open('ver_definicion.asp?id="&id_nota&"','def','width=600,height=400,scrollbars=yes,resizable=yes') style='cursor:hand'>"&nota&"</a>"
			else
				cadena_notas = cadena_notas & ", <a onclick=window.open('ver_definicion.asp?id="&id_nota&"','def','width=600,height=400,scrollbars=yes,resizable=yes') style='cursor:hand'>"&nota&"</a>"
			end if
		end if
	next
	response.write cadena_notas
end sub

' ##################################################################################

sub ap2_clasificacion_lista_negra()
	' Muestra el etiquetado

	if (esta_en_lista_cancer_rd or esta_en_lista_cancer_danesa or esta_en_lista_cancer_iarc_excepto_grupo_3 or esta_en_lista_cancer_otras or esta_en_lista_de or esta_en_lista_neurotoxico or  esta_en_lista_tpb or esta_en_lista_sensibilizante or esta_en_lista_sensibilizante_danesa or esta_en_lista_tpr or esta_en_lista_tpr_danesa or esta_en_lista_mutageno_rd or esta_en_lista_mutageno_danesa or esta_en_lista_cancer_mama or esta_en_lista_cop) then

    ' Esta en lista negra. Aprovechamos para marcarle el bit correspondiente para que aparezca en el listado de lista negra
    sqlListaNegra="UPDATE dn_risc_sustancias SET negra=1 WHERE id="&id_sustancia
    objConnection2.execute(sqlListaNegra),,adexecutenorecords
    
    ' OK, continuamos...

		razones = ""

		if (esta_en_lista_cancer_rd or esta_en_lista_cancer_danesa or esta_en_lista_cancer_iarc_excepto_grupo_3 or esta_en_lista_cancer_otras or esta_en_lista_cancer_mama) then
			if (razones = "") then
				razones = "cancer�gena"
			else
				razones = razones & ", cancer�gena"
			end if
		end if
	
		if (esta_en_lista_cop) then
			if (razones = "") then
				razones = "cop"
			else
				razones = razones & ", COP"
			end if
		end if

		if (esta_en_lista_mutageno_rd or esta_en_lista_mutageno_danesa) then
			if (razones = "") then
				razones = "mut�gena"
			else
				razones = razones & ", mut�gena"
			end if
		end if

		if (esta_en_lista_de) then
			if (razones = "") then
				razones = "disruptora endocrina"
			else
				razones = razones & ", disruptora endocrina"
			end if	
		end if

		if (esta_en_lista_neurotoxico) then
			if (razones = "") then
				razones = "neurot�xica"
			else
				razones = razones & ", neurot�xica"
			end if
		end if

		if (esta_en_lista_tpb) then
			if (razones = "") then
				razones = "t�xica, persistente y bioacumulativa"
			else
				razones = razones & ", t�xica, persistente y bioacumulativa"
			end if
		end if

		if (esta_en_lista_sensibilizante or esta_en_lista_sensibilizante_danesa) then
			if (razones = "") then
				razones = "sensibilizante"
			else
				razones = razones & ", sensibilizante"
			end if

		end if
	
		if (esta_en_lista_tpr or esta_en_lista_tpr_danesa) then
			if (razones = "") then
				razones = "t�xica para la reproducci�n"
			else
				razones = razones & ", t�xica para la reproducci�n"
			end if
		end if
%>
		<p id="ap2_clasificacion_lista_negra_titulo" class="subtitulo3">&nbsp;<img src="imagenes/icono_atencion_20.png" align="absmiddle" /> <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("Lista negra")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> Sustancia incluida en la Lista negra de ISTAS <% plegador "secc-listanegra", "img-listanegra" %></p>
		<p id="secc-listanegra" class="texto" style="display:none">Esta sustancia est� incluida en la Lista negra de ISTAS ya que es <%=razones%></p>

<%
	end if
end sub

' ###################################################################################

sub ap3_riesgos()
	' SALUD

	if (esta_en_lista_cancer_rd or esta_en_lista_cancer_danesa or esta_en_lista_cancer_iarc or esta_en_lista_cancer_otras or esta_en_lista_cancer_mama or esta_en_lista_de or esta_en_lista_neurotoxico or esta_en_lista_sensibilizante or esta_en_lista_sensibilizante_danesa or esta_en_lista_tpr or esta_en_lista_tpr_danesa or esta_en_lista_eepp or esta_en_lista_mutageno_rd or esta_en_lista_mutageno_danesa or esta_en_lista_salud) then
%>

		<!-- ################ Riesgos para la salud ###################### -->
		<br />
		<div id="ficha">
		<table width="100%" cellpadding=5>
			<tr>
				<td>
					<a name="identificacion"></a><img src="imagenes/risctox02.gif" alt="Riesgos espec�ficos para la salud" />
				</td>
				<td align="right">
					<a href="#"><img src="imagenes/subir.gif" border=0 alt=subir></a>
				</td>
			</tr>
		</table>

<%
		if (esta_en_lista_cancer_rd or esta_en_lista_cancer_danesa or esta_en_lista_cancer_iarc or esta_en_lista_cancer_otras or esta_en_lista_cancer_mama) then ap3_riesgos_tabla("Cancer�geno") end if
		if (esta_en_lista_mutageno_rd or esta_en_lista_mutageno_danesa) then ap3_riesgos_tabla("Mut�geno") end if
		if esta_en_lista_de then ap3_riesgos_tabla("Disruptor endocrino") end if
		if esta_en_lista_neurotoxico then ap3_riesgos_tabla("Neurot�xico") end if
		if esta_en_lista_sensibilizante or esta_en_lista_sensibilizante_danesa then ap3_riesgos_tabla("Sensibilizante") end if
		if esta_en_lista_tpr or esta_en_lista_tpr_danesa then ap3_riesgos_tabla("T�xico para la reproducci�n") end if
		if esta_en_lista_eepp then ap3_riesgos_enfermedades() end if
    if esta_en_lista_salud then ap7_salud() end if
%>
		</div>
		<!-- ################ Fin Riesgos para la salud ################## -->
<%
	end if ' salud 
%>

<% ' MEDIO AMBIENTE %>
<% if (esta_en_lista_tpb or esta_en_lista_directiva_aguas or esta_en_lista_alemana or esta_en_lista_ozono or esta_en_lista_clima or esta_en_lista_aire or esta_en_lista_cop) then %>

		<!-- ################ Riesgos para el medio ambiente ###################### -->
		<br />
		<div id="ficha">
		<table width="100%" cellpadding=5>
			<tr>
				<td>
					<a name="identificacion"></a><img src="imagenes/risctox03.gif" alt="Riesgos espec�ficos para el medio ambiente" />
				</td>
				<td align="right">
					<a href="#"><img src="imagenes/subir.gif" border=0 alt=subir></a>
				</td>
			</tr>
		</table>
<%
		if esta_en_lista_tpb then ap3_riesgos_tabla("T�xica, Persistente y Bioacumulativa") end if
		if (esta_en_lista_directiva_aguas or esta_en_lista_alemana) then ap3_riesgos_tabla("T�xica para el agua") end if
		if (esta_en_lista_ozono or esta_en_lista_clima or esta_en_lista_aire) then ap3_riesgos_tabla("Contaminante del aire") end if
		if (esta_en_lista_cop) then ap3_riesgos_tabla("Contaminante Org�nico Persistente (COP)") end if
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

        <% if ((tipo <> "COV") and (tipo <> "Vertidos") and (tipo <> "LPCIC (EPER Agua)") and (tipo <> "LPCIC (EPER Aire)") and (tipo <> "Residuos Peligrosos") and (tipo <> "Accidentes Graves") and (tipo <> "Emisiones Atmosf�ricas")) then %>

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
		case "Cancer�geno":
%>
			<a href="index.asp?idpagina=607"><img src="http://www.istas.net/ecoinformas/web/imagenes/ayuda.gif" align="absmiddle" border="0" /></a> 
<%
		case "Mut�geno":
%>
			<a href="index.asp?idpagina=607"><img src="http://www.istas.net/ecoinformas/web/imagenes/ayuda.gif" align="absmiddle" border="0" /></a> 
<%
		case "Disruptor endocrino":
%>
			<a href="index.asp?idpagina=610"><img src="http://www.istas.net/ecoinformas/web/imagenes/ayuda.gif" align="absmiddle" border="0" /></a> 
<%
		case "Neurot�xico":
%>
			<a href="index.asp?idpagina=611"><img src="http://www.istas.net/ecoinformas/web/imagenes/ayuda.gif" align="absmiddle" border="0" /></a> 
<%
		case "Sensibilizante":
%>
			<a href="index.asp?idpagina=612"><img src="http://www.istas.net/ecoinformas/web/imagenes/ayuda.gif" align="absmiddle" border="0" /></a> 
<%
		case "T�xico para la reproducci�n":
%>
			<a href="index.asp?idpagina=609"><img src="http://www.istas.net/ecoinformas/web/imagenes/ayuda.gif" align="absmiddle" border="0" /></a> 
<%
		case "T�xica, Persistente y Bioacumulativa":
%>
			<a href="index.asp?idpagina=613"><img src="http://www.istas.net/ecoinformas/web/imagenes/ayuda.gif" align="absmiddle" border="0" /></a> 
<%
		case "T�xica para el agua":
%>
			<a href="index.asp?idpagina=614"><img src="http://www.istas.net/ecoinformas/web/imagenes/ayuda.gif" align="absmiddle" border="0" /></a> 
<%
		case "Contaminante Org�nico Persistente (COP)":
%>
			<a href="index.asp?idpagina=1185"><img src="http://www.istas.net/ecoinformas/web/imagenes/ayuda.gif" align="absmiddle" border="0" /></a> 
<%
		case "Contaminante del aire":
%>
			<a href="index.asp?idpagina=615"><img src="http://www.istas.net/ecoinformas/web/imagenes/ayuda.gif" align="absmiddle" border="0" /></a> 
<%
		case "Residuos Peligrosos":
%>
			<a href="index.asp?idpagina=618"><img src="http://www.istas.net/ecoinformas/web/imagenes/ayuda.gif" align="absmiddle" border="0" /></a> 
<%
		case "Vertidos":
%>
			<a href="index.asp?idpagina=619"><img src="http://www.istas.net/ecoinformas/web/imagenes/ayuda.gif" align="absmiddle" border="0" /></a> 
<%
		case "Accidentes Graves":
%>
			<a href="index.asp?idpagina=623"><img src="http://www.istas.net/ecoinformas/web/imagenes/ayuda.gif" align="absmiddle" border="0" /></a> 
<%
		case "COV":
%>
			<a href="index.asp?idpagina=621"><img src="http://www.istas.net/ecoinformas/web/imagenes/ayuda.gif" align="absmiddle" border="0" /></a> 
<%
		case "LPCIC (EPER Agua)":
%>
			<a href="index.asp?idpagina=622"><img src="http://www.istas.net/ecoinformas/web/imagenes/ayuda.gif" align="absmiddle" border="0" /></a> 
<%
		case "LPCIC (EPER Aire)":
%>
			<a href="index.asp?idpagina=622"><img src="http://www.istas.net/ecoinformas/web/imagenes/ayuda.gif" align="absmiddle" border="0" /></a> 
<%
		case "Emisiones Atmosf�ricas":
%>
			<a href="index.asp?idpagina=620"><img src="http://www.istas.net/ecoinformas/web/imagenes/ayuda.gif" align="absmiddle" border="0" /></a> 
<%
	end select

end sub

' ###################################################################################

sub ap3_riesgos_tabla_contenidos(tipo)

	select case tipo
    case "Contaminante Org�nico Persistente (COP)":
%>
    <fieldset>
      <legend class="subtitulo3"><strong>Seg�n Convenio de Estocolmo</strong></legend>
      <ul>
<%
      if isNull(cop) then
        cop = ""
      end if

      array_anexos = split(cop, ";")
      for i=0 to ubound(array_anexos)
%>       
        <li><%=dame_definicion("COP Anexo "&trim(array_anexos(i)))%></li>
<%
      next
%>
      </ul>
    </fieldset>
<%
		case "Cancer�geno": 

		' Real Decreto ---------------------------------------------------------------
		if (esta_en_lista_cancer_rd) then
%>
			<fieldset>
				<legend class="subtitulo3"><strong>Seg�n Real Decreto 363/1995</strong></legend>
				<blockquote>
<%
		nivel_cancerigeno_rd = dame_nivel_cancerigeno_rd()
		if (nivel_cancerigeno_rd <> "") then
			response.write "<strong>Nivel cancer�geno:</strong> "&nivel_cancerigeno_rd
%>
			 <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("C"&nivel_cancerigeno_rd)%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>
<%
		end if
%>

<%
			if (notas_cancer_rd <> "") then
%>				
				<br/><strong>Notas:</strong> <%=notas_cancer_rd%>
<%
			end if
%>
				</blockquote>
			</fieldset>
<%
		end if



		' Lista danesa ---------------------------------------------------------------
		if (esta_en_lista_cancer_danesa) then
%>
			<fieldset>
				<legend class="subtitulo3"><strong>Seg�n <% plegador_texto "frases_r_danesa_cancer", "frases R", "subtitulo3" %> en la clasificaci�n de la EPA danesa <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("EPA Danesa")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>.</strong></legend>
				<blockquote>
<%
		nivel_cancerigeno_danesa = dame_nivel_cancerigeno_danesa()
		if (nivel_cancerigeno_danesa <> "") then
			response.write "<strong>Nivel cancer�geno:</strong> "&nivel_cancerigeno_danesa
%>
			 <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("C"&nivel_cancerigeno_danesa)%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>
<%
		end if
%>

<%
			if (notas_cancer_rd <> "") then
%>				
				<br/><strong>Notas:</strong> <%=notas_cancer_rd%>
<%
			end if
%>
        <div id="frases_r_danesa_cancer" style="display:none"><br />
        <% ap2_clasificacion_frases_r_danesa() %>
        </div>

				</blockquote>
			</fieldset>
<%
		end if



		' IARC -----------------------------------------------------------------------
		if (esta_en_lista_cancer_iarc) then
%>
			<fieldset>
				<legend class="subtitulo3"><strong>Seg�n IARC <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("IARC")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a></strong></legend>
<%
				if (grupo_iarc <> "") or (volumen_iarc <> "") or (notas_iarc <> "") then
%>
					<blockquote>
					<table>
<%
					if (grupo_iarc <> "") then						
%>
						<tr><td class="subtitulo3">Grupo:</td><td><%=trim(replace(ucase(grupo_iarc), "GRUPO", ""))%> <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion(trim(grupo_iarc))%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a></td></tr>
<%
					end if

					if (volumen_iarc <> "") then						
%>
						<tr><td class="subtitulo3">Volumen:</td><td><%=volumen_iarc%></td></tr>
<%
					end if
					if (notas_iarc <> "") then						
%>
						<tr><td class="subtitulo3">Notas:</td><td><%=notas_iarc%></td></tr>
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
		if (esta_en_lista_cancer_otras) then
%>
    <fieldset>
		  <legend class="subtitulo3"><strong>Seg�n otras fuentes</strong></legend>
<%
      if (isNull(categoria_cancer_otras)) then
        categoria_cancer_otras = ""
      end if

      if (isNull(fuente)) then
        fuente = ""
      end if

			array_categorias=split(categoria_cancer_otras, ",")
			array_fuentes=split(fuente, ",")

			' Damos por hecho que hay el mismo numero de categorias y fuentes y que coinciden en orden
			for i=0 to ubound(array_fuentes)
%>
			<fieldset>
				<legend class="subtitulo3"><strong>Seg�n <%=trim(array_fuentes(i))%> <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion(trim(array_fuentes(i)))%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a></strong></legend>
				<blockquote>	
				<table>
					<tr><td class="subtitulo3"><%=trim(array_categorias(i))%>:</td><td><%= dame_definicion(trim(array_categorias(i))) %></td></tr>
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
    if (esta_en_lista_cancer_mama) then
      if (isNull(cancer_mama_fuente)) then
        cancer_mama_fuente = ""
      end if
%>
			<fieldset>
				<legend class="subtitulo3"><strong>Seg�n SSI (c�ncer de mama) <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("SSI")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a></strong></legend>
				<blockquote>	
				<table>
					<tr><td class="subtitulo3"><strong>Fuente:</strong><br /><a href="<%= cancer_mama_fuente %>" target="_blank"><%= replace(cancer_mama_fuente, "http://", "") %></a></td></tr>
				</table>
				</blockquote>
			</fieldset>
<%
    end if

		case "Mut�geno":
      ' MUTAGENO RD -------------------------------------------------------------
      if (esta_en_lista_mutageno_rd) then
%>
			<fieldset>
				<legend class="subtitulo3"><strong>Seg�n Real Decreto 363/1995</strong></legend>
				<blockquote>
				<%
					nivel_mutageno_rd = dame_nivel_mutageno_rd()
					if (nivel_mutageno_rd <> "") then
					response.write "<br /><strong>Nivel mut�geno:</strong> "&nivel_mutageno_rd 
				%>
					 <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("M"&nivel_mutageno_rd)%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>
				<%
					end if
				%>
				</blockquote>
			</fieldset>
<%
      end if


      ' MUTAGENO DANESA -------------------------------------------------------------
      if (esta_en_lista_mutageno_danesa) then
%>
			<fieldset>
				<legend class="subtitulo3"><strong>Seg�n <% plegador_texto "frases_r_danesa_mutageno", "frases R", "subtitulo3" %> en la clasificaci�n de la EPA danesa <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("EPA Danesa")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>.</strong></legend>
				<blockquote>
				<%
					nivel_mutageno_danesa = dame_nivel_mutageno_danesa()
					if (nivel_mutageno_danesa <> "") then
					response.write "<br /><strong>Nivel mut�geno:</strong> "&nivel_mutageno_danesa 
				%>
					 <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("M"&nivel_mutageno_danesa)%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>
				<%
					end if
				%>

        <div id="frases_r_danesa_mutageno" style="display:none"><br />
        <% ap2_clasificacion_frases_r_danesa() %>
        </div>

				</blockquote>
			</fieldset>
<%
      end if




		case "Disruptor endocrino":
%>
			<blockquote>
			<table>
			<% if (nivel_disruptor <> "") then %>
				<tr>
					<td class="subtitulo3" valign="top">Fuente:</td>
					<td>
					<%
					array_niveles=split(nivel_disruptor, ",")
					for i=0 to ubound(array_niveles)
						nivel=dame_definicion(trim(array_niveles(i)))
						response.write nivel&"<br />"
					next
					%>
					</td>
				</tr>
			<% end if %>
			</table>
			</blockquote>
<%
		case "Neurot�xico":

        'response.write efecto_neurotoxico&"***"&fuente_neurotoxico

        if esta_en_lista_neurotoxico_rd or esta_en_lista_neurotoxico_danesa then
          ' A�adimos SNC a efecto neurotoxico si no exist�a ya
          if (efecto_neurotoxico = "") or (IsNull(efecto_neurotoxico)) then
            efecto_neurotoxico="SNC"
          else
            if (not (inStr(efecto_neurotoxico, "SNC") > 0)) then
              efecto_neurotoxico = efecto_neurotoxico & "/SNC"
            end if
          end if
        end if

        if esta_en_lista_neurotoxico_rd then
          if (fuente_neurotoxico = "") or (IsNull(fuente_neurotoxico)) then
            fuente_neurotoxico = "363"
          else
            fuente_neurotoxico = fuente_neurotoxico & ",363"
          end if
        end if

        if esta_en_lista_neurotoxico_danesa then
          if (fuente_neurotoxico = "") or (IsNull(fuente_neurotoxico)) then
            fuente_neurotoxico = "EPA-R67"
          else
            fuente_neurotoxico = fuente_neurotoxico & ",EPA-R67"
          end if
        end if
      %>


      <% if ((efecto_neurotoxico <> "") or (nivel_neurotoxico <> "") or (fuente_neurotoxico <> "")) then %>
			<blockquote>
			<table>
			<%	if (efecto_neurotoxico <> "") then %>
				<tr>
					<td class="subtitulo3" valign="top">Efecto:</td>
					<td>
						<%
							' Separamos el efecto neurotoxico por "/". Ejemplo: "SNC/NEUROTOXICO/OTOTOXICO" se convierte en 3 definiciones, cada una con su ayuda.
							array_neurotoxico = split(efecto_neurotoxico, "/")
							for i=0 to ubound(array_neurotoxico)
								efecto = trim(array_neurotoxico(i))
                efecto = ucase(efecto)
                'efecto = quitartildes(efecto)
                'efecto = montartildes(efecto)
						%>

						<%= efecto %> <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion(efecto)%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> 

						<%
							next
						%>
					</td>
				</tr>
			<% end if %>
			<% if (nivel_neurotoxico <> "") then %>
				<tr>
					<td class="subtitulo3" valign="top">Nivel:</td><td><%=nivel_neurotoxico%>

					 <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("Nivel "&nivel_neurotoxico)%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>

					</td>
			</tr>
			<% end if %>
			<% if (fuente_neurotoxico <> "") then %>
				<tr>
					<td class="subtitulo3" valign="top">Fuente:</td>
					<td>
					<%
					array_fuentes=split(fuente_neurotoxico, ",")
					for i=0 to ubound(array_fuentes)
						'fuente=dame_definicion(trim(array_fuentes(i)))
						'response.write fuente&"<br />"
          %>
            <%= trim(array_fuentes(i)) %>     
            <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion(trim(array_fuentes(i)))%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>        
          <%
            if (i < ubound(array_fuentes)) then
              response.write ", "
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
      if esta_en_lista_sensibilizante then
        response.write "<li class='subtitulo3'>Sensibilizante seg�n Real Decreto 363/1995</li>"
      end if

      if esta_en_lista_sensibilizante_danesa then
      %>
        <li class='subtitulo3'>Sensibilizante seg�n <% plegador_texto "frases_r_danesa_sensibilizante", "frases R", "subtitulo3" %> en la clasificaci�n de la EPA danesa <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("EPA Danesa")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a></li>
      <%      
      end if
      response.write "</ul>"
      %>    
        <div id="frases_r_danesa_sensibilizante" style="display:none"><br />
        <blockquote>
        <% ap2_clasificacion_frases_r_danesa() %>
        </blockquote>
        </div>
      <%

		case "T�xico para la reproducci�n":
      ' TPR SEGUN RD -------------------------------------------------------------
      if (esta_en_lista_tpr) then
%>
    			<fieldset>
  				<legend class="subtitulo3"><strong>Seg�n Real Decreto 363/1995</strong></legend>
<%
  			nivel_reproduccion_rd = dame_nivel_reproduccion_rd()
  			if (nivel_reproduccion_rd <> "") then
			  %>
  				<blockquote>
  					<strong>Categor�a:</strong> <%=nivel_reproduccion_rd%>
  				  <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("TR"&nivel_reproduccion_rd)%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>
  					</blockquote>
  			<%
  			end if
%>
          </fieldset>
<%
      end if

      ' TPR SEGUN LISTA DANESA ---------------------------------------------------
      if (esta_en_lista_tpr_danesa) then
%>
    			<fieldset>
  				<legend class="subtitulo3"><strong>Seg�n <% plegador_texto "frases_r_danesa_tpr", "frases R", "subtitulo3" %> en la clasificaci�n de la EPA danesa <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("EPA Danesa")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a></strong></legend>
<%
  			nivel_reproduccion_danesa = dame_nivel_reproduccion_danesa()
  			if (nivel_reproduccion_danesa <> "") then
			  %>
  				<blockquote>
  					<strong>Categor�a:</strong> <%=nivel_reproduccion_danesa%>
  				  <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("TR"&nivel_reproduccion_danesa)%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>
  					</blockquote>
  			<%
  			end if
%>
        <div id="frases_r_danesa_tpr" style="display:none"><br />
        <blockquote>
        <% ap2_clasificacion_frases_r_danesa() %>
        </blockquote>
        </div>
          </fieldset>
<%
      end if



		case "T�xica, Persistente y Bioacumulativa":
%>
			<blockquote>
			<table>
				<tr>
					<td class="subtitulo3">M�s informaci�n (en ingl�s):</td>
					<td><a href="<%= enlace_tpb %>"><%= corta(anchor_tpb, 70, "puntossuspensivos") %></a></td>
				</tr>
			</table>
			</blockquote>
<%

		case "T�xica para el agua":
			response.write "<table>"
			if (directiva_aguas or esta_en_lista_directiva_aguas) then
%>
				<tr>
					<td class="subtitulo3" colspan="2">� Seg�n Directiva de Aguas</td>
				</tr>
<%
			end if

			if (clasif_mma <> "") then
%>
				<tr>
					<td class="subtitulo3">
						� Seg�n Peligrosas Agua Alemania</td><td><strong>Clasificaci�n</strong>: <%=clasif_mma%>

					 <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion(parche_definicion(clasif_mma, "MMA"))%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>

					</td>
				</tr>
<%
			end if
			response.write "</table>"


			case "Contaminante del aire":
%>
				<table>
<%
				if (dano_calidad_aire or esta_en_lista_aire) then
%>
					<tr>
						<td class="subtitulo3">Calidad del aire:</td>
						<td>Sustancia incluida en la <a href="http://www.istas.net/ecoinformas/web/abreenlace.asp?idenlace=2234">Directiva 96/62/CE</a> de 27 de septiembre sobre evaluaci�n y gesti�n de la calidad del aire ambiente</td>
					</tr>
<%
				end if
%>			
<%
				if (dano_ozono) then
%>
					<tr>
						<td class="subtitulo3">Capa de ozono:</td>
						<td>Sustancia que agota la capa de ozono, seg�n <a href="http://www.istas.net/ecoinformas/web/abreenlace.asp?idenlace=2229">Reglamento (CE) 2037/2000</a> del Parlamento Europeo y del Consejo, de 29 de junio de 2000</td>
					</tr>
<%
				end if
%>			
<%
				if (dano_cambio_clima) then
%>
					<tr>
						<td class="subtitulo3">Cambio clim�tico:</td>
						<td>Sustancia incluida en el listado del <a href="http://www.istas.net/ecoinformas/web/abreenlace.asp?idenlace=2230">Protocolo de Kyoto</a></td>
					</tr>
<%
				end if
%>			
				</table>
<%	
	end select
end sub

' ###################################################################################

sub ap3_riesgos_enfermedades()
	' Se agrupan por listado, cada listado en una ficha blanca y dentro cada enfermedad
	sql_enf = "select distinct enf.id, enf.listado, enf.nombre, enf.sintomas, enf.actividades FROM dn_risc_enfermedades AS enf LEFT OUTER JOIN dn_risc_grupos_por_enfermedades AS gpe ON enf.id = gpe.id_enfermedad LEFT OUTER JOIN dn_risc_sustancias_por_grupos AS spg ON gpe.id_grupo = spg.id_grupo LEFT OUTER JOIN dn_risc_sustancias_por_enfermedades AS spe ON spe.id_enfermedad = enf.id WHERE spg.id_sustancia="&id_sustancia&" OR spe.id_sustancia="&id_sustancia&" ORDER BY enf.listado, enf.nombre"
	'response.write "<br />"&sql_enf
	set objRstEnf=objConnection2.execute(sql_enf)
	if (not objRstEnf.eof) then
		listado_antiguo = ""
		do while (not objRstEnf.eof)
			' Para mostrar agrupados por listado, solo escribimos la cabecera si el listado es nuevo
			if (listado_antiguo <> objRstEnf("listado")) then

				' Si el listado antiguo no es vac�o, es que ya habiamos abierto antes uno as� que primero cerramos el anterior
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
				<table cellpadding=0 cellspacing=0 width="100%" border="0"><tr><td width="100%" class="titulo3" align="left"><a href="index.asp?idpagina=617"><img src="http://www.istas.net/ecoinformas/web/imagenes/ayuda.gif" align="absmiddle" border="0" /></a> <%=objRstEnf("listado")%>  <% plegador "secc-enf"&objRstEnf("listado"), "img-enf"&objRstEnf("listado") %></td></tr></table>
			</td>
		</tr>
		<tr id="secc-enf<%= aplana(objRstEnf("listado")) %>" style="display:none">
			<td>
<%
				listado_antiguo = objRstEnf("listado")
			end if
%>

				<!-- Tabla enfermedad -->				
				<table>
					<tr>
						<td class="subtitulo3" colspan="2"><%=objRstEnf("nombre")%></td>
					</tr>
				<%
					if (objRstEnf("sintomas") <> "") then
				%>
					<tr>
						<td class="subtitulo3" align="right" valign="top">S�ntomas:</td><td><%=replace(objRstEnf("sintomas"), vbcrlf, "<br>")%></td>
					</tr>
				<%
					end if
				%>
				<%
					if (objRstEnf("actividades") <> "") then
				%>
					<tr>
						<td class="subtitulo3" align="right" valign="top">Actividades:</td><td><%=replace(objRstEnf("actividades"), vbcrlf, "<br>")%></td>
					</tr>
				<%
					end if
				%>
				</table>
				<!-- Fin tabla enfermedad -->

<%
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
	if esta_en_lista_cov or esta_en_lista_residuos or esta_en_lista_vertidos or esta_en_lista_lpcic or esta_en_lista_accidentes or esta_en_lista_emisiones then
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
' Para dividir los 7 posibles apartados en dos columnas, primero calculamos cu�ntos hay en total.
total = 0

if esta_en_lista_cov then total = total +1 end if
if esta_en_lista_vertidos then total = total +1 end if
if esta_en_lista_lpcic_agua then total = total +1 end if
if esta_en_lista_lpcic_aire then total = total +1 end if
if esta_en_lista_residuos then total = total +1 end if
if esta_en_lista_accidentes then total = total +1 end if
if esta_en_lista_emisiones then total = total +1 end if

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
' Contaremos cuantos llevamos para ver en qu� momento hay que poner la divisi�n de columnas
llevo = 0
%>

<%
		if esta_en_lista_cov then
			ap3_riesgos_tabla("COV")
			llevo = llevo +1
			if llevo >= mitad then
				response.write "</td><td valign='top' width='50%'>"
				llevo = 0 ' Lo reseteo para que no vuelva a dividir
			end if
		end if

		if esta_en_lista_vertidos then
			ap3_riesgos_tabla("Vertidos")
			llevo = llevo +1
			if llevo >= mitad then
				response.write "</td><td valign='top' width='50%'>"
				llevo = 0 ' Lo reseteo para que no vuelva a dividir
			end if
		end if

		if esta_en_lista_lpcic_agua then
			ap3_riesgos_tabla("LPCIC (EPER Agua)")
			llevo = llevo +1
			if llevo >= mitad then
				response.write "</td><td valign='top' width='50%'>"
				llevo = 0 ' Lo reseteo para que no vuelva a dividir
			end if		
		end if

		if esta_en_lista_lpcic_aire then
			ap3_riesgos_tabla("LPCIC (EPER Aire)")
			llevo = llevo +1
			if llevo >= mitad then
				response.write "</td><td valign='top' width='50%'>"
				llevo = 0 ' Lo reseteo para que no vuelva a dividir
			end if
		end if

		if esta_en_lista_residuos then
			ap3_riesgos_tabla("Residuos Peligrosos")
			llevo = llevo +1
			if llevo >= mitad then
				response.write "</td><td valign='top' width='50%'>"
				llevo = 0 ' Lo reseteo para que no vuelva a dividir
			end if
		end if

		if esta_en_lista_accidentes then 
			ap3_riesgos_tabla("Accidentes Graves")
			llevo = llevo +1
			if llevo >= mitad then
				response.write "</td><td valign='top' width='50%'>"
				llevo = 0 ' Lo reseteo para que no vuelva a dividir
			end if
		end if

		if esta_en_lista_emisiones then
			ap3_riesgos_tabla("Emisiones Atmosf�ricas") 
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

' ##################################################################################
sub ap5_alternativas()
'	sql_enf = "select distinct enf.id, enf.listado, enf.nombre, enf.sintomas, enf.actividades FROM dn_risc_enfermedades AS enf LEFT OUTER JOIN dn_risc_grupos_por_enfermedades AS gpe ON enf.id = gpe.id_enfermedad LEFT OUTER JOIN dn_risc_sustancias_por_grupos AS spg ON gpe.id_grupo = spg.id_grupo LEFT OUTER JOIN dn_risc_sustancias_por_enfermedades AS spe ON spe.id_enfermedad = enf.id WHERE spg.id_sustancia="&id_sustancia&" OR spe.id_sustancia="&id_sustancia&" ORDER BY enf.listado, enf.nombre"

'	sql="SELECT DISTINCT id_fichero, titulo FROM dn_alter_ficheros_por_sustancias INNER JOIN dn_alter_ficheros ON dn_alter_ficheros_por_sustancias.id_fichero = dn_alter_ficheros.id WHERE id_sustancia="&id_sustancia&" ORDER BY titulo"

	sql="SELECT DISTINCT f.id AS id_fichero, f.titulo FROM dn_alter_ficheros AS f LEFT OUTER JOIN dn_alter_ficheros_por_sustancias AS fps ON f.id = fps.id_fichero LEFT OUTER JOIN dn_alter_ficheros_por_grupos AS fpg ON f.id = fpg.id_fichero LEFT OUTER JOIN dn_risc_grupos AS g ON fpg.id_grupo = g.id LEFT OUTER JOIN dn_risc_sustancias_por_grupos AS spg ON g.id = spg.id_grupo WHERE fps.id_sustancia="&id_sustancia&" OR spg.id_sustancia = "& id_sustancia&" ORDER BY titulo"

  'response.write sql

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
	' Mostramos los ficheros, comprobando que no haya titulos repetidos. Como vienen ordenados por t�tulo, basta comparar con el t�tulo anterior
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

	sql="SELECT DISTINCT s.numero_cnae AS codigo, s.nombre AS nombre, s.id AS id_sector FROM dn_alter_sectores AS s LEFT OUTER JOIN dn_risc_sustancias_por_sectores AS sps ON s.id = sps.id_sector WHERE sps.id_sustancia="&id_sustancia&" ORDER BY s.codigo"

  ' Mejora: incluimos solo los sectores que contienen documentos asociados
  'sql="SELECT DISTINCT s.numero_cnae AS codigo, s.nombre AS nombre, s.id AS id_sector FROM dn_alter_sectores AS s LEFT OUTER JOIN dn_risc_sustancias_por_sectores AS sps ON s.id = sps.id_sector INNER JOIN dn_alter_ficheros_por_sectores AS fps ON sps.id_sector = fps.id_sector WHERE sps.id_sustancia="&id_sustancia&" ORDER BY s.codigo"

  'response.write sql

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
	' Mostramos los sectores, comprobando que no haya codigos repetidos. Como vienen ordenados por c�digo, basta comparar con el c�digo anterior
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

	sql="SELECT cardiocirculatorio, rinyon, respiratorio, reproductivo, piel_sentidos, neuro_toxicos, musculo_esqueletico, sistema_inmunitario, higado_gastrointestinal, sistema_endocrino, embrion, cancer FROM dn_risc_sustancias_salud WHERE id_sustancia="&id_sustancia&" AND (cardiocirculatorio=1 OR rinyon=1 OR respiratorio=1 OR reproductivo=1 OR piel_sentidos=1 OR neuro_toxicos=1 OR musculo_esqueletico=1 OR sistema_inmunitario=1 OR higado_gastrointestinal=1 OR sistema_endocrino=1 OR embrion=1 OR cancer=1)"

  'response.write sql

	set objRst=objConnection2.execute(sql)
	if (not objRst.eof) then
%>
	<!-- Efectos para la salud -->
	<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
   	<tr>
			<td class="celdaabajo" colspan="2" align="center">
				<table cellpadding=0 cellspacing=0 width="100%" border="0"><tr><td width="100%" class="titulo3" align="left">Otras alteraciones para la salud y sistemas y �rganos afectados <% plegador "secc-salud", "img-salud" %></td></tr></table>
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

    if (cardiocirculatorio OR respiratorio OR reproductivo OR musculo_esqueletico OR sistema_inmunitario OR higado_gastrointestinal OR sistema_endocrino) then
%>
        <td valign="top">
        <strong>- Sistemas a los que afecta:</strong><br/>
        <ul>
<%
          if (cardiocirculatorio) then response.write "<li>Cardiocirculatorio</li>" end if
          if (respiratorio) then response.write "<li>Respiratorio</li>" end if
          if (reproductivo) then response.write "<li>Reproductivo</li>" end if
          if (musculo_esqueletico) then response.write "<li>Musculoesquel�tico</li>" end if
          if (sistema_inmunitario) then response.write "<li>Inmunitario</li>" end if
          if (higado_gastrointestinal) then response.write "<li>Gastrointestinal - H�gado</li>" end if
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
          if (embrion) then response.write "<li>Da�os en el embri�n</li>" end if
          if (cancer) then response.write "<li>C�ncer</li>" end if
          if (rinyon) then response.write "<li>Da�os en el ri��n</li>" end if
          if (piel_sentidos) then response.write "<li>Piel y mucosas</li>" end if
          if (neuro_toxicos) then response.write "<li>Efectos neurot�xicos</li>" end if
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

function dame_nivel_cancerigeno_rd()
	' Juntamos todas las clasificaciones
	clasificacion_rd = clasificacion_1 & clasificacion_2 & clasificacion_3 & clasificacion_4 & clasificacion_5 & clasificacion_6 & clasificacion_7 & clasificacion_8 & clasificacion_9 & clasificacion_10 & clasificacion_11 & clasificacion_12 & clasificacion_13 & clasificacion_14 & clasificacion_15

	' Sustituimos "Carc. Cat." por "Carc.Cat." para unificar
	clasificacion_rd = replace(clasificacion_rd, "Carc. Cat.", "Carc.Cat.")

	' Quitamos los espacios en blanco
	clasificacion_rd = replace(clasificacion_rd, " ", "")

	' Buscamos la primera aparicion de "Carc.Cat."
	posicion = instr(1,clasificacion_rd, "Carc.Cat.")

	' Sacamos el nivel como el caracter que hay justo detr�s de la primera aparici�n de la subcadena

	if (posicion > 0) then
		dame_nivel_cancerigeno_rd = mid(clasificacion_rd, posicion+9, 1)
	else
		dame_nivel_cancerigeno_rd = ""
	end if
end function

' #############################################################################################

function dame_nivel_cancerigeno_danesa()
	' Buscamos la primera aparicion de "Carc"
	posicion = instr(1,frases_r_danesa, "Carc")

	' Sacamos el nivel como el caracter que hay justo detr�s de la primera aparici�n de la subcadena
	if (posicion > 0) then
		dame_nivel_cancerigeno_danesa = mid(frases_r_danesa, posicion+4, 1)
	else
		dame_nivel_cancerigeno_danesa = ""
	end if
end function

' #############################################################################################

function dame_nivel_mutageno_rd()
	' Juntamos todas las clasificaciones
	clasificacion_rd = clasificacion_1 & clasificacion_2 & clasificacion_3 & clasificacion_4 & clasificacion_5 & clasificacion_6 & clasificacion_7 & clasificacion_8 & clasificacion_9 & clasificacion_10 & clasificacion_11 & clasificacion_12 & clasificacion_13 & clasificacion_14 & clasificacion_15

	' Sustituimos "Muta. Cat." por "Muta.Cat." para unificar
	clasificacion_rd = replace(clasificacion_rd, "Muta. Cat.", "Muta.Cat.")

	' Quitamos los espacios en blanco
	clasificacion_rd = replace(clasificacion_rd, " ", "")

	'response.write "["&clasificacion_rd&"]"

	' Buscamos la primera aparicion de "Muta.Cat."
	posicion = instr(1,clasificacion_rd, "Muta.Cat.")

	' Sacamos el nivel como el caracter que hay justo detr�s de la primera aparici�n de la subcadena
	if (posicion > 0) then
		dame_nivel_mutageno_rd = mid(clasificacion_rd, posicion+9, 1)
	else
		dame_nivel_mutageno_rd = ""
	end if
end function

' #############################################################################################

function dame_nivel_mutageno_danesa()
	' Buscamos la primera aparicion de "Mut"
	posicion = instr(1,frases_r_danesa, "Mut")

	' Sacamos el nivel como el caracter que hay justo detr�s de la primera aparici�n de la subcadena
	if (posicion > 0) then
		dame_nivel_mutageno_danesa = mid(frases_r_danesa, posicion+3, 1)
	else
		dame_nivel_mutageno_danesa = ""
	end if
end function

' #############################################################################################

function dame_nivel_reproduccion_rd()
	' Juntamos todas las clasificaciones
	clasificacion_rd = clasificacion_1 & clasificacion_2 & clasificacion_3 & clasificacion_4 & clasificacion_5 & clasificacion_6 & clasificacion_7 & clasificacion_8 & clasificacion_9 & clasificacion_10 & clasificacion_11 & clasificacion_12 & clasificacion_13 & clasificacion_14 & clasificacion_15

	' Sustituimos "Repr. Cat." por "Repr.Cat." para unificar
	clasificacion_rd = replace(clasificacion_rd, "Repr. Cat.", "Repr.Cat.")

	' Quitamos los espacios en blanco
	clasificacion_rd = replace(clasificacion_rd, " ", "")

	'response.write "["&clasificacion_rd&"]"

	' Buscamos la primera aparicion de "Repr.Cat."
	posicion = instr(1,clasificacion_rd, "Repr.Cat.")

	' Sacamos el nivel como el caracter que hay justo detr�s de la primera aparici�n de la subcadena
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

	' Sacamos el nivel como el caracter que hay justo detr�s de la primera aparici�n de la subcadena
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
  <a href="javascript:toggle('<%= id_bloque %>', '<%= id_imagen %>');"><img src="imagenes/desplegar.gif" align="absmiddle" id="<%= id_imagen %>" alt="Pulse para desplegar la informaci�n" title="Pulse para desplegar la informaci�n" /></a>
<%
end sub

' #############################################################################################

sub plegador_texto(byval id_bloque, byval texto, byval clase)
  ' Pinta el HTML necesario para las llamadas a mostrar/ocultar el objeto
  ' Solo se emplea para el plegador de frases R danesas, en caso de que no se hayan mostrado ya.
  id_bloque=aplana(id_bloque)
  if (frases_r_danesa_mostradas) then
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
