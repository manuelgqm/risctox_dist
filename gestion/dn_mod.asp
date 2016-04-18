<!--#include file="dn_fun_comunes.asp"-->

<!--#include file="dn_fun_texto.asp"-->
<!--#include file="dn_auten.inc"-->
<!--#include file="adovbs.inc"-->
<!--#include file="dn_conexion.asp"-->


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Istas</title>
<link rel="stylesheet" type="text/css" href="dn_estilos.css">
<script type="text/javascript" src="niftycube.js"></script>
<script type="text/javascript">
window.onload=function(){
Nifty("div#box2","big");
}



function valida_longitud(campo, max)

{

  if (campo.value.length > max)

  {

    campo.value = campo.value.substr(0,max);

  }

}
</script>
<script language="JavaScript" src="gen_validatorv2.js" type="text/javascript"></script>
</head>

<body>
<%flashMsgShow()%>
<div id="box2" class="centcontenido">

<%
asociar=request("asociar")
id=request("id")
if id="" then
	cerrarconexion
%>
<script>window.close();</script>
<%
else
	select case asociar

		case "grupo":
			sql3="select * from dn_risc_grupos where id=" &id
			set objRst3=objconn1.execute(sql3)
			nombre=objRst3("nombre")
      		descripcion=objRst3("descripcion")
			nombre_ing=objRst3("nombre_ing")
      		descripcion_ing=objRst3("descripcion_ing")

			num_cas=objRst3("num_cas")


			call evaluaCamposListaAsociada("rd",split("",","))
			call evaluaCamposListaAsociada("enfermedades",split("",","))


			call evaluaCamposListaAsociada("eper_agua",split("",","))
			call evaluaCamposListaAsociada("eper_aire",split("",","))
			call evaluaCamposListaAsociada("eper_suelo",split("",","))

			call evaluaCamposListaAsociada("directiva_aguas",split("clasif_mma",","))

			call evaluaCamposListaAsociada("emisiones_atmosfericas",split("",","))
			call evaluaCamposListaAsociada("calidad_aire",split("",","))


			call evaluaCamposListaAsociada("vla",split("estado_1,ed_ppm_1,ed_mg_m3_1,ec_ppm_1,ec_mg_m3_1,notas_vla_1",","))
			call evaluaCamposListaAsociada("vla",split("estado_2,ed_ppm_2,ed_mg_m3_2,ec_ppm_2,ec_mg_m3_2,notas_vla_2",","))
			call evaluaCamposListaAsociada("vla",split("estado_3,ed_ppm_3,ed_mg_m3_3,ec_ppm_3,ec_mg_m3_3,notas_vla_3",","))
			call evaluaCamposListaAsociada("vla",split("estado_4,ed_ppm_4,ed_mg_m3_4,ec_ppm_4,ec_mg_m3_4,notas_vla_4",","))
			call evaluaCamposListaAsociada("vla",split("estado_5,ed_ppm_5,ed_mg_m3_5,ec_ppm_5,ec_mg_m3_5,notas_vla_5",","))
			call evaluaCamposListaAsociada("vla",split("estado_6,ed_ppm_6,ed_mg_m3_6,ec_ppm_6,ec_mg_m3_6,notas_vla_6",","))

			call evaluaCamposListaAsociada("vlb",split("ib_1,vlb_1,momento_1,notas_vlb_1",","))
			call evaluaCamposListaAsociada("vlb",split("ib_2,vlb_2,momento_2,notas_vlb_2",","))
			call evaluaCamposListaAsociada("vlb",split("ib_3,vlb_3,momento_3,notas_vlb_3",","))
			call evaluaCamposListaAsociada("vlb",split("ib_4,vlb_4,momento_4,notas_vlb_4",","))
			call evaluaCamposListaAsociada("vlb",split("ib_5,vlb_5,momento_5,notas_vlb_5",","))
			call evaluaCamposListaAsociada("vlb",split("ib_6,vlb_6,momento_6,notas_vlb_6",","))





			call evaluaCamposListaAsociada("cancer_rd",split("notas_cancer_rd",","))
			call evaluaCamposListaAsociada("cancer_iarc",split("grupo_iarc,volumen_iarc",","))

			call evaluaCamposListaAsociada("cancer_otras",split("categoria_cancer_otras,fuente",","))
			call evaluaCamposListaAsociada("cancer_mama",split("cancer_mama_fuente",","))
			call evaluaCamposListaAsociada("neuro_oto",split("efecto_neurotoxico,nivel_neurotoxico,fuente_neurotoxico",","))
			call evaluaCamposListaAsociada("disruptores",split("nivel_disruptor",","))


			call evaluaCamposListaAsociada("seveso",split("",","))
			call evaluaCamposListaAsociada("cov",split("",","))
			call evaluaCamposListaAsociada("reproduccion",split("",","))




			'-- xip

			asoc_peligrosas_agua_alemania = objRst3("asoc_peligrosas_agua_alemania")
			asoc_capa_ozono = objRst3("asoc_capa_ozono")
			asoc_cambio_climatico = objRst3("asoc_cambio_climatico")
			asoc_contaminantes_suelo = objRst3("asoc_contaminantes_suelo")

			'-- /xip

			'-- spl
			call evaluaCamposListaAsociada("alergenos",split("",","))

			call evaluaCamposListaAsociada("cop",split("enlace_cop",","))
			call evaluaCamposListaAsociada("mpmb",split("",","))
			call evaluaCamposListaAsociada("tpb",split("enlace_tpb,anchor_tpb,fuentes_tpb",","))

			call evaluaCamposListaAsociada("prohibidas",split("comentario_prohibida,comentario_prohibida_ing",","))
			call evaluaCamposListaAsociada("restringidas",split("comentario_restringida,comentario_restringida_ing",","))

			call evaluaCamposListaAsociada("prohibidas_embarazadas",split("comentario_prohibida,comentario_prohibida_ing",","))
			call evaluaCamposListaAsociada("prohibidas_lactantes",split("comentario_prohibida,comentario_prohibida_ing",","))
			call evaluaCamposListaAsociada("candidatas_reach",split("",","))
			call evaluaCamposListaAsociada("autorizacion_reach",split("",","))

			call evaluaCamposListaAsociada("biocidas_autorizadas",split("fuente,pureza_minima,condiciones,usos,condiciones_ing,usos_ing",","))
			call evaluaCamposListaAsociada("biocidas_prohibidas",split("fuente,fecha_limite,usos,usos_ing",","))

			call evaluaCamposListaAsociada("pesticidas_autorizadas",split("fuente,plazo_renovacion,pureza_minima,usos,plazo_renovacion_ing,pureza_minima_ing,usos_ing",","))
			call evaluaCamposListaAsociada("pesticidas_prohibidas",split("fuente,exenciones,exenciones_ing",","))
			'-- /spl


			objRst3.close
			set objRst3=nothing

		case "enfermedad":
			sql3="select * from dn_risc_enfermedades where id=" &id
			set objRst3=objconn1.execute(sql3)
			nombre=objRst3("nombre")
			listado=objRst3("listado")
			sintomas=objRst3("sintomas")
			actividades=objRst3("actividades")
			nombre_ing=objRst3("nombre_ing")
			listado_ing=objRst3("listado_ing")
			sintomas_ing=objRst3("sintomas_ing")
			actividades_ing=objRst3("actividades_ing")
			objRst3.close
			set objRst3=nothing

		case "uso":
			sql3="select nombre, descripcion, nombre_ing, descripcion_ing  from dn_risc_usos where id=" &id
			set objRst3=objconn1.execute(sql3)
			nombre=h2(objRst3("nombre"))
			descripcion=h2(objRst3("descripcion"))
			nombre_ing=h2(objRst3("nombre_ing"))
			descripcion_ing=h2(objRst3("descripcion_ing"))

			objRst3.close
			set objRst3=nothing

		case "compania":
			sql3="select * from dn_risc_companias where id=" &id
			set objRst3=objconn1.execute(sql3)
			nombre=objRst3("nombre")
			direccion=objRst3("direccion")
			fuente=objRst3("fuente")
			productora=objRst3("productora")
			distribuidora=objRst3("distribuidora")
			objRst3.close
			set objRst3=nothing

		case "sector":
			sql3="select * from dn_alter_sectores where id=" &id
			set objRst3=objconn1.execute(sql3)
			nombre=objRst3("nombre")
			numero_cnae=objRst3("numero_cnae")
			objRst3.close
			set objRst3=nothing

		case "proceso":
			sql3="select * from dn_alter_procesos where id=" &id
			set objRst3=objconn1.execute(sql3)
			nombre=objRst3("nombre")
			descripcion=objRst3("descripcion")
			objRst3.close
			set objRst3=nothing

		case "residuo":
			sql3="select * from rq_residuos where id=" &id
			set objRst3=objconn1.execute(sql3)
			codigo=objRst3("codigo")
			nombre=objRst3("nombre")
			objRst3.close
			set objRst3=nothing
	end select
end if
cerrarconexion
%>

<form name="myform" action="dn_mod2.asp?asociar=<%=asociar%>&id=<%=id%>" method="post" >

<fieldset>
<legend><strong>Modificar</strong></legend>

<%
select case asociar
	case "grupo":
%>
		<table border="0">
			<tr>
				<td align="left"><strong>Nombre</strong>:</td>
				<td align="left"><input name="nombre" type="text" value="<%=nombre%>" size="80" maxlength="750" /></td>
			</tr>
			<tr>
				<td align="left"><strong>Nombre en ingl&eacute;s</strong>:</td>
				<td align="left"><input name="nombre_ing" type="text" value="<%=nombre_ing%>" size="80" maxlength="750" /></td>
			</tr>
			<tr>
				<td align="left"><strong>Descripción</strong>:</td>
				<td align="left"><textarea name="descripcion" rows="10" cols="80"><%=descripcion%></textarea></td>
			</tr>
			<tr>
				<td align="left"><strong>Descripci&oacute;n en ingl&eacute;s</strong>:</td>
				<td align="left"><textarea name="descripcion_ing" rows="10" cols="80"><%=descripcion_ing%></textarea></td>
			</tr>

			<tr>
				<td align="left"><strong>Núm. CAS</strong>:</td>
				<td align="left"><input name="num_cas" type="text" value="<%=num_cas%>" size="20" maxlength="20" /></td>
			</tr>
		</table>

		<fieldset>
		<legend><strong>Asociado a las listas:</strong></legend>

			<table border="0">
				<tr>
					<td colspan="6" align="left"><br></td>
				</tr>
				<tr>
					<td colspan="6" align="left"><strong>Listas no clasificadas (no aparecen en inicio)</strong></td>
				</tr>
				<tr>
					<td align="left" valign="top" colspan="4">
						<% call generaCamposListaAsociada("Real Decreto 363", "rd", split("",","),split("",","))%>

			        </td>
				</tr>





				<tr>
					<td colspan="6" align="left"><br></td>
				</tr>
				<tr>
					<td colspan="6" align="left"><strong>Riesgos específicos para la salud</strong></td>
				</tr>
				<tr>
					<td align="left" valign="top" colspan="4">
						<% call generaCamposListaAsociada("Cancerígenos y mutágenos según RD 363/1995", "cancer_rd", split("Notas",","),split("notas_cancer_rd",","))%>
						<% call generaCamposListaAsociada("Cancerígenos y mutágenos según IARC", "cancer_iarc", split("Grupo,Volumen",","),split("grupo_iarc,volumen_iarc",","))%>
						<% call generaCamposListaAsociada("Cancerígenos y mutágenos según Otras Fuentes", "cancer_otras", split("Categoría,Fuente",","),split("categoria_cancer_otras,fuente",","))%>
						<% call generaCamposListaAsociada("Cancerígenos y mutágenos según SSI (cáncer de mama)", "cancer_mama", split("Fuente",","),split("cancer_mama_fuente",","))%>




					<br><br><br><br>
						<% call generaCamposListaAsociada("Tóxicos para la reproducción", "reproduccion", split("",","),split("",","))%>
						<% call generaCamposListaAsociada("Disruptores Endocrinos", "disruptores", split("Fuente",","),split("nivel_disruptor",","))%>
						<% call generaCamposListaAsociada("Neurotóxicos / Ototóxicos", "neuro_oto", split("Efecto,Nivel,Fuente",","),split("efecto_neurotoxico,nivel_neurotoxico,fuente_neurotoxico",","))%>
						<% call generaCamposListaAsociada("Alérgenos REACH", "alergenos", split("",","),split("",","))%>
			        </td>
				</tr>


				<tr>
					<td colspan="6" align="left"><br></td>
				</tr>
				<tr>
					<td colspan="6" align="left"><strong>Riesgos específicos medioambiente</strong></td>
				</tr>
				<tr>
					<td align="left" valign="top" colspan="4">
						<% call generaCamposListaAsociada("TPB (Tóxicas, persistentes y bioacumulativas)", "tpb", split("Enlace,Nombre sustancia,Fuentes",","),split("enlace_tpb,anchor_tpb,fuentes_tpb",","))%>
						<% call generaCamposListaAsociada("mPmB", "mpmb", split("",","),split("",","))%>
						<% call generaCamposListaAsociada("Toxicidad acuática: Directiva de Aguas", "directiva_aguas", split("Clasificación MMA",","),split("clasif_mma",","))%>
						<% call generaCamposListaAsociada("Toxicidad acuática: Peligrosas agua Alemania", "peligrosas_agua_alemania", split("",","),split("",","))%>

						<% call generaCamposListaAsociada("Daño a la atmósfera: Capa Ozono", "capa_ozono", split("",","),split("",","))%>
						<% call generaCamposListaAsociada("Daño a la atmósfera: Cambio climático", "cambio_climatico", split("",","),split("",","))%>
						<% call generaCamposListaAsociada("Daño a la atmósfera: Calidad del Aire", "calidad_aire", split("",","),split("",","))%>

						<% call generaCamposListaAsociada("Contaminantes suelo (Según RD 9/2005)", "contaminantes_suelo", split("",","),split("",","))%>
						<% call generaCamposListaAsociada("Contaminantes Orgánicos Persistentes (COP's)", "cop", split("Enlace",","),split("enlace_cop",","))%>
			        </td>
				</tr>


				<tr>
					<td colspan="6" align="left"><br></td>
				</tr>
				<tr>
					<td colspan="6" align="left"><strong>Normativa sobre salud laboral</strong></td>
				</tr>
				<tr>
					<td align="left" valign="top" colspan="4">
						<% call generaCamposListaAsociadaPorTipo("vl1", "Valores Límite Ambientales", "vla", split("ESTADO,VLA-ED (pmm),VLA-ED (mg/m3),VLA-EC (pmm),VLA-EC (mg/m3),NOTAS VLA",","),split("estado_1,ed_ppm_1,ed_mg_m3_1,ec_ppm_1,ec_mg_m3_1,notas_vla_1",","))%>
						<% call generaCamposListaAsociadaPorTipo("vl2", "Valores Límite Ambientales", "vla", split("ESTADO,VLA-ED (pmm),VLA-ED (mg/m3),VLA-EC (pmm),VLA-EC (mg/m3),NOTAS VLA",","),split("estado_2,ed_ppm_2,ed_mg_m3_2,ec_ppm_2,ec_mg_m3_2,notas_vla_2",","))%>
						<% call generaCamposListaAsociadaPorTipo("vl3", "Valores Límite Ambientales", "vla", split("ESTADO,VLA-ED (pmm),VLA-ED (mg/m3),VLA-EC (pmm),VLA-EC (mg/m3),NOTAS VLA",","),split("estado_3,ed_ppm_3,ed_mg_m3_3,ec_ppm_3,ec_mg_m3_3,notas_vla_3",","))%>
						<% call generaCamposListaAsociadaPorTipo("vl4", "Valores Límite Ambientales", "vla", split("ESTADO,VLA-ED (pmm),VLA-ED (mg/m3),VLA-EC (pmm),VLA-EC (mg/m3),NOTAS VLA",","),split("estado_4,ed_ppm_4,ed_mg_m3_4,ec_ppm_4,ec_mg_m3_4,notas_vla_4",","))%>
						<% call generaCamposListaAsociadaPorTipo("vl5", "Valores Límite Ambientales", "vla", split("ESTADO,VLA-ED (pmm),VLA-ED (mg/m3),VLA-EC (pmm),VLA-EC (mg/m3),NOTAS VLA",","),split("estado_5,ed_ppm_5,ed_mg_m3_5,ec_ppm_5,ec_mg_m3_5,notas_vla_5",","))%>
						<% call generaCamposListaAsociadaPorTipo("vl6", "Valores Límite Ambientales", "vla", split("ESTADO,VLA-ED (pmm),VLA-ED (mg/m3),VLA-EC (pmm),VLA-EC (mg/m3),NOTAS VLA",","),split("estado_6,ed_ppm_6,ed_mg_m3_6,ec_ppm_6,ec_mg_m3_6,notas_vla_6",","))%>

						<% call generaCamposListaAsociadaPorTipo("vl1", "Valores Límite Biológicos", "vlb", split("INDICADOR BIOLÓGICO,VLB<br>,MOMENTO DE MUESTREO,NOTAS VLB<br>",","),split("ib_1,vlb_1,momento_1,notas_vlb_1",","))%>
						<% call generaCamposListaAsociadaPorTipo("vl2", "Valores Límite Biológicos", "vlb", split("INDICADOR BIOLÓGICO,VLB,MOMENTO DE MUESTREO,NOTAS VLB",","),split("ib_2,vlb_2,momento_2,notas_vlb_2",","))%>
						<% call generaCamposListaAsociadaPorTipo("vl3", "Valores Límite Biológicos", "vlb", split("INDICADOR BIOLÓGICO,VLB,MOMENTO DE MUESTREO,NOTAS VLB",","),split("ib_3,vlb_3,momento_3,notas_vlb_3",","))%>
						<% call generaCamposListaAsociadaPorTipo("vl4", "Valores Límite Biológicos", "vlb", split("INDICADOR BIOLÓGICO,VLB,MOMENTO DE MUESTREO,NOTAS VLB",","),split("ib_4,vlb_4,momento_4,notas_vlb_4",","))%>
						<% call generaCamposListaAsociadaPorTipo("vl5", "Valores Límite Biológicos", "vlb", split("INDICADOR BIOLÓGICO,VLB,MOMENTO DE MUESTREO,NOTAS VLB",","),split("ib_5,vlb_5,momento_5,notas_vlb_5",","))%>
						<% call generaCamposListaAsociadaPorTipo("vl6", "Valores Límite Biológicos", "vlb", split("INDICADOR BIOLÓGICO,VLB,MOMENTO DE MUESTREO,NOTAS VLB",","),split("ib_6,vlb_6,momento_6,notas_vlb_6",","))%>

						<% call generaCamposListaAsociada("Enfermedades Profesionales", "enfermedades", split("",","),split("",","))%>
			        </td>
				</tr>



				<tr>
					<td colspan="6" align="left"><br></td>
				</tr>
				<tr>
					<td colspan="6" align="left"><strong>Normativa ambiental</strong></td>
				</tr>
				<tr>
					<td align="left" valign="top" colspan="4">
						<% call generaCamposListaAsociada("Emisiones Atmosféricas", "emisiones_atmosfericas", split("",","),split("",","))%>
						<% call generaCamposListaAsociada("COV", "cov", split("",","),split("",","))%>
						<% call generaCamposListaAsociada("IPPC (PRTR agua)", "eper_agua", split("",","),split("",","))%>
						<% call generaCamposListaAsociada("IPPC (PRTR aire)", "eper_aire", split("",","),split("",","))%>
						<% call generaCamposListaAsociada("IPPC (PRTR suelo)", "eper_suelo", split("",","),split("",","))%>
						<% call generaCamposListaAsociada("Accidentes graves (Seveso)", "seveso", split("",","),split("",","))%>
			        </td>
				</tr>



				<tr>
					<td colspan="6" align="left"><br></td>
				</tr>
				<tr>
					<td colspan="6" align="left"><strong>Normativa sobre restricción / prohibición de sustancias</strong></td>
				</tr>
				<tr>
					<td align="left" valign="top" colspan="4">

	<% call generaCamposListaAsociada("Prohibidas", "prohibidas", split("Comentario,Comentario (ingl&eacute;s)",","),split("comentario_prohibida,comentario_prohibida_ing",","))%>
	<% call generaCamposListaAsociada("Restringidas", "restringidas", split("Comentario,Comentario (ingl&eacute;s)",","),split("comentario_restringida,comentario_restringida_ing",","))%>

	<% call generaCamposListaAsociada("Prohibidas embarazadas", "prohibidas_embarazadas", split("Comentario,Comentario (ingl&eacute;s)",","),split("comentario_prohibida,comentario_prohibida_ing",","))%>
	<% call generaCamposListaAsociada("Prohibidas lactantes", "prohibidas_lactantes", split("Comentario,Comentario (ingl&eacute;s)",","),split("comentario_prohibida,comentario_prohibida_ing",","))%>
	<% call generaCamposListaAsociada("Candidatas REACH", "candidatas_reach", split("",","),split("",","))%>
	<% call generaCamposListaAsociada("Sujetas a autorización REACH", "autorizacion_reach", split("",","),split("",","))%>

	<% call generaCamposListaAsociada("Biocidas autorizadas", "biocidas_autorizadas", split("Fuente,Pureza m&iacute;nima,Condiciones,Usos,Condiciones (ingl&eacute;s),Usos (ingl&eacute;s)",","),split("fuente,pureza_minima,condiciones,usos,condiciones_ing,usos_ing",","))%>
	<% call generaCamposListaAsociada("Biocidas prohibidas", "biocidas_prohibidas", split("Fuente,Fecha l&iacute;mite,Usos,Usos (ingl&eacute;s)",","),split("fuente,fecha_limite,usos,usos_ing",","))%>

	<% call generaCamposListaAsociada("Pesticidas autorizadas", "pesticidas_autorizadas", split("Fuente,Plazo renovaci&oacute;n,Pureza m&iacute;nima,Usos,Plazo renovaci&oacute;n (ingl&eacute;s),Pureza m&iacute;nima (ingl&eacute;s),Usos (ingl&eacute;s)",","),split("fuente,plazo_renovacion,pureza_minima,usos,plazo_renovacion_ing,pureza_minima_ing,usos_ing",","))%>
	<% call generaCamposListaAsociada("Pesticidas prohibidas", "pesticidas_prohibidas", split("Fuente,Exenciones,Exenciones (ingl&eacute;s)",","),split("fuente,exenciones,exenciones_ing",","))%>
			        </td>
				</tr>
			</table>
		</fieldset>

<%
	case "compania":
%>
		<table align="center">
		<tr><td>Nombre</td><td> <input type="text" name="nombre" maxlength="2500" size="50" value="<%=nombre%>" /></td></tr>
		<tr><td>Dirección</td><td><textarea name="direccion" cols="50" rows="5"><%=direccion%></textarea></td></tr>
		<tr><td>Fuente</td><td><textarea name="fuente" cols="50" rows="5"><%=fuente%></textarea></td></tr>
		<tr><td colspan="2"><input type="checkbox" name="productora" value="1" <%if productora then response.write "checked"%> /> productora &nbsp;&nbsp;&nbsp; <input type="checkbox" name="distribuidora" value="1" <%if distribuidora then response.write "checked"%> />  distribuidora</td></tr>
		</table>
<%
	case "enfermedad":
%>
		<table>
		<tr><td>Nombre </td><td><input type="text" name="nombre" maxlength="500" size="100" value="<%=nombre%>" /></td></tr>
		<tr><td>Listado </td><td><textarea name="listado" cols="70"><%=listado%></textarea></td></tr>
		<tr><td>Síntomas </td><td><textarea name="sintomas" cols="70" id="sintomas"><%=sintomas%></textarea></td></tr>
		<tr><td>Actividades </td><td><textarea name="actividades" cols="70" id="actividades"><%=actividades%></textarea></td></tr>
		</table>
		<fieldset>
			<legend>En inglés:</legend>
			<table>
				<tr><td>Nombre </td><td><input type="text" name="nombre_ing" maxlength="500" size="100" value="<%=nombre_ing%>" /></td></tr>
				<tr><td>Listado </td><td><textarea name="listado_ing" cols="70"><%=listado_ing%></textarea></td></tr>
				<tr><td>Síntomas </td><td><textarea name="sintomas_ing" cols="70" id="sintomas"><%=sintomas_ing%></textarea></td></tr>
				<tr><td>Actividades </td><td><textarea name="actividades_ing" cols="70" id="actividades"><%=actividades_ing%></textarea></td></tr>
			</table>
		</fieldset>
<%
	case "sector":
%>
		<table align="center">
		<tr><td>Nombre</td><td> <input type="text" name="nombre" maxlength="1500" size="50" value="<%=nombre%>" /> </td></tr>
		<tr><td>Nº CNAE</td><td> <input type="text" name="numero_cnae" maxlength="10" size="50" value="<%=numero_cnae%>" /> </td></tr>
		</table>
<%
	case "proceso":
%>
		<table align="center">
		<tr><td>Nombre</td><td> <input type="text" name="nombre" maxlength="150" size="50" value="<%=nombre%>" /> </td></tr>
		<tr><td>Descripcion</td><td> <textarea name="descripcion" cols="45"><%=descripcion%></textarea> </td></tr>
		</table>

<%
	case "uso":
%>
		<table align="center">
		<tr><td>Nombre</td><td> <input type="text" name="nombre" maxlength="150" size="50" value="<%=nombre%>" /> </td></tr>
		<tr><td>Descripci&oacute;n (m&aacute;ximo: 500 caracteres)</td><td> <textarea name="descripcion" cols="45" onchange="valida_longitud(this,500);"><%=descripcion%></textarea> </td></tr>
		</table>
		<fieldset>
			<legend>En inglés:</legend>
			<table>
				<tr><td>Nombre </td><td><input type="text" name="nombre_ing" maxlength="500" size="100" value="<%=nombre_ing%>" /></td></tr>
				<tr><td>Descripci&oacute;n (m&aacute;ximo: 500 caracteres)</td><td><input type="text" name="descripcion_ing" maxlength="500" size="100" value="<%=descripcion_ing%>" /></td></tr>
			</table>
		</fieldset>
<%
	case "residuo":
%>
		<table align="center">
		<tr><td>Código</td><td>  <input type="text" name="codigo" maxlength="10" size="10" value="<%=codigo%>" /> </td></tr>
		<tr><td>Residuo</td><td> <input type="text" name="nombre" maxlength="2000" size="100" value="<%=nombre%>" /></td></tr>
		</table>


<%
	case else:
%>
		Nuevo nombre: <input name="nombre" type="text" value="<%=nombre%>" size="150" maxlength="500" />
<%
end select
%>

</fieldset>

<p><input type="submit" value="Enviar" class="centcontenido"  /></p>
</form>
<script language="JavaScript" type="text/javascript">
var frmvalidator = new Validator("myform");
<%
	if asociar <> "enfermedad" then
%>
		frmvalidator.addValidation("nombre","req","Debe escribir un nombre.");
<%
	end if
%>
<%validacionesadicionales asociar%>
</script>
</div>
</body>
</html>

<%
sub validacionesadicionales(asociar)
	select case asociar
		case "enfermedad":
%>
frmvalidator.addValidation("listado","maxlen=300");
frmvalidator.addValidation("sintomas","maxlen=1000");
frmvalidator.addValidation("actividades","maxlen=2000");
<%
		'case else:
	end select
end sub

sub generaCamposListaAsociada(titulo, lista, nombresCamposArray(),camposArray())
	tipo = ""
	call generaCamposListaAsociadaPorTipo(tipo, titulo, lista, nombresCamposArray,camposArray)
end sub
sub generaCamposListaAsociadaPorTipo(tipo, titulo, lista, nombresCamposArray(),camposArray())
	dim c
%>
	<div style="clear:left;">
	<br>
<%
		' SI ES EL PRIMER CAMPO DE UNA LISTA DE VALORES LÍMITE O CUALQUIER OTRO CAMPO DIFERENTE DE VALORES LÍMITE:
		if left(tipo,3)="vl1" or left(tipo,2)<>"vl" then
%>
	<input type="checkbox" name="asoc_<%=lista%>" value="1" <%if eval("asoc_"&lista) then response.write "checked"%> /> <%=titulo%><br/>
<%
		end if
%>
	<div style="margin-left:30px;">
	<%
	if left(tipo,2)="vl" then
		linea=right(tipo,1)
	%>
		<div style="float:left;"><br><b><%=linea%></b>&nbsp;</div>
	<%
	end if

		for i = 0 to UBound(camposArray)
			c = camposArray(i)
			if c<>"" then
				if left(tipo,3)="vl1" then
	%>
		<div style="float:left;width:130px;"><%=nombresCamposArray(i)%>:<br><input type="text" size="15" name="asoc_<%=lista%>_<%=c%>" value="<%=eval("asoc_"&lista&"_"&c)%>"></div>
	<%
				else
					if left(tipo,2)="vl" then
	%>
		<div style="float:left;width:130px;"><input type="text" size="15" name="asoc_<%=lista%>_<%=c%>" value="<%=eval("asoc_"&lista&"_"&c)%>"></div>
	<%
					else
	%>
		<div style="float:left;"><%=nombresCamposArray(i)%>:<br><textarea name="asoc_<%=lista%>_<%=c%>"><%=eval("asoc_"&lista&"_"&c)%></textarea></div>
	<%
					end if
				end if
			end if
		next
	%>
	</div>
	</div>
<%
end sub



sub evaluaCamposListaAsociada(lista,camposArray())
	dim c
	dim v
	execute("asoc_"&lista&" = objRst3(""asoc_"&lista&""")")
	for i = 0 to UBound(camposArray)
		c = camposArray(i)
		v = objRst3("asoc_"&lista&"_"&c)
		if isnull(v) then v=""
		v = replace(v,vbCrLf,"<br>")
'response.write v
'response.write "<br>asoc_"&lista&"_"&c&" = """ & objRst3("asoc_"&lista&"_"&c) & """"
		execute("asoc_"&lista&"_"&c&" = """ & v & """")
	next
end sub




%>


