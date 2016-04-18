<!--#include file="dn_fun_comunes.asp"-->

<!--#include file="dn_fun_texto.asp"-->
<!--#include file="dn_auten.inc"-->
<!--#include file="adovbs.inc"-->
<!--#include file="dn_conexion.asp"-->

<%'Permite auditar las acciones%>
<!--#include file="spl_auditoria.inc.asp"-->

<%
asociar=request("asociar")
id=request("id")
nombre=h(request.form("nombre"))
descripcion=h(request.form("descripcion"))
if id="" then
	cerrarconexion
%>
<script>window.close();</script>
<%
else
	select case asociar

		case "grupo":
			nombre_ing=h(request.form("nombre_ing"))
			descripcion_ing=h(request.form("descripcion_ing"))

			num_cas=request.form("num_cas")
			asoc_rd =0
			asoc_enfermedades =0

			asoc_cop =0
			asoc_eper =0
			asoc_directiva_aguas =0
			asoc_emisiones_atmosfericas =0
			asoc_calidad_aire =0
			asoc_vla =0
			asoc_vlb =0
			asoc_cancer_iarc =0
			asoc_cancer_rd = 0

		    asoc_cancer_mama = 0
			asoc_cancer_otras = 0
			asoc_seveso = 0
			asoc_neuro_oto = 0
			asoc_reproduccion = 0
			asoc_disruptores = 0

			'-- xip
			asoc_tpb = 0
			asoc_mpmb = 0
			asoc_prohibidas = 0
			asoc_restringidas = 0
			asoc_alergenos = 0
			asoc_peligrosas_agua_alemania = 0
			asoc_capa_ozono = 0
			asoc_cambio_climatico = 0
			asoc_contaminantes_suelo = 0
			asoc_cov = 0
			'-- /xip

			'-- spl
			asoc_prohibidas_embarazadas=0
			asoc_prohibidas_lactantes=0
			asoc_candidatas_reach=0
			asoc_autorizacion_reach=0
			asoc_biocidas_autorizadas=0
			asoc_biocidas_prohibidas=0
			asoc_pesticidas_autorizadas=0
			'-- /spl

			if request.form("asoc_rd")=1 then asoc_rd=1

			if request.form("asoc_enfermedades")=1 then asoc_enfermedades=1

			if request.form("asoc_cop")=1 then asoc_cop=1
			if (request.form("asoc_eper_agua")=1) and (request.form("asoc_eper_aire")=1) and (request.form("asoc_eper_suelo")=1) then asoc_eper=1
			if request.form("asoc_directiva_aguas")=1 then asoc_directiva_aguas=1
			if request.form("asoc_emisiones_atmosfericas")=1 then asoc_emisiones_atmosfericas=1
			if request.form("asoc_calidad_aire")=1 then asoc_calidad_aire=1
			if request.form("asoc_vla")=1 then asoc_vla=1
			if request.form("asoc_vlb")=1 then asoc_vlb=1
			if request.form("asoc_cancer_iarc")=1 then asoc_cancer_iarc=1
			if request.form("asoc_cancer_rd")=1 then asoc_cancer_rd=1

			if request.form("asoc_cancer_mama")=1 then asoc_cancer_mama=1
			if request.form("asoc_cancer_otras")=1 then asoc_cancer_otras=1
			if request.form("asoc_seveso")=1 then asoc_seveso=1
			if request.form("asoc_neuro_oto")=1 then asoc_neuro_oto=1
			if request.form("asoc_reproduccion")=1 then asoc_reproduccion=1
			if request.form("asoc_disruptores")=1 then asoc_disruptores=1

			'-- xip
			if request.form("asoc_tpb")=1 then asoc_tpb=1
			if request.form("asoc_mpmb")=1 then asoc_mpmb=1
			if request.form("asoc_prohibidas")=1 then asoc_prohibidas=1
			if request.form("asoc_restringidas")=1 then asoc_restringidas=1
			if request.form("asoc_alergenos")=1 then asoc_alergenos=1
			if request.form("asoc_peligrosas_agua_alemania")=1 then asoc_peligrosas_agua_alemania=1
			if request.form("asoc_capa_ozono")=1 then asoc_capa_ozono=1
			if request.form("asoc_cambio_climatico")=1 then asoc_cambio_climatico=1
			if request.form("asoc_contaminantes_suelo")=1 then asoc_contaminantes_suelo=1
			if request.form("asoc_cov")=1 then asoc_cov=1
			'-- /xip


			'-- spl
			if request.form("asoc_prohibidas_embarazadas")=1 then asoc_prohibidas_embarazadas=1
			if request.form("asoc_prohibidas_lactantes")=1 then asoc_prohibidas_lactantes=1
			if request.form("asoc_candidatas_reach")=1 then asoc_candidatas_reach=1
			if request.form("asoc_autorizacion_reach")=1 then asoc_autorizacion_reach=1
			'-- /spl

			sql3="update dn_risc_grupos set nombre='" &nombre& "', descripcion='"&descripcion&"',nombre_ing='" &nombre_ing& "', descripcion_ing='"&descripcion_ing&"', num_cas='" &num_cas& "'"
			sql3=sql3& ", asoc_rd=" &asoc_rd& " , asoc_enfermedades=" &asoc_enfermedades& ", asoc_eper=" &asoc_eper& ", asoc_reproduccion=" &asoc_reproduccion


			'-- spl
			sql3 = sql3 & generaSQLListaAsociada( "alergenos", Split("",","))

			sql3 = sql3 & generaSQLListaAsociada( "emisiones_atmosfericas", Split("",","))
			sql3 = sql3 & generaSQLListaAsociada( "cov", Split("",","))
			sql3 = sql3 & generaSQLListaAsociada( "seveso", Split("",","))

			sql3 = sql3 & generaSQLListaAsociada("vla",split("estado_1,ed_ppm_1,ed_mg_m3_1,ec_ppm_1,ec_mg_m3_1,notas_vla_1,estado_2,ed_ppm_2,ed_mg_m3_2,ec_ppm_2,ec_mg_m3_2,notas_vla_2,estado_3,ed_ppm_3,ed_mg_m3_3,ec_ppm_3,ec_mg_m3_3,notas_vla_3,estado_4,ed_ppm_4,ed_mg_m3_4,ec_ppm_4,ec_mg_m3_4,notas_vla_4,estado_5,ed_ppm_5,ed_mg_m3_5,ec_ppm_5,ec_mg_m3_5,notas_vla_5,estado_6,ed_ppm_6,ed_mg_m3_6,ec_ppm_6,ec_mg_m3_6,notas_vla_6",","))

			sql3 = sql3 & generaSQLListaAsociada("vlb",split("ib_1,vlb_1,momento_1,notas_vlb_1,ib_2,vlb_2,momento_2,notas_vlb_2,ib_3,vlb_3,momento_3,notas_vlb_3,ib_4,vlb_4,momento_4,notas_vlb_4,ib_5,vlb_5,momento_5,notas_vlb_5,ib_6,vlb_6,momento_6,notas_vlb_6",","))

			sql3 = sql3 & generaSQLListaAsociada( "eper_agua", Split("",","))
			sql3 = sql3 & generaSQLListaAsociada( "eper_aire", Split("",","))
			sql3 = sql3 & generaSQLListaAsociada( "eper_suelo", Split("",","))

			sql3 = sql3 & generaSQLListaAsociada( "capa_ozono", Split("",","))
			sql3 = sql3 & generaSQLListaAsociada( "cambio_climatico", Split("",","))
			sql3 = sql3 & generaSQLListaAsociada( "calidad_aire", Split("",","))
			sql3 = sql3 & generaSQLListaAsociada( "contaminantes_suelo", Split("",","))
			sql3 = sql3 & generaSQLListaAsociada( "cop", Split("enlace_cop",","))

			sql3 = sql3 & generaSQLListaAsociada( "mpmb", Split("",","))
			sql3 = sql3 & generaSQLListaAsociada( "directiva_aguas", Split("clasif_mma",","))
			sql3 = sql3 & generaSQLListaAsociada( "peligrosas_agua_alemania", Split("",","))
			sql3 = sql3 & generaSQLListaAsociada( "tpb", Split("enlace_tpb,anchor_tpb,fuentes_tpb",","))

			sql3 = sql3 & generaSQLListaAsociada( "disruptores", Split("nivel_disruptor",","))
			sql3 = sql3 & generaSQLListaAsociada( "neuro_oto", Split("efecto_neurotoxico,nivel_neurotoxico,fuente_neurotoxico",","))

			sql3 = sql3 & generaSQLListaAsociada( "cancer_rd", Split("notas_cancer_rd",","))
			sql3 = sql3 & generaSQLListaAsociada( "cancer_iarc", Split("grupo_iarc,volumen_iarc",","))
			sql3 = sql3 & generaSQLListaAsociada( "cancer_otras", Split("categoria_cancer_otras,fuente",","))
			sql3 = sql3 & generaSQLListaAsociada( "cancer_mama", Split("cancer_mama_fuente",","))

			sql3 = sql3 & generaSQLListaAsociada( "prohibidas", Split("comentario_prohibida,comentario_prohibida_ing",","))
			sql3 = sql3 & generaSQLListaAsociada( "restringidas", Split("comentario_restringida,comentario_restringida_ing",","))
			sql3 = sql3 & generaSQLListaAsociada( "prohibidas_embarazadas", Split("comentario_prohibida,comentario_prohibida_ing",","))
			sql3 = sql3 & generaSQLListaAsociada( "prohibidas_lactantes", Split("comentario_prohibida,comentario_prohibida_ing",","))
			sql3 = sql3 & generaSQLListaAsociada( "candidatas_reach", Split("",","))
			sql3 = sql3 & generaSQLListaAsociada( "autorizacion_reach", Split("",","))

			sql3 = sql3 & generaSQLListaAsociada( "pesticidas_prohibidas", Split("fuente,exenciones,exenciones_ing",","))
			sql3 = sql3 & generaSQLListaAsociada( "pesticidas_autorizadas", Split("fuente,plazo_renovacion,pureza_minima,usos,plazo_renovacion_ing,pureza_minima_ing,usos_ing",","))
			sql3 = sql3 & generaSQLListaAsociada( "biocidas_prohibidas", Split("fuente,fecha_limite,usos,usos_ing",","))
			sql3 = sql3 & generaSQLListaAsociada( "biocidas_autorizadas", Split("fuente,pureza_minima,condiciones,usos,condiciones_ing,usos_ing",","))

			'-- /spl
			sql3=sql3& " where id=" &id

			objconn1.execute(sql3)
			padre="dn_grupos.asp"

		case "enfermedad":

			listado=	request.Form("listado")
			sintomas=	request.Form("sintomas")
			actividades=	request.Form("actividades")

			nombre_ing=	request.Form("nombre_ing")
			listado_ing=	request.Form("listado_ing")
			sintomas_ing=	request.Form("sintomas_ing")
			actividades_ing=	request.Form("actividades_ing")

			sql3="update dn_risc_enfermedades set nombre='" &nombre& "', nombre_ing='" &nombre_ing& "', listado='" &listado& "', listado_ing='" &listado_ing& "', sintomas='" &sintomas& "', sintomas_ing='" &sintomas_ing& "', actividades='" &actividades& "', actividades_ing='" &actividades_ing& "'  where id=" &id
			objconn1.execute(sql3)
			padre="dn_enfermedades.asp"

		case "uso":

			nombre_ing=	request.Form("nombre_ing")
			descripcion_ing=	request.Form("descripcion_ing")
			sql3="update dn_risc_usos set nombre='" &nombre& "', nombre_ing='" &nombre_ing& "', descripcion='"&descripcion&"', descripcion_ing='"&descripcion_ing&"' where id=" &id
			objconn1.execute(sql3)
			padre="dn_usos.asp"

		case "compania":
			direccion=request.Form("direccion")
			fuente=request.Form("fuente")
			productora=request.Form("productora")
			distribuidora=request.Form("distribuidora")
			if productora<>1 then productora=0
			if distribuidora<>1 then distribuidora=0
			sql3="update dn_risc_companias set nombre='" &nombre& "', direccion='" &direccion& "', fuente='" &fuente& "', productora=" &productora& ", distribuidora=" &distribuidora& " where id=" &id
			objconn1.execute(sql3)
			padre="dn_companias.asp"

		case "sector":

			numero_cnae=	request.Form("numero_cnae")
			sql3="update dn_alter_sectores set nombre='" &nombre& "', numero_cnae='" &numero_cnae& "'  where id=" &id
			objconn1.execute(sql3)
			padre="dn_sectores.asp"

		case "proceso":

			sql3="update dn_alter_procesos set nombre='" &nombre& "', descripcion='" &descripcion& "'  where id=" &id
			objconn1.execute(sql3)
			padre="dn_procesos.asp"

		case "residuo":

			codigo=	request.Form("codigo")
			sql3="update rq_residuos set nombre='" &nombre& "', codigo='" &codigo& "'  where id=" &id
			objconn1.execute(sql3)
			padre="dn_residuos.asp"

	end select
end if
' ** AUDITORIA
spl_accion = "modificar"
spl_entidad = asociar
spl_descripcion = sql3

flashMsgCreate "Los datos se modificaron correctamente", "OK"
' ** AUDITORIA **
call auditaYCierraConexion(spl_accion,spl_entidad,spl_descripcion) ' accion, entidad, descripcion




function generaSQLListaAsociada(lista, camposArray())
	if request.form("asoc_"&lista)=1 then
		sql = ", asoc_"&lista&"=1"
		for each c in camposArray
			sql = sql& ", asoc_"&lista&"_" & c &"='" & request.form("asoc_"&lista&"_"&c) & "'"
		next
	else
		sql = ", asoc_"&lista&"=0"
		for each c in camposArray
			sql = sql& ", asoc_"&lista&"_" & c &"=''"
		next
	end if

	generaSQLListaAsociada = sql
end function


%>

<script>
window.opener.document.location='<%=padre%>';
window.close();
</script>

ipt>

