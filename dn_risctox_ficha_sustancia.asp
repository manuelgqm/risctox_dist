<!--#include file="dn_restringida.asp"-->
<!--#include file="dn_conexion.asp"-->
<!--#include file="dn_funciones_comunes.asp"-->
<!--#include file="dn_funciones_texto.asp"-->

<%
'----- Registrar la visita
	idpagina = 627	'--- página Resultado de la búsqueda, sólo para registrar estadísticas
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

on error resume next


' Borde para ver las tablas u ocultarlas
'borde=" border='1'"
borde=""

' Inicialmente no hay errores...
errores = ""

' Cogemos el id de la sustancia elegida y traemos sus datos
id_sustancia = request("id_sustancia")
id_sustancia = EliminaInyeccionSQL(id_sustancia)

function composeSubstanceQuery(id_sustancia)
	
	'TODO: nombrar la función con su intención
	sql = "SELECT *,dn_risc_sustancias_ambiente.comentarios as comentarios_ma, dn_risc_sustancias.comentarios as comentarios_sustancia "
	sql = sql & " FROM dn_risc_sustancias  "
	sql = sql & " FULL OUTER JOIN dn_risc_sustancias_vl ON dn_risc_sustancias.id = dn_risc_sustancias_vl.id_sustancia  "
	sql = sql & " FULL OUTER JOIN dn_risc_sustancias_iarc ON dn_risc_sustancias.id = dn_risc_sustancias_iarc.id_sustancia  "
	sql = sql & " FULL OUTER JOIN dn_risc_sustancias_cancer_otras ON dn_risc_sustancias.id = dn_risc_sustancias_cancer_otras.id_sustancia  "
	sql = sql & " FULL OUTER JOIN dn_risc_sustancias_neuro_disruptor ON dn_risc_sustancias.id = dn_risc_sustancias_neuro_disruptor.id_sustancia  "
	sql = sql & " FULL OUTER JOIN dn_risc_sustancias_ambiente ON dn_risc_sustancias.id = dn_risc_sustancias_ambiente.id_sustancia  "
	sql = sql & " FULL OUTER JOIN dn_risc_sustancias_mama_cop ON dn_risc_sustancias.id = dn_risc_sustancias_mama_cop.id_sustancia  "
	sql = sql & " FULL OUTER JOIN spl_risc_sustancias_prohibidas_embarazadas ON dn_risc_sustancias.id = spl_risc_sustancias_prohibidas_embarazadas.id_sustancia  "
	sql = sql & " WHERE dn_risc_sustancias.id="&id_sustancia
	
	composeSubstanceQuery = sql

end function

sql = composeSubstanceQuery( id_sustancia )
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
	cas_alternativos = objRst("cas_alternativos")
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





	' RD1272/2008
	clasificacion_rd1272_1 = trim(objRst("clasificacion_rd1272_1"))
	clasificacion_rd1272_2 = trim(objRst("clasificacion_rd1272_2"))
	clasificacion_rd1272_3 = trim(objRst("clasificacion_rd1272_3"))
	clasificacion_rd1272_4 = trim(objRst("clasificacion_rd1272_4"))
	clasificacion_rd1272_5 = trim(objRst("clasificacion_rd1272_5"))
	clasificacion_rd1272_6 = trim(objRst("clasificacion_rd1272_6"))
	clasificacion_rd1272_7 = trim(objRst("clasificacion_rd1272_7"))
	clasificacion_rd1272_8 = trim(objRst("clasificacion_rd1272_8"))
	clasificacion_rd1272_9 = trim(objRst("clasificacion_rd1272_9"))
	clasificacion_rd1272_10 = trim(objRst("clasificacion_rd1272_10"))
	clasificacion_rd1272_11 = trim(objRst("clasificacion_rd1272_11"))
	clasificacion_rd1272_12 = trim(objRst("clasificacion_rd1272_12"))
	clasificacion_rd1272_13 = trim(objRst("clasificacion_rd1272_13"))
	clasificacion_rd1272_14 = trim(objRst("clasificacion_rd1272_14"))
	clasificacion_rd1272_15 = trim(objRst("clasificacion_rd1272_15"))
	conc_rd1272_1 = objRst("conc_rd1272_1")
	eti_conc_rd1272_1 = objRst("eti_conc_rd1272_1")
	conc_rd1272_2 = objRst("conc_rd1272_2")
	eti_conc_rd1272_2 = objRst("eti_conc_rd1272_2")
	conc_rd1272_3 = objRst("conc_rd1272_3")
	eti_conc_rd1272_3 = objRst("eti_conc_rd1272_3")
	conc_rd1272_4 = objRst("conc_rd1272_4")
	eti_conc_rd1272_4 = objRst("eti_conc_rd1272_4")
	conc_rd1272_5 = objRst("conc_rd1272_5")
	eti_conc_rd1272_5 = objRst("eti_conc_rd1272_5")
	conc_rd1272_6 = objRst("conc_rd1272_6")
	eti_conc_rd1272_6 = objRst("eti_conc_rd1272_6")
	conc_rd1272_7 = objRst("conc_rd1272_7")
	eti_conc_rd1272_7 = objRst("eti_conc_rd1272_7")
	conc_rd1272_8 = objRst("conc_rd1272_8")
	eti_conc_rd1272_8 = objRst("eti_conc_rd1272_8")
	conc_rd1272_9 = objRst("conc_rd1272_9")
	eti_conc_rd1272_9 = objRst("eti_conc_rd1272_9")
	conc_rd1272_10 = objRst("conc_rd1272_10")
	eti_conc_rd1272_10 = objRst("eti_conc_rd1272_10")
	conc_rd1272_11 = objRst("conc_rd1272_11")
	eti_conc_rd1272_11 = objRst("eti_conc_rd1272_11")
	conc_rd1272_12 = objRst("conc_rd1272_12")
	eti_conc_rd1272_12 = objRst("eti_conc_rd1272_12")
	conc_rd1272_13 = objRst("conc_rd1272_13")
	eti_conc_rd1272_13 = objRst("eti_conc_rd1272_13")
	conc_rd1272_14 = objRst("conc_rd1272_14")
	eti_conc_rd1272_14 = objRst("eti_conc_rd1272_14")
	conc_rd1272_15 = objRst("conc_rd1272_15")
	eti_conc_rd1272_15 = objRst("eti_conc_rd1272_15")
	notas_rd1272 = replace(objRst("notas_rd1272"), "@", "@ ")
	simbolos_rd1272 = objRst("simbolos_rd1272")
	clases_categorias_peligro_rd1272 = objRst("clases_categorias_peligro_rd1272")




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
	' Parche: quitar las que diga "véase Tabla 3"
	notas_cancer_rd = replace(notas_cancer_rd, "véase Tabla 3", "")

	grupo_iarc = objRst("grupo_iarc")
	volumen_iarc = objRst("volumen_iarc")
	notas_iarc = objRst("notas_iarc")
	categoria_cancer_otras = objRst("categoria_cancer_otras")
	fuente = objRst("fuente")

	' Disruptor endocrino
	nivel_disruptor = objRst("nivel_disruptor")
	dim vector_disruptores
	vector_disruptores = split(nivel_disruptor,",")


	' Neurotóxico
	efecto_neurotoxico=objRst("efecto_neurotoxico")
	nivel_neurotoxico=objRst("nivel_neurotoxico")
	fuente_neurotoxico=objRst("fuente_neurotoxico")

	' TPB
	enlace_tpb = objRst("enlace_tpb")
	anchor_tpb = objRst("anchor_tpb")
	fuente_tpb = objRst("fuentes_tpb")

	' SPL (16/06/2014)
	' mPmB
	mpmb = objRst("mpmb")

	' Tóxica para el agua
	directiva_aguas = objRst("directiva_aguas")
	clasif_mma = objRst("clasif_mma")
	sustancia_prioritaria = objrst("sustancia_prioritaria")

	' Contaminante del aire
	dano_calidad_aire = objRst("dano_calidad_aire")
	dano_ozono = objRst("dano_ozono")
	dano_cambio_clima = objRst("dano_cambio_clima")


	comentarios_medio_ambiente = objrst("comentarios_ma")
	toxicidad_suelo = objrst("toxicidad_suelo")





	' Sustancia prohibida
	sustancia_prohibida = objrst("sustancia_prohibida")
	sustancia_restringida = objrst("sustancia_restringida")
	comentarios = trim(objrst("comentarios_sustancia"))
	'response.write comentarios
	'response.write comentarios

	' Cancer Mama
	cancer_mama = objRst("cancer_mama")
	cancer_mama_fuente = objRst("cancer_mama_fuente")

  	' COP
  	cop = objRst("cop")
  	enlace_cop = objRst("enlace_cop")


end if
objRst.close()
set objRst=nothing


' **** SPL
' A continuación buscamos la relación de la sustancia con grupos que tengan información de listas asociadas y se la añadimos a los campos
' Leemos todos los grupos relacionados con la sustancia
sql = "SELECT gr.* FROM dn_risc_grupos gr, dn_risc_sustancias_por_grupos sg WHERE sg.id_grupo=gr.id AND sg.id_sustancia="&id_sustancia

set objRst=objConnection2.execute(sql)
	' Recorremos todos los grupos
	do while not objRst.eof
		call evaluaCamposListaAsociada("cancer_rd",split("notas_cancer_rd",","))
		call evaluaCamposListaAsociada("cancer_iarc",split("grupo_iarc,volumen_iarc",","))

		call evaluaCamposListaAsociada("cancer_otras",split("categoria_cancer_otras,fuente",","))
		call evaluaCamposListaAsociada("cancer_mama",split("cancer_mama_fuente",","))
		call evaluaCamposListaAsociada("neuro_oto",split("efecto_neurotoxico,nivel_neurotoxico,fuente_neurotoxico",","))
		call evaluaCamposListaAsociada("disruptores",split("nivel_disruptor",","))


		call evaluaCamposListaAsociada("tpb",split("enlace_tpb,anchor_tpb,fuentes_tpb",","))

		call evaluaCamposListaAsociada("directiva_aguas",split("clasif_mma",","))

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

		call evaluaCamposListaAsociada("cop",split("enlace_cop",","))

		call evaluaCamposListaAsociada("mpmb",split("",","))

		call evaluaCamposListaAsociada("eper",split("",","))
		call evaluaCamposListaAsociada("eper_agua",split("",","))
		call evaluaCamposListaAsociada("eper_aire",split("",","))
		call evaluaCamposListaAsociada("eper_suelo",split("",","))


		call evaluaCamposListaAsociada("prohibidas",split("comentario_prohibida",","))
		call evaluaCamposListaAsociada("restringidas",split("comentario_restringida",","))

		call evaluaCamposListaAsociada("prohibidas_embarazadas",split("comentario_prohibida",","))
		call evaluaCamposListaAsociada("prohibidas_lactantes",split("comentario_prohibida",","))
		call evaluaCamposListaAsociada("candidatas_reach",split("",","))
		call evaluaCamposListaAsociada("autorizacion_reach",split("",","))

		call evaluaCamposListaAsociada("biocidas_autorizadas",split("fuente,pureza_minima,condiciones,usos",","))
		call evaluaCamposListaAsociada("biocidas_prohibidas",split("fuente,fecha_limite,usos",","))

		call evaluaCamposListaAsociada("pesticidas_autorizadas",split("fuente,plazo_renovacion,pureza_minima,usos",","))
		call evaluaCamposListaAsociada("pesticidas_prohibidas",split("fuente,exenciones",","))

		call evaluaCamposListaAsociada("alergeno",split("",","))

		call evaluaCamposListaAsociada("calidad_aire",split("",","))
		
		call evaluaCamposListaAsociada( "corap", split("", ",") )

		objRst.movenext
	loop
objRst.close()


' **** /SPL


' Sinonimos
sinonimos = dameSinonimos(id_sustancia)

' Comprobamos si está en cada lista, para no tener que buscar varias veces
esta_en_lista_cancer_rd = esta_en_lista_cancer_rd or esta_en_lista ("cancer_rd", id_sustancia)
esta_en_lista_cancer_danesa = esta_en_lista_cancer_danesa or esta_en_lista ("cancer_danesa", id_sustancia)
esta_en_lista_mutageno_rd = esta_en_lista_mutageno_rd or esta_en_lista ("mutageno_rd", id_sustancia)
esta_en_lista_mutageno_danesa = esta_en_lista_mutageno_danesa or esta_en_lista ("mutageno_danesa", id_sustancia)
esta_en_lista_cancer_iarc = esta_en_lista_cancer_iarc or esta_en_lista ("cancer_iarc", id_sustancia)
esta_en_lista_cancer_iarc_excepto_grupo_3 = esta_en_lista_cancer_iarc_excepto_grupo_3 or esta_en_lista ("cancer_iarc_excepto_grupo_3", id_sustancia)
esta_en_lista_cancer_otras = esta_en_lista_cancer_otras or esta_en_lista ("cancer_otras", id_sustancia)

esta_en_lista_cancer_mama = esta_en_lista_cancer_mama or esta_en_lista ("cancer_mama", id_sustancia)
esta_en_lista_tpr = esta_en_lista_tpr or esta_en_lista ("tpr", id_sustancia)
esta_en_lista_tpr_danesa = esta_en_lista_tpr_danesa or esta_en_lista ("tpr_danesa", id_sustancia)
esta_en_lista_de = esta_en_lista_de or esta_en_lista ("de", id_sustancia)
esta_en_lista_neurotoxico_rd = esta_en_lista_neurotoxico_rd or esta_en_lista ("neurotoxico_rd", id_sustancia)
esta_en_lista_neurotoxico_danesa = esta_en_lista_neurotoxico_danesa or  esta_en_lista ("neurotoxico_danesa", id_sustancia)
esta_en_lista_neurotoxico_nivel = esta_en_lista_neurotoxico_nivel or esta_en_lista ("neurotoxico_nivel", id_sustancia)
'dn_risc_sustancias_neuro_disruptor.efecto_neurotoxico='OTOTÓXICO'
esta_en_lista_neurotoxico = esta_en_lista_neurotoxico or esta_en_lista_neurotoxico_rd OR esta_en_lista_neurotoxico_danesa OR esta_en_lista_neurotoxico_nivel OR esta_en_lista ("neurotoxico", id_sustancia)


esta_en_lista_sensibilizante = esta_en_lista_sensibilizante or esta_en_lista ("sensibilizante", id_sustancia)
esta_en_lista_sensibilizante_danesa = esta_en_lista_sensibilizante_danesa or esta_en_lista ("sensibilizante_danesa", id_sustancia)
esta_en_lista_sensibilizante_reach = esta_en_lista_sensibilizante_reach or esta_en_lista_alergenos or esta_en_lista ("sensibilizante_reach", id_sustancia) 'en_lista_alergenos es el equivalente a sensibilizantes_reach para grupos.
esta_en_lista_eepp = esta_en_lista_eepp or esta_en_lista ("eepp", id_sustancia)
esta_en_lista_tpb = esta_en_lista_tpb or esta_en_lista ("tpb", id_sustancia)

' SPL (16/06/2014)
esta_en_lista_mpmb = mpmb or esta_en_lista_mpmb

esta_en_lista_directiva_aguas =  esta_en_lista_directiva_aguas or esta_en_lista ("directiva_aguas", id_sustancia)
esta_en_lista_sustancias_prioritarias = esta_en_lista_sustancias_prioritarias or esta_en_lista ("sustancias_prioritarias", id_sustancia)
esta_en_lista_alemana = esta_en_lista_alemana or esta_en_lista ("alemana", id_sustancia)

esta_en_lista_aire = esta_en_lista_aire  or esta_en_lista_calidad_aire or esta_en_lista ("aire", id_sustancia)
esta_en_lista_ozono = esta_en_lista_ozono or esta_en_lista ("ozono", id_sustancia)
esta_en_lista_clima = esta_en_lista_clima  or esta_en_lista ("clima", id_sustancia)
'esta_en_lista_aire = esta_en_lista_aire  or esta_en_lista ("aire", id_sustancia)

esta_en_lista_suelos = esta_en_lista_suelos or esta_en_lista ("suelos", id_sustancia)

esta_en_lista_cov = esta_en_lista_cov or esta_en_lista ("cov", id_sustancia)
esta_en_lista_vertidos = esta_en_lista_vertidos or esta_en_lista ("vertidos", id_sustancia)
' Como las listas en grupos tienen diferente nombre, en este caso el 'or' es entre listas diferentes
esta_en_lista_lpcic = esta_en_lista ("lpcic", id_sustancia) or esta_en_lista_eper
esta_en_lista_lpcic_agua = esta_en_lista ("lpcic-agua", id_sustancia) or esta_en_lista_eper_agua
esta_en_lista_lpcic_aire = esta_en_lista ("lpcic-aire", id_sustancia) or esta_en_lista_eper_aire
esta_en_lista_lpcic_suelo = esta_en_lista ("lpcic-suelo", id_sustancia) or esta_en_lista_eper_suelo
esta_en_lista_residuos = esta_en_lista_residuos or esta_en_lista ("residuos", id_sustancia)
esta_en_lista_accidentes = esta_en_lista_accidentes or esta_en_lista ("accidentes", id_sustancia)
esta_en_lista_emisiones = esta_en_lista_emisiones or esta_en_lista ("emisiones", id_sustancia)
esta_en_lista_salud = esta_en_lista_salud or esta_en_lista ("salud", id_sustancia)

esta_en_lista_prohibidas = esta_en_lista_prohibidas or esta_en_lista ("prohibidas", id_sustancia)
esta_en_lista_restringidas = esta_en_lista_restringidas or esta_en_lista ("restringidas", id_sustancia)


esta_en_lista_cop = esta_en_lista_cop or esta_en_lista ("cop", id_sustancia)


'--SPL
esta_en_lista_prohibidas_embarazadas = esta_en_lista_prohibidas_embarazadas or esta_en_lista ("prohibidas_embarazadas", id_sustancia)

esta_en_lista_prohibidas_lactantes = esta_en_lista_prohibidas_lactantes or esta_en_lista ("prohibidas_lactantes", id_sustancia)

esta_en_lista_candidatas_reach = esta_en_lista_candidatas_reach or esta_en_lista ("candidatas_reach", id_sustancia)
esta_en_lista_autorizacion_reach = esta_en_lista_autorizacion_reach or esta_en_lista ("autorizacion_reach", id_sustancia)

esta_en_lista_biocidas_autorizadas = esta_en_lista_biocidas_autorizadas or esta_en_lista ("biocidas_autorizadas", id_sustancia)
esta_en_lista_biocidas_prohibidas = esta_en_lista_biocidas_prohibidas or esta_en_lista ("biocidas_prohibidas", id_sustancia)
esta_en_lista_pesticidas_autorizadas = esta_en_lista_pesticidas_autorizadas or esta_en_lista ("pesticidas_autorizadas", id_sustancia)
esta_en_lista_pesticidas_prohibidas = esta_en_lista_pesticidas_prohibidas or esta_en_lista ("pesticidas_prohibidas", id_sustancia)

esta_en_lista_corap = esta_en_lista_corap or esta_en_lista ("corap", id_sustancia)

'--/SPL
' Condiciones para mostrar las frases R danesas en Clasificacion

' Se mostrarán si existen las frases R danesas y NO existen las de RD



' Montamos frases R
frases_r=trim(monta_frases("R", clasificacion_1, clasificacion_2, clasificacion_3, clasificacion_4, clasificacion_5, clasificacion_6, clasificacion_7, clasificacion_8, clasificacion_9, clasificacion_10, clasificacion_11, clasificacion_12, clasificacion_13, clasificacion_14, clasificacion_15))



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
<title>ISTAS: risctox</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
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
        ' Montamos enlace para abrir ventana emergente de descripción
        enlace_descripcion = " <a onclick=window.open('dn_glosario.asp?tabla=grupos&id="&id_grupo&"','def','width=500,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a>"
      else
        ' No hay descripción
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
        ' Montamos enlace para abrir ventana emergente de descripción
        enlace_descripcion = " <a onclick=window.open('dn_glosario.asp?tabla=usos&id="&id_uso&"','def','width=500,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'>"&nombre_uso&"</a>"
      else
        ' No hay descripción
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
	' Devuelve lista de compañías para la sustancia indicada

	lista = ""

	sql="SELECT dn_risc_companias.id as idcomp, nombre FROM dn_risc_sustancias_por_companias INNER JOIN dn_risc_companias ON dn_risc_sustancias_por_companias.id_compania = dn_risc_companias.id WHERE id_sustancia="&id_sustancia&" ORDER BY nombre"
	set objRst=objConnection2.execute(sql)
	if (not objRst.eof) then
		do while (not objRst.eof)
			if (lista = "") then
				lista = "<a onclick=window.open('dn_risctox_ficha_compania.asp?id="&objRst("idcomp")&"','comp','width=600,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'>"&objRst("nombre")&"</a>"
			else
				lista = lista&", <a onclick=window.open('dn_risctox_ficha_compania.asp?id="&objRst("idcomp")&"','comp','width=600,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'>"&objRst("nombre")&"</a>"
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
			<a onclick=window.open('ver_definicion.asp?id=82','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a> Nombre:
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
				<a onclick=window.open('ver_definicion.asp?id=83','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a> Sinónimos:
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
				Números de Identificación:
			</td>
			<td class="texto" valign="middle">
				<% if (num_cas <> "") then response.write "<a onclick=window.open('ver_definicion.asp?id=84','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a> <b>CAS</b>: "&num_cas&"<br/>" %>
				<% if (cas_alternativos <> "") then response.write "<a onclick=window.open('ver_definicion.asp?id=84','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a> <b>CAS alternativos</b>: "&cas_alternativos &"<br/>" %>
				<%
					if (num_ce_einecs <> "") then
						'Sergio, si empieza por 4 y num_ce_elincs<>'' muestro el num_ce_elincs
						if (mid(num_ce_einecs,1,1)="4" and num_ce_elincs <> "") then
							response.write "<a onclick=window.open('ver_definicion.asp?id=85','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a> <b>CE ELINCS</b>: "&num_ce_elincs&"<br/>"
						else
						response.write "<a onclick=window.open('ver_definicion.asp?id=85','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a> <b>CE EINECS</b>: "&num_ce_einecs&"<br/>"
						end if
					elseif (num_ce_elincs <> "") then
						response.write "<a onclick=window.open('ver_definicion.asp?id=85','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a> <b>CE ELINCS</b>: "&num_ce_elincs&"<br/>"
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
				 Ficha Internacional de Seguridad Química (<a onClick="window.open('ver_definicion.asp?id=<%=dame_id_definicion("INSHT")%>', 'def', 'width=300,height=200,scrollbars=yes,resizable=yes')" class="subtitulo3">INSHT</a>)
			</td>
			<td class="texto" valign="middle">
          <%

            array_icsc=split(num_icsc, "@")

            for i=0 to ubound(array_icsc)
            	num_icsc = cstr(array_icsc(i))
            	if len(num_icsc)=4 then
            		centena_icsc = mid(num_icsc,1,2)
            		icsc_max = cstr(clng(centena_icsc&"01"))
            		if icsc_max="1" then icsc_max="0"
            		icsc_min = cstr(clng(centena_icsc)+1) & "00"
            	end if

          %>

              <!--<a href="http://www.mtas.es/insht/ipcsnspn/nspn<%= array_icsc(i) %>.htm" target="_blank"><%= array_icsc(i) %></a> -->
              <a href="http://www.insht.es/InshtWeb/Contenidos/Documentacion/FichasTecnicas/FISQ/Ficheros/<%=icsc_max%>a<%=icsc_min%>/nspn<%= array_icsc(i) %>.pdf" target="_blank"><%= array_icsc(i) %></a>

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
				Compañías productoras/distribuidoras:
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
				Más información <% plegador "secc-masinformacion", "img-masinformacion" %>
			</td>
			<td class="texto" valign="middle" id="secc-masinformacion" style="display:none">



        <% if (nombre_ing <> "") then
            array_nombres_ingleses = split(nombre_ing, "@")
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
              <b>Nombre inglés</b>: <%= espaciar(nombre_ing) %><br/>
        <%
            end if
           end if %>

				<% if (num_rd <> "") then response.write "<a onclick=window.open('ver_definicion.asp?id=86','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' /></a> <b>Nº &iacute;ndice</b>: "&num_rd&"<br/>" %>
				<% if (formula_molecular <> "") then response.write "<b>Fórmula molecular</b>: "&formula_molecular&"<br/>" %>
				<% if (estructura_molecular <> "") then response.write "<b>Estructura molecular</b>:<br /><img src='../gestion/estructuras/"&estructura_molecular&"' /><br/>" %>

				<% if (notas_xml <> "") then %>
          <a onClick="window.open('ver_definicion.asp?id=<%=dame_id_definicion("ECB")%>', 'def', 'width=300,height=200,scrollbars=yes,resizable=yes')" style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>
          <b>Notas ECB</b>: <%= espaciar(notas_xml) %> <br />
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

			<% ap2_clasificacion_lista_negra() %>
		</td>
	</tr>
<%
end sub ' ap1_identificacion

' ###################################################################################

sub ap2_clasificacion()
	' Solo mostramos este apartado si hay información para él
'	if ((simbolos <> "") or (clasificacion_1 <> "") or (clasificacion_2 <> "") or (clasificacion_3 <> "") or (clasificacion_4 <> "") or (clasificacion_5 <> "") or (clasificacion_6 <> "") or (clasificacion_7 <> "") or (clasificacion_8 <> "") or (clasificacion_9 <> "") or (clasificacion_10 <> "") or (clasificacion_11 <> "") or (clasificacion_12 <> "") or (clasificacion_13 <> "") or (clasificacion_14 <> "") or (clasificacion_15 <> "") or (frases_r_danesa <> "") or (notas_rd_363 <> "") or (conc_1 <> "") or (eti_conc_1 <> "") or (conc_2 <> "") or (eti_conc_2 <> "") or (conc_3 <> "") or (eti_conc_3 <> "") or (conc_4 <> "") or (eti_conc_4 <> "") or (conc_5 <> "") or (eti_conc_5 <> "") or (conc_6 <> "") or (eti_conc_6 <> "") or (conc_7 <> "") or (eti_conc_7 <> "") or (conc_8 <> "") or (eti_conc_8 <> "") or (conc_9 <> "") or (eti_conc_9 <> "") or (conc_10 <> "") or (eti_conc_10 <> "") or (conc_11 <> "") or (eti_conc_11 <> "") or (conc_12 <> "") or (eti_conc_12 <> "") or (conc_13 <> "") or (eti_conc_13 <> "") or (conc_14 <> "") or (eti_conc_14 <> "") or (conc_15 <> "") or (eti_conc_15 <> "") or (estado_1 <> "") or (vla_ed_ppm_1 <> "") or (vla_ed_mg_m3_1 <> "") or (vla_ec_ppm_1 <> "") or (vla_ec_mg_m3_1 <> "") or (notas_vla_1 <> "") or (estado_2 <> "") or (vla_ed_ppm_2 <> "") or (vla_ed_mg_m3_2 <> "") or (vla_ec_ppm_2 <> "") or (vla_ec_mg_m3_2 <> "") or (notas_vla_2 <> "") or (estado_3 <> "") or (vla_ed_ppm_3 <> "") or (vla_ed_mg_m3_3 <> "") or (vla_ec_ppm_3 <> "") or (vla_ec_mg_m3_3 <> "") or (notas_vla_3 <> "") or (estado_4 <> "") or (vla_ed_ppm_4 <> "") or (vla_ed_mg_m3_4 <> "") or (vla_ec_ppm_4 <> "") or (vla_ec_mg_m3_4 <> "") or (notas_vla_4 <> "") or (estado_5 <> "") or (vla_ed_ppm_5 <> "") or (vla_ed_mg_m3_5 <> "") or (vla_ec_ppm_5 <> "") or (vla_ec_mg_m3_5 <> "") or (notas_vla_5 <> "") or (estado_6 <> "") or (vla_ed_ppm_6 <> "") or (vla_ed_mg_m3_6 <> "") or (vla_ec_ppm_6  <> "") or (vla_ec_mg_m3_6 <> "") or (notas_vla_6 <> "") or (ib_1 <> "") or  (vlb_1 <> "") or (momento_1 <> "") or (notas_vlb_1 <> "") or (ib_2 <> "") or  (vlb_2 <> "") or (momento_2 <> "") or (notas_vlb_2 <> "") or (ib_3 <> "") or  (vlb_3 <> "") or (momento_3 <> "") or (notas_vlb_3 <> "") or (ib_4 <> "") or  (vlb_4 <> "") or (momento_4 <> "") or (notas_vlb_4 <> "") or (ib_5 <> "") or  (vlb_5 <> "") or (momento_5 <> "") or (notas_vlb_5 <> "") or (ib_6 <> "") or  (vlb_6 <> "") or (momento_6 <> "") or (notas_vlb_6 <> "") or esta_en_lista_cancer_rd or esta_en_lista_cancer_iarc_excepto_grupo_3 or esta_en_lista_cancer_otras or esta_en_lista_de or esta_en_lista_neurotoxico or  esta_en_lista_tpb or esta_en_lista_sensibilizante or esta_en_lista_tpr or esta_en_lista_cancer_mama or esta_en_lista_cop or esta_en_lista_prohibidas_embarazadas or esta_en_lista_prohibidas_lactantes) then
	if ((simbolos <> "") or (clasificacion_1 <> "") or (clasificacion_2 <> "") or (clasificacion_3 <> "") or (clasificacion_4 <> "") or (clasificacion_5 <> "") or (clasificacion_6 <> "") or (clasificacion_7 <> "") or (clasificacion_8 <> "") or (clasificacion_9 <> "") or (clasificacion_10 <> "") or (clasificacion_11 <> "") or (clasificacion_12 <> "") or (clasificacion_13 <> "") or (clasificacion_14 <> "") or (clasificacion_15 <> "") or (frases_r_danesa <> "") or (notas_rd_363 <> "") or (conc_1 <> "") or (eti_conc_1 <> "") or (conc_2 <> "") or (eti_conc_2 <> "") or (conc_3 <> "") or (eti_conc_3 <> "") or (conc_4 <> "") or (eti_conc_4 <> "") or (conc_5 <> "") or (eti_conc_5 <> "") or (conc_6 <> "") or (eti_conc_6 <> "") or (conc_7 <> "") or (eti_conc_7 <> "") or (conc_8 <> "") or (eti_conc_8 <> "") or (conc_9 <> "") or (eti_conc_9 <> "") or (conc_10 <> "") or (eti_conc_10 <> "") or (conc_11 <> "") or (eti_conc_11 <> "") or (conc_12 <> "") or (eti_conc_12 <> "") or (conc_13 <> "") or (eti_conc_13 <> "") or (conc_14 <> "") or (eti_conc_14 <> "") or (conc_15 <> "") or (eti_conc_15 <> "") ) then

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


		</table>
		</td>
		</tr>
	</table>
<%
	end if
end sub ' ap2_clasificacion






sub ap2_clasificacion_rd1272()
	' Solo mostramos este apartado si hay información para él
'	if ((simbolos_rd1272 <> "") or (clasificacion_rd1272_1 <> "") or (clasificacion_rd1272_2 <> "") or (clasificacion_rd1272_3 <> "") or (clasificacion_rd1272_4 <> "") or (clasificacion_rd1272_5 <> "") or (clasificacion_rd1272_6 <> "") or (clasificacion_rd1272_7 <> "") or (clasificacion_rd1272_8 <> "") or (clasificacion_rd1272_9 <> "") or (clasificacion_rd1272_10 <> "") or (clasificacion_rd1272_11 <> "") or (clasificacion_rd1272_12 <> "") or (clasificacion_rd1272_13 <> "") or (clasificacion_rd1272_14 <> "") or (clasificacion_rd1272_15 <> "") or (conc_rd1272_1 <> "") or (eti_conc_rd1272_1 <> "") or (conc_rd1272_2 <> "") or (eti_conc_rd1272_2 <> "") or (conc_rd1272_3 <> "") or (eti_conc_rd1272_3 <> "") or (conc_rd1272_4 <> "") or (eti_conc_rd1272_4 <> "") or (conc_rd1272_5 <> "") or (eti_conc_rd1272_5 <> "") or (conc_rd1272_6 <> "") or (eti_conc_rd1272_6 <> "") or (conc_rd1272_7 <> "") or (eti_conc_rd1272_7 <> "") or (conc_rd1272_8 <> "") or (eti_conc_rd1272_8 <> "") or (conc_rd1272_9 <> "") or (eti_conc_rd1272_9 <> "") or (conc_rd1272_10 <> "") or (eti_conc_rd1272_10 <> "") or (conc_rd1272_11 <> "") or (eti_conc_rd1272_11 <> "") or (conc_rd1272_12 <> "") or (eti_conc_rd1272_12 <> "") or (conc_rd1272_13 <> "") or (eti_conc_rd1272_13 <> "") or (conc_rd1272_14 <> "") or (eti_conc_rd1272_14 <> "") or (conc_rd1272_15 <> "") or (eti_conc_rd1272_15 <> "") or (estado_1 <> "") or (vla_ed_ppm_1 <> "") or (vla_ed_mg_m3_1 <> "") or (vla_ec_ppm_1 <> "") or (vla_ec_mg_m3_1 <> "") or (notas_vla_1 <> "") or (estado_2 <> "") or (vla_ed_ppm_2 <> "") or (vla_ed_mg_m3_2 <> "") or (vla_ec_ppm_2 <> "") or (vla_ec_mg_m3_2 <> "") or (notas_vla_2 <> "") or (estado_3 <> "") or (vla_ed_ppm_3 <> "") or (vla_ed_mg_m3_3 <> "") or (vla_ec_ppm_3 <> "") or (vla_ec_mg_m3_3 <> "") or (notas_vla_3 <> "") or (estado_4 <> "") or (vla_ed_ppm_4 <> "") or (vla_ed_mg_m3_4 <> "") or (vla_ec_ppm_4 <> "") or (vla_ec_mg_m3_4 <> "") or (notas_vla_4 <> "") or (estado_5 <> "") or (vla_ed_ppm_5 <> "") or (vla_ed_mg_m3_5 <> "") or (vla_ec_ppm_5 <> "") or (vla_ec_mg_m3_5 <> "") or (notas_vla_5 <> "") or (estado_6 <> "") or (vla_ed_ppm_6 <> "") or (vla_ed_mg_m3_6 <> "") or (vla_ec_ppm_6  <> "") or (vla_ec_mg_m3_6 <> "") or (notas_vla_6 <> "") or (ib_1 <> "") or  (vlb_1 <> "") or (momento_1 <> "") or (notas_vlb_1 <> "") or (ib_2 <> "") or  (vlb_2 <> "") or (momento_2 <> "") or (notas_vlb_2 <> "") or (ib_3 <> "") or  (vlb_3 <> "") or (momento_3 <> "") or (notas_vlb_3 <> "") or (ib_4 <> "") or  (vlb_4 <> "") or (momento_4 <> "") or (notas_vlb_4 <> "") or (ib_5 <> "") or  (vlb_5 <> "") or (momento_5 <> "") or (notas_vlb_5 <> "") or (ib_6 <> "") or  (vlb_6 <> "") or (momento_6 <> "") or (notas_vlb_6 <> "") or esta_en_lista_cancer_rd or esta_en_lista_cancer_iarc_excepto_grupo_3 or esta_en_lista_cancer_otras or esta_en_lista_de or esta_en_lista_neurotoxico or  esta_en_lista_tpb or esta_en_lista_sensibilizante or esta_en_lista_tpr or esta_en_lista_cancer_mama or esta_en_lista_cop or esta_en_lista_prohibidas_embarazadas or esta_en_lista_prohibidas_lactantes) then
	if ((simbolos_rd1272 <> "") or (clasificacion_rd1272_1 <> "") or (clasificacion_rd1272_2 <> "") or (clasificacion_rd1272_3 <> "") or (clasificacion_rd1272_4 <> "") or (clasificacion_rd1272_5 <> "") or (clasificacion_rd1272_6 <> "") or (clasificacion_rd1272_7 <> "") or (clasificacion_rd1272_8 <> "") or (clasificacion_rd1272_9 <> "") or (clasificacion_rd1272_10 <> "") or (clasificacion_rd1272_11 <> "") or (clasificacion_rd1272_12 <> "") or (clasificacion_rd1272_13 <> "") or (clasificacion_rd1272_14 <> "") or (clasificacion_rd1272_15 <> "") or (conc_rd1272_1 <> "") or (eti_conc_rd1272_1 <> "") or (conc_rd1272_2 <> "") or (eti_conc_rd1272_2 <> "") or (conc_rd1272_3 <> "") or (eti_conc_rd1272_3 <> "") or (conc_rd1272_4 <> "") or (eti_conc_rd1272_4 <> "") or (conc_rd1272_5 <> "") or (eti_conc_rd1272_5 <> "") or (conc_rd1272_6 <> "") or (eti_conc_rd1272_6 <> "") or (conc_rd1272_7 <> "") or (eti_conc_rd1272_7 <> "") or (conc_rd1272_8 <> "") or (eti_conc_rd1272_8 <> "") or (conc_rd1272_9 <> "") or (eti_conc_rd1272_9 <> "") or (conc_rd1272_10 <> "") or (eti_conc_rd1272_10 <> "") or (conc_rd1272_11 <> "") or (eti_conc_rd1272_11 <> "") or (conc_rd1272_12 <> "") or (eti_conc_rd1272_12 <> "") or (conc_rd1272_13 <> "") or (eti_conc_rd1272_13 <> "") or (conc_rd1272_14 <> "") or (eti_conc_rd1272_14 <> "") or (conc_rd1272_15 <> "") or (eti_conc_rd1272_15 <> "") ) then

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
	if (simbolos <> "") then
%>
		<p id="ap2_clasificacion_simbolos_titulo" class="ficha_titulo_2">Símbolos</p>
		<p id="ap2_clasificacion_simbolos_cuerpo" class="texto" align="center">
<%
		' Tiene símbolos, muestro cada uno
		simbolos = replace(simbolos, ",", ";")
		array_simbolos = split(simbolos, ";")
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
	if (simbolos_rd1272 <> "") then
%>
		<p id="ap2_clasificacion_simbolos_titulo" class="ficha_titulo_2">Pictogramas y palabras de advertencia</p>
		<p id="ap2_clasificacion_simbolos_cuerpo" class="texto" align="center">
<%
		' Tiene símbolos, muestro cada uno
		simbolos = replace(simbolos_rd1272, ",", ";")
		array_simbolos = split(simbolos, ";")
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

sub ap2_clasificacion_frases_r()
	' Muestra las frases R segun clasificacion_1 hasta clasificacion_15
	' No incluye las frases R danesas

	' Montamos frases R
	frases_r=monta_frases("R", clasificacion_1, clasificacion_2, clasificacion_3, clasificacion_4, clasificacion_5, clasificacion_6, clasificacion_7, clasificacion_8, clasificacion_9, clasificacion_10, clasificacion_11, clasificacion_12, clasificacion_13, clasificacion_14, clasificacion_15)

	if (frases_r <> "") then
%>
		<p id="ap2_clasificacion_frases_r_titulo" class="ficha_titulo_2" style="margin-bottom: -10px;"><a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("Frases R")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> Frases R</p>
<%
		bucle_frases "r", frases_r
	end if
end sub

' ##################################################################################

sub ap2_clasificacion_frases_h()
	' Muestra las frases H segun clasificacion_rd1272_1 hasta clasificacion_rd1272_15

	' Montamos frases H
	frases_h=monta_frases("H", clasificacion_rd1272_1, clasificacion_rd1272_2, clasificacion_rd1272_3, clasificacion_rd1272_4, clasificacion_rd1272_5, clasificacion_rd1272_6, clasificacion_rd1272_7, clasificacion_rd1272_8, clasificacion_rd1272_9, clasificacion_rd1272_10, clasificacion_rd1272_11, clasificacion_rd1272_12, clasificacion_rd1272_13, clasificacion_rd1272_14, clasificacion_rd1272_15)

	if (frases_h <> "") then
%>
		<p id="ap2_clasificacion_frases_r_titulo" class="ficha_titulo_2" style="margin-bottom: -10px;"><a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("Frases H")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> Frases H</p>
<%
'		bucle_frases "h", frases_h
		muestra_clasificacion 1, clasificacion_rd1272_1
		muestra_clasificacion 2, clasificacion_rd1272_2
		muestra_clasificacion 3, clasificacion_rd1272_3
		muestra_clasificacion 4, clasificacion_rd1272_4
		muestra_clasificacion 5, clasificacion_rd1272_5
		muestra_clasificacion 6, clasificacion_rd1272_6
		muestra_clasificacion 7, clasificacion_rd1272_7
		muestra_clasificacion 8, clasificacion_rd1272_8
		muestra_clasificacion 9, clasificacion_rd1272_9
		muestra_clasificacion 10, clasificacion_rd1272_10
		muestra_clasificacion 11, clasificacion_rd1272_11
		muestra_clasificacion 12, clasificacion_rd1272_12
		muestra_clasificacion 13, clasificacion_rd1272_13
		muestra_clasificacion 14, clasificacion_rd1272_14
		muestra_clasificacion 15, clasificacion_rd1272_15
	end if
	' 23/06/2014 - SPL - Por indicación de Tatiana se pone esta definición.
	if (trim(clasificacion_rd1272_1)="Expl., ****;") then
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
	        <a href="javascript:toggle('<%= "secc-categpeligro-"+CStr(numero) %>', '<%= "img-fraseh-"+CStr(numero) %>');"><img src="imagenes/desplegar.gif" align="absmiddle" id="<%= "img-fraseh-"+CStr(numero) %>" alt="Pulse para ver el etiquetado" title="Pulse para ver el etiquetado" /></a>
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

sub ap2_clasificacion_frases_r_danesa()
	' Muestra las frases R danesas

	' Montamos frases R
	frases_r = monta_frases_r_danesa(frases_r_danesa)

	if (frases_r <> "") then
%>
	<p id="ap2_clasificacion_frases_r_danesa_titulo" class="ficha_titulo_2"><a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("Frases R según la lista danesa de la EPA")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> Frases R según clasificación de la EPA danesa</p>
<%
		bucle_frases "r", frases_r
	end if
end sub


' ##################################################################################

sub ap2_clasificacion_frases_s
	' Muestra las frases S

	if (frases_s <> "") then
		' Eliminamos los paréntesis de las frases S
		frases_s = replace (frases_s, "(", "")
		frases_s = replace (frases_s, ")", "")

%>
	<p id="ap2_clasificacion_frases_s_titulo" class="ficha_titulo_2" style="margin-top: 14px;"><a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("Frases S")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> Frases S <% plegador "secc-frasess", "img-frasess" %></p>
		<!-- <%= frases_s %> <a onclick="window.open('busca_frases_s.asp?id=<%= frases_s %>', 'fr','width=300,height=200,scrollbars=yes,resizable=yes')" style="text-decoration:none; cursor:pointer;"><img src="imagenes/ayuda.gif" border="0" align="absmiddle" alt="busca Frases S"></a> -->

		<% bucle_frases_s frases_s%>

<%
	end if
end sub



' ##################################################################################

sub ap2_clasificacion_categorias_peligro
	' Muestra las frases

	if (clases_categorias_peligro_rd1272 <> "") then
		' Eliminamos los paréntesis
		clases_categorias_peligro_rd1272 = replace (clases_categorias_peligro_rd1272, "(", "")
		clases_categorias_peligro_rd1272 = replace (clases_categorias_peligro_rd1272, ")", "")

%>
	<p id="ap2_clasificacion_frases_s_titulo" class="ficha_titulo_2" style="margin-top: 14px;"><a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("Frases S")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> Clase y categoría de peligro <% plegador "secc-categpeligro", "img-frasess" %></p>

		<% bucle_categorias_peligro_rd1272 clases_categorias_peligro_rd1272%>

<%
	end if
end sub 'ap2_clasificacion_categorias_peligro


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
	if (notas_rd1272 <> "") then

		' Dividimos las notas, separadas por puntos, en un array
		array_notas = split(notas_rd1272, ".")
%>
	<p id="ap2_clasificacion_notas_titulo" class="ficha_titulo_2">Notas <% plegador "secc-notas-rd1272", "img-notas-rd1272" %></p>
	<p class="texto" >
		<blockquote id="secc-notas-rd1272" style="display:none">
<%
		for i=0 to ubound(array_notas)
			nota = trim(array_notas(i))
			id_nota = dame_id_definicion("R.1272-"+nota)
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

sub ap2_clasificacion_etiquetado()
	' Muestra el etiquetado

	if ((conc_1 <> "") or (eti_conc_1 <> "") or (conc_2 <> "") or (eti_conc_2 <> "") or (conc_3 <> "") or (eti_conc_3 <> "") or (conc_4 <> "") or (eti_conc_4 <> "") or (conc_5 <> "") or (eti_conc_5 <> "") or (conc_6 <> "") or (eti_conc_6 <> "") or (conc_7 <> "") or (eti_conc_7 <> "") or (conc_8 <> "") or (eti_conc_8 <> "") or (conc_9 <> "") or (eti_conc_9 <> "") or (conc_10 <> "") or (eti_conc_10 <> "") or (conc_11 <> "") or (eti_conc_11 <> "") or (conc_12 <> "") or (eti_conc_12 <> "") or (conc_13 <> "") or (eti_conc_13 <> "") or (conc_14 <> "") or (eti_conc_14 <> "") or (conc_15 <> "") or (eti_conc_15 <> "")) then

%>
	<span id="ap2_clasificacion_etiquetado_titulo" class="ficha_titulo_2"><a onclick=window.open('ver_definicion.asp?id=88','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> Etiquetado <% plegador "secc-etiquetado", "img-etiquetado" %></span>


  <fieldset id="secc-etiquetado" style="display:none; margin: 15px 45px;">
	<table cellspacing="0" cellpadding="3" width="100%" align="center">
		<tr>
			<th class="subtitulo3 celdaabajo">Concentración</th><th class="subtitulo3 celdaabajo">Etiquetado</th>
		</tr>
<%
	ap2_clasificacion_etiquetado_fila	"r", conc_1, eti_conc_1
	ap2_clasificacion_etiquetado_fila	"r", conc_2, eti_conc_2
	ap2_clasificacion_etiquetado_fila	"r", conc_3, eti_conc_3
	ap2_clasificacion_etiquetado_fila	"r", conc_4, eti_conc_4
	ap2_clasificacion_etiquetado_fila	"r", conc_5, eti_conc_5
	ap2_clasificacion_etiquetado_fila	"r", conc_6, eti_conc_6
	ap2_clasificacion_etiquetado_fila	"r", conc_7, eti_conc_7
	ap2_clasificacion_etiquetado_fila	"r", conc_8, eti_conc_8
	ap2_clasificacion_etiquetado_fila	"r", conc_9, eti_conc_9
	ap2_clasificacion_etiquetado_fila	"r", conc_10, eti_conc_10
	ap2_clasificacion_etiquetado_fila	"r", conc_11, eti_conc_11
	ap2_clasificacion_etiquetado_fila	"r", conc_12, eti_conc_12
	ap2_clasificacion_etiquetado_fila	"r", conc_13, eti_conc_13
	ap2_clasificacion_etiquetado_fila	"r", conc_14, eti_conc_14
	ap2_clasificacion_etiquetado_fila	"r", conc_15, eti_conc_15
%>
	</table>
  </fieldset>

<%
	end if
end sub


' ##################################################################################

sub ap2_clasificacion_etiquetado_rd1272()
	' Muestra el etiquetado

	if ((conc_rd1272_1 <> "") or (eti_conc_rd1272_1 <> "") or (conc_rd1272_2 <> "") or (eti_conc_rd1272_2 <> "") or (conc_rd1272_3 <> "") or (eti_conc_rd1272_3 <> "") or (conc_rd1272_4 <> "") or (eti_conc_rd1272_4 <> "") or (conc_rd1272_5 <> "") or (eti_conc_rd1272_5 <> "") or (conc_rd1272_6 <> "") or (eti_conc_rd1272_6 <> "") or (conc_rd1272_7 <> "") or (eti_conc_rd1272_7 <> "") or (conc_rd1272_8 <> "") or (eti_conc_rd1272_8 <> "") or (conc_rd1272_9 <> "") or (eti_conc_rd1272_9 <> "") or (conc_rd1272_10 <> "") or (eti_conc_rd1272_10 <> "") or (conc_rd1272_11 <> "") or (eti_conc_rd1272_11 <> "") or (conc_rd1272_12 <> "") or (eti_conc_rd1272_12 <> "") or (conc_rd1272_13 <> "") or (eti_conc_rd1272_13 <> "") or (conc_rd1272_14 <> "") or (eti_conc_rd1272_14 <> "") or (conc_rd1272_15 <> "") or (eti_conc_rd1272_15 <> "")) then

%>
	<span id="ap2_clasificacion_etiquetado_titulo" class="ficha_titulo_2"><a onclick=window.open('ver_definicion.asp?id=279','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> Etiquetado <% plegador "secc-etiquetado_rd1272", "img-etiquetado-rd1272" %></span>


  <fieldset id="secc-etiquetado_rd1272" style="display:none; margin: 15px 45px;">
<%
	if (conc_rd1272_1+conc_rd1272_2)<>"" then
		if (conc_rd1272_1)="" then
			if eti_conc_rd1272_1<>"" then
%>
			Factor <%= eti_conc_rd1272_1 %>
<%
			end if
		end if

%>
	<table cellspacing="0" cellpadding="3" width="100%" align="center">
		<tr>
			<th class="subtitulo3 celdaabajo">Concentración</th><th class="subtitulo3 celdaabajo">Etiquetado</th>
		</tr>
<%
	ap2_clasificacion_etiquetado_fila	"h", conc_rd1272_1, eti_conc_rd1272_1
	ap2_clasificacion_etiquetado_fila	"h", conc_rd1272_2, eti_conc_rd1272_2
	ap2_clasificacion_etiquetado_fila	"h", conc_rd1272_3, eti_conc_rd1272_3
	ap2_clasificacion_etiquetado_fila	"h", conc_rd1272_4, eti_conc_rd1272_4
	ap2_clasificacion_etiquetado_fila	"h", conc_rd1272_5, eti_conc_rd1272_5
	ap2_clasificacion_etiquetado_fila	"h", conc_rd1272_6, eti_conc_rd1272_6
	ap2_clasificacion_etiquetado_fila	"h", conc_rd1272_7, eti_conc_rd1272_7
	ap2_clasificacion_etiquetado_fila	"h", conc_rd1272_8, eti_conc_rd1272_8
	ap2_clasificacion_etiquetado_fila	"h", conc_rd1272_9, eti_conc_rd1272_9
	ap2_clasificacion_etiquetado_fila	"h", conc_rd1272_10, eti_conc_rd1272_10
	ap2_clasificacion_etiquetado_fila	"h", conc_rd1272_11, eti_conc_rd1272_11
	ap2_clasificacion_etiquetado_fila	"h", conc_rd1272_12, eti_conc_rd1272_12
	ap2_clasificacion_etiquetado_fila	"h", conc_rd1272_13, eti_conc_rd1272_13
	ap2_clasificacion_etiquetado_fila	"h", conc_rd1272_14, eti_conc_rd1272_14
	ap2_clasificacion_etiquetado_fila	"h", conc_rd1272_15, eti_conc_rd1272_15
%>
	</table>
<%
	else
		if eti_conc_rd1272_1<>"" then
%>
			Factor <%= eti_conc_rd1272_1 %>
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
	if ((estado_1 <> "") or (vla_ed_ppm_1 <> "") or (vla_ed_mg_m3_1 <> "") or (vla_ec_ppm_1 <> "") or (vla_ec_mg_m3_1 <> "") or (notas_vla_1 <> "") or (estado_2 <> "") or (vla_ed_ppm_2 <> "") or (vla_ed_mg_m3_2 <> "") or (vla_ec_ppm_2 <> "") or (vla_ec_mg_m3_2 <> "") or (notas_vla_2 <> "") or (estado_3 <> "") or (vla_ed_ppm_3 <> "") or (vla_ed_mg_m3_3 <> "") or (vla_ec_ppm_3 <> "") or (vla_ec_mg_m3_3 <> "") or (notas_vla_3 <> "") or (estado_4 <> "") or (vla_ed_ppm_4 <> "") or (vla_ed_mg_m3_4 <> "") or (vla_ec_ppm_4 <> "") or (vla_ec_mg_m3_4 <> "") or (notas_vla_4 <> "") or (estado_5 <> "") or (vla_ed_ppm_5 <> "") or (vla_ed_mg_m3_5 <> "") or (vla_ec_ppm_5 <> "") or (vla_ec_mg_m3_5 <> "") or (notas_vla_5 <> "") or (estado_6 <> "") or (vla_ed_ppm_6 <> "") or (vla_ed_mg_m3_6 <> "") or (vla_ec_ppm_6  <> "") or (vla_ec_mg_m3_6 <> "") or (notas_vla_6 <> "") or (ib_1 <> "") or  (vlb_1 <> "") or (momento_1 <> "") or (notas_vlb_1 <> "") or (ib_2 <> "") or  (vlb_2 <> "") or (momento_2 <> "") or (notas_vlb_2 <> "") or (ib_3 <> "") or  (vlb_3 <> "") or (momento_3 <> "") or (notas_vlb_3 <> "") or (ib_4 <> "") or  (vlb_4 <> "") or (momento_4 <> "") or (notas_vlb_4 <> "") or (ib_5 <> "") or  (vlb_5 <> "") or (momento_5 <> "") or (notas_vlb_5 <> "") or (ib_6 <> "") or  (vlb_6 <> "") or (momento_6 <> "") or (notas_vlb_6 <> "")) then

%>

	<table id="tabla_valores_limite" class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
  	<tr>
		<td class="celdaabajo" colspan="2" align="center">
			<table cellpadding=0 cellspacing=0 width="100%" border="0"><tr><td width="100%" class="titulo3" align="left"> VALORES L&Iacute;MITE DE EXPOSICI&Oacute;N PROFESIONAL
<!--			<a href="javascript:toggle('secc-mas_valores_limite', 'img-mas_valores_limite');"><img src="imagenes/desplegar.gif" align="absmiddle" id="img-mas_valores_limite" alt="Pulse para desplegar la información" title="Pulse para desplegar la información" /></a>-->
			</td></tr></table>
		</td>
	</tr>

		<tr>
			<td valign="top" width="50%">
			<!-- VLA -->

<%
		' VLA
		if ((estado_1 <> "") or (vla_ed_ppm_1 <> "") or (vla_ed_mg_m3_1 <> "") or (vla_ec_ppm_1 <> "") or (vla_ec_mg_m3_1 <> "") or (notas_vla_1 <> "") or (estado_2 <> "") or (vla_ed_ppm_2 <> "") or (vla_ed_mg_m3_2 <> "") or (vla_ec_ppm_2 <> "") or (vla_ec_mg_m3_2 <> "") or (notas_vla_2 <> "") or (estado_3 <> "") or (vla_ed_ppm_3 <> "") or (vla_ed_mg_m3_3 <> "") or (vla_ec_ppm_3 <> "") or (vla_ec_mg_m3_3 <> "") or (notas_vla_3 <> "") or (estado_4 <> "") or (vla_ed_ppm_4 <> "") or (vla_ed_mg_m3_4 <> "") or (vla_ec_ppm_4 <> "") or (vla_ec_mg_m3_4 <> "") or (notas_vla_4 <> "") or (estado_5 <> "") or (vla_ed_ppm_5 <> "") or (vla_ed_mg_m3_5 <> "") or (vla_ec_ppm_5 <> "") or (vla_ec_mg_m3_5 <> "") or (notas_vla_5 <> "") or (estado_6 <> "") or (vla_ed_ppm_6 <> "") or (vla_ed_mg_m3_6 <> "") or (vla_ec_ppm_6  <> "") or (vla_ec_mg_m3_6 <> "") or (notas_vla_6 <> "")) then
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

sub ap2_clasificacion_lista_negra()
	' Muestra el etiquetado

	if (esta_en_lista_cancer_rd or esta_en_lista_cancer_danesa or esta_en_lista_cancer_iarc_excepto_grupo_3 or esta_en_lista_cancer_otras_excepto_grupo_4 or esta_en_lista_de or (esta_en_lista_neurotoxico and (instr(frases_r,"R67")=0)) or  esta_en_lista_tpb or esta_en_lista_sensibilizante or esta_en_lista_sensibilizante_danesa or esta_en_lista_sensibilizante_reach or esta_en_lista_tpr or esta_en_lista_tpr_danesa or esta_en_lista_mutageno_rd or esta_en_lista_mutageno_danesa or esta_en_lista_cancer_mama or esta_en_lista_cop) or (instr(frases_r,"R33")<>0) or (instr(frases_r,"R53")<>0) or (instr(frases_r,"R50-53")<>0) or (instr(frases_r,"R51-53")<>0) or (instr(frases_r,"R52-53")<>0) or (instr(frases_r,"R58")<>0) then

    ' Esta en lista negra. Aprovechamos para marcarle el bit correspondiente para que aparezca en el listado de lista negra
    sqlListaNegra="UPDATE dn_risc_sustancias SET negra=1 WHERE id="&id_sustancia
    objConnection2.execute(sqlListaNegra),,adexecutenorecords

    ' OK, continuamos...

		razones = ""

		if (esta_en_lista_cancer_rd or esta_en_lista_cancer_danesa or esta_en_lista_cancer_iarc_excepto_grupo_3 or esta_en_lista_cancer_otras_excepto_grupo_4 or esta_en_lista_cancer_mama) then
			if (razones = "") then
				razones = "cancerígena"
			else
				razones = razones & ", cancerígena"
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
				razones = "mutágena"
			else
				razones = razones & ", mutágena"
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
				razones = "neurotóxica"
			else
				razones = razones & ", neurotóxica"
			end if
		end if

		if (esta_en_lista_sensibilizante or esta_en_lista_sensibilizante_danesa or esta_en_lista_sensibilizante_reach) then
			if (razones = "") then
				razones = "sensibilizante"
			else
				razones = razones & ", sensibilizante"
			end if

		end if

		if (esta_en_lista_tpr or esta_en_lista_tpr_danesa) then
			if (razones = "") then
				razones = "tóxica para la reproducción"
			else
				razones = razones & ", tóxica para la reproducción"
			end if
		end if

		if (instr(frases_r,"R33")<>0) then
			if (razones = "") then
				razones = "bioacumulativa"
			else
				razones = razones & ", bioacumulativa"
			end if
		end if

		if (instr(frases_r,"R58")<>0) then
			if (razones = "") then
				razones = "Puede provocar a largo plazo efectos negativos en el medio ambiente"
			else
				razones = razones & ", puede provocar a largo plazo efectos negativos en el medio ambiente"
			end if
		end if

		if (esta_en_lista_tpb) then
			if (razones = "") then
				razones = "tóxica, persistente y bioacumulativa"
			else
				razones = razones & ", tóxica, persistente y bioacumulativa"
			end if
		end if

		' SPL (16/06/20014)
'		if num_cas="87-68-3" or num_cas="133-49-3" or num_cas="75-74-1" then
		if esta_en_lista_mpmb then
			if (razones = "") then
				razones = "Muy persistente y muy bioacumulativa"
			else
				razones = razones & ", muy persistente y muy bioacumulativa"
			end if
		end if

		if (instr(frases_r,"R53")<>0) or (instr(frases_r,"R50-53")<>0) or (instr(frases_r,"R51-53")<>0) or (instr(frases_r,"R52-53")<>0) then
			if (razones = "") then
				razones = "Puede provocar a largo plazo efectos negativos en el medio ambiente acuático"
			else
				razones = razones & ", puede provocar a largo plazo efectos negativos en el medio ambiente acuático"
			end if
		end if

%>
		<p id="ap2_clasificacion_lista_negra_titulo" class="subtitulo3">&nbsp;<img src="imagenes/icono_atencion_20.png" align="absmiddle" /> <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("Lista negra")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a> Sustancia incluida en la Lista negra de ISTAS <% plegador "secc-listanegra", "img-listanegra" %></p>
		<p id="secc-listanegra" class="texto" style="display:none">Esta sustancia está incluida en la Lista negra de ISTAS por los siguientes motivos: <%=razones%></p>

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


	if (esta_en_lista_cancer_rd or esta_en_lista_cancer_danesa or esta_en_lista_cancer_iarc or esta_en_lista_cancer_otras or esta_en_lista_cancer_mama or esta_en_lista_de or esta_en_lista_neurotoxico or efecto_neurotoxico="OTOTÓXICO" or esta_en_lista_sensibilizante or esta_en_lista_sensibilizante_reach or esta_en_lista_sensibilizante_danesa or esta_en_lista_tpr or esta_en_lista_tpr_danesa or esta_en_lista_eepp or esta_en_lista_mutageno_rd or esta_en_lista_mutageno_danesa or esta_en_lista_salud or esta_en_lista_prohibidas_embarazadas or esta_en_lista_prohibidas_lactantes or comentarios_sl <> "") then
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
		if (esta_en_lista_cancer_rd or esta_en_lista_cancer_danesa or esta_en_lista_cancer_iarc or esta_en_lista_cancer_otras or esta_en_lista_cancer_mama) then ap3_riesgos_tabla("Cancerígeno") end if
		'response.write esta_en_lista_mutageno_rd & esta_en_lista_mutageno_danesa
		if (esta_en_lista_mutageno_rd or esta_en_lista_mutageno_danesa ) then ap3_riesgos_tabla("Mutágeno") end if

		if esta_en_lista_de then ap3_riesgos_tabla("Disruptor endocrino") end if
		if esta_en_lista_neurotoxico or efecto_neurotoxico="OTOTÓXICO" then ap3_riesgos_tabla("Neurotóxico") end if
		if esta_en_lista_sensibilizante or esta_en_lista_sensibilizante_danesa or esta_en_lista_sensibilizante_reach then ap3_riesgos_tabla("Sensibilizante") end if
		'if esta_en_lista_sensibilizante_reach then ap3_riesgos_tabla("Sensibilizante para REACH") end if
		if esta_en_lista_tpr or esta_en_lista_tpr_danesa then ap3_riesgos_tabla("Tóxico para la reproducción") end if
		if esta_en_lista_eepp then ap3_riesgos_enfermedades() end if
    	if esta_en_lista_salud then ap7_salud() end if
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
if (esta_en_lista_tpb or esta_en_lista_directiva_aguas or esta_en_lista_alemana or esta_en_lista_sustancias_prioritarias  or esta_en_lista_ozono or esta_en_lista_clima or esta_en_lista_aire or esta_en_lista_cop or comentarios_medio_ambiente <>"" or esta_en_lista_suelos) then %>

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
		if esta_en_lista_tpb then
			ap3_riesgos_tabla("Tóxica, Persistente y Bioacumulativa")
		end if
		' SPL (16/06/20014)
'		if num_cas="87-68-3" or num_cas="133-49-3" or num_cas="75-74-1" then
		if esta_en_lista_mpmb then
			ap3_riesgos_tabla("mPmB")
		end if
		if (esta_en_lista_directiva_aguas or esta_en_lista_alemana) then ap3_riesgos_tabla("Tóxica para el agua") end if
		if (esta_en_lista_suelos) then ap3_riesgos_tabla("Contaminante de suelos") end if
		if (esta_en_lista_ozono or esta_en_lista_clima or esta_en_lista_aire) then ap3_riesgos_tabla("Contaminante del aire") end if

		if (esta_en_lista_cop) then ap3_riesgos_tabla("Contaminante Orgánico Persistente (COP)") end if
%>

		<%
		if (comentarios_medio_ambiente <>"") then
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
							<%=comentarios_medio_ambiente %>
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
			<a href="index.asp?idpagina=1194"><img src="../imagenes/ayuda.gif" align="absmiddle" border="0" /></a>
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
		<%
	  	if (trim(enlace_cop) <> "") then
			response.write "<li><a href='"&enlace_cop&"' target='_blank'>Más información</a></li>"
		end if
	  %>

      </ul>


    </fieldset>

<%
		case "Cancerígeno":

				' Real Decreto ---------------------------------------------------------------
				if (esta_en_lista_cancer_rd) then
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
				if (trim(notas_cancer_rd) <> "") then
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
					if (trim(notas_cancer_rd) <> "") then
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
						<legend class="subtitulo3"><strong>Según IARC <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("IARC")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a></strong></legend>
		<%
						if (trim(grupo_iarc) <> "") or (trim(volumen_iarc) <> "") or (trim(notas_iarc) <> "") then
		%>
							<blockquote>
							<table>
		<%
							if (trim(grupo_iarc) <> "") then
		%>
								<tr><td class="subtitulo3">Grupo:</td><td><%=trim(replace(ucase(grupo_iarc), "GRUPO", ""))%> <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion(trim(grupo_iarc))%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a></td></tr>
		<%
							end if

							if (trim(volumen_iarc) <> "") then
		%>
								<tr><td class="subtitulo3">Volumen:</td><td><%=volumen_iarc%></td></tr>
		<%
							end if
							if (trim(notas_iarc) <> "") then
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
				  <legend class="subtitulo3"><strong>Según otras fuentes</strong></legend>

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
						<legend class="subtitulo3"><strong>Según <%=trim(array_fuentes(i))%> <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion(trim(array_fuentes(i)))%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a></strong></legend>
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
						<legend class="subtitulo3"><strong>Según SSI (cáncer de mama) <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("SSI")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a></strong></legend>
						<blockquote>
						<table>
							<tr><td class="subtitulo3"><strong>Fuente:</strong><br /><a href="<%= cancer_mama_fuente %>" target="_blank"><%= replace(cancer_mama_fuente, "http://", "") %></a></td></tr>
						</table>
						</blockquote>
					</fieldset>

		<%

		    end if

		case "Mutágeno":
      ' MUTAGENO RD -------------------------------------------------------------
      if (esta_en_lista_mutageno_rd) then
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
      if (esta_en_lista_mutageno_danesa) then
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

        'response.write efecto_neurotoxico&"***"&fuente_neurotoxico

        if esta_en_lista_neurotoxico_rd or esta_en_lista_neurotoxico_danesa then
          ' Añadimos SNC a efecto neurotoxico si no existía ya
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

						<%= efecto %> <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion(efecto)%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>

						<%
							next
						%>
					</td>
				</tr>
			<% end if %>
			<% if (nivel_neurotoxico <> "") then %>
				<tr>
					<td class="subtitulo3" valign="top">Nivel:</td><td><%=nivel_neurotoxico%>

					 <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("Nivel "&nivel_neurotoxico)%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>

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
						'response.write trim(array_fuentes(i))
          				response.write dame_definicion(trim(array_fuentes(i)))

			'trim(array_fuentes(i))


			%>
            <!--
			 <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion(trim(array_fuentes(i)))%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>
			 -->
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
		      if esta_en_lista_sensibilizante then
		        response.write "<li class='subtitulo3'>Sensibilizante según R. 1272/2008</li>"
		      end if

			  if esta_en_lista_sensibilizante_reach then
		        response.write "<li class='subtitulo3'>Alérgeno REACH &nbsp;<a href='http://www.istas.net/web/abreenlace.asp?idenlace=6340' target='_blank'>Ver documento</a></li>"
		      end if

		      if esta_en_lista_sensibilizante_danesa then
		      %>
		        <li class='subtitulo3'>Sensibilizante según <% plegador_texto "frases_r_danesa_sensibilizante", "frases R", "subtitulo3" %> en la clasificación de la EPA danesa <a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion("EPA Danesa")%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a></li>
		      <%


		      %>
		        <div id="frases_r_danesa_sensibilizante" style="display:none"><br />
		        <blockquote>
		        <% ap2_clasificacion_frases_r_danesa() %>
		        </blockquote>
		        </div>
		      <%
			  end if
			  response.write "</ul>"


		case "Tóxico para la reproducción":
	      ' TPR SEGUN RD -------------------------------------------------------------
	      if (esta_en_lista_tpr) then
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
	      if (esta_en_lista_tpr_danesa) then
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
	        <% ap2_clasificacion_frases_r_danesa() %>
	        </blockquote>
	        </div>
	          </fieldset>
	<%
	      end if

	case "Prohibida para trabajadoras embarazadas":

      if (esta_en_lista_prohibidas_embarazadas) then
%>
  				<blockquote>
  					<strong>Fuente:</strong> Real Decreto 298/2009
				</blockquote>
<%
      end if

	case "Prohibida para trabajadoras lactantes":

      if (esta_en_lista_prohibidas_lactantes) then
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
					<td><a href="<%= enlace_tpb %>"><%= corta(anchor_tpb, 70, "puntossuspensivos") %></a></td>
				</tr>
				<tr>
					<td class="subtitulo3" valign="top">Fuente/s:</td>
					<td class="subtitulo3"><%
						if trim(fuentes_tpb) <> "" then
							array_tpb = split(fuentes_tpb,",")
							for i=0 to ubound(array_tpb)
								response.write "<li>"&dame_definicion(trim(array_tpb(i)))&"</li>"
							next
						end if
						if trim(fuente_tpb) <> "" then
							array_tpb = split(fuente_tpb,",")
							for i=0 to ubound(array_tpb)
								' response.write "<li>"&c&"</li>"
								response.write "<li>"&dame_definicion(trim(array_tpb(i)))&"</li>"
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
					<td class="subtitulo3"><%=dame_definicion("REACH")%></td>

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
			if (directiva_aguas or esta_en_lista_directiva_aguas) then
%>
				<tr>
					<td class="subtitulo3" colspan="2">· Según <a href="http://www.istas.net/web/abreenlace.asp?idenlace=2227" target="_blank">Directiva de aguas</a>, y sus posteriores <a href="http://www.istas.net/web/abreenlace.asp?idenlace=6323">modificaciones</a></td>
				</tr>
<%
			end if

			if (esta_en_lista_sustancias_prioritarias) then
%>
				<tr>
					<td class="subtitulo3" colspan="2">· Posible sustancia prioritaria según la <a href="http://www.istas.net/web/abreenlace.asp?idenlace=2227" target="_blank">Directiva de aguas</a>, y sus posteriores <a href="http://www.istas.net/web/abreenlace.asp?idenlace=6323" target="_blank">modificaciones</a></td>
				</tr>
<%
			end if

			if (trim(clasif_mma) <> "") and (trim(clasif_mma)<>"nwg")then
%>
				<tr>
					<td class="subtitulo3" colspan="2">
						· Según <a href="http://www.istas.net/risctox/abreenlace.asp?idenlace=2226" target="_blank">Ministerio de Medio Ambiente de Alemania</a>
					</td>
				</tr>
				<tr>
					<td>&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<td><strong>Clasificación</strong>: <%=clasif_mma%>
					 	<a onclick=window.open('ver_definicion.asp?id=<%=dame_id_definicion(parche_definicion(clasif_mma, "MMA"))%>','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:pointer'><img src='imagenes/ayuda.gif' width=14 height=14 align='absmiddle' border='0' /></a>
					</td>
				</tr>
<%
			end if
			if (sustancia_prioritaria=1)then
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
				if (dano_calidad_aire or esta_en_lista_aire) then
%>
					<tr>
						<td class="subtitulo3">Calidad del aire:</td>
						<td>Sustancia incluida en la <a href="abreenlace.asp?idenlace=2234" target="_blank">Directiva 96/62/CE</a> de 27 de septiembre sobre evaluación y gestión de la calidad del aire ambiente</td>
					</tr>
<%
				end if
%>
<%
				if (dano_ozono) then
%>
					<tr>
						<td class="subtitulo3">Capa de ozono:</td>
						<td>Sustancia que agota la capa de ozono, según <a href="abreenlace.asp?idenlace=2229" target="_blank">Reglamento (CE) 2037/2000</a> del Parlamento Europeo y del Consejo, de 29 de junio de 2000</td>
					</tr>
<%
				end if
%>
<%
				if (dano_cambio_clima) then
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
                    	Fuente: <a href="http://echa.europa.eu/chem_data/authorisation_process/candidate_list_table_en.asp" target="_blank">Agencia Europea de sustancias y mezclas químicas (ECHA)</a>
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
                    	Fuente: <a href="http://echa.europa.eu/reach/authorisation_under_reach/authorisation_list_en.asp" target="_blank">Agencia Europea de sustancias y mezclas químicas (ECHA)</a>
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
	'response.write "<br />"&sql_enf
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
	if esta_en_lista_cov or esta_en_lista_residuos or esta_en_lista_vertidos or esta_en_lista_lpcic  or esta_en_lista_accidentes or esta_en_lista_emisiones then
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

if esta_en_lista_cov then total = total +1 end if
if esta_en_lista_vertidos then total = total +1 end if
if esta_en_lista_lpcic_agua then total = total +1 end if
if esta_en_lista_lpcic_aire then total = total +1 end if
if esta_en_lista_lpcic_suelo then total = total +1 end if
if esta_en_lista_residuos then total = total +1 end if
if esta_en_lista_accidentes then total = total +1 end if
if esta_en_lista_emisiones then total = total +1 end if
'if esta_en_lista_prohibidas then total = total + 1
'if esta_en_lista_restringidas then total = total + 1

'response.write total

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
			ap3_riesgos_tabla("IPPC (PRTR Agua)")
			llevo = llevo +1
			if llevo >= mitad then
				response.write "</td><td valign='top' width='50%'>"
				llevo = 0 ' Lo reseteo para que no vuelva a dividir
			end if
		end if

		if esta_en_lista_lpcic_aire then
			ap3_riesgos_tabla("IPPC (PRTR Aire)")
			llevo = llevo +1
			if llevo >= mitad then
				response.write "</td><td valign='top' width='50%'>"
				llevo = 0 ' Lo reseteo para que no vuelva a dividir
			end if
		end if

		if esta_en_lista_lpcic_suelo then
			ap3_riesgos_tabla("IPPC (PRTR Suelo)")
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
	if esta_en_lista_prohibidas or esta_en_lista_restringidas or esta_en_lista_candidatas_reach or esta_en_lista_autorizacion_reach or esta_en_lista_biocidas_autorizadas or esta_en_lista_biocidas_prohibidas or esta_en_lista_pesticidas_autorizadas or esta_en_lista_pesticidas_prohibidas or esta_en_lista_prohibidas_embarazadas or esta_en_lista_prohibidas_lactantes or esta_en_lista_corap then
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
		if esta_en_lista_prohibidas then
			ap3_riesgos_tabla("Sustancia prohibida")
		end if

		if esta_en_lista_restringidas then
			ap3_riesgos_tabla("Sustancia restringida")
		end if

		if esta_en_lista_prohibidas_embarazadas then ap3_riesgos_tabla("Prohibida para trabajadoras embarazadas") end if

		if esta_en_lista_prohibidas_lactantes then ap3_riesgos_tabla("Prohibida para trabajadoras lactantes") end if

		if esta_en_lista_candidatas_reach then
			ap3_riesgos_tabla("Sustancia candidata REACH")
		end if
		if esta_en_lista_autorizacion_reach then
			ap3_riesgos_tabla("Sustancia REACH sujeta a autorización")
		end if
		if esta_en_lista_biocidas_autorizadas then
			ap3_riesgos_tabla("Sustancia biocida autorizada")
		end if
		if esta_en_lista_biocidas_prohibidas then
			ap3_riesgos_tabla("Sustancia biocida prohibida")
		end if
		if esta_en_lista_pesticidas_autorizadas then
			ap3_riesgos_tabla("Sustancia pesticida autorizada")
		end if
		if esta_en_lista_pesticidas_prohibidas then
			ap3_riesgos_tabla("Sustancia pesticida prohibida")
		end if
		if esta_en_lista_corap then
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

  'response.write sql

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
	clasificacion_rd = clasificacion_1 & clasificacion_2 & clasificacion_3 & clasificacion_4 & clasificacion_5 & clasificacion_6 & clasificacion_7 & clasificacion_8 & clasificacion_9 & clasificacion_10 & clasificacion_11 & clasificacion_12 & clasificacion_13 & clasificacion_14 & clasificacion_15

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
	posicion = instr(1,frases_r_danesa, "Carc")

	' Sacamos el nivel como el caracter que hay justo detrás de la primera aparición de la subcadena
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
	posicion = instr(1,frases_r_danesa, "Mut")

	' Sacamos el nivel como el caracter que hay justo detrás de la primera aparición de la subcadena
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



sub evaluaCamposListaAsociada(lista,camposArray())
	dim c, q, x
	if objRst("asoc_"&lista) then
		execute("esta_en_lista_"&lista&"=1")
		for i = 0 to UBound(camposArray)
			c = camposArray(i)
			execute( "q= " & c )
			x = objRst( "asoc_" & lista & "_" & c )
			if inStr(q, x) = 0 then execute(c&" = "&c& "& "", " & objRst("asoc_"&lista&"_"&c) & """")
		next
	else
		execute("esta_en_lista_"&lista&"=0")
	end if
end sub

%>


