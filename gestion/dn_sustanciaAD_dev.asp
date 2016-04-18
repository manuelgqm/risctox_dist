<!--#include file="dn_fun_comunes.asp"-->
<!--#include file="dn_auten.inc"-->

<%
id=EliminaInyeccionSQL(request("id"))
if id="" then
%>
<script>window.close();</script>
<%
else
%>
	<!--#include file="adovbs.inc"-->
	<!--#include file="dn_conexion.asp"-->
<%
	'CONSULTAMOS DATOS DE LAS DISTINTAS TABLAS A MOSTRAR

	't1
	sql3="select * from dn_risc_sustancias_vl where id_sustancia=" &id
	set objRst3=objconn1.execute(sql3)
	if objRst3.eof then
		hayt1=0
	else
		hayt1=1

		estado_1=objRst3("estado_1")
  		vla_ed_ppm_1=objRst3("vla_ed_ppm_1")
  		vla_ed_mg_m3_1=objRst3("vla_ed_mg_m3_1")
	  	vla_ec_ppm_1=objRst3("vla_ec_ppm_1")
  		vla_ec_mg_m3_1=objRst3("vla_ec_mg_m3_1")
		notas_vla_1=objRst3("notas_vla_1")

		ib_1=objRst3("ib_1")
		vlb_1=objRst3("vlb_1")
		momento_1=objRst3("momento_1")
		notas_vlb_1=objRst3("notas_vlb_1")

				estado_2=objRst3("estado_2")
  		vla_ed_ppm_2=objRst3("vla_ed_ppm_2")
  		vla_ed_mg_m3_2=objRst3("vla_ed_mg_m3_2")
	  	vla_ec_ppm_2=objRst3("vla_ec_ppm_2")
  		vla_ec_mg_m3_2=objRst3("vla_ec_mg_m3_2")
		notas_vla_2=objRst3("notas_vla_2")

		ib_2=objRst3("ib_2")
		vlb_2=objRst3("vlb_2")
		momento_2=objRst3("momento_2")
		notas_vlb_2=objRst3("notas_vlb_2")

				estado_3=objRst3("estado_3")
  		vla_ed_ppm_3=objRst3("vla_ed_ppm_3")
  		vla_ed_mg_m3_3=objRst3("vla_ed_mg_m3_3")
	  	vla_ec_ppm_3=objRst3("vla_ec_ppm_3")
  		vla_ec_mg_m3_3=objRst3("vla_ec_mg_m3_3")
		notas_vla_3=objRst3("notas_vla_3")

		ib_3=objRst3("ib_3")
		vlb_3=objRst3("vlb_3")
		momento_3=objRst3("momento_3")
		notas_vlb_3=objRst3("notas_vlb_3")

				estado_4=objRst3("estado_4")
  		vla_ed_ppm_4=objRst3("vla_ed_ppm_4")
  		vla_ed_mg_m3_4=objRst3("vla_ed_mg_m3_4")
	  	vla_ec_ppm_4=objRst3("vla_ec_ppm_4")
  		vla_ec_mg_m3_4=objRst3("vla_ec_mg_m3_4")
		notas_vla_4=objRst3("notas_vla_4")

		ib_4=objRst3("ib_4")
		vlb_4=objRst3("vlb_4")
		momento_4=objRst3("momento_4")
		notas_vlb_4=objRst3("notas_vlb_4")

				estado_5=objRst3("estado_5")
  		vla_ed_ppm_5=objRst3("vla_ed_ppm_5")
  		vla_ed_mg_m3_5=objRst3("vla_ed_mg_m3_5")
	  	vla_ec_ppm_5=objRst3("vla_ec_ppm_5")
  		vla_ec_mg_m3_5=objRst3("vla_ec_mg_m3_5")
		notas_vla_5=objRst3("notas_vla_5")

		ib_5=objRst3("ib_5")
		vlb_5=objRst3("vlb_5")
		momento_5=objRst3("momento_5")
		notas_vlb_5=objRst3("notas_vlb_5")

				estado_6=objRst3("estado_6")
  		vla_ed_ppm_6=objRst3("vla_ed_ppm_6")
  		vla_ed_mg_m3_6=objRst3("vla_ed_mg_m3_6")
	  	vla_ec_ppm_6=objRst3("vla_ec_ppm_6")
  		vla_ec_mg_m3_6=objRst3("vla_ec_mg_m3_6")
		notas_vla_6=objRst3("notas_vla_6")

		ib_6=objRst3("ib_6")
		vlb_6=objRst3("vlb_6")
		momento_6=objRst3("momento_6")
		notas_vlb_6=objRst3("notas_vlb_6")

	end if
	objRst3.close
	set objRst3=nothing

	't0: dn_risc_sustancias
	sql3="select notas_cancer_rd from dn_risc_sustancias where id=" &id
	set objRst3=objconn1.execute(sql3)
	hayt0=1 'esta fila siempre existe
	notas_cancer_rd=objRst3("notas_cancer_rd")
	objRst3.close
	set objRst3=nothing

	't2
	sql3="select * from dn_risc_sustancias_iarc where id_sustancia=" &id
	set objRst3=objconn1.execute(sql3)
	if objRst3.eof then
		hayt2=0
	else
		hayt2=1

		grupo_iarc=objRst3("grupo_iarc")
  		volumen_iarc=objRst3("volumen_iarc")
  		notas_iarc=objRst3("notas_iarc")
  		notas_iarc_ing=objRst3("notas_iarc_ing")

	end if
	objRst3.close
	set objRst3=nothing

	't3
	sql3="select * from dn_risc_sustancias_cancer_otras where id_sustancia=" &id
	set objRst3=objconn1.execute(sql3)
	if objRst3.eof then
		hayt3=0
	else
		hayt3=1

		categoria_cancer_otras=objRst3("categoria_cancer_otras")
  		fuente=objRst3("fuente")

	end if
	objRst3.close
	set objRst3=nothing


  ' ##########################
	't9: CANCER MAMA
	sql9="select * from dn_risc_sustancias_mama_cop where id_sustancia=" &id
	set objRst9=objconn1.execute(sql9)
	if objRst9.eof then
		hayt9=0
	else
		hayt9=1

		cancer_mama=objRst9("cancer_mama")
		cancer_mama_fuente=objRst9("cancer_mama_fuente")

	end if
	objRst9.close
	set objRst9=nothing
  ' ##########################




	't8: COP
	sql8="select * from dn_risc_sustancias_mama_cop where id_sustancia=" &id
	set objRst8=objconn1.execute(sql8)
	if objRst8.eof then
		hayt8=0
	else
		hayt8=1

		cop=objRst8("cop")

	end if
	objRst8.close
	set objRst8=nothing
  ' ##########################


	't4
	sql3="select * from dn_risc_sustancias_neuro_disruptor where id_sustancia=" &id
	set objRst3=objconn1.execute(sql3)
	if objRst3.eof then
		hayt4=0
	else
		hayt4=1

		efecto_neurotoxico=objRst3("efecto_neurotoxico")
  		nivel_neurotoxico=objRst3("nivel_neurotoxico")
		fuente_neurotoxico=objRst3("fuente_neurotoxico")
  		nivel_disruptor=objRst3("nivel_disruptor")

	end if

	't5
	sql3="select * from dn_risc_sustancias_ambiente where id_sustancia=" &id
	set objRst3=objconn1.execute(sql3)
	if objRst3.eof then
		hayt5=0
	else
		hayt5=1

		enlace_tpb=objRst3("enlace_tpb")
		anchor_tpb=objRst3("anchor_tpb")
		dano_calidad_aire=objRst3("dano_calidad_aire")
		dano_ozono=objRst3("dano_ozono")
		dano_cambio_clima=objRst3("dano_cambio_clima")
		directiva_aguas=objRst3("directiva_aguas")
		clasif_mma=objRst3("clasif_mma")
		emisiones_atmosfera=objRst3("emisiones_atmosfera")
		cov=objRst3("cov")
		seveso=objRst3("seveso")
		eper_agua=objRst3("eper_agua")
		eper_aire=objRst3("eper_aire")
		eper_suelo=objRst3("eper_suelo")


		'Sergio
		fuentes_tpb = objrst3("fuentes_tpb")
		am_comentarios = objrst3("comentarios")
		am_comentarios_ing = objrst3("comentarios_ing")
		toxicidad_suelo = objrst3("toxicidad_suelo")
		sustancia_prioritaria = objrst3("sustancia_prioritaria")

	end if


  't6: Efectos sobre la salud
  sql3="select * from dn_risc_sustancias_salud where id_sustancia=" &id
	set objRst3=objconn1.execute(sql3)
	if objRst3.eof then
		hayt6=0
	else
		hayt6=1

		salud_cardiocirculatorio=objRst3("cardiocirculatorio")
		salud_rinyon=objRst3("rinyon")
		salud_respiratorio=objRst3("respiratorio")
		salud_reproductivo=objRst3("reproductivo")
		salud_piel_sentidos=objRst3("piel_sentidos")
		salud_neuro_toxicos=objRst3("neuro_toxicos")
		salud_musculo_esqueletico=objRst3("musculo_esqueletico")
		salud_sistema_inmunitario=objRst3("sistema_inmunitario")
		salud_higado_gastrointestinal=objRst3("higado_gastrointestinal")
		salud_sistema_endocrino=objRst3("sistema_endocrino")
		salud_embrion=objRst3("embrion")
		salud_cancer=objRst3("cancer")
		salud_comentarios=objrst3("comentarios")
		salud_comentarios_ing=objrst3("comentarios_ing")

	end if


' *******  NUEVAS LISTAS SPL

  ' ##########################
	't10: PROHIBIDAS EMBARAZADAS
	sql10="select * from spl_risc_sustancias_prohibidas_embarazadas where id_sustancia=" &id
	set objRstLstSPL=objconn1.execute(sql10)
	if objRstLstSPL.eof then
		hayt10=0
	else
		hayt10=1
		t10_comentario_prohibida = objRstLstSPL("comentario_prohibida")
		t10_comentario_prohibida_ing = objRstLstSPL("comentario_prohibida_ing")

	end if
	objRstLstSPL.close
	set objRstLstSPL=nothing
  ' ##########################


  ' ##########################
	't11: PROHIBIDAS LACTANTES
	sql11="select * from spl_risc_sustancias_prohibidas_lactantes where id_sustancia=" &id
	set objRstLstSPL=objconn1.execute(sql11)
	if objRstLstSPL.eof then
		hayt11=0
	else
		hayt11=1
		t11_comentario_prohibida = objRstLstSPL("comentario_prohibida")
		t11_comentario_prohibida_ing = objRstLstSPL("comentario_prohibida_ing")
	end if
	objRstLstSPL.close
	set objRstLstSPL=nothing
  ' ##########################

  ' ##########################
	't12: PROHIBIDAS
	sqlLstSPL="select * from dn_risc_sustancias_prohibidas where id_sustancia=" &id
	set objRstLstSPL=objconn1.execute(sqlLstSPL)
	if objRstLstSPL.eof then
		hayt12=0
	else
		hayt12=1
		t12_comentario_prohibida = objRstLstSPL("comentario_prohibida")
		t12_comentario_prohibida_ing = objRstLstSPL("comentario_prohibida_ing")
	end if
	objRstLstSPL.close
	set objRstLstSPL=nothing
  ' ##########################

  ' ##########################
	't13: RESTRINGIDAS
	sqlLstSPL="select * from dn_risc_sustancias_restringidas where id_sustancia=" &id
	set objRstLstSPL=objconn1.execute(sqlLstSPL)
	if objRstLstSPL.eof then
		hayt13=0
	else
		hayt13=1
		t13_comentario_restringida = objRstLstSPL("comentario_restringida")
		t13_comentario_restringida_ing = objRstLstSPL("comentario_restringida_ing")
	end if
	objRstLstSPL.close
	set objRstLstSPL=nothing
  ' ##########################

  ' ##########################
	't14: Sustancias candidatas REACH
	sqlLstSPL="select * from spl_risc_sustancias_candidatas_reach where id_sustancia=" &id
	set objRstLstSPL=objconn1.execute(sqlLstSPL)
	if objRstLstSPL.eof then
		hayt14=0
	else
		hayt14=1
	end if
	objRstLstSPL.close
	set objRstLstSPL=nothing
  ' ##########################

  ' ##########################
	't15: Sustancias REACH sujetas a autorización
	sqlLstSPL="select * from spl_risc_sustancias_autorizacion_reach where id_sustancia=" &id
	set objRstLstSPL=objconn1.execute(sqlLstSPL)
	if objRstLstSPL.eof then
		hayt15=0
	else
		hayt15=1
	end if
	objRstLstSPL.close
	set objRstLstSPL=nothing
  ' ##########################

  ' ##########################
	't16: Sustancias Biocidas prohibidas
	sqlLstSPL="select * from spl_risc_sustancias_biocidas_prohibidas where id_sustancia=" &id
	set objRstLstSPL=objconn1.execute(sqlLstSPL)
	if objRstLstSPL.eof then
		hayt16=0
	else
		hayt16=1
		t16_fuente = objRstLstSPL("fuente")
		t16_fecha_limite = objRstLstSPL("fecha_limite")
		t16_usos = objRstLstSPL("usos")
		t16_usos_ing = objRstLstSPL("usos_ing")

	end if
	objRstLstSPL.close
	set objRstLstSPL=nothing
  ' ##########################

  ' ##########################
	't17: Sustancias Biocidas autorizadas
	sqlLstSPL="select * from spl_risc_sustancias_biocidas_autorizadas where id_sustancia=" &id
	set objRstLstSPL=objconn1.execute(sqlLstSPL)
	if objRstLstSPL.eof then
		hayt17=0
	else
		hayt17=1
		t17_fuente = objRstLstSPL("fuente")
		t17_pureza_minima = objRstLstSPL("pureza_minima")
		t17_condiciones = objRstLstSPL("condiciones")
		t17_condiciones_ing = objRstLstSPL("condiciones_ing")
		t17_usos = objRstLstSPL("usos")
		t17_usos_ing = objRstLstSPL("usos_ing")

	end if
	objRstLstSPL.close
	set objRstLstSPL=nothing
  ' ##########################

  ' ##########################
	't18: Sustancias Pesticidas autorizadas
	sqlLstSPL="select * from spl_risc_sustancias_pesticidas_autorizadas where id_sustancia=" &id
	set objRstLstSPL=objconn1.execute(sqlLstSPL)
	if objRstLstSPL.eof then
		hayt18=0
	else
		hayt18=1
		t18_fuente = objRstLstSPL("fuente")
		t18_plazo_renovacion = objRstLstSPL("plazo_renovacion")
		t18_plazo_renovacion_ing = objRstLstSPL("plazo_renovacion_ing")
		t18_pureza_minima = objRstLstSPL("pureza_minima")
		t18_pureza_minima_ing = objRstLstSPL("pureza_minima_ing")
		t18_usos = objRstLstSPL("usos")
		t18_usos_ing = objRstLstSPL("usos_ing")
	end if
	objRstLstSPL.close
	set objRstLstSPL=nothing
  ' ##########################

  ' ##########################
	't19: Sustancias Pesticidas prohibidas
	sqlLstSPL="select * from spl_risc_sustancias_pesticidas_prohibidas where id_sustancia=" &id
	set objRstLstSPL=objconn1.execute(sqlLstSPL)
	if objRstLstSPL.eof then
		hayt19=0
	else
		hayt19=1
		t19_fuente = objRstLstSPL("fuente")
		t19_exenciones = objRstLstSPL("exenciones")
		t19_exenciones_ing = objRstLstSPL("exenciones_ing")

	end if
	objRstLstSPL.close
	set objRstLstSPL=nothing
  ' ##########################


' ******* FIN NUEVAS LISTAS SPL

	'Sergio: Lo pongo bajo
	'cerrarconexion
end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Istas</title>
<link rel="stylesheet" type="text/css" href="dn_estilos.css">
<style type="text/css">
td {text-align:left}
table {text-align:center; margin:5px auto;}
</style>
<script type="text/javascript" src="niftycube.js"></script>
<script type="text/javascript">
window.onload=function(){
Nifty("div#box2","big");
}
</script>
<script language="JavaScript" src="gen_validatorv2.js" type="text/javascript"></script>
<script language="javascript">
	function cambia_pestanya(par){
		/*
		for (i=1; i<=4 ;i++){
			eval("document.getElementById('div"+i+"').style.visibility='hidden';");
			eval("document.getElementById('div"+i+"').style.display='none';");

		}
		*/
	    //Limpio


		eval("if(document.getElementById('div"+par+"').style.visibility == 'visible'){document.getElementById('div"+par+"').style.visibility='hidden';document.getElementById('div"+par+"').style.display='none';}else{document.getElementById('div"+par+"').style.visibility='visible';document.getElementById('div"+par+"').style.display='block';}");
		//eval("document.getElementById('div"+par+"').style.display='block';");
	}

</script>
</head>

<body>
<%flashMsgShow()%>
<div id="box2" class="centcontenido">
<form name="myform" action="dn_sustanciaAD2.asp?id=<%=id%>" method="post" >

<fieldset>
	<legend><strong><a href='javascript:cambia_pestanya(1)'>RIESGOS ESPECÍFICOS SOBRE LA SALUD</a></strong></legend>
	<div id='div1' style='display:block; visibility:visible'>
	<fieldset>
 		<legend><strong>Información sobre cáncer</strong></legend>
 		<fieldset>
 			<legend><strong>Real Decreto</strong></legend>
				 Notas cáncer RD: <input name="t0_notas_cancer_rd" type="text" value="<%=notas_cancer_rd%>" size="100" maxlength="100" />
 		</fieldset>
 		<fieldset>
 			<legend><strong>IARC</strong></legend><input type="hidden" name="hayt2" value="<%=hayt2%>" />
				 <table >
				 <tr><td>Grupo IARC</td><td><input name="t2_grupo_iarc" type="text" value="<%=grupo_iarc%>" size="100" maxlength="20" /></td></tr>
				 <tr><td>Volumen IARC</td><td><input name="t2_volumen_iarc" type="text" value="<%=volumen_iarc%>" size="100" maxlength="125" /></td></tr>
				 <tr><td>Notas IARC</td><td><textarea name="t2_notas_iarc" cols="75"><%=notas_iarc%></textarea></td></tr>
				 <tr><td>Notas IARC (ingl&eacute;s)</td><td><textarea name="t2_notas_iarc_ing" cols="75"><%=notas_iarc_ing%></textarea></td></tr>
				 </table>
 		</fieldset>

 		<fieldset>
			 <legend><strong>Otras fuentes</strong></legend><input type="hidden" name="hayt3" value="<%=hayt3%>" />
			 <table>
             <%
             	sql = "select id,palabra from rq_definiciones where fuente like '%CAF%'"
				 Set objRecordset = Server.CreateObject ("ADODB.Recordset")
				 Set objRecordset = objconn1.Execute(sql)

				 dim caf
				 if fuente <> "" then

					 caf = split(fuente,",")
				 end if

				 conta=0
				 while not objrecordset.eof
					tiene=""
					if fuente <> "" then
						for i=0 to ubound(caf)
							'response.write caf(i)
							if (trim(caf(i))=trim(objrecordset("palabra"))) then
								tiene = "checked"
							end if
						next
					end if
					if conta=0 then
						salida_cancer = salida_cancer & "<tr>"
					end if

					salida_cancer = salida_cancer & "<td align='left'><table cellspacing=1 cellpadding=1 border=0 align='left'><tr><td align='left'><input type='checkbox' name='t3_fuente' value='"&objrecordset("palabra")&"' "&tiene&"></td><td align='left'>"&objrecordset("palabra")&"</td><td>&nbsp;&nbsp;&nbsp;</td></tr></table></td>"

					conta = conta + 1

					if conta=7 then
						salida_cancer = salida_cancer & "</tr>"
						conta=0
					end if
					objrecordset.movenext
				 wend
				 if conta>0 then
					salida_cancer = salida_cancer & "</tr>"
				 end if

			%>
            <tr>
				  <td>FUENTE:  </td>
				  <td align="left">
				  	<table cellspacing=2 cellpadding=2 border=0 align='left'>

						<%= salida_cancer %>

					</table>

				  </td>


			</tr>

            <%
             	sql = "select id,palabra from rq_definiciones where fuente like '%CAC%'"
				 Set objRecordset = Server.CreateObject ("ADODB.Recordset")
				 Set objRecordset = objconn1.Execute(sql)

				 dim cac
				 if categoria_cancer_otras <> "" then
					cac = split(categoria_cancer_otras,",")
				 end if

				 conta=0
				 while not objrecordset.eof
					tiene=""
					if categoria_cancer_otras <> "" then
						for i=0 to ubound(cac)
							if (trim(cac(i))=trim(objrecordset("palabra"))) then
								tiene = "checked"
							end if
						next
					end if
					if conta=0 then
						salida_cancer_categoria = salida_cancer_categoria & "<tr>"
					end if

					salida_cancer_categoria = salida_cancer_categoria & "<td align='left'><table cellspacing=1 cellpadding=1 border=0 align='left'><tr><td align='left'><input type='checkbox' name='t3_categoria_cancer_otras' value='"&objrecordset("palabra")&"' "&tiene&"></td><td align='left'>"&objrecordset("palabra")&"</td><td>&nbsp;&nbsp;&nbsp;</td></tr></table></td>"

					conta = conta + 1

					if conta=7 then
						salida_cancer_categoria = salida_cancer_categoria & "</tr>"
						conta=0
					end if
					objrecordset.movenext
				 wend
				 if conta>0 then
					salida_cancer_categoria = salida_cancer_categoria & "</tr>"
				 end if

			%>
            <tr>
				  <td>CATEGORÍA:  </td>
				  <td align="left">
				  	<table cellspacing=2 cellpadding=2 border=0 align='left'>

						<%= salida_cancer_categoria %>

					</table>

				  </td>


			</tr>
            <!--
			 <tr><td>Categoría</td><td><input name="t3_categoria_cancer_otras" type="text" value="<%=categoria_cancer_otras%>" size="50" maxlength="50" /></td></tr>
			 <tr><td>Fuente</td><td><input name="t3_fuente" type="text" value="<%=fuente%>" size="50" maxlength="50" /></td></tr>
            -->
			 </table>
		 </fieldset>
		<fieldset>
 			<legend><strong>Cáncer de mama</strong></legend><input type="hidden" name="hayt9" value="<%=hayt9%>" />
			 <table>
			  <tr>
				<td>Cáncer de mama (si/no)</td>
				<td><input type="checkbox" value="1" name="t9_cancer_mama" <%=dimechecked(cancer_mama)%> /></td>
			  </tr>
			  <tr>
				<td>Cáncer de mama (fuente)</td>
				<td><input name="t9_cancer_mama_fuente" type="text" value="<%=cancer_mama_fuente%>" size="50" maxlength="500" /></td>
			  </tr>
			 </table>
 		</fieldset>
	</fieldset>
	<fieldset>
  		<legend><strong>Neurot&oacute;xico</strong></legend>
  			<input type="hidden" name="hayt4" value="<%=hayt4%>" />
			  <table >
				<tr>
				  <td>Efecto neurot&oacute;xico</td>
				  <td><input name="t4_efecto_neurotoxico" type="text" value="<%=efecto_neurotoxico%>" size="100" maxlength="100" /></td>
				</tr>
				<tr>
				  <td>Nivel de neurotoxicidad </td>
				  <td>
				  <select name="t4_nivel_neurotoxico">
				  <option><%=i%></option>
				  <%
				  for i=1 to 4
					marcado=""
					if nivel_neurotoxico<>"" then
						if isnumeric(nivel_neurotoxico) then
							if cint(i)=cint(nivel_neurotoxico) then marcado="selected"
						end if
					end if
				  %>
				  <option <%=marcado%> ><%=i%></option>
				  <%
				  next
				  %>
				  </select>
				  </td>
				</tr>
				<%
				 function selec(valor1, valor2)

					if (valor1 = valor2) then
						salida = "selected"
					end if
					selec = salida
				 end function

				 sql = "select id,palabra from rq_definiciones where fuente like '%NT%'"
				 Set objRecordset = Server.CreateObject ("ADODB.Recordset")
				 Set objRecordset = objconn1.Execute(sql)

				 dim nt
				 if fuente_neurotoxico <> "" then

					 nt = split(fuente_neurotoxico,",")
				 end if

				 conta=0
				 while not objrecordset.eof
					tiene=""
					if fuente_neurotoxico <> "" then
						for i=0 to ubound(nt)

							if (trim(nt(i))=trim(objrecordset("palabra"))) then
								tiene = "checked"
							end if
						next
					end if
					if conta=0 then
						neuro = neuro & "<tr>"
					end if

					neuro = neuro & "<td align='left'><table cellspacing=1 cellpadding=1 border=0 align='left'><tr><td align='left'><input type='checkbox' name='t4_fuente_neurotoxico' value='"&objrecordset("palabra")&"' "&tiene&"></td><td align='left'>"&objrecordset("palabra")&"</td><td>&nbsp;&nbsp;&nbsp;</td></tr></table></td>"

					conta = conta + 1

					if conta=7 then
						neuro = neuro & "</tr>"
						conta=0
					end if
					objrecordset.movenext
				 wend
				 if conta>0 then
					neuro = neuro & "</tr>"
				 end if

				%>
				<tr>
				  <td>Fuente de neurotoxicidad:  </td>
				  <td align="left">
				  	<table cellspacing=2 cellpadding=2 border=0 align='left'>

						<%= neuro %>

					</table>

				  </td>


				</tr>
			  </table>
	</fieldset>

 	<fieldset>
	 <legend><strong>Disruptor endocrino </strong></legend>
			 <%
			 'Sergio, saco aquellas definiciones con el texto: 'DE', en el campo fuente
			 sql = "select id,palabra from rq_definiciones where fuente like '%DE%'"
			 Set objRecordset = Server.CreateObject ("ADODB.Recordset")
			 Set objRecordset = objconn1.Execute(sql)

			 dim disruptores
			 if nivel_disruptor <> "" then
				 disruptores = split(nivel_disruptor,",")
			 end if

			 conta=0
			 while not objrecordset.eof
				tiene=""
				if nivel_disruptor <> "" then
					for i=0 to ubound(disruptores)
						'response.write disruptores(i)
						if (trim(disruptores(i))=trim(objrecordset("palabra"))) then
							tiene = "checked"
						end if
					next
				end if

				if conta=0 then
						endo = endo & "<tr>"
				end if



				endo = endo & "<td align='left'><table cellspacing=1 cellpadding=1 border=0 align='left'><tr><td><input type='checkbox' name='t4_nivel_disruptor' value='"&objrecordset("palabra")&"' "&tiene&"></td><td>"&objrecordset("palabra")&"</td><td>&nbsp;&nbsp;&nbsp;</td></tr></table></td>"

				conta = conta + 1

				if conta=7 then
					endo = endo & "</tr>"
					conta=0
				end if

				objrecordset.movenext

			 wend

			 if conta>0 then
					endo = endo & "</tr>"
			 end if

			 %>
		 <table>
		  <tr>
			  <td>Fuente: </td>
			  <td>
				<table cellspacing=2 cellpadding=2 border=0>
					<tr>
						<%=endo %>
					</tr>
				</table>
				</td>
		</tr>
		</table>
		</fieldset>
		<fieldset>
		  <legend><strong>Efectos sobre la salud y órganos afectados</strong></legend>
		  <input type="hidden" name="hayt6" value="<%=hayt6%>" />

		  <table border="0">
			<tr>
			  <td align="left">
				<input type="checkbox" value="1"  name="t6_cardiocirculatorio" <%=dimechecked(salud_cardiocirculatorio)%> /> Cardiocirculatorio<br/>
				<input type="checkbox" value="1"  name="t6_rinyon" <%=dimechecked(salud_rinyon)%> /> Riñón<br/>
				<input type="checkbox" value="1"  name="t6_respiratorio" <%=dimechecked(salud_respiratorio)%> /> Respiratorio<br/>
				<input type="checkbox" value="1"  name="t6_reproductivo" <%=dimechecked(salud_reproductivo)%> /> Reproductivo<br/>
				<input type="checkbox" value="1"  name="t6_piel_sentidos" <%=dimechecked(salud_piel_sentidos)%> /> Piel/sentidos<br/>
				<input type="checkbox" value="1"  name="t6_neuro_toxicos" <%=dimechecked(salud_neuro_toxicos)%> /> Neurotóxicos<br/>
			  </td>
			  <td align="left">
				<input type="checkbox" value="1"  name="t6_musculo_esqueletico" <%=dimechecked(salud_musculo_esqueletico)%> /> Musculoesquelético<br/>
				<input type="checkbox" value="1"  name="t6_sistema_inmunitario" <%=dimechecked(salud_sistema_inmunitario)%> /> Sistema  inmunitario<br/>
				<input type="checkbox" value="1"  name="t6_higado_gastrointestinal" <%=dimechecked(salud_higado_gastrointestinal)%> /> Hígado / gastrointestinal<br/>
				<input type="checkbox" value="1"  name="t6_sistema_endocrino" <%=dimechecked(salud_sistema_endocrino)%> /> Sistema endocrino<br/>
				<input type="checkbox" value="1"  name="t6_embrion" <%=dimechecked(salud_embrion)%> /> Embrión<br/>
				<input type="checkbox" value="1"  name="t6_cancer" <%=dimechecked(salud_cancer)%> /> Cáncer<br/>
			  </td>
			</tr>
		  </table>
			<fieldset>
			 <legend><strong>Más información en salud laboral</strong></legend>
			 <table>
			  <tr>
				<td align='center'>
					<textarea name='t6_comentarios' id='t6_comentarios' class="campo" rows='5' cols='60'><%=salud_comentarios%></textarea>
				</td>
			  </tr>
			 </table>
			</fieldset>
			<fieldset>
			 <legend><strong>Más información en salud laboral (ingl&eacute;s)</strong></legend>
			 <table>
			  <tr>
				<td align='center'>
					<textarea name='t6_comentarios_ing' id='t6_comentarios_ing' class="campo" rows='5' cols='60'><%=salud_comentarios_ing%></textarea>
				</td>
			  </tr>
			 </table>
			</fieldset>
		  </fieldset>
	 </div>

</fieldset>
<fieldset>
  <legend><strong><a href='javascript:cambia_pestanya(2)'>RIESGOS ESPECÍFICOS PARA EL MEDIO AMBIENTE</a></strong></legend>
  <div id='div2' style='display:block; visibility:visible'>
  <input type="hidden" name="hayt5" value="<%=hayt5%>" />
  <fieldset>
  <legend><strong>Tóxica, persistente y bioacumulativa</strong></legend>
  <table >
  <tr><td>Enlace TPB</td><td><input name="t5_enlace_tpb" type="text" value="<%=enlace_tpb%>" size="66" maxlength="250" />
    (ej: http://www.google.com)</td>
  </tr>
  <tr><td>Nombre TPB</td><td><input name="t5_anchor_tpb" type="text" value="<%=anchor_tpb%>" size="100" maxlength="750" /></td></tr>
  <%
 'Sergio, saco aquellas definiciones con el texto: 'Listado tpbs', en el campo Nota interna


 sql = "select id,palabra from rq_definiciones where fuente like '%TPB%'"
 Set objRecordset = Server.CreateObject ("ADODB.Recordset")
 Set objRecordset = objconn1.Execute(sql)

 dim bioacumulativas
 if fuentes_tpb <> "" then
	 bioacumulativas = split(fuentes_tpb,",")
 end if

 conta = 0
 while not objrecordset.eof
 	tiene=""

	if fuentes_tpb <> "" then
		for i=0 to ubound(bioacumulativas)
			'response.write disruptores(i)
			if (trim(bioacumulativas(i))=trim(objrecordset("palabra"))) then
				tiene = "checked"
			end if
		next
	end if

	if conta=0 then
			salida_tpb = salida_tpb & "<tr>"
	end if

 	salida_tpb = salida_tpb & "<td align='left'><table cellspacing=1 cellpadding=1 border=0 align='left'><tr><td><input type='checkbox' name='t5_fuentes_tpb' value='"&objrecordset("palabra")&"' "&tiene&"></td><td>"&objrecordset("palabra")&"</td><td>&nbsp;&nbsp;&nbsp;</td></tr></table></td>"

	conta = conta + 1

	if conta=7 then
		salida_tpb = salida_tpb & "</tr>"
		conta=0
	end if


 	objrecordset.movenext
 wend

 if conta>0 then
	 	salida_tpb = salida_tpb & "</tr>"
 end if

 %>
  <tr>
  	<td >
	Fuente:
	</td>
	<td align='left'>
		<table cellspacing=2 cellpadding=2 border=0 align='left'>
		<tr>
			<%=salida_tpb %>
		</tr>
	</table>
	</td>

  </tr>
  </table>
  </fieldset>


  <table >
  <tr>
  	<td valign="top">
  <fieldset><legend><strong>Daño a la atmósfera</strong></legend>
  <input type="checkbox" value="1"  name="t5_dano_calidad_aire" <%=dimechecked(dano_calidad_aire)%> /> Calidad del aire
  <br />
  <input type="checkbox" value="1"  name="t5_dano_ozono" <%=dimechecked(dano_ozono)%> /> Capa de ozono
  <br />
   <input type="checkbox" value="1"  name="t5_dano_cambio_clima" <%=dimechecked(dano_cambio_clima)%> /> Cambio climático
  <br />
  </fieldset>
  </td>
  <td valign="top">
  	<table cellspacing=0 cellpadding=0 bordedr=0>
		<tr>
			<td>
				  <fieldset style="margin-left:10px; "><legend><strong>Toxicidad acuática</strong></legend>
				  <input type="checkbox" value="1"  name="t5_directiva_aguas" <%=dimechecked(directiva_aguas)%> /> Aparece en la directiva de aguas
				  <br>
				  <input type="checkbox" value="1"  name="t5_sustancia_prioritaria" <%=dimechecked(sustancia_prioritaria)%> /> Posible sustancia prioritaria
				   <br /><br>
				  Clasificación MMA Alemania
				  <select name="t5_clasif_mma">
					<option value=""> </option>
					  <%
					  for i=1 to 3
						marcado=""
						if clasif_mma<>"" then
							if isnumeric(clasif_mma) then
								if cint(i)=cint(clasif_mma) then marcado="selected"
							end if
						end if
					  %>
					  <option <%=marcado%> ><%=i%></option>
					  <%
					  next
					  %>
					  <option value="nwg" <%if clasif_mma="nwg" then response.write "selected"%>>nwg</option>
					  </select>
				  </fieldset>
			</td>

		</tr>
		<tr>
			<td>

			</td>
		</tr>
	 </table>
  </td>
  <td valign='top'>
  	<fieldset style="margin-left:10px; "><legend><strong>Contaminantes de suelos</strong></legend>
				  <input type="checkbox" value="1"  name="t5_toxicidad_suelo" <%=dimechecked(toxicidad_suelo)%> /> Real decreto 9/2005
				  <br /> <br />
	</fieldset>
  </td>
   <td colspan="2" valign="top">

  </td>
  </tr>
  </table>

  <fieldset>
 <legend><strong>COP</strong></legend><input type="hidden" name="hayt8" value="<%=hayt8%>" />
 <table>
  <tr>
    <td>Anexos COP (letras separadas por punto y coma)</td>
    <td><input name="t8_cop" type="text" value="<%=cop%>" size="50" maxlength="100" /></td>
  </tr>
 </table>
 </fieldset>
 <fieldset>
 <legend><strong>Más información en medio ambiente</strong></legend>
 <table>
  <tr>
    <td align='center'>
		<textarea name='t5_comentarios' id='t5_comentarios' class="campo" rows='5' cols='60'><%=am_comentarios%></textarea>
	</td>
  </tr>
 </table>
 </fieldset>
 <fieldset>
 <legend><strong>Más información en medio ambiente (ingl&eacute;s)</strong></legend>
 <table>
  <tr>
    <td align='center'>
		<textarea name='t5_comentarios_ing' id='t5_comentarios_ing' class="campo" rows='5' cols='60'><%=am_comentarios_ing%></textarea>
	</td>
  </tr>
 </table>
 </fieldset>
</div>
</fieldset>












 <fieldset>
  <legend><strong><a href='javascript:cambia_pestanya(3)'>NORMATIVA SOBRE SALUD LABORAL</a></strong></legend>
  <div id='div3' style='display:block; visibility:visible'>
  <input type="hidden" name="hayt1" value="<%=hayt1%>" />
       <fieldset>
			  <legend><strong>Valores límites de exposición Ambiental</strong></legend>
				<table >
			  <tr>
			  <th> </th>
			  <th>ESTADO</th>
			  <th>VLA-ED (pmm)</th>
			  <th>VLA-ED (mg/m3)</th>
			  <th>VLA-EC (pmm)</th>
			  <th>VLA-EC (mg/m3)</th>
			  <th>NOTAS VLA</th>
			  </tr>
			  <tr>
			  <td> 1 </td>
			  <td><input name='t1_estado_1' type='text' value='<%=estado_1%>' maxlength='200' /></td>
			  <td><input type='text' name='t1_vla_ed_ppm_1' value='<%=vla_ed_ppm_1%>' maxlength='50' size='16' /></td>
			  <td><input type='text' name='t1_vla_ed_mg_m3_1' value='<%=vla_ed_mg_m3_1%>' maxlength='50' size='16' /></td>
			  <td><input type='text' name='t1_vla_ec_ppm_1' value='<%=vla_ec_ppm_1%>' maxlength='50' size='16' /></td>
			  <td><input type='text' name='t1_vla_ec_mg_m3_1' value='<%=vla_ec_mg_m3_1%>' maxlength='50' size='16' /></td>
			  <td><input type='text' name='t1_notas_vla_1' value='<%=notas_vla_1%>' maxlength='50' size='16' /></td>
			  </tr>
			  <tr>
			  <td> 2 </td>
			  <td><input type='text' name='t1_estado_2' value='<%=estado_2%>' maxlength='200' /></td>
			  <td><input type='text' name='t1_vla_ed_ppm_2' value='<%=vla_ed_ppm_2%>' maxlength='50' size='16' /></td>
			  <td><input type='text' name='t1_vla_ed_mg_m3_2' value='<%=vla_ed_mg_m3_2%>' maxlength='50' size='16' /></td>
			  <td><input type='text' name='t1_vla_ec_ppm_2' value='<%=vla_ec_ppm_2%>' maxlength='50' size='16' /></td>
			  <td><input type='text' name='t1_vla_ec_mg_m3_2' value='<%=vla_ec_mg_m3_2%>' maxlength='50' size='16' /></td>
			  <td><input type='text' name='t1_notas_vla_2' value='<%=notas_vla_2%>' maxlength='50' size='16' /></td>
			  </tr>
				<tr>
			  <td> 3 </td>
			  <td><input type='text' name='t1_estado_3' value='<%=estado_3%>' maxlength='200' /></td>
			  <td><input type='text' name='t1_vla_ed_ppm_3' value='<%=vla_ed_ppm_3%>' maxlength='50' size='16' /></td>
			  <td><input type='text' name='t1_vla_ed_mg_m3_3' value='<%=vla_ed_mg_m3_3%>' maxlength='50' size='16' /></td>
			  <td><input type='text' name='t1_vla_ec_ppm_3' value='<%=vla_ec_ppm_3%>' maxlength='50' size='16' /></td>
			  <td><input type='text' name='t1_vla_ec_mg_m3_3' value='<%=vla_ec_mg_m3_3%>' maxlength='50' size='16' /></td>
			  <td><input type='text' name='t1_notas_vla_3' value='<%=notas_vla_3%>' maxlength='50' size='16' /></td>
			  </tr>
			  <tr>
			  <td> 4 </td>
			  <td><input type='text' name='t1_estado_4' value='<%=estado_4%>' maxlength='200' /></td>
			  <td><input type='text' name='t1_vla_ed_ppm_4' value='<%=vla_ed_ppm_4%>' maxlength='50' size='16' /></td>
			  <td><input type='text' name='t1_vla_ed_mg_m3_4' value='<%=vla_ed_mg_m3_4%>' maxlength='50' size='16' /></td>
			  <td><input type='text' name='t1_vla_ec_ppm_4' value='<%=vla_ec_ppm_4%>' maxlength='50' size='16' /></td>
			  <td><input type='text' name='t1_vla_ec_mg_m3_4' value='<%=vla_ec_mg_m3_4%>' maxlength='50' size='16' /></td>
			  <td><input type='text' name='t1_notas_vla_4' value='<%=notas_vla_4%>' maxlength='50' size='16' /></td>
			  </tr>
			   <tr>
			  <td> 5 </td>
			  <td><input type='text' name='t1_estado_5' value='<%=estado_5%>' maxlength='200' /></td>
			  <td><input type='text' name='t1_vla_ed_ppm_5' value='<%=vla_ed_ppm_5%>' maxlength='50' size='16' /></td>
			  <td><input type='text' name='t1_vla_ed_mg_m3_5' value='<%=vla_ed_mg_m3_5%>' maxlength='50' size='16' /></td>
			  <td><input type='text' name='t1_vla_ec_ppm_5' value='<%=vla_ec_ppm_5%>' maxlength='50' size='16' /></td>
			  <td><input type='text' name='t1_vla_ec_mg_m3_5' value='<%=vla_ec_mg_m3_5%>' maxlength='50' size='16' /></td>
			  <td><input type='text' name='t1_notas_vla_5' value='<%=notas_vla_5%>' maxlength='50' size='16' /></td>
			  </tr>
			  <tr>
			  <td> 6 </td>
			  <td><input type='text' name='t1_estado_6' value='<%=estado_6%>' maxlength='200' /></td>
			  <td><input type='text' name='t1_vla_ed_ppm_6' value='<%=vla_ed_ppm_6%>' maxlength='50' size='16' /></td>
			  <td><input type='text' name='t1_vla_ed_mg_m3_6' value='<%=vla_ed_mg_m3_6%>' maxlength='50' size='16' /></td>
			  <td><input type='text' name='t1_vla_ec_ppm_6' value='<%=vla_ec_ppm_6%>' maxlength='50' size='16' /></td>
			  <td><input type='text' name='t1_vla_ec_mg_m3_6' value='<%=vla_ec_mg_m3_6%>' maxlength='50' size='16' /></td>
			  <td><input type='text' name='t1_notas_vla_6' value='<%=notas_vla_6%>' maxlength='50' size='16' /></td>
			  </tr>
			  </table>
 	  </fieldset>

	  <fieldset>
			  <legend><strong>Valores límites de exposición Biol&oacute;gico</strong></legend>
			  <table >
			  <tr>
			  <th> </th>
			  <th>INDICADOR BIOL&Oacute;GICO</th>
			  <th>VLB</th>
			  <th>MOMENTO DE MUESTREO </th>
			  <th>NOTAS VLB </th>
			  </tr>
			  <tr>
			  <td> 1 </td>
			  <td><input name='t1_ib_1' type='text' value='<%=ib_1%>' size="31" maxlength='200' /></td>
			  <td><input type='text' name='t1_vlb_1' value='<%=vlb_1%>' maxlength='100' size='16' /></td>
			  <td><input name='t1_momento_1' type='text' value='<%=momento_1%>' size="30" maxlength='100' /></td>
			  <td><input type='text' name='t1_notas_vlb_1' value='<%=notas_vlb_1%>' maxlength='40' /></td>
			  </tr>
			 <tr>
			  <td> 2 </td>
			  <td><input name='t1_ib_2' type='text' value='<%=ib_2%>' size="31" maxlength='200' /></td>
			  <td><input type='text' name='t1_vlb_2' value='<%=vlb_2%>' maxlength='100' size='16' /></td>
			  <td><input name='t1_momento_2' type='text' value='<%=momento_2%>' size="30" maxlength='100' /></td>
			  <td><input type='text' name='t1_notas_vlb_2' value='<%=notas_vlb_2%>' maxlength='40' /></td>
			  </tr>
				<tr>
			  <td> 3 </td>
			  <td><input name='t1_ib_3' type='text' value='<%=ib_3%>' size="31" maxlength='200' /></td>
			  <td><input type='text' name='t1_vlb_3' value='<%=vlb_3%>' maxlength='100' size='16' /></td>
			  <td><input name='t1_momento_3' type='text' value='<%=momento_3%>' size="30" maxlength='100' /></td>
			  <td><input type='text' name='t1_notas_vlb_3' value='<%=notas_vlb_3%>' maxlength='40' /></td>
			  </tr>
				<tr>
			  <td> 4 </td>
			  <td><input name='t1_ib_4' type='text' value='<%=ib_4%>' size="31" maxlength='200' /></td>
			  <td><input type='text' name='t1_vlb_4' value='<%=vlb_4%>' maxlength='100' size='16' /></td>
			  <td><input name='t1_momento_4' type='text' value='<%=momento_4%>' size="30" maxlength='100' /></td>
			  <td><input type='text' name='t1_notas_vlb_4' value='<%=notas_vlb_4%>' maxlength='40' /></td>
			  </tr>
				<tr>
			  <td> 5 </td>
			  <td><input name='t1_ib_5' type='text' value='<%=ib_5%>' size="31" maxlength='200' /></td>
			  <td><input type='text' name='t1_vlb_5' value='<%=vlb_5%>' maxlength='100' size='16' /></td>
			  <td><input name='t1_momento_5' type='text' value='<%=momento_5%>' size="30" maxlength='100' /></td>
			  <td><input type='text' name='t1_notas_vlb_5' value='<%=notas_vlb_5%>' maxlength='40' /></td>
			  </tr>
				<tr>
			  <td> 6 </td>
			  <td><input name='t1_ib_6' type='text' value='<%=ib_6%>' size="31" maxlength='200' /></td>
			  <td><input type='text' name='t1_vlb_6' value='<%=vlb_6%>' maxlength='100' size='16' /></td>
			  <td><input name='t1_momento_6' type='text' value='<%=momento_6%>' size="30" maxlength='100' /></td>
			  <td><input type='text' name='t1_notas_vlb_6' value='<%=notas_vlb_6%>' maxlength='40' /></td>
			  </tr>
			  </table>
	  </fieldset>
 </fieldset>
 <fieldset style="margin-left:10px; ">
	<legend><strong><a href='javascript:cambia_pestanya(4)'>NORMATIVA AMBIENTAL</a></strong></strong></legend>
	<div id='div4' style='display:block; visibility:visible'>
	<table cellspacing=1 cellpadding=1 border=0>
		<tr>
			<td><input type="checkbox" value="1"  name="t5_emisiones_atmosfera" <%=dimechecked(emisiones_atmosfera)%> />
  		Emisiones atmosf&eacute;ricas
			</td>
		</tr>
 	 	<tr>
			<td><input type="checkbox" value="1"  name="t5_cov" <%=dimechecked(cov)%> />
   		Compuestos org&aacute;nicos vol&aacute;tiles </td>
		</tr>
   	    <tr>
			<td><input type="checkbox" value="1"  name="t5_seveso" <%=dimechecked(seveso)%> />
   		Seveso accidentes graves </td>
		</tr>
        <tr>
			<td><input type="checkbox" value="1"  name="t5_eper_agua" <%=dimechecked(eper_agua)%> />
  		EPER (IPPC) agua </td>
		</tr>
  	    <tr>
			<td><input type="checkbox" value="1"  name="t5_eper_aire" <%=dimechecked(eper_aire)%> />
   		EPER (IPPC) aire </td>
		</tr>
        <tr>
			<td><input type="checkbox" value="1"  name="t5_eper_suelo" <%=dimechecked(eper_suelo)%> />
   		EPER (IPPC) suelo </td>
		</tr>
	</table>
	</div>
  </fieldset>


 <fieldset style="margin-left:10px; ">
	<legend><strong>Normativa sobre restricción / prohibición de sustancias</strong></legend>
			 <table>
			  <tr>
				<td><input type="checkbox" value="1" name="hayt12" <%=dimechecked(hayt12)%> /></td>
				<td>Prohibidas (s&iacute;/no)</td>
			  </tr>
			  <tr>
				<td>&nbsp;</td>
				<td>Comentarios:<br><textarea cols="50" rows="3" name='t12_comentario_prohibida'><%= t12_comentario_prohibida %></textarea></td>
			  </tr>
			  <tr>
				<td>&nbsp;</td>
				<td>Comentarios (ingl&eacute;s):<br><textarea cols="50" rows="3" name='t12_comentario_prohibida_ing'><%= t12_comentario_prohibida_ing %></textarea></td>
			  </tr>

			  <tr>
				<td><input type="checkbox" value="1" name="hayt13" <%=dimechecked(hayt13)%> /></td>
				<td>Restringidas(s&iacute;/no)</td>
			  </tr>
			  <tr>
				<td>&nbsp;</td>
				<td>Comentarios:<br><textarea cols="50" rows="3" name='t13_comentario_restringida'><%= t13_comentario_restringida %></textarea></td>
			  </tr>
			  <tr>
				<td>&nbsp;</td>
				<td>Comentarios (ingl&eacute;s):<br><textarea cols="50" rows="3" name='t13_comentario_restringida_ing'><%= t13_comentario_restringida_ing %></textarea></td>
			  </tr>

			  <tr>
				<td><input type="checkbox" value="1" name="hayt10" <%=dimechecked(hayt10)%> /></td>
				<td>Prohibidas para trabajadoras embarazadas (s&iacute;/no)</td>
			  </tr>
			  <tr>
				<td>&nbsp;</td>
				<td>Comentarios:<br><textarea cols="50" rows="3" name='t10_comentario_prohibida'><%= t10_comentario_prohibida %></textarea></td>
			  </tr>
			  <tr>
				<td>&nbsp;</td>
				<td>Comentarios (ingl&eacute;s):<br><textarea cols="50" rows="3" name='t10_comentario_prohibida_ing'><%= t10_comentario_prohibida_ing %></textarea></td>
			  </tr>

			  <tr>
				<td><input type="checkbox" value="1" name="hayt11" <%=dimechecked(hayt11)%> /></td>
				<td>Prohibidas para trabajadoras lactantes (s&iacute;/no)</td>
			  </tr>
			  <tr>
				<td>&nbsp;</td>
				<td>Comentarios:<br><textarea cols="50" rows="3" name='t11_comentario_prohibida'><%= t11_comentario_prohibida %></textarea></td>
			  </tr>
			  <tr>
				<td>&nbsp;</td>
				<td>Comentarios (ingl&eacute;s):<br><textarea cols="50" rows="3" name='t11_comentario_prohibida_ing'><%= t11_comentario_prohibida_ing %></textarea></td>
			  </tr>

			  <tr>
				<td><input type="checkbox" value="1" name="hayt14" <%=dimechecked(hayt14)%> /></td>
				<td>Sustancias candidatas REACH (s&iacute;/no)</td>
			  </tr>
			  <tr>
				<td><input type="checkbox" value="1" name="hayt15" <%=dimechecked(hayt15)%> /></td>
				<td>Sustancias REACH sujetas a autorizaci&oacute;n (s&iacute;/no)</td>
			  </tr>

			  <tr>
				<td><input type="checkbox" value="1" name="hayt16" <%=dimechecked(hayt16)%> /></td>
				<td>Sustancias Biocidas prohibidas (s&iacute;/no)</td>
			  </tr>
			  <tr>
				<td>&nbsp;</td>
				<td>Fuente:<br><textarea cols="50" rows="3" name='t16_fuente'><%= t16_fuente %></textarea></td>
			  </tr>
			  <tr>
				<td>&nbsp;</td>
				<td>Fecha l&iacute;mite:<br><textarea cols="50" rows="3" name='t16_fecha_limite'><%= t16_fecha_limite %></textarea></td>
			  </tr>
			  <tr>
				<td>&nbsp;</td>
				<td>Usos:<br><textarea cols="50" rows="3" name='t16_usos'><%= t16_usos %></textarea></td>
			  </tr>
			  <tr>
				<td>&nbsp;</td>
				<td>Usos (ingl&eacute;s):<br><textarea cols="50" rows="3" name='t16_usos_ing'><%= t16_usos_ing %></textarea></td>
			  </tr>

			  <tr>
				<td><input type="checkbox" value="1" name="hayt17" <%=dimechecked(hayt17)%> /></td>
				<td>Sustancias Biocidas autorizadas (s&iacute;/no)</td>
			  </tr>
			  <tr>
				<td>&nbsp;</td>
				<td>Fuente:<br><textarea cols="50" rows="3" name='t17_fuente'><%= t17_fuente %></textarea></td>
			  </tr>
			  <tr>
				<td>&nbsp;</td>
				<td>Pureza m&iacute;nima:<br><textarea cols="50" rows="3" name='t17_pureza_minima'><%= t17_pureza_minima %></textarea></td>
			  </tr>
			  <tr>
				<td>&nbsp;</td>
				<td>Condiciones:<br><textarea cols="50" rows="3" name='t17_condiciones'><%= t17_condiciones %></textarea></td>
			  </tr>
			  <tr>
				<td>&nbsp;</td>
				<td>Condiciones (ingl&eacute;s):<br><textarea cols="50" rows="3" name='t17_condiciones_ing'><%= t17_condiciones_ing %></textarea></td>
			  </tr>
			  <tr>
				<td>&nbsp;</td>
				<td>Usos:<br><textarea cols="50" rows="3" name='t17_usos'><%= t17_usos %></textarea></td>
			  </tr>
			  <tr>
				<td>&nbsp;</td>
				<td>Usos (ingl&eacute;s):<br><textarea cols="50" rows="3" name='t17_usos_ing'><%= t17_usos_ing %></textarea></td>
			  </tr>

			  <tr>
				<td><input type="checkbox" value="1" name="hayt18" <%=dimechecked(hayt18)%> /></td>
				<td>Sustancias Pesticidas autorizadas (s&iacute;/no)</td>
			  </tr>
			  <tr>
				<td>&nbsp;</td>
				<td>Fuente:<br><textarea cols="50" rows="3" name='t18_fuente'><%= t18_fuente %></textarea></td>
			  </tr>
			  <tr>
				<td>&nbsp;</td>
				<td>Plazo renovaci&oacute;n:<br><textarea cols="50" rows="3" name='t18_plazo_renovacion'><%= t18_plazo_renovacion %></textarea></td>
			  </tr>
			  <tr>
				<td>&nbsp;</td>
				<td>Plazo renovaci&oacute;n (ingl&eacute;s):<br><textarea cols="50" rows="3" name='t18_plazo_renovacion_ing'><%= t18_plazo_renovacion_ing %></textarea></td>
			  </tr>
			  <tr>
				<td>&nbsp;</td>
				<td>Pureza m&iacute;nima:<br><textarea cols="50" rows="3" name='t18_pureza_minima'><%= t18_pureza_minima %></textarea></td>
			  </tr>
			  <tr>
				<td>&nbsp;</td>
				<td>Pureza m&iacute;nima (ingl&eacute;s):<br><textarea cols="50" rows="3" name='t18_pureza_minima_ing'><%= t18_pureza_minima_ing %></textarea></td>
			  </tr>
			  <tr>
				<td>&nbsp;</td>
				<td>Usos:<br><textarea cols="50" rows="3" name='t18_usos'><%= t18_usos %></textarea></td>
			  </tr>
			  <tr>
				<td>&nbsp;</td>
				<td>Usos (ingl&eacute;s):<br><textarea cols="50" rows="3" name='t18_usos_ing'><%= t18_usos_ing %></textarea></td>
			  </tr>

			  <tr>
				<td><input type="checkbox" value="1" name="hayt19" <%=dimechecked(hayt19)%> /></td>
				<td>Sustancias Pesticidas prohibidas (s&iacute;/no)</td>
			  </tr>
			  <tr>
				<td>&nbsp;</td>
				<td>Fuente:<br><textarea cols="50" rows="3" name='t19_fuente'><%= t19_fuente %></textarea></td>
			  </tr>
			  <tr>
				<td>&nbsp;</td>
				<td>Exenciones:<br><textarea cols="50" rows="3" name='t19_exenciones'><%= t19_exenciones %></textarea></td>
			  </tr>
			  <tr>
				<td>&nbsp;</td>
				<td>Exenciones (ingl&eacute;s):<br><textarea cols="50" rows="3" name='t19_exenciones_ing'><%= t19_exenciones_ing %></textarea></td>
			  </tr>
			 </table>
  </fieldset>


   <p ><input type="submit" value="Enviar" class="centcontenido"  /></p>
  </form>

<script language="JavaScript" type="text/javascript">
var frmvalidator = new Validator("myform");
//frmvalidator.addValidation("eti_conc_15","maxlen=50");
</script>
<%
'Sergio
cerrarconexion

%>
</div>
</body>
</html>
