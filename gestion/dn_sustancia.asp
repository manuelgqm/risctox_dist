<!--#include file="dn_fun_comunes.asp"-->
<!--#include file="dn_auten.inc"-->



<%
'si nos pasan id de sustancia, consultamos datos
id=request("id")
if id<>"" then
%>
	<!--#include file="adovbs.inc"-->
	<!--#include file="dn_conexion.asp"-->
<%
	'DATOS GENERALES

	sql3="select * from dn_risc_sustancias where id=" &id
	'response.write sqls
	set objRst3=objconn1.execute(sql3)

	'id=objRst3("id")
	nombre=objRst3("nombre")
	nombre_ing=objRst3("nombre_ing")
	num_rd=objRst3("num_rd")
	num_ce_einecs=objRst3("num_ce_einecs")
	num_ce_elincs=objRst3("num_ce_elincs")
	num_cas=objRst3("num_cas")
	cas_alternativos=objRst3("cas_alternativos")
	num_onu=objRst3("num_onu")
	formula_molecular=objRst3("formula_molecular")
	estructura_molecular=objRst3("estructura_molecular")
	simbolos=objRst3("simbolos")
	clasificacion_1=objRst3("clasificacion_1")
	clasificacion_2=objRst3("clasificacion_2")
	clasificacion_3=objRst3("clasificacion_3")
	clasificacion_4=objRst3("clasificacion_4")
	clasificacion_5=objRst3("clasificacion_5")
	clasificacion_6=objRst3("clasificacion_6")
	clasificacion_7=objRst3("clasificacion_7")
	clasificacion_8=objRst3("clasificacion_8")
	clasificacion_9=objRst3("clasificacion_9")
	clasificacion_10=objRst3("clasificacion_10")
	clasificacion_11=objRst3("clasificacion_11")
	clasificacion_12=objRst3("clasificacion_12")
	clasificacion_13=objRst3("clasificacion_13")
	clasificacion_14=objRst3("clasificacion_14")
	clasificacion_15=objRst3("clasificacion_15")
	frases_s=objRst3("frases_s")
	conc_1=objRst3("conc_1")
	eti_conc_1=objRst3("eti_conc_1")
	conc_2=objRst3("conc_2")
	eti_conc_2=objRst3("eti_conc_2")
	conc_3=objRst3("conc_3")
	eti_conc_3=objRst3("eti_conc_3")
	conc_4=objRst3("conc_4")
	eti_conc_4=objRst3("eti_conc_4")
	conc_5=objRst3("conc_5")
	eti_conc_5=objRst3("eti_conc_5")
	conc_6=objRst3("conc_6")
	eti_conc_6=objRst3("eti_conc_6")
	conc_7=objRst3("conc_7")
	eti_conc_7=objRst3("eti_conc_7")
	conc_8=objRst3("conc_8")
	eti_conc_8=objRst3("eti_conc_8")
	conc_9=objRst3("conc_9")
	eti_conc_9=objRst3("eti_conc_9")
	conc_10=objRst3("conc_10")
	eti_conc_10=objRst3("eti_conc_10")
	conc_11=objRst3("conc_11")
	eti_conc_11=objRst3("eti_conc_11")
	conc_12=objRst3("conc_12")
	eti_conc_12=objRst3("eti_conc_12")
	conc_13=objRst3("conc_13")
	eti_conc_13=objRst3("eti_conc_13")
	conc_14=objRst3("conc_14")
	eti_conc_14=objRst3("eti_conc_14")
	conc_15=objRst3("conc_15")
	eti_conc_15=objRst3("eti_conc_15")
	notas_rd_363=objRst3("notas_rd_363")
	notas_xml=objRst3("notas_xml")
	notas_xml_ing=objRst3("notas_xml_ing")

	frases_r_danesa=objRst3("frases_r_danesa")




	'SPL
	' RD1272/2008
	clasificacion_rd1272_1 = trim(objRst3("clasificacion_rd1272_1"))
	clasificacion_rd1272_2 = trim(objRst3("clasificacion_rd1272_2"))
	clasificacion_rd1272_3 = trim(objRst3("clasificacion_rd1272_3"))
	clasificacion_rd1272_4 = trim(objRst3("clasificacion_rd1272_4"))
	clasificacion_rd1272_5 = trim(objRst3("clasificacion_rd1272_5"))
	clasificacion_rd1272_6 = trim(objRst3("clasificacion_rd1272_6"))
	clasificacion_rd1272_7 = trim(objRst3("clasificacion_rd1272_7"))
	clasificacion_rd1272_8 = trim(objRst3("clasificacion_rd1272_8"))
	clasificacion_rd1272_9 = trim(objRst3("clasificacion_rd1272_9"))
	clasificacion_rd1272_10 = trim(objRst3("clasificacion_rd1272_10"))
	clasificacion_rd1272_11 = trim(objRst3("clasificacion_rd1272_11"))
	clasificacion_rd1272_12 = trim(objRst3("clasificacion_rd1272_12"))
	clasificacion_rd1272_13 = trim(objRst3("clasificacion_rd1272_13"))
	clasificacion_rd1272_14 = trim(objRst3("clasificacion_rd1272_14"))
	clasificacion_rd1272_15 = trim(objRst3("clasificacion_rd1272_15"))
	conc_rd1272_1 = objRst3("conc_rd1272_1")
	eti_conc_rd1272_1 = objRst3("eti_conc_rd1272_1")
	conc_rd1272_2 = objRst3("conc_rd1272_2")
	eti_conc_rd1272_2 = objRst3("eti_conc_rd1272_2")
	conc_rd1272_3 = objRst3("conc_rd1272_3")
	eti_conc_rd1272_3 = objRst3("eti_conc_rd1272_3")
	conc_rd1272_4 = objRst3("conc_rd1272_4")
	eti_conc_rd1272_4 = objRst3("eti_conc_rd1272_4")
	conc_rd1272_5 = objRst3("conc_rd1272_5")
	eti_conc_rd1272_5 = objRst3("eti_conc_rd1272_5")
	conc_rd1272_6 = objRst3("conc_rd1272_6")
	eti_conc_rd1272_6 = objRst3("eti_conc_rd1272_6")
	conc_rd1272_7 = objRst3("conc_rd1272_7")
	eti_conc_rd1272_7 = objRst3("eti_conc_rd1272_7")
	conc_rd1272_8 = objRst3("conc_rd1272_8")
	eti_conc_rd1272_8 = objRst3("eti_conc_rd1272_8")
	conc_rd1272_9 = objRst3("conc_rd1272_9")
	eti_conc_rd1272_9 = objRst3("eti_conc_rd1272_9")
	conc_rd1272_10 = objRst3("conc_rd1272_10")
	eti_conc_rd1272_10 = objRst3("eti_conc_rd1272_10")
	conc_rd1272_11 = objRst3("conc_rd1272_11")
	eti_conc_rd1272_11 = objRst3("eti_conc_rd1272_11")
	conc_rd1272_12 = objRst3("conc_rd1272_12")
	eti_conc_rd1272_12 = objRst3("eti_conc_rd1272_12")
	conc_rd1272_13 = objRst3("conc_rd1272_13")
	eti_conc_rd1272_13 = objRst3("eti_conc_rd1272_13")
	conc_rd1272_14 = objRst3("conc_rd1272_14")
	eti_conc_rd1272_14 = objRst3("eti_conc_rd1272_14")
	conc_rd1272_15 = objRst3("conc_rd1272_15")
	eti_conc_rd1272_15 = objRst3("eti_conc_rd1272_15")
	notas_rd1272 = objRst3("notas_rd1272")
	simbolos_rd1272 = objRst3("simbolos_rd1272")
	notas_rd1272=objRst3("notas_rd1272")








	'sergio
	'negra = objrst3("negra")
	sustancia_prohibida = objrst3("sustancia_prohibida")
	sustancia_restringida = objrst3("sustancia_restringida")
	objRst3.close
	set objRst3=nothing

	comentarios_pro = ""
	sql_pro = "SELECT comentario_prohibida FROM dn_risc_sustancias_prohibidas WHERE id_sustancia="&id
	set obj_pro = objconn1.execute(sql_pro)
	do while not obj_pro.eof
		comentarios_pro = comentarios_pro & obj_pro("comentario_prohibida")
		obj_pro.movenext
	loop
	'-- puede haber varios registros con comentarios de la misma sustancia, por eso los acumulo
	comentarios_res = ""
	sql_res = "SELECT comentario_restringida FROM dn_risc_sustancias_restringidas WHERE id_sustancia="&id
	set obj_res = objconn1.execute(sql_res)
	do while not obj_res.eof
		comentarios_res = comentarios_res & obj_res("comentario_restringida")
		obj_res.movenext
		if not obj_res.eof then comentarios_res = comentarios_res & "---------------------------------------------------------------------------"
	loop



	'SINONIMOS
	sql3="select nombre from dn_risc_sinonimos where id_sustancia=" &id
	set objRst3=objconn1.execute(sql3)
	if objRst3.eof then
		sinonimos=""
	else
		sinonimos=objRst3.GetString(adClipString, -1, "", "@ ", "")
	end if
	objRst3.close
	set objRst3=nothing

	'NOMBRES COMERCIALES
	sql3="select nombre from dn_risc_nombres_comerciales where id_sustancia=" &id
	set objRst3=objconn1.execute(sql3)
	if objRst3.eof then
		nombres_comerciales=""
	else
		nombres_comerciales=objRst3.GetString(adClipString, -1, "", "@ ", "")
	end if
	objRst3.close
	set objRst3=nothing

	cerrarconexion

end if
%>

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
function DoCustomValidation()
{

  var frm = document.forms["myform"];
  if( (frm.num_cas.value=="") && (frm.num_ce_einecs.value=="") && (frm.num_ce_elincs.value=="") && (frm.num_rd.value==""))
  {
    alert("Debe escribir el nº CAS, el nº CE y/o el nº RD.");
    return false;
  }
  else
  {
    return true;
  }
}
</script>
<script language="JavaScript" src="gen_validatorv2.js" type="text/javascript"></script>
</head>

<body>
<%flashMsgShow()%>
<div id="box2" class="centcontenido">
<form name="myform" action="dn_sustancia2.asp?id=<%=id%>" method="post" enctype="multipart/form-data">

  <fieldset><legend><strong>Identificaci&oacute;n de la sustancia</strong></legend>
Nombre <span class="exp">(obligatorio,max 1000 c.)</span><br>  <textarea cols="60" rows="3" name="nombre"><%=nombre%></textarea><br>
<br>Nombre en ingl&eacute;s <span class="exp">(max 1000 c.)</span><br>  <textarea cols="60" rows="3" name="nombre_ing"><%=nombre_ing%></textarea><br>
<br>Sin&oacute;nimos <span class="exp">(separados por <strong>@</strong> Cada uno max 750 c.)</span><br>  <textarea cols="60" rows="3" name="sinonimos"><%=sinonimos%></textarea><br>
<br>Nombres comerciales <span class="exp">(Separados por <strong>@</strong>; Cada uno max 100 c.)</span><br>  <textarea cols="60" rows="3" name="nombres_comerciales"><%=nombres_comerciales%></textarea><br>  <br>


Notas RD 363 <br>  <textarea cols="60" rows="3" name="notas_rd_363"><%=notas_rd_363%></textarea><br>  <br>



Notas XML <span class="exp">(max 2500 c.)</span><br>  <textarea cols="60" rows="3" name="notas_xml"><%=notas_xml%></textarea><br>  <br>
Notas XML (ingl&eacute;s)<span class="exp">(max 2500 c.)</span><br>  <textarea cols="60" rows="3" name="notas_xml_ing"><%=notas_xml_ing%></textarea><br>  <br>

<table cellspacing=1 cellpadding=1 border=0 align='center'>
	<tr>
		<td colspan=2><br></td>
	</tr>
</table>

  <fieldset><legend><strong>N&ordm;
Identificaci&oacute;n</strong>&nbsp;</legend>N&uacute;mero
RD <span class="exp">(max 20 c.) </span><input maxlength="50" size="20" name="num_rd" value="<%=num_rd%>" /><br>
<br>N&uacute;mero CE/EINECS <span class="exp">(max 20 c.) </span><input maxlength="50" size="20" name="num_ce_einecs" value="<%=num_ce_einecs%>" /><br>
<br>N&uacute;mero CE/ELINCS <span class="exp">(max 20 c.) </span><input maxlength="50" size="20" name="num_ce_elincs" value="<%=num_ce_elincs%>" /><br>

<br>N&uacute;mero CAS <span class="exp">(max 20 c.) </span><input maxlength="50" size="20" name="num_cas" value="<%=num_cas%>" /><br>
<br>CAS alternativos<span class="exp">(separados por <strong>,</strong> Cada uno max 1000 c.)</span><br>  <textarea cols="80" rows="5" name="cas_alternativos"><%=cas_alternativos%></textarea><br>

<br>N&uacute;mero ONU <span class="exp">(max 20 c.) </span><input maxlength="50" size="20" name="num_onu" value="<%=num_onu%>" /><br>  <br>
  </fieldset>  <br>
  <fieldset><legend><strong>Informaci&oacute;n
molecular</strong></legend>F&oacute;rmula molecular <span class="exp">(max 100 c.)</span>&nbsp;<input maxlength="100" size="35" name="formula_molecular" value="<%=formula_molecular%>">
<br><br>

<%
if estructura_molecular="" then
%>
Estructura molecular <input name="estructura_molecular" type="file">  <input name="imagen" type="hidden" value="nueva"  /><br>
<%
else
%>
<table>
<tr><td>Estructura molecular<br />
<img style="width: 200px; " alt="Estructura molecular <%=nombre%> - <%=estructura_molecular%>" src="estructuras/<%=estructura_molecular%>?rand=<%=now()%>">
</td>
<td align="left">
<input name="imagen" type="radio" value="mantener" checked="checked" /> mantener
<br>  <input name="imagen" type="radio" value="cambiar"  /> cambiar por: <input name="estructura_molecular" type="file">
<br>  <input name="imagen" type="radio" value="eliminar" /> eliminar
</td>
</tr>
</table>
<%
end if
%>
  </fieldset>  <br>
  <fieldset><legend><strong>Clasificaci&oacute;n (RD 363/1995)</strong></legend>S&iacute;mbolos:
  <span class="exp">(max 50c.)</span>
  <input maxlength="50" size="35" name="simbolos" value="<%=simbolos%>" /><br>  <br>
  <fieldset><legend><strong>Frases R</strong></legend>
Clasificaci&oacute;n 01 <input value="<%=clasificacion_1%>" maxlength="100" size="70" name="clasificacion_1" />
<br>Clasificaci&oacute;n 02 <input value="<%=clasificacion_2%>" maxlength="100" size="70" name="clasificacion_2">
<br>Clasificaci&oacute;n 03 <input value="<%=clasificacion_3%>" maxlength="100" size="70" name="clasificacion_3">
<br>Clasificaci&oacute;n 04 <input value="<%=clasificacion_4%>" maxlength="100" size="70" name="clasificacion_4">
<br>Clasificaci&oacute;n 05 <input value="<%=clasificacion_5%>" maxlength="100" size="70" name="clasificacion_5">
<br>Clasificaci&oacute;n 06 <input value="<%=clasificacion_6%>" maxlength="100" size="70" name="clasificacion_6">
<br>Clasificaci&oacute;n 07 <input value="<%=clasificacion_7%>" maxlength="100" size="70" name="clasificacion_7">
<br>Clasificaci&oacute;n 08 <input value="<%=clasificacion_8%>" maxlength="100" size="70" name="clasificacion_8">
<br>Clasificaci&oacute;n 09 <input value="<%=clasificacion_9%>" maxlength="100" size="70" name="clasificacion_9">
<br>Clasificaci&oacute;n 10 <input value="<%=clasificacion_10%>" maxlength="100" size="70" name="clasificacion_10">
<br>Clasificaci&oacute;n 11 <input value="<%=clasificacion_11%>" maxlength="100" size="70" name="clasificacion_11">
<br>Clasificaci&oacute;n 12 <input value="<%=clasificacion_12%>" maxlength="100" size="70" name="clasificacion_12">
<br>Clasificaci&oacute;n 13 <input value="<%=clasificacion_13%>" maxlength="100" size="70" name="clasificacion_13">
<br>Clasificaci&oacute;n 14 <input value="<%=clasificacion_14%>" maxlength="100" size="70" name="clasificacion_14">
<br>Clasificaci&oacute;n 15 <input value="<%=clasificacion_15%>" maxlength="100" size="70" name="clasificacion_15">

<br>Frases R Danesa <input value="<%=frases_r_danesa%>" maxlength="100" size="70" name="frases_r_danesa">
<br>

  <span class="exp">(max 100c. cada una)</span><br>  </fieldset>  <br>
  <fieldset><legend><strong>Frases S</strong></legend>Frases S:&nbsp;<span class="exp">(max 100c.)</span>
  <input maxlength="200" size="70" name="frases_s" value="<%=frases_s%>" />
  </fieldset>  <br>
  <fieldset><legend><strong>Etiquetado RD</strong></legend>
Concentraci&oacute;n 01&nbsp;<input value="<%=conc_1%>" maxlength="100" size="30" name="conc_1"> &nbsp;Etiquetado
01&nbsp;<input value="<%=eti_conc_1%>" maxlength="200" size="35" name="eti_conc_1"><br>Concentraci&oacute;n 02&nbsp;<input value="<%=conc_2%>" maxlength="100" size="30" name="conc_2"> &nbsp;Etiquetado
02&nbsp;<input value="<%=eti_conc_2%>" maxlength="200" size="35" name="eti_conc_2"><br>Concentraci&oacute;n 03&nbsp;<input value="<%=conc_3%>" maxlength="100" size="30" name="conc_3"> &nbsp;Etiquetado
03&nbsp;<input value="<%=eti_conc_3%>" maxlength="200" size="35" name="eti_conc_3"><br>Concentraci&oacute;n 04&nbsp;<input value="<%=conc_4%>" maxlength="100" size="30" name="conc_4"> &nbsp;Etiquetado
04&nbsp;<input value="<%=eti_conc_4%>" maxlength="200" size="35" name="eti_conc_4"><br>Concentraci&oacute;n 05&nbsp;<input value="<%=conc_5%>" maxlength="100" size="30" name="conc_5"> &nbsp;Etiquetado
05&nbsp;<input value="<%=eti_conc_5%>" maxlength="200" size="35" name="eti_conc_5"><br>Concentraci&oacute;n 06&nbsp;<input value="<%=conc_6%>" maxlength="100" size="30" name="conc_6"> &nbsp;Etiquetado
06&nbsp;<input value="<%=eti_conc_6%>" maxlength="200" size="35" name="eti_conc_6"><br>Concentraci&oacute;n 07&nbsp;<input value="<%=conc_7%>" maxlength="100" size="30" name="conc_7"> &nbsp;Etiquetado
07&nbsp;<input value="<%=eti_conc_7%>" maxlength="200" size="35" name="eti_conc_7"><br>Concentraci&oacute;n 08&nbsp;<input value="<%=conc_8%>" maxlength="100" size="30" name="conc_8"> &nbsp;Etiquetado
08&nbsp;<input value="<%=eti_conc_8%>" maxlength="200" size="35" name="eti_conc_8"><br>Concentraci&oacute;n 09&nbsp;<input value="<%=conc_9%>" maxlength="100" size="30" name="conc_9"> &nbsp;Etiquetado
09&nbsp;<input value="<%=eti_conc_9%>" maxlength="200" size="35" name="eti_conc_9"><br>Concentraci&oacute;n 10 <input value="<%=conc_10%>" maxlength="100" size="30" name="conc_10"> &nbsp;Etiquetado 10 <input value="<%=eti_conc_10%>" maxlength="200" size="35" name="eti_conc_10"><br>Concentraci&oacute;n 11 <input value="<%=conc_11%>" maxlength="100" size="30" name="conc_11"> &nbsp;Etiquetado 11 <input value="<%=eti_conc_11%>" maxlength="200" size="35" name="eti_conc_11"><br>Concentraci&oacute;n&nbsp;12 <input value="<%=conc_12%>" maxlength="100" size="30" name="conc_12"> &nbsp;Etiquetado 12 <input value="<%=eti_conc_12%>" maxlength="200" size="35" name="eti_conc_12"><br>Concentraci&oacute;n 13 <input value="<%=conc_13%>" maxlength="100" size="30" name="conc_13"> &nbsp;Etiquetado 13 <input value="<%=eti_conc_13%>" maxlength="200" size="35" name="eti_conc_13"><br>Concentraci&oacute;n 14 <input value="<%=conc_14%>" maxlength="100" size="30" name="conc_14"> &nbsp;Etiquetado 14 <input value="<%=eti_conc_14%>" maxlength="200" size="35" name="eti_conc_14"><br>Concentraci&oacute;n 15 <input value="<%=conc_15%>" maxlength="100" size="30" name="conc_15"> &nbsp;Etiquetado 15 <input value="<%=eti_conc_15%>" maxlength="200" size="35" name="eti_conc_15"><br>  <br class="exp">  <span class="exp">(concentraciones,
max 100 c cada; etiquetas, m&aacute;x 200 c cada)</span><br>
</fieldset>
  </fieldset>

    </fieldset>




<br>
  <fieldset><legend><strong>Clasificaci&oacute;n y etiquetado (Reglamento 1272/2008)</strong></legend>S&iacute;mbolos:
  <span class="exp">(max 50c.)</span>
  <input maxlength="50" size="35" name="simbolos_rd1272" value="<%=simbolos_rd1272%>" /><br>  <br>
  <fieldset><legend><strong>Frases H</strong></legend>
Clasificaci&oacute;n 01 <input value="<%=clasificacion_rd1272_1%>" maxlength="100" size="70" name="clasificacion_rd1272_1" />
<br>Clasificaci&oacute;n 02 <input value="<%=clasificacion_rd1272_2%>" maxlength="100" size="70" name="clasificacion_rd1272_2">
<br>Clasificaci&oacute;n 03 <input value="<%=clasificacion_rd1272_3%>" maxlength="100" size="70" name="clasificacion_rd1272_3">
<br>Clasificaci&oacute;n 04 <input value="<%=clasificacion_rd1272_4%>" maxlength="100" size="70" name="clasificacion_rd1272_4">
<br>Clasificaci&oacute;n 05 <input value="<%=clasificacion_rd1272_5%>" maxlength="100" size="70" name="clasificacion_rd1272_5">
<br>Clasificaci&oacute;n 06 <input value="<%=clasificacion_rd1272_6%>" maxlength="100" size="70" name="clasificacion_rd1272_6">
<br>Clasificaci&oacute;n 07 <input value="<%=clasificacion_rd1272_7%>" maxlength="100" size="70" name="clasificacion_rd1272_7">
<br>Clasificaci&oacute;n 08 <input value="<%=clasificacion_rd1272_8%>" maxlength="100" size="70" name="clasificacion_rd1272_8">
<br>Clasificaci&oacute;n 09 <input value="<%=clasificacion_rd1272_9%>" maxlength="100" size="70" name="clasificacion_rd1272_9">
<br>Clasificaci&oacute;n 10 <input value="<%=clasificacion_rd1272_10%>" maxlength="100" size="70" name="clasificacion_rd1272_10">
<br>Clasificaci&oacute;n 11 <input value="<%=clasificacion_rd1272_11%>" maxlength="100" size="70" name="clasificacion_rd1272_11">
<br>Clasificaci&oacute;n 12 <input value="<%=clasificacion_rd1272_12%>" maxlength="100" size="70" name="clasificacion_rd1272_12">
<br>Clasificaci&oacute;n 13 <input value="<%=clasificacion_rd1272_13%>" maxlength="100" size="70" name="clasificacion_rd1272_13">
<br>Clasificaci&oacute;n 14 <input value="<%=clasificacion_rd1272_14%>" maxlength="100" size="70" name="clasificacion_rd1272_14">
<br>Clasificaci&oacute;n 15 <input value="<%=clasificacion_rd1272_15%>" maxlength="100" size="70" name="clasificacion_rd1272_15">

<br>

  <span class="exp">(max 100c. cada una)</span><br>  </fieldset>  <br>

  <fieldset><legend><strong>Etiquetado CLP</strong></legend>
Concentraci&oacute;n 01&nbsp;<input value="<%=conc_rd1272_1%>" maxlength="25" size="25" name="conc_rd1272_1"> &nbsp;Etiquetado 01&nbsp;<input value="<%=eti_conc_rd1272_1%>" maxlength="100" size="35" name="eti_conc_rd1272_1">
<br>Concentraci&oacute;n 02&nbsp;<input value="<%=conc_rd1272_2%>" maxlength="25" size="25" name="conc_rd1272_2"> &nbsp;Etiquetado 02&nbsp;<input value="<%=eti_conc_rd1272_2%>" maxlength="100" size="35" name="eti_conc_rd1272_2">
<br>Concentraci&oacute;n 03&nbsp;<input value="<%=conc_rd1272_3%>" maxlength="25" size="25" name="conc_rd1272_3"> &nbsp;Etiquetado 03&nbsp;<input value="<%=eti_conc_rd1272_3%>" maxlength="100" size="35" name="eti_conc_rd1272_3">
<br>Concentraci&oacute;n 04&nbsp;<input value="<%=conc_rd1272_4%>" maxlength="25" size="25" name="conc_rd1272_4"> &nbsp;Etiquetado 04&nbsp;<input value="<%=eti_conc_rd1272_4%>" maxlength="100" size="35" name="eti_conc_rd1272_4">
<br>Concentraci&oacute;n 05&nbsp;<input value="<%=conc_rd1272_5%>" maxlength="25" size="25" name="conc_rd1272_5"> &nbsp;Etiquetado 05&nbsp;<input value="<%=eti_conc_rd1272_5%>" maxlength="100" size="35" name="eti_conc_rd1272_5">
<br>Concentraci&oacute;n 06&nbsp;<input value="<%=conc_rd1272_6%>" maxlength="25" size="25" name="conc_rd1272_6"> &nbsp;Etiquetado 06&nbsp;<input value="<%=eti_conc_rd1272_6%>" maxlength="100" size="35" name="eti_conc_rd1272_6">
<br>Concentraci&oacute;n 07&nbsp;<input value="<%=conc_rd1272_7%>" maxlength="25" size="25" name="conc_rd1272_7"> &nbsp;Etiquetado 07&nbsp;<input value="<%=eti_conc_rd1272_7%>" maxlength="100" size="35" name="eti_conc_rd1272_7">
<br>Concentraci&oacute;n 08&nbsp;<input value="<%=conc_rd1272_8%>" maxlength="25" size="25" name="conc_rd1272_8"> &nbsp;Etiquetado 08&nbsp;<input value="<%=eti_conc_rd1272_8%>" maxlength="100" size="35" name="eti_conc_rd1272_8">
<br>Concentraci&oacute;n 09&nbsp;<input value="<%=conc_rd1272_9%>" maxlength="25" size="25" name="conc_rd1272_9"> &nbsp;Etiquetado 09&nbsp;<input value="<%=eti_conc_rd1272_9%>" maxlength="100" size="35" name="eti_conc_rd1272_9">
<br>Concentraci&oacute;n 10 <input value="<%=conc_rd1272_10%>" maxlength="25" size="25" name="conc_rd1272_10"> &nbsp;Etiquetado 10 <input value="<%=eti_conc_rd1272_10%>" maxlength="100" size="35" name="eti_conc_rd1272_10">
<br>Concentraci&oacute;n 11 <input value="<%=conc_rd1272_11%>" maxlength="25" size="25" name="conc_rd1272_11"> &nbsp;Etiquetado 11 <input value="<%=eti_conc_rd1272_11%>" maxlength="100" size="35" name="eti_conc_rd1272_11">
<br>Concentraci&oacute;n&nbsp;12 <input value="<%=conc_12%>" maxlength="25" size="25" name="conc_rd1272_12"> &nbsp;Etiquetado 12 <input value="<%=eti_conc_rd1272_12%>" maxlength="100" size="35" name="eti_conc_rd1272_12">
<br>Concentraci&oacute;n 13 <input value="<%=conc_rd1272_13%>" maxlength="25" size="25" name="conc_rd1272_13"> &nbsp;Etiquetado 13 <input value="<%=eti_conc_rd1272_13%>" maxlength="100" size="35" name="eti_conc_rd1272_13">
<br>Concentraci&oacute;n 14 <input value="<%=conc_rd1272_14%>" maxlength="25" size="25" name="conc_rd1272_14"> &nbsp;Etiquetado 14 <input value="<%=eti_conc_rd1272_14%>" maxlength="100" size="35" name="eti_conc_rd1272_14">
<br>Concentraci&oacute;n 15 <input value="<%=conc_rd1272_15%>" maxlength="25" size="25" name="conc_rd1272_15"> &nbsp;Etiquetado 15 <input value="<%=eti_conc_rd1272_15%>" maxlength="100" size="35" name="eti_conc_rd1272_15">
<br>  <br class="exp">  <span class="exp">(concentraciones, max 25 c cada; etiquetas, m&aacute;x 50 c cada)</span><br>  </fieldset>

Notas <span class="exp">(max 2500 c.)</span><br>  <textarea cols="60" rows="3" name="notas_rd1272"><%=notas_rd1272%></textarea><br>  <br>

  </fieldset>

    </fieldset>



  <p><input type="submit" value="Enviar" class="centcontenido"  /></p>
  </form>

<script language="JavaScript" type="text/javascript">
var frmvalidator = new Validator("myform");
frmvalidator.addValidation("nombre","req","El nombre de la sustancia es obligatorio.");
frmvalidator.addValidation("nombre","maxlen=1000");
frmvalidator.addValidation("nombre_ing","maxlen=1000");
frmvalidator.addValidation("notas_xml","maxlen=2500");
frmvalidator.addValidation("num_rd","maxlen=50");
frmvalidator.addValidation("num_ce_einecs","maxlen=50");
frmvalidator.addValidation("num_ce_elincs","maxlen=50");
frmvalidator.addValidation("formula_molecular","maxlen=100");
frmvalidator.addValidation("simbolos","maxlen=50");
frmvalidator.addValidation("cas_alternativos","maxlen=1000");
frmvalidator.addValidation("clasificacion_1","maxlen=100");
frmvalidator.addValidation("clasificacion_2","maxlen=100");
frmvalidator.addValidation("clasificacion_3","maxlen=100");
frmvalidator.addValidation("clasificacion_4","maxlen=100");
frmvalidator.addValidation("clasificacion_5","maxlen=100");
frmvalidator.addValidation("clasificacion_6","maxlen=100");
frmvalidator.addValidation("clasificacion_7","maxlen=100");
frmvalidator.addValidation("clasificacion_8","maxlen=100");
frmvalidator.addValidation("clasificacion_9","maxlen=100");
frmvalidator.addValidation("clasificacion_10","maxlen=100");
frmvalidator.addValidation("clasificacion_11","maxlen=100");
frmvalidator.addValidation("clasificacion_12","maxlen=100");
frmvalidator.addValidation("clasificacion_13","maxlen=100");
frmvalidator.addValidation("clasificacion_14","maxlen=100");
frmvalidator.addValidation("clasificacion_15","maxlen=100");
frmvalidator.addValidation("frases_s","maxlen=100");
frmvalidator.addValidation("conc_1","maxlen=100");
frmvalidator.addValidation("conc_2","maxlen=100");
frmvalidator.addValidation("conc_3","maxlen=100");
frmvalidator.addValidation("conc_4","maxlen=100");
frmvalidator.addValidation("conc_5","maxlen=100");
frmvalidator.addValidation("conc_6","maxlen=100");
frmvalidator.addValidation("conc_7","maxlen=100");
frmvalidator.addValidation("conc_8","maxlen=100");
frmvalidator.addValidation("conc_9","maxlen=100");
frmvalidator.addValidation("conc_10","maxlen=100");
frmvalidator.addValidation("conc_11","maxlen=100");
frmvalidator.addValidation("conc_12","maxlen=100");
frmvalidator.addValidation("conc_13","maxlen=100");
frmvalidator.addValidation("conc_14","maxlen=100");
frmvalidator.addValidation("conc_15","maxlen=100");
frmvalidator.addValidation("eti_conc_1","maxlen=200");
frmvalidator.addValidation("eti_conc_2","maxlen=200");
frmvalidator.addValidation("eti_conc_3","maxlen=200");
frmvalidator.addValidation("eti_conc_4","maxlen=200");
frmvalidator.addValidation("eti_conc_5","maxlen=200");
frmvalidator.addValidation("eti_conc_6","maxlen=200");
frmvalidator.addValidation("eti_conc_7","maxlen=200");
frmvalidator.addValidation("eti_conc_8","maxlen=200");
frmvalidator.addValidation("eti_conc_9","maxlen=200");
frmvalidator.addValidation("eti_conc_10","maxlen=200");
frmvalidator.addValidation("eti_conc_11","maxlen=200");
frmvalidator.addValidation("eti_conc_12","maxlen=200");
frmvalidator.addValidation("eti_conc_13","maxlen=200");
frmvalidator.addValidation("eti_conc_14","maxlen=200");
frmvalidator.addValidation("eti_conc_15","maxlen=200");

frmvalidator.addValidation("clasificacion_rd1272_1","maxlen=100");
frmvalidator.addValidation("clasificacion_rd1272_2","maxlen=100");
frmvalidator.addValidation("clasificacion_rd1272_3","maxlen=100");
frmvalidator.addValidation("clasificacion_rd1272_4","maxlen=100");
frmvalidator.addValidation("clasificacion_rd1272_5","maxlen=100");
frmvalidator.addValidation("clasificacion_rd1272_6","maxlen=100");
frmvalidator.addValidation("clasificacion_rd1272_7","maxlen=100");
frmvalidator.addValidation("clasificacion_rd1272_8","maxlen=100");
frmvalidator.addValidation("clasificacion_rd1272_9","maxlen=100");
frmvalidator.addValidation("clasificacion_rd1272_10","maxlen=100");
frmvalidator.addValidation("clasificacion_rd1272_11","maxlen=100");
frmvalidator.addValidation("clasificacion_rd1272_12","maxlen=100");
frmvalidator.addValidation("clasificacion_rd1272_13","maxlen=100");
frmvalidator.addValidation("clasificacion_rd1272_14","maxlen=100");
frmvalidator.addValidation("clasificacion_rd1272_15","maxlen=100");
frmvalidator.addValidation("conc_rd1272_1","maxlen=25");
frmvalidator.addValidation("conc_rd1272_2","maxlen=25");
frmvalidator.addValidation("conc_rd1272_3","maxlen=25");
frmvalidator.addValidation("conc_rd1272_4","maxlen=25");
frmvalidator.addValidation("conc_rd1272_5","maxlen=25");
frmvalidator.addValidation("conc_rd1272_6","maxlen=25");
frmvalidator.addValidation("conc_rd1272_7","maxlen=25");
frmvalidator.addValidation("conc_rd1272_8","maxlen=25");
frmvalidator.addValidation("conc_rd1272_9","maxlen=25");
frmvalidator.addValidation("conc_rd1272_10","maxlen=25");
frmvalidator.addValidation("conc_rd1272_11","maxlen=25");
frmvalidator.addValidation("conc_rd1272_12","maxlen=25");
frmvalidator.addValidation("conc_rd1272_13","maxlen=25");
frmvalidator.addValidation("conc_rd1272_14","maxlen=25");
frmvalidator.addValidation("conc_rd1272_15","maxlen=25");
frmvalidator.addValidation("eti_conc_rd1272_1","maxlen=100");
frmvalidator.addValidation("eti_conc_rd1272_2","maxlen=100");
frmvalidator.addValidation("eti_conc_rd1272_3","maxlen=100");
frmvalidator.addValidation("eti_conc_rd1272_4","maxlen=100");
frmvalidator.addValidation("eti_conc_rd1272_5","maxlen=100");
frmvalidator.addValidation("eti_conc_rd1272_6","maxlen=100");
frmvalidator.addValidation("eti_conc_rd1272_7","maxlen=100");
frmvalidator.addValidation("eti_conc_rd1272_8","maxlen=100");
frmvalidator.addValidation("eti_conc_rd1272_9","maxlen=100");
frmvalidator.addValidation("eti_conc_rd1272_10","maxlen=100");
frmvalidator.addValidation("eti_conc_rd1272_11","maxlen=100");
frmvalidator.addValidation("eti_conc_rd1272_12","maxlen=100");
frmvalidator.addValidation("eti_conc_rd1272_13","maxlen=100");
frmvalidator.addValidation("eti_conc_rd1272_14","maxlen=100");
frmvalidator.addValidation("eti_conc_rd1272_15","maxlen=100");
frmvalidator.addValidation("notas_rd1272","maxlen=2500");

frmvalidator.setAddnlValidationFunction("DoCustomValidation");
</script>

</div>
</body>
</html>


