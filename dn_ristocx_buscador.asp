<!--#include file="adovbs.inc"-->
<!--#include file="dn_conexion.asp"-->
<!--#include file="dn_funciones_comunes.asp"-->
<!--#include file="dn_funciones_texto.asp"-->

<%
'si busc está vacio, mostramos formulario; si es 1, han dado a "buscar"; si es dos, han dado a paginación 
busc=request.form("busc")
busc = EliminaInyeccionSQL(busc)

if busc="" then
	'valores de busqueda por defecto
	ordenacion="sus.nombre"
	sentido=""
	nregs=20
else
	ordenacion=request("ordenacion")
	sentido=request("sentido") 
	nregs=request("nregs") 
	ordenacion = EliminaInyeccionSQL(ordenacion)
	sentido = EliminaInyeccionSQL(sentido)
	nregs = EliminaInyeccionSQL(nregs)
	if isnumeric(nregs) then
		nregs=round(nregs,0)
	else
		nregs=20
	end if
	
	nombre=request.form("nombre")
	tipobus=request.form("tipobus")
	numero=request.form("numero")
	nombre = EliminaInyeccionSQL(nombre)
	tipobus = EliminaInyeccionSQL(tipobus)
	numero = EliminaInyeccionSQL(numero)
	
	select case busc
	
	case 1: 'han dado a buscar
					
		condicion=""	
		
			if nombre<>"" or numero<>"" then
			condicion=" WHERE "
			if nombre<>"" then	
				nombre2=h(nombre)	
				nombre2=quitartildes(nombre2)
				nombre2=montartildes(nombre2)
				if tipobus="exacto" then
					condicion=condicion& " (sus.nombre='" &nombre2& "' or sin.nombre='" &nombre2& "')  "
				else
					condicion=condicion& " (sus.nombre like '%" &nombre2& "%' or sin.nombre like '%" &nombre2& "%')  "
				end if
			end if
			if numero<>"" then
				if nombre<>"" then condicion=condicion& " OR "
				condicion=condicion& " (num_ce_einecs = '" &numero& "' OR num_ce_elincs  = '" &numero& "' OR  num_rd = '" &numero& "' OR  num_cas = '" &numero& "')"
			end if
		end if
		sqls="select sus.id from dn_risc_sustancias as sus FULL OUTER JOIN dn_risc_sinonimos as sin ON (sus.id=sin.id_sustancia) " &condicion& " ORDER BY " &ordenacion&  " " &sentido
		'response.write sqls
		
			Set objRst = Server.CreateObject("ADODB.Recordset")
			objRst.Open sqls, objConnection2, adOpenStatic, adCmdText
			hr=objRst.recordcount
		
			IF not objRst.eof THEN 
				arr=objRst.GetString(adClipString, -1, "", ",", "")				
				'esta sera la pagina 1
				pag = 1	
			END IF
					
			objRst.Close
			Set objRst=Nothing

	case 2: 'paginando
		
		hr=request("hr")
		pag=request("pag")
		arr=request("arr")
		hr = EliminaInyeccionSQL(hr)
	  pag = EliminaInyeccionSQL(pag)
		arr = EliminaInyeccionSQL(arr)		
					
	end select 'cual busc
	
	'RESULTADOS DE BUSQUEDA (para busc 1 y busc 2)
	'seleccionamos datos a mostrar de los x registros que toquen
	if hr>0 then
				
		arrayx = split(arr, ",")
		
		FOR i=0 to UBound(arrayx)-1		 'OJO, luego hay que hacer que muestre solo las de paginacion
			cadenaids=cadenaids  &arrayx(i)&","		
		NEXT	
		
		'quitamos la ultima coma
		cadenaids= left(cadenaids,len(cadenaids)-1)
		sqlpag="select id, nombre from dn_risc_sustancias as sus WHERE id IN(" &cadenaids& ") ORDER BY " &ordenacion&  " " &sentido
		set rstpag=objConnection2.execute(sqlpag)
		if not rstpag.eof then
			'strDBDataTable = rstpag.GetString(adClipString, -1, "</td><td>", "</td></tr>" & vbCrLf & "<tr><td>", "&nbsp;")
			'strDBDataTable= left(strDBDataTable,len(strDBDataTable)-8)
			'tablares= "<table border='1'><tr><td>" &strDBDataTable& "</table>" 
			arrayDatos = rstpag.GetRows			
      		
			registroini=(pag*nregs)-nregs
			'response.write "<p>registroini=" &registroini& "</p>"
			
			registrofin=registroini+nregs
			'response.write "<p>registrofin=" &registrofin& "</p>"
			
			if registrofin>=hr-1 then
				registrofin=hr
      			'response.write "<p>registrofin era mayor, ahora=" &registrofin& "</p>"
			end if
			
			registrofin=registrofin-1
			'response.write "<p>registrofin-1=" &registrofin& "</p>"
			
			
			for contadorFilas=registroini to registrofin				
				
					'if contadorfilas>=hr-1 then
							'exit for
					'else
						  'arrayDatos(0,contadorFilas)
						  tablares=tablares & "<tr>" 
						  'tablares=tablares & "<td>" & contadorFilas+1 & "</td>"
						  tablares=tablares & "<td><a href='dn_risctox_ficha_sustancia.asp?id_sustancia=" &arrayDatos(0,contadorFilas)& "'>" &corta(arrayDatos(1,contadorFilas),100, "puntossuspensivos")& "</a><br />" &dameSinonimos(arrayDatos(0,contadorFilas))& "</td>"							
						  tablares=tablares & "</tr>" 
					'end if				
			next
		end if
		rstpag.close
		set rstpag=nothing
		tablares="<table id='resultados' cellspacing='0' cellpadding='3' style='margin:0 10px;'>" &tablares& "</table>" 
	end if

end if 'busc<>""

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
	frm.busc.value=1;
	frm.pag.value=1;
	frm.submit();
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

<p class=campo>Est&aacute;s en: <a href=index.asp?idpagina=550>Plataforma prevención de riesgo químico</a>&nbsp;&gt;&nbsp;BBDD Sustancias / Buscador </p>

<p class=titulo3>Buscador de Sustancias </p>

<%

%>

<form action="dn_ristocx_buscador.asp?busc=1" method="post" name="myform" onSubmit="primerapag();">
 <input type="hidden" name='busc' value='<%=busc%>' />	
 <input type="hidden" name='pag' value='<%=pag%>' />	
 <input type="hidden" name='hr' value='<%=hr%>' />		
 <input type="hidden" name='arr' value='<%=arr%>' />
 <input type="hidden" name='ordenacion' value='<%=ordenacion%>' />
 <input type="hidden" name='nregs' value='<%=nregs%>' />				
<table border="0" align="center" cellspacing="5">
	<tr>
		<td><strong>Nombre</strong></td>
		<td><input type="text" name="nombre" value="<%=nombre%>" /></td>
		<td><select name="tipobus">
		<option value="exacto" <%if tipobus="exacto" then response.write "selected"%>>exacto</option>
		<option value="parte" <%if tipobus="parte" then response.write "selected"%>>parte</option>
		</select></td>
	</tr>
	<tr>
		<td><strong>Número CAS/CE/RD</strong></td>
		<td><input type="text" name="numero" value="<%=numero%>" /></td>
		<td></td>
	</tr>	
	<tr>
		<td colspan="2" align="center"><input type="button" value="Buscar" onclick="primerapag();" /> <input type="reset" value="Borrar" /></td>
	</tr>
</table>

<%
if busc<>"" then
	if hr=0  then
%>
		<fieldset id="flashmsg"><legend class="advertencia"><strong>Advertencia</strong></legend>No se encontraron registros que coincidan con su consulta.</fieldset>
<%
	else
		response.Write("<p class='neg' style='margin:15px 0; padding:10px;'>Se han encontrado " &hr& " registros. Se muestran registros del " &registroini+1& " al " &registrofin+1& ":</p>")
%>		
		<%=tablares%>
		<div align='center' style="margin:20px 10px; background-color: #3399CC; padding:3px;"><%paginacion%></div>
<%
	end if
end if
%>
</form>


<!-- ############ FIN DE CONTENIDO ################## -->



<p align=center><object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0" width=530 height=100 id="anima0" align="middle"><param name="allowScriptAccess" value="sameDomain" /><param name="movie" value="http://www.istas.net/recursos/ANI/ISTAS_01076.swf" /><param name="quality" value="high" /><param name="wmode" value="transparent" /><param name="bgcolor" value="#ffffff" /><embed src="http://www.istas.net/recursos/ANI/ISTAS_01076.swf" quality="high" wmode="transparent" bgcolor="#ffffff" width=530 height=100 name="" align="middle" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" /></object></p><p> </p><br>
<br>
<br>
Esta página ha sido desarrollada por <a href="http://www.istas.ccoo.es/" target="_blank"><b>ISTAS</b></a> que es una Fundación de <a href="http://www.ccoo.es/" target="_blank"><font color="#FF0000"><b>CC.OO.</b></font></a><br>

				
				</div>
				<p>&nbsp;</p>
			</div>
			
			
			<map name="Map1" id="Map1">
            		<area shape="rect" coords="307,38,399,104" href="http://www.fundacion-biodiversidad.es" target="_blank" alt="Fundación Biodiversidad" />
            		<area shape="rect" coords="400,38,546,104" href="http://www.istas.ccoo.es" target="_blank" alt="Instituto Sindical de Trabajo, Ambiente y Salud" />
      			<area shape="rect" coords="547,38,701,104" href="http://www.mtas.es/UAFSE/default.htm" target="_blank" alt="Fondo Social Europeo" />
      			</map>

			<map name="Map2" id="Map2">
            		<area shape="rect" coords="300,8,392,66" href="http://www.fundacion-biodiversidad.es" target="_blank" alt="Fundación Biodiversidad" />
            		<area shape="rect" coords="393,8,539,66" href="http://www.istas.ccoo.es" target="_blank" alt="Instituto Sindical de Trabajo, Ambiente y Salud" />
      			<area shape="rect" coords="540,8,694,66" href="http://www.mtas.es/UAFSE/default.htm" target="_blank" alt="Fondo Social Europeo" />
      			</map>
      			<map name="Map3" id="Map3">
            		<area shape="rect" coords="300,18,392,80" href="http://www.fundacion-biodiversidad.es" target="_blank" alt="Fundación Biodiversidad" />
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
