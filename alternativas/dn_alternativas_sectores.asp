<%
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'muestra sectores paginados; pulsando sobre CNAE o SECTOR, se ordena por ese campo
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>

<!--#include file="../adovbs.inc"-->
<!--#include file="../dn_conexion.asp"-->
<!--#include file="../dn_funciones_comunes.asp"-->
<!--#include file="../dn_funciones_texto.asp"-->
<!--#include file="../dn_restringida.asp"-->

<%
'si busc está vacio, mostramos formulario; si es 1, han dado a "buscar"; si es dos, han dado a paginación 
busc=EliminaInyeccionSQL(request.form("busc"))
%>
<%
	'aqui no hay busqueda, siempre se muestran resultados, indicamos valores por defecto si no los hay
	ordenacion=EliminaInyeccionSQL(request("ordenacion"))
	if ordenacion="" then ordenacion="numero_cnae"
	sentido=EliminaInyeccionSQL(request("sentido"))
	'if sentido="" then sentido="" 
	nregs=EliminaInyeccionSQL(request("nregs"))
	if nregs="" then nregs=100
	
	if busc="" then busc=1 

	if isnumeric(nregs) then
		nregs=round(nregs,0)
	else
		nregs=100
	end if
	
		
	select case busc
	
	case 1: 'han dado a buscar
		' Solo mostramos los sectores que tienen ficheros asociados, mediante un inner join
		
		sqls="select distinct dn_alter_sectores.id as id, dn_alter_sectores.numero_cnae, dn_alter_sectores.nombre from dn_alter_sectores inner join dn_alter_ficheros_por_sectores on dn_alter_sectores.id = dn_alter_ficheros_por_sectores.id_sector order by " &ordenacion& " " &sentido 
		'response.write sqls
		
			Set objRst = Server.CreateObject("ADODB.Recordset")
			objRst.Open sqls, objConnection2, adOpenStatic, adCmdText
			hr=objRst.recordcount
		
			IF not objRst.eof THEN 
				'arr=objRst.GetString(adClipString, 1, "", ",", "")		
				arrayDatos=objRst.getrows	
				
				for I = 0 to UBound(arrayDatos,2) 
					arr=arr& arrayDatos(0,I) & ","
				next	
				'esta sera la pagina 1
				pag = 1	
			END IF
					
			objRst.Close
			Set objRst=Nothing

	case 2: 'paginando
		
		hr=EliminaInyeccionSQL(request("hr"))
		pag=EliminaInyeccionSQL(request("pag"))
		arr=EliminaInyeccionSQL(request("arr"))
					
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

		' Solo mostramos los sectores que tienen ficheros asociado, mediante inner join

		sqlpag="select distinct dn_alter_sectores.id as id, numero_cnae, nombre from dn_alter_sectores inner join dn_alter_ficheros_por_sectores on dn_alter_sectores.id = dn_alter_ficheros_por_sectores.id_sector ORDER BY " &ordenacion&  " " &sentido &", dn_alter_sectores.id"
		set rstpag=objConnection2.execute(sqlpag)
		if not rstpag.eof then
			
			arrayDatos = rstpag.GetRows			
      		
			registroini=(pag*nregs)-nregs
			registrofin=registroini+nregs
			
			if registrofin>=hr-1 then
				registrofin=hr
			end if
			
			registrofin=registrofin-1

			for contadorFilas=registroini to registrofin				
				
					'if contadorfilas>=hr-1 then
							'exit for
					'else
						  'arrayDatos(0,contadorFilas)
						  tablares=tablares & "<tr>" 
						  'tablares=tablares & "<td>" & contadorFilas+1 & "</td>"
						  tablares=tablares & "<td class='celda_risctox'><a href='dn_alternativas_ficha_sector.asp?id=" &arrayDatos(0,contadorFilas)& "'>" &corta(arrayDatos(1,contadorFilas),1000, "puntossuspensivos")& "</a></td><td class='celda_risctox'><a href='dn_alternativas_ficha_sector.asp?id=" &arrayDatos(0,contadorFilas)& "'>" &corta(arrayDatos(2,contadorFilas),1000, "puntossuspensivos")& "</a></td>"							
						  tablares=tablares & "</tr>" 
					'end if				
			next
		end if
		rstpag.close
		set rstpag=nothing
		
		tablares="<table class='tabla3' width='90%' align='center' border='0' cellpadding='4' cellspacing='0'><tr><th nowrap><a href='dn_alternativas_sectores.asp?ordenacion=numero_cnae&sentido=ASC' title='Pulse para ordenar por CNAE'>CNAE</a></th><th><a href='dn_alternativas_sectores.asp?ordenacion=nombre&sentido=ASC' title='Pulse para ordenar por nombre'>Sector</a></th></tr>" &tablares& "</table>"
		
	end if

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ISTAS: BBDD Alternativas</title>
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

<link rel="stylesheet" type="text/css" href="../estructura.css">
<link rel="stylesheet" type="text/css" href="dn_estilos.css">
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
		<!--#include file="../dn_cabecera.asp"-->
		<div id="texto">
			
<div class="texto">
<!-- ################ CONTENIDO ###################### -->

<table width="100%" border="0">
<tr>
<td></td>
<td align='right'><input type="button" name="volver" class="boton" value="Volver a la portada de Alternativas" onClick="window.location='./index.asp';"></td>
</tr>
</table>

<p class=titulo3>Sectores destacados</p>
<p align="center">

<input type="button" class="boton2" value="Artes gráficas" onClick="window.location='dn_alternativas_ficha_sector_destacado.asp?s=artesgraficas';">&nbsp;
<input type="button" class="boton2" value="Papel" onClick="window.location='dn_alternativas_ficha_sector_destacado.asp?s=papel';">&nbsp;
<input type="button" class="boton2" value="Madera" onClick="window.location='dn_alternativas_ficha_sector_destacado.asp?s=madera';">&nbsp;
<input type="button" class="boton2" value="Construcción" onClick="window.location='dn_alternativas_ficha_sector_destacado.asp?s=construccion';">&nbsp;<input type="button" class="boton2" value="Textil" onClick="window.location='dn_alternativas_ficha_sector_destacado.asp?s=textil';">&nbsp;<input type="button" class="boton2" value="Limpieza" onClick="window.location='dn_alternativas_ficha_sector_destacado.asp?s=limpieza';">&nbsp;

<!--
	<ul>
			<li><a href="dn_alternativas_ficha_sector_destacado.asp?s=artesgraficas">Artes gráficas</a></li>
			<li><a href="dn_alternativas_ficha_sector_destacado.asp?s=papel">Papel</a></li>
			<li><a href="dn_alternativas_ficha_sector_destacado.asp?s=madera">Madera</a></li>
			<li><a href="dn_alternativas_ficha_sector_destacado.asp?s=construccion">Construcción</a></li>
	</ul>
-->
</p>

<p class=titulo3>Sectores</p>

<form action="dn_alternativas_sectores.asp?busc=1" method="post" name="myform">
 <input type="hidden" name='busc' value='<%=busc%>' />	
 <input type="hidden" name='pag' value='<%=pag%>' />	
 <input type="hidden" name='hr' value='<%=hr%>' />		
 <input type="hidden" name='arr' value='<%=arr%>' />
 <input type="hidden" name='nregs' value='<%=nregs%>' />				
 <input type="hidden" name='ordenacion' value='<%=ordenacion%>' />				
 <input type="hidden" name='sentido' value='<%=sentido%>' />				


<%
		response.Write("<p class='neg' style='margin:15px 0; padding:10px;'>Se han encontrado " &hr& " registros. Se muestran registros del " &registroini+1& " al " &registrofin+1& ":</p>")
%>		
		<%=tablares%>
		<div align='center' style="margin:20px 10px; background-color: #3399CC; padding:3px;"><%paginacion%></div>
</form>


<!-- ############ FIN DE CONTENIDO ################## -->
<!--#include file="spl_pie.inc.asp"-->

<%
cerrarconexion
%>
