<%
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'muestra residuos paginados; pulsando sobre CODIGO o RESIDUO, se ordena por ese campo
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>

<!--#include file="../adovbs.inc"-->
<!--#include file="../dn_conexion.asp"-->
<!--#include file="../dn_funciones_comunes.asp"-->
<!--#include file="../dn_funciones_texto.asp"-->
<!--#include file="../dn_restringida.asp"-->


<%
'si busc está vacio, mostramos formulario; si es 1, han dado a "buscar"; si es dos, han dado a paginación 
busc = request.form("busc")
busc = EliminaInyeccionSQL(busc)
texto = request.form("texto")
texto = EliminaInyeccionSQL(texto)
%>
<%
	'aqui no hay busqueda, siempre se muestran resultados, indicamos valores por defecto si no los hay
	ordenacion=request("ordenacion")
	ordenacion = EliminaInyeccionSQL(ordenacion)
	if ordenacion="" then ordenacion="codigo"
	sentido=request("sentido")
	sentido = EliminaInyeccionSQL(sentido)
	'if sentido="" then sentido="" 
	nregs=request("nregs") 
	nregs = EliminaInyeccionSQL(nregs)
	if nregs="" then nregs=100
	
	if busc="" then busc=1 

	if isnumeric(nregs) then
		nregs=round(nregs,0)
	else
		nregs=100
	end if
	
		
	select case busc
	
	case 1: 'han dado a buscar

		' Solo mostramos los residuos que tienen ficheros asociados, mediante un inner join
		
		'sqls="select distinct dn_alter_sectores.id as id, dn_alter_sectores.numero_cnae, dn_alter_sectores.nombre from dn_alter_sectores inner join dn_alter_ficheros_por_sectores on dn_alter_sectores.id = dn_alter_ficheros_por_sectores.id_sector order by " &ordenacion& " " &sentido 
		'sqls="select distinct dn_alter_residuos.id as id, dn_alter_residuos.codigo, dn_alter_residuos.nombre from dn_alter_residuos inner join dn_alter_ficheros_por_residuos on dn_alter_residuos.id = dn_alter_ficheros_por_residuos.id_residuo order by " &ordenacion& " " &sentido 
		sqls="select distinct id,codigo,nombre from rq_residuos where codigo like '%"&texto&"%' or nombre like '%"&texto&"%' order by " &ordenacion& " " &sentido 
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



		' Solo mostramos los residuos que tienen ficheros asociado, mediante inner join


		'sqlpag="select distinct dn_alter_sectores.id as id, numero_cnae, nombre from dn_alter_sectores inner join dn_alter_ficheros_por_sectores on dn_alter_sectores.id = dn_alter_ficheros_por_sectores.id_sector ORDER BY " &ordenacion&  " " &sentido &", dn_alter_sectores.id"
		'sqlpag="select distinct dn_alter_residuos.id as id, codigo, nombre from dn_alter_residuos inner join dn_alter_ficheros_por_residuos on dn_alter_residuos.id = dn_alter_ficheros_por_residuos.id_residuo ORDER BY " &ordenacion&  " " &sentido &", dn_alter_residuos.id"
		sqlpag="select distinct rq_residuos.id as id,codigo,nombre from rq_residuos where codigo like '%"&texto&"%' or nombre like '%"&texto&"%' order by " &ordenacion& " " &sentido 
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
						  tablares=tablares & "<td class='celda_risctox'><a href='dn_alternativas_ficha_residuo.asp?id=" &arrayDatos(0,contadorFilas)& "'>" &corta(arrayDatos(1,contadorFilas),1000, "puntossuspensivos")& "</a></td><td class='celda_risctox'><a href='dn_alternativas_ficha_residuo.asp?id=" &arrayDatos(0,contadorFilas)& "'>" &corta(arrayDatos(2,contadorFilas),1000, "puntossuspensivos")& "</a></td>"							
						  tablares=tablares & "</tr>" 
					'end if				
			next
		end if
		rstpag.close
		set rstpag=nothing
		
		tablares="<table class='tabla3' width='90%' align='center' border='0' cellpadding='4' cellspacing='0'><tr><th nowrap><a href='dn_alternativas_residuos.asp?ordenacion=codigo&sentido=ASC' title='Pulse para ordenar por Código'>CÓDIGO LER</a></th><th><a href='dn_alternativas_residuos.asp?ordenacion=nombre&sentido=ASC' title='Pulse para ordenar por nombre'>RESIDUO</a></th></tr>" &tablares& "</table>"
		
	end if

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

<link rel="stylesheet" type="text/css" href="../estructura.css">

<link rel="stylesheet" type="text/css" href="../dn_estilos.css">
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
<td align='right'><input type="button" name="volver" class="boton2" value="Volver a la portada de Alternativas" onclick="window.location='../dn_alternativas_portada.asp';"></td>

</tr>

</table>

<form class="texto" action="dn_alternativas_residuos.asp" method="post">
<fieldset><legend><strong>Código LER o nombre del residuo</strong></legend>
<p class="texto">Introduzca el texto a buscar, o deje el campo en blanco si prefiere que se muestren todos los resultados:</p>
<input type="hidden" name="busc" value="1" />
<input name="texto" type="text" value="<%=texto%>" size="40" />
<input type="submit" value="Buscar" class="boton2" />
</fieldset>
</form>

<p class=titulo3>Residuos</p>

<form action="dn_alternativas_residuos.asp?busc=1" method="post" name="myform">
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
