<%
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'muestra buscador y resultados para alternativas y casos practicos (para estos ultimos, se pasa el parametro cp=1)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>

<!--#include file="../adovbs.inc"-->
<!--#include file="../dn_conexion.asp"-->
<!--#include file="../dn_funciones_comunes.asp"-->
<!--#include file="../dn_funciones_texto.asp"-->
<!--#include file="../dn_restringida.asp"-->
<!--#include file="../EliminaInyeccionSQL.asp"-->

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

<%
'valores por defecto
nregs=20

'recogemos valores del formulario
pag=EliminaInyeccionSQL(request("pag"))
texto=EliminaInyeccionSQL(request.form("texto"))
tema=EliminaInyeccionSQL(request.form("tema"))
cp=EliminaInyeccionSQL(request.QueryString("cp"))
if cp="" then cp=0
select case cp
	case 1: mititulo="Experiencias sindicales y Casos Prácticos"
	case 2: mititulo="Casos Prácticos"
	case 3: mititulo="Experiencias sindicales"
	case 4: mititulo="Experiencias sindicales / Casos"
	case 5: mititulo="Experiencias sindicales / Fichas de sustitución de disolventes"
	case 0: mititulo="Alternativas"
end select
%>
<table width="100%" border="0">
<tr>
<td></td>
<%
if cp>=3 and cp<=5 then 'EXPERIENCIAS SINDICALES (todas, los casos y las fichas de sustitución)

'<td align='right'><input type="button" name="volver" class="boton" value="Casos" onClick="window.location='./dn_alternativasycasosprac.asp?cp=3';"></td>
'<td align='right'><input type="button" name="volver" class="boton" value="Fichas de sustitución de disolventes" onClick="window.location='./dn_alternativasycasosprac.asp?cp=4';"></td>
'<td align='right'><input type="button" name="volver" class="boton" value="volver a Experiencias" onClick="window.location='./dn_alternativasycasosprac.asp?cp=1';"></td>

else

'<td align='right'><input type="button" name="volver" class="boton" value="volver a Experiencias" onClick="window.location='./dn_alternativasycasosprac.asp?cp=1';"></td>

end if
%>
<% 
	if cp>1 then 
%>
<td align='right'><input type="button" name="volver" class="boton" value="volver a Experiencias" onClick="window.location='./dn_alternativasycasosprac.asp?cp=1';"></td>
<% 
	end if
%>
<td align='right'><input type="button" name="volver" class="boton" value="Volver a la portada de Alternativas" onClick="window.location='./index.asp';"></td>
</tr>
</table>
<%
	' 
%>
<% 
	if cp=1 then 'EXPERIENCIAS 
%>
<p>En éste apartado podrás encontrar tanto <b><a href="./dn_alternativasycasosprac.asp?cp=3">experiencias sindicales</a></b> como <b><a href="./dn_alternativasycasosprac.asp?cp=2">casos prácticos (no sindicales)</a></b> de sustitución de sustancias peligrosas.</p>
<% 
	end if
%>

<% 
	if cp>=3 and cp<=6 then 'EXPERIENCIAS SINDICALES (todas, los casos y las fichas de sustitución)
%>
<p class=titulo3>Experiencias sindicales</p>

<p>En este apartado hemos diferenciado los <a href="./dn_alternativasycasosprac.asp?cp=4">casos de sustitución</a> sindicales de las <a href="./dn_alternativasycasosprac.asp?cp=5">fichas de experiencias sindicales de sustitución de disolventes</a>, para facilitar la búsqueda. 
Sin embargo, podrás realizar una <a href="./dn_alternativasycasosprac.asp?cp=3">búsqueda global en el buscador de experiencias sindicales</a>.</p>

<% 
	end if
%>
<p class=titulo3>Buscador de <%=mititulo%></p>

<div align="center">
<form class="texto" action="dn_alternativasycasosprac.asp?cp=<%=cp%>" method="post">
<fieldset><legend><strong>Texto</strong></legend>
<p class="texto">Introduzca el texto a buscar, o deje el campo en blanco si prefiere que se muestren todos los resultados:</p>
<input type="hidden" name="busc" value="1" />
<input name="texto" type="text" value="<%=texto%>" size="40" />
<input type="submit" value="Buscar" class="boton2" />
</fieldset>
</form>
<br />
<%
if cp=0 then 'DOCUMENTOS -> SELECTOR DE TEMA
%>
<form class="texto" action="dn_alternativasycasosprac.asp?busc=1&cp=<%=cp%>" method="post">
<fieldset><legend><strong>Tema</strong></legend>
<p class="texto">O bien seleccione un tema:</p>
<select name="tema" class="campo">
<option>SELECCIONE UN TEMA</option>
<%
sqll="SELECT DISTINCT tema  FROM dn_alter_ficheros order by tema"
Set rstt=objConnection2.Execute(sqll)
if not  rstt.eof then
	do while not rstt.eof
		if tema<>"" then
			if tema<>rstt("tema") then
				marcado=""
			else
				marcado="selected"
			end if
		end if
		response.write "<option value='" &rstt("tema")& "'" &marcado&">" &corta (rstt("tema"), 95, "puntossuspensivos")& "</option>" 
	rstt.movenext
	loop
end if
rstt.Close
Set rstt = Nothing
%>
</select>
<input type="hidden" name="busc" value="1" />
<input type="submit" value="Buscar" class="boton2" />
</fieldset>
</form>
<%
end if
%>

<%
'vemos si tenemos que mostrar resultados
busc=EliminaInyeccionSQL(request.form("busc"))

'para cp=1 mostramos resultados ya la primera vez que entramos
if (cp>=1 and cp<=5) and busc="" then
	busc=1
end if


if busc<>""   then
	if request("tema")="SELECCIONE UN TEMA" then
%>
		<fieldset id="flashmsg" style="width:60%; margin:50px; "><legend class="advertencia"><strong>Advertencia</strong></legend>Debe seleccionar un tema. Si lo que desea es que se muestren todos los registros, deje en blanco el campo <strong>Texto</strong> y pulse el botón <strong>Buscar</strong> que se encuentra a su derecha.</fieldset>
<%
	else
		select case busc
			
			case 1: 'han dado a buscar
							
						select case cp
							case 1: condicion=condicion& " (tema like '%Casos Prácticos%' OR tema like '%Experiencias sindicales%' )AND" 
							case 2: condicion=condicion& " tema like '%Casos Prácticos%' AND" 
							case 3: condicion=condicion& " tema like '%Experiencias sindicales%' AND" 
							case 4: condicion=condicion& " tema like '%Experiencias sindicales / Casos%' AND" 
							case 5: condicion=condicion& " tema like '%Experiencias sindicales / Fichas de sustitución de disolventes%' AND" 
																	   
						end select
'						if tema<>"" then condicion=condicion& " tema like '" &tema& "' AND"

							
'						if cp=1 then
'							 condicion=condicion& " tema like '%Casos práctico%' AND"
'						else
'							if cp=2 then
'								 condicion=condicion& " tema like '%Experiencias sindicales%' AND"
'							else
'								if tema<>"" then condicion=condicion& " tema like '" &tema& "' AND"
'							end if
'						end if						
							
						if texto<>"" then
							condicion = condicion & "("
							cualquier=replace(texto,"  "," ")		
							' Dividimos la cadena en palabras separadas por espacios
							arraycualquier = split(cualquier, " ")					
							' Montamos la condicion de texto (reemplazando caracteres)	
							FOR i=0 to UBound(arraycualquier)		
								mitermino=arraycualquier(i)		
								mitermino=quitartildes(mitermino)
								mitermino=montartildes(mitermino)
								condicion = condicion & " (titulo LIKE '%" &mitermino& "%') or (resumen LIKE '%" &mitermino& "%') or (tema LIKE '%" &mitermino& "%')  OR"
							NEXT		
							
							if condicion<>"" then condicion= left(condicion,len(condicion)-3) & ") AND"
						end if
						
						if condicion<>"" then condicion=" WHERE "  &left(condicion,len(condicion)-3) 'quitamos el último or/and	
				
						'COMPLICACION: lo primero que tenemos que hacer es un select distinc exclusivamente de titulos, ya que puede haber registros con el mismo titulo, pero distinto tema/id

				sqls="select distinct titulo, id from dn_alter_ficheros " &condicion
'				response.write sqls
				
					Set objRst = Server.CreateObject("ADODB.Recordset")
					objRst.Open sqls, objConnection2, adOpenStatic, adCmdText
					
				
					IF not objRst.eof THEN 
					
						'hacemos lista de ids
						mititulo=""
						do while not objRst.eof
							if mititulo<>objRst("titulo") then
								arr=arr& objRst("id")& ","
								mititulo=objRst("titulo")
								hr=hr+1
							 end if
						objRst.movenext
						loop
						
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
			if hr=0 then
	%>
				<fieldset id="flashmsg" style="width:60%; margin:50px; "><legend class="advertencia"><strong>Advertencia</strong></legend>No se encontraron registros que coincidan con su consulta.</fieldset>
	<%
			else
				arrayx = split(arr, ",")
				
				FOR i=0 to UBound(arrayx)-1		 'OJO, luego hay que hacer que muestre solo las de paginacion
					cadenaids=cadenaids  &arrayx(i)&","		
				NEXT	
				
				'quitamos la ultima coma
				cadenaids= left(cadenaids,len(cadenaids)-1)
				sqlpag="select id, tema, titulo from dn_alter_ficheros as sus WHERE id IN(" &cadenaids& ") ORDER BY tema, titulo" 
				'response.write "<p>" &sqlpag
				set rstpag=objConnection2.execute(sqlpag)
				if not rstpag.eof then
				
					arrayDatos = rstpag.GetRows			
					registroini=(pag*nregs)-nregs
					registrofin=registroini+nregs
					
					if registrofin>=hr-1 then
						registrofin=hr
					end if
					
					registrofin=registrofin-1
		
					' Generamos tabla de resultados, comprobando que no se repiten los títulos
					for contadorFilas=registroini to registrofin				
								  tablares=tablares & "<tr>" 
								  'tablares=tablares & "<td>" & contadorFilas+1 & "</td>"
								  'tablares=tablares & "<td class='celda_risctox'>" &corta(arrayDatos(1,contadorFilas),1000, "puntossuspensivos")& "</td>"
								  tablares=tablares & "<td class='celda_risctox' align='left'><a href='dn_alternativas_ficha_fichero.asp?id_fichero=" &arrayDatos(0,contadorFilas)& "'>" &corta(arrayDatos(2,contadorFilas),1000, "puntossuspensivos")& "</a></td>"							
								  tablares=tablares & "</tr>" 

					next
					
				end if
				rstpag.close
				set rstpag=nothing
				
				'iniciotabla="<table class='tabla3' width='90%' align='center' border='0' cellpadding='4' cellspacing='0'> <tr><td class='subtitulo3'>	<table width='100%' align='center'><tr><td>"
				'if cp<>1 then 'alternativas
					'iniciotabla=iniciotabla& "Alternativas &nbsp;</td><td align=right><!-- <img src='imagenes/ico_alt_procesos.gif'> -->"
				'else
					'iniciotabla=iniciotabla& "Casos prácticos &nbsp;</td><td align=right><!-- <img src='imagenes/ico_alt_procesos.gif'> -->"
				'end if
				'iniciotabla=iniciotabla& "</td></tr></table></td> </tr>"
		
				response.Write("<p align='left' class='neg' style='margin:15px 0; padding:10px;'>Se han encontrado " &hr& " registros. Se muestran registros del " &registroini+1& " al " &registrofin+1& ":</p>")
				'response.write "<br />"

				tablares="<table class='tabla3' width='90%' align='center' border='0' cellpadding='4' cellspacing='0'><tr><th>Título</th></tr>" &tablares& "</table>"
				response.write tablares
				
				if hr>nregs then
%>		
				<form name="myform" method="post" action="dn_alternativasycasosprac.asp?cp=<%=cp%>">
				<input type="hidden" name="texto" value="<%=texto%>" />
				<input type="hidden" name="tema" value="<%=tema%>" />
				<input type="hidden" name="pag" value="<%=pag%>" />
				<input type="hidden" name="arr" value="<%=arr%>" />
				<input type="hidden" name="hr" value="<%=hr%>" />
				<input type="hidden" name="busc" value="2" />
				<div align='center' style="margin:20px 10px; background-color: #3399CC; padding:3px;"><%paginacion%></div>
				</form>
<%
				end if
			end if
		end if
end if 'busc
%>


</div>

<!-- ############ FIN DE CONTENIDO ################## -->
<%
	variacion_pie= "_alternativas_experiencias"
%>
<!--#include file="spl_pie.inc.asp"-->

<%
cerrarconexion
%>
