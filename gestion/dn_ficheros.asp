<%
'+++++++pagina paralela de dn_sustancias
%>

<!--#include file="dn_fun_comunes.asp"-->
<!--#include file="dn_fun_texto.asp"-->
<!--#include file="dn_auten.inc"-->
<!--#include file="adovbs.inc"-->
<!--#include file="dn_conexion.asp"-->
<%'Permite auditar las acciones%>
<!--#include file="spl_auditoria.inc.asp"-->

<%
'no hay buscador

'si busc está vacio, mostramos formulario; si es 1, han dado a "buscar"; si es dos, han dado a paginación
busc=EliminaInyeccionSQL(request.form("busc"))

'siempre mostramos resultados, aunque sea la primera vez que entramos
if busc="" then
	busc=1
	'valores de busqueda por defecto
	ordenacion="num_alternativa, tema, titulo"
	sentido=""
	nregs=100
else
	ordenacion=EliminaInyeccionSQL(request("ordenacion"))
	sentido=EliminaInyeccionSQL(request("sentido"))
	nregs=EliminaInyeccionSQL(request("nregs"))
end if

if busc="" then 'aqui no va a entrar
	'valores de busqueda por defecto
	'NOMBRE=1
	ordenacion="num_alternativa, tema, titulo"
	sentido=""
	nregs=100
else

	if isnumeric(nregs) then
		nregs=round(nregs,0)
	else
		nregs=100
	end if

	titulo=h(EliminaInyeccionSQL(request.form("titulo")))
	tema=h(EliminaInyeccionSQL(request.form("tema")))
	num_alternativa=h(EliminaInyeccionSQL(request.form("num_alternativa")))

	select case busc

		case 1: 'han dado a buscar

			condicion=""

			'si mandan texto, creamos condicion
			if titulo<>"" then condicion=condicion& " OR  titulo like '%" &titulo& "%' "
			if tema<>"" then condicion=condicion& " OR  tema like '%" &tema& "%' "
			if num_alternativa<>"" then condicion=condicion& " OR num_alternativa=" &num_alternativa

			'quitamos el primer or
			if condicion<>"" then condicion= " WHERE " &right(condicion,len(condicion)-3)
			sqls="select id from dn_alter_ficheros " &condicion& " ORDER BY " &ordenacion&  " " &sentido
			'response.write sqls

				Set objRst = Server.CreateObject("ADODB.Recordset")
				objRst.Open sqls, objConn1, adOpenStatic, adCmdText
				hr=objRst.recordcount

				IF not objRst.eof THEN
					arr=objRst.GetString(adClipString, -1, "", ",", "")
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
		sqlpag="select id, tema, titulo, num_alternativa from dn_alter_ficheros WHERE id IN(" &cadenaids& ") ORDER BY " &ordenacion&  " " &sentido
		set rstpag=objconn1.execute(sqlpag)
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
						  'pijamar=true
						  if pijamar then
						  	if  (contadorFilas mod 2)=0 then
								estilotr=" class='rayado'"
							else
								estilotr=""
							end if
						  end if
						  tablares=tablares & "<tr " &estilotr& ">"
						  tablares=tablares & "<td><input type='checkbox' name='idcheck' value=" &arrayDatos(0,contadorFilas)& " /></td>"
						  tablares=tablares & "<td>" & contadorFilas+1 & "</td>"
				  		  tablares=tablares & "<td>" & arrayDatos(1,contadorFilas) & "</td>"
						  tablares=tablares & "<td><a href='#' onclick=" &chr(34)& "abreVentanaCentrada('dn_asociaciones_mostrar.asp?tipo=fichero&id=" &arrayDatos(0,contadorFilas)& "', 800, 580); return false;" &chr(34)& "	><img src='imagenes/relacion.png' align='left' alt='Ver relación' /></a>&nbsp;<a href='#' onclick=" &chr(34)& "abreVentanaCentrada('dn_fichero.asp?id=" &arrayDatos(0,contadorFilas)& "', 800, 580); return false;" &chr(34)& "	>("&arrayDatos(0,contadorFilas)&") <b>" &arrayDatos(2,contadorFilas)& "</b></a></td>"
						  tablares=tablares & "<td>" & arrayDatos(3,contadorFilas) & "</td>"
						  tablares=tablares & "<td align='center'><a href='#' onclick=" &chr(34)& "abreVentanaCentrada('dn_desasociar.asp?asociar=fich_sustancia&id=" &arrayDatos(0,contadorFilas)& "', 800, 500); return false;" &chr(34)& "	><img src='imagenes/icono_1.gif' alt='Sustancias'></a></td>"
						  tablares=tablares & "<td align='center'><a href='#' onclick=" &chr(34)& "abreVentanaCentrada('dn_desasociar.asp?asociar=fich_grupo&id=" &arrayDatos(0,contadorFilas)& "', 800, 500); return false;" &chr(34)& "	><img src='imagenes/icono_2.gif' alt='Ficheros'></a></td>"
						  tablares=tablares & "<td align='center'><a href='#' onclick=" &chr(34)& "abreVentanaCentrada('dn_desasociar.asp?asociar=fich_sector&id=" &arrayDatos(0,contadorFilas)& "', 800, 500); return false;" &chr(34)& "	><img src='imagenes/icono_3.gif' alt='Sectores'></a></td>"
						  tablares=tablares & "<td align='center'><a href='#' onclick=" &chr(34)& "abreVentanaCentrada('dn_desasociar.asp?asociar=fich_proceso&id=" &arrayDatos(0,contadorFilas)& "', 800, 500); return false;" &chr(34)& "	><img src='imagenes/icono_4.gif' alt='Procesos'></a></td>"
						  tablares=tablares & "<td align='center'><a href='#' onclick=" &chr(34)& "abreVentanaCentrada('dn_desasociar.asp?asociar=fich_uso&id=" &arrayDatos(0,contadorFilas)& "', 800, 500); return false;" &chr(34)& "	><img src='imagenes/icono_5.gif' alt='Usos'></a></td>"
						  tablares=tablares & "<td align='center'><a href='#' onclick=" &chr(34)& "abreVentanaCentrada('dn_desasociar.asp?asociar=fich_residuo&id=" &arrayDatos(0,contadorFilas)& "', 800, 500); return false;" &chr(34)& "	><img src='imagenes/icono_2.gif' alt='Residuos'></a></td>"
						  tablares=tablares & "</tr>"
					'end if
			next
		end if
		rstpag.close
		set rstpag=nothing
		tablares="<table id='resultados' cellspacing='0' cellpadding='3' border='1' align='center' class='sortable'><tr><th><input type='checkbox' name='selector' onchange='cambiachecks(this)' /></th><th>Nº</th><th>TEMA</th><th>TITULO (pulse sobre él para ver/modificar)</th><th>NUM ALTERNATIVA</th><th>Sustancias</th><th>Grupos</th><th>Sectores</th><th>Procesos</th><th>Usos</th><th>Residuos</th></tr>" &tablares& "</table>"
	end if

end if 'busc<>""

' ** AUDITORIA **
spl_accion = "buscar"
spl_entidad = "fichero"
spl_descripcion = condicion

' ** AUDITORIA **
call auditaYCierraConexion(spl_accion,spl_entidad,spl_descripcion) ' accion, entidad, descripcion
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Istas</title>
<link rel="stylesheet" type="text/css" href="dn_estilos.css">
<link rel="stylesheet" type="text/css" href="dn_estilosmenu.css">
<script type="text/javascript" src="dn_scripts.js"></script>
<script type="text/javascript" src="niftycube.js"></script>
<script type="text/javascript">
<!--
window.onload=function(){
Nifty("ul#split h3","top");
Nifty("ul#split div","bottom same-height");
}



function checkbox_checker()
{
	var frm = document.forms["myform"];
	if ( !(frm.NOMBRE.checked || frm.CAS.checked || frm.CE.checked || frm.RD.checked || frm.ONU.checked ) && frm.cualquier.value!='')
	{
		return false;
	}
	else
	{
		return (true);
	}
}

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

function eliminar()
{
	//abreVentanaCentrada('dn_fichero_eliminar.asp', 400, 400):
	var frm = document.forms["myform"];
	frm.action = 'dn_ficheros_eliminar.asp';
	frm.submit();
}


function cambiachecks(c)
{
	var frm = document.forms["myform"];
	for (i=0; i<frm.elements.length; i++)
	{
		if(frm.elements[i].name=='idcheck')
		{
			frm.elements[i].checked=c.checked;
		}
	}
}

function UncheckAll()
{
	var frm = document.forms["myform"];
	for (i=0; i<frm.elements.length; i++)
	{
		if(frm.elements[i].name=='idcheck')
		{
			frm.elements[i].checked=false;
		}
	}
}
-->
</script>
<script src="sorttable.js"></script>
<script language="JavaScript" src="gen_validatorv2.js" type="text/javascript"></script>
</head>

<body>
<!--#include file="dn_menu.asp"-->

<h1>Alternativas / Ficheros </h1>
<%flashMsgShow()%>
<div align="right"><input type="button" onclick="abreVentanaCentrada('dn_fichero.asp', 800, 780)" value='Añadir fichero' /></div>

<form action="dn_ficheros.asp"  method="post" name="myform" onSubmit="primerapag();">
 <input type="hidden" name='busc' value='<%=busc%>' />
 <input type="hidden" name='pag' value='<%=pag%>' />
 <input type="hidden" name='hr' value='<%=hr%>' />
 <input type="hidden" name='arr' value='<%=arr%>' />
<table class="pq" width="100%" cellpadding="3" cellspacing="1" align='center' style='width:40%; border:1px solid #999;'>
<tr>
  <td colspan='2' bgcolor="#C8C866" ><b style='font-size:13px;'>Buscar</b></td>
</tr>

<tr bgcolor='#F7F6F6'>
  <td align="right"><strong>T&iacute;tulo:</strong>&nbsp;    </td>
  <td nowrap='nowrap'>
    <input type='text' size='24' name='titulo' value='<%=titulo%>'></td>
</tr>
<tr bgcolor='#F7F6F6'>
  <td align="right"><strong>Tema:</strong>&nbsp; </td>
  <td nowrap='nowrap'>
    <input type='text' size='24' name='tema' value='<%=tema%>'></td>
</tr>
<tr bgcolor='#F7F6F6'>
  <td align="right"><strong>Número de alternativa:</strong>&nbsp; </td>
  <td nowrap='nowrap'>
    <input type='text' size='24' name='num_alternativa' value='<%=num_alternativa%>'></td>
</tr>
<tr bgcolor='#F7F6F6'>
  <td align="right"><strong>Mostrar: </strong></td>
  <td nowrap='nowrap'>  <span class="negro">
    <input type="text" name="nregs" size=3 maxlength=3 value="<%=nregs%>" class="campo">
	<input type="hidden" name="ordenacion"  value="<%=ordenacion%>">
	<input type="hidden" name="sentido"  value="<%=sentido%>">

ficheros por p&aacute;gina </span></td>
</tr>
<tr bgcolor='#F7F6F6'><td colspan='2' align='center'><input type="submit" value="Buscar" /> </td></tr>
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
		<div align='center' class='margengr'><%paginacion%></div>
		<fieldset class="margengr"><legend><strong>Acciones</strong></legend>
		La acción sobre la que pulse se llevará a cabo sobre los ficheros marcados <input type="checkbox" name="ejemplo" checked />
		.
		<div align="center" class='margengr'><input type="button" onClick="eliminar();" value="ELIMINAR" />

	    </div>
		</fieldset>
<%
	end if
end if
%>

</form>

</body>
</html>

<%
sub paginacion
%>
 Páginas:
<%
	totalpags=roundsup(hr/nregs)
	if pag>1 then
%>
	<a href='#' onclick='cambiapag(<%=pag-1%>)'>&lt; Anterior</a>
<%
	end if

	for i=1 to totalpags
		if (cint(i)=cint(pag)) then
			mipag=" <b>" &i& "</b>"
		else
			mipag=" <a href='#' onclick='cambiapag(" &i& ")'>" &i& "</a>"
		end if
		response.write mipag
	next

	if cint(pag)<cint(totalpags) then
%>
	<a href='#' onclick='cambiapag(<%=pag+1%>)'>Siguiente &gt;</a>
<%
	end if

end sub
%>
