<!--#include file="dn_fun_comunes.asp"-->
<!--#include file="dn_fun_texto.asp"-->
<!--#include file="dn_auten.inc"-->
<!--#include file="adovbs.inc"-->
<!--#include file="dn_conexion.asp"-->
<%'Permite auditar las acciones%>
<!--#include file="spl_auditoria.inc.asp"-->

<%
' ** AUDITORIA **
spl_accion = "buscar"
spl_entidad = "sector"
spl_descripcion = ""
'PRIMERO: si nos mandan nombre, lo añadimos
if EliminaInyeccionSQL(request.Form("nombre"))<>"" then

	nombre=h(EliminaInyeccionSQL(request.Form("nombre")))
	numero_cnae=h(EliminaInyeccionSQL(request.Form("numero_cnae")))
	
	sqls="insert into  dn_alter_sectores (nombre, numero_cnae) VALUES ('" &nombre& "', '" &numero_cnae& "')"
	objConn1.execute(sqls)
	flashMsgCreate "Se ha añadido el nuevo registro.", "OK"

	spl_accion = "crear"
	spl_entidad = "sector"
	spl_descripcion = sqls		
end if

	'mostramos todos los resultados (paginados)
	
	
	'aqui no hay busqueda, siempre se muestran resultados, indicamos valores por defecto si no los hay
	ordenacion=EliminaInyeccionSQL(request("ordenacion"))
	if ordenacion="" then ordenacion="nombre"
	sentido=EliminaInyeccionSQL(request("sentido"))
	'if sentido="" then sentido="" 
	nregs=EliminaInyeccionSQL(request("nregs")) 
	if nregs="" then nregs=100
	busc=EliminaInyeccionSQL(request.form("busc"))
	if busc="" then busc=1 

	if isnumeric(nregs) then
		nregs=round(nregs,0)
	else
		nregs=100
	end if
	
		
	select case busc
	
	case 1: 'han dado a buscar
		
		sqls="select id from dn_alter_sectores order by " &ordenacion& " " &sentido 
		'response.write sqls
		
			Set objRst = Server.CreateObject("ADODB.Recordset")
			objRst.Open sqls, objConn1, adOpenStatic, adCmdText
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
		sqlpag="select id, numero_cnae, nombre from dn_alter_sectores ORDER BY " &ordenacion&  " " &sentido
		set rstpag=objConn1.execute(sqlpag)
		if not rstpag.eof then
			
			arrayDatos = rstpag.GetRows			
      		
			registroini=(pag*nregs)-nregs
			registrofin=registroini+nregs
			
			if registrofin>=hr-1 then
				registrofin=hr
			end if
			
			registrofin=registrofin-1

			if hr>0 then
				for contadorFilas=registroini to registrofin					
					
						  tablares=tablares & "<tr>" 
						  tablares=tablares & "<td><input type='checkbox' name='idcheck' value=" &arrayDatos(0,contadorFilas)& " /></td>"
						  tablares=tablares & "<td>" & contadorFilas+1 & "</td>"				  		  
						  tablares=tablares & "<td>" &arrayDatos(1,contadorFilas)& "</td>"		
						  tablares=tablares & "<td><a href='#' onclick=" &chr(34)& "abreVentanaCentrada('dn_asociaciones_mostrar.asp?tipo=sector&id=" &arrayDatos(0,contadorFilas)& "', 800, 580); return false;" &chr(34)& "	><img src='imagenes/relacion.png' align='left' alt='Ver relación' /></a>&nbsp;" &arrayDatos(2,contadorFilas)& "</td>"		
						  
						  tablares=tablares & "<td align='center'><a href='#' onclick=" &chr(34)& "abreVentanaCentrada('dn_mod.asp?asociar=sector&id=" &arrayDatos(0,contadorFilas)& "', 800, 300); return false;" &chr(34)& "	><img src='imagenes/icono_modificar.gif' alt='Modificar'></a></td>"
						

              tablares=tablares & "<td align='center'><a href='#' onclick=" &chr(34)& "abreVentanaCentrada('dn_desasociar.asp?asociar=sector&id=" &arrayDatos(0,contadorFilas)& "', 800, 580); return false;" &chr(34)& "	><img src='imagenes/icono_estructura.gif' alt='Sustancias asociadas' /></a></td>"
						 				
						  tablares=tablares & "</tr>" 
					'end if				
				next		
		
				tablares="<table id='resultados' cellspacing='0' cellpadding='3' border='1' align='center'><tr><th><input type='checkbox' name='selector' onchange='cambiachecks(this)' /></th><th>Nº</th><th nowrap>Nº CNAE <a href='dn_sectores.asp?ordenacion=numero_cnae&sentido=ASC' style='text-decoration:none'>&uarr;</a> <a href='dn_sectores.asp?ordenacion=numero_cnae&sentido=DESC' style='text-decoration:none'>&darr;</a></th><th>NOMBRE <a href='dn_sectores.asp?ordenacion=nombre&sentido=ASC' style='text-decoration:none'>&uarr;</a> <a href='dn_sectores.asp?ordenacion=nombre&sentido=DESC' style='text-decoration:none'>&darr;</a></th><th>Modificar</th><th>Ver/Modificar Sustancias asociadas</th></tr>" &tablares& "</table>" 
			end if
			
		end if
		rstpag.close
		set rstpag=nothing
		
	end if
	
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
<script type="text/javascript">
<!--
window.onload=function(){
Nifty("ul#split h3","top");
Nifty("ul#split div","bottom same-height");
}

function asociar(tabla)
{
	abreVentanaCentrada('dn_asociar.asp?asociar='+tabla, 1000, 250)
	var frm = document.forms["myform"].target="nueva"; 
}

function eliminar(tabla)
{
	var frm = document.forms["myform"]; 	
	frm.action = 'dn_eliminar.asp?tabla='+tabla;
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



-->
</script>
<script language="JavaScript" src="sorttable.js" type="text/javascript"></script>
<script language="JavaScript" src="gen_validatorv2.js" type="text/javascript"></script>
</head>

<body>
<!--#include file="dn_menu.asp"-->

<h1>Alternativas / Sectores </h1>
<%flashMsgShow()%>

<form action="dn_sectores.asp" method="post" name="myform">
  <input type="hidden" name='busc' value='<%=busc%>' />	
 <input type="hidden" name='pag' value='<%=pag%>' />	
 <input type="hidden" name='hr' value='<%=hr%>' />		
 <input type="hidden" name='arr' value='<%=arr%>' />
 <input type="hidden" name='nregs' value='<%=nregs%>' />				
 <input type="hidden" name='ordenacion' value='<%=ordenacion%>' />				
 <input type="hidden" name='sentido' value='<%=sentido%>' />		
<%
if hr=0 then
		response.Write("No se encontraron registros.")
else
		response.Write("<p class='neg' style='margin:15px 0; padding:10px;'>Se han encontrado " &hr& " registros. </p>")
%>		
		<%=tablares%>
		
		<div align='center' style="margin:20px 10px;   padding:3px;"><%paginacion%></div>
		
				<fieldset class="margengr"><legend><strong>Acciones</strong></legend>
		La acción sobre la que pulse se llevará a cabo sobre los sectores marcados 
		<input type="checkbox" name="ejemplo" checked />
		. 
		<div align="center" class='margengr'>
		<input type="button" onClick="eliminar('sectores');" value="ELIMINAR" />  
		<p>Asociar a: <input type="button" onClick="asociar('fich_sector');" value="Ficheros" /> 
		</p>
	    </div>
		</fieldset>
		
<%
end if
%>
</form>

<form action="dn_sectores.asp" method="post" name="myform2">
		<fieldset class="margengr centcontenido"><legend><strong>Nuevo sector</strong></legend>
		<table align="center">
		<tr><td>Nombre</td><td> <input type="text" name="nombre" maxlength="750" size="100" /> </td></tr>
		<tr><td>Nº CNAE</td><td align="left"> <input type="text" name="numero_cnae" maxlength="10" size="10" /> </td></tr>
		</table>
		<br /><input type="submit"  value="Añadir" />  
		</fieldset>
</form>


<script language="JavaScript" type="text/javascript">
var frmvalidator = new Validator("myform2");
frmvalidator.addValidation("nombre","req","Por favor, introduzca el nombre");
</script>

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

