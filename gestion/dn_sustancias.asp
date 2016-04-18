<!--#include file="dn_fun_comunes.asp"-->
<!--#include file="dn_fun_texto.asp"-->
<!--#include file="dn_auten.inc"-->
<!--#include file="adovbs.inc"-->
<!--#include file="dn_conexion.asp"-->
<%'Permite auditar las acciones%>
<!--#include file="spl_auditoria.inc.asp"-->

<%
'si busc está vacio, mostramos formulario; si es 1, han dado a "buscar"; si es dos, han dado a paginación 
busc=EliminaInyeccionSQL(request.form("busc"))

if busc="" then
	'valores de busqueda por defecto
	NOMBRE=1
	'CAS=1
	ordenacion=""
	sentido=""
	nregs=20
else
	ordenacion=request("ordenacion")
	sentido=request("sentido") 
	nregs=request("nregs") 
	if isnumeric(nregs) then
		nregs=round(nregs,0)
	else
		nregs=20
	end if
	
	if request.form("cualquier")<>"" then cualquier=Trim(h(request("cualquier")))
	'campos en que tenemos que buscar (en busc=2 no los usamos para hacer la busqueda, pero si para rellenar el formulario con lo que se ha buscado)
	NOMBRE=request.form("NOMBRE") 
	CAS=request.form("CAS")
	CE=request.form("CE") 
	RD=request.form("RD") 
	ONU=request.form("ONU")
	tipobus=request("tipobus")
		
	select case busc
	
	case 1: 'han dado a buscar
					
		condicion=""	
		
		'si mandan texto, creamos condicion
		if cualquier<>"" then			
			if CAS then condicion=condicion& " OR  num_cas = '" &cualquier& "' "
			if CE then condicion=condicion& " OR  num_ce_einecs = '" &cualquier& "' OR num_ce_elincs  = '" &cualquier& "'"
			if RD then condicion=condicion& " OR  num_rd = '" &cualquier& "'"
			if ONU then condicion=condicion& " OR  num_onu = '" &cualquier& "'"	
						
			'si tenemos que buscar en nombre, hacemos busqueda SEGUN TIPOBUS
			if NOMBRE then		
			
				condicion = condicion & " OR  ("			
				
				if tipobus="exacta" then
					condicion = condicion & " (dn_risc_sustancias.nombre= '" &cualquier& "' or dn_risc_sinonimos.nombre= '" &cualquier& "') " 
				else
					' Quitamos espacios dobles
					cualquier=replace(cualquier,"  "," ")		
					' Dividimos la cadena en palabras separadas por espacios
					arraycualquier = split(cualquier, " ")					
					' Montamos la condicion de texto (reemplazando caracteres)	
					FOR i=0 to UBound(arraycualquier)		
						mitermino=arraycualquier(i)		
						mitermino=quitartildes(mitermino)
						mitermino=montartildes(mitermino)
						condicion = condicion & " (dn_risc_sustancias.nombre LIKE '%" &mitermino& "%' or dn_risc_sinonimos.nombre LIKE '%" &mitermino& "%')  " &tipobus
					NEXT		
					'quitamos el último or/and				
					condicion=left(condicion,len(condicion)-3)
				end if

				condicion = condicion & ")"						
				
			end if	'nombre
		end if	'if request.form("cualquier")<>""		
		
		'quitamos el primer or
		if condicion<>"" then condicion= " WHERE " &right(condicion,len(condicion)-3)
		
		miordenacion=replace(ordenacion, ordenacion, "dn_risc_sustancias." &ordenacion) 

		if miordenacion<>"" then miordenacion=miordenacion& " " &sentido&", dn_risc_sustancias.id"

		sqls="select distinct dn_risc_sustancias.id"
		if ordenacion<>"" then sqls=sqls & " , dn_risc_sustancias." &ordenacion
		sqls=sqls & " from dn_risc_sustancias "
		sqls=sqls & "  FULL OUTER JOIN dn_risc_sinonimos ON (dn_risc_sustancias.id=dn_risc_sinonimos.id_sustancia) " 
		sqls=sqls  &condicion& " ORDER BY " &miordenacion

'		response.write sqls

			Set objRst = Server.CreateObject("ADODB.Recordset")
			objRst.Open sqls, objConn1, adOpenStatic, adCmdText
			hr=objRst.recordcount
		
			IF not objRst.eof THEN 
				'arr=objRst.GetString(adClipString, -1, "", ",", "")				
				do while not objRst.eof
					arr=arr& objRst("id")& ","
			  objRst.movenext
				loop
				'esta sera la pagina 1
				pag = 1	
			END IF
					
			objRst.Close
			Set objRst=Nothing
			
	case 2: 'paginando
		
		hr=request("hr")
		pag=request("pag")
		arr=request("arr")
					
	end select 'cual busc
	
	'RESULTADOS DE BUSQUEDA (para busc 1 y busc 2)
	'seleccionamos datos a mostrar de los x registros que toquen
	if hr>0 then
				
		arrayx = split(arr, ",")
		
		FOR i=0 to UBound(arrayx)-1		 'OJO, luego hay que hacer que muestre solo las de paginacion
			if len(trim(arrayx(i)))>0 then cadenaids=cadenaids  &arrayx(i)&","		
		NEXT	
		
		'quitamos la ultima coma
		cadenaids= left(cadenaids,len(cadenaids)-1)
		sqlpag="select id, num_cas, nombre from dn_risc_sustancias WHERE id IN(" &cadenaids& ") ORDER BY " &ordenacion&  " " &sentido
'		response.write sqlpag
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
		'	response.write "<p>registrofin-1=" &registrofin& "</p>"
			
			
			for contadorFilas=registroini to registrofin				
				
					'if contadorfilas>=hr-1 then
							'exit for
					'else
						  'arrayDatos(0,contadorFilas)
						  tablares=tablares & "<tr>" 
						  tablares=tablares & "<td><input type='checkbox' name='idcheck' value=" &arrayDatos(0,contadorFilas)& " /></td>"
						  tablares=tablares & "<td>" & contadorFilas+1 & "</td>"
				  		  tablares=tablares & "<td>" & arrayDatos(1,contadorFilas)  & "</td>"	
						  tablares=tablares & "<td><a href='#' onclick=" &chr(34)& "abreVentanaCentrada('dn_asociaciones_mostrar.asp?tipo=sustancia&id=" &arrayDatos(0,contadorFilas)& "', 800, 580); return false;" &chr(34)& "	><img src='imagenes/relacion.png' align='left' alt='Ver relación' /></a>&nbsp;<a href='#' onclick=" &chr(34)& "abreVentanaCentrada('dn_sustancia.asp?id=" &arrayDatos(0,contadorFilas)& "', 800, 580); return false;" &chr(34)& "	>" &arrayDatos(2,contadorFilas)& "</a><br />" & dameSinonimos(arrayDatos(0,contadorFilas))&"</td>"
						  tablares=tablares & "<td align='center'><a href='#' onclick=" &chr(34)& "abreVentanaCentrada('dn_sustanciaAD.asp?id=" &arrayDatos(0,contadorFilas)& "', 1000, 700); return false;" &chr(34)& "	><img src='imagenes/icono_modificar.gif' alt='Campos adicionales'></a></td>"							
						  tablares=tablares & "</tr>" 
					'end if				
			next
		end if
		rstpag.close
		set rstpag=nothing
		tablares="<table id='resultados' cellspacing='0' cellpadding='3' border='1' align='center'><tr><th><input type='checkbox' name='selector' onchange='cambiachecks(this)' /></th><th>Nº</th><th>CAS</th><th>SUSTANCIA (pulse sobre el nombre para ver/modificar)</th><th>Campos<br/>adicionales</th></tr>" &tablares& "</table>" 
	end if

end if 'busc<>""

' ** AUDITORIA **
call auditaYCierraConexion("buscar","sustancia",condicion) ' accion, entidad, descripcion
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

function DoCustomValidation()
{
  var frm = document.forms["myform"]; 
  if(false == checkbox_checker())
  {
    alert("Por favor, seleccione al menos un campo en el que buscar.");
    return false;
  }
  else
  {
    return true;
  }
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
	//abreVentanaCentrada('dn_sustancia_eliminar.asp', 400, 400):
	var frm = document.forms["myform"]; 	
	frm.action = 'dn_sustancias_eliminar.asp';
	frm.submit();
}

function asociar(tabla)
{
	abreVentanaCentrada('dn_asociar.asp?asociar='+tabla, 1000, 250)
	var frm = document.forms["myform"].target="nueva"; 
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
<script language="JavaScript" src="gen_validatorv2.js" type="text/javascript"></script>
</head>

<body>
<!--#include file="dn_menu.asp"-->

<h1>Risctox / Sustancias </h1>
<%flashMsgShow()%>
<div align="right"><input type="button" onclick="abreVentanaCentrada('dn_sustancia.asp', 800, 780)" value='Añadir sustancia' /></div>

<form action="dn_sustancias.asp" method="post" name="myform" >
 <input type="hidden" name='busc' value='<%=busc%>' />	
 <input type="hidden" name='pag' value='<%=pag%>' />	
 <input type="hidden" name='hr' value='<%=hr%>' />		
 <input type="hidden" name='arr' value='<%=arr%>' />		
<table class="pq" width="100%" cellpadding="3" cellspacing="1" align='center' style='width:40%; border:1px solid #999;'>
<tr>
  <td colspan='2' bgcolor="#C8C866" ><b style='font-size:13px;'>Buscar</b></td>
</tr>

<tr bgcolor='#F7F6F6'>
  <td valign="top" nowrap="nowrap"><strong>Texto:</strong>&nbsp;
    <input type='text' size='24' name='cualquier' value='<%=cualquier%>'></td>
  <td nowrap='nowrap'>
  <strong>en</strong>  <input type="checkbox" name="NOMBRE" value="1" <%if NOMBRE=1 then response.write("checked")%> /> Nombre 
  <input type="checkbox" name="CAS" value="1" class="paden" <%if CAS=1 then response.write("checked")%> /> CAS
 <input type="checkbox" name="CE" value="1" class="paden" <%if CE=1 then response.write("checked")%> /> CE
 <input type="checkbox" name="RD" value="1" class="paden" <%if RD=1 then response.write("checked")%> /> RD
<input type="checkbox" name="ONU" value="1" class="paden" <%if ONU=1 then response.write("checked")%> /> ONU</td>
</tr>
<tr bgcolor='#F7F6F6'>
  <td align="right"><strong>Ordenar por: </strong></td>
  <td nowrap='nowrap'>
<select name="ordenacion" class="campo">
<option value="nombre" >nombre</option>
<option value="ID" <%if ordenacion="ID" then response.write ("selected")%>>fecha introducción</option>
<option value="num_cas" <%if ordenacion="num_cas" then response.write ("selected")%>>CAS</option>
<option value="num_ce_einecs,num_ce_elincs" <%if ordenacion="num_ce_einecs,num_ce_elincs" then response.write ("selected")%>>CEE</option>
<option value="num_rd" <%if ordenacion="num_rd" then response.write ("selected")%>>RD</option>
</select>
<select name="sentido" class="campo"><option value="">ascendente</option><option value="DESC" <%if sentido="DESC" then response.write ("selected")%>>descendente</option></select></td></tr>
<tr bgcolor='#F7F6F6'>
  <td align="right"><strong>Mostrar: </strong></td>
  <td nowrap='nowrap'>  <span class="negro">
    <input type="text" name="nregs" size=3 maxlength=3 value="<%=nregs%>" class="campo">
sustancias por p&aacute;gina </span></td>
</tr>
<tr bgcolor='#F7F6F6'>
  <td align="right"><strong>Tipo de b&uacute;squeda: </strong></td>
  <td nowrap='nowrap'>  <span class="negro">
    <select name="tipobus" class="campo">      
      <option value="and" <%if tipobus="and" then response.write ("selected")%>>todas las palabras (AND)</option>
      <option value="or" <%if tipobus="or" then response.write ("selected")%>>alguna de las palabras (OR)</option>
      <option value="exacta"<%if tipobus="exacta" then response.write ("selected")%>>exacta</option>
     </select> 
    (para Nombre)
</span></td>
</tr>
<tr bgcolor='#F7F6F6'><td colspan='2' align='center'><input type="submit" value="Buscar" onclick="primerapag();"/> </td></tr>
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
		La acción sobre la que pulse se llevará a cabo sobre las sustancias marcadas <input type="checkbox" name="ejemplo" checked />
		. 
		<div align="center" class='margengr'><input type="button" onClick="eliminar();" value="ELIMINAR" />
          <br>
          <br>
          Asociar a:
          <input type="button" onClick="asociar('grupo');" value="Grupo" /> 
		<input type="button" onClick="asociar('uso');" value="Uso" /> 
		<input type="button" onClick="asociar('compania');" value="Compañía" />
    <input type="button" onClick="asociar('sector');" value="Sector" />		
		<input type="button" onClick="asociar('enfermedad');" value="Enfermedad" /> 
		<input type="button" onClick="asociar('fich_sustancia');" value="Fichero" /> 
		<br /> <br />
	    </div>
		</fieldset>
<%
	end if
end if
%>

</form>
<script language="JavaScript" type="text/javascript">
var frmvalidator = new Validator("myform");
frmvalidator.addValidation("cualquier","req","Por favor, introduzca el texto a buscar");
//frmvalidator.addValidation("cualquier","minlen=3","La longitud mínima del texto a buscar es de tres caracteres"); 
frmvalidator.setAddnlValidationFunction("DoCustomValidation");
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

