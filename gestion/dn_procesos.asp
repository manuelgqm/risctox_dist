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
spl_entidad = "proceso"
spl_descripcion = ""

'PRIMERO: si nos mandan nombre, lo añadimos
if EliminaInyeccionSQL(request.Form("nombre"))<>"" then

	nombre=h(EliminaInyeccionSQL(request.Form("nombre")))
	descripcion=h(EliminaInyeccionSQL(request.Form("descripcion")))
	
	sqls="insert into  dn_alter_procesos (nombre, descripcion) VALUES ('" &nombre& "', '" &descripcion& "')"
	objConn1.execute(sqls)
	flashMsgCreate "Se ha añadido el nuevo registro.", "OK"

	spl_accion = "crear"
	spl_entidad = "proceso"
	spl_descripcion = sqls		
end if

		sqls="select id, nombre, descripcion from dn_alter_procesos ORDER BY nombre"
		'response.write sqls

			Set objRst = Server.CreateObject("ADODB.Recordset")
			objRst.Open sqls, objConn1, adOpenStatic, adCmdText
					
			hr=objRst.recordcount
			IF not objRst.eof THEN 
				arrayDatos=objRst.GetRows	
			END IF
					
			objRst.Close
			Set objRst=Nothing
			' ** AUDITORIA **
			call auditaYCierraConexion(spl_accion,spl_entidad,spl_descripcion) ' accion, entidad, descripcion			
			
			if hr>0 then
				for contadorFilas=0 to hr-1			
					
						  tablares=tablares & "<tr>" 
						  tablares=tablares & "<td><input type='checkbox' name='idcheck' value=" &arrayDatos(0,contadorFilas)& " /></td>"
						  tablares=tablares & "<td>" & contadorFilas+1 & "</td>"				  		  
						  tablares=tablares & "<td><a href='#' onclick=" &chr(34)& "abreVentanaCentrada('dn_asociaciones_mostrar.asp?tipo=proceso&id=" &arrayDatos(0,contadorFilas)& "', 800, 580); return false;" &chr(34)& "	><img src='imagenes/relacion.png' align='left' alt='Ver relación' /></a>&nbsp;" &arrayDatos(1,contadorFilas)& "</td>"		
						  tablares=tablares & "<td>" & corta (arrayDatos(2,contadorFilas), 40, "puntossuspensivos")& "</td>"		
						  
						  tablares=tablares & "<td align='center'><a href='#' onclick=" &chr(34)& "abreVentanaCentrada('dn_mod.asp?asociar=proceso&id=" &arrayDatos(0,contadorFilas)& "', 800, 300); return false;" &chr(34)& "	><img src='imagenes/icono_modificar.gif' alt='Modificar'></a></td>"
						 				
						  tablares=tablares & "</tr>" 
					'end if				
				next		
		
				tablares="<table id='resultados' cellspacing='0' cellpadding='3' border='1' align='center'><tr><th><input type='checkbox' name='selector' onchange='cambiachecks(this)' /></th><th>Nº</th><th>NOMBRE</th><th>Descripción</th><th>Modificar</th></tr>" &tablares& "</table>" 
			end if


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

<h1>Alternativas / Procesos </h1>
<%flashMsgShow()%>

<form action="dn_procesoes.asp" method="post" name="myform">
 
<%
if hr=0 then
		response.Write("No se encontraron registros.")
else
		response.Write("<p class='neg' style='margin:15px 0; padding:10px;'>Se han encontrado " &hr& " registros. </p>")
%>		
		<%=tablares%>
		
		
		
						<fieldset class="margengr"><legend><strong>Acciones</strong></legend>
		La acción sobre la que pulse se llevará a cabo sobre los procesos marcados 
		<input type="checkbox" name="ejemplo" checked />
		. 
		<div align="center" class='margengr'>
		<input type="button" onClick="eliminar('procesos');" value="ELIMINAR" />  
		<p>Asociar a: <input type="button" onClick="asociar('fich_proceso');" value="Ficheros" /> 
		</p>
	    </div>
		</fieldset>
		
<%
end if
%>
</form>

<form action="dn_procesos.asp" method="post" name="myform2">

		<fieldset class="margengr"><legend><strong>Nuevo proceso</strong></legend>
		Nombre<br /><input type="text" name="nombre" maxlength="250" size="80" /><br/><br/>
    Descripción<br /><textarea name="descripcion" rows="10" cols="80"></textarea><br />
		<br /><input type="submit"  value="Añadir" />  
		</fieldset>

</form>


<script language="JavaScript" type="text/javascript">
var frmvalidator = new Validator("myform2");
frmvalidator.addValidation("nombre","req","Por favor, introduzca el nombre");
</script>

</body>
</html>
