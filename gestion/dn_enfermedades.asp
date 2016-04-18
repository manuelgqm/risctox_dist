<!--#include file="dn_fun_comunes.asp"-->
<!--#include file="dn_auten.inc"-->
<!--#include file="adovbs.inc"-->
<!--#include file="dn_conexion.asp"-->

<%'Permite auditar las acciones%>
<!--#include file="spl_auditoria.inc.asp"-->

<%
' ** AUDITORIA **
spl_accion = "buscar"
spl_entidad = "enfermedad"
spl_descripcion = ""

'PRIMERO: si nos mandan nombre, lo añadimos
if EliminaInyeccionSQL(request.Form("nombre"))<>"" then
	nombre=	EliminaInyeccionSQL(request.Form("nombre"))
	listado=	EliminaInyeccionSQL(request.Form("listado"))
	sintomas=	EliminaInyeccionSQL(request.Form("sintomas"))
	actividades=	EliminaInyeccionSQL(request.Form("actividades"))
	sqls="insert into  dn_risc_enfermedades (nombre, listado, sintomas, actividades) VALUES ('" &nombre& "', '" &listado& "', '" &sintomas& "', '" &actividades& "')"
	objConn1.execute(sqls)
	flashMsgCreate "Se ha añadido el nuevo registro. Puede asociar las sustancias correspondientes desde la <a href='dn_sustancias.asp'>página de sustancias</a>", "OK"

	spl_accion = "crear"
	spl_entidad = "enfermedad"
	spl_descripcion = sqls
end if

		sqls="select id, nombre, nombre_ing from dn_risc_enfermedades ORDER BY nombre"
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
						  tablares=tablares & "<td><a href='#' onclick=" &chr(34)& "abreVentanaCentrada('dn_asociaciones_mostrar.asp?tipo=enfermedad&id=" &arrayDatos(0,contadorFilas)& "', 800, 580); return false;" &chr(34)& "	><img src='imagenes/relacion.png' align='left' alt='Ver relación' /></a>&nbsp;" &arrayDatos(1,contadorFilas)& "</td>"
						  tablares=tablares & "<td>&nbsp;" &arrayDatos(2,contadorFilas)& "</td>"
						  tablares=tablares & "<td align='center'><a href='#' onclick=" &chr(34)& "abreVentanaCentrada('dn_mod.asp?asociar=enfermedad&id=" &arrayDatos(0,contadorFilas)& "', 800, 500); return false;" &chr(34)& "	><img src='imagenes/icono_modificar.gif' alt='Cambiar nombre'></a></td>"
						  tablares=tablares & "<td align='center'><a href='#' onclick=" &chr(34)& "abreVentanaCentrada('dn_desasociar.asp?asociar=enfermedad&id=" &arrayDatos(0,contadorFilas)& "', 800, 580); return false;" &chr(34)& "	><img src='imagenes/icono_estructura.gif' alt='Sustancias asociadas' /></a></td>"
						  tablares=tablares & "<td align='center'><a href='#' onclick=" &chr(34)& "abreVentanaCentrada('dn_desasociar.asp?asociar=enfermedad_gr&id=" &arrayDatos(0,contadorFilas)& "', 800, 580); return false;" &chr(34)& "	><img src='imagenes/icono_grupo.gif' alt='Grupos asociados' /></a></td>"
						  tablares=tablares & "</tr>"
					'end if
				next

				tablares="<table id='resultados' cellspacing='0' cellpadding='3' border='1' align='center'><tr><th><input type='checkbox' name='selector' onchange='cambiachecks(this)' /></th><th>Nº</th><th>NOMBRE</th><th>NOMBRE EN INGLÉS</th><th>Modificar</th><th>Ver/Modificar<br />Sustancias asociadas</th><th>Ver/Modificar<br />Grupos asociados</th></tr>" &tablares& "</table>"
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
<script language="JavaScript" src="gen_validatorv2.js" type="text/javascript"></script>
</head>

<body>
<!--#include file="dn_menu.asp"-->

<h1>Risctox / Enfermedades </h1>
<%flashMsgShow()%>

<form action="dn_enfermedades.asp" method="post" name="myform">

<%
if hr=0 then
		response.Write("No se encontraron registros.")
else
		response.Write("<p class='neg' style='margin:15px 0; padding:10px;'>Se han encontrado " &hr& " registros. </p>")
%>
		<%=tablares%>
		<div align="center" class='margengr'><input type="button" onClick="eliminar('enfermedades');" value="ELIMINAR MARCADOS" />
	    </div>
<%
end if
%>
</form>

<form action="dn_enfermedades.asp" method="post" name="myform2">
		<fieldset class="margengr"><legend><strong>Nueva enfermedad</strong></legend>
		<table>
		<tr><td>Nombre </td><td><input type="text" name="nombre" maxlength="500" size="100"  /></td></tr>
		<tr><td>Listado </td><td><textarea name="listado" cols="70"></textarea></td></tr>
		<tr><td>Síntomas </td><td><textarea name="sintomas" cols="70" id="sintomas"></textarea></td></tr>
		<tr><td>Actividades </td><td><textarea name="actividades" cols="70" id="actividades"></textarea></td></tr>
		<tr><td colspan="2" align="center"> <input type="submit"  value="Añadir" /></td></tr>
		</table>

		</fieldset>
</form>


<script language="JavaScript" type="text/javascript">
var frmvalidator = new Validator("myform2");
frmvalidator.addValidation("nombre","req","Por favor, introduzca el nombre");
frmvalidator.addValidation("listado","maxlen=300");
frmvalidator.addValidation("sintomas","maxlen=1000");
frmvalidator.addValidation("actividades","maxlen=2000");
</script>

</body>
</html>
