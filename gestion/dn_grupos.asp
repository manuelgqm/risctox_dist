<!--#include file="dn_fun_comunes.asp"-->
<!--#include file="dn_fun_texto.asp"-->
<!--#include file="dn_auten.inc"-->
<!--#include file="adovbs.inc"-->
<!--#include file="dn_conexion.asp"-->
<%'Permite auditar las acciones%>
<!--#include file="spl_auditoria.inc.asp"-->


<%
' ** AUDITORIA **
accion = "buscar"
entidad = "grupo"
descripcion = ""
'PRIMERO: si nos mandan nombre, lo a�adimos
if EliminaInyeccionSQL(request.Form("nombre"))<>"" then
	nombre=	h(EliminaInyeccionSQL(request.Form("nombre")))
  descripcion = h(EliminaInyeccionSQL(request.form("descripcion")))

	sqls="insert into dn_risc_grupos (nombre, descripcion) VALUES ('" &nombre& "', '"&descripcion&"')"
	objConn1.execute(sqls)
	flashMsgCreate "Se ha a�adido el nuevo registro. Puede asociar las sustancias correspondientes desde la <a href='dn_sustancias.asp'>p�gina de sustancias</a>", "OK"

	accion = "crear"
	entidad = "grupo"
	descripcion = sqls
end if

		sqls="select id, nombre, num_cas, descripcion, nombre_ing, descripcion_ing from dn_risc_grupos ORDER BY num_cas, nombre"
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
			call auditaYCierraConexion(accion,entidad,descripcion) ' accion, entidad, descripcion

			if hr>0 then
				for contadorFilas=0 to hr-1

						  tablares=tablares & "<tr>"
						  tablares=tablares & "<td><input type='checkbox' name='idcheck' value=" &arrayDatos(0,contadorFilas)& " /></td>"
						  tablares=tablares & "<td>" &arrayDatos(2,contadorFilas)& "</td>"
						  tablares=tablares & "<td><a href='#' onclick=" &chr(34)& "abreVentanaCentrada('dn_asociaciones_mostrar.asp?tipo=grupo&id=" &arrayDatos(0,contadorFilas)& "', 800, 580); return false;" &chr(34)& "	><img src='imagenes/relacion.png' align='left' alt='Ver relaci�n' /></a>&nbsp;" &arrayDatos(1,contadorFilas)& "<br/><blockquote><i>"&nl2br(arrayDatos(3,contadorFilas))&"</i></blockquote></td>"
						  tablares=tablares & "<td>" &arrayDatos(4,contadorFilas)& "</td>"
						  tablares=tablares & "<td align='center'><a href='#' onclick=" &chr(34)& "abreVentanaCentrada('dn_mod.asp?asociar=grupo&id=" &arrayDatos(0,contadorFilas)& "', 1000, 375); return false;" &chr(34)& "	><img src='imagenes/icono_modificar.gif' alt='Cambiar nombre'></a></td>"
						  tablares=tablares & "<td align='center'><a href='#' onclick=" &chr(34)& "abreVentanaCentrada('dn_desasociar.asp?asociar=grupo&id=" &arrayDatos(0,contadorFilas)& "', 985, 575); return false;" &chr(34)& "	><img src='imagenes/icono_estructura.gif' alt='Asociar en bloque' /></a></td>"
						  tablares=tablares & "<td align='center'><a href='#' onclick=" &chr(34)& "abreVentanaCentrada('dn_asociarbloque.asp?asociar=grupo&id=" &arrayDatos(0,contadorFilas)& "', 985, 575); return false;" &chr(34)& "	><img src='imagenes/icono_asociarenbloque.gif' alt='Sustancias asociadas' /></a></td>"
						  tablares=tablares & "</tr>"
					'end if
				next

				tablares="<table id='resultados' cellspacing='0' cellpadding='3' border='1' align='center'><tr><th><input type='checkbox' name='selector' onchange='cambiachecks(this)' /></th><th>N�</th><th>NOMBRE</th><th>NOMBRE EN INGL&Eacute;S</th><th>Modificar</th><th>Ver/Modificar<br />Sustancias asociadas</th><th>Asociar<br />en bloque</th></tr>" &tablares& "</table>"
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

function asociar(tabla)
{
	abreVentanaCentrada('dn_asociar.asp?asociar='+tabla, 640, 250)
	var frm = document.forms["myform"].target="nueva";
}
-->
</script>
<script language="JavaScript" src="gen_validatorv2.js" type="text/javascript"></script>
</head>

<body>
<!--#include file="dn_menu.asp"-->

<h1>Risctox / Grupos </h1>
<%flashMsgShow()%>

<form action="dn_grupos.asp" method="post" name="myform">

<%
if hr=0 then
		response.Write("No se encontraron registros.")
else
		response.Write("<p class='neg' style='margin:15px 0; padding:10px;'>Se han encontrado " &hr& " registros. </p>")
%>
		<%=tablares%>
		<fieldset class="margengr"><legend><strong>Acciones</strong></legend>
		La acci�n sobre la que pulse se llevar� a cabo sobre los grupos marcados
		<input type="checkbox" name="ejemplo" checked />
		.
		<div align="center" class='margengr'>
		<input type="button" onClick="eliminar('grupos');" value="ELIMINAR" />
		<p>Asociar a: <input type="button" onClick="asociar('uso_gr');" value="Uso" />
		<input type="button" onClick="asociar('enfermedad_gr');" value="Enfermedad" />
		<input type="button" onClick="asociar('fich_grupo');" value="Fichero" />
		</p>
	    </div>
		</fieldset>

<%
end if
%>
</form>

<form action="dn_grupos.asp" method="post" name="myform2">
		<fieldset class="margengr"><legend><strong>Nuevo grupo</strong></legend>
		Nombre<br /><input type="text" name="nombre" maxlength="750" size="100" /> <br />
    Descripci�n<br /><textarea name="descripcion" rows="10" cols="80"></textarea><br />
		<br /><input type="submit"  value="A�adir" />
		</fieldset>
</form>


<script language="JavaScript" type="text/javascript">
var frmvalidator = new Validator("myform2");
frmvalidator.addValidation("nombre","req","Por favor, introduzca el nombre");
frmvalidator.addValidation("descripcion","maxlen=2500")
</script>

</body>
</html>
