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
spl_entidad = "enlace"
spl_descripcion = ""

'PRIMERO: si nos mandan nombre, lo aÃ±adimos
if EliminaInyeccionSQL(request.Form("nombre"))<>"" then

	titulo=h(EliminaInyeccionSQL(request.Form("titulo")))
	enlace=h(EliminaInyeccionSQL(request.Form("enlace")))
	
	sqls="insert into  dn_alter_enlaces (titulo, enlace, texto, clasificacion) VALUES ('" &titulo& "', '" &enlace& "','"&texto&"','"&clasificacion&"')"
	objConn1.execute(sqls)
	flashMsgCreate "Se ha aÃ±adido el nuevo registro.", "OK"

' ** AUDITORIA **
	spl_accion = "crear"
	spl_entidad = "enlace"
	spl_descripcion = sqls				
end if

		sqls="select * from dn_alter_enlaces ORDER BY titulo"
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
						  tablares=tablares & "<td>&nbsp;" &arrayDatos(1,contadorFilas)& "</td>"		
						  tablares=tablares & "<td>" & corta (arrayDatos(2,contadorFilas), 40, "puntossuspensivos")& "</td>"
						  tablares=tablares & "<td>" & corta (arrayDatos(3,contadorFilas), 40, "puntossuspensivos")& "</td>"		
						  select case arrayDatos(4,contadorFilas)
						  		case 1:
									clasificacion="Fuentes de información generales"
								case 2:
									clasificacion="Información sobre eliminación/sustitución"
						  end select
						  tablares=tablares & "<td>" & clasificacion& "</td>"
						  
						  tablares=tablares & "<td align='center'><a href='#' onclick=" &chr(34)& "abreVentanaCentrada('dn_enlaces_mod.asp?id=" &arrayDatos(0,contadorFilas)& "', 800, 300); return false;" &chr(34)& "	><img src='imagenes/icono_modificar.gif' alt='Modificar'></a></td>"
						 				
						  tablares=tablares & "</tr>" 
					'end if				
				next		
		
				tablares="<table id='resultados' cellspacing='0' cellpadding='3' border='1' align='center'><tr><th><input type='checkbox' name='selector' onchange='cambiachecks(this)' /></th><th>Título</th><th>Enlace</th><th>Texto</th><th>Clasificación</th><th> </th></tr>" &tablares& "</table>" 
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
	frm.action = 'dn_enlaces_sql.asp?hidaccion=3';
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

<h1>Alternativas / Enlaces </h1>
<%flashMsgShow()%>

<form action="dn_enlaces_sql.asp" method="post" name="myform">
 
<%
if hr=0 then
		response.Write("No se encontraron registros.")
else
		response.Write("<p class='neg' style='margin:15px 0; padding:10px;'>Se han encontrado " &hr& " registros. </p>")
%>		
		<%=tablares%>
		
		
		
		<fieldset class="margengr"><legend><strong>Acciones</strong></legend>
		La acción sobre la que pulse se llevará a cabo sobre los enlaces marcados 
		<input type="checkbox" name="ejemplo" checked />
		. 
		<div align="center" class='margengr'>
		<input type="button" onClick="eliminar('enlaces');" value="ELIMINAR" />  
		 
		</p>
	    </div>
		</fieldset>
		
<%
end if
%>
</form>

<form action="dn_enlaces_sql.asp" method="post" name="myform2">

		<fieldset class="margengr"><legend><strong>Nuevo enlace</strong></legend>
		Título<br /><input type="text" name="titulo" maxlength="250" size="80" /><br/><br/>
    	Enlace<br /><input type="text" name="enlace" maxlength="250" size="80" /><br/><br/>
        Texto<br /><textarea name="texto" rows="6" cols="80"></textarea><br />
        Clasificación<br />
        	<select name='clasificacion'>
            	<option value='1'>Fuentes de información generales</option>
                <option value='2'>Información sobre eliminación/sustitución</option>
            </select>
        <br />
		<br /><input type="submit"  value="Añadir" />  
		</fieldset>

</form>

<!--
<script language="JavaScript" type="text/javascript">
var frmvalidator = new Validator("myform2");
frmvalidator.addValidation("nombre","req","Por favor, introduzca el nombre");
</script>
-->

</body>
</html>
