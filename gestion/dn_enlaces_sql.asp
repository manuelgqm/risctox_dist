<!--#include file="dn_fun_comunes.asp"-->
<!--#include file="dn_fun_texto.asp"-->
<!--#include file="dn_auten.inc"-->
<!--#include file="adovbs.inc"-->
<!--#include file="dn_conexion.asp"-->
<%'Permite auditar las acciones%>
<!--#include file="spl_auditoria.inc.asp"-->

<%
' ** AUDITORIA **
spl_entidad = "enlace"

'Recojo los datos
id = request("id")
titulo = request("titulo")
enlace = request("enlace")
texto = request("texto")
clasificacion = request("clasificacion")
hidAccion=request("hidAccion")

if hidAccion = 3 then 'Eliminación
	ids=request.form("idcheck")
	sqld="delete from dn_alter_enlaces where id IN (" &ids& ")"
	objconn1.execute(sqld)

	' ** AUDITORIA **
	spl_accion = "eliminar"
	spl_descripcion = ids
	call auditaYCierraConexion(spl_accion,spl_entidad,spl_descripcion) ' accion, entidad, descripcion			

	response.redirect("./dn_enlaces.asp")

end if

if (id = "") then
	sql = "insert into dn_alter_enlaces(titulo, enlace, texto, clasificacion)"
	sql = sql & " values('"&titulo&"','"&enlace&"','"&texto&"','"&clasificacion&"')"
	objconn1.execute(sql)

	' ** AUDITORIA **
	spl_accion = "crear"
	spl_descripcion = sql
	call auditaYCierraConexion(spl_accion,spl_entidad,spl_descripcion) ' accion, entidad, descripcion			

	response.redirect("./dn_enlaces.asp")	
	
else
	sql = "update dn_alter_enlaces set titulo='"&titulo&"', enlace='"&enlace&"', texto='"&texto&"',clasificacion='"&clasificacion&"' where id="&id
	objconn1.execute(sql)
	spl_accion = "modificar"
	spl_descripcion = sql
	call auditaYCierraConexion(spl_accion,spl_entidad,spl_descripcion) ' accion, entidad, descripcion			
end if



%>
<html>
<head></head>
<body>
	<script language='javascript'>
		opener.document.location.reload();
		window.close();
	</script>
</body>
</html>