<!--#include file="dn_fun_comunes.asp"-->
<!--#include file="dn_auten.inc"-->
<!--#include file="adovbs.inc"-->
<!--#include file="dn_conexion.asp"-->
<%'Permite auditar las acciones%>
<!--#include file="spl_auditoria.inc.asp"-->


<%
if request.form("idcheck")="" then
	flashMsgCreate "No se indicó ningún fichero para eliminar.", "Advertencia"
else
	ids=request.form("idcheck")

	spl_accion = "eliminar"
	spl_entidad = "fichero"
	spl_descripcion = "dn_alter_ficheros,dn_alter_ficheros_por_sustancias,dn_alter_ficheros_por_grupos,dn_alter_ficheros_por_procesos,dn_alter_ficheros_por_sectores,dn_alter_ficheros_por_usos para ficheros con ids "&ids
	
	'consultamos que archivos hay que borrar
	set rstb=objconn1.execute("select archivo from  dn_alter_ficheros where id IN (" &ids& ")")
	do while not rstb.eof
		archivo=rstb("archivo")
		if archivo<>"" then borrarfichero "\estructuras\" &archivo
	rstb.movenext
	loop
	rstb.close
	set rstb=nothing
	
	objconn1.execute("delete from dn_alter_ficheros where id IN (" &ids& ")")
	
	'borrar tb las asociaciones
	objconn1.execute("delete from dn_alter_ficheros_por_sustancias where id_fichero IN (" &ids& ")")	
	objconn1.execute("delete from dn_alter_ficheros_por_grupos where id_fichero IN (" &ids& ")")	
	objconn1.execute("delete from dn_alter_ficheros_por_procesos where id_fichero IN (" &ids& ")")	
	objconn1.execute("delete from dn_alter_ficheros_por_sectores where id_fichero IN (" &ids& ")")	
	objconn1.execute("delete from dn_alter_ficheros_por_usos where id_fichero IN (" &ids& ")")		
	
	flashMsgCreate "Los ficheros seleccionados han sido eliminados.", "OK"
end if
' ** AUDITORIA **
call auditaYCierraConexion(spl_accion,spl_entidad,spl_descripcion) ' accion, entidad, descripcion			

response.redirect "dn_ficheros.asp"
'flashMsgShow()
%>
