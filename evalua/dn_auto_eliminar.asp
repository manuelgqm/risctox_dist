<!--#include file="../dn_conexion.asp"-->
<!--#include file="../adovbs.inc"-->
<!--#include file="../dn_restringida.asp"-->
<!--#include file="../dn_funciones_comunes.asp"-->

<%
' #################################################
' ### ELIMINACIÓN DE PRODUCTOS
' #################################################

' Cogemos los IDs de los productos a eliminar
ids_eliminar = EliminaInyeccionSQL(request.form("idcheck"))

' Eliminamos los componentes de los productos. Adjuntamos id de usuario para mayor seguridad
sql = "DELETE FROM dn_auto_componentes WHERE id_producto IN ("&ids_eliminar&") AND id_ecogente="&session("id_ecogente")
objConnection2.execute(sql),,adexecutenorecords

' Eliminamos los productos. Adjuntamos id de usuario para mayor seguridad
sql = "DELETE FROM dn_auto_productos WHERE id IN ("&ids_eliminar&") AND id_ecogente="&session("id_ecogente")
'response.write sql
objConnection2.execute(sql),,adexecutenorecords

' Cerramos la conexión
cerrarconexion

' Redirigimos a herramienta
response.redirect "dn_auto_herramienta.asp?desp=0"
%>
