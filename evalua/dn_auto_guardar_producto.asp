<!--#include file="../dn_conexion.asp"-->
<!--#include file="../adovbs.inc"-->
<!--#include file="../dn_restringida.asp"-->
<!--#include file="../dn_funciones_comunes.asp"-->
<!--#include file="../dn_funciones_texto.asp"-->

<%
' ##########################################################
' ### INSERCIÓN DE NUEVO PRODUCTO CON TODOS SUS COMPONENTES
' ### Llega del formulario en dn_auto_herramienta.asp
' ##########################################################

' ##########################
' ### PRODUCTO
' ##########################
' Cogemos datos del producto
prod_nombre = EliminaInyeccionSQL(h(trim(request.form("prod_nombre"))))
prod_cod_proceso = EliminaInyeccionSQL(h(trim(request.form("prod_cod_proceso"))))
prod_frases_r = EliminaInyeccionSQL(h(trim(request.form("prod_frases_r"))))
prod_cod_estado = EliminaInyeccionSQL(h(trim(request.form("prod_cod_estado"))))
prod_cod_presion = EliminaInyeccionSQL(h(trim(request.form("prod_cod_presion"))))
prod_cod_temperatura = EliminaInyeccionSQL(h(trim(request.form("prod_cod_temperatura"))))
prod_cod_inflamabilidad = EliminaInyeccionSQL(h(trim(request.form("prod_cod_inflamabilidad"))))


' Lo insertamos en la tabla de productos
sql="INSERT INTO dn_auto_productos (nombre, cod_proceso, id_ecogente, frases_r, cod_estado, cod_presion, cod_temperatura, cod_inflamabilidad) VALUES ('"&prod_nombre&"','"&prod_cod_proceso&"',"&session("id_ecogente2")&", '"&prod_frases_r&"', '"&prod_cod_estado&"', '"&prod_cod_presion&"', '"&prod_cod_temperatura&"', '"&prod_cod_inflamabilidad&"')"
objConnection2.execute(sql),,adexecutenorecords
'response.write "<br><br>"&sql


' ##########################
' ### COMPONENTES
' ##########################

num_componentes = EliminaInyeccionSQL(request.form("num_componentes"))
response.write "<br/>Hay "&num_componentes&" componentes"

' Buscamos el ID del producto recién creado
sql="SELECT TOP 1 id FROM dn_auto_productos WHERE nombre = '"&prod_nombre&"' AND id_ecogente="&session("id_ecogente2")
set objRst = objConnection2.execute(sql)
id_producto = objRst("id")
objRst.close
set objRst = nothing
'response.write "<br/>El producto recién creado tiene ID = "&id_producto

' Insertamos los componentes

for i=1 to num_componentes
	' Cogemos los datos del componente actual
	comp_nombre = EliminaInyeccionSQL(request.form("comp"&i&"_nombre"))
	comp_numero_tipo = EliminaInyeccionSQL(request.form("comp"&i&"_numero_tipo"))
	comp_numero = EliminaInyeccionSQL(request.form("comp"&i&"_numero"))
	comp_porcentaje = EliminaInyeccionSQL(request.form("comp"&i&"_porcentaje"))
	comp_frases_r = EliminaInyeccionSQL(request.form("comp"&i&"_frases_r"))

	sql="INSERT INTO dn_auto_componentes (id_producto, nombre, numero_tipo, numero, porcentaje, frases_r, id_ecogente) VALUES ("&id_producto&", '"&comp_nombre&"', '"&comp_numero_tipo&"', '"&comp_numero&"', '"&comp_porcentaje&"', '"&comp_frases_r&"', "&session("id_ecogente2")&")"
	'response.write "<br><br>"&sql
	objConnection2.execute(sql),,adexecutenorecords
next

' Cerramos la conexión
cerrarconexion

' Redirigimos a herramienta
response.redirect("dn_auto_herramienta.asp")
%>
