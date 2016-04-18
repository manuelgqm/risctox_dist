<%
ordenacion=EliminaInyeccionSQL(request("ordenacion"))
sentido=EliminaInyeccionSQL(request("sentido"))
nregs=EliminaInyeccionSQL(request("nregs"))

'valores de busqueda por defecto

if ordenacion="" then ordenacion="sus.nombre"
if sentido="" then sentido=""
if nregs="" then nregs=50


if busc="" then

else

%>
<!--#include file="../dn_buscador_sustancias_condiciones.asp"-->
<%

	'RESULTADOS DE BUSQUEDA (para busc 1 y busc 2)
	'seleccionamos datos a mostrar de los x registros que toquen
	if hr>0 then

			'vemos que registros hay que mostrar
			registroini=(pag*nregs)-nregs
			'response.write "<p>registroini=" &registroini& "</p>"

			registrofin=registroini+nregs
			'response.write "<p>registrofin=" &registrofin& "</p>"

			if registrofin>=hr-1 then
				registrofin=hr
      			'response.write "<p>registrofin era mayor, ahora=" &registrofin& "</p>"
			end if

			registrofin=registrofin-1
			'response.write "<p>registrofin corregido=" &registrofin& "</p>"

		arrayx = split(arr, ",")

		FOR i=registroini to registrofin
			cadenaids=cadenaids  &arrayx(i)&","
		NEXT
		'quitamos la ultima coma
		cadenaids= left(cadenaids,len(cadenaids)-1)

%>
<!--#include file="dn_buscador_sustancias_lista.asp"-->
<%
	end if 'hr>0
end if 'busc

if hr=1 then
			response.redirect(unico_enlace)
end if
%>