<!--#include file="../dn_conexion.asp"-->
<!--#include file="../adovbs.inc"-->
<!--#include file="../dn_restringida.asp"-->
<!--#include file="../dn_funciones_comunes.asp"-->

<%
response.ContentType="text/html; charset=iso-8859-1"
%>

<%
' #################################################
' ### EVALUACIÓN DE PRODUCTOS
' #################################################

dim riesgo, razon
dim mayorriesgo_aguda, mayorrazon_aguda
dim mayorriesgo_cronica, mayorrazon_cronica
dim mayorriesgo_ecotoxicidad, mayorrazon_ecotoxicidad
dim mayorriesgo_fuego, mayorrazon_fuego
dim mayorriesgo_exposicion, mayorrazon_exposicion
dim mayorriesgo_proceso, mayorrazon_proceso

' FASE 1: Cogemos los IDs de los productos a evaluar, sus datos y contamos cuántos hay
' ------------------------------------------------------------------------------------

ids_evaluar = EliminaInyeccionSQL(request("idcheck"))
'response.write ids_evaluar

if (ids_evaluar = "") then
	' No ha seleccionado 

	cuantos = 0
else
	' Hay IDs seleccionados, los mostramos
	' Mostramos  los productos. Añadimos id de usuario para mayor seguridad

	sql = "SELECT nombre, id FROM dn_auto_productos WHERE ((id IN ("&ids_evaluar&")) AND (id_ecogente="&session("id_ecogente2")&"))"
	set objRst = objConnection2.execute(sql)

	' Cogemos registros

	if not objRst.eof then 
		arrayDatos=objRst.GetRows
		cuantos = ubound(arrayDatos,2)+1
	end if

	' Cerrar

	objRst.close
	set objRst=nothing
	
end if

' FASE 2: Mostramos los datos
' ---------------------------

if (cuantos = 0) then
	' No hay productos seleccionados

%>
	<p>Seleccione los productos desde la cesta y pulse el botón "Autoevaluar".<p>
<%
else
	' Hay productos	seleccionados
%>
<div id="tabla_resultado_autoevaluacion">
	<table class="dn_auto_tabla" border="0"  cellpadding="2" cellspacing="2" align="center">
		<tr>
			<th><a href="javascript:mostrar_ayuda('TOXICIDAD AGUDA');">Toxicidad aguda</a></th>
			<th><a href="javascript:mostrar_ayuda('TOXICIDAD CRONICA');">Toxicidad cr&oacute;nica</a></th>
			<th><a href="javascript:mostrar_ayuda('MEDIO AMBIENTE');">Medio ambiente</a></th>
			<th><a href="javascript:mostrar_ayuda('FUEGO Y EXPLOSION');">Fuego y explosi&oacute;n</a></th>
			<th><a href="javascript:mostrar_ayuda('FACILIDAD DE EXPOSICION');">Facilidad de exposici&oacute;n</a></th>
			<th><a href="javascript:mostrar_ayuda('PROCESO');">Proceso</a></th>
		</tr>
	<%
	for contadorFilas=0 to (cuantos-1)
		
		'PARA CADA PRODUCTO:
		'1/ escribimos su nombre
		
		nombre = arrayDatos(0,contadorFilas)
		id = arrayDatos(1,contadorFilas)
%>
		<tr>
			<td nowrap="nowrap" colspan="6" class="producto" id="<%=id%>_producto"><a href="javascript:despliegame(<%=id%>)"><%=replace(nombre," ", "&nbsp;")%></a>&nbsp;<a href="javascript:pliegame(<%=id%>)" id='<%=id%>_plegar' class="oculta" style="margin:0 10px; color:#006699; ">&uarr;&nbsp;Plegar</a></td>
		</tr>
<%
		'2/ consultamos componentes  y toxicidades 
%>
		<%comptox(id)%>
<%
		
	next	' Siguiente producto
%>
	</table>
</div>
<br/>
	<!-- xip -->
	<p style="width:100%; text-align:center">
		<form name="form_imprimir" action="imprimir_cesta.asp" target="_blank" method="post" style="border-bottom:0">
			<textarea name="cesta" style="display:none; visibility:hidden;"></textarea>
			<input type="button" class="boton" value="imprimir resultado" onclick="document.form_imprimir.cesta.value=document.getElementById('tabla_resultado_autoevaluacion').innerHTML; document.form_imprimir.submit();" />
		</form>
	</p>
	<!-- /xip -->	
<%
end if ' if cuantos = 0
%>

<p>
<!--
La información que aparece es una evaluación básica sobre los niveles de riesgo de cada uno de los productos que has consultado. Puedes ver también la evaluación de cada una de las sustancias que componen el producto pinchando sobre el nombre del producto.<br/><br/>
Los productos y sustancias se comparan por columnas, esto es  por tipos de riesgo (toxicidad aguda; ecotoxicidad; etc.). Además, se deben tener en cuenta las condiciones de uso del producto. A la vista de los niveles de riesgo identificados por la herramienta deberás de optar por el producto o sustancia que presente los niveles más bajos.<br/><br/>
No existen niveles de exposición seguros para las sustancias con riesgo de toxicidad crónica muy alta, por lo que la única medida segura es eliminar el riesgo por sustitución del agente químico por otro menos peligroso.<br/><br/>
Puedes encontrar más información sobre cómo actuar frente al riesgo químico consultando nuestras <a href="http://www.istas.net/ecoinformas/web/index.asp?idpagina=720">guías</a>.
-->
<p>
La información que aparece es una evaluación básica sobre los niveles de riesgo de cada uno de los productos que has consultado. Puedes ver también la evaluación de cada una de las sustancias que componen el producto pinchando sobre el nombre del producto.
</p>
<p>
Los productos y sustancias se comparan por columnas, esto es por tipos de riesgo (toxicidad aguda; ecotoxicidad; etc.). Además, se deben tener en cuenta las condiciones de uso del producto. A la vista de los niveles de riesgo identificados por la herramienta deberás de optar por el producto o sustancia que presente los niveles más bajos.
</p>
<p>
Cuando no existe información sobre ensayos de toxicidad o de sensibilización de la piel, el riesgo de toxicidad aguda se considera alto.
</p>
<p>
Cuando no existe información sobre ensayos de mutagenicidad, la sustancia o preparado debería categorizarse al menos en alto riesgo, en la columna de toxicidad crónica.
</p>
<p>
Si no existe información disponible de ensayos de efectos irritantes sobre la piel o mucosas, la sustancia o preparado debería categorizarse al menos, en el apartado de bajo riesgo para toxicidad aguda.
</p>
No existen niveles de exposición seguros para las sustancias con riesgo de toxicidad crónica muy alta, por lo que la única medida segura es eliminar el riesgo por sustitución del agente químico por otro menos peligroso.
<p>
Puedes encontrar más información sobre cómo actuar frente al riesgo químico consultando nuestras <a href="http://www.istas.net/risctox/index.asp?idpagina=720" target="_blank">guías</a>. 
</p>

</p>

<%
cerrarconexion
%>


<!--<input type="button" name="boton_autoevaluar" id="boton_autoevaluar" value="Autoevaluar" onclick="autoevaluar();" /> -->

<%
sub comptox(id) 'maqueta componentes y toxicidad; a partir de un id de producto (la fila de nombre ya esta escrita), escribirá la fila de toxicidades del producto (para cada tiporiesgo, la mayor, sea esta la del producto o la de alguno de sus componentes), y para cada componente, su nombre y sus toxicidades
		
		
		
		filatoxproducto=""
		filascomponentes=""

		'consultamos riesgo y razon del producto para cada columna; provisionalmente, son las mayores toxicidades, pero si un componente la tiene mayor, la sustituirá
		consultox id,0,"aguda"
		mayorriesgo_aguda=riesgo
		mayorrazon_aguda=razon
		
		consultox id,0,"cronica"
		mayorriesgo_cronica=riesgo
		mayorrazon_cronica=razon
		
		consultox id,0,"ecotoxicidad"
		mayorriesgo_ecotoxicidad=riesgo
		mayorrazon_ecotoxicidad=razon
		
		consultox id,0,"fuego"
		mayorriesgo_fuego=riesgo
		mayorrazon_fuego=razon
		
		consultox id,0,"exposicion"
		mayorriesgo_exposicion=riesgo
		mayorrazon_exposicion=razon
		
		consultox id,0,"proceso"
		mayorriesgo_proceso=riesgo
		mayorrazon_proceso=razon

		'COMPONENTES
		sql = "SELECT nombre, id FROM dn_auto_componentes WHERE id_producto=" &id
		set objRst = objConnection2.execute(sql)
	
		' Cogemos registros
		if not objRst.eof then 
			arrayComps=objRst.GetRows
			cuantosComps = ubound(arrayComps,2)
			haycomponentes=true
		else
			haycomponentes=false
		end if
		
		' Cerrar
		objRst.close
		set objRst=nothing

		if haycomponentes then 'todos los componentes y sus toxicidades se muestran en una tabla con el id  

			FOR contadorFilasComp=0 to cuantosComps
				nombrecomp = arrayComps(0,contadorFilasComp)
				idcomp = arrayComps(1,contadorFilasComp)
				
				'PARA CADA COMPONENTE:
				'1/ creamos fila con su nombre
				filascomponentes=filascomponentes& "<tr><td colspan='6' class='componente2'>" &nombrecomp& "</td></tr>"
	
				'2/ creamos fila toxicidades
				'y si el riesgo de este componente es mayor que el mayor riesgo que teniamos (para cada tipo de riesgo), sustituimos; si es igual, concatenamos razones (luego habra que sustituir repetidas); si es menor, no cambia nada
	
				filascomponentes=filascomponentes& "<tr>"
				
				consultox idcomp,1,"aguda"
				
				filascomponentes=filascomponentes& pintatox(riesgo,razon)	
				if riesgo>mayorriesgo_aguda then
						mayorriesgo_aguda=riesgo
						mayorrazon_aguda=razon
				else
						if riesgo=mayorriesgo_aguda then
							'mayorriesgo_aguda=riesgo 'lo dejamos como está
							mayorrazon_aguda=mayorrazon_aguda &razon 'concatenamos razones
						end if
				end if
				
				consultox idcomp,1,"cronica"
				filascomponentes=filascomponentes& pintatox(riesgo,razon)	
				if riesgo>mayorriesgo_cronica then
						mayorriesgo_cronica=riesgo
						mayorrazon_cronica=razon
				else
						if riesgo=mayorriesgo_cronica then
							'mayorriesgo_cronica=riesgo 'lo dejamos como está
							mayorrazon_cronica=mayorrazon_cronica &razon 'concatenamos razones
						end if
				end if
				
				consultox idcomp,1,"ecotoxicidad"
				filascomponentes=filascomponentes& pintatox(riesgo,razon)	
				if riesgo>mayorriesgo_ecotoxicidad then
						mayorriesgo_ecotoxicidad=riesgo
						mayorrazon_ecotoxicidad=razon
				else
						if riesgo=mayorriesgo_ecotoxicidad then
							'mayorriesgo_ecotoxicidad=riesgo 'lo dejamos como está
							mayorrazon_ecotoxicidad=mayorrazon_ecotoxicidad &razon 'concatenamos razones
						end if
				end if
				
				consultox idcomp,1,"fuego"
				filascomponentes=filascomponentes& pintatox(riesgo,razon)	
				if riesgo>mayorriesgo_fuego then
						mayorriesgo_fuego=riesgo
						mayorrazon_fuego=razon
				else
						if riesgo=mayorriesgo_fuego then
							'mayorriesgo_fuego=riesgo 'lo dejamos como está
							mayorrazon_fuego=mayorrazon_fuego &razon 'concatenamos razones
						end if
				end if
				
				consultox idcomp,1,"exposicion"
				filascomponentes=filascomponentes& pintatox(riesgo,razon)	
				if riesgo>mayorriesgo_exposicion then
						mayorriesgo_exposicion=riesgo
						mayorrazon_exposicion=razon
				else
						if riesgo=mayorriesgo_exposicion then
							'mayorriesgo_exposicion=riesgo 'lo dejamos como está
							mayorrazon_exposicion=mayorrazon_exposicion &razon 'concatenamos razones
						end if
				end if
				
				consultox idcomp,1,"proceso"
				filascomponentes=filascomponentes& pintatox(riesgo,razon)	
				if riesgo>mayorriesgo_proceso then
						mayorriesgo_proceso=riesgo
						mayorrazon_proceso=razon
				else
						if riesgo=mayorriesgo_proceso then
							'mayorriesgo_proceso=riesgo 'lo dejamos como está
							mayorrazon_proceso=mayorrazon_proceso &razon 'concatenamos razones
						end if
				end if

				filascomponentes=filascomponentes& "</tr>"
	
			NEXT	' Siguiente componente
			
		end if 'si hay componentes
		
		'creamos fila toxicidades producto, con las mayores de cada columna
		filatoxproducto="<tr>"
		filatoxproducto=filatoxproducto& pintatox(mayorriesgo_aguda,mayorrazon_aguda)
		filatoxproducto=filatoxproducto& pintatox(mayorriesgo_cronica,mayorrazon_cronica)
		filatoxproducto=filatoxproducto& pintatox(mayorriesgo_ecotoxicidad,mayorrazon_ecotoxicidad)
		filatoxproducto=filatoxproducto& pintatox(mayorriesgo_fuego,mayorrazon_fuego)
		filatoxproducto=filatoxproducto& pintatox(mayorriesgo_exposicion,mayorrazon_exposicion)
		filatoxproducto=filatoxproducto& pintatox(mayorriesgo_proceso,mayorrazon_proceso)
		filatoxproducto=filatoxproducto& "</tr>"
		
		'escribimos cadena de toxicidades producto
		response.write filatoxproducto
		'escribimos cadena de componentes (si había, metiendolos en una tabla que se oculta/desoculta)
		if haycomponentes then 'todos los componentes y sus toxicidades se muestran en una tabla con el id  
%>
		<tr><td colspan="6">
		<table id='<%=id%>' class="oculta dn_auto_tabla">
		<%=filascomponentes%>
		</table>		
		</td></tr>
<%
		end if

end sub
%>

<%
sub consultox(id,tipo,tiporiesgo)

				set cmdsus=Server.CreateObject("ADODB.Command")
				   With cmdsus
					.ActiveConnection=objConnection2
					.CommandText="dn_autoevaluar"
					.CommandType=adCmdStoredProc					
					.Parameters.Append  .CreateParameter("@id", adinteger, adParamInput, , id) 'id de producto/componente
					.Parameters.Append  .CreateParameter("@tipo", adboolean, adParamInput, , tipo) '0:producto 1:componente
					.Parameters.Append  .CreateParameter("@tiporiesgo", advarchar, adParamInput, 20, tiporiesgo)					
					.Parameters.Append  .CreateParameter("@riesgo", adtinyint, adParamOutput) 					
					.Parameters.Append  .CreateParameter("@razon", advarchar, adParamoutput, 1000)							
					.Execute,,adexecutenorecords
					riesgo=.Parameters("@riesgo")
					razon=.Parameters("@razon")
					'response.write "<p>RIESGO: " &riesgo& "</p>"
					'response.write "<p>RAZON: " &razon& "</p>"
					
				   End With 
				set cmdsus=nothing		

end sub
%>

<%
function pintatox(riesgo,razon)

			cadtox="<td class='toxicidad tox_" &riesgo& "'>" 
			if riesgo=0 then
				cadtox=cadtox& "<a href=" &chr(34)& "javascript:mostrar_ayuda('INFORMACION INSUFICIENTE')" &chr(34)& "><i>Información insuficiente</i></a>"
			else			
				'cadtox=cadtox& "<span class='tox_" &riesgo& "'>"
				cadtox=cadtox& "<strong>"
				select case riesgo
					case 5: cadtox=cadtox& "Muy alto: "
					case 4: cadtox=cadtox& "Alto: "
					case 3: cadtox=cadtox& "Medio: "
					case 2: cadtox=cadtox& "Bajo: "
					case 1: cadtox=cadtox& "Muy bajo: "
				end select
				cadtox=cadtox& "</strong>"
				
				'quitamos duplicados y añadimos enlace
				cadrazones=""
				arrenl= split(razon, ", ")
				for i=0 to ubound(arrenl)
          'response.write "<br/>"&arrenl(i)        
          'response.write "<br/>"&cadrazones
          'response.write "<br>*"&instr(cadrazones, ">" &arrenl(i)& "<" ) & "<br/>"

					if ( (instr(cadrazones, ">" &arrenl(i))=0) and arrenl(i)<>"" ) then
            ' Parche para no mostrar dos veces cancerígeno
            if not ((arrenl(i)="Cancerígeno") and ((instr(cadrazones, "R40")>0) or (instr(cadrazones, "R45")>0) or (instr(cadrazones, "R49")>0))) then
              cadrazones=cadrazones & "<a href=" &chr(34)& "javascript:mostrar_ayuda('" &arrenl(i)& "')" &chr(34)& ">" &arrenl(i)& explica_razon(arrenl(i))& "</a>, "
            end if
          end if
				next
				cadtox=cadtox& left(cadrazones,len(cadrazones)-2)
				
				
			end if
			'cadtox=cadtox& "</span>"
			cadtox=cadtox& "</td>"	
			pintatox=cadtox  
end function


function explica_razon(byval cadena)
	select case cadena
		case "R40","R45","R49","R40/20","R40/21","R40/22","R40/20/21","R40/20/22","R40/21/22","R40/20/21/22": explica_razon = "(cancer&iacute;geno)"
		case else: explica_razon = ""
	end select
end function
%>


