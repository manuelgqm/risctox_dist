		<%
		' INCLUDE QUE MUESTRA LA CESTA DE PRODUCTOS DEL USUARIO, COMO FORMULARIO
		sql="select id, nombre FROM dn_auto_productos WHERE id_ecogente="&session("id_ecogente2")
		'response.write sql
		set objRst = Server.CreateObject("ADODB.Recordset")
		objRst.Open sql, objConnection2, adOpenStatic, adCmdText
					
		cuantos=objRst.recordcount
		if not objRst.eof then 
			arrayDatos=objRst.GetRows	
		end if
				
		objRst.close
		set objRst = nothing
		%>

		<%
		if (cuantos = 0) then
			' No hay productos
		%>
			<p>Actualmente no hay ningún producto en tu cesta.<p>
		<%
		else
			' Hay productos	
		%>
			<p>Estos son los productos que has introducido. Puedes pulsar sobre ellos para ver sus características, o seleccionar uno o varios para evaluar y comparar su peligrosidad, o eliminarlos de la cesta.</p>
			<form id="cesta_productos" name="cesta_productos" action="#" method="post">
			<table border="0" width="100%">
			<%
			for contadorFilas=0 to (cuantos-1)
				id = arrayDatos(0,contadorFilas)
				nombre = arrayDatos(1,contadorFilas)
			%>
				<!-- PRODUCTO <%=id%> -->
				<%
				sql = "SELECT COUNT(*) AS num_componentes FROM dn_auto_componentes WHERE id_producto="&id
				set objRst=objConnection2.execute(sql)
				num_componentes = objRst("num_componentes")
				objRst.close()
				set objRst = nothing
				%>
				<tr>
					<td align="right"><input type="checkbox" name="idcheck" value="<%=id%>" /></td>
					<td><strong><a href="javascript:mostrar_producto(<%=id%>)"><%=nombre%></a></strong> (<%=num_componentes%> componentes)</td>
				</tr>
			<%
				next	' Siguiente producto
			%>
				<tr>
					<td colspan="2" align="center"><input type="button" class="boton2" name="boton_autoevaluar" id="boton_autoevaluar" value="evaluar / comparar" onclick="autoevaluar();" />&nbsp;<input type="button" class="boton2"  name="boton_eliminar" id="boton_eliminar" value="eliminar" onclick="eliminar();" /></td>
				</tr>
			</table>
			<input type="hidden" id="ids_chequeados" name="ids_chequeados" value="">
			</form>
		<%
		end if ' if cuantos = 0
		%>
	<div id="mensaje_cesta" class="oculta"></div>
