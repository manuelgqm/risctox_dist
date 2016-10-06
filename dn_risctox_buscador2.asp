<!--#include file="adovbs.inc"-->
<!--#include file="config/dbConnection.asp"-->
<!--#include file="dn_funciones_comunes.asp"-->
<!--#include file="dn_funciones_texto.asp"-->
<!--#include file="dn_restringida.asp"-->
<%
const filtro = 0 'este filtro diferencia a este buscador del de sustancias: si esta a true, muestra solo las que son toxicas y tienen alternativas
dim displayMode : displayMode = EliminaInyeccionSQL(request.form("displayMode"))
if displayMode = "" then
	displayMode = "search"
end if
%>
<!--#include file="lib/db/substancesSearch.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ISTAS: risctox</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="Title" content="ECOinformas" />
<meta name="Author" content="XiP multimèdia" />
<meta name="description" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Subject" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Keywords" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Language" content="Spanish" />
<meta name="Revisit" content="15 days" />
<meta name="Distribution" content="Global" />
<meta name="Robots" content="All" />

<link rel="stylesheet" type="text/css" href="estructura.css">
<script type="text/javascript">
	function cambiapag(paginadest){
		var frm = document.forms["myform"]; 
		frm.displayMode.value = "pagination";
		frm.currentPageNumber.value = paginadest;
		frm.submit();
	};

	function primerapag(){
		var frm = document.forms["myform"]; 
		if ( (frm.nombre.value.length < 3) && (frm.tipobus.options[frm.tipobus.selectedIndex].value == "parte") ){
			alert("Por favor, teclee al menos 3 caracteres para buscar por nombre");
		} else {
			frm.displayMode.value=1;
			frm.currentPageNumber.value=1;
			frm.submit();
		};
	};
</script>
</head>
<body>
<div id="contenedor">
	<div id="sombra_arriba"></div>
  	<div id="sombra_lateral">
		<div id="caja">
			<!--#include file="dn_cabecera.asp"-->
			<div id="texto">

				<div class="texto">
				<!-- ################ CONTENIDO ###################### -->
					<table width="100%" border="0">
						<tr>
							<td></td>
							<td align="right"><input type="button" name="volver" class="boton" value="Volver a la portada de risctox" onClick="window.location='./index.asp';"></td>
						</tr>
					</table>
	
					<p class=titulo3>Base de datos de sustancias t&oacute;xicas y peligrosas RISCTOX </p>

					<form action="dn_risctox_buscador2.asp" method="post" name="myform" onSubmit="primerapag();">
						<input type="hidden" name='displayMode' value='<%=displayMode%>' />	
						<input type="hidden" name='currentPageNumber' value='<%=currentPageNumber%>' />	
						<input type="hidden" name='numRecordsFound' value='<%=numRecordsFound%>' />		
						<input type="hidden" name='arr' value='<%=arr%>' />
						<input type="hidden" name='ordenacion' value='<%=ordenacion%>' />
						<input type="hidden" name='numRecordsByPage' value='<%=numRecordsByPage%>' />				
						<table class="tabla3" width="90%" align="center">
							<tr><td colspan="3" class="subtitulo3">Buscador de sustancias</td></td></tr>
							<tr>
								<td align="right"><strong>Nombre</strong></td>
								<td><input type="text" name="nombre" value="<%=nombre%>" />
								<select name="tipobus">
								<option value="exacto" <%if tipobus="exacto" then response.write "selected"%>>nombre exacto</option>

								<option value="parte" <%if tipobus="parte" then response.write "selected"%>>parte del nombre</option>

								</select></td>
							</tr>
							<tr>
								<td align="right"><strong>Número CAS/CE/RD</strong></td>
								<td><input type="text" name="numero" value="<%=numero%>" /></td>
							</tr>	
							<tr>
								<td colspan="2" align="center"><input type="submit" value="Buscar" /> <input type="reset" value="Borrar" /></td>
							</tr>
						</table>
						<%
						if displayMode<>"" then
							if numRecordsFound = 0  then
							%>
							<fieldset id="flashmsg">
								<legend class="advertencia"><strong>Advertencia</strong></legend>
								No se encontraron registros que coincidan con su consulta.
							</fieldset>
							<%
							else
								response.Write("<p class='neg' style='margin:15px 0; padding:10px;'>Se han encontrado " & numRecordsFound & " registros. Se muestran registros del " &currentPageInitialRecordNumber+1& " al " &currentPageFinalRecordNumber+1& ":</p>")
							%>		
							<%= tablares %>
							<% if numRecordsFound>numRecordsByPage then %>		
							<div align='center' style="margin:20px 10px; background-color: #3399CC; padding:3px;"><%= obtainPagerHtml(numRecordsFound, numRecordsByPage, currentPageNumber)%></div>
							<% end if %>		
						<%
							end if
						end if
						%>
					</form>
					<!-- ############ FIN DE CONTENIDO ################## -->
					<br>
					<br>
					Esta página ha sido desarrollada por <a href="http://www.istas.ccoo.es/" target="_blank"><b>ISTAS</b></a> que es una Fundación de <a href="http://www.ccoo.es/" target="_blank"><font color="#FF0000"><b>CC.OO.</b></font></a><br>
				</div>
				<p>&nbsp;</p>
			</div>
			<img src="imagenes/pie_risctox.gif" width="708" border="0">
		</div>
	</div> <!-- /sombra_lateral -->
	<div id="sombra_abajo"></div>
</div>
</body>
</html>

<% cerrarconexion %>

<%
function obtainPagerHtml(numRecordsFound, numRecordsByPage, currentPageNumber)
	dim html : html = "<strong>Páginas: </strong><br />"
	dim pagesCount : pagesCount = roundsup(numRecordsFound/numRecordsByPage)
	if currentPageNumber>1 then
		html = html & "<a href=""#"" onclick=""cambiapag(" & currentPageNumber-1 & ")"">&lt; Anterior</a>"
	end if
	dim i
	for i = 1 to pagesCount
		if cint(i) = cint(currentPageNumber) then
			currentPageHtml = " <b>" & i & "</b>"
		else
			currentPageHtml = " <a href=""#"" onclick=""cambiapag(" & i & ")"">" & i & "</a>"
		end if
		html = html & currentPageHtml
	next
	if cint(currentPageNumber) < cint(pagesCount) then
		html = html & "&nbsp;<a href=""#"" onclick=""cambiapag(" & currentPageNumber + 1 & ")"">Siguiente &gt;</a>"
	end if

	obtainPagerHtml = html
end function
%>