<!--#include file="eco_conexion2.asp"-->
<%
  id_tema = EliminaInyeccionSQL(request("idtema"))
	
	mos1 = EliminaInyeccionSQL(request("mos1"))
	mos2 = EliminaInyeccionSQL(request("mos2"))
	mos3 = EliminaInyeccionSQL(request("mos3"))
	mos4 = EliminaInyeccionSQL(request("mos4"))
	mos5 = EliminaInyeccionSQL(request("mos5"))
	mos6 = EliminaInyeccionSQL(request("mos6"))
	ord  = EliminaInyeccionSQL(request("ord"))
	fil1 = EliminaInyeccionSQL(request("fil1"))
	fil2 = EliminaInyeccionSQL( request("fil2"))
	fil3 = EliminaInyeccionSQL(request("fil3"))
	fil4 = EliminaInyeccionSQL(request("fil4"))
	fil5 = EliminaInyeccionSQL(request("fil5"))
	fil6 = EliminaInyeccionSQL(request("fil6"))
	
	if mos1="" and mos2="" and mos3="" and mos4="" and mos5="" and mos6="" then 
		mos1="1"
		mos2="1"
		mos3="1"
	end if
	if fil2="" then fil2="1/1/01"
	if fil3="" then fil3=date()
	if fil4="" then fil4="1/1/01"
	if fil5="" then fil5=date()

%>
<html>
<head>
<title>Opciones para visualizar el árbol</title>
<link rel="stylesheet" type="text/css" href="panelcontrol.css">
<script>

function aplica()
{	mos1="0"; mos2="0"; mos3="0"; mos4="0"; mos5="0"; mos6="0";
	if (formulario.ver_num.checked) {mos1="1"}
	if (formulario.ver_nom.checked) {mos2="1"}
	if (formulario.ver_ico.checked) {mos3="1"}
	if (formulario.ver_fal.checked) {mos4="1"}
	if (formulario.ver_fmo.checked) {mos5="1"}
	if (formulario.ver_aut.checked) {mos6="1"}
	
	ord = formulario.orden_temas.value;

	fil1 = formulario.filtro_ram.value;
	fil2 = formulario.filtro_fal1.value;
	fil3 = formulario.filtro_fal2.value;
	fil4 = formulario.filtro_fmo1.value;
	fil5 = formulario.filtro_fmo2.value;
	fil6 = formulario.filtro_pal.value;

	window.opener.location.href='eco_tema_arbol.asp?mos1='+mos1+'&mos2='+mos2+'&mos3='+mos3+'&mos4='+mos4+'&mos5='+mos5+'&mos6='+mos6+'&ord='+ord+'&fil1='+fil1+'&fil2='+fil2+'&fil3='+fil3+'&fil4='+fil4+'&fil5='+fil5+'&fil6='+fil6+'&idtema=<%=id_tema%>';
	window.close();
}

</script>

</head>

<body bgcolor="#FFFFFF" topmargin="20" leftmargin="20">
<form name="formulario">
<table align=center cellpadding=0 cellspacing=0 width="98%" style="background: #EEEEEE">
<tr><td class=negro colspan=2 align=center><u><b>Mostrar:</b></u></td></tr>
<tr>
  <td class=negro align=right><input type="checkbox" name="ver_num" <% if mos1="1" then response.write "checked" %>></td>
  <td class=negro align=left>Numeración</td>
</tr>
<tr>
  <td class=negro align=right><input type="checkbox" name="ver_nom" <% if mos2="1" then response.write "checked" %>></td>
  <td class=negro align=left>Nombre tema</td>
</tr>
<tr>
  <td class=negro align=right><input type="checkbox" name="ver_ico" <% if mos3="1" then response.write "checked" %>></td>
  <td class=negro align=left>Iconos</td>
</tr>
<tr>
  <td class=negro align=right><input type="checkbox" name="ver_fal" <% if mos4="1" then response.write "checked" %>></td>
  <td class=negro align=left>Fecha creación</td>
</tr>
<tr>
  <td class=negro align=right><input type="checkbox" name="ver_fmo" <% if mos5="1" then response.write "checked" %>></td>
  <td class=negro align=left>Fecha última modificación</td>
</tr>
<tr>
  <td class=negro align=right><input type="checkbox" name="ver_aut" <% if mos6="1" then response.write "checked" %>></td>
  <td class=negro align=left>Autor/a</td>
</tr>
<tr><td class=negro colspan=2>&nbsp;</td></tr>
</table>
<br>
<table align=center width="98%" style="background: #EEEEEE">
<tr>
  <td class=negro colspan=2 align=center><u><b>Ordenar por:</b></u>&nbsp;</td>
  <td class=negro colspan=2 align="center"><select name="orden_temas" class=campo>
    <option value="ord_num" <% if ord="ord_num" then response.write "selected" %>>Numeración</option>
    <option value="ord_nom" <% if ord="ord_nom" then response.write "selected" %>>Nombre</option>
    <option value="ord_fal" <% if ord="ord_fal" then response.write "selected" %>>Fecha alta</option>
    <option value="ord_fmo" <% if ord="ord_fmo" then response.write "selected" %>>Fecha última modificación</option>
    <option value="ord_aut" <% if ord="ord_aut" then response.write "selected" %>>Autor/a</option>
  </select>
</td></tr>
<tr><td class=negro colspan=2>&nbsp;</td></tr>
</table>
<br>
<table align=center width="98%" style="background: #EEEEEE">
<tr><td class=negro colspan=2 align=center><u><b>Filtrar:</b></u></td></tr>
<tr>
  <td class=negro colspan=2 align=center><select name="filtro_ram" class=campo>
<option value="">- ver ramas principales y toda la rama seleccionada -</option>
<option value="AI" <%if cstr(fil1)="AI" then response.write "SELECTED"%>>- ver árbol completo con todas las ramas abiertas -</option>
<% 	sql = "SELECT idpagina,titulo,numeracion FROM WEBISTAS_PAGINAS WHERE numeracion LIKE 'AI%' ORDER BY "
	for i=1 to 9
	  sql = sql & "ascii(substring(numeracion,"&cstr(i)&",1)),"
	next
	sql = sql & "ascii(substring(numeracion,10,1));"
	
	set objRecordset = Server.CreateObject ("ADODB.Recordset")
	objRecordset.Open sql,objConnection,adOpenKeyset
	do while not objRecordset.eof 
	  numeracion = objrecordset("numeracion")
	  nivel = len(numeracion)
	  texto = ""

	  if len(numeracion)>2 then 
	  	situacion = asc(mid(numeracion,3,1))-64
	  	for i = 4 to nivel
	   	  situacion = situacion & "." & (asc(mid(numeracion,i,1))-64)
	   	  texto = texto&"    "
	  	next
	  	situacion = situacion & ".&nbsp;"
	  else
	  	situacion = ""
	  end if

	  	  if cstr(fil1)=cstr(objRecordset("idpagina")) then
	  	    elegido = "SELECTED"
	  	  else
	  	    elegido = ""
	  	  end if

%>
<option value="<%=objRecordset("idpagina")%>" <%=elegido%>><%=texto & situacion & mid(ucase(objRecordset("titulo")),1,30)%></option>
<%	  objrecordset.movenext
	loop %>
</select></td>
</tr>
<tr>
  <td class=negro align=right>Fecha de alta entre:&nbsp;<input type="text" name="filtro_fal1" class=campo size="10" maxlenght="10" value="<%=fil2%>"></td>
  <td class=negro align=left>y&nbsp;<input type="text" name="filtro_fal2" class=campo size="10" maxlenght="10" value="<%=fil3%>"></td>
</tr>
<tr>
  <td class=negro align=right>Modificado entre:&nbsp;<input type="text" name="filtro_fmo1" class=campo size="10" maxlenght="10" value="<%=fil4%>"></td>
  <td class=negro align=left>y&nbsp;<input type="text" name="filtro_fmo2" class=campo size="10" maxlenght="10" value="<%=fil5%>"></td>
</tr>
<tr>
  <td class=negro colspan=2 align=center>Con la palabra:&nbsp;<input type="text" name="filtro_pal" class=campo size="30" maxlenght="50" value="<%=fil6%>"></td>
</tr>
</table>
<br>
<table align=center width="98%">
<tr>
  <td class=negro colspan=2 align=center><input type="button" value="APLICAR" onClick="aplica()" class=boton></td>
</tr>
</table>
</body>
</html>