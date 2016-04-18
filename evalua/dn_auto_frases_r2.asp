<!--#include file="../dn_conexion.asp"-->
<!--#include file="../adovbs.inc"-->
<!--#include file="../dn_restringida.asp"-->
<!--#include file="../dn_funciones_comunes.asp"-->
<!--#include file="../dn_funciones_texto.asp"-->

<%
' Cogemos frases R del formulario anterior, las pasamos a opener y cerramos esta ventana
check = EliminaInyeccionSQL(request.form("check"))
idcampo = EliminaInyeccionSQL(request.form("idcampo"))
%>
<script language="JavaScript">
window.opener.document.getElementById("<%=idcampo%>").value="<%=check%>";
window.close();
</script>
<%
cerrarconexion
%>
