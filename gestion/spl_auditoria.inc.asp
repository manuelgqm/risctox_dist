<!--#include file="clases/auditoria.inc.asp"-->
<%
sub auditaYCierraConexion(aud_accion,aud_entidad,aud_descripcion)
	if not(session("usuario")="") and (not aud_accion="buscar") then
		dim objAuditoria
		set objAuditoria = new Auditoria	
		call objAuditoria.setProperty("usuario",session("usuario"))	
		call objAuditoria.setProperty("fecha",FormatDateTime(Now))
		call objAuditoria.setProperty("ip",Request.ServerVariables("remote_addr"))
		call objAuditoria.setProperty("navegador",Request.ServerVariables("http_user_agent"))
		call objAuditoria.setProperty("accion",aud_accion)
		call objAuditoria.setProperty("entidad",aud_entidad)
		call objAuditoria.setProperty("descripcion",aud_descripcion)		
		objAuditoria.registra()
	end if
	cerrarconexion
end sub
%>