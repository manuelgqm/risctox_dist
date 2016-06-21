<!--#include file="../EliminaInyeccionSQL.asp"-->
<%
function obtainSanitizedQueryParameter(parameterName)
  id_sustancia = EliminaInyeccionSQL(request(parameterName))
  obtainSanitizedQueryParameter = id_sustancia
end function
%>
