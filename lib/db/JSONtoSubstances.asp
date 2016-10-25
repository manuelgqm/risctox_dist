<!--#include file="../vbsJSON.asp"-->
<%
Dim fso, json, str, o, i
Set json = New VbsJson
Set fso = Server.CreateObject("Scripting.FileSystemObject")
str = fso.OpenTextFile(Server.MapPath("local/substancesRaw.JSON")).ReadAll
Set o = json.Decode(str)
response.write o("data")(3)("nombre")
%>