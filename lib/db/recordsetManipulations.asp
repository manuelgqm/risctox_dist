<!--#include file="../JSON151.asp"-->
<%
function recordsetToFile(query, connection, url)
	dim fs : set fs = Server.CreateObject("Scripting.FileSystemObject")
	dim path : path = Server.MapPath(url)
	dim fname : set fname = fs.CreateTextFile(path, true)
	dim recordsetJSON : recordsetJSON = recordsetToJSON(query, connection)
	fname.WriteLine(recordsetJSON)
	fname.Close
	set fname = nothing
	set fs = nothing
end function

function recordsetToJSON(query, connection)
	dim recordset
	set recordset = connection.execute(query)
	recordsetToJSON = (new JSON).toJSON("data", recordset, false)
end function
%>