<%
 	'Option Explicit
	 
 	Dim upl, NewFileName
 	 
 	Set upl = Server.CreateObject("ASPSimpleUpload.Upload")
 	
 	If Len(upl.Form("File1")) > 0 Then
		
		NewFileName = "../estructuras/" & upl.ExtractFileName(upl.Form("File1"))
 	    If upl.SaveToWeb("File1", NewFileName) Then
 	        Response.Write("File successfully written to disk.") 
 	    Else
 	        Response.Write("There was an error saving the file to disk.")
 	    End If
 	End If
%>
	 
 	<html><head><title>ASP Simple Upload Example #1</title></head></title>
 	<body>
 	

<form method="POST" action="aspsimpleupload.asp" enctype="multipart/form-data">
 	Select a file to upload: <input type="file" name="File1" size="30">
 	<input type="submit" name="submit" value="Upload Now">
 	</form>
 	</body>
</html>
