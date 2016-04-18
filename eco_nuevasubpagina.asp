<!--#include file="eco_conexion.asp"-->
<%

	numeracion = EliminaInyeccionSQL(request("numeracion"))
	titulo = unquote(EliminaInyeccionSQL(request("titulo")))
	pagina = unquote(EliminaInyeccionSQL(request("pagina")))
	tipo = EliminaInyeccionSQL(request("tipo"))
	destino = EliminaInyeccionSQL(request("destino"))
	if destino="" then destino=0
	visible = EliminaInyeccionSQL(request("visible"))

	if tipo=3 then						'-- otra página existente
		sql = "SELECT titulo FROM WEBISTAS_PAGINAS WHERE idpagina="&destino
		set objRecordset = Server.CreateObject ("ADODB.Recordset")
		set objRecordset = OBJConnection.Execute(sql)
		if objrecordset.eof then
			grabar = "no"
		else
			titulo = objRecordset("titulo")
		end if
	end if

	if tipo=4 then						'-- ficha de recursos para pymes
		sql = "SELECT nomficha FROM PYM_FICHAS WHERE idficha="&destino
		set objRecordset = Server.CreateObject ("ADODB.Recordset")
		set objRecordset = OBJConnection.Execute(sql)
		if objrecordset.eof then
			grabar = "no"
		else
			titulo = objRecordset("nomficha")
		end if
	end if

	if tipo=5 then						'-- carpeta de recursos para pymes
		sql = "SELECT nombre FROM PYM_INTRANET WHERE id="&destino
		set objRecordset = Server.CreateObject ("ADODB.Recordset")
		set objRecordset = OBJConnection.Execute(sql)
		if objrecordset.eof then
			grabar = "no"
		else
			titulo = objRecordset("nombre")
		end if
	end if

	if tipo=6 then						'-- carpeta de enlaces
		sql = "SELECT nombre FROM ENL_TEMAS WHERE id="&destino
		set objRecordset = Server.CreateObject ("ADODB.Recordset")
		set objRecordset = OBJConnection.Execute(sql)
		if objrecordset.eof then
			grabar = "no"
		else
			titulo = objRecordset("nombre")
		end if
	end if

	if grabar<>"no" then
		sql = "INSERT into WEBISTAS_PAGINAS (titulo,pagina,numeracion,fecha,hora,fecha_modificacion,autor,tipo,destino,visible) values ('"&titulo&"','"&pagina&"','"&numeracion&"','"&date()&"','"&time()&"','"&date()&"',"&session("idgente")&","&tipo&","&destino&","&visible&");"
		Set objRecordset = Server.CreateObject ("ADODB.Recordset")
		Set objRecordset = OBJConnection.Execute(sql)
		'response.write sql

		sql5 = "SELECT MAX(idpagina) AS nuevoid FROM WEBISTAS_PAGINAS"
		Set objRecordset5 = Server.CreateObject ("ADODB.Recordset")
		objRecordset5.Open sql5,OBJConnection,adOpenKeyset
		nuevoid = objRecordset5("nuevoid")
	end if

FUNCTION unQuote(s)
  pos = Instr(s, "'")
  While pos > 0
    s = Mid(s,1,pos) & "'" & Mid(s,pos+1)
    pos = InStr(pos+2, s, "'")
  Wend
  pos = Instr(s, """")
  While pos > 0
    s = Mid(s,1,pos-1) & "''" & Mid(s,pos+1)
    pos = InStr(pos+2, s, """")
  Wend
  unQuote = Trim(s)
END FUNCTION
%>

<html>

<head>
<title>Nueva subpágina de la página <%=id%></title>
<base target="_self">

</head>
<link rel="stylesheet" type="text/css" href="ot.css">
<body bgcolor="#F4F4F9" topmargin="20" leftmargin="20">
<% if grabar<>"no" then %>
<p class="negro">Guardada nueva página...</p>
<script>
	parent.frames.izquierda.indice_arriba.location.reload();
	location.href='eco_editarpagina.asp?id=<%=nuevoid%>';
</script>
<% else %>
<script>
	alert('El destino no tiene correspondencia con ninguna página, ficha o carpeta');
	history.back();
</script>
<% end if %>
</body>
</html>