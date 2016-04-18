<!--#include file="dn_fun_comunes.asp"-->
<!--#include file="dn_auten.inc"-->
<!--#include file="adovbs.inc"-->
<!--#include file="dn_conexion.asp"-->

<%'Permite auditar las acciones%>
<!--#include file="spl_auditoria.inc.asp"-->

<%
'se hace distinto asociar a sustancias que a ficheros; en el primer caso, ya tenemos el id, en el segundo, tenemos num_alternativa, tenemos que buscar el id

asociar=EliminaInyeccionSQL(request("asociar"))
idcheck=EliminaInyeccionSQL(request.form("idcheck"))

select case asociar


	case "fich_sustancia", "fich_grupo", "fich_sector", "fich_proceso", "fich_uso", "fich_residuo": 'ficheros a


		num_alternativa=request.form("num_alternativa")


		if num_alternativa="" then
			alerta="Debe escribir un número de alternativa."
		else
			if not isnumeric(num_alternativa) then
				alerta="El número de alternativa debe ser un número."
			else
				'buscamos id correspondiente al num_alternativa

				sqln="select id from dn_alter_ficheros where num_alternativa=" &num_alternativa

				set rstn=objconn1.execute(sqln)


				'response.write "<p>" &sqln

				rstnabierto=true

				if rstn.eof then
					alerta="No se encontró el número de alternativa. Compruebe que es correcto en la sección Ficheros."
				else
					id=rstn("id")
					if id="" then alerta="No se encontró el número de alternativa. Compruebe que es correcto en la sección Ficheros."
				end if
				'rstn queda abierto, para recorrerlo en el paso siguiente, YA QUE PUEDE HABER VARIOS FICHEROS DE ALTERNATIVAS CON EL MISMO NÚMERO, y lo cerramos al final
			end if
		end if

	case else: 'sustancias a


		id=request.form("id") 'id de elemento al que se asocia


end select

if alerta<>"" then
	flashMsgCreate alerta & "<br /><br />", "Advertencia"
else
	if idcheck="" then
		flashMsgCreate "No se había marcado <input type='checkbox' value='checked' /> ningún elemento para asociar. Marque los elementos que desea que se asocien. ", "Advertencia"
	else
		arr=split(idcheck, ",")
		FOR i=0 to UBound(arr)



			'Si estamos asociando una sustancia, tenemos un único id;

			'PERO si estamos asociando un fichero, podemos tener varios. Recorremos el recordset:

			if not rstnabierto then 'si estamos asociando una sustancia

				'el procedure, si ya existia la asociacion, no hace nada; else, la inserta
				set cmdsus=Server.CreateObject("ADODB.Command")
			   With cmdsus
				.ActiveConnection=objconn1
				.CommandText="dn_asociar"
				.CommandType=adCmdStoredProc
				.Parameters.Append  .CreateParameter("@asociar", advarchar, adParamInput, 15, asociar)
				.Parameters.Append  .CreateParameter("@id", adinteger, adParamInput, , id)
				.Parameters.Append  .CreateParameter("@idcheck", adinteger, adParamInput, , arr(i))
				if asociar="uso" or asociar="uso_gr" then
					toxico=0
					if request("toxico")=1 then toxico=1
					.Parameters.Append  .CreateParameter("@toxico", adboolean, adParamInput, , toxico)
				end if
				.Execute,,adexecutenorecords
			   End With
				set cmdsus=nothing

			else 'asociando ficheros, recorremos el recordset





			'response.write "<p>ASOCIAR= "&asociar&" - asociando sustancia " &arr(i)& " a fichero id:"

			'movemos el recordset al principio, y empezamos otra vez

			rstn.movefirst

			do while not rstn.eof

				id=rstn("id")



				'response.write "<p>" &id

				'el procedure, si ya existia la asociacion, no hace nada; else, la inserta
				set cmdsus=Server.CreateObject("ADODB.Command")
			   With cmdsus
				.ActiveConnection=objconn1
				.CommandText="dn_asociar"
				.CommandType=adCmdStoredProc
				.Parameters.Append  .CreateParameter("@asociar", advarchar, adParamInput, 15, asociar)
				.Parameters.Append  .CreateParameter("@id", adinteger, adParamInput, , id)
				.Parameters.Append  .CreateParameter("@idcheck", adinteger, adParamInput, , arr(i))
				if asociar="uso" or asociar="uso_gr" then
					toxico=0
					if request("toxico")=1 then toxico=1
					.Parameters.Append  .CreateParameter("@toxico", adboolean, adParamInput, , toxico)
				end if
				.Execute,,adexecutenorecords
			   End With
				set cmdsus=nothing

			 rstn.movenext

			 loop

			end if
		NEXT
		flashMsgCreate "Los elementos se han asociado", "OK"

		' ** AUDITORIA **
		spl_descripcion = "ids sustancias="&idcheck&"<br>id de asociado="&id&"<br>num alternativa="&num_alternativa
		call auditaYCierraConexion("asociar","sustancia a " & asociar,spl_descripcion)
	end if
end if



'if rstnabierto then

'	rstn.close
'	set rstn=nothing

'end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Istas</title>
<link rel="stylesheet" type="text/css" href="dn_estilos.css">
<script type="text/javascript" src="niftycube.js"></script>
<script type="text/javascript">
window.onload=function(){
Nifty("div#box2","big");
}

</script>
</head>

<body>
<%flashMsgShow()%>
<div align="center"><input type="button" value="Cerrar ventana" onClick="window.close()" /></div>
</body>
</html>
