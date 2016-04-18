<!--#include file="adovbs.inc"--><!--#include file="dn_conexion.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd"><html><head><meta http-equiv="Content-Type" content="text/html; charset=windows-1252"><title>Istas</title><link rel="stylesheet" type="text/css" href="dn_estilos.css"><link rel="stylesheet" type="text/css" href="dn_estilosmenu.css"><script type="text/javascript" src="niftycube.js"></script><script type="text/javascript">window.onload=function(){Nifty("ul#split h3","top");Nifty("ul#split div","bottom same-height");}</script></head><body><!--#include file="dn_menu.asp"--><h1>Importador de ICSC</h1>
<p>Importando relación ICSC-CAS...</p>

<%
response.buffer=true

' Primero reseteamos el campo de todas
sqlActualiza = "UPDATE dn_risc_sustancias SET num_icsc='' WHERE num_icsc <> ''"
'response.write "<br />" & sqlActualiza
objConn1.execute(sqlActualiza),,adexecutenorecords

' A continuación procesamos lo enviado
actualizadas = 0
no_encontradas = 0

listado = request.form("listado")
filas = split(listado, vbCrLf)
for i=0 to ubound(filas)
  ' Para la fila actual, procesamos los caracteres.
  ' Lo primero será el número ICSC, lo segundo será el CAS, y entre medias
  ' habrá espacios o tabuladores, que ignoraremos y marcarán el comienzo del CAS
  icsc = ""
  cas = ""
  comienza_cas = false
  fila = trim(filas(i))
  for j=1 to len(fila)
    ' Cogemos caracter actual
    caracter=mid(fila, j, 1)
    if (caracter = " " or caracter = "\t" or caracter = "/t" or caracter="	" or caracter="	") then
      comienza_cas = true
    else
      if comienza_cas then
        cas = cas & caracter
      else
        icsc = icsc & caracter
      end if
    end if    
  next
  
  icsc = trim(icsc)
  cas  = trim(cas)

  if (cas <> "" and icsc <> "") then
    ' Primero buscamos para apuntar cuántas se actualizarán
    sqlBusca = "SELECT COUNT(*) AS num FROM dn_risc_sustancias WHERE num_cas = '"&cas&"'"
    set objRstBusca=objConn1.execute(sqlBusca)
    if (objRstBusca("num") = 0) then
      ' No encontradas
      no_encontradas = no_encontradas + 1
      response.write "<br />No se han encontrado las sustancias con CAS ("&cas&")<br />"
    else
      actualizadas = actualizadas + objRstBusca("num")

      ' Bucleamos para actualizar cada una por separado, porque si tienen ya un valor se lo concatenamos separando por @
      sqlBucleCas="SELECT id, num_icsc FROM dn_risc_sustancias WHERE num_cas = '"&cas&"'"
      set objRstBucleCas=objConn1.execute(sqlBucleCas)
      do while (not objRstBucleCas.eof)
        if (objRstBucleCas("num_icsc")="") then
          sqlActualiza = "UPDATE dn_risc_sustancias SET num_icsc='"&icsc&"' WHERE id="&objRstBucleCas("id")
        else
          sqlActualiza = "UPDATE dn_risc_sustancias SET num_icsc='"&objRstBucleCas("num_icsc")&"@"&icsc&"' WHERE id="&objRstBucleCas("id")
        end if
        'response.write "<br />"&sqlActualiza
        response.write ". "
        objConn1.execute(sqlActualiza),,adexecutenorecords
        
        objRstBucleCas.movenext
      loop
      objRstBucleCas.close()
      set objRstBucleCas=nothing

    end if
    objRstBusca.close()
    set objRstBusca=nothing

  else
    response.write "<br />No hay ICSC y CAS, no hacemos nada (fila: "&fila&")<br />"
  end if
  response.flush()
next
%>
<hr />
<strong><%= actualizadas %> sustancias actualizadas, <%= no_encontradas %> sustancias no encontradas.</strong>
</body>
</html>

<% cerrarconexion %>
