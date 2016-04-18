<!--#include file="../dn_funciones_comunes.asp"-->
<!--#include file="dn_fun_comunes.asp"-->
<!--#include file="dn_fun_texto.asp"-->
<!--#include file="dn_auten.inc"-->
<!--#include file="adovbs.inc"-->
<!--#include file="dn_conexion.asp"-->
<!--#include file="dn_conexion_access.asp"-->
<!--#include file="../lib/listas_dev.asp"-->

<%
response.buffer = true
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Istas</title>
<link rel="stylesheet" type="text/css" href="dn_estilos.css">
<link rel="stylesheet" type="text/css" href="dn_estilosmenu.css">
<script type="text/javascript" src="niftycube.js"></script>
<script type="text/javascript">
window.onload=function(){
Nifty("ul#split h3","top");
Nifty("ul#split div","bottom same-height");
}
</script>
</head>

<body>
<!--#include file="dn_menu.asp"-->
<h1>Exportando a Access...<img src="imagenes/spinner.gif" align="absmiddle" hspace="5" id="spinner"></h1>
<p>Por favor, espera... el proceso de exportación puede durar varias horas (unos 5 minutos cada 1.000 sustancias).
<br /><strong>No interrumpas el proceso</strong>.</p>
<p><strong>Inicio: <%= now %></strong></p>
<%
'log_this("Comenzamos exportación a Access")

resetear_access() '## desactivar para debug

bucle_sustancias(EliminaInyeccionSQL(request("listado")))
bucle_grupos()
bucle_companias()
bucle_enfermedades()
bucle_enfermedades_por_grupos()
bucle_sectores()
bucle_usos()

'log_this("Fin de la exportación a Access")
%>

<% cerrar_conexion_access %>
<% compactar_access %>
<% cerrarconexion %>

<script type="text/javascript">
document.getElementById("spinner").src="imagenes/emoticon_smile.png";
</script>
<p><strong>Fin: <%= now %></strong></p>
<h2>Exportación finalizada</h2>
<center><blink><strong><a href="estructuras/risctox.mdb">Descargar fichero Access</a></strong></blink></center>

</body>
</html>

<%
'##################################################
function h3(byval cadena)	' Reemplaza caracteres codificados en HTML por una alternativa que también se pueda almacenar en SQL, para deshacer
	' lo cambiado por la función h (usado en la exportación a Access)
	if (isNull(cadena)) then
		h3 = ""
	else
		cadena = Replace(cadena,"'","´") 	' Apóstrofe de verdad por acento
		cadena = Replace(cadena,"&#39;","´") 	' Apóstrofe por acento
		cadena = Replace(cadena,"&#34;","´´") ' La comilla doble por dos acentos
		cadena = Replace(cadena,"&#37;","%") 	' Porcentaje
		cadena = Replace(cadena,"&#91;","[")	' Corchete izq
		cadena = Replace(cadena,"&#93;","]")	' Corchete dch
    cadena = Replace(cadena,"&mu;", "mu")
    cadena = Replace(cadena,"&alpha;", "alfa")
    cadena = Replace(cadena,"&beta;", "beta")
    cadena = Replace(cadena,"&plusmn;", "+-")
    cadena = Replace(cadena,"&xi;", "xi")
    cadena = Replace(cadena,"&ccedil;", "ç")
    cadena = Replace(cadena,"&epsilon;", "epsilon")
    cadena = Replace(cadena,"&eta;", "eta")
    cadena = Replace(cadena,"&kappa;", "kappa")
    cadena = Replace(cadena,"&lambda;", "lambda")
    cadena = Replace(cadena,"&omega;", "omega")
    cadena = Replace(cadena,"&gamma;", "gamma")
    cadena = Replace(cadena,"&delta;", "delta")
		cadena = Replace(cadena,"&#947;", "gamma")
		cadena = Replace(cadena,"&#945;", "alfa")
		cadena = Replace(cadena,"&#946;", "beta")
		cadena = Replace(cadena,"&#948;", "delta")
		cadena = Replace(cadena,"&#949;", "epsilon")
		cadena = Replace(cadena,"&#950;", "zeta")
		cadena = Replace(cadena,"&#8805;", ">=")
		cadena = Replace(cadena,"&gt;", ">")

		h3 = cadena
		end if
end function

' #####################################################

sub bucle_sustancias( byval listado )
  ' Seleccionamos las sustancias del SQL Server
  contador = 0
  sqlRix = get_sqlstr_listas( listado )
  'sqlRix = dame_sql_busqueda( listado )
  response.write "<!-- " & sqlRix & " -->"
  
	'response.write sqlRix '## activar para Debug
	'response.end() '## breakpoint
	
  response.write "<h2>Copiando sustancias...</h2>"

  set objRstRix = Server.CreateObject("ADODB.recordset")
	objRstRix.CursorLocation = adUseClient
	objRstRix.CursorType = adOpenForwardOnly
	objRstRix.LockType = adLockReadOnly
	objRstRix.Open sqlRix, objConn1

  do while (not objRstRix.eof)
    if ( objRstRix("id") <> "" ) then copia_sustancia( objRstRix( "id" ) )

    ' Mostramos contador
    contador = contador +1
		
    if ( (contador mod 100) = 0 ) then
      response.write "<br /><strong>"&now&"</strong>: "&contador&" sustancias...<br/>"
    else
      response.write ". "
    end if
		
    response.flush()

    ' Siguiente fila en Rix
    objRstRix.movenext
  loop

  objRstRix.close()
  set objRstRix = nothing

  response.write "<br />"&contador&" sustancias..."
end sub

' #####################################################

sub bucle_grupos()
  ' A ejecutar después del bucle de sustancias
  ' Recorre la tabla sustancias_por_grupos, distinct id_grupo, copiando los datos de los grupos
  contador = 0
  response.write "<h2>Copiando grupos...</h2>"
  sqlGrupos = "SELECT DISTINCT (id_grupo) FROM dn_risc_sustancias_por_grupos"
  set objRstGrupos=objConnAccess.execute(sqlGrupos)
  do while (not objRstGrupos.eof)
    copia_fila "dn_risc_grupos","id",objRstGrupos("id_grupo")
    copia_fila "dn_risc_grupos_por_enfermedades","id_grupo",objRstGrupos("id_grupo")
    copia_fila "dn_risc_grupos_por_usos","id_grupo",objRstGrupos("id_grupo")

    ' Mostramos contador
    contador = contador +1
    if ((contador mod 100) = 0) then
      response.write "<br />"&contador&" grupos...<br />"
    else
      response.write ". "
    end if
    response.flush()

    objRstGrupos.movenext
  loop

  objRstGrupos.close()
  set objRstGrupos = nothing

  response.write "<br />"&contador&" grupos..."
end sub

' #####################################################

sub bucle_companias()
  ' A ejecutar después del bucle de sustancias
  ' Recorre la tabla sustancias_por_companias, distinct id_compania, copiando los datos de las compañías
  contador = 0
  response.write "<h2>Copiando compañías...</h2>"
  sqlCompanias = "SELECT DISTINCT (id_compania) FROM dn_risc_sustancias_por_companias"
  set objRstCompanias=objConnAccess.execute(sqlCompanias)
  do while (not objRstCompanias.eof)
    copia_fila "dn_risc_companias","id",objRstCompanias("id_compania")

    ' Mostramos contador
    contador = contador +1
    if ((contador mod 100) = 0) then
      response.write "<br />"&contador&" compañías...<br />"
    else
      response.write ". "
    end if
    response.flush()

    objRstCompanias.movenext
  loop

  objRstCompanias.close()
  set objRstCompanias = nothing

  response.write "<br />"&contador&" compañías..."
end sub

' #####################################################

sub bucle_enfermedades()
  ' A ejecutar después del bucle de sustancias
  ' Recorre la tabla sustancias_por_enfermedades, distinct id_enfermedad, copiando los datos de las enfermedades
  contador = 0
  response.write "<h2>Copiando enfermedades asociadas a sustancias...</h2>"
  sqlEnfermedades = "SELECT DISTINCT (id_enfermedad) FROM dn_risc_sustancias_por_enfermedades"
  set objRstEnfermedades=objConnAccess.execute(sqlEnfermedades)
  do while (not objRstEnfermedades.eof)
    copia_fila "dn_risc_enfermedades","id",objRstEnfermedades("id_enfermedad")

    ' Mostramos contador
    contador = contador +1
    if ((contador mod 100) = 0) then
      response.write "<br />"&contador&" enfermedades...<br />"
    else
      response.write ". "
    end if
    response.flush()

    objRstEnfermedades.movenext
  loop

  objRstEnfermedades.close()
  set objRstEnfermedades = nothing

  response.write "<br />"&contador&" enfermedades..."

end sub

' #####################################################

sub bucle_enfermedades_por_grupos()
  ' A ejecutar después del bucle de grupos
  ' Recorre la tabla grupos_por_enfermedades, distinct id_enfermedad, copiando los datos de las enfermedades
  contador = 0
  response.write "<h2>Copiando enfermedades asociadas a grupos...</h2>"
  sqlEnfermedades = "SELECT DISTINCT (id_enfermedad) FROM dn_risc_grupos_por_enfermedades"
  set objRstEnfermedades=objConnAccess.execute(sqlEnfermedades)
  do while (not objRstEnfermedades.eof)
    ' Insertamos si no existe ya esa enfermedad (puede venir de antes, de cuando las sustancias)
    sqlBusca = "SELECT id FROM dn_risc_enfermedades WHERE id="&objRstEnfermedades("id_enfermedad")
    set objRstBusca = objConnAccess.execute(sqlBusca)
    if objRstBusca.eof then
      ' No existe, la copiamos
      copia_fila "dn_risc_enfermedades","id",objRstEnfermedades("id_enfermedad")

      ' Mostramos contador
      contador = contador +1
      if ((contador mod 100) = 0) then
        response.write "<br />"&contador&" enfermedades...<br />"
      else
        response.write ". "
      end if
    else
      ' Ya existe, no la copiamos de nuevo
      response.write "- "
    end if

    objRstBusca.close()
    set objRstBusca=nothing

    response.flush()

    objRstEnfermedades.movenext
  loop

  objRstEnfermedades.close()
  set objRstEnfermedades = nothing

  response.write "<br />"&contador&" enfermedades..."

end sub

' #####################################################

sub bucle_usos()
  ' A ejecutar después del bucle de sustancias
  ' Recorre la tabla sustancias_por_usos, distinct id_uso, copiando los datos de los usos
  contador = 0
  response.write "<h2>Copiando usos asociados a sustancias...</h2>"
  sqlUsos = "SELECT DISTINCT (id_uso) FROM dn_risc_sustancias_por_usos"
  set objRstUsos=objConnAccess.execute(sqlUsos)
  do while (not objRstUsos.eof)
    copia_fila "dn_risc_usos","id",objRstUsos("id_uso")

    ' Mostramos contador
    contador = contador +1
    if ((contador mod 100) = 0) then
      response.write "<br />"&contador&" usos...<br />"
    else
      response.write ". "
    end if
    response.flush()

    objRstUsos.movenext
  loop

  objRstUsos.close()
  set objRstUsos = nothing

  response.write "<br />"&contador&" usos..."

end sub

' #####################################################

sub bucle_usos_por_grupos()
  ' A ejecutar después del bucle de grupos
  ' Recorre la tabla grupos_por_usos, distinct id_uso, copiando los datos de los usos
  contador = 0
  response.write "<h2>Copiando usos asociados a grupos...</h2>"
  sqlEnfermedades = "SELECT DISTINCT (id_uso) FROM dn_risc_grupos_por_usos"
  set objRstUsos=objConnAccess.execute(sqlUsos)
  do while (not objRstUsos.eof)
    ' Insertamos si no existe ya ese uso (puede venir de antes, de cuando las sustancias)
    sqlBusca = "SELECT id FROM dn_risc_usos WHERE id="&objRstUsos("id_uso")
    set objRstBusca = objConnAccess.execute(sqlBusca)
    if objRstBusca.eof then
      ' No existe, la copiamos
      copia_fila "dn_risc_usos","id",objRstUsos("id_uso")

      ' Mostramos contador
      contador = contador +1
      if ((contador mod 100) = 0) then
        response.write "<br />"&contador&" usos...<br />"
      else
        response.write ". "
      end if
    else
      ' Ya existe, no la copiamos de nuevo
      response.write "- "
    end if

    objRstBusca.close()
    set objRstBusca=nothing

    response.flush()

    objRstUsos.movenext
  loop

  objRstUsos.close()
  set objRstUsos = nothing

  response.write "<br />"&contador&" usos..."

end sub

' #####################################################

sub bucle_sectores()
  ' A ejecutar después del bucle de sustancias
  ' Recorre la tabla sustancias_por_sectores, distinct id_sector, copiando los datos de los sectores
  contador = 0
  response.write "<h2>Copiando sectores...</h2>"
  sqlSectores = "SELECT DISTINCT (id_sector) FROM dn_risc_sustancias_por_sectores"
  set objRstSectores=objConnAccess.execute(sqlSectores)
  do while (not objRstSectores.eof)
    copia_fila "dn_alter_sectores","id",objRstSectores("id_sector")

    ' Mostramos contador
    contador = contador +1
    if ((contador mod 100) = 0) then
      response.write "<br />"&contador&" sectores...<br />"
    else
      response.write ". "
    end if
    response.flush()

    objRstSectores.movenext
  loop

  objRstSectores.close()
  set objRstSectores = nothing

  response.write "<br />"&contador&" sectores..."

end sub

' #####################################################

sub copia_sustancia(byval id)
  ' Copia los datos de la sustancia del SQL Server al Access, para todas las tablas de sustancias
  ' en las que aparece:

  copia_fila "dn_risc_sustancias","id",id
  copia_fila "dn_risc_sustancias_ambiente","id_sustancia",id
  copia_fila "dn_risc_sustancias_cancer_otras","id_sustancia",id
  copia_fila "dn_risc_sustancias_mama_cop","id_sustancia",id
  copia_fila "dn_risc_sustancias_iarc","id_sustancia",id
  copia_fila "dn_risc_sustancias_neuro_disruptor","id_sustancia",id
  copia_fila "dn_risc_sustancias_vl","id_sustancia",id
  copia_fila "dn_risc_sustancias_salud","id_sustancia",id

  copia_fila "dn_risc_sustancias_por_companias","id_sustancia",id
  copia_fila "dn_risc_sustancias_por_enfermedades","id_sustancia",id
  copia_fila "dn_risc_sustancias_por_grupos","id_sustancia",id
  copia_fila "dn_risc_sustancias_por_usos","id_sustancia",id
  copia_fila "dn_risc_sustancias_por_sectores","id_sustancia",id
  copia_fila "dn_risc_nombres_comerciales","id_sustancia",id
  copia_fila "dn_risc_sinonimos","id_sustancia",id
end sub

' #####################################################

sub copia_fila(byval tabla, byval columna_id, byval id)
  ' Copia la fila de la tabla indicada, indicando tambien la columna y valor del id

  ' Sacamos los datos
  sqlFila = "SELECT * FROM " & tabla & " WHERE " & columna_id & "=" & id
  'response.write "<br />"&sqlFila&"<br />"

  set objRstFila = Server.CreateObject("ADODB.recordset")
	objRstFila.CursorLocation = adUseClient
	objRstFila.CursorType = adOpenForwardOnly
	objRstFila.LockType = adLockReadOnly
	objRstFila.Open sqlFila, objConn1

  if (not objRstFila.eof) then
    ' Filas encontradas, bucleamos
    do while (not objRstFila.eof)
      ' Sacamos los nombres de las columnas y generamos la cadena del INSERT
      campos = ""
      valores = ""
      for each columna in objRstFila.fields
        if (campos = "") then
          campos = columna.name
          valores = "'"&objRstFila(columna.name)&"'"
        else
          campos = campos & ", " & columna.name
          valores = valores & ", " & "'" & h3(objRstFila(columna.name)) & "'"
        end if
      next

      sqlAccess = "INSERT INTO "&tabla&" ("&campos&") VALUES ("&valores&")"
			'response.write(sqlAccess) & "<br>"
			' response.end
'      response.write "<!--<br />"&sqlAccess&"-->"
      objConnAccess.execute(sqlAccess),,adexecutenorecords

      objRstFila.movenext
    loop
  end if

  objRstFila.close()
  set objRstFila = nothing

  ' Pausa para no saturar al servidor
  'contador = 0
  'for i = 1 to 100000
  '  contador = contador + 1
  'next
  ' Fin de pausa
end sub

' #####################################################
sub resetear_access

  ' Eliminamos todos los datos del access

  borra_tabla_access("dn_alter_sectores")
  borra_tabla_access("dn_risc_companias")
  borra_tabla_access("dn_risc_enfermedades")
  borra_tabla_access("dn_risc_frases_r")
  borra_tabla_access("dn_risc_frases_s")
  borra_tabla_access("dn_risc_grupos")
  borra_tabla_access("dn_risc_grupos_por_enfermedades")
  borra_tabla_access("dn_risc_nombres_comerciales")
  borra_tabla_access("dn_risc_sinonimos")
  borra_tabla_access("dn_risc_sustancias")
  borra_tabla_access("dn_risc_sustancias_ambiente")
  borra_tabla_access("dn_risc_sustancias_cancer_otras")
  borra_tabla_access("dn_risc_sustancias_mama_cop")
  borra_tabla_access("dn_risc_sustancias_iarc")
  borra_tabla_access("dn_risc_sustancias_neuro_disruptor")
  borra_tabla_access("dn_risc_sustancias_salud")
  borra_tabla_access("dn_risc_sustancias_por_companias")
  borra_tabla_access("dn_risc_sustancias_por_enfermedades")
  borra_tabla_access("dn_risc_sustancias_por_grupos")
  borra_tabla_access("dn_risc_sustancias_por_sectores")
  borra_tabla_access("dn_risc_sustancias_vl")
  borra_tabla_access("dn_simbolos")
  borra_tabla_access("dn_risc_usos")
  borra_tabla_access("dn_risc_sustancias_por_usos")
  borra_tabla_access("dn_risc_grupos_por_usos")

end sub

' #####################################################

sub borra_tabla_access(byval nombre_tabla)

  sql="DELETE * FROM "&nombre_tabla
  objConnAccess.execute(sql),,adexecutenorecords

end sub

' #####################################################

sub compactar_access

  'log_this("Compactando Access...")

  ruta_access = server.mappath(".")&"\estructuras\risctox.mdb"
  ruta_compactado = server.mappath(".")&"\estructuras\risctox_compactado.mdb"

  ' Creamos copia compactada
  Set Engine = CreateObject("JRO.JetEngine")
  Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&ruta_access,"Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&ruta_compactado

  ' Borramos el anterior y renombramos
  Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
  objFSO.DeleteFile ruta_access
  objFSO.MoveFile ruta_compactado, ruta_access

  set objFSO = nothing

  'log_this("Se terminó de compactar el Access...")
end sub

' #####################################################

function dame_sql_busqueda(byval listado) '## Obsoleta!
  
	' Devuelve la cadena SQL para buscar las sustancias dependiendo del listado
  ' sql = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus " &_
  sql = "select distinct sus.id from dn_risc_sustancias as sus " &_
				" FULL OUTER JOIN dn_risc_sinonimos as sin ON (sus.id=sin.id_sustancia) " &_
				" FULL OUTER JOIN dn_risc_nombres_comerciales as com ON (sus.id=com.id_sustancia) "

  sql = sql & get_string_tablas( listado )

  ' AÑADIMOS LAS CONDICIONES DE LOS GRUPOS
	'según filtro, agregamos distintas condiciones
	if filtro<>"0" then
		if InStr(sql, "WHERE")>0 then
			sql = sql & " AND "
		else
			sql = sql & " WHERE "
		end if
		
		sql = sql & get_string_codicion( listado )
	
	end if

  dame_sql_busqueda = sql
end function

%>
