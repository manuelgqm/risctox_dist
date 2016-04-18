
<!--#include file="../dn_funciones_comunes.asp"-->
<!--#include file="dn_fun_comunes.asp"-->
<!--#include file="dn_fun_texto.asp"-->
<!--#include file="dn_auten.inc"-->
<!--#include file="adovbs.inc"-->
<!--#include file="dn_conexion.asp"-->
<!--#include file="dn_conexion_access.asp"-->

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

resetear_access()

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
	'Sergio
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

sub bucle_sustancias(byval listado)
  ' Seleccionamos las sustancias del SQL Server
  contador = 0
  sqlRix = dame_sql_busqueda(listado)
  response.write "<!-- "&sqlRix&" -->"
  response.write "<h2>Copiando sustancias...</h2>"

  set objRstRix = Server.CreateObject("ADODB.recordset")
	objRstRix.CursorLocation = adUseClient
	objRstRix.CursorType = adOpenForwardOnly
	objRstRix.LockType = adLockReadOnly
	objRstRix.Open sqlRix, objConn1
  'set objRstRix = objConn1.execute(sqlRix)

  do while (not objRstRix.eof)
    if (objRstRix("id")<>"") then copia_sustancia objRstRix("id")

    ' Mostramos contador
    contador = contador +1
    if ((contador mod 100) = 0) then
      response.write "<br /><strong>"&now&"</strong>: "&contador&" sustancias...<br/>"
    else
      response.write ". "
    end if
    response.flush()

    ' Si es la lista negra, aprovechamos para marcar este campo en la base
    'if (request("listado")="negra") then
    '  sqlNegra="UPDATE dn_risc_sustancias SET negra=1 WHERE id="&objRstRix("id")
    '  objConn1.execute(sqlNegra),,adexecutenorecords
    'end if

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
  sqlFila = "SELECT * FROM "&tabla&" WHERE "&columna_id&"="&id
  'response.write "<br />"&sqlFila&"<br />"

  set objRstFila = Server.CreateObject("ADODB.recordset")
	objRstFila.CursorLocation = adUseClient
	objRstFila.CursorType = adOpenForwardOnly
	objRstFila.LockType = adLockReadOnly
	objRstFila.Open sqlFila, objConn1
  'set objRstFila=objConn1.execute(sqlFila)

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

function dame_sql_busqueda(byval listado)
  ' Devuelve la cadena SQL para buscar las sustancias dependiendo del listado
  sql = ""
  sql = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus "
  sql=sql & " FULL OUTER JOIN dn_risc_sinonimos as sin ON (sus.id=sin.id_sustancia) "
  sql=sql & " FULL OUTER JOIN dn_risc_nombres_comerciales as com ON (sus.id=com.id_sustancia) "

  select case listado
    case "todas":
      sql = "select distinct sus.id from dn_risc_sustancias as sus"
	  sql = sql & " LEFT OUTER JOIN dn_risc_sustancias_por_usos ON (sus.id=dn_risc_sustancias_por_usos.id_sustancia)"
	  sql = sql & " LEFT OUTER JOIN dn_alter_ficheros_por_sustancias ON (sus.id=dn_alter_ficheros_por_sustancias.id_sustancia)"
	  sql = sql & " LEFT OUTER JOIN dn_risc_sustancias_por_grupos ON (sus.id=dn_risc_sustancias_por_grupos.id_sustancia)"
	  sql = sql & " LEFT OUTER JOIN dn_risc_grupos_por_usos ON (dn_risc_sustancias_por_grupos.id_grupo=dn_risc_grupos_por_usos.id_grupo)"
	  sql = sql & " LEFT OUTER JOIN dn_alter_ficheros_por_grupos ON (dn_risc_sustancias_por_grupos.id_grupo=dn_alter_ficheros_por_grupos.id_grupo)"

	case "cym": 'cancerigenos y mutagenos segun RD 363/1995
	'no unimos a mas tablas

	case "cym2": 'cancerigenos y mutagenos segun IARC
	sql = sql & " LEFT OUTER JOIN dn_risc_sustancias_iarc ON (sus.id=dn_risc_sustancias_iarc.id_sustancia)"

	case "cym3": 'cancerigenos y mutagenos segun otras
	sql  = sql  & " LEFT OUTER JOIN dn_risc_sustancias_cancer_otras ON (sus.id=dn_risc_sustancias_cancer_otras.id_sustancia)"

	case "mama": 'cáncer de mama
	sql = sql & " LEFT OUTER JOIN dn_risc_sustancias_mama_cop ON (sus.id=dn_risc_sustancias_mama_cop.id_sustancia)"

	case "cop": 'cop
	sql = sql  & " LEFT OUTER JOIN dn_risc_sustancias_mama_cop ON (sus.id=dn_risc_sustancias_mama_cop.id_sustancia)"

    case "tpr":
      'sql = "select distinct sus.id from dn_risc_sustancias as sus FULL OUTER JOIN dn_risc_sinonimos as sin ON (sus.id=sin.id_sustancia) LEFT OUTER JOIN dn_risc_sustancias_iarc iarc ON sus.id = iarc.id_sustancia LEFT OUTER JOIN dn_risc_sustancias_cancer_otras otras ON sus.id = otras.id_sustancia LEFT OUTER JOIN dn_risc_sustancias_neuro_disruptor neuro ON sus.id = neuro.id_sustancia LEFT OUTER JOIN dn_risc_sustancias_ambiente ambiente ON sus.id = ambiente.id_sustancia WHERE ((sus.clasificacion_1 LIKE '%R60') OR (sus.clasificacion_1 LIKE '%R60;%') OR (sus.clasificacion_1 LIKE '%R61') OR (sus.clasificacion_1 LIKE '%R61;%') OR (sus.clasificacion_1 LIKE '%R62') OR (sus.clasificacion_1 LIKE '%R62;%') OR (sus.clasificacion_1 LIKE '%R63') OR (sus.clasificacion_1 LIKE '%R63;%') OR (sus.clasificacion_2 LIKE '%R60') OR (sus.clasificacion_2 LIKE '%R60;%') OR (sus.clasificacion_2 LIKE '%R61') OR (sus.clasificacion_2 LIKE '%R61;%') OR (sus.clasificacion_2 LIKE '%R62') OR (sus.clasificacion_2 LIKE '%R62;%') OR (sus.clasificacion_2 LIKE '%R63') OR (sus.clasificacion_2 LIKE '%R63;%') OR (sus.clasificacion_3 LIKE '%R60') OR (sus.clasificacion_3 LIKE '%R60;%') OR (sus.clasificacion_3 LIKE '%R61') OR (sus.clasificacion_3 LIKE '%R61;%') OR (sus.clasificacion_3 LIKE '%R62') OR (sus.clasificacion_3 LIKE '%R62;%') OR (sus.clasificacion_3 LIKE '%R63') OR (sus.clasificacion_3 LIKE '%R63;%') OR (sus.clasificacion_4 LIKE '%R60') OR (sus.clasificacion_4 LIKE '%R60;%') OR (sus.clasificacion_4 LIKE '%R61') OR (sus.clasificacion_4 LIKE '%R61;%') OR (sus.clasificacion_4 LIKE '%R62') OR (sus.clasificacion_4 LIKE '%R62;%') OR (sus.clasificacion_4 LIKE '%R63') OR (sus.clasificacion_4 LIKE '%R63;%') OR (sus.clasificacion_5 LIKE '%R60') OR (sus.clasificacion_5 LIKE '%R60;%') OR (sus.clasificacion_5 LIKE '%R61') OR (sus.clasificacion_5 LIKE '%R61;%') OR (sus.clasificacion_5 LIKE '%R62') OR (sus.clasificacion_5 LIKE '%R62;%') OR (sus.clasificacion_5 LIKE '%R63') OR (sus.clasificacion_5 LIKE '%R63;%') OR (sus.clasificacion_6 LIKE '%R60') OR (sus.clasificacion_6 LIKE '%R60;%') OR (sus.clasificacion_6 LIKE '%R61') OR (sus.clasificacion_6 LIKE '%R61;%') OR (sus.clasificacion_6 LIKE '%R62') OR (sus.clasificacion_6 LIKE '%R62;%') OR (sus.clasificacion_6 LIKE '%R63') OR (sus.clasificacion_6 LIKE '%R63;%') OR (sus.clasificacion_7 LIKE '%R60') OR (sus.clasificacion_7 LIKE '%R60;%') OR (sus.clasificacion_7 LIKE '%R61') OR (sus.clasificacion_7 LIKE '%R61;%') OR (sus.clasificacion_7 LIKE '%R62') OR (sus.clasificacion_7 LIKE '%R62;%') OR (sus.clasificacion_7 LIKE '%R63') OR (sus.clasificacion_7 LIKE '%R63;%') OR (sus.clasificacion_8 LIKE '%R60') OR (sus.clasificacion_8 LIKE '%R60;%') OR (sus.clasificacion_8 LIKE '%R61') OR (sus.clasificacion_8 LIKE '%R61;%') OR (sus.clasificacion_8 LIKE '%R62') OR (sus.clasificacion_8 LIKE '%R62;%') OR (sus.clasificacion_8 LIKE '%R63') OR (sus.clasificacion_8 LIKE '%R63;%') OR (sus.clasificacion_9 LIKE '%R60') OR (sus.clasificacion_9 LIKE '%R60;%') OR (sus.clasificacion_9 LIKE '%R61') OR (sus.clasificacion_9 LIKE '%R61;%') OR (sus.clasificacion_9 LIKE '%R62') OR (sus.clasificacion_9 LIKE '%R62;%') OR (sus.clasificacion_9 LIKE '%R63') OR (sus.clasificacion_9 LIKE '%R63;%') OR (sus.clasificacion_10 LIKE '%R60') OR (sus.clasificacion_10 LIKE '%R60;%') OR (sus.clasificacion_10 LIKE '%R61') OR (sus.clasificacion_10 LIKE '%R61;%') OR (sus.clasificacion_10 LIKE '%R62') OR (sus.clasificacion_10 LIKE '%R62;%') OR (sus.clasificacion_10 LIKE '%R63') OR (sus.clasificacion_10 LIKE '%R63;%') OR (sus.clasificacion_11 LIKE '%R60') OR (sus.clasificacion_11 LIKE '%R60;%') OR (sus.clasificacion_11 LIKE '%R61') OR (sus.clasificacion_11 LIKE '%R61;%') OR (sus.clasificacion_11 LIKE '%R62') OR (sus.clasificacion_11 LIKE '%R62;%') OR (sus.clasificacion_11 LIKE '%R63') OR (sus.clasificacion_11 LIKE '%R63;%') OR (sus.clasificacion_12 LIKE '%R60') OR (sus.clasificacion_12 LIKE '%R60;%') OR (sus.clasificacion_12 LIKE '%R61') OR (sus.clasificacion_12 LIKE '%R61;%') OR (sus.clasificacion_12 LIKE '%R62') OR (sus.clasificacion_12 LIKE '%R62;%') OR (sus.clasificacion_12 LIKE '%R63') OR (sus.clasificacion_12 LIKE '%R63;%') OR (sus.clasificacion_13 LIKE '%R60') OR (sus.clasificacion_13 LIKE '%R60;%') OR (sus.clasificacion_13 LIKE '%R61') OR (sus.clasificacion_13 LIKE '%R61;%') OR (sus.clasificacion_13 LIKE '%R62') OR (sus.clasificacion_13 LIKE '%R62;%') OR (sus.clasificacion_13 LIKE '%R63') OR (sus.clasificacion_13 LIKE '%R63;%') OR (sus.clasificacion_14 LIKE '%R60') OR (sus.clasificacion_14 LIKE '%R60;%') OR (sus.clasificacion_14 LIKE '%R61') OR (sus.clasificacion_14 LIKE '%R61;%') OR (sus.clasificacion_14 LIKE '%R62') OR (sus.clasificacion_14 LIKE '%R62;%') OR (sus.clasificacion_14 LIKE '%R63') OR (sus.clasificacion_14 LIKE '%R63;%') OR (sus.clasificacion_15 LIKE '%R60') OR (sus.clasificacion_15 LIKE '%R60;%') OR (sus.clasificacion_15 LIKE '%R61') OR (sus.clasificacion_15 LIKE '%R61;%') OR (sus.clasificacion_15 LIKE '%R62') OR (sus.clasificacion_15 LIKE '%R62;%') OR (sus.clasificacion_15 LIKE '%R63') OR (sus.clasificacion_15 LIKE '%R63;%'))"
	  sql = sql & " LEFT OUTER JOIN dn_risc_sustancias_iarc iarc ON sus.id = iarc.id_sustancia LEFT OUTER JOIN dn_risc_sustancias_cancer_otras otras ON sus.id = otras.id_sustancia LEFT OUTER JOIN dn_risc_sustancias_neuro_disruptor neuro ON sus.id = neuro.id_sustancia LEFT OUTER JOIN dn_risc_sustancias_ambiente ambiente ON sus.id = ambiente.id_sustancia "

    case "dis":
      sql=sql & " LEFT OUTER JOIN dn_risc_sustancias_neuro_disruptor ON (sus.id=dn_risc_sustancias_neuro_disruptor.id_sustancia) "


	case "neu": 'neurotoxico
	sql = sql  & " LEFT OUTER JOIN dn_risc_sustancias_neuro_disruptor ON (sus.id=dn_risc_sustancias_neuro_disruptor.id_sustancia)"

	'Sergio
	case "oto": 'neurotoxico
	sql = sql & " LEFT OUTER JOIN dn_risc_sustancias_neuro_disruptor ON (sus.id=dn_risc_sustancias_neuro_disruptor.id_sustancia)"

	case "sen": ' sensibilizante... no hace falta más tablas que la principal

	case "senreach": 'sensibilizantes reach
	sql = sql  & " LEFT OUTER JOIN dn_risc_sensibilizantes_reach ON (sus.id=dn_risc_sensibilizantes_reach.id_sustancia)"

	case "pyb": 'Sustancias tóxicas, persistentes y bioacumulativas
	sql = sql & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"

	case "tac": 'Sustancias de toxicidad acuática según Directiva de aguas
	sql = sql  & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"

	case "tac2": 'Sustancias peligrosas agua Alemania
	sql = sql & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"

	case "dat": 'Sustancias de daño a la capa de ozono
	sql = sql & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"

	case "dat2": 'Sustancias cambio climático
	sql = sql & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"

	case "dat3": 'Sustancias calidad aire
	sql = sql & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"

	case "vl1": 'Límites de exposición profesional: Valores Límite Ambientales
	sql = sql & " LEFT OUTER JOIN dn_risc_sustancias_vl ON (sus.id=dn_risc_sustancias_vl.id_sustancia)"

	case "vl2": '	Valores Límite Ambientales Cancerígenos
	sql = sql & " LEFT OUTER JOIN dn_risc_sustancias_vl as vl ON (sus.id=vl.id_sustancia)"

	case "vl3": 'Límites de exposición profesional: Valores Límite Biológicos
	sql = sql & " LEFT OUTER JOIN dn_risc_sustancias_vl ON (sus.id=dn_risc_sustancias_vl.id_sustancia)"

	case "enf": 'Enfermedades profesionales (borrador)
	sql = sql & " INNER JOIN dn_risc_sustancias_por_grupos ON (sus.id=dn_risc_sustancias_por_grupos.id_sustancia) INNER JOIN dn_risc_grupos_por_enfermedades ON (dn_risc_sustancias_por_grupos.id_grupo=dn_risc_grupos_por_enfermedades.id_grupo)"

	case "res": 'residuos
		sql = sql & " FULL OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"
		sql = sql & " FULL OUTER JOIN dn_risc_sustancias_cancer_otras ON (sus.id=dn_risc_sustancias_cancer_otras.id_sustancia)"
		sql = sql & " FULL OUTER JOIN dn_risc_sustancias_iarc ON (sus.id=dn_risc_sustancias_iarc.id_sustancia)"
		sql = sql & " FULL OUTER JOIN dn_risc_sustancias_neuro_disruptor ON (sus.id=dn_risc_sustancias_neuro_disruptor.id_sustancia)"
		sql = sql & " FULL OUTER JOIN dn_risc_sustancias_vl ON (sus.id=dn_risc_sustancias_vl.id_sustancia)"

	case "ver": 'vertidos
		sql = sql & " LEFT OUTER JOIN dn_risc_sustancias_iarc iarc ON sus.id = iarc.id_sustancia LEFT OUTER JOIN dn_risc_sustancias_cancer_otras otras ON sus.id = otras.id_sustancia LEFT OUTER JOIN dn_risc_sustancias_neuro_disruptor neuro ON sus.id = neuro.id_sustancia LEFT OUTER JOIN dn_risc_sustancias_ambiente ambiente ON sus.id = ambiente.id_sustancia "

	case "emi": 'emisiones
		sql = sql & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"

	case "cov": 'Compuestos orgánicos volátiles (COV)
		sql = sql  & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"

	'Sergio
	case "lpc": 'Sustancias (LPCIC)
	'sqls=sqls & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"
		sql=sql & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"
	case "ep1":
		sql = sql & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"

	case "ep2":
		sql = sql & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"

	case "ep3":
		sql = sql & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"



	case "acm": 'Sustancias que pueden provocar accidentes graves
		sql = sql & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"



	case "cos":
		sql = sql & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"

	case "anexo_reach"
		sql = sql & " LEFT OUTER JOIN dn_risc_sustancias_por_usos ON (sus.id=dn_risc_sustancias_por_usos.id_sustancia)"

    case "todo":
      sql = "SELECT TOP 100 id FROM dn_risc_sustancias WHERE 1=1"

    case "negra":
' Se sustituye por la cadena de búsqueda del buscador.
'      sql = "SELECT distinct sus.id FROM dn_risc_sustancias AS sus WHERE (sus.id IN (select distinct sus.id from dn_risc_sustancias as sus WHERE ((sus.clasificacion_1 LIKE '%R40') OR (sus.clasificacion_1 LIKE '%R40;%') OR (sus.clasificacion_1 LIKE '%R40 %') OR (sus.clasificacion_1 LIKE '%R45') OR (sus.clasificacion_1 LIKE '%R45;%') OR (sus.clasificacion_1 LIKE '%R45 %') OR (sus.clasificacion_1 LIKE '%R49') OR (sus.clasificacion_1 LIKE '%R49;%') OR (sus.clasificacion_1 LIKE '%R49 %') OR (sus.clasificacion_1 LIKE '%R40/20') OR (sus.clasificacion_1 LIKE '%R40/20;%') OR (sus.clasificacion_1 LIKE '%R40/20 %') OR (sus.clasificacion_1 LIKE '%R40/21') OR (sus.clasificacion_1 LIKE '%R40/21;%') OR (sus.clasificacion_1 LIKE '%R40/21 %') OR (sus.clasificacion_1 LIKE '%R40/22') OR (sus.clasificacion_1 LIKE '%R40/22;%') OR (sus.clasificacion_1 LIKE '%R40/22 %') OR (sus.clasificacion_1 LIKE '%R40/20/21') OR (sus.clasificacion_1 LIKE '%R40/20/21;%') OR (sus.clasificacion_1 LIKE '%R40/20/21 %') OR (sus.clasificacion_1 LIKE '%R40/20/22') OR (sus.clasificacion_1 LIKE '%R40/20/22;%') OR (sus.clasificacion_1 LIKE '%R40/20/22 %') OR (sus.clasificacion_1 LIKE '%R40/21/22') OR (sus.clasificacion_1 LIKE '%R40/21/22;%') OR (sus.clasificacion_1 LIKE '%R40/21/22 %') OR (sus.clasificacion_1 LIKE '%R40/20/21/22') OR (sus.clasificacion_1 LIKE '%R40/20/21/22;%') OR (sus.clasificacion_1 LIKE '%R40/20/21/22 %') OR (sus.clasificacion_2 LIKE '%R40') OR (sus.clasificacion_2 LIKE '%R40;%') OR (sus.clasificacion_2 LIKE '%R40 %') OR (sus.clasificacion_2 LIKE '%R45') OR (sus.clasificacion_2 LIKE '%R45;%') OR (sus.clasificacion_2 LIKE '%R45 %') OR (sus.clasificacion_2 LIKE '%R49') OR (sus.clasificacion_2 LIKE '%R49;%') OR (sus.clasificacion_2 LIKE '%R49 %') OR (sus.clasificacion_2 LIKE '%R40/20') OR (sus.clasificacion_2 LIKE '%R40/20;%') OR (sus.clasificacion_2 LIKE '%R40/20 %') OR (sus.clasificacion_2 LIKE '%R40/21') OR (sus.clasificacion_2 LIKE '%R40/21;%') OR (sus.clasificacion_2 LIKE '%R40/21 %') OR (sus.clasificacion_2 LIKE '%R40/22') OR (sus.clasificacion_2 LIKE '%R40/22;%') OR (sus.clasificacion_2 LIKE '%R40/22 %') OR (sus.clasificacion_2 LIKE '%R40/20/21') OR (sus.clasificacion_2 LIKE '%R40/20/21;%') OR (sus.clasificacion_2 LIKE '%R40/20/21 %') OR (sus.clasificacion_2 LIKE '%R40/20/22') OR (sus.clasificacion_2 LIKE '%R40/20/22;%') OR (sus.clasificacion_2 LIKE '%R40/20/22 %') OR (sus.clasificacion_2 LIKE '%R40/21/22') OR (sus.clasificacion_2 LIKE '%R40/21/22;%') OR (sus.clasificacion_2 LIKE '%R40/21/22 %') OR (sus.clasificacion_2 LIKE '%R40/20/21/22') OR (sus.clasificacion_2 LIKE '%R40/20/21/22;%') OR (sus.clasificacion_2 LIKE '%R40/20/21/22 %') OR (sus.clasificacion_3 LIKE '%R40') OR (sus.clasificacion_3 LIKE '%R40;%') OR (sus.clasificacion_3 LIKE '%R40 %') OR (sus.clasificacion_3 LIKE '%R45') OR (sus.clasificacion_3 LIKE '%R45;%') OR (sus.clasificacion_3 LIKE '%R45 %') OR (sus.clasificacion_3 LIKE '%R49') OR (sus.clasificacion_3 LIKE '%R49;%') OR (sus.clasificacion_3 LIKE '%R49 %') OR (sus.clasificacion_3 LIKE '%R40/20') OR (sus.clasificacion_3 LIKE '%R40/20;%') OR (sus.clasificacion_3 LIKE '%R40/20 %') OR (sus.clasificacion_3 LIKE '%R40/21') OR (sus.clasificacion_3 LIKE '%R40/21;%') OR (sus.clasificacion_3 LIKE '%R40/21 %') OR (sus.clasificacion_3 LIKE '%R40/22') OR (sus.clasificacion_3 LIKE '%R40/22;%') OR (sus.clasificacion_3 LIKE '%R40/22 %') OR (sus.clasificacion_3 LIKE '%R40/20/21') OR (sus.clasificacion_3 LIKE '%R40/20/21;%') OR (sus.clasificacion_3 LIKE '%R40/20/21 %') OR (sus.clasificacion_3 LIKE '%R40/20/22') OR (sus.clasificacion_3 LIKE '%R40/20/22;%') OR (sus.clasificacion_3 LIKE '%R40/20/22 %') OR (sus.clasificacion_3 LIKE '%R40/21/22') OR (sus.clasificacion_3 LIKE '%R40/21/22;%') OR (sus.clasificacion_3 LIKE '%R40/21/22 %') OR (sus.clasificacion_3 LIKE '%R40/20/21/22') OR (sus.clasificacion_3 LIKE '%R40/20/21/22;%') OR (sus.clasificacion_3 LIKE '%R40/20/21/22 %') OR (sus.clasificacion_4 LIKE '%R40') OR (sus.clasificacion_4 LIKE '%R40;%') OR (sus.clasificacion_4 LIKE '%R40 %') OR (sus.clasificacion_4 LIKE '%R45') OR (sus.clasificacion_4 LIKE '%R45;%') OR (sus.clasificacion_4 LIKE '%R45 %') OR (sus.clasificacion_4 LIKE '%R49') OR (sus.clasificacion_4 LIKE '%R49;%') OR (sus.clasificacion_4 LIKE '%R49 %') OR (sus.clasificacion_4 LIKE '%R40/20') OR (sus.clasificacion_4 LIKE '%R40/20;%') OR (sus.clasificacion_4 LIKE '%R40/20 %') OR (sus.clasificacion_4 LIKE '%R40/21') OR (sus.clasificacion_4 LIKE '%R40/21;%') OR (sus.clasificacion_4 LIKE '%R40/21 %') OR (sus.clasificacion_4 LIKE '%R40/22') OR (sus.clasificacion_4 LIKE '%R40/22;%') OR (sus.clasificacion_4 LIKE '%R40/22 %') OR (sus.clasificacion_4 LIKE '%R40/20/21') OR (sus.clasificacion_4 LIKE '%R40/20/21;%') OR (sus.clasificacion_4 LIKE '%R40/20/21 %') OR (sus.clasificacion_4 LIKE '%R40/20/22') OR (sus.clasificacion_4 LIKE '%R40/20/22;%') OR (sus.clasificacion_4 LIKE '%R40/20/22 %') OR (sus.clasificacion_4 LIKE '%R40/21/22') OR (sus.clasificacion_4 LIKE '%R40/21/22;%') OR (sus.clasificacion_4 LIKE '%R40/21/22 %') OR (sus.clasificacion_4 LIKE '%R40/20/21/22') OR (sus.clasificacion_4 LIKE '%R40/20/21/22;%') OR (sus.clasificacion_4 LIKE '%R40/20/21/22 %') OR (sus.clasificacion_5 LIKE '%R40') OR (sus.clasificacion_5 LIKE '%R40;%') OR (sus.clasificacion_5 LIKE '%R40 %') OR (sus.clasificacion_5 LIKE '%R45') OR (sus.clasificacion_5 LIKE '%R45;%') OR (sus.clasificacion_5 LIKE '%R45 %') OR (sus.clasificacion_5 LIKE '%R49') OR (sus.clasificacion_5 LIKE '%R49;%') OR (sus.clasificacion_5 LIKE '%R49 %') OR (sus.clasificacion_5 LIKE '%R40/20') OR (sus.clasificacion_5 LIKE '%R40/20;%') OR (sus.clasificacion_5 LIKE '%R40/20 %') OR (sus.clasificacion_5 LIKE '%R40/21') OR (sus.clasificacion_5 LIKE '%R40/21;%') OR (sus.clasificacion_5 LIKE '%R40/21 %') OR (sus.clasificacion_5 LIKE '%R40/22') OR (sus.clasificacion_5 LIKE '%R40/22;%') OR (sus.clasificacion_5 LIKE '%R40/22 %') OR (sus.clasificacion_5 LIKE '%R40/20/21') OR (sus.clasificacion_5 LIKE '%R40/20/21;%') OR (sus.clasificacion_5 LIKE '%R40/20/21 %') OR (sus.clasificacion_5 LIKE '%R40/20/22') OR (sus.clasificacion_5 LIKE '%R40/20/22;%') OR (sus.clasificacion_5 LIKE '%R40/20/22 %') OR (sus.clasificacion_5 LIKE '%R40/21/22') OR (sus.clasificacion_5 LIKE '%R40/21/22;%') OR (sus.clasificacion_5 LIKE '%R40/21/22 %') OR (sus.clasificacion_5 LIKE '%R40/20/21/22') OR (sus.clasificacion_5 LIKE '%R40/20/21/22;%') OR (sus.clasificacion_5 LIKE '%R40/20/21/22 %') OR (sus.clasificacion_6 LIKE '%R40') OR (sus.clasificacion_6 LIKE '%R40;%') OR (sus.clasificacion_6 LIKE '%R40 %') OR (sus.clasificacion_6 LIKE '%R45') OR (sus.clasificacion_6 LIKE '%R45;%') OR (sus.clasificacion_6 LIKE '%R45 %') OR (sus.clasificacion_6 LIKE '%R49') OR (sus.clasificacion_6 LIKE '%R49;%') OR (sus.clasificacion_6 LIKE '%R49 %') OR (sus.clasificacion_6 LIKE '%R40/20') OR (sus.clasificacion_6 LIKE '%R40/20;%') OR (sus.clasificacion_6 LIKE '%R40/20 %') OR (sus.clasificacion_6 LIKE '%R40/21') OR (sus.clasificacion_6 LIKE '%R40/21;%') OR (sus.clasificacion_6 LIKE '%R40/21 %') OR (sus.clasificacion_6 LIKE '%R40/22') OR (sus.clasificacion_6 LIKE '%R40/22;%') OR (sus.clasificacion_6 LIKE '%R40/22 %') OR (sus.clasificacion_6 LIKE '%R40/20/21') OR (sus.clasificacion_6 LIKE '%R40/20/21;%') OR (sus.clasificacion_6 LIKE '%R40/20/21 %') OR (sus.clasificacion_6 LIKE '%R40/20/22') OR (sus.clasificacion_6 LIKE '%R40/20/22;%') OR (sus.clasificacion_6 LIKE '%R40/20/22 %') OR (sus.clasificacion_6 LIKE '%R40/21/22') OR (sus.clasificacion_6 LIKE '%R40/21/22;%') OR (sus.clasificacion_6 LIKE '%R40/21/22 %') OR (sus.clasificacion_6 LIKE '%R40/20/21/22') OR (sus.clasificacion_6 LIKE '%R40/20/21/22;%') OR (sus.clasificacion_6 LIKE '%R40/20/21/22 %') OR (sus.clasificacion_7 LIKE '%R40') OR (sus.clasificacion_7 LIKE '%R40;%') OR (sus.clasificacion_7 LIKE '%R40 %') OR (sus.clasificacion_7 LIKE '%R45') OR (sus.clasificacion_7 LIKE '%R45;%') OR (sus.clasificacion_7 LIKE '%R45 %') OR (sus.clasificacion_7 LIKE '%R49') OR (sus.clasificacion_7 LIKE '%R49;%') OR (sus.clasificacion_7 LIKE '%R49 %') OR (sus.clasificacion_7 LIKE '%R40/20') OR (sus.clasificacion_7 LIKE '%R40/20;%') OR (sus.clasificacion_7 LIKE '%R40/20 %') OR (sus.clasificacion_7 LIKE '%R40/21') OR (sus.clasificacion_7 LIKE '%R40/21;%') OR (sus.clasificacion_7 LIKE '%R40/21 %') OR (sus.clasificacion_7 LIKE '%R40/22') OR (sus.clasificacion_7 LIKE '%R40/22;%') OR (sus.clasificacion_7 LIKE '%R40/22 %') OR (sus.clasificacion_7 LIKE '%R40/20/21') OR (sus.clasificacion_7 LIKE '%R40/20/21;%') OR (sus.clasificacion_7 LIKE '%R40/20/21 %') OR (sus.clasificacion_7 LIKE '%R40/20/22') OR (sus.clasificacion_7 LIKE '%R40/20/22;%') OR (sus.clasificacion_7 LIKE '%R40/20/22 %') OR (sus.clasificacion_7 LIKE '%R40/21/22') OR (sus.clasificacion_7 LIKE '%R40/21/22;%') OR (sus.clasificacion_7 LIKE '%R40/21/22 %') OR (sus.clasificacion_7 LIKE '%R40/20/21/22') OR (sus.clasificacion_7 LIKE '%R40/20/21/22;%') OR (sus.clasificacion_7 LIKE '%R40/20/21/22 %') OR (sus.clasificacion_8 LIKE '%R40') OR (sus.clasificacion_8 LIKE '%R40;%') OR (sus.clasificacion_8 LIKE '%R40 %') OR (sus.clasificacion_8 LIKE '%R45') OR (sus.clasificacion_8 LIKE '%R45;%') OR (sus.clasificacion_8 LIKE '%R45 %') OR (sus.clasificacion_8 LIKE '%R49') OR (sus.clasificacion_8 LIKE '%R49;%') OR (sus.clasificacion_8 LIKE '%R49 %') OR (sus.clasificacion_8 LIKE '%R40/20') OR (sus.clasificacion_8 LIKE '%R40/20;%') OR (sus.clasificacion_8 LIKE '%R40/20 %') OR (sus.clasificacion_8 LIKE '%R40/21') OR (sus.clasificacion_8 LIKE '%R40/21;%') OR (sus.clasificacion_8 LIKE '%R40/21 %') OR (sus.clasificacion_8 LIKE '%R40/22') OR (sus.clasificacion_8 LIKE '%R40/22;%') OR (sus.clasificacion_8 LIKE '%R40/22 %') OR (sus.clasificacion_8 LIKE '%R40/20/21') OR (sus.clasificacion_8 LIKE '%R40/20/21;%') OR (sus.clasificacion_8 LIKE '%R40/20/21 %') OR (sus.clasificacion_8 LIKE '%R40/20/22') OR (sus.clasificacion_8 LIKE '%R40/20/22;%') OR (sus.clasificacion_8 LIKE '%R40/20/22 %') OR (sus.clasificacion_8 LIKE '%R40/21/22') OR (sus.clasificacion_8 LIKE '%R40/21/22;%') OR (sus.clasificacion_8 LIKE '%R40/21/22 %') OR (sus.clasificacion_8 LIKE '%R40/20/21/22') OR (sus.clasificacion_8 LIKE '%R40/20/21/22;%') OR (sus.clasificacion_8 LIKE '%R40/20/21/22 %') OR (sus.clasificacion_9 LIKE '%R40') OR (sus.clasificacion_9 LIKE '%R40;%') OR (sus.clasificacion_9 LIKE '%R40 %') OR (sus.clasificacion_9 LIKE '%R45') OR (sus.clasificacion_9 LIKE '%R45;%') OR (sus.clasificacion_9 LIKE '%R45 %') OR (sus.clasificacion_9 LIKE '%R49') OR (sus.clasificacion_9 LIKE '%R49;%') OR (sus.clasificacion_9 LIKE '%R49 %') OR (sus.clasificacion_9 LIKE '%R40/20') OR (sus.clasificacion_9 LIKE '%R40/20;%') OR (sus.clasificacion_9 LIKE '%R40/20 %') OR (sus.clasificacion_9 LIKE '%R40/21') OR (sus.clasificacion_9 LIKE '%R40/21;%') OR (sus.clasificacion_9 LIKE '%R40/21 %') OR (sus.clasificacion_9 LIKE '%R40/22') OR (sus.clasificacion_9 LIKE '%R40/22;%') OR (sus.clasificacion_9 LIKE '%R40/22 %') OR (sus.clasificacion_9 LIKE '%R40/20/21') OR (sus.clasificacion_9 LIKE '%R40/20/21;%') OR (sus.clasificacion_9 LIKE '%R40/20/21 %') OR (sus.clasificacion_9 LIKE '%R40/20/22') OR (sus.clasificacion_9 LIKE '%R40/20/22;%') OR (sus.clasificacion_9 LIKE '%R40/20/22 %') OR (sus.clasificacion_9 LIKE '%R40/21/22') OR (sus.clasificacion_9 LIKE '%R40/21/22;%') OR (sus.clasificacion_9 LIKE '%R40/21/22 %') OR (sus.clasificacion_9 LIKE '%R40/20/21/22') OR (sus.clasificacion_9 LIKE '%R40/20/21/22;%') OR (sus.clasificacion_9 LIKE '%R40/20/21/22 %') OR (sus.clasificacion_10 LIKE '%R40') OR (sus.clasificacion_10 LIKE '%R40;%') OR (sus.clasificacion_10 LIKE '%R40 %') OR (sus.clasificacion_10 LIKE '%R45') OR (sus.clasificacion_10 LIKE '%R45;%') OR (sus.clasificacion_10 LIKE '%R45 %') OR (sus.clasificacion_10 LIKE '%R49') OR (sus.clasificacion_10 LIKE '%R49;%') OR (sus.clasificacion_10 LIKE '%R49 %') OR (sus.clasificacion_10 LIKE '%R40/20') OR (sus.clasificacion_10 LIKE '%R40/20;%') OR (sus.clasificacion_10 LIKE '%R40/20 %') OR (sus.clasificacion_10 LIKE '%R40/21') OR (sus.clasificacion_10 LIKE '%R40/21;%') OR (sus.clasificacion_10 LIKE '%R40/21 %') OR (sus.clasificacion_10 LIKE '%R40/22') OR (sus.clasificacion_10 LIKE '%R40/22;%') OR (sus.clasificacion_10 LIKE '%R40/22 %') OR (sus.clasificacion_10 LIKE '%R40/20/21') OR (sus.clasificacion_10 LIKE '%R40/20/21;%') OR (sus.clasificacion_10 LIKE '%R40/20/21 %') OR (sus.clasificacion_10 LIKE '%R40/20/22') OR (sus.clasificacion_10 LIKE '%R40/20/22;%') OR (sus.clasificacion_10 LIKE '%R40/20/22 %') OR (sus.clasificacion_10 LIKE '%R40/21/22') OR (sus.clasificacion_10 LIKE '%R40/21/22;%') OR (sus.clasificacion_10 LIKE '%R40/21/22 %') OR (sus.clasificacion_10 LIKE '%R40/20/21/22') OR (sus.clasificacion_10 LIKE '%R40/20/21/22;%') OR (sus.clasificacion_10 LIKE '%R40/20/21/22 %') OR (sus.clasificacion_11 LIKE '%R40') OR (sus.clasificacion_11 LIKE '%R40;%') OR (sus.clasificacion_11 LIKE '%R40 %') OR (sus.clasificacion_11 LIKE '%R45') OR (sus.clasificacion_11 LIKE '%R45;%') OR (sus.clasificacion_11 LIKE '%R45 %') OR (sus.clasificacion_11 LIKE '%R49') OR (sus.clasificacion_11 LIKE '%R49;%') OR (sus.clasificacion_11 LIKE '%R49 %') OR (sus.clasificacion_11 LIKE '%R40/20') OR (sus.clasificacion_11 LIKE '%R40/20;%') OR (sus.clasificacion_11 LIKE '%R40/20 %') OR (sus.clasificacion_11 LIKE '%R40/21') OR (sus.clasificacion_11 LIKE '%R40/21;%') OR (sus.clasificacion_11 LIKE '%R40/21 %') OR (sus.clasificacion_11 LIKE '%R40/22') OR (sus.clasificacion_11 LIKE '%R40/22;%') OR (sus.clasificacion_11 LIKE '%R40/22 %') OR (sus.clasificacion_11 LIKE '%R40/20/21') OR (sus.clasificacion_11 LIKE '%R40/20/21;%') OR (sus.clasificacion_11 LIKE '%R40/20/21 %') OR (sus.clasificacion_11 LIKE '%R40/20/22') OR (sus.clasificacion_11 LIKE '%R40/20/22;%') OR (sus.clasificacion_11 LIKE '%R40/20/22 %') OR (sus.clasificacion_11 LIKE '%R40/21/22') OR (sus.clasificacion_11 LIKE '%R40/21/22;%') OR (sus.clasificacion_11 LIKE '%R40/21/22 %') OR (sus.clasificacion_11 LIKE '%R40/20/21/22') OR (sus.clasificacion_11 LIKE '%R40/20/21/22;%') OR (sus.clasificacion_11 LIKE '%R40/20/21/22 %') OR (sus.clasificacion_12 LIKE '%R40') OR (sus.clasificacion_12 LIKE '%R40;%') OR (sus.clasificacion_12 LIKE '%R40 %') OR (sus.clasificacion_12 LIKE '%R45') OR (sus.clasificacion_12 LIKE '%R45;%') OR (sus.clasificacion_12 LIKE '%R45 %') OR (sus.clasificacion_12 LIKE '%R49') OR (sus.clasificacion_12 LIKE '%R49;%') OR (sus.clasificacion_12 LIKE '%R49 %') OR (sus.clasificacion_12 LIKE '%R40/20') OR (sus.clasificacion_12 LIKE '%R40/20;%') OR (sus.clasificacion_12 LIKE '%R40/20 %') OR (sus.clasificacion_12 LIKE '%R40/21') OR (sus.clasificacion_12 LIKE '%R40/21;%') OR (sus.clasificacion_12 LIKE '%R40/21 %') OR (sus.clasificacion_12 LIKE '%R40/22') OR (sus.clasificacion_12 LIKE '%R40/22;%') OR (sus.clasificacion_12 LIKE '%R40/22 %') OR (sus.clasificacion_12 LIKE '%R40/20/21') OR (sus.clasificacion_12 LIKE '%R40/20/21;%') OR (sus.clasificacion_12 LIKE '%R40/20/21 %') OR (sus.clasificacion_12 LIKE '%R40/20/22') OR (sus.clasificacion_12 LIKE '%R40/20/22;%') OR (sus.clasificacion_12 LIKE '%R40/20/22 %') OR (sus.clasificacion_12 LIKE '%R40/21/22') OR (sus.clasificacion_12 LIKE '%R40/21/22;%') OR (sus.clasificacion_12 LIKE '%R40/21/22 %') OR (sus.clasificacion_12 LIKE '%R40/20/21/22') OR (sus.clasificacion_12 LIKE '%R40/20/21/22;%') OR (sus.clasificacion_12 LIKE '%R40/20/21/22 %') OR (sus.clasificacion_13 LIKE '%R40') OR (sus.clasificacion_13 LIKE '%R40;%') OR (sus.clasificacion_13 LIKE '%R40 %') OR (sus.clasificacion_13 LIKE '%R45') OR (sus.clasificacion_13 LIKE '%R45;%') OR (sus.clasificacion_13 LIKE '%R45 %') OR (sus.clasificacion_13 LIKE '%R49') OR (sus.clasificacion_13 LIKE '%R49;%') OR (sus.clasificacion_13 LIKE '%R49 %') OR (sus.clasificacion_13 LIKE '%R40/20') OR (sus.clasificacion_13 LIKE '%R40/20;%') OR (sus.clasificacion_13 LIKE '%R40/20 %') OR (sus.clasificacion_13 LIKE '%R40/21') OR (sus.clasificacion_13 LIKE '%R40/21;%') OR (sus.clasificacion_13 LIKE '%R40/21 %') OR (sus.clasificacion_13 LIKE '%R40/22') OR (sus.clasificacion_13 LIKE '%R40/22;%') OR (sus.clasificacion_13 LIKE '%R40/22 %') OR (sus.clasificacion_13 LIKE '%R40/20/21') OR (sus.clasificacion_13 LIKE '%R40/20/21;%') OR (sus.clasificacion_13 LIKE '%R40/20/21 %') OR (sus.clasificacion_13 LIKE '%R40/20/22') OR (sus.clasificacion_13 LIKE '%R40/20/22;%') OR (sus.clasificacion_13 LIKE '%R40/20/22 %') OR (sus.clasificacion_13 LIKE '%R40/21/22') OR (sus.clasificacion_13 LIKE '%R40/21/22;%') OR (sus.clasificacion_13 LIKE '%R40/21/22 %') OR (sus.clasificacion_13 LIKE '%R40/20/21/22') OR (sus.clasificacion_13 LIKE '%R40/20/21/22;%') OR (sus.clasificacion_13 LIKE '%R40/20/21/22 %') OR (sus.clasificacion_14 LIKE '%R40') OR (sus.clasificacion_14 LIKE '%R40;%') OR (sus.clasificacion_14 LIKE '%R40 %') OR (sus.clasificacion_14 LIKE '%R45') OR (sus.clasificacion_14 LIKE '%R45;%') OR (sus.clasificacion_14 LIKE '%R45 %') OR (sus.clasificacion_14 LIKE '%R49') OR (sus.clasificacion_14 LIKE '%R49;%') OR (sus.clasificacion_14 LIKE '%R49 %') OR (sus.clasificacion_14 LIKE '%R40/20') OR (sus.clasificacion_14 LIKE '%R40/20;%') OR (sus.clasificacion_14 LIKE '%R40/20 %') OR (sus.clasificacion_14 LIKE '%R40/21') OR (sus.clasificacion_14 LIKE '%R40/21;%') OR (sus.clasificacion_14 LIKE '%R40/21 %') OR (sus.clasificacion_14 LIKE '%R40/22') OR (sus.clasificacion_14 LIKE '%R40/22;%') OR (sus.clasificacion_14 LIKE '%R40/22 %') OR (sus.clasificacion_14 LIKE '%R40/20/21') OR (sus.clasificacion_14 LIKE '%R40/20/21;%') OR (sus.clasificacion_14 LIKE '%R40/20/21 %') OR (sus.clasificacion_14 LIKE '%R40/20/22') OR (sus.clasificacion_14 LIKE '%R40/20/22;%') OR (sus.clasificacion_14 LIKE '%R40/20/22 %') OR (sus.clasificacion_14 LIKE '%R40/21/22') OR (sus.clasificacion_14 LIKE '%R40/21/22;%') OR (sus.clasificacion_14 LIKE '%R40/21/22 %') OR (sus.clasificacion_14 LIKE '%R40/20/21/22') OR (sus.clasificacion_14 LIKE '%R40/20/21/22;%') OR (sus.clasificacion_14 LIKE '%R40/20/21/22 %') OR (sus.clasificacion_15 LIKE '%R40') OR (sus.clasificacion_15 LIKE '%R40;%') OR (sus.clasificacion_15 LIKE '%R40 %') OR (sus.clasificacion_15 LIKE '%R45') OR (sus.clasificacion_15 LIKE '%R45;%') OR (sus.clasificacion_15 LIKE '%R45 %') OR (sus.clasificacion_15 LIKE '%R49') OR (sus.clasificacion_15 LIKE '%R49;%') OR (sus.clasificacion_15 LIKE '%R49 %') OR (sus.clasificacion_15 LIKE '%R40/20') OR (sus.clasificacion_15 LIKE '%R40/20;%') OR (sus.clasificacion_15 LIKE '%R40/20 %') OR (sus.clasificacion_15 LIKE '%R40/21') OR (sus.clasificacion_15 LIKE '%R40/21;%') OR (sus.clasificacion_15 LIKE '%R40/21 %') OR (sus.clasificacion_15 LIKE '%R40/22') OR (sus.clasificacion_15 LIKE '%R40/22;%') OR (sus.clasificacion_15 LIKE '%R40/22 %') OR (sus.clasificacion_15 LIKE '%R40/20/21') OR (sus.clasificacion_15 LIKE '%R40/20/21;%') OR (sus.clasificacion_15 LIKE '%R40/20/21 %') OR (sus.clasificacion_15 LIKE '%R40/20/22') OR (sus.clasificacion_15 LIKE '%R40/20/22;%') OR (sus.clasificacion_15 LIKE '%R40/20/22 %') OR (sus.clasificacion_15 LIKE '%R40/21/22') OR (sus.clasificacion_15 LIKE '%R40/21/22;%') OR (sus.clasificacion_15 LIKE '%R40/21/22 %') OR (sus.clasificacion_15 LIKE '%R40/20/21/22') OR (sus.clasificacion_15 LIKE '%R40/20/21/22;%') OR (sus.clasificacion_15 LIKE '%R40/20/21/22 %'))) OR sus.id IN (select distinct sus.id from dn_risc_sustancias as sus WHERE ((sus.frases_r_danesa LIKE '%R40') OR (sus.frases_r_danesa LIKE '%R40;%') OR (sus.frases_r_danesa LIKE '%R40 %') OR (sus.frases_r_danesa LIKE '%R45') OR (sus.frases_r_danesa LIKE '%R45;%') OR (sus.frases_r_danesa LIKE '%R45 %') OR (sus.frases_r_danesa LIKE '%R49') OR (sus.frases_r_danesa LIKE '%R49;%') OR (sus.frases_r_danesa LIKE '%R49 %') OR (sus.frases_r_danesa LIKE '%R40/20') OR (sus.frases_r_danesa LIKE '%R40/20;%') OR (sus.frases_r_danesa LIKE '%R40/20 %') OR (sus.frases_r_danesa LIKE '%R40/21') OR (sus.frases_r_danesa LIKE '%R40/21;%') OR (sus.frases_r_danesa LIKE '%R40/21 %') OR (sus.frases_r_danesa LIKE '%R40/22') OR (sus.frases_r_danesa LIKE '%R40/22;%') OR (sus.frases_r_danesa LIKE '%R40/22 %') OR (sus.frases_r_danesa LIKE '%R40/20/21') OR (sus.frases_r_danesa LIKE '%R40/20/21;%') OR (sus.frases_r_danesa LIKE '%R40/20/21 %') OR (sus.frases_r_danesa LIKE '%R40/20/22') OR (sus.frases_r_danesa LIKE '%R40/20/22;%') OR (sus.frases_r_danesa LIKE '%R40/20/22 %') OR (sus.frases_r_danesa LIKE '%R40/21/22') OR (sus.frases_r_danesa LIKE '%R40/21/22;%') OR (sus.frases_r_danesa LIKE '%R40/21/22 %') OR (sus.frases_r_danesa LIKE '%R40/20/21/22') OR (sus.frases_r_danesa LIKE '%R40/20/21/22;%') OR (sus.frases_r_danesa LIKE '%R40/20/21/22 %'))) OR sus.id IN (select distinct sus.id from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_iarc ON (sus.id=dn_risc_sustancias_iarc.id_sustancia) WHERE (dn_risc_sustancias_iarc.grupo_iarc<>'' AND dn_risc_sustancias_iarc.grupo_iarc NOT LIKE '%3%')) OR sus.id IN (select distinct sus.id from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_cancer_otras ON (sus.id=dn_risc_sustancias_cancer_otras.id_sustancia) WHERE (dn_risc_sustancias_cancer_otras.categoria_cancer_otras<>'')) OR sus.id IN (select distinct sus.id from dn_risc_sustancias as sus WHERE ((sus.clasificacion_1 LIKE '%R46') OR (sus.clasificacion_1 LIKE '%R46;%') OR (sus.clasificacion_1 LIKE '%R46 %') OR (sus.clasificacion_1 LIKE '%R68') OR (sus.clasificacion_1 LIKE '%R68;%') OR (sus.clasificacion_1 LIKE '%R68 %') OR (sus.clasificacion_1 LIKE '%R68/20') OR (sus.clasificacion_1 LIKE '%R68/20;%') OR (sus.clasificacion_1 LIKE '%R68/20 %') OR (sus.clasificacion_1 LIKE '%R68/21') OR (sus.clasificacion_1 LIKE '%R68/21;%') OR (sus.clasificacion_1 LIKE '%R68/21 %') OR (sus.clasificacion_1 LIKE '%R68/22') OR (sus.clasificacion_1 LIKE '%R68/22;%') OR (sus.clasificacion_1 LIKE '%R68/22 %') OR (sus.clasificacion_1 LIKE '%R68/20/21') OR (sus.clasificacion_1 LIKE '%R68/20/21;%') OR (sus.clasificacion_1 LIKE '%R68/20/21 %') OR (sus.clasificacion_1 LIKE '%R68/20/22') OR (sus.clasificacion_1 LIKE '%R68/20/22;%') OR (sus.clasificacion_1 LIKE '%R68/20/22 %') OR (sus.clasificacion_1 LIKE '%R68/21/22') OR (sus.clasificacion_1 LIKE '%R68/21/22;%') OR (sus.clasificacion_1 LIKE '%R68/21/22 %') OR (sus.clasificacion_1 LIKE '%R68/20/21/22') OR (sus.clasificacion_1 LIKE '%R68/20/21/22;%') OR (sus.clasificacion_1 LIKE '%R68/20/21/22 %') OR (sus.clasificacion_2 LIKE '%R46') OR (sus.clasificacion_2 LIKE '%R46;%') OR (sus.clasificacion_2 LIKE '%R46 %') OR (sus.clasificacion_2 LIKE '%R68') OR (sus.clasificacion_2 LIKE '%R68;%') OR (sus.clasificacion_2 LIKE '%R68 %') OR (sus.clasificacion_2 LIKE '%R68/20') OR (sus.clasificacion_2 LIKE '%R68/20;%') OR (sus.clasificacion_2 LIKE '%R68/20 %') OR (sus.clasificacion_2 LIKE '%R68/21') OR (sus.clasificacion_2 LIKE '%R68/21;%') OR (sus.clasificacion_2 LIKE '%R68/21 %') OR (sus.clasificacion_2 LIKE '%R68/22') OR (sus.clasificacion_2 LIKE '%R68/22;%') OR (sus.clasificacion_2 LIKE '%R68/22 %') OR (sus.clasificacion_2 LIKE '%R68/20/21') OR (sus.clasificacion_2 LIKE '%R68/20/21;%') OR (sus.clasificacion_2 LIKE '%R68/20/21 %') OR (sus.clasificacion_2 LIKE '%R68/20/22') OR (sus.clasificacion_2 LIKE '%R68/20/22;%') OR (sus.clasificacion_2 LIKE '%R68/20/22 %') OR (sus.clasificacion_2 LIKE '%R68/21/22') OR (sus.clasificacion_2 LIKE '%R68/21/22;%') OR (sus.clasificacion_2 LIKE '%R68/21/22 %') OR (sus.clasificacion_2 LIKE '%R68/20/21/22') OR (sus.clasificacion_2 LIKE '%R68/20/21/22;%') OR (sus.clasificacion_2 LIKE '%R68/20/21/22 %') OR (sus.clasificacion_3 LIKE '%R46') OR (sus.clasificacion_3 LIKE '%R46;%') OR (sus.clasificacion_3 LIKE '%R46 %') OR (sus.clasificacion_3 LIKE '%R68') OR (sus.clasificacion_3 LIKE '%R68;%') OR (sus.clasificacion_3 LIKE '%R68 %') OR (sus.clasificacion_3 LIKE '%R68/20') OR (sus.clasificacion_3 LIKE '%R68/20;%') OR (sus.clasificacion_3 LIKE '%R68/20 %') OR (sus.clasificacion_3 LIKE '%R68/21') OR (sus.clasificacion_3 LIKE '%R68/21;%') OR (sus.clasificacion_3 LIKE '%R68/21 %') OR (sus.clasificacion_3 LIKE '%R68/22') OR (sus.clasificacion_3 LIKE '%R68/22;%') OR (sus.clasificacion_3 LIKE '%R68/22 %') OR (sus.clasificacion_3 LIKE '%R68/20/21') OR (sus.clasificacion_3 LIKE '%R68/20/21;%') OR (sus.clasificacion_3 LIKE '%R68/20/21 %') OR (sus.clasificacion_3 LIKE '%R68/20/22') OR (sus.clasificacion_3 LIKE '%R68/20/22;%') OR (sus.clasificacion_3 LIKE '%R68/20/22 %') OR (sus.clasificacion_3 LIKE '%R68/21/22') OR (sus.clasificacion_3 LIKE '%R68/21/22;%') OR (sus.clasificacion_3 LIKE '%R68/21/22 %') OR (sus.clasificacion_3 LIKE '%R68/20/21/22') OR (sus.clasificacion_3 LIKE '%R68/20/21/22;%') OR (sus.clasificacion_3 LIKE '%R68/20/21/22 %') OR (sus.clasificacion_4 LIKE '%R46') OR (sus.clasificacion_4 LIKE '%R46;%') OR (sus.clasificacion_4 LIKE '%R46 %') OR (sus.clasificacion_4 LIKE '%R68') OR (sus.clasificacion_4 LIKE '%R68;%') OR (sus.clasificacion_4 LIKE '%R68 %') OR (sus.clasificacion_4 LIKE '%R68/20') OR (sus.clasificacion_4 LIKE '%R68/20;%') OR (sus.clasificacion_4 LIKE '%R68/20 %') OR (sus.clasificacion_4 LIKE '%R68/21') OR (sus.clasificacion_4 LIKE '%R68/21;%') OR (sus.clasificacion_4 LIKE '%R68/21 %') OR (sus.clasificacion_4 LIKE '%R68/22') OR (sus.clasificacion_4 LIKE '%R68/22;%') OR (sus.clasificacion_4 LIKE '%R68/22 %') OR (sus.clasificacion_4 LIKE '%R68/20/21') OR (sus.clasificacion_4 LIKE '%R68/20/21;%') OR (sus.clasificacion_4 LIKE '%R68/20/21 %') OR (sus.clasificacion_4 LIKE '%R68/20/22') OR (sus.clasificacion_4 LIKE '%R68/20/22;%') OR (sus.clasificacion_4 LIKE '%R68/20/22 %') OR (sus.clasificacion_4 LIKE '%R68/21/22') OR (sus.clasificacion_4 LIKE '%R68/21/22;%') OR (sus.clasificacion_4 LIKE '%R68/21/22 %') OR (sus.clasificacion_4 LIKE '%R68/20/21/22') OR (sus.clasificacion_4 LIKE '%R68/20/21/22;%') OR (sus.clasificacion_4 LIKE '%R68/20/21/22 %') OR (sus.clasificacion_5 LIKE '%R46') OR (sus.clasificacion_5 LIKE '%R46;%') OR (sus.clasificacion_5 LIKE '%R46 %') OR (sus.clasificacion_5 LIKE '%R68') OR (sus.clasificacion_5 LIKE '%R68;%') OR (sus.clasificacion_5 LIKE '%R68 %') OR (sus.clasificacion_5 LIKE '%R68/20') OR (sus.clasificacion_5 LIKE '%R68/20;%') OR (sus.clasificacion_5 LIKE '%R68/20 %') OR (sus.clasificacion_5 LIKE '%R68/21') OR (sus.clasificacion_5 LIKE '%R68/21;%') OR (sus.clasificacion_5 LIKE '%R68/21 %') OR (sus.clasificacion_5 LIKE '%R68/22') OR (sus.clasificacion_5 LIKE '%R68/22;%') OR (sus.clasificacion_5 LIKE '%R68/22 %') OR (sus.clasificacion_5 LIKE '%R68/20/21') OR (sus.clasificacion_5 LIKE '%R68/20/21;%') OR (sus.clasificacion_5 LIKE '%R68/20/21 %') OR (sus.clasificacion_5 LIKE '%R68/20/22') OR (sus.clasificacion_5 LIKE '%R68/20/22;%') OR (sus.clasificacion_5 LIKE '%R68/20/22 %') OR (sus.clasificacion_5 LIKE '%R68/21/22') OR (sus.clasificacion_5 LIKE '%R68/21/22;%') OR (sus.clasificacion_5 LIKE '%R68/21/22 %') OR (sus.clasificacion_5 LIKE '%R68/20/21/22') OR (sus.clasificacion_5 LIKE '%R68/20/21/22;%') OR (sus.clasificacion_5 LIKE '%R68/20/21/22 %') OR (sus.clasificacion_6 LIKE '%R46') OR (sus.clasificacion_6 LIKE '%R46;%') OR (sus.clasificacion_6 LIKE '%R46 %') OR (sus.clasificacion_6 LIKE '%R68') OR (sus.clasificacion_6 LIKE '%R68;%') OR (sus.clasificacion_6 LIKE '%R68 %') OR (sus.clasificacion_6 LIKE '%R68/20') OR (sus.clasificacion_6 LIKE '%R68/20;%') OR (sus.clasificacion_6 LIKE '%R68/20 %') OR (sus.clasificacion_6 LIKE '%R68/21') OR (sus.clasificacion_6 LIKE '%R68/21;%') OR (sus.clasificacion_6 LIKE '%R68/21 %') OR (sus.clasificacion_6 LIKE '%R68/22') OR (sus.clasificacion_6 LIKE '%R68/22;%') OR (sus.clasificacion_6 LIKE '%R68/22 %') OR (sus.clasificacion_6 LIKE '%R68/20/21') OR (sus.clasificacion_6 LIKE '%R68/20/21;%') OR (sus.clasificacion_6 LIKE '%R68/20/21 %') OR (sus.clasificacion_6 LIKE '%R68/20/22') OR (sus.clasificacion_6 LIKE '%R68/20/22;%') OR (sus.clasificacion_6 LIKE '%R68/20/22 %') OR (sus.clasificacion_6 LIKE '%R68/21/22') OR (sus.clasificacion_6 LIKE '%R68/21/22;%') OR (sus.clasificacion_6 LIKE '%R68/21/22 %') OR (sus.clasificacion_6 LIKE '%R68/20/21/22') OR (sus.clasificacion_6 LIKE '%R68/20/21/22;%') OR (sus.clasificacion_6 LIKE '%R68/20/21/22 %') OR (sus.clasificacion_7 LIKE '%R46') OR (sus.clasificacion_7 LIKE '%R46;%') OR (sus.clasificacion_7 LIKE '%R46 %') OR (sus.clasificacion_7 LIKE '%R68') OR (sus.clasificacion_7 LIKE '%R68;%') OR (sus.clasificacion_7 LIKE '%R68 %') OR (sus.clasificacion_7 LIKE '%R68/20') OR (sus.clasificacion_7 LIKE '%R68/20;%') OR (sus.clasificacion_7 LIKE '%R68/20 %') OR (sus.clasificacion_7 LIKE '%R68/21') OR (sus.clasificacion_7 LIKE '%R68/21;%') OR (sus.clasificacion_7 LIKE '%R68/21 %') OR (sus.clasificacion_7 LIKE '%R68/22') OR (sus.clasificacion_7 LIKE '%R68/22;%') OR (sus.clasificacion_7 LIKE '%R68/22 %') OR (sus.clasificacion_7 LIKE '%R68/20/21') OR (sus.clasificacion_7 LIKE '%R68/20/21;%') OR (sus.clasificacion_7 LIKE '%R68/20/21 %') OR (sus.clasificacion_7 LIKE '%R68/20/22') OR (sus.clasificacion_7 LIKE '%R68/20/22;%') OR (sus.clasificacion_7 LIKE '%R68/20/22 %') OR (sus.clasificacion_7 LIKE '%R68/21/22') OR (sus.clasificacion_7 LIKE '%R68/21/22;%') OR (sus.clasificacion_7 LIKE '%R68/21/22 %') OR (sus.clasificacion_7 LIKE '%R68/20/21/22') OR (sus.clasificacion_7 LIKE '%R68/20/21/22;%') OR (sus.clasificacion_7 LIKE '%R68/20/21/22 %') OR (sus.clasificacion_8 LIKE '%R46') OR (sus.clasificacion_8 LIKE '%R46;%') OR (sus.clasificacion_8 LIKE '%R46 %') OR (sus.clasificacion_8 LIKE '%R68') OR (sus.clasificacion_8 LIKE '%R68;%') OR (sus.clasificacion_8 LIKE '%R68 %') OR (sus.clasificacion_8 LIKE '%R68/20') OR (sus.clasificacion_8 LIKE '%R68/20;%') OR (sus.clasificacion_8 LIKE '%R68/20 %') OR (sus.clasificacion_8 LIKE '%R68/21') OR (sus.clasificacion_8 LIKE '%R68/21;%') OR (sus.clasificacion_8 LIKE '%R68/21 %') OR (sus.clasificacion_8 LIKE '%R68/22') OR (sus.clasificacion_8 LIKE '%R68/22;%') OR (sus.clasificacion_8 LIKE '%R68/22 %') OR (sus.clasificacion_8 LIKE '%R68/20/21') OR (sus.clasificacion_8 LIKE '%R68/20/21;%') OR (sus.clasificacion_8 LIKE '%R68/20/21 %') OR (sus.clasificacion_8 LIKE '%R68/20/22') OR (sus.clasificacion_8 LIKE '%R68/20/22;%') OR (sus.clasificacion_8 LIKE '%R68/20/22 %') OR (sus.clasificacion_8 LIKE '%R68/21/22') OR (sus.clasificacion_8 LIKE '%R68/21/22;%') OR (sus.clasificacion_8 LIKE '%R68/21/22 %') OR (sus.clasificacion_8 LIKE '%R68/20/21/22') OR (sus.clasificacion_8 LIKE '%R68/20/21/22;%') OR (sus.clasificacion_8 LIKE '%R68/20/21/22 %') OR (sus.clasificacion_9 LIKE '%R46') OR (sus.clasificacion_9 LIKE '%R46;%') OR (sus.clasificacion_9 LIKE '%R46 %') OR (sus.clasificacion_9 LIKE '%R68') OR (sus.clasificacion_9 LIKE '%R68;%') OR (sus.clasificacion_9 LIKE '%R68 %') OR (sus.clasificacion_9 LIKE '%R68/20') OR (sus.clasificacion_9 LIKE '%R68/20;%') OR (sus.clasificacion_9 LIKE '%R68/20 %') OR (sus.clasificacion_9 LIKE '%R68/21') OR (sus.clasificacion_9 LIKE '%R68/21;%') OR (sus.clasificacion_9 LIKE '%R68/21 %') OR (sus.clasificacion_9 LIKE '%R68/22') OR (sus.clasificacion_9 LIKE '%R68/22;%') OR (sus.clasificacion_9 LIKE '%R68/22 %') OR (sus.clasificacion_9 LIKE '%R68/20/21') OR (sus.clasificacion_9 LIKE '%R68/20/21;%') OR (sus.clasificacion_9 LIKE '%R68/20/21 %') OR (sus.clasificacion_9 LIKE '%R68/20/22') OR (sus.clasificacion_9 LIKE '%R68/20/22;%') OR (sus.clasificacion_9 LIKE '%R68/20/22 %') OR (sus.clasificacion_9 LIKE '%R68/21/22') OR (sus.clasificacion_9 LIKE '%R68/21/22;%') OR (sus.clasificacion_9 LIKE '%R68/21/22 %') OR (sus.clasificacion_9 LIKE '%R68/20/21/22') OR (sus.clasificacion_9 LIKE '%R68/20/21/22;%') OR (sus.clasificacion_9 LIKE '%R68/20/21/22 %') OR (sus.clasificacion_10 LIKE '%R46') OR (sus.clasificacion_10 LIKE '%R46;%') OR (sus.clasificacion_10 LIKE '%R46 %') OR (sus.clasificacion_10 LIKE '%R68') OR (sus.clasificacion_10 LIKE '%R68;%') OR (sus.clasificacion_10 LIKE '%R68 %') OR (sus.clasificacion_10 LIKE '%R68/20') OR (sus.clasificacion_10 LIKE '%R68/20;%') OR (sus.clasificacion_10 LIKE '%R68/20 %') OR (sus.clasificacion_10 LIKE '%R68/21') OR (sus.clasificacion_10 LIKE '%R68/21;%') OR (sus.clasificacion_10 LIKE '%R68/21 %') OR (sus.clasificacion_10 LIKE '%R68/22') OR (sus.clasificacion_10 LIKE '%R68/22;%') OR (sus.clasificacion_10 LIKE '%R68/22 %') OR (sus.clasificacion_10 LIKE '%R68/20/21') OR (sus.clasificacion_10 LIKE '%R68/20/21;%') OR (sus.clasificacion_10 LIKE '%R68/20/21 %') OR (sus.clasificacion_10 LIKE '%R68/20/22') OR (sus.clasificacion_10 LIKE '%R68/20/22;%') OR (sus.clasificacion_10 LIKE '%R68/20/22 %') OR (sus.clasificacion_10 LIKE '%R68/21/22') OR (sus.clasificacion_10 LIKE '%R68/21/22;%') OR (sus.clasificacion_10 LIKE '%R68/21/22 %') OR (sus.clasificacion_10 LIKE '%R68/20/21/22') OR (sus.clasificacion_10 LIKE '%R68/20/21/22;%') OR (sus.clasificacion_10 LIKE '%R68/20/21/22 %') OR (sus.clasificacion_11 LIKE '%R46') OR (sus.clasificacion_11 LIKE '%R46;%') OR (sus.clasificacion_11 LIKE '%R46 %') OR (sus.clasificacion_11 LIKE '%R68') OR (sus.clasificacion_11 LIKE '%R68;%') OR (sus.clasificacion_11 LIKE '%R68 %') OR (sus.clasificacion_11 LIKE '%R68/20') OR (sus.clasificacion_11 LIKE '%R68/20;%') OR (sus.clasificacion_11 LIKE '%R68/20 %') OR (sus.clasificacion_11 LIKE '%R68/21') OR (sus.clasificacion_11 LIKE '%R68/21;%') OR (sus.clasificacion_11 LIKE '%R68/21 %') OR (sus.clasificacion_11 LIKE '%R68/22') OR (sus.clasificacion_11 LIKE '%R68/22;%') OR (sus.clasificacion_11 LIKE '%R68/22 %') OR (sus.clasificacion_11 LIKE '%R68/20/21') OR (sus.clasificacion_11 LIKE '%R68/20/21;%') OR (sus.clasificacion_11 LIKE '%R68/20/21 %') OR (sus.clasificacion_11 LIKE '%R68/20/22') OR (sus.clasificacion_11 LIKE '%R68/20/22;%') OR (sus.clasificacion_11 LIKE '%R68/20/22 %') OR (sus.clasificacion_11 LIKE '%R68/21/22') OR (sus.clasificacion_11 LIKE '%R68/21/22;%') OR (sus.clasificacion_11 LIKE '%R68/21/22 %') OR (sus.clasificacion_11 LIKE '%R68/20/21/22') OR (sus.clasificacion_11 LIKE '%R68/20/21/22;%') OR (sus.clasificacion_11 LIKE '%R68/20/21/22 %') OR (sus.clasificacion_12 LIKE '%R46') OR (sus.clasificacion_12 LIKE '%R46;%') OR (sus.clasificacion_12 LIKE '%R46 %') OR (sus.clasificacion_12 LIKE '%R68') OR (sus.clasificacion_12 LIKE '%R68;%') OR (sus.clasificacion_12 LIKE '%R68 %') OR (sus.clasificacion_12 LIKE '%R68/20') OR (sus.clasificacion_12 LIKE '%R68/20;%') OR (sus.clasificacion_12 LIKE '%R68/20 %') OR (sus.clasificacion_12 LIKE '%R68/21') OR (sus.clasificacion_12 LIKE '%R68/21;%') OR (sus.clasificacion_12 LIKE '%R68/21 %') OR (sus.clasificacion_12 LIKE '%R68/22') OR (sus.clasificacion_12 LIKE '%R68/22;%') OR (sus.clasificacion_12 LIKE '%R68/22 %') OR (sus.clasificacion_12 LIKE '%R68/20/21') OR (sus.clasificacion_12 LIKE '%R68/20/21;%') OR (sus.clasificacion_12 LIKE '%R68/20/21 %') OR (sus.clasificacion_12 LIKE '%R68/20/22') OR (sus.clasificacion_12 LIKE '%R68/20/22;%') OR (sus.clasificacion_12 LIKE '%R68/20/22 %') OR (sus.clasificacion_12 LIKE '%R68/21/22') OR (sus.clasificacion_12 LIKE '%R68/21/22;%') OR (sus.clasificacion_12 LIKE '%R68/21/22 %') OR (sus.clasificacion_12 LIKE '%R68/20/21/22') OR (sus.clasificacion_12 LIKE '%R68/20/21/22;%') OR (sus.clasificacion_12 LIKE '%R68/20/21/22 %') OR (sus.clasificacion_13 LIKE '%R46') OR (sus.clasificacion_13 LIKE '%R46;%') OR (sus.clasificacion_13 LIKE '%R46 %') OR (sus.clasificacion_13 LIKE '%R68') OR (sus.clasificacion_13 LIKE '%R68;%') OR (sus.clasificacion_13 LIKE '%R68 %') OR (sus.clasificacion_13 LIKE '%R68/20') OR (sus.clasificacion_13 LIKE '%R68/20;%') OR (sus.clasificacion_13 LIKE '%R68/20 %') OR (sus.clasificacion_13 LIKE '%R68/21') OR (sus.clasificacion_13 LIKE '%R68/21;%') OR (sus.clasificacion_13 LIKE '%R68/21 %') OR (sus.clasificacion_13 LIKE '%R68/22') OR (sus.clasificacion_13 LIKE '%R68/22;%') OR (sus.clasificacion_13 LIKE '%R68/22 %') OR (sus.clasificacion_13 LIKE '%R68/20/21') OR (sus.clasificacion_13 LIKE '%R68/20/21;%') OR (sus.clasificacion_13 LIKE '%R68/20/21 %') OR (sus.clasificacion_13 LIKE '%R68/20/22') OR (sus.clasificacion_13 LIKE '%R68/20/22;%') OR (sus.clasificacion_13 LIKE '%R68/20/22 %') OR (sus.clasificacion_13 LIKE '%R68/21/22') OR (sus.clasificacion_13 LIKE '%R68/21/22;%') OR (sus.clasificacion_13 LIKE '%R68/21/22 %') OR (sus.clasificacion_13 LIKE '%R68/20/21/22') OR (sus.clasificacion_13 LIKE '%R68/20/21/22;%') OR (sus.clasificacion_13 LIKE '%R68/20/21/22 %') OR (sus.clasificacion_14 LIKE '%R46') OR (sus.clasificacion_14 LIKE '%R46;%') OR (sus.clasificacion_14 LIKE '%R46 %') OR (sus.clasificacion_14 LIKE '%R68') OR (sus.clasificacion_14 LIKE '%R68;%') OR (sus.clasificacion_14 LIKE '%R68 %') OR (sus.clasificacion_14 LIKE '%R68/20') OR (sus.clasificacion_14 LIKE '%R68/20;%') OR (sus.clasificacion_14 LIKE '%R68/20 %') OR (sus.clasificacion_14 LIKE '%R68/21') OR (sus.clasificacion_14 LIKE '%R68/21;%') OR (sus.clasificacion_14 LIKE '%R68/21 %') OR (sus.clasificacion_14 LIKE '%R68/22') OR (sus.clasificacion_14 LIKE '%R68/22;%') OR (sus.clasificacion_14 LIKE '%R68/22 %') OR (sus.clasificacion_14 LIKE '%R68/20/21') OR (sus.clasificacion_14 LIKE '%R68/20/21;%') OR (sus.clasificacion_14 LIKE '%R68/20/21 %') OR (sus.clasificacion_14 LIKE '%R68/20/22') OR (sus.clasificacion_14 LIKE '%R68/20/22;%') OR (sus.clasificacion_14 LIKE '%R68/20/22 %') OR (sus.clasificacion_14 LIKE '%R68/21/22') OR (sus.clasificacion_14 LIKE '%R68/21/22;%') OR (sus.clasificacion_14 LIKE '%R68/21/22 %') OR (sus.clasificacion_14 LIKE '%R68/20/21/22') OR (sus.clasificacion_14 LIKE '%R68/20/21/22;%') OR (sus.clasificacion_14 LIKE '%R68/20/21/22 %') OR (sus.clasificacion_15 LIKE '%R46') OR (sus.clasificacion_15 LIKE '%R46;%') OR (sus.clasificacion_15 LIKE '%R46 %') OR (sus.clasificacion_15 LIKE '%R68') OR (sus.clasificacion_15 LIKE '%R68;%') OR (sus.clasificacion_15 LIKE '%R68 %') OR (sus.clasificacion_15 LIKE '%R68/20') OR (sus.clasificacion_15 LIKE '%R68/20;%') OR (sus.clasificacion_15 LIKE '%R68/20 %') OR (sus.clasificacion_15 LIKE '%R68/21') OR (sus.clasificacion_15 LIKE '%R68/21;%') OR (sus.clasificacion_15 LIKE '%R68/21 %') OR (sus.clasificacion_15 LIKE '%R68/22') OR (sus.clasificacion_15 LIKE '%R68/22;%') OR (sus.clasificacion_15 LIKE '%R68/22 %') OR (sus.clasificacion_15 LIKE '%R68/20/21') OR (sus.clasificacion_15 LIKE '%R68/20/21;%') OR (sus.clasificacion_15 LIKE '%R68/20/21 %') OR (sus.clasificacion_15 LIKE '%R68/20/22') OR (sus.clasificacion_15 LIKE '%R68/20/22;%') OR (sus.clasificacion_15 LIKE '%R68/20/22 %') OR (sus.clasificacion_15 LIKE '%R68/21/22') OR (sus.clasificacion_15 LIKE '%R68/21/22;%') OR (sus.clasificacion_15 LIKE '%R68/21/22 %') OR (sus.clasificacion_15 LIKE '%R68/20/21/22') OR (sus.clasificacion_15 LIKE '%R68/20/21/22;%') OR (sus.clasificacion_15 LIKE '%R68/20/21/22 %'))) OR sus.id IN (select distinct sus.id from dn_risc_sustancias as sus WHERE ((sus.frases_r_danesa LIKE '%R46') OR (sus.frases_r_danesa LIKE '%R46;%') OR (sus.frases_r_danesa LIKE '%R46 %') OR (sus.frases_r_danesa LIKE '%R68') OR (sus.frases_r_danesa LIKE '%R68;%') OR (sus.frases_r_danesa LIKE '%R68 %') OR (sus.frases_r_danesa LIKE '%R68/20') OR (sus.frases_r_danesa LIKE '%R68/20;%') OR (sus.frases_r_danesa LIKE '%R68/20 %') OR (sus.frases_r_danesa LIKE '%R68/21') OR (sus.frases_r_danesa LIKE '%R68/21;%') OR (sus.frases_r_danesa LIKE '%R68/21 %') OR (sus.frases_r_danesa LIKE '%R68/22') OR (sus.frases_r_danesa LIKE '%R68/22;%') OR (sus.frases_r_danesa LIKE '%R68/22 %') OR (sus.frases_r_danesa LIKE '%R68/20/21') OR (sus.frases_r_danesa LIKE '%R68/20/21;%') OR (sus.frases_r_danesa LIKE '%R68/20/21 %') OR (sus.frases_r_danesa LIKE '%R68/20/22') OR (sus.frases_r_danesa LIKE '%R68/20/22;%') OR (sus.frases_r_danesa LIKE '%R68/20/22 %') OR (sus.frases_r_danesa LIKE '%R68/21/22') OR (sus.frases_r_danesa LIKE '%R68/21/22;%') OR (sus.frases_r_danesa LIKE '%R68/21/22 %') OR (sus.frases_r_danesa LIKE '%R68/20/21/22') OR (sus.frases_r_danesa LIKE '%R68/20/21/22;%') OR (sus.frases_r_danesa LIKE '%R68/20/21/22 %'))) OR sus.id IN (select distinct sus.id from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_neuro_disruptor ON (sus.id=dn_risc_sustancias_neuro_disruptor.id_sustancia) WHERE (dn_risc_sustancias_neuro_disruptor.nivel_disruptor<>'')) OR sus.id IN (select distinct sus.id from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_neuro_disruptor ON (sus.id=dn_risc_sustancias_neuro_disruptor.id_sustancia) WHERE (dn_risc_sustancias_neuro_disruptor.nivel_neurotoxico<>'')) OR sus.id IN (select distinct sus.id from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (dn_risc_sustancias_ambiente.anchor_tpb<>'')) OR sus.id IN (select distinct sus.id from dn_risc_sustancias as sus WHERE ((sus.clasificacion_1 LIKE '%R42') OR (sus.clasificacion_1 LIKE '%R42;%') OR (sus.clasificacion_1 LIKE '%R42 %') OR (sus.clasificacion_1 LIKE '%R43') OR (sus.clasificacion_1 LIKE '%R43;%') OR (sus.clasificacion_1 LIKE '%R43 %') OR (sus.clasificacion_1 LIKE '%R42/43') OR (sus.clasificacion_1 LIKE '%R42/43;%') OR (sus.clasificacion_1 LIKE '%R42/43 %') OR (sus.clasificacion_2 LIKE '%R42') OR (sus.clasificacion_2 LIKE '%R42;%') OR (sus.clasificacion_2 LIKE '%R42 %') OR (sus.clasificacion_2 LIKE '%R43') OR (sus.clasificacion_2 LIKE '%R43;%') OR (sus.clasificacion_2 LIKE '%R43 %') OR (sus.clasificacion_2 LIKE '%R42/43') OR (sus.clasificacion_2 LIKE '%R42/43;%') OR (sus.clasificacion_2 LIKE '%R42/43 %') OR (sus.clasificacion_3 LIKE '%R42') OR (sus.clasificacion_3 LIKE '%R42;%') OR (sus.clasificacion_3 LIKE '%R42 %') OR (sus.clasificacion_3 LIKE '%R43') OR (sus.clasificacion_3 LIKE '%R43;%') OR (sus.clasificacion_3 LIKE '%R43 %') OR (sus.clasificacion_3 LIKE '%R42/43') OR (sus.clasificacion_3 LIKE '%R42/43;%') OR (sus.clasificacion_3 LIKE '%R42/43 %') OR (sus.clasificacion_4 LIKE '%R42') OR (sus.clasificacion_4 LIKE '%R42;%') OR (sus.clasificacion_4 LIKE '%R42 %') OR (sus.clasificacion_4 LIKE '%R43') OR (sus.clasificacion_4 LIKE '%R43;%') OR (sus.clasificacion_4 LIKE '%R43 %') OR (sus.clasificacion_4 LIKE '%R42/43') OR (sus.clasificacion_4 LIKE '%R42/43;%') OR (sus.clasificacion_4 LIKE '%R42/43 %') OR (sus.clasificacion_5 LIKE '%R42') OR (sus.clasificacion_5 LIKE '%R42;%') OR (sus.clasificacion_5 LIKE '%R42 %') OR (sus.clasificacion_5 LIKE '%R43') OR (sus.clasificacion_5 LIKE '%R43;%') OR (sus.clasificacion_5 LIKE '%R43 %') OR (sus.clasificacion_5 LIKE '%R42/43') OR (sus.clasificacion_5 LIKE '%R42/43;%') OR (sus.clasificacion_5 LIKE '%R42/43 %') OR (sus.clasificacion_6 LIKE '%R42') OR (sus.clasificacion_6 LIKE '%R42;%') OR (sus.clasificacion_6 LIKE '%R42 %') OR (sus.clasificacion_6 LIKE '%R43') OR (sus.clasificacion_6 LIKE '%R43;%') OR (sus.clasificacion_6 LIKE '%R43 %') OR (sus.clasificacion_6 LIKE '%R42/43') OR (sus.clasificacion_6 LIKE '%R42/43;%') OR (sus.clasificacion_6 LIKE '%R42/43 %') OR (sus.clasificacion_7 LIKE '%R42') OR (sus.clasificacion_7 LIKE '%R42;%') OR (sus.clasificacion_7 LIKE '%R42 %') OR (sus.clasificacion_7 LIKE '%R43') OR (sus.clasificacion_7 LIKE '%R43;%') OR (sus.clasificacion_7 LIKE '%R43 %') OR (sus.clasificacion_7 LIKE '%R42/43') OR (sus.clasificacion_7 LIKE '%R42/43;%') OR (sus.clasificacion_7 LIKE '%R42/43 %') OR (sus.clasificacion_8 LIKE '%R42') OR (sus.clasificacion_8 LIKE '%R42;%') OR (sus.clasificacion_8 LIKE '%R42 %') OR (sus.clasificacion_8 LIKE '%R43') OR (sus.clasificacion_8 LIKE '%R43;%') OR (sus.clasificacion_8 LIKE '%R43 %') OR (sus.clasificacion_8 LIKE '%R42/43') OR (sus.clasificacion_8 LIKE '%R42/43;%') OR (sus.clasificacion_8 LIKE '%R42/43 %') OR (sus.clasificacion_9 LIKE '%R42') OR (sus.clasificacion_9 LIKE '%R42;%') OR (sus.clasificacion_9 LIKE '%R42 %') OR (sus.clasificacion_9 LIKE '%R43') OR (sus.clasificacion_9 LIKE '%R43;%') OR (sus.clasificacion_9 LIKE '%R43 %') OR (sus.clasificacion_9 LIKE '%R42/43') OR (sus.clasificacion_9 LIKE '%R42/43;%') OR (sus.clasificacion_9 LIKE '%R42/43 %') OR (sus.clasificacion_10 LIKE '%R42') OR (sus.clasificacion_10 LIKE '%R42;%') OR (sus.clasificacion_10 LIKE '%R42 %') OR (sus.clasificacion_10 LIKE '%R43') OR (sus.clasificacion_10 LIKE '%R43;%') OR (sus.clasificacion_10 LIKE '%R43 %') OR (sus.clasificacion_10 LIKE '%R42/43') OR (sus.clasificacion_10 LIKE '%R42/43;%') OR (sus.clasificacion_10 LIKE '%R42/43 %') OR (sus.clasificacion_11 LIKE '%R42') OR (sus.clasificacion_11 LIKE '%R42;%') OR (sus.clasificacion_11 LIKE '%R42 %') OR (sus.clasificacion_11 LIKE '%R43') OR (sus.clasificacion_11 LIKE '%R43;%') OR (sus.clasificacion_11 LIKE '%R43 %') OR (sus.clasificacion_11 LIKE '%R42/43') OR (sus.clasificacion_11 LIKE '%R42/43;%') OR (sus.clasificacion_11 LIKE '%R42/43 %') OR (sus.clasificacion_12 LIKE '%R42') OR (sus.clasificacion_12 LIKE '%R42;%') OR (sus.clasificacion_12 LIKE '%R42 %') OR (sus.clasificacion_12 LIKE '%R43') OR (sus.clasificacion_12 LIKE '%R43;%') OR (sus.clasificacion_12 LIKE '%R43 %') OR (sus.clasificacion_12 LIKE '%R42/43') OR (sus.clasificacion_12 LIKE '%R42/43;%') OR (sus.clasificacion_12 LIKE '%R42/43 %') OR (sus.clasificacion_13 LIKE '%R42') OR (sus.clasificacion_13 LIKE '%R42;%') OR (sus.clasificacion_13 LIKE '%R42 %') OR (sus.clasificacion_13 LIKE '%R43') OR (sus.clasificacion_13 LIKE '%R43;%') OR (sus.clasificacion_13 LIKE '%R43 %') OR (sus.clasificacion_13 LIKE '%R42/43') OR (sus.clasificacion_13 LIKE '%R42/43;%') OR (sus.clasificacion_13 LIKE '%R42/43 %') OR (sus.clasificacion_14 LIKE '%R42') OR (sus.clasificacion_14 LIKE '%R42;%') OR (sus.clasificacion_14 LIKE '%R42 %') OR (sus.clasificacion_14 LIKE '%R43') OR (sus.clasificacion_14 LIKE '%R43;%') OR (sus.clasificacion_14 LIKE '%R43 %') OR (sus.clasificacion_14 LIKE '%R42/43') OR (sus.clasificacion_14 LIKE '%R42/43;%') OR (sus.clasificacion_14 LIKE '%R42/43 %') OR (sus.clasificacion_15 LIKE '%R42') OR (sus.clasificacion_15 LIKE '%R42;%') OR (sus.clasificacion_15 LIKE '%R42 %') OR (sus.clasificacion_15 LIKE '%R43') OR (sus.clasificacion_15 LIKE '%R43;%') OR (sus.clasificacion_15 LIKE '%R43 %') OR (sus.clasificacion_15 LIKE '%R42/43') OR (sus.clasificacion_15 LIKE '%R42/43;%') OR (sus.clasificacion_15 LIKE '%R42/43 %'))) OR sus.id IN (select distinct sus.id from dn_risc_sustancias as sus WHERE ((sus.frases_r_danesa LIKE '%R42') OR (sus.frases_r_danesa LIKE '%R42;%') OR (sus.frases_r_danesa LIKE '%R42 %') OR (sus.frases_r_danesa LIKE '%R43') OR (sus.frases_r_danesa LIKE '%R43;%') OR (sus.frases_r_danesa LIKE '%R43 %') OR (sus.frases_r_danesa LIKE '%R42/43') OR (sus.frases_r_danesa LIKE '%R42/43;%') OR (sus.frases_r_danesa LIKE '%R42/43 %'))) OR sus.id IN (select distinct sus.id from dn_risc_sustancias as sus WHERE ((sus.clasificacion_1 LIKE '%R60') OR (sus.clasificacion_1 LIKE '%R60;%') OR (sus.clasificacion_1 LIKE '%R60 %') OR (sus.clasificacion_1 LIKE '%R61') OR (sus.clasificacion_1 LIKE '%R61;%') OR (sus.clasificacion_1 LIKE '%R61 %') OR (sus.clasificacion_1 LIKE '%R62') OR (sus.clasificacion_1 LIKE '%R62;%') OR (sus.clasificacion_1 LIKE '%R62 %') OR (sus.clasificacion_1 LIKE '%R63') OR (sus.clasificacion_1 LIKE '%R63;%') OR (sus.clasificacion_1 LIKE '%R63 %') OR (sus.clasificacion_2 LIKE '%R60') OR (sus.clasificacion_2 LIKE '%R60;%') OR (sus.clasificacion_2 LIKE '%R60 %') OR (sus.clasificacion_2 LIKE '%R61') OR (sus.clasificacion_2 LIKE '%R61;%') OR (sus.clasificacion_2 LIKE '%R61 %') OR (sus.clasificacion_2 LIKE '%R62') OR (sus.clasificacion_2 LIKE '%R62;%') OR (sus.clasificacion_2 LIKE '%R62 %') OR (sus.clasificacion_2 LIKE '%R63') OR (sus.clasificacion_2 LIKE '%R63;%') OR (sus.clasificacion_2 LIKE '%R63 %') OR (sus.clasificacion_3 LIKE '%R60') OR (sus.clasificacion_3 LIKE '%R60;%') OR (sus.clasificacion_3 LIKE '%R60 %') OR (sus.clasificacion_3 LIKE '%R61') OR (sus.clasificacion_3 LIKE '%R61;%') OR (sus.clasificacion_3 LIKE '%R61 %') OR (sus.clasificacion_3 LIKE '%R62') OR (sus.clasificacion_3 LIKE '%R62;%') OR (sus.clasificacion_3 LIKE '%R62 %') OR (sus.clasificacion_3 LIKE '%R63') OR (sus.clasificacion_3 LIKE '%R63;%') OR (sus.clasificacion_3 LIKE '%R63 %') OR (sus.clasificacion_4 LIKE '%R60') OR (sus.clasificacion_4 LIKE '%R60;%') OR (sus.clasificacion_4 LIKE '%R60 %') OR (sus.clasificacion_4 LIKE '%R61') OR (sus.clasificacion_4 LIKE '%R61;%') OR (sus.clasificacion_4 LIKE '%R61 %') OR (sus.clasificacion_4 LIKE '%R62') OR (sus.clasificacion_4 LIKE '%R62;%') OR (sus.clasificacion_4 LIKE '%R62 %') OR (sus.clasificacion_4 LIKE '%R63') OR (sus.clasificacion_4 LIKE '%R63;%') OR (sus.clasificacion_4 LIKE '%R63 %') OR (sus.clasificacion_5 LIKE '%R60') OR (sus.clasificacion_5 LIKE '%R60;%') OR (sus.clasificacion_5 LIKE '%R60 %') OR (sus.clasificacion_5 LIKE '%R61') OR (sus.clasificacion_5 LIKE '%R61;%') OR (sus.clasificacion_5 LIKE '%R61 %') OR (sus.clasificacion_5 LIKE '%R62') OR (sus.clasificacion_5 LIKE '%R62;%') OR (sus.clasificacion_5 LIKE '%R62 %') OR (sus.clasificacion_5 LIKE '%R63') OR (sus.clasificacion_5 LIKE '%R63;%') OR (sus.clasificacion_5 LIKE '%R63 %') OR (sus.clasificacion_6 LIKE '%R60') OR (sus.clasificacion_6 LIKE '%R60;%') OR (sus.clasificacion_6 LIKE '%R60 %') OR (sus.clasificacion_6 LIKE '%R61') OR (sus.clasificacion_6 LIKE '%R61;%') OR (sus.clasificacion_6 LIKE '%R61 %') OR (sus.clasificacion_6 LIKE '%R62') OR (sus.clasificacion_6 LIKE '%R62;%') OR (sus.clasificacion_6 LIKE '%R62 %') OR (sus.clasificacion_6 LIKE '%R63') OR (sus.clasificacion_6 LIKE '%R63;%') OR (sus.clasificacion_6 LIKE '%R63 %') OR (sus.clasificacion_7 LIKE '%R60') OR (sus.clasificacion_7 LIKE '%R60;%') OR (sus.clasificacion_7 LIKE '%R60 %') OR (sus.clasificacion_7 LIKE '%R61') OR (sus.clasificacion_7 LIKE '%R61;%') OR (sus.clasificacion_7 LIKE '%R61 %') OR (sus.clasificacion_7 LIKE '%R62') OR (sus.clasificacion_7 LIKE '%R62;%') OR (sus.clasificacion_7 LIKE '%R62 %') OR (sus.clasificacion_7 LIKE '%R63') OR (sus.clasificacion_7 LIKE '%R63;%') OR (sus.clasificacion_7 LIKE '%R63 %') OR (sus.clasificacion_8 LIKE '%R60') OR (sus.clasificacion_8 LIKE '%R60;%') OR (sus.clasificacion_8 LIKE '%R60 %') OR (sus.clasificacion_8 LIKE '%R61') OR (sus.clasificacion_8 LIKE '%R61;%') OR (sus.clasificacion_8 LIKE '%R61 %') OR (sus.clasificacion_8 LIKE '%R62') OR (sus.clasificacion_8 LIKE '%R62;%') OR (sus.clasificacion_8 LIKE '%R62 %') OR (sus.clasificacion_8 LIKE '%R63') OR (sus.clasificacion_8 LIKE '%R63;%') OR (sus.clasificacion_8 LIKE '%R63 %') OR (sus.clasificacion_9 LIKE '%R60') OR (sus.clasificacion_9 LIKE '%R60;%') OR (sus.clasificacion_9 LIKE '%R60 %') OR (sus.clasificacion_9 LIKE '%R61') OR (sus.clasificacion_9 LIKE '%R61;%') OR (sus.clasificacion_9 LIKE '%R61 %') OR (sus.clasificacion_9 LIKE '%R62') OR (sus.clasificacion_9 LIKE '%R62;%') OR (sus.clasificacion_9 LIKE '%R62 %') OR (sus.clasificacion_9 LIKE '%R63') OR (sus.clasificacion_9 LIKE '%R63;%') OR (sus.clasificacion_9 LIKE '%R63 %') OR (sus.clasificacion_10 LIKE '%R60') OR (sus.clasificacion_10 LIKE '%R60;%') OR (sus.clasificacion_10 LIKE '%R60 %') OR (sus.clasificacion_10 LIKE '%R61') OR (sus.clasificacion_10 LIKE '%R61;%') OR (sus.clasificacion_10 LIKE '%R61 %') OR (sus.clasificacion_10 LIKE '%R62') OR (sus.clasificacion_10 LIKE '%R62;%') OR (sus.clasificacion_10 LIKE '%R62 %') OR (sus.clasificacion_10 LIKE '%R63') OR (sus.clasificacion_10 LIKE '%R63;%') OR (sus.clasificacion_10 LIKE '%R63 %') OR (sus.clasificacion_11 LIKE '%R60') OR (sus.clasificacion_11 LIKE '%R60;%') OR (sus.clasificacion_11 LIKE '%R60 %') OR (sus.clasificacion_11 LIKE '%R61') OR (sus.clasificacion_11 LIKE '%R61;%') OR (sus.clasificacion_11 LIKE '%R61 %') OR (sus.clasificacion_11 LIKE '%R62') OR (sus.clasificacion_11 LIKE '%R62;%') OR (sus.clasificacion_11 LIKE '%R62 %') OR (sus.clasificacion_11 LIKE '%R63') OR (sus.clasificacion_11 LIKE '%R63;%') OR (sus.clasificacion_11 LIKE '%R63 %') OR (sus.clasificacion_12 LIKE '%R60') OR (sus.clasificacion_12 LIKE '%R60;%') OR (sus.clasificacion_12 LIKE '%R60 %') OR (sus.clasificacion_12 LIKE '%R61') OR (sus.clasificacion_12 LIKE '%R61;%') OR (sus.clasificacion_12 LIKE '%R61 %') OR (sus.clasificacion_12 LIKE '%R62') OR (sus.clasificacion_12 LIKE '%R62;%') OR (sus.clasificacion_12 LIKE '%R62 %') OR (sus.clasificacion_12 LIKE '%R63') OR (sus.clasificacion_12 LIKE '%R63;%') OR (sus.clasificacion_12 LIKE '%R63 %') OR (sus.clasificacion_13 LIKE '%R60') OR (sus.clasificacion_13 LIKE '%R60;%') OR (sus.clasificacion_13 LIKE '%R60 %') OR (sus.clasificacion_13 LIKE '%R61') OR (sus.clasificacion_13 LIKE '%R61;%') OR (sus.clasificacion_13 LIKE '%R61 %') OR (sus.clasificacion_13 LIKE '%R62') OR (sus.clasificacion_13 LIKE '%R62;%') OR (sus.clasificacion_13 LIKE '%R62 %') OR (sus.clasificacion_13 LIKE '%R63') OR (sus.clasificacion_13 LIKE '%R63;%') OR (sus.clasificacion_13 LIKE '%R63 %') OR (sus.clasificacion_14 LIKE '%R60') OR (sus.clasificacion_14 LIKE '%R60;%') OR (sus.clasificacion_14 LIKE '%R60 %') OR (sus.clasificacion_14 LIKE '%R61') OR (sus.clasificacion_14 LIKE '%R61;%') OR (sus.clasificacion_14 LIKE '%R61 %') OR (sus.clasificacion_14 LIKE '%R62') OR (sus.clasificacion_14 LIKE '%R62;%') OR (sus.clasificacion_14 LIKE '%R62 %') OR (sus.clasificacion_14 LIKE '%R63') OR (sus.clasificacion_14 LIKE '%R63;%') OR (sus.clasificacion_14 LIKE '%R63 %') OR (sus.clasificacion_15 LIKE '%R60') OR (sus.clasificacion_15 LIKE '%R60;%') OR (sus.clasificacion_15 LIKE '%R60 %') OR (sus.clasificacion_15 LIKE '%R61') OR (sus.clasificacion_15 LIKE '%R61;%') OR (sus.clasificacion_15 LIKE '%R61 %') OR (sus.clasificacion_15 LIKE '%R62') OR (sus.clasificacion_15 LIKE '%R62;%') OR (sus.clasificacion_15 LIKE '%R62 %') OR (sus.clasificacion_15 LIKE '%R63') OR (sus.clasificacion_15 LIKE '%R63;%') OR (sus.clasificacion_15 LIKE '%R63 %'))) OR sus.id IN (select distinct sus.id from dn_risc_sustancias as sus WHERE ((sus.frases_r_danesa LIKE '%R60') OR (sus.frases_r_danesa LIKE '%R60;%') OR (sus.frases_r_danesa LIKE '%R60 %') OR (sus.frases_r_danesa LIKE '%R61') OR (sus.frases_r_danesa LIKE '%R61;%') OR (sus.frases_r_danesa LIKE '%R61 %') OR (sus.frases_r_danesa LIKE '%R62') OR (sus.frases_r_danesa LIKE '%R62;%') OR (sus.frases_r_danesa LIKE '%R62 %') OR (sus.frases_r_danesa LIKE '%R63') OR (sus.frases_r_danesa LIKE '%R63;%') OR (sus.frases_r_danesa LIKE '%R63 %'))) OR sus.id in (select distinct sus.id from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_mama_cop ON (sus.id=dn_risc_sustancias_mama_cop.id_sustancia) WHERE (dn_risc_sustancias_mama_cop.cancer_mama=1)) OR sus.id IN (select distinct sus.id from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_mama_cop ON (sus.id=dn_risc_sustancias_mama_cop.id_sustancia) WHERE (dn_risc_sustancias_mama_cop.cop<>'')))"

	  sql = sql & " FULL OUTER JOIN dn_risc_sustancias_iarc as iarc ON (sus.id=iarc.id_sustancia)"
	  sql = sql & " FULL OUTER JOIN dn_risc_sustancias_cancer_otras as caoc ON (sus.id=caoc.id_sustancia)"
	  sql = sql & " FULL OUTER JOIN dn_risc_sustancias_neuro_disruptor as neuro ON (sus.id=neuro.id_sustancia)"

    case else:
      sql = "SELECT TOP 100 id FROM dn_risc_sustancias WHERE 1=0"

  end select




  ' AÑADIMOS LAS CONDICIONES DE LOS GRUPOS
		'según filtro, agregamos distintas condiciones
		if filtro<>"0" then
'			sql=sql & " OR "
			if InStr(sql, "WHERE")>0 then
				sql=sql & " AND "
			else
				sql=sql & " WHERE "
			end if

			select case listado

				case "todas": 'la sustancia (o el grupo al que pertenece) debe tener un uso toxico, o debe existir una alternativa
				  sql=sql & " (dn_risc_sustancias_por_usos.toxico=1 OR dn_alter_ficheros_por_sustancias.id_fichero is not null OR dn_risc_grupos_por_usos.toxico=1 OR dn_alter_ficheros_por_grupos.id_fichero is not null) "

 				case "cym":
          campos="sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15"
          frases="R40, R45, R49, R40/20, R40/21, R40/22, R40/20/21, R40/20/22, R40/21/22, R40/20/21/22, R46, R68, R68/20, R68/21, R68/22, R68/20/21, R68/20/22, R68/21/22, R68/20/21/22"

				  sql=sql & "( " & monta_condicion(campos, frases) & " OR ("&monta_condicion_grupo("asoc_cancer_rd")&") )"

				case "cym2": 'la condicion simplemente es que exista la fila, pero ya que estamos, comprobamos que no este vacía
				  sql=sql & "( dn_risc_sustancias_iarc.grupo_iarc<>'GRUPO 4' and ((dn_risc_sustancias_iarc.grupo_iarc<>'')) OR ("&monta_condicion_grupo("asoc_cancer_iarc")&") )"
				  'Sergio
				  'sql=sql & "( ((dn_risc_sustancias_iarc.grupo_iarc<>'') OR ("&monta_condicion_grupo("asoc_cancer_iarc")&")) and dn_risc_sustancias_iarc.grupo_iarc<>'GRUPO 4')"



				case "cym3": 'la condicion simplemente es que exista la fila, pero ya que estamos, comprobamos que no este vacía
				  'sql=sql & "  ( ((dn_risc_sustancias_cancer_otras.categoria_cancer_otras<>'')) OR ("&monta_condicion_grupo("asoc_cancer_otras")&") )"
				  sql=sql & " (not(dn_risc_sustancias_cancer_otras.fuente like '%ACGIH%' and (dn_risc_sustancias_cancer_otras.categoria_cancer_otras like '%G-A5%' or dn_risc_sustancias_cancer_otras.categoria_cancer_otras like '%G-A4%' ) ) or (dn_risc_sustancias_cancer_otras.categoria_cancer_otras is null))  and ( ((dn_risc_sustancias_cancer_otras.categoria_cancer_otras<>'')) OR ("&monta_condicion_grupo("asoc_cancer_otras")&") )"

				case "mama": 'que cancer_mama sea 1
				  sql=sql & "( ((dn_risc_sustancias_mama_cop.cancer_mama=1)) OR ("&monta_condicion_grupo("asoc_cancer_mama")&") )"

				case "cop": 'que cop no sea vacío
				  sql=sql & "( ((dn_risc_sustancias_mama_cop.cop <> '')) OR ("&monta_condicion_grupo("asoc_cop")&") )"

				case "tpr":
          ' Buscando las frases R: TPR R60, R61, R62, R63, en las columnas CLASIFICACION_1, hasta CLASIFICACION_6 del RD 363/1995.

          campos="sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15"
          frases="R60, R61, R62, R63"

				  sql=sql & "( " & monta_condicion(campos, frases) & " OR ("&monta_condicion_grupo("asoc_reproduccion")&") )"

				case "dis": 'nivel_disruptor no esta vacio
				  sql=sql & "( ((dn_risc_sustancias_neuro_disruptor.nivel_disruptor<>'')) OR ("&monta_condicion_grupo("asoc_disruptores")&") )"

				case "neu": 'en campos adicionales nivel_neurotoxico no esta vacio
				  'sql=sql & " ((dn_risc_sustancias_neuro_disruptor.nivel_neurotoxico<>'')) "
          sql = sql & "( " & sql_lista_neurotoxico & " OR ("&monta_condicion_grupo("asoc_neuro_oto")&") )"


		  		'Sergio
		  		case "oto":
          		'Comentar con Xavi
				'sql = sql & " dn_risc_sustancias_neuro_disruptor.efecto_neurotoxico='OTOTÓXICO' and ( " & sql_lista_neurotoxico & " OR ("&monta_condicion_grupo("asoc_neuro_oto")&") )"
				sql = sql & " dn_risc_sustancias_neuro_disruptor.efecto_neurotoxico='OTOTÓXICO'" ' and ( " & sql_lista_neurotoxico & " OR ("&monta_condicion_grupo("asoc_neuro_oto")&") )"


				case "sen": 'determinadas frases R en clasif_1 a clasif_15 y frases_r_danesa
          campos="sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15, sus.frases_r_danesa"
          frases = "R42, R43, R42/43, R42-43"
          sql=sql & monta_condicion(campos, frases)
				case "senreach":
				  sql=sql & " ((dn_risc_sensibilizantes_reach.id_sustancia<>'')) OR ("&monta_condicion_grupo("asoc_alergenos")&")"
				case "pyb":
				  sql=sql & " ((dn_risc_sustancias_ambiente.anchor_tpb<>''))  OR ("&monta_condicion_grupo("asoc_tpb")&")"

				case "tac":
				  sql=sql & "( ((dn_risc_sustancias_ambiente.directiva_aguas=1)) OR ("&monta_condicion_grupo("asoc_directiva_aguas")&") )"

				case "tac2":
				  sql=sql & " ((dn_risc_sustancias_ambiente.clasif_MMA<>'' and dn_risc_sustancias_ambiente.clasif_MMA<>'nwg')) OR ("&monta_condicion_grupo("asoc_peligrosas_agua_alemania")&")"

				case "dat":
				  sql=sql & " ((dn_risc_sustancias_ambiente.dano_ozono=1)) OR ("&monta_condicion_grupo("asoc_capa_ozono")&")"

				case "dat2":
				  sql=sql & " ((dn_risc_sustancias_ambiente.dano_cambio_clima=1)) OR ("&monta_condicion_grupo("asoc_cambio_climatico")&")"

				case "dat3":
				  sql=sql & "( ((dn_risc_sustancias_ambiente.dano_calidad_aire=1)) OR ("&monta_condicion_grupo("asoc_calidad_aire")&") )"

				case "vl1":
				  sql=sql & "( ((vla_ed_ppm_1<>'') or (vla_ed_mg_m3_1<>'') or (vla_ed_ppm_1<>'') or (vla_ec_mg_m3_1<>''))  OR ("&monta_condicion_grupo("asoc_vla")&") )"

				case "vl2":
          campos="sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15"
          frases = "R40, R45, R49, R40/20, R40/21, R40/22, R40/20/22, R46, R68, R68/20, R68/21, R68/22, R68/20/21, R68/21/22, R68/20/21/22"
				  sql=sql & "( ((vl.vla_ed_ppm_1 <> '') OR (vl.vla_ed_mg_m3_1 <> '') OR (vl.vla_ec_ppm_1 <> '') OR (vl.vla_ec_mg_m3_1 <> '') OR (vl.vla_ed_ppm_2 <> '') OR (vl.vla_ed_mg_m3_2 <> '') OR (vl.vla_ec_ppm_2 <> '') OR (vl.vla_ec_mg_m3_2 <> '') OR (vl.vla_ed_ppm_3 <> '') OR (vl.vla_ed_mg_m3_3 <> '') OR (vl.vla_ec_ppm_3 <> '') OR (vl.vla_ec_mg_m3_3 <> '') OR (vl.vla_ed_ppm_4 <> '') OR (vl.vla_ed_mg_m3_4 <> '') OR (vl.vla_ec_ppm_4 <> '') OR (vl.vla_ec_mg_m3_4 <> '') OR (vl.vla_ed_ppm_5 <> '') OR (vl.vla_ed_mg_m3_5 <> '') OR (vl.vla_ec_ppm_5 <> '') OR (vl.vla_ec_mg_m3_5 <> '') OR (vl.vla_ed_ppm_6 <> '') OR (vl.vla_ed_mg_m3_6 <> '') OR (vl.vla_ec_ppm_6 <> '') OR (vl.vla_ec_mg_m3_6 <> '')) and ("&monta_condicion(campos, frases)&")  OR ("&monta_condicion_grupo("asoc_vla")&") )"

				case "vl3":
				  sql=sql & "( ((vlb_1<>''))  OR ("&monta_condicion_grupo("asoc_vlb")&") )"

				case "enf": 'la condicion simplemente es que exista la fila, pero ya que estamos, comprobamos que no este vacía
				  sql=sql & "( ((sus.id<>'')) OR ("&monta_condicion_grupo("asoc_enfermedades")&") )"

				case "emi":
				  sql=sql & "( ((dn_risc_sustancias_ambiente.emisiones_atmosfera=1)) OR ("&monta_condicion_grupo("asoc_emisiones_atmosfericas")&") )"

				case "res": 'la condicion es que exista la fila O QUE que num_rd <>'' / como tenemos full outer, añadimos la condicion de que la tabla de sustancias no este vacia, por si acaso
				  sql=sql & " ((sus.num_rd<>'' or sus.num_rd <> '' or dn_risc_sustancias_ambiente.id is not null or dn_risc_sustancias_cancer_otras.id is not null or dn_risc_sustancias_iarc.id is not null or dn_risc_sustancias_neuro_disruptor.id is not null or dn_risc_sustancias_vl.id is not null ) AND sus.id is not null) "

				case "ver":
				  sql=sql & " ((sus.num_rd <> '') OR (sus.frases_r_danesa <> '') OR (iarc.grupo_iarc <> '') OR (otras.categoria_cancer_otras <> '') OR (neuro.nivel_disruptor <> '') OR (ambiente.enlace_tpb <> '') OR (ambiente.directiva_aguas <> '') OR (ambiente.clasif_mma <> '')) "

				case "cov":
				  sql=sql & " ((dn_risc_sustancias_ambiente.cov=1)) OR ("&monta_condicion_grupo("asoc_cov")&")"

				case "lpc":
				  sql=sql & "( ((eper_agua<>'' or eper_aire<>'' or eper_suelo<>'')) OR ("&monta_condicion_grupo("asoc_eper")&") )"
				case "ep1":
				sql=sql & "(eper_agua<>'')"

				case "ep2":
				sql=sql & "(eper_aire<>'')"

				case "ep3":
				sql=sql & "(eper_suelo<>'')"

				case "mpmb":
				sql=sql & "(sus.num_cas='87-68-3' or sus.num_cas='133-49-3' or sus.num_cas='75-74-1') OR ("&monta_condicion_grupo("asoc_mpmb")&")"

				case "pro":
				sql=sql & "(sustancia_prohibida=1) OR ("&monta_condicion_grupo("asoc_prohibidas")&")"

				case "rest":
				sql=sql & "(sustancia_restringida=1) OR ("&monta_condicion_grupo("asoc_restringidas")&")"


				case "acm":
				  sql=sql & "( ((seveso<>'')) OR ("&monta_condicion_grupo("asoc_seveso")&") )"

				case "cos":
				  sql=sql & "( dn_risc_sustancias_ambiente.toxicidad_suelo=1 ) OR ("&monta_condicion_grupo("asoc_contaminantes_suelo")&")"

				case "anexo_reach":
				  sql=sql & " (dn_risc_sustancias_por_usos.anexo_reach=1)"

        case "negra": 'Lista negra
				  'sql = sql & " ((sus.clasificacion_1 like '%33%' or sus.clasificacion_2 like '%33%' or sus.clasificacion_3 like '%33%' or sus.clasificacion_4 like '%33%' or sus.clasificacion_5 like '%33%' or sus.clasificacion_6 like '%33%' or sus.clasificacion_7 like '%33%' or sus.clasificacion_8 like '%33%' or sus.clasificacion_9 like '%33%' or sus.clasificacion_10 like '%33%' or sus.clasificacion_11 like '%33%' or sus.clasificacion_12 like '%33%' or sus.clasificacion_13 like '%33%' or sus.clasificacion_14 like '%33%' or sus.clasificacion_15 like '%33%')"
				  'sql = sql & " OR (sus.clasificacion_1 like '%R53%' or sus.clasificacion_2 like '%R53%' or sus.clasificacion_3 like '%R53%' or sus.clasificacion_4 like '%R53%' or sus.clasificacion_5 like '%R53%' or sus.clasificacion_6 like '%R53%' or sus.clasificacion_7 like '%R53%' or sus.clasificacion_8 like '%R53%' or sus.clasificacion_9 like '%R53%' or sus.clasificacion_10 like '%R53%' or sus.clasificacion_11 like '%R53%' or sus.clasificacion_12 like '%R53%' or sus.clasificacion_13 like '%R53%' or sus.clasificacion_14 like '%R53%' or sus.clasificacion_15 like '%R53%')"
				  'sql = sql & " OR (sus.clasificacion_1 like '%58%' or sus.clasificacion_2 like '%58%' or sus.clasificacion_3 like '%58%' or sus.clasificacion_4 like '%58%' or sus.clasificacion_5 like '%58%' or sus.clasificacion_6 like '%58%' or sus.clasificacion_7 like '%58%' or sus.clasificacion_8 like '%58%' or sus.clasificacion_9 like '%58%' or sus.clasificacion_10 like '%58%' or sus.clasificacion_11 like '%58%' or sus.clasificacion_12 like '%58%' or sus.clasificacion_13 like '%58%' or sus.clasificacion_14 like '%58%' or sus.clasificacion_15 like '%58%'))"

				  campos = "sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15"
          		  frases = "R33, R53, R58, R50-53, R51-53, R52-53"
          		  sql = sql & "(" & monta_condicion(campos, frases)
				  sql = sql & " or ((negra=1)"
				  sql = sql & " and (iarc.grupo_iarc<>'GRUPO 3' or iarc.grupo_iarc is null) and (iarc.grupo_iarc<>'GRUPO 4' or iarc.grupo_iarc is null)"
				  sql = sql & " and (not(caoc.fuente like '%ACGIH%' and (caoc.categoria_cancer_otras like '%G-A5%' or caoc.categoria_cancer_otras like '%G-A4%')) or (caoc.categoria_cancer_otras is null))"
				  sql = sql & " and not (sus.clasificacion_1 like '%67%' or sus.clasificacion_2 like '%67%' or sus.clasificacion_3 like '%67%' or sus.clasificacion_4 like '%67%' or sus.clasificacion_5 like '%67%' or sus.clasificacion_6 like '%67%' or sus.clasificacion_7 like '%67%' or sus.clasificacion_8 like '%67%' or sus.clasificacion_9 like '%67%' or sus.clasificacion_10 like '%67%' or sus.clasificacion_11 like '%67%' or sus.clasificacion_12 like '%67%' or sus.clasificacion_13 like '%67%' or sus.clasificacion_14 like '%67%' or sus.clasificacion_15 like '%67%')))"


			end select

		end if











  dame_sql_busqueda = sql
end function

%>
