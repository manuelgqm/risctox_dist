<!--#include file="dn_fun_comunes.asp"-->
<!--#include file="dn_auten.inc"-->
<!--#include file="adovbs.inc"-->
<!--#include file="dn_conexion.asp"-->

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
<div id="box2" class="contenido">
<div id="navegacion" align="right">
<strong>[<a href="javascript:close();">Cerrar</a>] [<a href="javascript:back();">Volver atrás</a>]</strong>
</div>

<%
tipo=request("tipo")
id=request("id")

select case tipo
  case "sustancia":
%>
  <h1>Sustancia: <%= dame_campo ("nombre", "dn_risc_sustancias", id) %></h1>
  <p>
    <strong>ID:</strong> <%= id %><br/>
  </p>

  <% muestra_grupos_sustancia(id) %>
  <% muestra_usos_sustancia(id) %>
  <% muestra_companias_sustancia(id) %>
  <% muestra_sectores_sustancia(id) %>
  <% muestra_enfermedades_sustancia(id) %>
  <% muestra_ficheros_sustancia(id) %>

<%
  case "grupo":
%>

  <h1>Grupo: <%= dame_campo ("nombre", "dn_risc_grupos", id) %></h1>
  <p>
    <strong>ID:</strong> <%= id %><br/>
  </p>

  <% muestra_sustancias_grupo(id) %>
  <% muestra_usos_grupo(id) %>
  <% muestra_enfermedades_grupo(id) %>
  <% muestra_ficheros_grupo(id) %>

<%
  case "uso":
%>

  <h1>Uso: <%= dame_campo ("nombre", "dn_risc_usos", id) %></h1>
  <p>
    <strong>ID:</strong> <%= id %><br/>
  </p>

  <% muestra_sustancias_uso(id) %>
  <% muestra_grupos_uso(id) %>
  <% muestra_ficheros_uso(id) %>

<%
  case "compania":
%>

  <h1>Compañía: <%= dame_campo ("nombre", "dn_risc_companias", id) %></h1>
  <p>
    <strong>ID:</strong> <%= id %><br/>
  </p>

  <% muestra_sustancias_compania(id) %>

<%
  case "enfermedad":
%>

  <h1>Enfermedad: <%= dame_campo ("nombre", "dn_risc_enfermedades", id) %></h1>
  <p>
    <strong>ID:</strong> <%= id %><br/>
  </p>

  <% muestra_sustancias_enfermedad(id) %>
  <% muestra_grupos_enfermedad(id) %>

<%
  case "fichero":
%>

  <h1>Fichero: <%= dame_campo ("titulo", "dn_alter_ficheros", id) %></h1>
  <p>
    <strong>ID:</strong> <%= id %><br/>
  </p>

  <% muestra_sustancias_fichero(id) %>
  <% muestra_grupos_fichero(id) %>
  <% muestra_sectores_fichero(id) %>
  <% muestra_procesos_fichero(id) %>
  <% muestra_usos_fichero(id) %>

<%
  case "sector":
%>

  <h1>Sector: <%= dame_campo ("nombre", "dn_alter_sectores", id) %></h1>
  <p>
    <strong>ID:</strong> <%= id %><br/>
  </p>

  <% muestra_ficheros_sector(id) %>

<%
  case "proceso":
%>

  <h1>Proceso: <%= dame_campo ("nombre", "dn_alter_procesos", id) %></h1>
  <p>
    <strong>ID:</strong> <%= id %><br/>
  </p>

  <% muestra_ficheros_proceso(id) %>

<%
  case "uso":
%>

  <h1>Uso: <%= dame_campo ("nombre", "dn_risc_usos", id) %></h1>
  <p>
    <strong>ID:</strong> <%= id %><br/>
  </p>

  <% muestra_sustancias_uso(id) %>
  <% muestra_grupos_uso(id) %>
  <% muestra_ficheros_uso(id) %>

<%
  case "residuo":
%>

  <h1>Residuo: <%= dame_campo ("nombre", "rq_residuos", id) %></h1>
  <p>
    <strong>ID:</strong> <%= id %><br/>
  </p>

  <% muestra_ficheros_residuo(id) %>
  
<%
  case else:
%>
  <h1>No se reconoce el tipo</h1>
<%
end select
%>

</div>
</body>
</html>

<%
	cerrarconexion
%>

<%
function dame_campo(byval columna, byval tabla, byval id)
  sql = "SELECT "&columna&" AS valor FROM "&tabla&" WHERE id="&id
  set obj_rst=objconn1.execute(sql)
  if (not obj_rst.eof) then
    valor = obj_rst("valor")
  else
    valor = ""
  end if
  obj_rst.close()
  set obj_rst=nothing

  dame_campo = valor
end function

' ##################################################################
sub muestra_grupos_sustancia(byval id)
  sql = "SELECT * FROM dn_risc_sustancias_por_grupos AS spg INNER JOIN dn_risc_grupos AS g ON spg.id_grupo = g.id WHERE id_sustancia="&id
  set obj_rst = objconn1.execute(sql)
  if (obj_rst.eof) then
%>
    <h2>No asociada a grupos</h2>
<%
  else
%>
    <h2>Asociada a los grupos...</h2>
    <ul>
<%
    do while(not obj_rst.eof)
      response.write("<li><a href='dn_asociaciones_mostrar.asp?tipo=grupo&id="&obj_rst("id_grupo")&"'>"&obj_rst("nombre")&"</a></li>")
      obj_rst.movenext
    loop
  end if
%>
    </ul>
<%

  obj_rst.close()
  set obj_rst = nothing
end sub

' ##################################################################
sub muestra_usos_sustancia(byval id)
  sql = "SELECT * FROM dn_risc_sustancias_por_usos AS spu INNER JOIN dn_risc_usos AS u ON spu.id_uso = u.id WHERE id_sustancia="&id
  set obj_rst = objconn1.execute(sql)
  if (obj_rst.eof) then
%>
    <h2>No asociada a usos</h2>
<%
  else
%>
    <h2>Asociada a los usos...</h2>
    <ul>
<%
    do while(not obj_rst.eof)
      response.write("<li><a href='dn_asociaciones_mostrar.asp?tipo=uso&id="&obj_rst("id_uso")&"'>"&obj_rst("nombre")&"</a></li>")
      obj_rst.movenext
    loop
  end if
%>
    </ul>
<%

  obj_rst.close()
  set obj_rst = nothing
end sub

' ##################################################################
sub muestra_companias_sustancia(byval id)
  sql = "SELECT * FROM dn_risc_sustancias_por_companias AS spc INNER JOIN dn_risc_companias AS c ON spc.id_compania = c.id WHERE id_sustancia="&id
  set obj_rst = objconn1.execute(sql)
  if (obj_rst.eof) then
%>
    <h2>No asociada a compañías</h2>
<%
  else
%>
    <h2>Asociada a las compañías...</h2>
    <ul>
<%
    do while(not obj_rst.eof)
      response.write("<li><a href='dn_asociaciones_mostrar.asp?tipo=compania&id="&obj_rst("id_compania")&"'>"&obj_rst("nombre")&"</a></li>")
      obj_rst.movenext
    loop
  end if
%>
    </ul>
<%

  obj_rst.close()
  set obj_rst = nothing
end sub

' ##################################################################
sub muestra_sectores_sustancia(byval id)
  sql = "SELECT * FROM dn_risc_sustancias_por_sectores AS sps INNER JOIN dn_alter_sectores AS s ON sps.id_sector = s.id WHERE id_sustancia="&id
  set obj_rst = objconn1.execute(sql)
  if (obj_rst.eof) then
%>
    <h2>No asociada a sectores</h2>
<%
  else
%>
    <h2>Asociada a los sectores...</h2>
    <ul>
<%
    do while(not obj_rst.eof)
      response.write("<li><a href='dn_asociaciones_mostrar.asp?tipo=sector&id="&obj_rst("id_sector")&"'>"&obj_rst("numero_cnae")&" - "&obj_rst("nombre")&"</a></li>")
      obj_rst.movenext
    loop
  end if
%>
    </ul>
<%

  obj_rst.close()
  set obj_rst = nothing
end sub


' ##################################################################
sub muestra_enfermedades_sustancia(byval id)
  sql = "SELECT * FROM dn_risc_sustancias_por_enfermedades AS spe INNER JOIN dn_risc_enfermedades AS e ON spe.id_enfermedad = e.id WHERE id_sustancia="&id
  set obj_rst = objconn1.execute(sql)
  if (obj_rst.eof) then
%>
    <h2>No asociada a enfermedades</h2>
<%
  else
%>
    <h2>Asociada a las enfermedades...</h2>
    <ul>
<%
    do while(not obj_rst.eof)
      response.write("<li><a href='dn_asociaciones_mostrar.asp?tipo=enfermedad&id="&obj_rst("id_enfermedad")&"'>"&obj_rst("nombre")&"</a></li>")
      obj_rst.movenext
    loop
  end if
%>
    </ul>
<%

  obj_rst.close()
  set obj_rst = nothing
end sub

' ##################################################################
sub muestra_ficheros_sustancia(byval id)
  sql = "SELECT * FROM dn_alter_ficheros_por_sustancias AS fps INNER JOIN dn_alter_ficheros AS f ON fps.id_fichero = f.id WHERE id_sustancia="&id
  set obj_rst = objconn1.execute(sql)
  if (obj_rst.eof) then
%>
    <h2>No asociada a ficheros</h2>
<%
  else
%>
    <h2>Asociada a los ficheros...</h2>
    <ul>
<%
    do while(not obj_rst.eof)
      response.write("<li><a href='dn_asociaciones_mostrar.asp?tipo=fichero&id="&obj_rst("id_fichero")&"'>"&obj_rst("titulo")&"</a></li>")
      obj_rst.movenext
    loop
  end if
%>
    </ul>
<%

  obj_rst.close()
  set obj_rst = nothing
end sub

' ##################################################################
sub muestra_sustancias_grupo(byval id)
  sql = "SELECT * FROM dn_risc_sustancias_por_grupos AS spg INNER JOIN dn_risc_sustancias AS s ON spg.id_sustancia = s.id WHERE id_grupo="&id
  set obj_rst = objconn1.execute(sql)
  if (obj_rst.eof) then
%>
    <h2>No asociado a sustancias</h2>
<%
  else
%>
    <h2>Asociado a las sustancias...</h2>
    <ul>
<%
    do while(not obj_rst.eof)
      response.write("<li><a href='dn_asociaciones_mostrar.asp?tipo=sustancia&id="&obj_rst("id_sustancia")&"'>"&obj_rst("nombre")&"</a></li>")
      obj_rst.movenext
    loop
  end if
%>
    </ul>
<%

  obj_rst.close()
  set obj_rst = nothing
end sub

' ##################################################################
sub muestra_usos_grupo(byval id)
  sql = "SELECT * FROM dn_risc_grupos_por_usos AS gpu INNER JOIN dn_risc_usos AS u ON gpu.id_uso = u.id WHERE id_grupo="&id
  set obj_rst = objconn1.execute(sql)
  if (obj_rst.eof) then
%>
    <h2>No asociado a usos</h2>
<%
  else
%>
    <h2>Asociado a los usos...</h2>
    <ul>
<%
    do while(not obj_rst.eof)
      response.write("<li><a href='dn_asociaciones_mostrar.asp?tipo=uso&id="&obj_rst("id_uso")&"'>"&obj_rst("nombre")&"</a></li>")
      obj_rst.movenext
    loop
  end if
%>
    </ul>
<%

  obj_rst.close()
  set obj_rst = nothing
end sub

' ##################################################################

sub muestra_enfermedades_grupo(byval id)
  sql = "SELECT * FROM dn_risc_grupos_por_enfermedades AS gpe INNER JOIN dn_risc_enfermedades AS e ON gpe.id_enfermedad = e.id WHERE id_grupo="&id
  set obj_rst = objconn1.execute(sql)
  if (obj_rst.eof) then
%>
    <h2>No asociado a enfermedades</h2>
<%
  else
%>
    <h2>Asociado a las enfermedades...</h2>
    <ul>
<%
    do while(not obj_rst.eof)
      response.write("<li><a href='dn_asociaciones_mostrar.asp?tipo=enfermedad&id="&obj_rst("id_enfermedad")&"'>"&obj_rst("nombre")&"</a></li>")
      obj_rst.movenext
    loop
  end if
%>
    </ul>
<%

  obj_rst.close()
  set obj_rst = nothing
end sub

' ##################################################################
sub muestra_ficheros_grupo(byval id)
  sql = "SELECT * FROM dn_alter_ficheros_por_grupos AS fpg INNER JOIN dn_alter_ficheros AS f ON fpg.id_fichero = f.id WHERE id_grupo="&id
  set obj_rst = objconn1.execute(sql)
  if (obj_rst.eof) then
%>
    <h2>No asociado a ficheros</h2>
<%
  else
%>
    <h2>Asociado a los ficheros...</h2>
    <ul>
<%
    do while(not obj_rst.eof)
      response.write("<li><a href='dn_asociaciones_mostrar.asp?tipo=fichero&id="&obj_rst("id_fichero")&"'>"&obj_rst("titulo")&"</a></li>")
      obj_rst.movenext
    loop
  end if
%>
    </ul>
<%

  obj_rst.close()
  set obj_rst = nothing
end sub

' ##################################################################
sub muestra_sustancias_uso(byval id)
  sql = "SELECT * FROM dn_risc_sustancias_por_usos AS spu INNER JOIN dn_risc_sustancias AS s ON spu.id_sustancia = s.id WHERE id_uso="&id
  set obj_rst = objconn1.execute(sql)
  if (obj_rst.eof) then
%>
    <h2>No asociado a sustancias</h2>
<%
  else
%>
    <h2>Asociado a las sustancias...</h2>
    <ul>
<%
    do while(not obj_rst.eof)
      response.write("<li><a href='dn_asociaciones_mostrar.asp?tipo=sustancia&id="&obj_rst("id_sustancia")&"'>"&obj_rst("nombre")&"</a></li>")
      obj_rst.movenext
    loop
  end if
%>
    </ul>
<%

  obj_rst.close()
  set obj_rst = nothing
end sub

' ##################################################################
sub muestra_grupos_uso(byval id)
  sql = "SELECT * FROM dn_risc_grupos_por_usos AS gpu INNER JOIN dn_risc_grupos AS g ON gpu.id_grupo = g.id WHERE id_uso="&id
  set obj_rst = objconn1.execute(sql)
  if (obj_rst.eof) then
%>
    <h2>No asociado a grupos</h2>
<%
  else
%>
    <h2>Asociado a los grupos...</h2>
    <ul>
<%
    do while(not obj_rst.eof)
      response.write("<li><a href='dn_asociaciones_mostrar.asp?tipo=grupo&id="&obj_rst("id_grupo")&"'>"&obj_rst("nombre")&"</a></li>")
      obj_rst.movenext
    loop
  end if
%>
    </ul>
<%

  obj_rst.close()
  set obj_rst = nothing
end sub

' ##################################################################
sub muestra_ficheros_uso(byval id)
  sql = "SELECT * FROM dn_alter_ficheros_por_usos AS fpu INNER JOIN dn_alter_ficheros AS f ON fpu.id_fichero = f.id WHERE id_uso="&id
  set obj_rst = objconn1.execute(sql)
  if (obj_rst.eof) then
%>
    <h2>No asociado a ficheros</h2>
<%
  else
%>
    <h2>Asociado a los ficheros...</h2>
    <ul>
<%
    do while(not obj_rst.eof)
      response.write("<li><a href='dn_asociaciones_mostrar.asp?tipo=fichero&id="&obj_rst("id_fichero")&"'>"&obj_rst("titulo")&"</a></li>")
      obj_rst.movenext
    loop
  end if
%>
    </ul>
<%

  obj_rst.close()
  set obj_rst = nothing
end sub

' ##################################################################
sub muestra_sustancias_compania(byval id)
  sql = "SELECT * FROM dn_risc_sustancias_por_companias AS spc INNER JOIN dn_risc_sustancias AS s ON spc.id_sustancia = s.id WHERE id_compania="&id
  set obj_rst = objconn1.execute(sql)
  if (obj_rst.eof) then
%>
    <h2>No asociada a sustancias</h2>
<%
  else
%>
    <h2>Asociada a las sustancias...</h2>
    <ul>
<%
    do while(not obj_rst.eof)
      response.write("<li><a href='dn_asociaciones_mostrar.asp?tipo=sustancia&id="&obj_rst("id_sustancia")&"'>"&obj_rst("nombre")&"</a></li>")
      obj_rst.movenext
    loop
  end if
%>
    </ul>
<%

  obj_rst.close()
  set obj_rst = nothing
end sub

' ##################################################################
sub muestra_sustancias_enfermedad(byval id)
  sql = "SELECT * FROM dn_risc_sustancias_por_enfermedades AS spe INNER JOIN dn_risc_sustancias AS s ON spe.id_sustancia = s.id WHERE id_enfermedad="&id
  set obj_rst = objconn1.execute(sql)
  if (obj_rst.eof) then
%>
    <h2>No asociada a sustancias</h2>
<%
  else
%>
    <h2>Asociada a las sustancias...</h2>
    <ul>
<%
    do while(not obj_rst.eof)
      response.write("<li><a href='dn_asociaciones_mostrar.asp?tipo=sustancia&id="&obj_rst("id_sustancia")&"'>"&obj_rst("nombre")&"</a></li>")
      obj_rst.movenext
    loop
  end if
%>
    </ul>
<%

  obj_rst.close()
  set obj_rst = nothing
end sub

' ##################################################################

sub muestra_grupos_enfermedad(byval id)
  sql = "SELECT * FROM dn_risc_grupos_por_enfermedades AS gpe INNER JOIN dn_risc_grupos AS g ON gpe.id_grupo = g.id WHERE id_enfermedad="&id
  set obj_rst = objconn1.execute(sql)
  if (obj_rst.eof) then
%>
    <h2>No asociada a grupos</h2>
<%
  else
%>
    <h2>Asociada a los grupos...</h2>
    <ul>
<%
    do while(not obj_rst.eof)
      response.write("<li><a href='dn_asociaciones_mostrar.asp?tipo=grupo&id="&obj_rst("id_grupo")&"'>"&obj_rst("nombre")&"</a></li>")
      obj_rst.movenext
    loop
  end if
%>
    </ul>
<%

  obj_rst.close()
  set obj_rst = nothing
end sub

' ##################################################################
sub muestra_sustancias_fichero(byval id)
  sql = "SELECT * FROM dn_alter_ficheros_por_sustancias AS fps INNER JOIN dn_risc_sustancias AS s ON fps.id_sustancia = s.id WHERE id_fichero="&id
  set obj_rst = objconn1.execute(sql)
  if (obj_rst.eof) then
%>
    <h2>No asociado a sustancias</h2>
<%
  else
%>
    <h2>Asociado a las sustancias...</h2>
    <ul>
<%
    do while(not obj_rst.eof)
      response.write("<li><a href='dn_asociaciones_mostrar.asp?tipo=sustancia&id="&obj_rst("id_sustancia")&"'>"&obj_rst("nombre")&"</a></li>")
      obj_rst.movenext
    loop
  end if
%>
    </ul>
<%

  obj_rst.close()
  set obj_rst = nothing
end sub

' ##################################################################
sub muestra_grupos_fichero(byval id)
  sql = "SELECT * FROM dn_alter_ficheros_por_grupos AS fpg INNER JOIN dn_risc_grupos AS g ON fpg.id_grupo = g.id WHERE id_fichero="&id
  set obj_rst = objconn1.execute(sql)
  if (obj_rst.eof) then
%>
    <h2>No asociado a grupos</h2>
<%
  else
%>
    <h2>Asociado a los grupos...</h2>
    <ul>
<%
    do while(not obj_rst.eof)
      response.write("<li><a href='dn_asociaciones_mostrar.asp?tipo=grupo&id="&obj_rst("id_grupo")&"'>"&obj_rst("nombre")&"</a></li>")
      obj_rst.movenext
    loop
  end if
%>
    </ul>
<%

  obj_rst.close()
  set obj_rst = nothing
end sub

' ##################################################################
sub muestra_sectores_fichero(byval id)
  sql = "SELECT * FROM dn_alter_ficheros_por_sectores AS fps INNER JOIN dn_alter_sectores AS s ON fps.id_sector = s.id WHERE id_fichero="&id
  set obj_rst = objconn1.execute(sql)
  if (obj_rst.eof) then
%>
    <h2>No asociado a sectores</h2>
<%
  else
%>
    <h2>Asociado a los sectores...</h2>
    <ul>
<%
    do while(not obj_rst.eof)
      response.write("<li><a href='dn_asociaciones_mostrar.asp?tipo=sector&id="&obj_rst("id_sector")&"'>"&obj_rst("nombre")&"</a></li>")
      obj_rst.movenext
    loop
  end if
%>
    </ul>
<%

  obj_rst.close()
  set obj_rst = nothing
end sub

' ##################################################################
sub muestra_procesos_fichero(byval id)
  sql = "SELECT * FROM dn_alter_ficheros_por_procesos AS fpp INNER JOIN dn_alter_procesos AS p ON fpp.id_proceso = p.id WHERE id_fichero="&id
  set obj_rst = objconn1.execute(sql)
  if (obj_rst.eof) then
%>
    <h2>No asociado a procesos</h2>
<%
  else
%>
    <h2>Asociado a los procesos...</h2>
    <ul>
<%
    do while(not obj_rst.eof)
      response.write("<li><a href='dn_asociaciones_mostrar.asp?tipo=proceso&id="&obj_rst("id_proceso")&"'>"&obj_rst("nombre")&"</a></li>")
      obj_rst.movenext
    loop
  end if
%>
    </ul>
<%

  obj_rst.close()
  set obj_rst = nothing
end sub

' ##################################################################
sub muestra_usos_fichero(byval id)
  sql = "SELECT * FROM dn_alter_ficheros_por_usos AS fpu INNER JOIN dn_risc_usos AS u ON fpu.id_uso = u.id WHERE id_fichero="&id
  set obj_rst = objconn1.execute(sql)
  if (obj_rst.eof) then
%>
    <h2>No asociado a usos</h2>
<%
  else
%>
    <h2>Asociado a los usos...</h2>
    <ul>
<%
    do while(not obj_rst.eof)
      response.write("<li><a href='dn_asociaciones_mostrar.asp?tipo=uso&id="&obj_rst("id_uso")&"'>"&obj_rst("nombre")&"</a></li>")
      obj_rst.movenext
    loop
  end if
%>
    </ul>
<%

  obj_rst.close()
  set obj_rst = nothing
end sub

' ##################################################################
sub muestra_ficheros_sector(byval id)
  sql = "SELECT * FROM dn_alter_ficheros_por_sectores AS fps INNER JOIN dn_alter_ficheros AS f ON fps.id_fichero = f.id WHERE id_sector="&id
  set obj_rst = objconn1.execute(sql)
  if (obj_rst.eof) then
%>
    <h2>No asociado a ficheros</h2>
<%
  else
%>
    <h2>Asociado a los ficheros...</h2>
    <ul>
<%
    do while(not obj_rst.eof)
      response.write("<li><a href='dn_asociaciones_mostrar.asp?tipo=fichero&id="&obj_rst("id_fichero")&"'>"&obj_rst("titulo")&"</a></li>")
      obj_rst.movenext
    loop
  end if
%>
    </ul>
<%

  obj_rst.close()
  set obj_rst = nothing
end sub

' ##################################################################
sub muestra_ficheros_proceso(byval id)
  sql = "SELECT * FROM dn_alter_ficheros_por_procesos AS fpp INNER JOIN dn_alter_ficheros AS f ON fpp.id_fichero = f.id WHERE id_proceso="&id
  set obj_rst = objconn1.execute(sql)
  if (obj_rst.eof) then
%>
    <h2>No asociado a ficheros</h2>
<%
  else
%>
    <h2>Asociado a los ficheros...</h2>
    <ul>
<%
    do while(not obj_rst.eof)
      response.write("<li><a href='dn_asociaciones_mostrar.asp?tipo=fichero&id="&obj_rst("id_fichero")&"'>"&obj_rst("titulo")&"</a></li>")
      obj_rst.movenext
    loop
  end if
%>
    </ul>
<%

  obj_rst.close()
  set obj_rst = nothing
end sub

' ##################################################################
sub muestra_sustancias_uso(byval id)
  sql = "SELECT * FROM dn_risc_sustancias_por_usos AS spu INNER JOIN dn_risc_sustancias AS s ON spu.id_sustancia = s.id WHERE id_uso="&id
  set obj_rst = objconn1.execute(sql)
  if (obj_rst.eof) then
%>
    <h2>No asociado a sustancias</h2>
<%
  else
%>
    <h2>Asociado a las sustancias...</h2>
    <ul>
<%
    do while(not obj_rst.eof)
      response.write("<li><a href='dn_asociaciones_mostrar.asp?tipo=sustancia&id="&obj_rst("id_sustancia")&"'>"&obj_rst("nombre")&"</a></li>")
      obj_rst.movenext
    loop
  end if
%>
    </ul>
<%

  obj_rst.close()
  set obj_rst = nothing
end sub

' ##################################################################
sub muestra_grupos_uso(byval id)
  sql = "SELECT * FROM dn_risc_grupos_por_usos AS gpu INNER JOIN dn_risc_grupos AS g ON gpu.id_grupo = g.id WHERE id_uso="&id
  set obj_rst = objconn1.execute(sql)
  if (obj_rst.eof) then
%>
    <h2>No asociado a grupos</h2>
<%
  else
%>
    <h2>Asociado a los grupos...</h2>
    <ul>
<%
    do while(not obj_rst.eof)
      response.write("<li><a href='dn_asociaciones_mostrar.asp?tipo=grupo&id="&obj_rst("id_grupo")&"'>"&obj_rst("nombre")&"</a></li>")
      obj_rst.movenext
    loop
  end if
%>
    </ul>
<%

  obj_rst.close()
  set obj_rst = nothing
end sub

' ##################################################################
sub muestra_ficheros_uso(byval id)
  sql = "SELECT * FROM dn_alter_ficheros_por_usos AS fpu INNER JOIN dn_alter_ficheros AS f ON fpu.id_fichero = f.id WHERE id_uso="&id
  set obj_rst = objconn1.execute(sql)
  if (obj_rst.eof) then
%>
    <h2>No asociado a ficheros</h2>
<%
  else
%>
    <h2>Asociado a los ficheros...</h2>
    <ul>
<%
    do while(not obj_rst.eof)
      response.write("<li><a href='dn_asociaciones_mostrar.asp?tipo=fichero&id="&obj_rst("id_fichero")&"'>"&obj_rst("titulo")&"</a></li>")
      obj_rst.movenext
    loop
  end if
%>
    </ul>
<%

  obj_rst.close()
  set obj_rst = nothing
end sub

' ##################################################################
sub muestra_ficheros_residuo(byval id)
  sql = "SELECT * FROM dn_alter_ficheros_por_residuos AS fpu INNER JOIN dn_alter_ficheros AS f ON fpu.id_fichero = f.id WHERE id_residuo="&id
  set obj_rst = objconn1.execute(sql)
  if (obj_rst.eof) then
%>
    <h2>No asociado a ficheros</h2>
<%
  else
%>
    <h2>Asociado a los ficheros...</h2>
    <ul>
<%
    do while(not obj_rst.eof)
      response.write("<li><a href='dn_asociaciones_mostrar.asp?tipo=fichero&id="&obj_rst("id_fichero")&"'>"&obj_rst("titulo")&"</a></li>")
      obj_rst.movenext
    loop
  end if
%>
    </ul>
<%

  obj_rst.close()
  set obj_rst = nothing
end sub

%>
