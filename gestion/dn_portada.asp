<!--#include file="dn_fun_comunes.asp"-->
<!--#include file="dn_auten.inc"-->
<!--#include file="adovbs.inc"-->
<!--#include file="dn_conexion.asp"-->



<%
		set cmdsus=Server.CreateObject("ADODB.Command")
		   With cmdsus
			.ActiveConnection=objconn1
			.CommandText="dn_cuenta"
			.CommandType=adCmdStoredProc
			.Parameters.Append  .CreateParameter("@sustancias", adinteger, adParamOutput)
			.Parameters.Append  .CreateParameter("@sinonimos", adinteger, adParamOutput)
			.Parameters.Append  .CreateParameter("@grupos", adinteger, adParamOutput)
			.Parameters.Append  .CreateParameter("@usos", adinteger, adParamOutput)
			.Parameters.Append  .CreateParameter("@companias", adinteger, adParamOutput)
			.Parameters.Append  .CreateParameter("@enfermedades", adinteger, adParamOutput)
			.Parameters.Append  .CreateParameter("@ficheros", adinteger, adParamOutput)
			.Parameters.Append  .CreateParameter("@sectores", adinteger, adParamOutput)
			.Parameters.Append  .CreateParameter("@procesos", adinteger, adParamOutput)
			.Execute,,adexecutenorecords
			sustancias=.Parameters("@sustancias")
			sinonimos=.Parameters("@sinonimos")
			grupos=.Parameters("@grupos")
			usos=.Parameters("@usos")
			companias=.Parameters("@companias")
			enfermedades=.Parameters("@enfermedades")
			ficheros=.Parameters("@ficheros")
			sectores=.Parameters("@sectores")
			procesos=.Parameters("@procesos")
		   End With 
		set cmdsus=nothing		




sql="SELECT COUNT(DISTINCT id_sustancia) AS num FROM dn_risc_sustancias_por_usos WHERE toxico=0"

set objRst=objConn1.execute(sql)

num_alternativas=objRst("num")

objRst.close()

set objRst=nothing


cerrarconexion
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

<h1>Bienvenid@ a la herramienta de Gestión de Ristox y Alternativas</h1>

<h2>Risctox</h2>
<ul>
<li>Podrá gestionar SUSTANCIAS, GRUPOS, COMPAÑIAS Y ENFERMEDADES</li>
<li>Desde SUSTANCIAS, podrá asociar sustancias a GRUPOS, USOS, COMPAÑIAS, ENFERMEDADES Y FICHEROS</li>
<li>Desde GRUPOS, podrá asociar grupos a  USOS, ENFERMEDADES Y FICHEROS</li>
<li>Para ver que SUSTANCIAS Y/O GRUPOS están asociados a determinados USOS, COMPAÑIAS Y ENFERMEDADES, vaya a la sección correspondiente (respectivamente USOS, COMPAÑIAS, ENFERMEDADES). Desde ella, podrá también eliminar la asociación.</li>
</ul>

<h2>Alternativas</h2>
<ul>
<li>Podrá gestionar FICHEROS, SECTORES, PROCESOS Y USOS</li>
<li>Desde SECTORES, podr&aacute; asociar sectores a FICHEROS </li>
<li>Desde PROCESOS, podr&aacute; asociar procesos a FICHEROS </li>
<li>Desde USOS, podr&aacute; asociar usos a FICHEROS </li>
<li>Desde FICHEROS, podrá ver que SUSTANCIAS, GRUPOS, SECTORES, PROCESOS Y USOS hay asociados a cada fichero, as&iacute; como eliminar la asociaci&oacute;n</li>
</ul>


<ul id="split">
<li id="one">
  <h3>Resumen Risctox</h3>

  <div>
<p><strong><%=sustancias%></strong> sustancias </p>
<p><strong><%=sinonimos%></strong> sinónimos </p>
<p><strong><%=grupos%></strong> grupos </p>
<p><strong><%=companias%></strong> companias </p>
<p><strong><%=enfermedades%></strong> enfermedades </p>
</div>
</li>
<li id="two">
  <h3>Resumen Alternativas </h3>
<div>
<p><strong><%=num_alternativas%></strong> sustancias alternativas </p>

<p><strong><%=ficheros%></strong> ficheros </p>
<p><strong><%=sectores%></strong> sectores </p>
<p><strong><%=procesos%></strong> procesos </p>
<p><strong><%=usos%></strong> usos </p>
</div>
</li>
</ul>
</ul>

</body>
</html>


