
 <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ISTAS: BBDD Evalúa lo que usas</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="Title" content="ECOinformas" />
<meta name="Author" content="DABNE" />
<meta name="description" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Subject" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Keywords" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Language" content="Spanish" />
<meta name="Revisit" content="15 days" />
<meta name="Distribution" content="Global" />
<meta name="Robots" content="All" />

<link rel="stylesheet" type="text/css" media="screen" href="../estructura.css">
<link rel="stylesheet" type="text/css" media="print" href="../estructura_impresion.css">
<link rel="stylesheet" type="text/css" media="screen" href="../dn_estilos.css">
<link rel="stylesheet" type="text/css" media="print" href="../dn_estilos_impresion_cesta.css">

</head>
<body onload="print();">
<div class="texto">
	<% 
	cesta = request("cesta")
	cesta = replace(cesta,"<table class=""dn_auto_tabla"" align=""center"" border=""0"" cellpadding","<table class=""dn_auto_tabla"" align=""center"" border=""1"" cellpadding")
	
	response.write cesta
	
	%>
</div>
</body>
</html>