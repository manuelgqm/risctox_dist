<%@LANGUAGE="VBSCRIPT" LCID="1034" CODEPAGE="65001"%>
<%
	Option Explicit
	Response.ContentType = "text/html"
	Response.AddHeader "Content-Type", "text/html;charset=UTF-8"
	Response.CodePage = 65001
	Response.CharSet = "UTF-8"
%>
<!DOCTYPE html>
<html lang="es">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Risctox mobile home page</title>
    <link href="../bootstrap-3.2.0-dist/css/bootstrap.min.css" rel="stylesheet">
	<link rel="stylesheet" href="styles.css" />
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
    <script src="../bootstrap-3.2.0-dist/js/bootstrap.min.js"></script>
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/html5shiv/3.7.2/html5shiv.min.js"></script>
      <script src="https://oss.maxcdn.com/respond/1.4.2/respond.min.js"></script>
    <![endif]-->
  </head>
  <body>
	<div id="container" class="text-center">
		<div class="jumbotron">
			<h1>Base de datos Risctox</h1>
			<p>Información sobre peligros y toxicidad para más de 100.000 sustancias químicas</p>
		</div>
		<div class="well well-lg">
			<form action="substance_card.html" role="form">
				<div class="form-group">
					<label for="substance-name">Sustancia</label>
					<input name="subtance-name" type="text" class="form-control" />
				</div>
				<div class="form-group">
					<label for="substance-number">Número CAS, CE o RD</label>
					<input name="substance-number" type="text" class="form-control" />
				</div>
				<button type="submit" class="btn btn-primary">Mostrar</button>
			</form>
		</div>
	</div>
  </body>
</html>