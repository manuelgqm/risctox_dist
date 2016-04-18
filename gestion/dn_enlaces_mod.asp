<!--#include file="dn_fun_comunes.asp"-->
<!--#include file="dn_fun_texto.asp"-->
<!--#include file="dn_auten.inc"-->
<!--#include file="adovbs.inc"-->
<!--#include file="dn_conexion.asp"-->
<%
id =request("id")
sql3="select titulo, enlace, texto, clasificacion from dn_alter_enlaces where id="& id


set objRst3=objconn1.execute(sql3)


titulo = objrst3("titulo")
enlace = objrst3("enlace")
texto = objrst3("texto")
clasificacion = objrst3("clasificacion")

function selecd(valor1, valor2)
	
	if (valor1 = valor2) then 
		salida = "selected"	
	end if
	selecd = salida
	
end function

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

function valida_longitud(campo, max)
{
  if (campo.value.length > max)
  {
    campo.value = campo.value.substr(0,max);
  }
}
</script>
<script language="JavaScript" src="gen_validatorv2.js" type="text/javascript"></script>
</head>

<body>
<%flashMsgShow()%>
<div id="box2" class="centcontenido">

<%

id=request("id")
if id="" then
	cerrarconexion
%>
	<script>window.close();</script>
<%
else

end if
cerrarconexion
%>

<form name="myform" action="dn_enlaces_sql.asp">
<div align='left'>
    <fieldset>
            <legend><strong>Modificar enlace</strong></legend>
    
            Título<br /><input type="text" name="titulo" maxlength="250" size="80" value='<%=titulo%>' /><br/><br/>
            Enlace<br /><input type="text" name="enlace" maxlength="250" size="80" value='<%=enlace%>' /><br/><br/>
            Texto<br /><textarea name="texto" rows="6" cols="80"><%=texto%></textarea><br />
            Clasificación<br />
                <select name='clasificacion'>
                    <option value='1' <%=selecd(cstr(clasificacion),"1")%>>Fuentes de información generales</option>
                    <option value='2' <%=selecd(cstr(clasificacion),"2")%>>Información sobre eliminación/sustitución</option>
                </select>
            <br />
            <br /><input type="submit"  value="Modificar" />  
    
    </fieldset>
    <input type='hidden' name='id' value='<%=id%>'>
    </div>
</form>

</div>
</body>
</html>
