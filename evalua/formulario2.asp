<!--#include file="../dn_conexion.asp"-->
<%

if session("id_ecogente2")<>"" then response.redirect "formulario_identificado.asp"


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


FUNCTION valores (vgrupo,vname,tipo)

orden2 = "SELECT * FROM ECOINFORMAS_VALORES WHERE grupo='"&vgrupo&"' ORDER BY valor,desc1"
Set dSQL2 = Server.CreateObject ("ADODB.Recordset")
dSQL2.Open orden2,objConnection,adOpenKeyset
if tipo="1" then 				'--- tipo 1 = desplegable
%>	<select name=<%=vname%> class="campo">
	<option value="">- Selecciona de la lista -</option><%
	if not(DSQL2.bof and DSQL2.eof) then
		dSQL2.movefirst
		DO while not dSQL2.eof
		  %><option <%=sele%> value="<%=dSQL2("valor")%>"><%=dSQL2("desc1")%></option><%
	        dSQL2.movenext
	        loop
	end if
%>	
	</select>
<%
else 						'--- tipo 2 = selecci�n visible por radio
	if not(DSQL2.bof and DSQL2.eof) then
		dSQL2.movefirst
		DO while not dSQL2.eof
	%>
	<input type="radio" name="<%=vname%>" class="campo" value="<%=dSQL2("valor")%>">&nbsp;<%=dSQL2("desc1")%>&nbsp;&nbsp;
	<%      dSQL2.movenext
	        loop
	end if
end if
dSQL2.close        
END FUNCTION

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ISTAS: BBDD Eval�a lo que usas</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="Title" content="ECOinformas" />
<meta name="Author" content="XiP multim�dia" />
<meta name="description" content="Informaci�n, formaci�n y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Subject" content="Informaci�n, formaci�n y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Keywords" content="Informaci�n, formaci�n y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Language" content="Spanish" />
<meta name="Revisit" content="15 days" />
<meta name="Distribution" content="Global" />
<meta name="Robots" content="All" />
<link rel="stylesheet" type="text/css" href="estructura.css"  />
<SCRIPT LANGUAGE="JavaScript">
<!--
function enviar() {
	
	if (document.getElementById('email').value==''){
	  alert ('Introduzca su correo electr�nico');
	}
	else{
	  document.getElementById('formulario').submit();
	}
}

// -->
</SCRIPT>
</head>
<body>
<script src="../valida_fecha.js"></script>
<script src="../valida_mail.js"></script>
<script src="../valida_textarea.js"></script>
<script src="../valida_dni.js"></script>

<div id="contenedor">
	<div id="sombra_arriba"></div>
  	<div id="sombra_lateral">
		<div id="caja">
			<!--#include file="../dn_cabecera.asp"-->
			<div id="texto">
			
			<div class="texto">
				<table width="100%" border="0">
                <tr>
                <td></td>
                <td align='right'><input type="button" name="volver" class="boton" value="Volver a la portada de eval�a lo que usas" onClick="window.location='./index.asp';"></td>
                </tr>
                </table>
				<p class="titulo3">Solicitud de acceso libre</p>
			
				<br />
                <br />
				
<form method="POST" name="formulario" id='formulario' action="formulario4_grabar.asp">

<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">

<tr><td class="texto" align="justify" colspan=2>Por favor introduce tu correo electr�nico para que podamos atender tu solicitud de acceso libre a toda nuestra p�gina web*. Autom�ticamente obtendr�s una clave y una contrase�a que te permiten entrar en la web de Riesto Qu�mico y aprovechar los materiales y herramientas que te ofrecemos. 
<tr><td class="texto" align="center" colspan=2>&nbsp;</td></tr> 
    <tr>
    	<td colspan=2 align="center">
        	<table cellspacing=0 cellpadding=0 border=0>
        		<tr>
	                <td  valign="middle" align='right'>Correo electr�nico:</td>
			        <td ><input type="text" name="email" id='email' size="50" maxlength="200" class="campo" ONBLUR='valida_mail(this)'></td>
				</tr>    
            </table>        
        </td>
    </tr>
	<tr><td class=texto align=center colspan=2>
   <input type="button" value="ENVIAR DATOS" name="comprobar" class="boton" onClick="enviar()">
	</td></tr>
</table>


<br>&nbsp;
<br>&nbsp;
<br>&nbsp;
<br>&nbsp;

<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">

<tr><td class=texto align=center>
(*) Los datos que nos facilites ser�n incorporados a un fichero bajo titularidad de ISTAS. La finalidad del tratamiento de sus datos la constituye la posibilidad de difusi�n por correo electr�nico y ordinario de informaci�n y materiales de ECOinformas; la promoci�n de la salud laboral y la protecci�n del medio ambiente a trav�s de la remisi�n de informaci�n sobre los productos editoriales y actividades de ISTAS; auditor�a por parte de la Fundaci�n Biodiversidad que se compromete a su vez a cumplir la Ley Org�nica de Protecci�n de Datos de car�cter Personal (LOPD). Para m�s informaci�n: <a href="http://www.istas.net/ecoinformas/index.asp?idpagina=558" target="_blank">pol�tica de privacidad.</a>
</td></tr>
   
</table>
</form>

<!-- fin formulario -->				
				</div>
				<p>&nbsp;</p>
			</div>
				
                <img src="../imagenes/pie_risctox.gif" width="708" border="0">

    			</div>
    		</div>
		<div id="sombra_abajo"></div>
	</div>
</body>
</html>
