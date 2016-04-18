<!--#include file="dn_fun_comunes.asp"-->
<!--#include file="dn_fun_texto.asp"-->
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
<script src="sorttable.js"></script>
<script type="text/javascript" src="niftycube.js"></script>
<script type="text/javascript">
window.onload=function(){
Nifty("div#box2","big"); 
}
function cambiachecks(c) 
{
	var frm = document.forms["myform"]; 
	for (i=0; i<frm.elements.length; i++)
	{
		if(frm.elements[i].name=='idcheck')
		{
			frm.elements[i].checked=c.checked;
		}
	}
}
</script>
<script language="JavaScript" src="gen_validatorv2.js" type="text/javascript"></script>
</head>

<body>
<%flashMsgShow()%>
<div id="box2" class="centcontenido">

<%
asociar=request("asociar")
id=request("id")
if id="" then
	cerrarconexion
%>
<script>window.close();</script>
<%
else
	select case asociar
		
		case "grupo":
		
			sql3="select sg.id_sustancia, s.num_cas, s.nombre from dn_risc_sustancias s INNER JOIN dn_risc_sustancias_por_grupos sg ON s.id=sg.id_sustancia  WHERE  sg.id_grupo=" &id
			set objRst3=objconn1.execute(sql3)		
			if objRst3.eof then
				tablares="No hay sustancias asociadas. Para asociar una sustancia, vaya a la sección SUSTANCIAS"
			else
				do while not objRst3.eof
						  tablares=tablares & "<tr>" 
						  tablares=tablares & "<td><input type='checkbox' name='idcheck' value=" &objRst3("id_sustancia")& " /></td>"						  
				  		  tablares=tablares & "<td align='left'>"  &objRst3("num_cas")&  "</td>"	
						  tablares=tablares & "<td align='left'>" &objRst3("nombre")&  "</td>"							
						  tablares=tablares & "</tr>"	
				objRst3.movenext
				loop
				
				tablares="<p>Para desasociar una sustancia, márquela y pulse en <em>Desasociar</em></p><table id='resultados' cellspacing='0' cellpadding='3' align='center'><tr><th><input type='checkbox' name='selector' onchange='cambiachecks(this)' /></th><th>CAS</th><th>SUSTANCIA</th></tr>" &tablares& "</table><p><input type='submit' value='Desasociar' class='centcontenido'  /></p>"
 			
			end if
					
			objRst3.close
			set objRst3=nothing		
			
		case "enfermedad":
		
			sql3="select sg.id_sustancia, s.num_cas, s.nombre from dn_risc_sustancias s INNER JOIN dn_risc_sustancias_por_enfermedades sg ON s.id=sg.id_sustancia  WHERE  sg.id_enfermedad=" &id
			set objRst3=objconn1.execute(sql3)		
			if objRst3.eof then
				tablares="No hay sustancias asociadas. Para asociar una sustancia, vaya a la sección SUSTANCIAS"
			else
				do while not objRst3.eof
						  tablares=tablares & "<tr>" 
						  tablares=tablares & "<td><input type='checkbox' name='idcheck' value=" &objRst3("id_sustancia")& " /></td>"						  
				  		  tablares=tablares & "<td align='left'>"  &objRst3("num_cas")&  "</td>"	
						  tablares=tablares & "<td align='left'>" &objRst3("nombre")&  "</td>"							
						  tablares=tablares & "</tr>"	
				objRst3.movenext
				loop
				
				tablares="<p>Para desasociar una sustancia, márquela y pulse en <em>Desasociar</em></p><table id='resultados' cellspacing='0' cellpadding='3' align='center'><tr><th><input type='checkbox' name='selector' onchange='cambiachecks(this)' /></th><th>CAS</th><th>SUSTANCIA</th></tr>" &tablares& "</table><p><input type='submit' value='Desasociar' class='centcontenido'  /></p>"
 			
			end if
					
			objRst3.close
			set objRst3=nothing		
			
		case "enfermedad_gr":
		
			sql3="select sg.id_grupo,  s.nombre from dn_risc_grupos s INNER JOIN dn_risc_grupos_por_enfermedades sg ON s.id=sg.id_grupo  WHERE  sg.id_enfermedad=" &id
			set objRst3=objconn1.execute(sql3)		
			if objRst3.eof then
				tablares="No hay grupos asociados. Para asociar un grupo, vaya a la sección GRUPOS."
			else
				do while not objRst3.eof
						  tablares=tablares & "<tr>" 
						  tablares=tablares & "<td><input type='checkbox' name='idcheck' value=" &objRst3("id_grupo")& " /></td>"						  
						  tablares=tablares & "<td align='left'>" &objRst3("nombre")&  "</td>"							
						  tablares=tablares & "</tr>"	
				objRst3.movenext
				loop
				
				tablares="<p>Para desasociar un grupo, márquelo y pulse en <em>Desasociar</em></p><table id='resultados' cellspacing='0' cellpadding='3' align='center'><tr><th><input type='checkbox' name='selector' onchange='cambiachecks(this)' /></th><th>GRUPO</th></tr>" &tablares& "</table><p><input type='submit' value='Desasociar' class='centcontenido'  /></p>"
 			
			end if
					
			objRst3.close
			set objRst3=nothing		
			
		case "uso":
		
			sql3="select toxico, sg.id_sustancia, s.num_cas, s.nombre from dn_risc_sustancias s INNER JOIN dn_risc_sustancias_por_usos sg ON s.id=sg.id_sustancia  WHERE  sg.id_uso=" &id
			set objRst3=objconn1.execute(sql3)		
			if objRst3.eof then
				tablares="No hay sustancias asociadas. Para asociar una sustancia, vaya a la sección SUSTANCIAS"
			else
				do while not objRst3.eof
							if objRst3("toxico") then
								trstyle=" style='color:red;'"
							else
								trstyle=""
							end if
						  tablares=tablares & "<tr" &trstyle& ">" 
						  tablares=tablares & "<td><input type='checkbox' name='idcheck' value=" &objRst3("id_sustancia")& " /></td>"						  
				  		  tablares=tablares & "<td align='left'>"  &objRst3("num_cas")&  "</td>"	
						  tablares=tablares & "<td align='left'>" &objRst3("nombre")&  "</td>"			
						  tablares=tablares & "<td align='left'>" &objRst3("toxico")&  "</td>"								
						  tablares=tablares & "</tr>"	
				objRst3.movenext
				loop
				
				tablares="<p>Para desasociar una sustancia, márquela y pulse en <em>Desasociar</em></p><table class='sortable' id='resultados' cellspacing='0' cellpadding='3' align='center'><tr><th><input type='checkbox' name='selector' onchange='cambiachecks(this)' /></th><th>CAS</th><th>SUSTANCIA</th><th>TÓXICO</th></tr>" &tablares& "</table><p><input type='submit' value='Desasociar' class='centcontenido'  /></p>"
 			
			end if
					
			objRst3.close
			set objRst3=nothing		
			
		case "uso_gr":
		
			sql3="select toxico, sg.id_grupo,  s.nombre from dn_risc_grupos s INNER JOIN dn_risc_grupos_por_usos sg ON s.id=sg.id_grupo  WHERE  sg.id_uso=" &id
			set objRst3=objconn1.execute(sql3)		
			if objRst3.eof then
				tablares="No hay grupos asociados. Para asociar un grupo, vaya a la sección GRUPOS."
			else
				do while not objRst3.eof
						  if objRst3("toxico") then
								trstyle=" style='color:red;'"
							else
								trstyle=""
							end if
						  tablares=tablares & "<tr" &trstyle& ">" 
						  tablares=tablares & "<td><input type='checkbox' name='idcheck' value=" &objRst3("id_grupo")& " /></td>"						  
						  tablares=tablares & "<td align='left'>" &objRst3("nombre")&  "</td>"				
						  tablares=tablares & "<td align='left'>" &objRst3("toxico")&  "</td>"				
						  tablares=tablares & "</tr>"	
				objRst3.movenext
				loop
				
				tablares="<p>Para desasociar un grupo, márquelo y pulse en <em>Desasociar</em></p><table class='sortable' id='resultados' cellspacing='0' cellpadding='3' align='center'><tr><th><input type='checkbox' name='selector' onchange='cambiachecks(this)' /></th><th>GRUPO</th><th>TÓXICO</th></tr>" &tablares& "</table><p><input type='submit' value='Desasociar' class='centcontenido'  /></p>"
 			
			end if
					
			objRst3.close
			set objRst3=nothing		
			
		case "compania":
		
			sql3="select sg.id_sustancia, s.num_cas, s.nombre from dn_risc_sustancias s INNER JOIN dn_risc_sustancias_por_companias sg ON s.id=sg.id_sustancia  WHERE  sg.id_compania=" &id
			set objRst3=objconn1.execute(sql3)		
			if objRst3.eof then
				tablares="No hay sustancias asociadas. Para asociar una sustancia, vaya a la sección SUSTANCIAS"
			else
				do while not objRst3.eof
						  tablares=tablares & "<tr>" 
						  tablares=tablares & "<td><input type='checkbox' name='idcheck' value=" &objRst3("id_sustancia")& " /></td>"						  
				  		  tablares=tablares & "<td align='left'>"  &objRst3("num_cas")&  "</td>"	
						  tablares=tablares & "<td align='left'>" &objRst3("nombre")&  "</td>"							
						  tablares=tablares & "</tr>"	
				objRst3.movenext
				loop
				
				tablares="<p>Para desasociar una sustancia, márquela y pulse en <em>Desasociar</em></p><table id='resultados' cellspacing='0' cellpadding='3' align='center'><tr><th><input type='checkbox' name='selector' onchange='cambiachecks(this)' /></th><th>CAS</th><th>SUSTANCIA</th></tr>" &tablares& "</table><p><input type='submit' value='Desasociar' class='centcontenido'  /></p>"
 			
			end if
					
			objRst3.close
			set objRst3=nothing		



' ############################



		case "sector":
		
			sql3="select ss.id_sustancia, s.num_cas, s.nombre from dn_risc_sustancias s INNER JOIN dn_risc_sustancias_por_sectores AS ss ON s.id=ss.id_sustancia WHERE ss.id_sector=" &id
			set objRst3=objconn1.execute(sql3)		
			if objRst3.eof then
				tablares="No hay sustancias asociadas. Para asociar una sustancia, vaya a la sección SUSTANCIAS"
			else
				do while not objRst3.eof
						  tablares=tablares & "<tr>" 
						  tablares=tablares & "<td><input type='checkbox' name='idcheck' value=" &objRst3("id_sustancia")& " /></td>"						  
				  		  tablares=tablares & "<td align='left'>"  &objRst3("num_cas")&  "</td>"	
						  tablares=tablares & "<td align='left'>" &objRst3("nombre")&  "</td>"							
						  tablares=tablares & "</tr>"	
				objRst3.movenext
				loop
				
				tablares="<p>Para desasociar una sustancia, márquela y pulse en <em>Desasociar</em></p><table id='resultados' cellspacing='0' cellpadding='3' align='center'><tr><th><input type='checkbox' name='selector' onchange='cambiachecks(this)' /></th><th>CAS</th><th>SUSTANCIA</th></tr>" &tablares& "</table><p><input type='submit' value='Desasociar' class='centcontenido'  /></p>"
 			
			end if
					
			objRst3.close
			set objRst3=nothing		



' ############################




			
'FICHEROS

		case "fich_sustancia":
		
			sql3="select sg.id_sustancia, s.num_cas, s.nombre from dn_risc_sustancias s INNER JOIN dn_alter_ficheros_por_sustancias sg ON s.id=sg.id_sustancia  WHERE  sg.id_fichero=" &id
			set objRst3=objconn1.execute(sql3)		
			if objRst3.eof then
				tablares="No hay ficheros asociados a sustancias. Para asociar una sustancia, vaya a la sección SUSTANCIAS"
			else
				do while not objRst3.eof
						  tablares=tablares & "<tr>" 
						  tablares=tablares & "<td><input type='checkbox' name='idcheck' value=" &objRst3("id_sustancia")& " /></td>"						  
				  		  tablares=tablares & "<td align='left'>"  &objRst3("num_cas")&  "</td>"	
						  tablares=tablares & "<td align='left'>" &objRst3("nombre")&  "</td>"							
						  tablares=tablares & "</tr>"	
				objRst3.movenext
				loop
				
				tablares="<p>Para desasociar un elemento, marqueló y pulse en <em>Desasociar</em></p><table id='resultados' cellspacing='0' cellpadding='3' align='center'><tr><th><input type='checkbox' name='selector' onchange='cambiachecks(this)' /></th><th>CAS</th><th>SUSTANCIA</th></tr>" &tablares& "</table><p><input type='submit' value='Desasociar' class='centcontenido'  /></p>"
 			
			end if
					
			objRst3.close
			set objRst3=nothing		
			
		case "fich_grupo":
		
			sql3="select asociado.id as idcheck, asociado.num_cas, asociado.nombre from dn_risc_grupos as asociado INNER JOIN dn_alter_ficheros_por_grupos as por ON asociado.id=por.id_grupo  WHERE  por.id_fichero=" &id
			
			set objRst3=objconn1.execute(sql3)		
			if objRst3.eof then
				tablares="No hay elementos asociados al fichero."
			else
				do while not objRst3.eof
						  tablares=tablares & "<tr>" 
						  tablares=tablares & "<td><input type='checkbox' name='idcheck' value=" &objRst3("idcheck")& " /></td>"						  
				  		  tablares=tablares & "<td align='left'>"  &objRst3("num_cas")&  "</td>"	
						  tablares=tablares & "<td align='left'>" &objRst3("nombre")&  "</td>"							
						  tablares=tablares & "</tr>"	
				objRst3.movenext
				loop
				
				tablares="<p>Para desasociar un elemento, marqueló y pulse en <em>Desasociar</em></p><table id='resultados' cellspacing='0' cellpadding='3' align='center'><tr><th><input type='checkbox' name='selector' onchange='cambiachecks(this)' /></th><th>CAS</th><th>GRUPO</th></tr>" &tablares& "</table><p><input type='submit' value='Desasociar' class='centcontenido'  /></p>"
 			
			end if
					
			objRst3.close
			set objRst3=nothing		
			
		case "fich_sector":
		
			sql3="select asociado.id as idcheck, asociado.numero_cnae, asociado.nombre from dn_alter_sectores as asociado INNER JOIN dn_alter_ficheros_por_sectores as por ON asociado.id=por.id_sector  WHERE  por.id_fichero=" &id
			
			set objRst3=objconn1.execute(sql3)		
			if objRst3.eof then
				tablares="No hay elementos asociados al fichero."
			else
				do while not objRst3.eof
						  tablares=tablares & "<tr>" 
						  tablares=tablares & "<td><input type='checkbox' name='idcheck' value=" &objRst3("idcheck")& " /></td>"						  
				  		  tablares=tablares & "<td align='left'>"  &objRst3("numero_cnae")&  "</td>"	
						  tablares=tablares & "<td align='left'>" &objRst3("nombre")&  "</td>"							
						  tablares=tablares & "</tr>"	
				objRst3.movenext
				loop
				
				tablares="<p>Para desasociar un elemento, marqueló y pulse en <em>Desasociar</em></p><table id='resultados' cellspacing='0' cellpadding='3' align='center'><tr><th><input type='checkbox' name='selector' onchange='cambiachecks(this)' /></th><th>NUMERO CNAE</th><th>SECTOR</th></tr>" &tablares& "</table><p><input type='submit' value='Desasociar' class='centcontenido'  /></p>"
 			
			end if
					
			objRst3.close
			set objRst3=nothing	
			
		case "fich_proceso":
		
			sql3="select asociado.id as idcheck, asociado.descripcion, asociado.nombre from dn_alter_procesos as asociado INNER JOIN dn_alter_ficheros_por_procesos as por ON asociado.id=por.id_proceso  WHERE  por.id_fichero=" &id
			
			set objRst3=objconn1.execute(sql3)		
			if objRst3.eof then
				tablares="No hay elementos asociados al fichero."
			else
				do while not objRst3.eof
						  tablares=tablares & "<tr>" 
						  tablares=tablares & "<td><input type='checkbox' name='idcheck' value=" &objRst3("idcheck")& " /></td>"						  
				  		  tablares=tablares & "<td align='left'>"  &objRst3("nombre")&  "</td>"	
						  tablares=tablares & "<td align='left'>" &corta (objRst3("descripcion"), 40, "puntossuspensivos")&  "</td>"							
						  tablares=tablares & "</tr>"	
				objRst3.movenext
				loop
				
				tablares="<p>Para desasociar un elemento, marqueló y pulse en <em>Desasociar</em></p><table id='resultados' cellspacing='0' cellpadding='3' align='center'><tr><th><input type='checkbox' name='selector' onchange='cambiachecks(this)' /></th><th>PROCESO</th><th>DESCRIPCION</th></tr>" &tablares& "</table><p><input type='submit' value='Desasociar' class='centcontenido'  /></p>"
 			
			end if
					
			objRst3.close
			set objRst3=nothing		
			
		case "fich_uso":
		
			sql3="select asociado.id as idcheck,  asociado.nombre from dn_risc_usos as asociado INNER JOIN dn_alter_ficheros_por_usos as por ON asociado.id=por.id_uso  WHERE  por.id_fichero=" &id
			
			set objRst3=objconn1.execute(sql3)		
			if objRst3.eof then
				tablares="No hay elementos asociados al fichero."
			else
				do while not objRst3.eof
						  tablares=tablares & "<tr>" 
						  tablares=tablares & "<td><input type='checkbox' name='idcheck' value=" &objRst3("idcheck")& " /></td>"						  
				  		  tablares=tablares & "<td align='left'>"  &objRst3("nombre")&  "</td>"	
						  
						  tablares=tablares & "</tr>"	
				objRst3.movenext
				loop
				
				tablares="<p>Para desasociar un elemento, marqueló y pulse en <em>Desasociar</em></p><table id='resultados' cellspacing='0' cellpadding='3' align='center'><tr><th><input type='checkbox' name='selector' onchange='cambiachecks(this)' /></th><th>USO</th></tr>" &tablares& "</table><p><input type='submit' value='Desasociar' class='centcontenido'  /></p>"
 			
			end if
					
			objRst3.close
			set objRst3=nothing		
			


		case "fich_residuo":
		
			sql3="select asociado.id as idcheck,asociado.nombre,asociado.codigo from rq_residuos as asociado INNER JOIN dn_alter_ficheros_por_residuos as por ON asociado.id=por.id_residuo  WHERE  por.id_fichero=" &id
			
			set objRst3=objconn1.execute(sql3)		
			if objRst3.eof then
				tablares="No hay residuos asociados al fichero "&id
			else
				do while not objRst3.eof
						  tablares=tablares & "<tr>" 
						  tablares=tablares & "<td><input type='checkbox' name='idcheck' value=" &objRst3("idcheck")& " /></td>"						  
				  		tablares=tablares & "<td align='left'>"  &objRst3("codigo")&  "</td>"	
				  		tablares=tablares & "<td align='left'>"  &objRst3("nombre")&  "</td>"	
						  tablares=tablares & "</tr>"	
				objRst3.movenext
				loop
				
				tablares="<p>Para desasociar un elemento, márquelo y pulse en <em>Desasociar</em></p><table id='resultados' cellspacing='0' cellpadding='3' align='center'><tr><th><input type='checkbox' name='selector' onchange='cambiachecks(this)' /></th><th>RESIDUO</th></tr>" &tablares& "</table><p><input type='submit' value='Desasociar' class='centcontenido'  /></p>"
 			
			end if
					
			objRst3.close
			set objRst3=nothing		
			

					
	end select
end if
cerrarconexion
%>

<form name="myform" action="dn_desasociar2.asp?asociar=<%=asociar%>&id=<%=id%>" method="post" >
 
<fieldset><legend><strong>Asociaciones</strong></legend>
<%=tablares%>
</fieldset>

</form>

</div>
</body>
</html>


