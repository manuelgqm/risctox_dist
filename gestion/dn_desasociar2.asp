<!--#include file="dn_fun_comunes.asp"-->
<!--#include file="dn_auten.inc"-->
<!--#include file="adovbs.inc"-->
<!--#include file="dn_conexion.asp"-->

<%'Permite auditar las acciones%>
<!--#include file="spl_auditoria.inc.asp"-->

<%
asociar=request("asociar")
id=request("id") 'id del elemento que desasociamos
idcheck=request.form("idcheck")

'montamos un array con los ids de sustancia 
'arr = split(id_sustancia, ",")
'FOR i=0 to UBound(arr)		
	
	select case asociar
		
		case "grupo":
				
			sql3="delete from dn_risc_sustancias_por_grupos where id_sustancia IN (" &idcheck& ") and id_grupo=" &id
			objconn1.execute(sql3)	

		case "enfermedad":
				
			sql3="delete from dn_risc_sustancias_por_enfermedades where id_sustancia IN (" &idcheck& ") and id_enfermedad=" &id
			objconn1.execute(sql3)	
			
		case "enfermedad_gr":
				
			sql3="delete from dn_risc_grupos_por_enfermedades where id_grupo IN (" &idcheck& ") and id_enfermedad=" &id
			objconn1.execute(sql3)	
			
		case "uso":
				
			sql3="delete from dn_risc_sustancias_por_usos where id_sustancia IN (" &idcheck& ") and id_uso=" &id
			objconn1.execute(sql3)	
			
		case "uso_gr":
				
			sql3="delete from dn_risc_grupos_por_usos where id_grupo IN (" &idcheck& ") and id_uso=" &id
			objconn1.execute(sql3)				
			
		case "compania":
				
			sql3="delete from dn_risc_sustancias_por_companias where id_sustancia IN (" &idcheck& ") and id_compania=" &id
			objconn1.execute(sql3)	

		case "sector":
				
			sql3="delete from dn_risc_sustancias_por_sectores where id_sustancia IN (" &idcheck& ") and id_sector=" &id
			objconn1.execute(sql3)	
			
		'FICHEROS
		
		case "fich_sustancia":
				
			sql3="delete from dn_alter_ficheros_por_sustancias where id_sustancia IN (" &idcheck& ") and id_fichero=" &id
			objconn1.execute(sql3)	
			
		case "fich_grupo":
				
			sql3="delete from dn_alter_ficheros_por_grupos where id_grupo IN (" &idcheck& ") and id_fichero=" &id
			objconn1.execute(sql3)	
			
		case "fich_sector":
				
			sql3="delete from dn_alter_ficheros_por_sectores where id_sector IN (" &idcheck& ") and id_fichero=" &id
			objconn1.execute(sql3)	
			
		case "fich_proceso":
				
			sql3="delete from dn_alter_ficheros_por_procesos where id_proceso IN (" &idcheck& ") and id_fichero=" &id
			objconn1.execute(sql3)	
			
		case "fich_uso":
				
			sql3="delete from dn_alter_ficheros_por_usos where id_uso IN (" &idcheck& ") and id_fichero=" &id
			objconn1.execute(sql3)	

		case "fich_residuo":
				
			sql3="delete from dn_alter_ficheros_por_residuos where id_residuo IN (" &idcheck& ") and id_fichero=" &id
			objconn1.execute(sql3)
													
	end select
'NEXT		

' ** AUDITORIA
spl_accion = "desasociar"
spl_entidad = asociar
spl_descripcion = sql3
' ** AUDITORIA **
call auditaYCierraConexion(spl_accion,spl_entidad,spl_descripcion) ' accion, entidad, descripcion		

flashMsgCreate "Los elementos seleccionados se han desasociado.", "OK"		
response.Redirect("dn_desasociar.asp?asociar=" &asociar& "&id="&id)
%>
