<!--#include file="dn_fun_comunes.asp"-->
<!--#include file="dn_fun_texto.asp"-->
<!--#include file="dn_auten.inc"-->

<%
id=request.QueryString("id")
if id="" then
%>
<script>window.close();</script>
<%
else
%>
	<!--#include file="adovbs.inc"-->
	<!--#include file="dn_conexion.asp"-->
<%
	'RECOGEMOS VALORES Y CREAMOS INSTRUCCIONES
	
	for each item in request.form
		
		tabla=left(item,3)
		campo=replace(item,tabla,"")
		valor=h(request.form(item))
		
		'si realmente es una tabla
		
			SELECT CASE tabla
			
				CASE "t1_":
				t1_sqlupdate=t1_sqlupdate& campo & "='" & valor & "', "
				t1_sqlinsertcampos=t1_sqlinsertcampos& campo& ", "
				t1_sqlinsertvalores=t1_sqlinsertvalores& "'" & valor & "', "
				
				CASE "t0_": 'siempre hay, siempre es update del campo notas_cancer_rd
				t0_sqlupdate=t0_sqlupdate& campo & "='" & valor & "', "
				
				CASE "t2_":
				t2_sqlupdate=t2_sqlupdate& campo & "='" & valor & "', "
				t2_sqlinsertcampos=t2_sqlinsertcampos& campo& ", "
				t2_sqlinsertvalores=t2_sqlinsertvalores& "'" & valor & "', "
				
				CASE "t3_":
				t3_sqlupdate=t3_sqlupdate& campo & "='" & valor & "', "
				t3_sqlinsertcampos=t3_sqlinsertcampos& campo& ", "
				t3_sqlinsertvalores=t3_sqlinsertvalores& "'" & valor & "', "

        ' ##########################################################
				CASE "t9_":
				t9_sqlupdate=t9_sqlupdate& campo & "='" & valor & "', "
				t9_sqlinsertcampos=t9_sqlinsertcampos& campo& ", "
				t9_sqlinsertvalores=t9_sqlinsertvalores& "'" & valor & "', "      
        ' ##########################################################

        ' ##########################################################
				CASE "t8_":
				t8_sqlupdate=t8_sqlupdate& campo & "='" & valor & "', "
				t8_sqlinsertcampos=t8_sqlinsertcampos& campo& ", "
				t8_sqlinsertvalores=t8_sqlinsertvalores& "'" & valor & "', "      
        ' ##########################################################
				
				case "t4_":
				t4_sqlupdate=t4_sqlupdate& campo & "='" & valor & "', "
				t4_sqlinsertcampos=t4_sqlinsertcampos& campo& ", "
				t4_sqlinsertvalores=t4_sqlinsertvalores& "'" & valor & "', "
				
				
				CASE "t5_":
				t5_sqlupdate=t5_sqlupdate& campo & "='" & valor & "', "
				t5_sqlinsertcampos=t5_sqlinsertcampos& campo& ", "
				t5_sqlinsertvalores=t5_sqlinsertvalores& "'" & valor & "', "

				CASE "t6_":
				t6_sqlupdate=t6_sqlupdate& campo & "='" & valor & "', "
				t6_sqlinsertcampos=t6_sqlinsertcampos& campo& ", "
				t6_sqlinsertvalores=t6_sqlinsertvalores& "'" & valor & "', "
									
			END SELECT
				
	next
	
	'EJECUTAMOS INSTRUCCIONES
	
	't1_
	nombretabla="dn_risc_sustancias_vl"
	if request.form("hayt1")=0 then
		sqla="INSERT INTO " &nombretabla& " (" &t1_sqlinsertcampos& " id_sustancia) VALUES  (" &t1_sqlinsertvalores& id& ")"		
	else
		sqla="UPDATE " &nombretabla& " SET " &quitaultimoscar(t1_sqlupdate,2)& " WHERE id_sustancia= " &id
	end if
	'response.write("<br>sqla t1_: " &sqla)
	objconn1.execute(sqla)
	
	't0_
	nombretabla="dn_risc_sustancias"
	sqla="UPDATE " &nombretabla& " SET " &quitaultimoscar(t0_sqlupdate,2)& " WHERE id= " &id
	'response.write("<br>sqla t0_: " &sqla)
	objconn1.execute(sqla)

	't2_
	nombretabla="dn_risc_sustancias_iarc"
	if request.form("hayt2")=0 then
		sqla="INSERT INTO " &nombretabla& " (" &t2_sqlinsertcampos& " id_sustancia) VALUES  (" &t2_sqlinsertvalores& id& ")"		
	else
		sqla="UPDATE " &nombretabla& " SET " &quitaultimoscar(t2_sqlupdate,2)& " WHERE id_sustancia= " &id
	end if
	'response.write("<br>sqla t2_: " &sqla)
	objconn1.execute(sqla)	
	
	't3_
	nombretabla="dn_risc_sustancias_cancer_otras"
	if request.form("hayt3")=0 then
		sqla="INSERT INTO " &nombretabla& " (" &t3_sqlinsertcampos& " id_sustancia) VALUES  (" &t3_sqlinsertvalores& id& ")"		
	else
		sqla="UPDATE " &nombretabla& " SET " &quitaultimoscar(t3_sqlupdate,2)& " WHERE id_sustancia= " &id
	end if
	'response.write("<br>sqla t3_: " &sqla)
	objconn1.execute(sqla)

  ' ####################################################################################
	't9_
	nombretabla="dn_risc_sustancias_mama_cop"
	if request.form("hayt9")=0 then
		sqla="INSERT INTO " &nombretabla& " (" &t9_sqlinsertcampos& " id_sustancia) VALUES  (" &t9_sqlinsertvalores& id& ")"		
	else
		't9 tiene checkboxes, cuyo valor es 1 si estan marcados
		'pero si no estan marcados, son como disabled, no se envian
		'de modo que primero hacemos un update para ponerlos todos a 0, y en el siguiente paso, ya se pondran a 1 si es lo que nos han pedido
		sqlb="UPDATE dn_risc_sustancias_mama_cop SET cancer_mama=0 WHERE id_sustancia=" &id
		objconn1.execute(sqlb)

    ' Y ahora el update normal
		sqla="UPDATE " &nombretabla& " SET " &quitaultimoscar(t9_sqlupdate,2)& " WHERE id_sustancia= " &id
	end if
	'response.write("<br>sqla t9_: " &sqla)
	objconn1.execute(sqla)

  ' ####################################################################################
	't8_
  ' Como va en la misma tabla que t9, si hubo t9 no hay que insertar
	nombretabla="dn_risc_sustancias_mama_cop"
	if (request.form("hayt8")=0) and (request.form("hayt9")=1) then
		sqla="INSERT INTO " &nombretabla& " (" &t8_sqlinsertcampos& " id_sustancia) VALUES  (" &t8_sqlinsertvalores& id& ")"		
	else
    	
		' Y si no, el update normal
		sqla="UPDATE " &nombretabla& " SET " &quitaultimoscar(t8_sqlupdate,2)& " WHERE id_sustancia= " &id
	end if
	'response.write("<br>sqla t8_: " &sqla)
	objconn1.execute(sqla)
  ' ####################################################################################
	
	't4_
	nombretabla="dn_risc_sustancias_neuro_disruptor"
	if request.form("hayt4")=0 then
		sqla="INSERT INTO " &nombretabla& " (" &t4_sqlinsertcampos& " id_sustancia) VALUES  (" &t4_sqlinsertvalores& id& ")"		
	else
		'Sergio
		'Por si no marca ninguno, y había alguno marcado que lo borre
		sqla="UPDATE dn_risc_sustancias_neuro_disruptor SET nivel_disruptor='' WHERE id_sustancia= " &id
		objconn1.execute(sqla)
		
		sqla="UPDATE dn_risc_sustancias_neuro_disruptor SET fuente_neurotoxico='' WHERE id_sustancia= " &id
		objconn1.execute(sqla)
				
		sqla="UPDATE " &nombretabla& " SET " &quitaultimoscar(t4_sqlupdate,2)& " WHERE id_sustancia= " &id
	end if

	objconn1.execute(sqla)
	
	't5_
	nombretabla="dn_risc_sustancias_ambiente"
	if request.form("hayt5")=0 then
		sqla="INSERT INTO " &nombretabla& " (" &t5_sqlinsertcampos& " id_sustancia) VALUES  (" &t5_sqlinsertvalores& id& ")"		
	else
		't5 tiene checkboxes, cuyo valor es 1 si estan marcados
		'pero si no estan marcados, son como disabled, no se envian
		'de modo que primero hacemos un update para ponerlos todos a 0, y en el siguiente paso, ya se pondran a 1 si es lo que nos han pedido
		'Sergio -> Añado fuentes tpb
		sqlb="UPDATE dn_risc_sustancias_ambiente SET dano_calidad_aire=0, dano_ozono=0, dano_cambio_clima=0, directiva_aguas=0, cov=0, enlace_tpb=0, anchor_tpb=0, eper_agua=0, eper_aire=0, emisiones_atmosfera=0, seveso=0, fuentes_tpb='' WHERE id_sustancia= " &id
		objconn1.execute(sqlb)	
		sqla="UPDATE " &nombretabla& " SET " &quitaultimoscar(t5_sqlupdate,2)& " WHERE id_sustancia= " &id
	end if
	'response.write("<br>sqla t5_: " &sqla)
	objconn1.execute(sqla)

	't6_
	nombretabla="dn_risc_sustancias_salud"
	if request.form("hayt6")=0 then
		sqla="INSERT INTO " &nombretabla& " (" &t6_sqlinsertcampos& " id_sustancia) VALUES  (" &t6_sqlinsertvalores& id& ")"
  	objconn1.execute(sqla)				
	else
		't6 tiene checkboxes, cuyo valor es 1 si estan marcados
		'pero si no estan marcados, son como disabled, no se envian
		'de modo que primero hacemos un update para ponerlos todos a 0, y en el siguiente paso, ya se pondran a 1 si es lo que nos han pedido
		sqlb="UPDATE dn_risc_sustancias_salud SET cardiocirculatorio=0, rinyon=0, respiratorio=0, reproductivo=0, piel_sentidos=0, neuro_toxicos=0, musculo_esqueletico=0, sistema_inmunitario=0, higado_gastrointestinal=0, sistema_endocrino=0, embrion=0, cancer=0 WHERE id_sustancia= " &id
		objconn1.execute(sqlb)	

    if (quitaultimoscar(t6_sqlupdate,2) <> "") then
      ' Solo actualizamos si hay campos que meter
  		sqla="UPDATE " &nombretabla& " SET " &quitaultimoscar(t6_sqlupdate,2)& " WHERE id_sustancia= " &id
  	  objconn1.execute(sqla)
    end if

	end if
	response.write("<br>sqla t6_: " &sqla)

	
	cerrarconexion
end if

flashMsgCreate "Los datos se han actualizado.", "OK"
response.redirect "dn_sustanciaAD.asp?id=" &id
%>
