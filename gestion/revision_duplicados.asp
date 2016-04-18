<!--#include file="dn_fun_comunes.asp"-->
<!--#include file="dn_auten.inc"-->
<!--#include file="adovbs.inc"-->
<!--#include file="dn_conexion.asp"-->

<%


valor1=983955
valor2=983956


sql = "update dn_alter_ficheros_por_sustancias set id_sustancia="&valor1& " where id_sustancia="&valor2
response.write sql & "<br>"
objconn1.execute(sql)

sql = "update dn_risc_nombres_comerciales set id_sustancia="&valor1& " where id_sustancia="&valor2
response.write sql & "<br>"
objconn1.execute(sql)

sql = "update dn_risc_sustancias_ambiente set id_sustancia="&valor1& " where id_sustancia="&valor2
response.write sql & "<br>"
objconn1.execute(sql)

sql = "update dn_risc_sustancias_cancer_otras set id_sustancia="&valor1& " where id_sustancia="&valor2
response.write sql & "<br>"
objconn1.execute(sql)

sql = "update dn_risc_sustancias_iarc set id_sustancia="&valor1& " where id_sustancia="&valor2
response.write sql & "<br>"
objconn1.execute(sql)

sql = "update dn_risc_sustancias_mama_cop set id_sustancia="&valor1& " where id_sustancia="&valor2
response.write sql & "<br>"
objconn1.execute(sql)

sql = "update dn_risc_sustancias_neuro_disruptor set id_sustancia="&valor1& " where id_sustancia="&valor2
response.write sql & "<br>"
objconn1.execute(sql)

sql = "update dn_risc_sustancias_por_companias set id_sustancia="&valor1& " where id_sustancia="&valor2
response.write sql & "<br>"
objconn1.execute(sql)

sql = "update dn_risc_sustancias_por_enfermedades set id_sustancia="&valor1& " where id_sustancia="&valor2
response.write sql & "<br>"
objconn1.execute(sql)

sql = "update dn_risc_sustancias_por_grupos set id_sustancia="&valor1& " where id_sustancia="&valor2
response.write sql & "<br>"
objconn1.execute(sql)

sql = "update dn_risc_sustancias_por_usos set id_sustancia="&valor1& " where id_sustancia="&valor2
response.write sql & "<br>"
objconn1.execute(sql)

sql = "update dn_risc_sustancias_salud set id_sustancia="&valor1& " where id_sustancia="&valor2
response.write sql & "<br>"
objconn1.execute(sql)

sql = "update dn_risc_sustancias_vl set id_sustancia="&valor1& " where id_sustancia="&valor2
response.write sql & "<br>"
objconn1.execute(sql)

sql = "delete from dn_risc_sustancias where id="&valor2
response.write sql & "<br>"
objconn1.execute(sql)


%>
