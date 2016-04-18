<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"
	
	dim gente
  gente = Array(603,606,607,608,610,611,613,614,616,618,620,625,626,628,631,636,637,638,640,642,644,645,646,650,652,653,654,657,658,659,660,664,666,667,671,673,675,677,679,682,685,686,687,688,689,690,692,697,698,699,700,702,704,705,706,707,708,709,710,711,713,716,717,720,721,722,723,725,728,731,732,734,736,737,738,740,743,747,750,751,753,754,755,756,759,760,765,767,768,769,770,772,774,775,779,781,783,785,787,789,791,794,796,800,801,803,804,805,806,809,813,816,817,823,824,825,826,828,830,831,832,835,836,841,842,843,849,850,851,852,853,854,856,857,860,864,866,867,868,869,875,877,879,882,885,888,894,895,896,897,909,917,922,927,928,929,933,935,940,944,945,948,951,954,957,961,964,967,974,975,977,980,982,984,987,988,990,993,999,1007,1010,1013,1018,1019,1020,1021,1034,1035,1036,1043,1044,1045,1047,1051,1052,1053,1060,1061,1064,1066,1067,1071,1075,1076,1077,1086,1087,1088,1089,1094,1098,1099,1100,1103,1109,1110,1114,1119,1123,1124,1128,1130,1136,1146,1147,1148,1149,1153,1154,1156,1158,1161,1162,1163,1164,1165,1167,1168,1169,1170,1171,1172,1173,1174,1175,1178,1179,1183,1185,1187,1190,1191,1193,1194,1197,1198,1202,1207,1210,1211,1212,1213,1215,1216,1217,1220,1221,1224,1225,1229,1230,1231,1237,1238,1240,1244,1245,1246,1252,1254,1255,1258,1260,1264,1269,1270,1280,1284,1288,1291,1292,1293,1294,1296,1298,1301,1303,1304,1306,1308,1309,1310,1312,1315,1316,1317,1318,1319,1320,1322,1323,1324,1329,1331,1332,1333,1334,1335,1338,1341,1342,1343,1344,1345,1347,1351)
	idpagina = 576	'-- ALTERNATIVAS: inicio
	
  response.write "<table style='font-family: Verdana; font-size: 8pt; border: 1 solid #000000'><tr><td>ELEGIBLES AÑADIDOS A LAS VISITAS DE ALTERNATIVAS INICIO (pag=576)</td></tr></table><br>"
  response.write "<table style='font-family: Verdana; font-size: 8pt; border: 1 solid #000000'><tr><td style='border-bottom: 1 solid #000000'>ORDEN</td><td style='border-bottom: 1 solid #000000'>ID_GENTE</td><td style='border-bottom: 1 solid #000000'>NOMBRE</td><td style='border-bottom: 1 solid #000000'>APELLIDOS</td><td style='border-bottom: 1 solid #000000'>DIRECCIÓN IP</td><td style='border-bottom: 1 solid #000000'>FECHA ALTA</td><td style='border-bottom: 1 solid #000000'>FECHA VISITA FICTICIA</td></tr>"
  
	for i=1 to ubound(gente)
	
	  orden = "SELECT IP,fec_hor,nombre,apellidos FROM ECOINFORMAS_GENTE WHERE idgente=" & gente(i)
	  set objRecordset = Server.CreateObject ("ADODB.Recordset")
	  set objRecordset = OBJConnection.Execute(orden)
	  IP = objRecordset("IP")
	  fec_hor = objRecordset("fec_hor")
	  
	  'fecha = FormatDateTime(fec_hor,2)
	  'hora = FormatDateTime(fec_hor,3)
	  
	  'fecha2 = dateadd("n",1,fec_hor)	'--suma 1 minuto
	  nAleatorio = Int(120) * Rnd
	  fec_hor2 = dateadd("s",nAleatorio,fec_hor)	'-- suma un valor de 0 a 120 segundos
	  fecha2 = FormatDateTime(fec_hor2,2)
	  hora2 = FormatDateTime(fec_hor2,3)
	  
	  nombre = objRecordset("nombre")
	  apellidos = objRecordset("apellidos")
	  
	  orden2 = "INSERT INTO WEBISTAS_VISITAS (fecha,hora,IP,navegador,idpagina,idgente) VALUES ('"&fecha2&"','"&hora2&"','"&IP&"','IE2',"&idpagina&","&gente(i)&")"
	  Set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
	  Set objRecordset2 = OBJConnection.Execute(orden2)
	  
	  response.write "<tr><td>" & cstr(i) & "</td><td>" & cstr(gente(i)) & "</td><td>" & nombre & "</td><td>" & apellidos & "</td><td>" & IP & "</td><td>" & fec_hor & "</td><td>" & fec_hor2 & "</td></tr>"
	  
	next
  response.write "</table>"

	%>
	
	
