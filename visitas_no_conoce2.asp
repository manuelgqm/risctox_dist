<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"
	
	dim gente
  gente = Array(1005,1006,1007,1010,1011,1013,1016,1019,1020,1021,1034,1035,1036,1039,1040,1043,1044,1045,1047,1050,1051,1052,1053,1061,1062,1063,1064,1066,1067,1071,1075,1076,1077,1086,1087,1088,1089,1094,1098,1100,1109,1110,1114,1119,1123,1124,1125,1128,1130,1145,1146,1148,1149,1152,1153,1154,1156,1160,1161,1162,1163,1164,1165,1166,1167,1168,1169,1170,1171,1172,1173,1174,1175,1176,1177,1179,1182,1183,1185,1187,1189,1190,1191,1193,1194,1197,1198,1202,1204,1207,1210,1211,1212,1213,1215,1217,1220,1221,1224,1225,1229,1230,1231,1234,1237,1238,1239,1240,1244,1245,1246,1254,1255,1258,1260,1264,1269,1270,1284,1291,1292,1293,1294,1296,1298,1301,1303,1308,1309,1312,1315,1316,1317,1318,1319,1320,1322,1323,1324,1329,1331,1334,1335,1338,1340,1341,1342,1343,1344,1345,1347,1351,1354,1361,1362,1363)
	idpagina = 521	'-- Conoces lo que usas
	
  response.write "<table style='font-family: Verdana; font-size: 8pt; border: 1 solid #000000'><tr><td>ELEGIBLES AÑADIDOS A LAS VISITAS DE CONOCES LO QUE USAS</td></tr></table><br>"
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
	
	
