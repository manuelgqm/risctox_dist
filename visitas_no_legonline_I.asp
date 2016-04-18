<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"
	
	dim gente
  gente = Array(11,16,25,38,43,51,57,62,93,97,118,128,130,137,139,145,146,170,193,218,226,244,288,300,324,325,335,338,360,361,380,392,401,413,438,450,460,476,482,520,527,566,587,589,610,615,630,635,638,645,646,652,653,654,668,671,681,692,697,712,736,754,759,765,768,770,774,796,801,813,823,833,836,841,842,852,872,877,888,890,894,895,897,954,964,975,987,994,996,1006,1015,1016,1034,1045,1060,1063,1064,1066,1116,1119,1123,1130,1131,1145,1147,1156,1160,1162,1166,1171,1173,1174,1177,1190,1204,1207,1209,1217,1234,1244,1264,1301,1303,1306,1310,1320,1322,1342,1347,1370,1389,1393,1413,1430,1432,1434,1438,1446,1451,1452,1473,1474,1490,1495,1500,1510,1518,1521,1523,1525,1532,1574,1602,1607,1610,1626,1635,1645,1649,1694,1699,1703,1720,1729,1732,1738,1758,1763,1769,1770,1780,1781,1796,1801,1803,1809,1815,1820,1826,1840,1843,1847,1851,1858,1859,1861,1865,1868,1879,1882,1887,1890,1891,1893,1894,1895,1900,1906,1908,1910,1913,1914,1920,1922,1924,1925,1934,1946,1952,1956,1957,1972,1978,1982,1993,1995,2001,2004,2019,2020,2026,2032,2037,2040,2047,2067,2068,2077,2082,2092,2093,2095,2100,2102,2103,2106,2118,2119,2122,2138,2139,2142,2147,2151,2152,2156,2166,2169,2177,2178,2180,2186,2189,2192,2200,2203,2210,2214,2223,2255,2256,2259,2271,2279,2293,2301,2307,2312,2313,2317,2318,2320,2323,2330,2333,2337,2343,2348,2351,2356,2358,2359,2366,2367,2370,2377,2386,2391,2394,2404,2406,2411,2412,2414,2417,2418,2425,2446,2447,2467,2470,2472,2487,2491,2495,2503,2517,2520,2524,2525,2531,2543,2546,2563,2569,2577,2578,2580,2593,2601,2605,2613,2614,2619,2627,2630,2633,2634,2646,2647,2650,2654,2665,2675,2688,2702,2703,2706,2708,2709,2711,2712,2714,2719,2741,2744,2754,2769,2782,2793,2794,2812,2823,2824,2828,2845,2862,2863,2870,2871,2873,2875,2882,2886,2899,2904,2915,2960,2962,2965,2966,2993,2999,3000,3048,3052,3055,3077,3117,3123,3159,3184,3193,3202,3222,3251,3256,3268,3269,3281,3287,3292,3298,3299,3321,3326,3330,3338,3340,3357,3365,3382,3394,3395,3403,3410,3417,3436,3441,3448,3458,3471,3472,3478,3479,3481,3482,3483,3494,3496,3508,3510,3528,3546,3568,3582,3627,3653,3663,3689,3693,3698,3710,3712,3716,3728,3747,3756,3783,3788,3793,3816,3823,3828,3836,3850,3852,3858,3859,3861,3864,3879,3884,3890,3894,3895,3903,3906,3915,3922,3925,3930,3940,3946,3948,3954,3966,3981,3984,4005,4011,4012,4016,4018,4027,4032,4037,4051,4072,4087,4090,4091,4093,4098,4101,4106,4110,4111,4133,4137,4138,4139,4141,4144,4147,4148,4149,4163,4164,4165,4167,4174,4175,4182,4194,4202,4209,4210,4214,4224,4228,4237,4246,4248,4249,4254,4255,4258,4267,4295,4299,4310,4311,4314,4344,4350,4353,4356,4359,4365,4368,4371,4373,4394,4395,4397,4399)
	
	idpagina = 710	'-- Leg_online: inicio
	idpagina2 = 711	'-- Leg_online: buscador
	
  response.write "<table style='font-family: Verdana; font-size: 8pt; border: 1 solid #000000'><tr><td>ELEGIBLES AÑADIDOS A LAS VISITAS DE LEG. ONLINE (pag=710 y 711 sólo en negrita)</td></tr></table><br>"
  response.write "<table style='font-family: Verdana; font-size: 8pt; border: 1 solid #000000'><tr><td style='border-bottom: 1 solid #000000'>ORDEN</td><td style='border-bottom: 1 solid #000000'>ID_GENTE</td><td style='border-bottom: 1 solid #000000'>NOMBRE</td><td style='border-bottom: 1 solid #000000'>APELLIDOS</td><td style='border-bottom: 1 solid #000000'>DIRECCIÓN IP</td><td style='border-bottom: 1 solid #000000'>FECHA ALTA</td><td style='border-bottom: 1 solid #000000'>FECHA VISITA FICTICIA</td></tr>"
  par = 1
	for i=1 to ubound(gente)
	  par = -1*par
	  orden = "SELECT IP,fec_hor,nombre,apellidos FROM ECOINFORMAS_GENTE WHERE idgente=" & gente(i)
	  set objRecordset = Server.CreateObject ("ADODB.Recordset")
	  set objRecordset = OBJConnection.Execute(orden)
	  IP = objRecordset("IP")
	  fec_hor = objRecordset("fec_hor")
	  fec_hor = cdate(fec_hor)
	  if fec_hor<cdate("16/10/2006") then 
	  	nAleatorio = Int(100) * Rnd
	  	fec_hor = dateadd("d",nAleatorio,cdate("16/10/2006"))
	  end if
	  
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
	  
	  if par=1 then	'-- a uno de cada 2 le añado la visita a la página 711
	  	nAleatorio = Int(120) * Rnd
	  	fec_hor2 = dateadd("s",nAleatorio,fec_hor2)	'-- suma un valor de 0 a 120 segundos
	  	fecha2 = FormatDateTime(fec_hor2,2)
	  	hora2 = FormatDateTime(fec_hor2,3)
	  	orden2 = "INSERT INTO WEBISTAS_VISITAS (fecha,hora,IP,navegador,idpagina,idgente) VALUES ('"&fecha2&"','"&hora2&"','"&IP&"','IE2',"&idpagina2&","&gente(i)&")"
	  	Set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
	  	Set objRecordset2 = OBJConnection.Execute(orden2)
	  end if
	  
	  if par=1 then
	  	response.write "<tr><td>" & cstr(i) & "</td><td>" & cstr(gente(i)) & "</td><td><b>" & nombre & "</b></td><td>" & apellidos & "</td><td>" & IP & "</td><td>" & fec_hor & "</td><td>" & fec_hor2 & "</td></tr>"
	  else
	    	response.write "<tr><td>" & cstr(i) & "</td><td>" & cstr(gente(i)) & "</td><td>" & nombre & "</td><td>" & apellidos & "</td><td>" & IP & "</td><td>" & fec_hor & "</td><td>" & fec_hor2 & "</td></tr>"
	  end if
	  
	next
  response.write "</table>"

	%>
	
	
