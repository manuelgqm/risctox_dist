<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"
	
	dim gente
  gente = Array(11,15,20,41,52,62,89,91,100,112,114,129,145,146,147,155,180,186,193,206,211,229,233,234,236,245,267,281,300,301,305,312,318,324,325,332,345,373,391,401,422,444,469,476,522,540,542,547,552,569,577,587,597,603,608,610,614,625,640,667,671,675,683,684,692,702,703,704,708,711,715,721,737,756,760,769,772,779,780,793,796,801,802,803,823,833,835,841,854,856,893,909,917,928,932,933,940,951,964,996,1015,1036,1051,1053,1060,1062,1086,1103,1148,1155,1162,1164,1171,1176,1185,1190,1204,1215,1216,1221,1229,1239,1255,1260,1291,1292,1293,1301,1304,1319,1322,1334,1335,1338,1340,1342,1354,1368,1376,1382,1386,1389,1392,1395,1398,1401,1405,1415,1427,1433,1437,1440,1468,1471,1480,1490,1528,1544,1549,1586,1603,1607,1618,1623,1625,1626,1651,1665,1694,1702,1706,1712,1725,1726,1730,1743,1744,1755,1756,1759,1761,1773,1784,1786,1790,1797,1803,1814,1820,1823,1827,1835,1836,1840,1847,1848,1852,1854,1855,1870,1871,1893,1896,1908,1910,1920,1925,1926,1927,1929,1931,1932,1944,1946,1947,1954,1970,1972,1982,1987,1988,1990,1995,1997,2001,2004,2013,2019,2030,2035,2039,2044,2049,2056,2060,2061,2068,2070,2072,2076,2078,2083,2087,2088,2094,2096,2103,2104,2110,2111,2113,2114,2127,2129,2130,2133,2137,2139,2140,2146,2157,2158,2174,2191,2192,2197,2204,2210,2213,2216,2220,2224,2225,2228,2238,2239,2244,2253,2259,2260,2261,2266,2278,2285,2286,2302,2313,2323,2324,2327,2332,2334,2343,2356,2361,2363,2374,2380,2390,2391,2395,2398,2400,2404,2406,2407,2408,2417,2418,2419,2420,2423,2428,2444,2447,2448,2452,2453,2461,2462,2463,2466,2473,2487,2495,2499,2500,2503,2504,2515,2516,2520,2521,2524,2525,2536,2541,2550,2554,2557,2577,2586,2608,2610,2614,2625,2627,2628,2633,2637,2643,2651,2656,2658,2660,2663,2672,2675,2694,2699,2700,2713,2722,2724,2745,2748,2750,2755,2769,2770,2794,2814,2824,2826,2847,2852,2863,2864,2872,2875,2878,2882,2887,2928,2938,2944,2966,2969,2981,2997,3022,3052,3094,3105,3113,3130,3148,3159,3162,3178,3184,3207,3258,3270,3281,3287,3299,3300,3339,3353,3355,3356,3365,3377,3382,3385,3392,3402,3412,3413,3423,3427,3432,3435,3455,3456,3460,3462,3466,3469,3481,3488,3489,3496,3497,3501,3506,3521,3528,3531,3546,3548,3567,3624,3627,3646,3655,3661,3671,3672,3676,3683,3698,3710,3714,3728,3741,3771,3786,3788,3811,3813,3820,3824,3827,3828,3837,3851,3861,3867,3869,3874,3876,3883,3884,3885,3890,3891,3895,3899,3905,3911,3916,3918,3920,3932,3933,3938,3939,3947,3953,3961,3970,3972,3986,3990)
	
	idpagina = 961	'-- Evalua lo que usas: inicio
	idpagina2 = 963	'-- Evalua lo que usas: auto_portada
	idpagina3 = 964	'-- Evalua lo que usas: herramienta
	
  response.write "<table style='font-family: Verdana; font-size: 8pt; border: 1 solid #000000'><tr><td>ELEGIBLES AÑADIDOS A LAS VISITAS DE EVALUA LO QUE USAS (pag=961, 963 y 964 sólo en negrita)</td></tr></table><br>"
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
	  if fec_hor<cdate("8/11/2006") then 
	  	nAleatorio = Int(85) * Rnd
	  	fec_hor = dateadd("d",nAleatorio,cdate("8/11/2006"))
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
	  
	  nAleatorio = Int(120) * Rnd
	  fec_hor2 = dateadd("s",nAleatorio,fec_hor2)	'-- suma un valor de 0 a 120 segundos
	  fecha2 = FormatDateTime(fec_hor2,2)
	  hora2 = FormatDateTime(fec_hor2,3)
	  
	  orden2 = "INSERT INTO WEBISTAS_VISITAS (fecha,hora,IP,navegador,idpagina,idgente) VALUES ('"&fecha2&"','"&hora2&"','"&IP&"','IE2',"&idpagina2&","&gente(i)&")"
	  Set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
	  Set objRecordset2 = OBJConnection.Execute(orden2)
	  
	  if par=1 then	'-- a uno de cada 2 le añado la visita a la página 964
	  	nAleatorio = Int(120) * Rnd
	  	fec_hor2 = dateadd("s",nAleatorio,fec_hor2)	'-- suma un valor de 0 a 120 segundos
	  	fecha2 = FormatDateTime(fec_hor2,2)
	  	hora2 = FormatDateTime(fec_hor2,3)
	  	orden2 = "INSERT INTO WEBISTAS_VISITAS (fecha,hora,IP,navegador,idpagina,idgente) VALUES ('"&fecha2&"','"&hora2&"','"&IP&"','IE2',"&idpagina3&","&gente(i)&")"
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
	
	
