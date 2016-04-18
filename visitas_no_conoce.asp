<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"
	
	dim gente
  gente = Array(145,196,200,212,226,229,231,233,234,235,238,239,240,242,244,245,248,267,273,274,275,278,279,281,283,284,288,298,299,301,306,309,312,315,316,321,324,325,331,332,335,336,337,338,341,343,345,349,350,353,357,360,361,363,364,369,371,374,376,380,382,391,392,393,394,397,399,401,407,409,410,413,414,435,440,443,444,446,450,456,459,470,471,473,480,485,490,491,492,495,498,513,517,520,521,522,523,527,540,542,545,547,552,553,559,560,564,566,569,576,577,583,586,587,589,592,595,597,603,604,606,610,611,612,613,614,615,616,618,619,620,626,628,631,635,636,637,639,640,642,645,646,648,650,652,654,657,658,659,660,666,667,671,673,674,675,676,677,682,683,684,685,686,687,688,690,692,694,697,698,699,700,702,704,706,707,708,709,710,713,715,716,717,721,722,723,725,728,731,732,734,736,737,740,743,746,747,750,751,753,755,756,759,765,768,769,770,771,772,774,775,777,780,781,785,787,789,790,791,794,796,800,801,804,805,806,809,816,823,824,825,828,830,831,832,833,835,836,841,842,849,850,852,853,856,857,860,864,866,867,868,872,875,878,882,888,894,895,897,907,909,917,918,922,928,929,935,940,944,945,946,948,954,957,961,964,967,975,977,979,980,982,984,987,988,993,994,996,999)
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
	
	
