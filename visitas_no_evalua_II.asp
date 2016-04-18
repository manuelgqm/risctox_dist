<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"
	
	dim gente
  	gente = Array(4010,4013,4015,4022,4024,4028,4032,4047,4057,4065,4072,4081,4104,4106,4107,4125,4127,4128,4136,4146,4153,4154,4156,4164,4176,4177,4179,4181,4183,4188,4189,4210,4212,4223,4234,4238,4239,4249,4258,4261,4266,4268,4271,4273,4296,4297,4303,4305,4312,4317,4325,4332,4349,4358,4364,4366,4367,4387,4408,4413,4426,4431,4441,4447,4455,4457,4463,4465,4466,4467,4479,4487,4488,4489,4506,4507,4513,4520,4522,4524,4528,4529,4537,4542,4557,4565,4573,4574,4579,4589,4590,4597,4598,4599,4610,4611,4614,4617,4623,4625,4626,4636,4639,4642,4650,4659,4660,4663,4665,4673,4686,4706,4710,4727,4740,4765,4766,4777,4783,4785,4790,4798,4801,4803,4811,4813,4819,4820,4837,4839,4841,4861,4873,4880,4885,4892,4893,4898,4904,4929,4942,4943,4945,4960,4979,5015,5020,5038,5043,5057,5081,5090,5091,5124,5128,5139,5145,5154,5156,5163,5178,5205,5206,5220,5222,5229,5260,5264,5269,5274,5288,5300,5314,5328,5345,5350,5351,5356,5386,5400,5402,5413,5415,5418,5426,5438,5441,5448,5449,5450,5452,5475,5476,5481,5483,5487,5489,5503,5505,5511,5545,5555,5564,5576,5578,5580,5582,5583,5585,5599,5600,5616,5618,5632,5634,5645,5653,5660,5663,5665,5673,5714,5718,5725,5738,5740,5749,5757,5770,5792,5827,5830,5836,5839,5859,5891,5907,5917,5920,5939,5943,5955,5968,5970,5975,5990,5997,6012,6014,6024,6025,6042,6044,6076,6083,6094,6117,6119,6143,6149,6150,6156,6169,6171,6186,6203,6220,6238,6253,6271,6290,6314,6322,6323,6325,6355,6376,6384,6389,6393,6400,6405,6407,6408,6420,6424,6450,6458,6463,6469,6479,6483,6485,6494,6512,6521,6526,6538,6539,6557,6570,6572,6574,6578,6590,6606,6617,6625,6627,6628,6634,6637,6652,6679,6682,6685,6699,6735,6739,6741,6768,6784,6789,6790,6791,6800,6817,6871,6878,6880,6882,6908,6914,6930,6933,6956,6974,6990,6992,7002,7013,7016,7025,7029,7048,7087,7089,7098,7104,7106,7109,7117,7121,7136,7146,7181,7192,7215,7224,7240,7246,7248,7263,7267,7275,7279,7300,7329,7334,7337,7351,7355,7365,7370,7388,7389,7393,7395,7406,7408,7416,7438,7444,7446,7453,7454,7469,7483,7485,7486,7488,7500,7502,7503,7530,7533,7540,7561,7571,7581,7585,7595,7597,7600,7609,7613,7631,7662,7695,7708,7713,7728,7733,7741,7746,7754,7766,7768,7788,7793,7803,7808,7824,7843,7847,7865,7884,7904,7909,7917,7923,7931,7944,7954,7956,7959,7965,7974,7991,7996,8012,8013,8016,8019,8022,8034,8038,8047,8052,8060,8078,8093,8101,8130,8147,8148,8160,8169,8197,8204,8212,8220,8228,8240,8246,8247,8258,8272,8275,8303,8343,8345,8347,8385,8404,8422,8448,8460,8463,8465,8500,8502,8547,8548,8550,8552,8555,8558,8593,8604,8608,8613,8615,8616,8619,8634,8650,8652,8654,8660,8663,8678,8730,8737,8741,8763,8788,8792,8802,8817,8827,8831,8832,8857,8866,8868,8869,8885,8886,8894,8913,8921,8926,8927,8928,8936,8944,8953,8991,9013,9019,9027,9028,9035,9041,9042,9051,9053,9056,9062,9070,9072,9074,9076,9111,9112,9122,9130,9145,9197,9206,9215,9216,9223,9231,9253,9257,9294,9303,9341,9355,9369,9396,9446,9468,9479,9502,9575,9655)
	
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
	
	
