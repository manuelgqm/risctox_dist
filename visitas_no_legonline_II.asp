<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"
	
	dim gente
  gente = Array(4400,4417,4429,4433,4438,4443,4449,4450,4453,4458,4465,4467,4470,4473,4475,4480,4482,4483,4484,4486,4500,4512,4515,4518,4522,4525,4526,4528,4530,4531,4537,4538,4545,4546,4548,4552,4563,4572,4580,4581,4582,4587,4595,4596,4600,4604,4617,4618,4620,4624,4626,4629,4638,4650,4652,4656,4663,4674,4686,4689,4690,4693,4702,4727,4741,4747,4750,4761,4766,4770,4776,4789,4802,4819,4830,4831,4863,4871,4872,4894,4901,4903,4917,4939,4951,4958,4960,4963,4964,4966,4970,4994,4997,5033,5061,5110,5177,5187,5189,5200,5206,5215,5224,5236,5259,5264,5273,5274,5276,5278,5280,5314,5340,5345,5350,5356,5370,5378,5386,5389,5400,5402,5404,5411,5413,5419,5425,5441,5450,5458,5467,5476,5481,5482,5483,5484,5485,5489,5490,5507,5577,5578,5587,5588,5590,5591,5618,5632,5652,5666,5680,5687,5690,5712,5714,5718,5748,5785,5792,5799,5819,5825,5826,5827,5828,5829,5830,5836,5842,5851,5853,5880,5896,5899,5901,5908,5917,5919,5921,5923,5934,5936,5942,5943,5948,5951,5961,5963,5965,5980,5990,5997,5999,6007,6019,6028,6029,6030,6031,6039,6063,6073,6103,6106,6108,6117,6118,6119,6126,6154,6167,6170,6171,6183,6187,6200,6209,6219,6222,6233,6234,6243,6278,6290,6295,6300,6303,6313,6315,6323,6326,6342,6357,6358,6368,6376,6379,6397,6399,6402,6409,6424,6433,6458,6464,6466,6470,6476,6487,6488,6489,6494,6496,6514,6519,6521,6525,6526,6533,6547,6551,6558,6560,6562,6577,6578,6583,6588,6604,6605,6613,6622,6625,6628,6632,6634,6652,6655,6660,6664,6666,6682,6686,6692,6715,6724,6735,6740,6744,6757,6761,6768,6777,6778,6779,6788,6789,6791,6811,6819,6823,6825,6839,6856,6857,6858,6864,6871,6880,6907,6917,6923,6930,6942,6954,6965,6967,6969,6971,6973,7013,7020,7023,7031,7036,7040,7064,7065,7107,7109,7118,7140,7141,7156,7169,7173,7174,7198,7209,7218,7219,7220,7224,7227,7236,7242,7245,7270,7288,7308,7312,7320,7365,7370,7373,7385,7416,7458,7461,7467,7469,7473,7484,7494,7497,7500,7524,7538,7561,7568,7575,7604,7632,7644,7651,7653,7664,7668,7671,7675,7691,7695,7708,7736,7741,7750,7751,7766,7788,7795,7808,7815,7824,7826,7828,7833,7836,7843,7847,7861,7869,7874,7878,7890,7896,7908,7909,7922,7944,7953,7954,7969,7979,7990,7999,8000,8006,8016,8032,8034,8052,8095,8101,8117,8134,8137,8159,8186,8208,8230,8240,8242,8258,8330,8338,8351,8358,8412,8415,8431,8435,8456,8467,8473,8480,8498,8500,8504,8532,8541,8555,8556,8568,8596,8600,8604,8605,8608,8613,8625,8634,8645,8650,8660,8661,8667,8668,8713,8717,8746,8764,8780,8786,8788,8794,8795,8798,8802,8817,8818,8828,8845,8872,8874,8875,8885,8905,8914,8961,8981,8991,9015,9026,9031,9036,9040,9044,9048,9052,9056,9064,9071,9072,9096,9097,9111,9117,9158,9162,9177,9181,9192,9205,9231,9260,9321,9331,9335,9336,9338,9386,9396,9443,9445,9457,9528,9556,9587,9602,9603,9618,9662,9706)
	
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
	
	
