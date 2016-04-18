<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"
	
	dim gente
  gente = Array(3487,3488,3489,3494,3496,3497,3501,3502,3503,3505,3506,3508,3510,3511,3513,3515,3521,3528,3529,3530,3531,3534,3536,3543,3546,3548,3552,3554,3557,3561,3567,3570,3573,3582,3604,3609,3624,3627,3630,3646,3647,3649,3653,3655,3656,3657,3658,3659,3661,3663,3669,3671,3672,3676,3677,3678,3680,3681,3682,3683,3685,3686,3689,3693,3703,3704,3705,3708,3710,3711,3712,3713,3714,3715,3716,3717,3718,3719,3720,3721,3722,3723,3724,3728,3731,3734,3737,3738,3741,3746,3747,3749,3751,3756,3764,3766,3768,3770,3771,3777,3779,3782,3783,3784,3786,3787,3788,3790,3791,3793,3795,3799,3801,3802,3803,3805,3807,3808,3809,3810,3811,3812,3813,3814,3816,3817,3818,3820,3821,3823,3824,3825,3826,3827,3828,3829,3830,3831,3832,3833,3834,3835,3836,3837,3839,3840,3841,3842,3844,3847,3848,3849,3850,3852,3853,3854,3855,3856,3857,3858,3859,3861,3862,3863,3864,3866,3867,3868,3869,3870,3871,3872,3873,3874,3876,3877,3878,3879,3880,3881,3882,3883,3884,3885,3886,3887,3888,3889,3890,3891,3892,3893,3894,3895,3896,3897,3898,3899,3900,3902,3903,3904,3905,3906,3907,3908,3909,3910,3911,3912,3913,3915,3916,3917,3918,3919,3920,3921,3922,3924,3927,3929,3930,3932,3933,3934,3935,3936,3937,3938,3939,3940,3941,3942,3945,3946,3947,3948,3949,3950,3951,3952,3953,3954,3955,3956,3959,3961,3962,3963,3964,3965,3966,3967,3968,3969,3970,3971,3972,3973,3974,3975,3976,3977,3978,3979,3980,3981,3984,3985,3986,3987,3988,3989,3990,3991,3992,3993,3994,3995,3997,3998,3999,4000,4002,4003,4004,4005,4006,4007,4008,4009,4010,4011,4012,4013,4014,4015,4016,4017)
	idpagina = 562	'-- Asesoría
	
  response.write "<table style='font-family: Verdana; font-size: 8pt; border: 1 solid #000000'><tr><td>ELEGIBLES AÑADIDOS A LAS VISITAS DEL OBSERVATORIO: Asesoría (pag=562)</td></tr></table><br>"
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
	
	
