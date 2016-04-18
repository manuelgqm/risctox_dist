<!--#include file="eco_conexion.asp"-->
<%
Server.ScriptTimeout = 1000000

dim usuario(106)

usuario(1  ) = 219
usuario(2  ) = 218
usuario(3  ) = 217
usuario(4  ) = 211
usuario(5  ) = 206
usuario(6  ) = 203
usuario(7  ) = 201
usuario(8  ) = 198
usuario(9  ) = 193
usuario(10 ) = 190
usuario(11 ) = 189
usuario(12 ) = 188
usuario(13 ) = 186
usuario(14 ) = 181
usuario(15 ) = 180
usuario(16 ) = 171
usuario(17 ) = 170
usuario(18 ) = 161
usuario(19 ) = 160
usuario(20 ) = 159
usuario(21 ) = 158
usuario(22 ) = 156
usuario(23 ) = 155
usuario(24 ) = 154
usuario(25 ) = 153
usuario(26 ) = 147
usuario(27 ) = 146
usuario(28 ) = 144
usuario(29 ) = 141
usuario(30 ) = 139
usuario(31 ) = 138
usuario(32 ) = 137
usuario(33 ) = 136
usuario(34 ) = 135
usuario(35 ) = 134
usuario(36 ) = 133
usuario(37 ) = 130
usuario(38 ) = 129
usuario(39 ) = 128
usuario(40 ) = 127
usuario(41 ) = 124
usuario(42 ) = 121
usuario(43 ) = 119
usuario(44 ) = 117
usuario(45 ) = 116
usuario(46 ) = 115
usuario(47 ) = 111
usuario(48 ) = 100
usuario(49 ) = 93
usuario(50 ) = 91
usuario(51 ) = 89
usuario(52 ) = 88
usuario(53 ) = 86
usuario(54 ) = 80
usuario(55 ) = 79
usuario(56 ) = 78
usuario(57 ) = 77
usuario(58 ) = 75
usuario(59 ) = 74
usuario(60 ) = 73
usuario(61 ) = 72
usuario(62 ) = 70
usuario(63 ) = 69
usuario(64 ) = 68
usuario(65 ) = 67
usuario(66 ) = 66
usuario(67 ) = 64
usuario(68 ) = 62
usuario(69 ) = 61
usuario(70 ) = 58
usuario(71 ) = 57
usuario(72 ) = 55
usuario(73 ) = 54
usuario(74 ) = 52
usuario(75 ) = 51
usuario(76 ) = 50
usuario(77 ) = 49
usuario(78 ) = 48
usuario(79 ) = 47
usuario(80 ) = 46
usuario(81 ) = 45
usuario(82 ) = 44
usuario(83 ) = 43
usuario(84 ) = 41
usuario(85 ) = 40
usuario(86 ) = 39
usuario(87 ) = 38
usuario(88 ) = 36
usuario(89 ) = 35
usuario(90 ) = 34
usuario(91 ) = 33
usuario(92 ) = 32
usuario(93 ) = 31
usuario(94 ) = 25
usuario(95 ) = 24
usuario(96 ) = 23
usuario(97 ) = 22
usuario(98 ) = 19
usuario(99 ) = 18
usuario(100) = 16
usuario(101) = 15
usuario(102) = 13
usuario(103) = 11
usuario(104) = 8
usuario(105) = 7
usuario(106) = 6

num_usuarios = Ubound(usuario)

'for i=1 to num_usuarios
' response.write i&": "&usuario(i)&"<br>"
'next


i=1
orden = "SELECT idvisita FROM WEBISTAS_VISITAS WHERE idgente=4 AND idvisita>900000 ORDER BY idvisita"
Set objRecordset = Server.CreateObject ("ADODB.Recordset")
Set objRecordset = OBJConnection.Execute(orden)

do while not objRecordset.eof
	orden2 = "UPDATE WEBISTAS_VISITAS SET idgente="&usuario(i)&" WHERE idvisita="&objRecordset("idvisita")
	Set objRecordset2 = OBJConnection.Execute(orden2)
	response.write objRecordset("idvisita")&": "&usuario(i)&"<br>"

	i=i+1
	if i>106 then i=1
	objRecordset.movenext
loop



%>
 