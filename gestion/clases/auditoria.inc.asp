<!--#include file="funciones.inc.asp"-->
<%
'Fecha: 25/01/2010
'Desar: SPL - Sistemas de informaci�n
'Autor: Jos� Sanchis P�rez de Le�n
'Descr: Clase Auditor�a

class Auditoria
	dim arrCampos ' Array de campos
	dim arrValores ' Array de valores

	private sub Class_Initialize()
		arrCampos = array(	"usuario",_
							"fecha",_
							"accion",_
							"entidad",_
							"navegador",_
							"descripcion",_
							"ip")
		redim arrValores(ubound(arrCampos))
	end sub

	' registra en la BD, la �ltima acci�n llevada a cabo
	public sub registra()
		sql="INSERT INTO spl_auditoria ("
		for i=0 to (Me.numeroCampos-1)
			sql = sql & Me.getName(i) & ","
		next
		sql = sql & Me.getName(Me.numeroCampos)
		sql = sql & ") VALUES("
		for i=0 to Me.numeroCampos-1
			campo = trim(Me.getProperty(Me.getName(i)))
			if (Len(campo) > 5000) then campo = Left(campo,5000)
			if Me.getName(i) = "fecha" then
				campo = FechaGenerica(campo)
			end if
			sql = sql & "'" & addslashes(trim(campo)) & "',"

		next
		sql = sql & "'" & addslashes(trim(Me.getProperty(Me.getName(Me.numeroCampos)))) & "'"
		sql=sql & ")"

'response.write("prueba:" & prueba_tonta)
'response.write("final: " & fechagenerica(prueba_tonta))
'response.flush
'response.write(sql)
'		conn.Execute "Set language Spanish"
'		Session.lcid = 1034
		objConn1.Execute sql
	end sub

	' devuelve el nombre del campo
	public function getName(indice)
		getName = arrCampos(indice)
	end function

	' obtiene el valor de la propiedad
	public function getProperty(campo)
		getProperty = arrValores(indice(campo))
	end function

	' devuelve el n�mero de campos que posee la clase
	public function numeroCampos()
		numeroCampos = ubound(arrCampos)
	end function

	' inicializa campo con el valor indicado
	public sub setProperty(campo,valor)
		arrValores(indice(campo)) = trim(valor)
	end sub

	' devuelve la posici�n en la que se encuentra el campo indicado
	public function indice(campo)
		dim i
		for i=0 to ubound(arrCampos)
			if arrCampos(i)=campo then
				indice = i
			end if
		next
	end function

	' Aqu�, en Class_Terminate() manejamos cualquier detalle de limpieza de la clase.
	private sub Class_Terminate()
	end sub
end class
%>