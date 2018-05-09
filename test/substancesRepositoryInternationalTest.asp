<%@ LANGUAGE="VBSCRIPT" LCID="1034" CODEPAGE="65001"%>
<!-- #include file="Lib/ASPUnit.asp" -->
<!--#include file="../lib/dn_funciones_texto_utf-8.asp"-->
<!--#include file="../config/dbConnection.asp"-->
<!--#include file="../lib/db/substancesRepositoryInternational.asp"-->

<%
	Response.ContentType = "text/html"
	Response.AddHeader "Content-Type", "text/html;charset=UTF-8"
	Response.CodePage = 65001
	Response.CharSet = "UTF-8"

	' Register unit test modules (groups of tests)
	Call ASPUnit.AddModule( _
		ASPUnit.CreateModule( _
			"substancesRepositoryInternational Funcional Tests", _
			Array _
				( _
					ASPUnit.CreateTest("obtainNombreEnglishTest"), _
					ASPUnit.CreateTest("obtainNombreSpanishTest") _
			), ASPUnit.CreateLifeCycle("Setup", "Teardown") _
		) _
	)

	Call ASPUnit.Run()

	Sub Setup()
		' Call ExecuteGlobal("Dim mySubstance")
		' Set mySubstance = New SubstanceClass
	End Sub

	Sub Teardown()
		' Set mySubstance = Nothing
	End Sub

	function obtainNombreEnglishTest()
		dim nombre_raw : nombre_raw = "hydrogen cyanide@hydrocyanic acid"
		dim language : language = "en"
		dim expected_nombre : expected_nombre = "hydrogen cyanide"

		dim actual_nombre : actual_nombre = obtainNombre(nombre_raw, language)

		call ASPUnit.Equal(actual_nombre, expected_nombre, "Value obtained '" & actual_nombre & "' should match expected value '" & expected_nombre & "'")

	end function

	function obtainNombreSpanishTest()
		dim nombre_raw : nombre_raw = "ácido cianhídrico"
		dim language : language = ""
		dim expected_nombre : expected_nombre = "ácido cianhídrico"

		dim actual_nombre : actual_nombre = obtainNombre(nombre_raw, language)

		call ASPUnit.Equal(actual_nombre, expected_nombre, "Value obtained '" & actual_nombre & "' should match expected value '" & expected_nombre & "'")

	end function
%>
