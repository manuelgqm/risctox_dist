<%@ LANGUAGE="VBSCRIPT" LCID="1034" CODEPAGE="65001"%>	
<!-- #include file="Lib/ASPUnit.asp" -->
<!-- #include file="../lib/class/SubstanceClass.asp" -->
<!--#include file="../lib/dn_funciones_texto_utf-8.asp"-->
<!--#include file="../lib/dn_funciones_comunes_utf-8.asp"-->
<!--#include file="../config/dbConnection.asp"-->
<!--#include file="../lib/db/substanceLocalRepository.asp"-->

<%
	Response.ContentType = "text/html"
	Response.AddHeader "Content-Type", "text/html;charset=UTF-8"
	Response.CodePage = 65001
	Response.CharSet = "UTF-8"
	
	' Register unit test modules (groups of tests)
	Call ASPUnit.AddModule( _
		ASPUnit.CreateModule( _
			"Single Substance Tests", _
			Array _ 
				( ASPUnit.CreateTest("Name") _
				, ASPUnit.CreateTest("local") _
			), ASPUnit.CreateLifeCycle("Setup", "Teardown") _
		) _
	)

	Call ASPUnit.Run()

	' Lifecycle methods

	Sub Setup()
		Call ExecuteGlobal("Dim mySubstance")
		Set mySubstance = New SubstanceClass
	End Sub

	Sub Teardown()
		Set mySubstance = Nothing
	End Sub

	' Test methods

	Function Name()
		mySubstance.find 957597, objConnection2
		dim substance : set substance = mySubstance.fields
		Call ASPUnit.Equal(substance.item("nombre"), "formaldehído", "Name loaded '" & substance.item("nombre") & "' should match mocked formaldehído")
	End Function

	function local()
		dim localDataFile : localDataFile = Server.MapPath("/istas/risctox/lib/db/local/substancesLevelOneFields.JSON")
		dim formol : set formol = findLocalJSONSubstance(localDataFile, 957597)
		call ASPUnit.Equal(formol("nombre"), "formaldehído", "Name loaded '" & formol("nombre") & "' should match mocked formaldehído")
	end function

%>