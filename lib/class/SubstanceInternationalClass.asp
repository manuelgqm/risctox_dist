<!--#include file="../arrayManipulations.asp"-->
<!--#include file="../stringManipulations.asp"-->
<!--#include file="../dictionaryManipulations.asp"-->
<!--#include file="../db/substancesRepositoryInternational.asp"-->

<%
Class SubstanceClassInternational
	Private mFields

	Public property Get Fields()
		set Fields = mFields
	End property
	Public property Let Fields(pData)
		set mFields = pData
	End property

' PUBLIC
	Public function obtainIdentification(id_sustancia, lang, connection)
		dim fields
		set fields = findIdentification(id_sustancia, lang, connection)
		Me.fields = fields
	End function

End Class
%>
