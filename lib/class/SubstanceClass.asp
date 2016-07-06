<!--#include file="../db/substancesRepository.asp"-->
<!--#include file="../db/synonymsRepository.asp"-->
<!--#include file="../db/substanceListsRepository.asp"-->
<%
Class SubstanceClass
	Private mNombre
	Private mFields

	Public property Get nombre()
		nombre = mNombre
	End property
	Public property Let nombre(pData)
		mNombre = pData
	End property

	Public property Get Fields()
		set Fields = mFields
	End property
	Public property Let Fields(pData)
		set mFields = pData
	End property

	Public function find(id_sustancia, connection)
		dim fields
		set fields = findSubstance(id_sustancia, connection)
		Me.Fields = fields
		find = id_sustancia
	end function

End Class
%>