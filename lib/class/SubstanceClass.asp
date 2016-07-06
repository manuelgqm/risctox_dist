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

	Public function inList(listName)
		dim result
		result = false

		if listName = "" Then
			inList = result
		end if
		result = in_array(listName, Me.fields.Item("featuredLists"))

		inList = result
	end function

	Private function in_array(element, arr)
	  in_array = False
	  For i=0 To Ubound(arr)
	     If Trim(arr(i)) = Trim(element) Then
	        in_array = True
	        Exit Function
	     End If
	  Next
End Function

End Class
%>