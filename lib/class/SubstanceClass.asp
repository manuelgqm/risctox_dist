<!--#include file="../db/substancesRepository.asp"-->
<!--#include file="../db/synonymsRepository.asp"-->
<!--#include file="../db/substanceListsRepository.asp"-->
<%
Class SubstanceClass
	Private mNombre
	Private mFields
	Private mFieldsShown

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

	Public property Get fieldsShown()
		fieldsShown = mFieldsShown
	End property
	Public property Let fieldsShown(pData)
		mFieldsShown = pData
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
			exit function
		end if
		result = in_array(listName, Me.fields.Item("featuredLists"))

		inList = result
	end function

	public function presentInLists(lists)
		dim result
		result = false

		if not isArray(lists) then
			presentInLists = result
			exit function
		end if

		dim i, list
		for i = 0 to Ubound(lists)
			list = lists(i)
			result = Me.inList(list)
			if result then 
				presentInLists = true
				exit function
			end if
		next

		presentInLists = result
	end function

	public function inMpmbList()
		inMpmbList = Me.Fields.Item("mpmb")
	end function

	public function addShown(fieldName)
		Me.fieldsShown = arrayPush(Me.fieldsShown, fieldName)
	end function

	public function showed(fieldName)
		isShown = in_array(fieldName, Me.fieldsShown)
	end function

	Private function in_array(element, arrayParameter)
		in_array = false

		if not isArray(arrayParameter) then
			in_array = false
			exit function
		end if
		For i = 0 To Ubound(arrayParameter)
			If Trim(arrayParameter(i)) = Trim(element) Then 
				in_array = true
				Exit Function
			end if
		Next
	End Function

	private function arrayPush(arrayParameter, valueParameter) 
		dim uba
		dim result()

		uba = getUBound(arrayParameter) 
		redim preserve result(uba + 1) 
		result(uba + 1) = valueParameter

		arrayPush = result
	end function

	function getUbound(arrayParameter) 
		dim result
		result = -1 
		
		on error resume next
		if not vartype(arrayParameter) = 8204 then 
			getUbound = result
			exit function
		end if
		result = ubound(arrayParameter) 

		getUbound = result
	end function 

End Class
%>