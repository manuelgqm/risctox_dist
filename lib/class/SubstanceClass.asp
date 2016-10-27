<!--#include file="../db/substancesRepository.asp"-->
<%
Class SubstanceClass
	Private mFields

	Public property Get Fields()
		set Fields = mFields
	End property
	Public property Let Fields(pData)
		set mFields = pData
	End property

' PUBLIC METHODS'
	Public function find(id_sustancia, connection)
		dim fields
		set fields = findSubstance(id_sustancia, connection)
		Me.Fields = fields
	end function

	Public function obtainLevelOneFields(id_sustancia, connection)
		dim fields
		set fields = findSubstanceLevelOne(id_sustancia, connection)
		Me.fields = fields
	End function

	Public function inList(listName)
		inList = inArray(listName, Me.fields.Item("featuredLists"))
	end function

	public function inLists(lists)
		inLists = false
		if not isArray(lists) then
			exit function
		end if
		dim i
		for i = 0 to Ubound(lists)
			if Me.inList(lists(i)) then 
				inLists = true
				exit function
			end if
		next
	end function

	public function inMpmbList()
		inMpmbList = Me.Fields.Item("mpmb")
	end function

	public function inNeurotoxicosLists()
		dim NEUROTOXICO_LISTS : NEUROTOXICO_LISTS = array("neurotoxico", "neurotoxico_rd", "neurotoxico_danesa", "neurotoxico_nivel")
		inNeurotoxicosLists = inLists(NEUROTOXICO_LISTS)
	end function

	public function containsFraseR(frase)
		dim result : result = false
		if frase = "" then
			containsFraseR = result
			exit function
		end if
		if instr(Me.fields.Item("frasesR"), frase) > 0 then
			result = true
		end if

		containsFraseR = result
	end function

	public function hasFrasesRdanesa()
		dim result : result = false
		if mFields.Item("frasesR") = "" and mFields.Item("frases_r_danesa") <> "" then
			result = true
		end if

		hasFrasesRdanesas = result
	end function

	public function hasListaNegraClassifications()
		dim result : result = false

		if ubound(mFields.item("listaNegraClassifications")) > -1 then
			result = true
		end if

		hasListaNegraClassifications = result
	end function

'PRIVATE'

End Class
%>