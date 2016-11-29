<!--#include file="../arrayManipulations.asp"-->
<!--#include file="../stringManipulations.asp"-->
<!--#include file="../dictionaryManipulations.asp"-->
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

	Public function obtainSaludFields(id_sustancia, connection)
		dim fields : set fields = findSaludFields(id_sustancia, connection)
		Me.fields = fields
	end function

	Public function inList(listName)
		inList = inArray(listName, Me.fields.Item("featuredLists"))
	end function

	public function inLists(lists)
		inLists = anyElementInArray(lists, Me.fields.Item("featuredLists"))
	end function

	public function inMpmbList()
		inMpmbList = Me.Fields.Item("mpmb")
	end function

	public function inNeurotoxicosLists()
		inNeurotoxicosLists = inLists( array _
			( "neurotoxico" _
			, "neurotoxico_rd" _
			, "neurotoxico_danesa" _
			, "neurotoxico_nivel" _
			) _
		)
	end function

	public function containsFraseR(frase)
		containsFraseR = stringContains(Me.fields.Item("frasesR"), frase)
	end function

	public function hasFrasesRdanesa()
		hasFrasesRdanesa = false
		if mFields.Item("frasesR") = "" and mFields.Item("frases_r_danesa") <> "" then
			hasFrasesRdanesa = true
		end if
	end function

	public function hasListaNegraClassifications()
		hasListaNegraClassifications = false
		if ubound(mFields.item("listaNegraClassifications")) > -1 then
			hasListaNegraClassifications = true
		end if
	end function

'PRIVATE'

End Class
%>