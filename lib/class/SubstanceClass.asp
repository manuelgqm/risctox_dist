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
		Me.Fields.item("listaNegraClassifications") = getListaNegraClassifications()
		find = id_sustancia
	end function

	Public function obtainLevelOneFields(id_sustancia, connection)
		dim fields
		set fields = findSubstanceLevelOne(id_sustancia, connection)
		Me.fields = fields
		Me.Fields.item("listaNegraClassifications") = getListaNegraClassifications()
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
	Private function getListaNegraClassifications()
		dim result : result = Array()

		if (Me.inList("cancer_rd") or Me.inList("cancer_danesa") or Me.inList("cancer_iarc_excepto_grupo_3") or Me.inList("cancer_otras_excepto_grupo_4") or Me.inList("cancer_mama")) then
			arrayPush result, "cancerígena"
		end if
		if (Me.inList("cop")) then
			arrayPush result, "cop"
		end if
		if (Me.inList("mutageno_rd") or Me.inList("mutageno_danesa")) then
			arrayPush result, "mutágena"
		end if
		if (Me.inList("de")) then
			arrayPush result, "disruptora endocrina"
		end if
		if Me.inNeurotoxicosLists() then 'Businnes Concern: original condition contains and not MySubstance.containsFraseR("R67"), removed due to bad logic
			arrayPush result, "neurotóxica"
		end if
		if (Me.inList("sensibilizante") or Me.inList("sensibilizante_danesa") or Me.inList("sensibilizante_reach")) then
			arrayPush result, "sensibilizante"
		end if
		if (Me.inList("tpr") or Me.inList("tpr_danesa")) then
			arrayPush result, "tóxica para la reproducción"
		end if
		if Me.containsFraseR("R33")then
			arrayPush result, "bioacumulativa"
		end if
		if Me.containsFraseR("R58") then
			arrayPush result, "puede provocar a largo plazo efectos negativos en el medio ambiente"
		end if
		if (Me.inList("tpb")) then
			arrayPush result, "tóxica, persistente y bioacumulativa"
		end if
		if Me.inMpmbList() then
			arrayPush result, "muy persistente y muy bioacumulativa"
		end if
		if Me.containsFraseR("R53") or Me.containsFraseR("R50-53") or Me.containsFraseR("R51-53") or Me.containsFraseR("R52-53") then
			arrayPush result, "puede provocar a largo plazo efectos negativos en el medio ambiente acuático"
		end if

		getListaNegraClassifications = result
	end function

End Class
%>