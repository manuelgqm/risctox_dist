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

' PUBLIC METHODS'
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

	public function inNeurotoxicosLists()
		dim NEUROTOXICO_LISTS : NEUROTOXICO_LISTS = array("neurotoxico", "neurotoxico_rd", "neurotoxico_danesa", "neurotoxico_nivel")
		inNeurotoxicosLists = presentInLists(NEUROTOXICO_LISTS)
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

	public function addShown(fieldName)
		arrayPush Me.fieldsShown, fieldName
	end function

	public function showed(fieldName)
		isShown = in_array(fieldName, Me.fieldsShown)
	end function

'PRIVATE METHODS'
	public function getListaNegraClassifications()
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
		if (Me.inNeurotoxicosLists() and not MySubstance.containsFraseR("R67")) then
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

	Private Sub arrayPush(byRef arrayParameter, valueParameter) 
		redim preserve arrayParameter(uBound(arrayParameter) + 1)
		arrayParameter(uBound(arrayParameter)) = valueParameter
	End Sub

End Class
%>