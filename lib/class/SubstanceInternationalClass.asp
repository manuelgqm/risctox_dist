<!--#include file="../arrayManipulations.asp"-->
<!--#include file="../stringManipulations.asp"-->
<!--#include file="../dictionaryManipulations.asp"-->
<!--#include file="../db/substancesRepositoryInternational.asp"-->

<%
Class SubstanceClassInternational
	Private m_identification
	Public property Get identification()
		set identification = m_identification
	End property

	Private m_classification
	Public property Get classification()
		set classification = m_classification
	End property

	Private m_health_effects
	Public property Get health_effects()
		set health_effects = m_health_effects
	End property

	Private m_environment_effects
	Public property Get environment_effects()
		set environment_effects = m_environment_effects
	End property

	Public default function init(substance_id, lang, connection)
		set m_identification = findIdentification(substance_id, lang, connection)
		set m_classification = findClassification(substance_id, lang, connection)
		set m_health_effects = find_health_effects(substance_id, lang, connection)
		set m_environment_effects = find_environment_effects(substance_id, lang, connection)

		set init = Me
	End function

	public function inLists(lists)
		inLists = anyElementInArray(lists, m_identification.item("featuredLists"))
	end function

	Public function inList(listName)
		inList = inArray(listName, m_identification.Item("featuredLists"))
	end function

	Public Function has_health_effects()
		has_health_effects = False
		if _
			m_health_effects.item("cardiocirculatorio") OR _
			m_health_effects.item("rinyon") OR _
			m_health_effects.item("respiratorio") OR _
			m_health_effects.item("reproductivo") OR _
			m_health_effects.item("piel_sentidos") OR _
			m_health_effects.item("neuro_toxicos") OR _
			m_health_effects.item("musculo_esqueletico") OR _
			m_health_effects.item("sistema_inmunitario") OR _
			m_health_effects.item("higado_gastrointestinal") OR _
			m_health_effects.item("sistema_endocrino") OR _
			m_health_effects.item("embrion") OR _
			m_health_effects.item("cancer") _
		then _
			has_health_effects = True
	End Function

End Class
%>
