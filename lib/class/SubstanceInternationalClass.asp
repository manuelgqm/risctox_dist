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

	Public default function init(substance_id, lang, connection_string)
		set m_identification = findIdentification(substance_id, lang, connection_string)
		set m_classification = findClassification(substance_id, lang, connection_string)

		set init = Me
	End function

	public function inLists(lists)
		inLists = anyElementInArray(lists, m_identification.item("featuredLists"))
	end function

	Public function inList(listName)
		inList = inArray(listName, m_identification.Item("featuredLists"))
	end function

End Class
%>
