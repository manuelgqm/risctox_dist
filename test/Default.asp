<% Option Explicit %>

<!-- Include ASPUnit library -->
<!-- #include file="Lib/ASPUnit.asp" -->

<%
	' Register pages to test
	Call ASPUnit.AddPage("TestSubstance.asp")

	' Execute tests
	Call ASPUnit.Run()
%>