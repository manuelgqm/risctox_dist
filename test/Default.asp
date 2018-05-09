<% Option Explicit %>

<!-- Include ASPUnit library -->
<!-- #include file="Lib/ASPUnit.asp" -->

<%
	' Register pages to test
	Call ASPUnit.AddPage("substancesRepositoryInternationalTest.asp")

	' Execute tests
	Call ASPUnit.Run()
%>
