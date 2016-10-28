<%@ LANGUAGE="VBSCRIPT" LCID="1034" CODEPAGE="65001"%>	
<!--#include file="recordsetManipulations.asp"-->
<!--#include file="../arrayManipulations.asp"-->
<!--#include file="substancesRepository.asp"-->
<!--#include file="../../config/dbConnection.asp"-->
<%
Response.ContentType = "text/html"
Response.AddHeader "Content-Type", "text/html;charset=UTF-8"
Response.CodePage = 65001
Response.CharSet = "UTF-8"

dim query : query = composeSubstanceLevelOneFieldsQuery(957597)
dim url : url = "local/substancesLevelOneFields.JSON"
dim substancesIds : substancesIds = Array("956228", "953905", "957555", "957597", "954578", "955906", "956773", "956772")
dim substancesIdsSrz : substancesIdsSrz = Join(substancesIds, "', '")
query = query & " or sus.id in ('" & substancesIdsSrz & "')"
call recordsetToFile(query, objConnection2, url)
%>