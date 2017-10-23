<%@ LANGUAGE=VBScript %>
<% Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"-->
<%
	Dim lgOpModeCRUD
	Dim objBDC005
	Dim iErrorPosition
	Dim iStrSpread

	Call HideStatusWnd

	Call LoadBasisGlobalInf()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
	lgOpModeCRUD = Request("txtMode")

	On Error Resume Next
	Err.Clear
	Set objBDC005 = Server.CreateObject("BDC005.clsBatchWorker")
'	Set objBDC005 = GetObject("queue:/new:BDC005.clsBatchWorker")
	
	If CheckSYSTEMError(Err,True) = True Then
		On Error Goto 0
		Response.End
	End If

	iStrSpread = FilterVar(Request("txtSpread"),"","SNM")

	Call objBDC005.ExecuteJobs(gStrGlobalCollection, istrSpread)
	
	If CheckSYSTEMError(Err,True) = True Then 
		If Not (objBDC005 Is Nothing) Then  Set objBDC005 = Nothing       
		Response.End
	End If
	
	If Not (objBDC005 Is Nothing) Then  Set objBDC005 = Nothing
	
	On Error Goto 0
	
	Response.Write "<Script Language=vbscript>"  & vbCr
	Response.Write "Parent.OpenRunJobOk "            & vbCr
	Response.Write "</Script>"
%>
