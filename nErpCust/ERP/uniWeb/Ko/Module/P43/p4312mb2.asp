<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4312mb2.asp
'*  4. Program Name			: Cancel Manual Issue
'*  5. Program Desc			: 
'*  6. Comproxy List		: PP4G302.cPCnclManualIssSvr
'*  7. Modified date(First) : 2000/03/21
'*  8. Modified date(Last)  : 2002/11/22
'*  9. Modifier (First)     : Park, BumSoo
'* 10. Modifier (Last)      : Jeon, Jae Hyun
'* 11. Comment              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 

Call HideStatusWnd

On Error Resume Next

Dim pPP4G302				'PP4G302.cPCnclManualIssSvr
Dim iErrorPosition
Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount

Dim iDCount

Dim ii

    Err.Clear																		'¢Ð: Protect system from crashing

	itxtSpread = ""
	
	iDCount  = Request.Form("txtDSpread").Count
	             
	itxtSpreadArrCount = -1
	
	ReDim itxtSpreadArr(iDCount)
	             
	For ii = 1 To iDCount
	    itxtSpreadArrCount = itxtSpreadArrCount + 1
	    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtDSpread")(ii)
	Next
	
	itxtSpread = Join(itxtSpreadArr,"")
	
	'-----------------------
	'Com Action Area
	'-----------------------
	Set pPP4G302 = Server.CreateObject("PP4G302_ko441.cPCnclManualIssSvr")
		    
	If CheckSYSTEMError(Err,True) = True Then	
		Response.Write "<Script Language=VBScript>" & vbCrLF
		Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
		Response.Write "</Script>" & vbCrLF
		Response.End
	End If

	Call pPP4G302.P_CANCEL_MANUAL_ISSUE_SRV(gStrGlobalCollection, _
											itxtSpread, _
											iErrorPosition)

	If CheckSYSTEMError2(Err, True, iErrorPosition & "Çà", "", "", "", "") = True Then
		Set pPP4G302 = Nothing	
		Response.Write "<Script Language=VBScript>" & vbCrLF
		Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
		Response.Write "Call parent.SheetFocus(" & iErrorPosition & ", 1)" & vbCrLF
		Response.Write "</Script>" & vbCrLF														'¢Ð: Unload Component
		Response.End
	End If
	
	Set pPP4G302 = Nothing 

	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "parent.DbSaveOk" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End

%>
