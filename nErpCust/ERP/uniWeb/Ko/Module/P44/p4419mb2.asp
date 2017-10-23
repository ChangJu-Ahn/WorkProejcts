<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		:
'*  3. Program ID			: p4419mb2.asp
'*  4. Program Name			: Cancel Confirmation By Order
'*  5. Program Desc			: Cancel Confirmation By Order
'*  6. Comproxy List		: +PP4G418.cPCnclCnfmRsltByOrdSvr
'*  7. Modified date(First)	: 2000/03/30
'*  8. Modified date(Last) 	: 2002/08/21
'*  9. Modifier (First)		: Park, Bum-Soo
'* 10. Modifier (Last)		: Chen, Jae Hyun
'* 11. Comment				:
'**********************************************************************************************
'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다.
%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("I", "*", "NOCOOKIE","MB")

Call HideStatusWnd

On Error Resume Next

Dim pPP4G418

Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount

Dim iCUCount
Dim iDCount

Dim ii


    Err.Clear																		'☜: Protect system from crashing

	
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
	Set pPP4G418 = Server.CreateObject("PP4G418.cPCnclCnfmRsltByOrdSvr")
		    
	If CheckSYSTEMError(Err,True) = True Then
		Set pPP4G418 = Nothing		
		Response.Write "<Script Language=VBScript>" & vbCrLF
		Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
		Response.Write "</Script>" & vbCrLF
		Response.End
	End If
	
	Call pPP4G418.P_CNCL_CNFM_RSLT_BY_ORD_SVR(gStrGlobalCollection,_
									 itxtSpread)

	If CheckSYSTEMError(Err, True) = True Then
		Set pPP4G418 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLF
		Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
		Response.Write "</Script>" & vbCrLF															'☜: Unload Component
		Response.End
	End If
	
	Set pPP4G418 = Nothing 

	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "parent.DbSaveOk" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
%>

