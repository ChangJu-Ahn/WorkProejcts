<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4417mb2.asp
'*  4. Program Name			: Cancel Confirm By Operation
'*  5. Program Desc			: Confirm Production Results (Called By p3221ma3.asp, p3221ma6.asp)
'*  6. Comproxy List		: +P32214CnclCnfmRsltByOpr
'*  7. Modified date(First)	: 2000/03/30
'*  8. Modified date(Last) 	: 2002/08/21
'*  9. Modifier (First)		: Park, Bum-Soo
'* 10. Modifier (Last)		: Chen, Jae Hyun
'* 11. Comment		:
'**********************************************************************************************
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

Dim pPP4G409												'☆ : 입력/수정용 ComProxy Dll 사용 변수 
Dim I1_b_plant_cd 
Dim IG1_import_group_String 

Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount

Dim iCUCount
Dim iDCount

Dim ii

On Error Resume Next

    Err.Clear											'☜: Protect system from crashing
    
	Call HideStatusWnd

    I1_b_plant_cd = Request("txtPlantCd")
	
	itxtSpread = ""
    
	iDCount  = Request.Form("txtDSpread").Count
	             
	itxtSpreadArrCount = -1
	             
	ReDim itxtSpreadArr(iDCount)
	             
	For ii = 1 To iDCount
	    itxtSpreadArrCount = itxtSpreadArrCount + 1
	    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtDSpread")(ii)
	Next

	itxtSpread = Join(itxtSpreadArr,"")
	
	
	Set pPP4G409 = Server.CreateObject("PP4G409.cPCnclCnfmRsltByOp")

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------

	If CheckSYSTEMError(Err, True) = True Then
		Response.Write "<Script Language=VBScript>" & vbCrLF
		Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
		Response.Write "</Script>" & vbCrLF
		Response.End
	End If


	Call pPP4G409.P_CNCL_CNFM_RSLT_BY_OPR_SVR(gStrGlobalCollection, _
								I1_b_plant_cd, _
								itxtSpread)
	If CheckSYSTEMError(Err, True) = True Then
	
		Set pPP4G409 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLF
		Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
		Response.Write "</Script>" & vbCrLF															'☜: Unload Component
		Response.End
	End If

    Set pPP4G409= Nothing	
    
    Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "parent.DbSaveOk" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
%>

