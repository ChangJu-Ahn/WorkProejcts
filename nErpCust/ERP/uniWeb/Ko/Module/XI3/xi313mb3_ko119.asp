<%
'**********************************************************************************************
'*  1. Module Name			: Interface 
'*  2. Function Name		: 
'*  3. Program ID			: xi313mb3_ko119.asp
'*  4. Program Name			: 최종실적등록
'*  5. Program Desc			: 
'*  6. Comproxy List		: +PXI3G132_KO119
'*  7. Modified date(First)	: 2006-04-24
'*  8. Modified date(Last) 	:
'*  9. Modifier (First)		:HJO
'* 10. Modifier (Last)		: 
'* 11. Comment		:
'**********************************************************************************************
'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
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

    Session.Timeout = 60          ' minute 
    Server.ScriptTimeOut = 3600   ' NumSeconds

Dim strPlantCd											'☆ : Lookup 용 코드 저장 변수 
Dim iErrorPosition										'☆ : Error Position									
Dim iErrorProdtOrdNo, iErrorOprNo, iErrorGoodMvmt		'☆ : Error Return Value
Dim msgStr1, msgStr2

Dim oPXI3122

Dim iCUCount

Dim ii											'☆ : Lookup 용 코드 저장 변수 

	Err.Clear											'☜: Protect system from crashing
	
	Set oPXI3122 = Server.CreateObject("PXI3G132_KO119.cUploadErpProdRcpt")
	If CheckSYSTEMError(Err,True) = True Then
		Set oPXI3122 = Nothing
		Response.End
	End If		

	Call oPXI3122.UPLOAD_ERP_PROD_RCPT_MAIN(gStrGlobalCollection,"")

	If CheckSYSTEMError(Err,True) = True Then
		Set oPXI3122 = Nothing
		Response.End
	
	End If
	Set oPXI3122 = Nothing
		
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End	
	%>
