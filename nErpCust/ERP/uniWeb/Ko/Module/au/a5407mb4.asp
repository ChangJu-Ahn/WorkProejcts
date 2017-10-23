
<%
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : 
'*  3. Program ID           : A5406mb1
'*  4. Program Name         : 미결반제(만기어음)
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'                             
'*  7. Modified date(First) : 2002/11/05
'*  8. Modified date(Last)  : 2002/11/05
'*  9. Modifier (First)     : KIM HO YOUNG
'* 10. Modifier (Last)      : KIM HO YOUNG
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2002/11/05 : ..........
'**********************************************************************************************


Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True								'☜ : ASP가 버퍼에 저장되어 마지막에 바로 Client에 내려간다.


'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 

%>
<%
'#########################################################################################################
'												1. Include
'##########################################################################################################
%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
'#########################################################################################################
'												2. 조건부 
'##########################################################################################################

													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call HideStatusWnd	
On Error Resume Next														'☜: 
Call LoadBasisGlobalInf() 
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

'#########################################################################################################
'												2.1 조건 체크 
'##########################################################################################################

If strMode = "" Then
'	Response.End 
End If

'#########################################################################################################
'												2. 업무 처리 수행부 
'##########################################################################################################

'#########################################################################################################
'												2.1. 변수, 상수 선언 
'##########################################################################################################
' 수정을 요함 
Dim pDelGlCardAcct																'☆ : 조회용 ComProxy Dll 사용 변수 

Dim IntRows
Dim IntCols
Dim sList
Dim vbIntRet
Dim intCount
Dim IntCount1
Dim LngMaxRow
Dim LngMaxRow1
Dim StrNextKey
Dim lgStrPrevKey
Dim lgIntFlgMode
dim test

' Com+ Conv. 변수 선언 
Dim pvStrGlobalCollection 
Dim I1_cls_no
Dim arrCount
					'☜: 현재 조회/Prev/Next 요청을 받음 
	'#########################################################################################################
	'												2.2. 요청 변수 처리 
	'##########################################################################################################
	lgStrPrevKey = Request("lgStrPrevKey")

	'#########################################################################################################
	'												2.3. 업무 처리 
	'##########################################################################################################

	Set pDelGlCardAcct = Server.CreateObject("PAUG035.cADelGlCardAcctSvr")
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		Set pDelGlCardAcct = Nothing												'☜: ComProxy Unload
		Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)	'⊙:
		Response.End														'☜: 비지니스 로직 처리를 종료함 
	End If

		LngMaxRow  = CLng(Request("txtMaxRows"))												'☜: Fetechd Count      
		LngMaxRow1  = CLng(Request("txtMaxRows1"))

		I1_cls_no = Request("txtClsNo")

		On Error Resume next
		Call pDelGlCardAcct.A_DELETE_GL_CARD_ACCT_SVR(gStrGlobalCollection,Trim(I1_cls_no))
						
	'-----------------------
	'Com Action Area
	'-----------------------

		If CheckSYSTEMError(Err,True) = True Then
		
			Set pDelGlCardAcct = Nothing																	'☜: ComProxy Unload
			Response.End																			'☜: 비지니스 로직 처리를 종료함 
		End If

		Set pDelGlCardAcct = Nothing

    Response.Write "<Script Language=VBScript> " & vbCr         
    Response.Write "With parent "				 & vbCr	
	Response.Write " .DbDeleteOK() "								  & vbCr
    Response.Write "End With "				 & vbCr	  
    Response.Write "</Script>"       																	  & vbCr	
		
	%>		

