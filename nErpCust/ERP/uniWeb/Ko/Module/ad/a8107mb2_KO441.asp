<%
'**********************************************************************************************
'*  1. Module명          : 회계 
'*  2. Function명        : A_Confirm_TempGl,A_Unconfirm_TempGl
'*  3. Program ID        : a5103mb2
'*  4. Program 이름      : 본지점 결의전표 일괄승인,승인취소 
'*  5. Program 설명      : 본지점 결의전표 일괄승인,승인취소를 위한 Logic
'*  6. Comproxy 리스트   : 
'*  7. 최초 작성년월일   : 2001/02/07
'*  8. 최종 수정년월일   : 
'*  9. 최초 작성자       : hersheys
'*  9. 최종 작성자       : 
'* 10. 전체 comment      :
'* 11. 공통 Coding Guide : 주석에 mark(☜)로 되어있는 부분은 업무 담당자가 변경(X)
'*                         주석에 mark(⊙)로 되어있는 부분은 업무 담당자가 변경(O)
'* 12. History           : 
'**********************************************************************************************
Response.Expires = -1		'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True		'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>

<!-- #Include file="../../inc/IncServer.asp"  -->

<%					

'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

'On Error Resume Next			' ☜: 

Dim a53013					' 결의전표승인 ComProxy Dll 사용 변수 
Dim strMode						'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strWkfg

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
		
strMode = Request("txtMode")	'☜ : 현재 상태를 받음 
strWkfg = Request("htxtWorkFg")

Select Case strMode
	'-------------------------------------------------------------------------------
	'								    일괄처리 로직 
	'-------------------------------------------------------------------------------   
	Case CStr(UID_M0002)																'☜: 저장 요청을 받음 
	    					
	    Err.Clear 												'☜: Protect system from crashing

		Set a53013 = Server.CreateObject("A53013.A53012HqConfirmTempGlSvr")

		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			Set a53013 = Nothing												'☜: ComProxy Unload
			Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)
			Response.End														'☜: 비지니스 로직 처리를 종료함 
		End If	

		'-----------------------
		'Data manipulate area
		'-----------------------
		a53013.ImportFromDtATempGlTempGlDt      = UNIConvDate(Request("txtFromTempGlDt"))
		a53013.ImportToDtATempGlTempGlDt        = UNIConvDate(Request("txtToTempGlDt"))
		A53013.ImportBBizAreaBizAreaCd			= Trim(Request("txtBizAreaCd"))
		a53013.ImportATempGlUpdtUserId          = gUsrId
		a53013.ImportBAcctDeptOrgChangeId       = gChangeOrgId
		a53013.ServerLocation                   = ggServerIP

		If UCase(strWkfg) = "CONF"  then		
			a53013.ImportIefSuppliedCommand = "CONF"			
		Elseif UCase(strWkfg) = "UNCONF"  then		
			a53013.ImportIefSuppliedCommand = "UNCONF"
		End if					

		a53013.ComCfg = gConnectionString
'		a53013.ComCfg = "TCP letitbe 2055"
        a53013.Execute

        '-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.number <> 0 Then
			Set a53013 = Nothing
			Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)						'⊙:
			Response.End 
		End If

		'-----------------------
		'Com action result check area(DB,internal)
		'-----------------------				
		If Not (a53013.OperationStatusMessage = MSG_OK_STR) Then
			Call DisplayMsgBox(a53013.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
			Set a53013 = Nothing
			Response.End 
		End If
                   
		a53013.Clear
		lGrpCnt = 0
		
		'ggoSpread.SSDeleteFlag lStartRow, lEndRow	'뭐하는건지 모르겠음       
	
	Set a53013 = Nothing                                                   '☜: Unload Comproxy
    
%>
<Script Language=vbscript>
	With parent																		'☜: 화면 처리 ASP 를 지칭함 
        .InitSpreadSheet
        .InitComboBox
        .InitVariables 															'⊙: Initializes local global variables	
		.dbQuery		
	End With
 
</Script>

<%					

End Select

%>

