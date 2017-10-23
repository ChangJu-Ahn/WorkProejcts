<%
'**********************************************************************************************
'*  1. Module명          : 자금관리 
'*  2. Function명        : 예적금관리 
'*  3. Program ID        : f3101mb2
'*  4. Program 이름      : 예적금정보 등록(환율처리)
'*  5. Program 설명      : 예적금정보 등록(환율처리)
'*  6. Comproxy 리스트   : 
'*  7. 최초 작성년월일   : 2000/09/21
'*  8. 최종 수정년월일   : 2000/09/21
'*  9. 최초 작성자       : 김종환 
'* 10. 최종 작성자       : 
'* 11. 전체 comment      :
'* 12. 공통 Coding Guide : this mark(☜) means that "Do not change"
'*                         this mark(⊙) Means that "may  change"
'*                         this mark(☆) Means that "must change"
'* 13. History           :
'*                         -2000/09/21 : ..........
'**********************************************************************************************




'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->


<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next														'☜: 

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Dim B17014

Dim objXchRate
Dim objCnclXchRate

Call LoadBasisGlobalInf()

Call HideStatusWnd


strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

Select Case strMode
	Case "XchRate"   
	  '********************************************************  
	'              LOOKUP For Exchange Rate
	'********************************************************  

	    Err.Clear                                                               '☜: Clear error no
	
		Set objXchRate = Server.CreateObject("B17014.B17014LookupExchangeRate")
     
	    '-----------------------
		'Com action result check area(OS,internal)
	    '-----------------------
		If Err.Number <> 0 Then
			Set objXchRate = Nothing																'☜: ComProxy UnLoad
			Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)						'⊙:
			Call HideStatusWnd
			Response.End																		'☜: Process End
		End If
    
	    '-----------------------
		'Data manipulate  area(import view match)
	    '-----------------------

		objXchRate.ImportBCurrencyCurrency   = Request("txtLocCur")
		objXchRate.ImportToBCurrencyCurrency = Request("txtDocCur")
		objXchRate.ImportBDailyExchangeRateApprlDt = Request("txtAppDt")
        
	    objXchRate.ServerLocation = ggServerIP
		objXchRate.CommandSent    = "LOOKUP"
    
	    '-----------------------
		'Com action area
		'-----------------------       
		objXchRate.ComCfg = gConnectionString
	    objXchRate.Execute 
		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)                          '⊙:
			Set objXchRate = Nothing																	    '☜: ComProxy UnLoad
			Call HideStatusWnd
			Response.End																				'☜: Process End
		End If
    
		'-----------------------
		'Com action result check area(DB,internal)
		'-----------------------
		'If Not (objXchRate.OperationStatusMessage = MSG_OK_STR) Then
		'	Call DisplayMsgBox(objXchRate.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		'	Set objXchRate = Nothing												'☜: ComProxy Unload
		'	Call HideStatusWnd
		'	Response.End														'☜: 비지니스 로직 처리를 종료함 
		'End If    
		If Not (objXchRate.OperationStatusMessage = MSG_OK_STR) Then
		    Select Case objXchRate.OperationStatusMessage
		        Case MSG_DEADLOCK_STR
		            Call DisplayMsgBox2("999999", "25", "deadlock or timeout" , I_MKSCRIPT)
		        Case MSG_DBERROR_STR
		            Call DisplayMsgBox2(objXchRate.ExportErrEabSqlCodeSqlcode, _
										objXchRate.ExportErrEabSqlCodeSeverity, _
										objXchRate.ExportErrEabSqlCodeErrorMsg, I_MKSCRIPT)
		        Case Else
		            Call DisplayMsgBox(objXchRate.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		    End Select
			Call HideStatusWnd
			Response.End																				'☜: Process End
		End If

		'-----------------------
		'Result data display area
		'----------------------- 
%>
<Script Language=vbscript>
		With parent.frm1
		    .txtXchRate.Value = "<%=objXchRate.ExportBDailyExchangeRateStdRate%>"                          '환율 
		End With
</Script>
<%	 
		Set objXchRate = Nothing															    '☜: Unload Comproxy
		Call HideStatusWnd
		Response.End																		'☜: Process End   

	Case "CnclXchRate"   
	  '********************************************************  
	'              LOOKUP For Exchange Rate
	'********************************************************  

	    Err.Clear                                                               '☜: Clear error no
	
		Set objCnclXchRate = Server.CreateObject("B17014.B17014LookupExchangeRate")
     
	    '-----------------------
		'Com action result check area(OS,internal)
	    '-----------------------
		If Err.Number <> 0 Then
			Set objCnclXchRate = Nothing																'☜: ComProxy UnLoad
			Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)						'⊙:
			Call HideStatusWnd
			Response.End																		'☜: Process End
		End If
    
	    '-----------------------
		'Data manipulate  area(import view match)
	    '-----------------------

		objCnclXchRate.ImportBCurrencyCurrency   = Request("txtLocCur")
		objCnclXchRate.ImportToBCurrencyCurrency = Request("txtDocCur")
		objCnclXchRate.ImportBDailyExchangeRateApprlDt = Request("txtAppDt")
        
	    objCnclXchRate.ServerLocation = ggServerIP
		objCnclXchRate.CommandSent    = "LOOKUP"
    
	    '-----------------------
		'Com action area
		'-----------------------       
		objCnclXchRate.ComCfg = gConnectionString
	    objCnclXchRate.Execute 
		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)                          '⊙:
			Set objCnclXchRate = Nothing																	    '☜: ComProxy UnLoad
			Call HideStatusWnd
			Response.End																				'☜: Process End
		End If
    
		'-----------------------
		'Com action result check area(DB,internal)
		'-----------------------
		'If Not (objCnclXchRate.OperationStatusMessage = MSG_OK_STR) Then
		'	Call DisplayMsgBox(objCnclXchRate.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		'	Set objCnclXchRate = Nothing												'☜: ComProxy Unload
		'	Call HideStatusWnd
		'	Response.End														'☜: 비지니스 로직 처리를 종료함 
		'End If    
		If Not (objCnclXchRate.OperationStatusMessage = MSG_OK_STR) Then
		    Select Case objCnclXchRate.OperationStatusMessage
		        Case MSG_DEADLOCK_STR
		            Call DisplayMsgBox2("999999", "25", "deadlock or timeout" , I_MKSCRIPT)
		        Case MSG_DBERROR_STR
		            Call DisplayMsgBox2(objCnclXchRate.ExportErrEabSqlCodeSqlcode, _
										objCnclXchRate.ExportErrEabSqlCodeSeverity, _
										objCnclXchRate.ExportErrEabSqlCodeErrorMsg, I_MKSCRIPT)
		        Case Else
		            Call DisplayMsgBox(objCnclXchRate.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		    End Select
			Call HideStatusWnd
			Response.End																				'☜: Process End
		End If

		'-----------------------
		'Result data display area
		'----------------------- 
%>
<Script Language=vbscript>
		With parent.frm1
		    .txtCnclXchRate.Value = "<%=objCnclXchRate.ExportBDailyExchangeRateStdRate%>"                          '환율 
		End With
</Script>
<%	 
		Set objCnclXchRate = Nothing															    '☜: Unload Comproxy
		Call HideStatusWnd
		Response.End																		'☜: Process End   

End Select

%>
