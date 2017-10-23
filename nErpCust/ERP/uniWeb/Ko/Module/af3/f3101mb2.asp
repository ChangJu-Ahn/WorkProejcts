<%
'**********************************************************************************************
'*  1. Module��          : �ڱݰ��� 
'*  2. Function��        : �����ݰ��� 
'*  3. Program ID        : f3101mb2
'*  4. Program �̸�      : ���������� ���(ȯ��ó��)
'*  5. Program ����      : ���������� ���(ȯ��ó��)
'*  6. Comproxy ����Ʈ   : 
'*  7. ���� �ۼ������   : 2000/09/21
'*  8. ���� ���������   : 2000/09/21
'*  9. ���� �ۼ���       : ����ȯ 
'* 10. ���� �ۼ���       : 
'* 11. ��ü comment      :
'* 12. ���� Coding Guide : this mark(��) means that "Do not change"
'*                         this mark(��) Means that "may  change"
'*                         this mark(��) Means that "must change"
'* 13. History           :
'*                         -2000/09/21 : ..........
'**********************************************************************************************




'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->


<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next														'��: 

Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

Dim B17014

Dim objXchRate
Dim objCnclXchRate

Call LoadBasisGlobalInf()

Call HideStatusWnd


strMode = Request("txtMode")												'�� : ���� ���¸� ���� 

Select Case strMode
	Case "XchRate"   
	  '********************************************************  
	'              LOOKUP For Exchange Rate
	'********************************************************  

	    Err.Clear                                                               '��: Clear error no
	
		Set objXchRate = Server.CreateObject("B17014.B17014LookupExchangeRate")
     
	    '-----------------------
		'Com action result check area(OS,internal)
	    '-----------------------
		If Err.Number <> 0 Then
			Set objXchRate = Nothing																'��: ComProxy UnLoad
			Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)						'��:
			Call HideStatusWnd
			Response.End																		'��: Process End
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
			Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)                          '��:
			Set objXchRate = Nothing																	    '��: ComProxy UnLoad
			Call HideStatusWnd
			Response.End																				'��: Process End
		End If
    
		'-----------------------
		'Com action result check area(DB,internal)
		'-----------------------
		'If Not (objXchRate.OperationStatusMessage = MSG_OK_STR) Then
		'	Call DisplayMsgBox(objXchRate.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		'	Set objXchRate = Nothing												'��: ComProxy Unload
		'	Call HideStatusWnd
		'	Response.End														'��: �����Ͻ� ���� ó���� ������ 
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
			Response.End																				'��: Process End
		End If

		'-----------------------
		'Result data display area
		'----------------------- 
%>
<Script Language=vbscript>
		With parent.frm1
		    .txtXchRate.Value = "<%=objXchRate.ExportBDailyExchangeRateStdRate%>"                          'ȯ�� 
		End With
</Script>
<%	 
		Set objXchRate = Nothing															    '��: Unload Comproxy
		Call HideStatusWnd
		Response.End																		'��: Process End   

	Case "CnclXchRate"   
	  '********************************************************  
	'              LOOKUP For Exchange Rate
	'********************************************************  

	    Err.Clear                                                               '��: Clear error no
	
		Set objCnclXchRate = Server.CreateObject("B17014.B17014LookupExchangeRate")
     
	    '-----------------------
		'Com action result check area(OS,internal)
	    '-----------------------
		If Err.Number <> 0 Then
			Set objCnclXchRate = Nothing																'��: ComProxy UnLoad
			Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)						'��:
			Call HideStatusWnd
			Response.End																		'��: Process End
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
			Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)                          '��:
			Set objCnclXchRate = Nothing																	    '��: ComProxy UnLoad
			Call HideStatusWnd
			Response.End																				'��: Process End
		End If
    
		'-----------------------
		'Com action result check area(DB,internal)
		'-----------------------
		'If Not (objCnclXchRate.OperationStatusMessage = MSG_OK_STR) Then
		'	Call DisplayMsgBox(objCnclXchRate.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		'	Set objCnclXchRate = Nothing												'��: ComProxy Unload
		'	Call HideStatusWnd
		'	Response.End														'��: �����Ͻ� ���� ó���� ������ 
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
			Response.End																				'��: Process End
		End If

		'-----------------------
		'Result data display area
		'----------------------- 
%>
<Script Language=vbscript>
		With parent.frm1
		    .txtCnclXchRate.Value = "<%=objCnclXchRate.ExportBDailyExchangeRateStdRate%>"                          'ȯ�� 
		End With
</Script>
<%	 
		Set objCnclXchRate = Nothing															    '��: Unload Comproxy
		Call HideStatusWnd
		Response.End																		'��: Process End   

End Select

%>
