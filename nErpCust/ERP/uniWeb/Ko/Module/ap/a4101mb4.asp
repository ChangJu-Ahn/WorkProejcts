<%
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : Open Ap
'*  3. Program ID           : a4101mb4
'*  4. Program Name         : Vat�� ������ ���� Asp
'*  5. Program Desc         :
'*  6. Comproxy List        : +AP001M
'*  7. Modified date(First) : 2000/04/10
'*  8. Modified date(Last)  : 2000/04/10
'*  9. Modifier (First)     : You So Eun
'* 10. Modifier (Last)      : You So Eun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -1999/09/12 : ..........
'**********************************************************************************************

								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
								'�� : ASP�� ���ۿ� ����Ǿ� �������� �ٷ� Client�� ��������.


'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/IncServer.asp"  -->
<%																			'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next														'��: 

Dim pB1a059
    Dim intMaxRow
    Dim intLoopCnt
        
    Err.Clear                                                               '��: Clear error no
	On Error Resume Next

	Set pB1a059 = Server.CreateObject("B1a059.B1a059LookupConfiguration")
	
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		Set pB1a059 = Nothing												'��: ComProxy Unload
		Call SvrMsgBox(Err.description , vbInformation, I_MKSCRIPT)			
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If

	pB1a059.ImportBMajorMajorCd = "B9001"								'Major Code
	pB1a059.ImportBMinorMinorCd = Trim(Request("cboVatType")	)						'Major Code
	pB1a059.ImportBConfigurationSeqNo = 1								'Major Code
    pB1a059.ServerLocation = ggServerIP
    
	pB1a059.ComCfg = gConnectionString
    pB1a059.Execute															'��:
    
    If Err.Number <> 0 Then
		Set pB1a059 = Nothing												'��: ComProxy Unload
		Call SvrMsgBox(Err.description , vbInformation, I_MKSCRIPT)			
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If
    
    '-----------------------
	'Com action result check area(DB,internal)
	'-----------------------
	
	If Not (pB1a059.OperationStatusMessage = MSG_OK_STR) Then
		Select Case pB1a059.OperationStatusMessage
			Case MSG_DEADLOCK_STR
				Call DisplayMsgBox2("999999", "25", "deadlock or timeout" , I_MKSCRIPT)
			Case MSG_DBERROR_STR
				Call DisplayMsgBox2(pB1a059.ExportErrEabSqlCodeSqlcode, _
						    pB1a059.ExportErrEabSqlCodeSeverity, _
						    pB1a059.ExportErrEabSqlCodeErrorMsg, I_MKSCRIPT)
			Case Else
				Call DisplayMsgBox(pB1a059.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		End Select

		Set pB1a059 = Nothing
		Response.End
	End If
	
%>
<Script Language=vbscript>
    parent.frm1.txtVatRate.value = "<%=UNINumClientFormat(pB1a059.ExportBConfigurationReference, ggAmtOfMoney.DecPoint, 0)%>"
</Script>    
<%
	Set pB1a059 = Nothing
%>	
