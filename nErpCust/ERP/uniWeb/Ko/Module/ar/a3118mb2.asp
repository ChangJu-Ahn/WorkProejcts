<%
'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Prepayment
'*  3. Program ID           : f6102mb2
'*  4. Program Name         : ä��û�� 
'*  5. Program Desc         : ä��û���� LookUp
'*  6. Modified date(First) : 2000/09/27
'*  7. Modified date(Last)  : 2000/12/20
'*  8. Modifier (First)     : ���ͼ� 
'*  9. Modifier (Last)      : hersheys
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
								'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/IncServer.asp"  -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

Call HideStatusWnd

On Error Resume Next														'��: 
																	'�Է�/������ ComProxy Dll ��� ���� 
Dim pAr0081																	'��ȸ�� ComProxy Dll ��� ���� 
Dim strCode																	'Lookup �� �ڵ� ���� ���� 

Dim strMode																	'���� MyBiz.asp �� ������¸� ��Ÿ�� 

strMode = Request("txtMode")												'���� ���¸� ���� 

Select Case strMode

Case CStr(UID_M0001)														'��: ���� ��ȸ/Prev/Next ��û�� ���� 

    Err.Clear                                                               '��: Protect system from crashing
    
    Set pAr0081 = Server.CreateObject("Ar0089.ALookupArAdjustSvrv")

    '------------------------------------------
    'Com action result check area(OS,internal)
    '------------------------------------------
    If Err.Number <> 0 Then
		Set pAr0081 = Nothing																'��: ComProxy UnLoad
		Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)						'��������, �޼���Ÿ��, ��ũ��Ʈ���� 
		Response.End																		'��: Process End
	End If
    
    '------------------------------------------
    'Data manipulate  area(import view match)
    '------------------------------------------
	pAr0081.ImportAOpenArArNo		= Trim(Request("txtArNo"))
	'pAr0081.ImportIefSuppliedSelectChar	= Request("SelectChar")
	'pAr0081.ImportBAcctDeptOrgChangeId	= gChangeOrgID
    pAr0081.ServerLocation              = ggServerIP
    
    '------------------------------------------
    'Com action area
    '------------------------------------------
    pAr0081.ComCfg = gConnectionString
    pAr0081.Execute					

	'------------------------------------------
	'Com action result check area(OS,internal)
	'------------------------------------------
	If Err.Number <> 0 Then
	   Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)                      '��������, �޼���Ÿ��, ��ũ��Ʈ���� 
	   Set pAr0081 = Nothing													            '��: ComProxy UnLoad
	   Response.End																			'��: Process End
	End If
    
	'------------------------------------------
	'Com action result check area(DB,internal)
	'------------------------------------------
	If Not (pAr0081.OperationStatusMessage = MSG_OK_STR) Then
	    Select Case pAr0081.OperationStatusMessage
            Case MSG_DEADLOCK_STR
                Call DisplayMsgBox2("999999", "25", "deadlock or timeout" , I_MKSCRIPT)
            Case MSG_DBERROR_STR
				Call DisplayMsgBox2(pAr0081.ExportErrEabSqlCodeSqlcode, _
						    pAr0081.ExportErrEabSqlCodeSeverity, _
						    pAr0081.ExportErrEabSqlCodeErrorMsg, I_MKSCRIPT)
            Case Else
                Call DisplayMsgBox(pAr0081.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
            End Select

		Set pAr0081 = Nothing
		Response.End 
	End If

	'------------------------------------------
	'Result data display area
	'------------------------------------------
	' ����Ű�� ����Ű�� �������� ���� ��� Blank�� ������ ������ ������.
%>
<Script Language=vbscript>
	With parent.frm1
<%	
	Dim lgCurrency
	lgCurrency = ConvSPChars(pAr0081.ExportAOpenArDocCur)
%>	
		.txtArNo.value			= "<%=ConvSPChars(pAr0081.ExportAOpenArArNo)%>"
		.txtDeptCd.value	   = "<%=ConvSPChars(pAr0081.ExportBAcctDeptDeptCd)%>"
		.txtDeptNm.value	   = "<%=ConvSPChars(pAr0081.ExportBAcctDeptDeptNm)%>"
		.txtArDt.text		   = "<%=UNIDateClientFormat(pAr0081.ExportAOpenArArDt)%>"
		.txtBpCd.value		   = "<%=ConvSPChars(pAr0081.ExportBBizPartnerBpCd)%>"
		.txtBpNm.value		   = "<%=ConvSPChars(pAr0081.ExportBBizPartnerBpNm)%>"
		.txtRefNo.value		   = "<%=ConvSPChars(pAr0081.ExportAOpenArRefNo)%>"
		.txtDocCur.value	   = "<%=ConvSPChars(pAr0081.ExportAOpenArDocCur)%>"		
		
		.txtArAmt.value			= "<%=UNIConvNumDBToCompanyByCurrency(pAr0081.ExportAOpenArArAmt, lgCurrency,ggAmtOfMoneyNo, "X" , "X")%>" 
		.txtArLocAmt.value		= "<%=UNIConvNumDBToCompanyByCurrency(pAr0081.ExportAOpenArArLocAmt, gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")%>" 
		.txtBalAmt.value	   = "<%=UNIConvNumDBToCompanyByCurrency(pAr0081.ExportAOpenArArAmt - opAr0081.ExportAOpenArClsAmt - ExportAOpenArAdjustAmt, lgCurrency,ggAmtOfMoneyNo, "X" , "X")%>" 
		.txtBalLocAmt.value	   = "<%=UNIConvNumDBToCompanyByCurrency(pAr0081.ExportAOpenArArLocAmt - opAr0081.ExportAOpenArClsLocAmt - ExportAOpenArAdjustLocAmt,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")%>".txtGlNo.value		   = "<%=pAr0081.ExportAGlGlNo%>"
		.txtGlNo.value		   = "<%=ConvSPChars(opAr0081.ExportAGlGlNo)%>"
		.txtArDesc.value		= "<%=ConvSPChars(pAr0081.ExportAOpenArArDesc)%>"

		parent.lgNextNo = ""		' ���� Ű �� �Ѱ��� 
		parent.lgPrevNo = ""		' ���� Ű �� �Ѱ��� , ���� ComProxy�� ����� �ȵ� ���� 
		
		parent.DbQueryOk															'��: ��ȸ�� ���� 
	End With
</Script>
<%
	 
    Set pAr0081 = Nothing															'��: Unload Comproxy

	Response.End																	'��: Process End

End Select
%>
