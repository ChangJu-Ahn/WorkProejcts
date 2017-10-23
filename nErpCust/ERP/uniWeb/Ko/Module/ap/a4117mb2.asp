<%
'======================================================================================================
'*  1. Module Name          : Finance
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

Dim pFp0011																	'�Է�/������ ComProxy Dll ��� ���� 
Dim pAp0061																	'��ȸ�� ComProxy Dll ��� ���� 
Dim strCode																	'Lookup �� �ڵ� ���� ���� 

Dim strMode																	'���� MyBiz.asp �� ������¸� ��Ÿ�� 

strMode = Request("txtMode")												'���� ���¸� ���� 

Select Case strMode

Case CStr(UID_M0001)														'��: ���� ��ȸ/Prev/Next ��û�� ���� 

    Err.Clear                                                               '��: Protect system from crashing
    
    Set pAp0061 = Server.CreateObject("Ap0069.ALookupRcptAdjustSvr")

    '------------------------------------------
    'Com action result check area(OS,internal)
    '------------------------------------------
    If Err.Number <> 0 Then
		Set pAp0061 = Nothing																'��: ComProxy UnLoad
		Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)						'��������, �޼���Ÿ��, ��ũ��Ʈ���� 
		Response.End																		'��: Process End
	End If
    
    '------------------------------------------
    'Data manipulate  area(import view match)
    '------------------------------------------
	pAp0061.ImprotARcptRcptNo		= Request("txtApNo")
	'pAp0061.ImportIefSuppliedSelectChar	= Request("SelectChar")
	'pAp0061.ImportBAcctDeptOrgChangeId	= gChangeOrgID
    pAp0061.ServerLocation              = ggServerIP
    
    '------------------------------------------
    'Com action area
    '------------------------------------------
    pAp0061.ComCfg = gConnectionString
    pAp0061.Execute					

	'------------------------------------------
	'Com action result check area(OS,internal)
	'------------------------------------------
	If Err.Number <> 0 Then
	   Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)                      '��������, �޼���Ÿ��, ��ũ��Ʈ���� 
	   Set pAp0061 = Nothing													            '��: ComProxy UnLoad
	   Response.End																			'��: Process End
	End If
    
	'-----------------------
	'Com action result check area(DB,internal)
	'-----------------------
	If Not (pAp0061.OperationStatusMessage = MSG_OK_STR) Then
		Select Case pAp0061.OperationStatusMessage
			Case MSG_DEADLOCK_STR
				Call DisplayMsgBox2("999999", "25", "deadlock or timeout" , I_MKSCRIPT)
			Case MSG_DBERROR_STR
				Call DisplayMsgBox2(pAp0061.ExportErrEabSqlCodeSqlcode, _
						    pAp0061.ExportErrEabSqlCodeSeverity, _
						    pAp0061.ExportErrEabSqlCodeErrorMsg, I_MKSCRIPT)
			Case Else
				Call DisplayMsgBox(pAp0061.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		End Select

		Set pAp0061 = Nothing
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
	lgCurrency = ConvSPChars(pAp0061.ExportARcptDocCur)
%>	
		.txtApNo.value	   = "<%=ConvSPChars(pAp0061.ExportARcptRcptNo)%>"

		.txtDeptCd.value	   = "<%=ConvSPChars(pAp0061.ExportBAcctDeptDeptCd)%>"
		.txtDeptNm.value	   = "<%=ConvSPChars(pAp0061.ExportBAcctDeptDeptNm)%>"
		.txtApDt.text		   = "<%=UNIDateClientFormat(pAp0061.ExportARcptRcptDt)%>"
		.txtBpCd.value		   = "<%=ConvSPChars(pAp0061.ExportBBizPartnerBpCd)%>"
		.txtBpNm.value		   = "<%=ConvSPChars(pAp0061.ExportBBizPartnerBpNm)%>"
		.txtRefNo.value		   = "<%=ConvSPChars(pAp0061.ExportARcptRcptNo)%>"
		.txtDocCur.value	   = "<%=ConvSPChars(pAp0061.ExportARcptDocCur)%>"		
		
		.txtApAmt.value	    = "<%=UNIConvNumDBToCompanyByCurrency(pAp0061.ExportARcptRcptAmt, lgCurrency,ggAmtOfMoneyNo, "X" , "X")%>"
		.txtApLocAmt.value  = "<%=UNIConvNumDBToCompanyByCurrency(pAp0061.ExportARcptRcptLocAmt, gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")%>"
		.txtBalAmt.value	= "<%=UNIConvNumDBToCompanyByCurrency(pAp0061.ExportARcptBalAmt, lgCurrency,ggAmtOfMoneyNo, "X" , "X")%>"
		.txtBalLocAmt.value	= "<%=UNIConvNumDBToCompanyByCurrency(pAp0061.ExportARcptBalLocAmt, gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")%>"
		
		.txtGlNo.value		   = "<%=ConvSPChars(pAp0061.ExportAGlGlNo)%>"
		.txtApDesc.value   = "<%=ConvSPChars(pAp0061.ExportARcptRcptDesc)%>"

		parent.lgNextNo = ""		' ���� Ű �� �Ѱ��� 
		parent.lgPrevNo = ""		' ���� Ű �� �Ѱ��� , ���� ComProxy�� ����� �ȵ� ���� 
		
		parent.DbQueryOk															'��: ��ȸ�� ���� 
	End With
</Script>
<%
	 
    Set pAp0061 = Nothing															'��: Unload Comproxy

	Response.End																	'��: Process End

End Select
%>
