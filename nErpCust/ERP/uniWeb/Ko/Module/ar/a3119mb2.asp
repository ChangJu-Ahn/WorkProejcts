<%
'======================================================================================================
'*  1. Module Name          : Finance
'*  2. Function Name        : Prepayment
'*  3. Program ID           : f6102mb2
'*  4. Program Name         : �Ա�û�� 
'*  5. Program Desc         : �Ա�û���� LookUp
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
Dim pAr0071																	'��ȸ�� ComProxy Dll ��� ���� 
Dim strCode																	'Lookup �� �ڵ� ���� ���� 

Dim strMode																	'���� MyBiz.asp �� ������¸� ��Ÿ�� 

strMode = Request("txtMode")												'���� ���¸� ���� 

Select Case strMode

Case CStr(UID_M0001)														'��: ���� ��ȸ/Prev/Next ��û�� ���� 

    Err.Clear                                                               '��: Protect system from crashing
    
    Set pAr0071 = Server.CreateObject("Ar0071.ALookupRcptAdjustSvr")

    '------------------------------------------
    'Com action result check area(OS,internal)
    '------------------------------------------
    If Err.Number <> 0 Then
		Set pAr0071 = Nothing																'��: ComProxy UnLoad
		Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)						'��������, �޼���Ÿ��, ��ũ��Ʈ���� 
		Response.End																		'��: Process End
	End If
    
    '------------------------------------------
    'Data manipulate  area(import view match)
    '------------------------------------------
	pAr0071.ImprotARcptRcptNo		= Request("txtRcptNo")
	pAr0071.ImportARcptAdjustAdjustNo = lgStrPrevKey
	'pAr0071.ImportIefSuppliedSelectChar	= Request("SelectChar")
	'pAr0071.ImportBAcctDeptOrgChangeId	= gChangeOrgID
    pAr0071.ServerLocation              = ggServerIP
    
    '------------------------------------------
    'Com action area
    '------------------------------------------
    pAr0071.ComCfg = gConnectionString
    pAr0071.Execute					

	'------------------------------------------
	'Com action result check area(OS,internal)
	'------------------------------------------
	If Err.Number <> 0 Then
	   Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)                      '��������, �޼���Ÿ��, ��ũ��Ʈ���� 
	   Set pAr0071 = Nothing													            '��: ComProxy UnLoad
	   Response.End																			'��: Process End
	End If
    
	'-----------------------
	'Com action result check area(DB,internal)
	'-----------------------
	If Not (pAr0071.OperationStatusMessage = MSG_OK_STR) Then
		Select Case pAr0071.OperationStatusMessage
			Case MSG_DEADLOCK_STR
				Call DisplayMsgBox2("999999", "25", "deadlock or timeout" , I_MKSCRIPT)
			Case MSG_DBERROR_STR
				Call DisplayMsgBox2(pAr0071.ExportErrEabSqlCodeSqlcode, _
						    pAr0071.ExportErrEabSqlCodeSeverity, _
						    pAr0071.ExportErrEabSqlCodeErrorMsg, I_MKSCRIPT)
			Case Else
				Call DisplayMsgBox(pAr0071.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		End Select

		Set pAr0071 = Nothing
		Response.End
	End If  

	'------------------------------------------
	'Result data display area
	'------------------------------------------
	' ����Ű�� ����Ű�� �������� ���� ��� Blank�� ������ ������ ������.
%>
<Script Language=vbscript>
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp
    Dim strData
	
	With parent.frm1
<%	
	Dim lgCurrency
	lgCurrency = ConvSPChars(pAr0071.ExportARcptDocCur)
%>	
		.txtRcptNo.value	   = "<%=ConvSPChars(pAr0071.ExportARcptRcptNo)%>"

		.txtDeptCd.value	   = "<%=ConvSPChars(pAr0071.ExportBAcctDeptDeptCd)%>"
		.txtDeptNm.value	   = "<%=ConvSPChars(pAr0071.ExportBAcctDeptDeptNm)%>"
		.txtRcptDt.text		   = "<%=UNIDateClientFormat(pAr0071.ExportARcptRcptDt)%>"
		.txtBpCd.value		   = "<%=ConvSPChars(pAr0071.ExportBBizPartnerBpCd)%>"
		.txtBpNm.value		   = "<%=ConvSPChars(pAr0071.ExportBBizPartnerBpNm)%>"
		.txtRefNo.value		   = "<%=ConvSPChars(pAr0071.ExportARcptRcptNo)%>"
		.txtDocCur.value	   = "<%=ConvSPChars(pAr0071.ExportARcptDocCur)%>"		
		
		.txtRcptAmt.value	   = "<%=UNIConvNumDBToCompanyByCurrency(pAr0071.ExportARcptRcptAmt, lgCurrency,ggAmtOfMoneyNo, "X" , "X")%>" 
		.txtRcptLocAmt.value   = "<%=UNIConvNumDBToCompanyByCurrency(pAr0071.ExportARcptRcptLocAmt, gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")%>" 
		.txtBalAmt.value	   = "<%=UNIConvNumDBToCompanyByCurrency(pAr0071.ExportARcptBalAmt, lgCurrency,ggAmtOfMoneyNo, "X" , "X")%>" 
		.txtBalLocAmt.value	   = "<%=UNIConvNumDBToCompanyByCurrency(pAr0071.ExportARcptBalLocAmt, gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")%>" 
		
		.txtGlNo.value		   = "<%=ConvSPChars(pAr0071.ExportAGlGlNo)%>"
		.txtRcptDesc.value   = "<%=ConvSPChars(pAr0071.ExportARcptRcptDesc)%>"

		parent.lgNextNo = ""		' ���� Ű �� �Ѱ��� 
		parent.lgPrevNo = ""		' ���� Ű �� �Ѱ��� , ���� ComProxy�� ����� �ȵ� ���� 
		
		parent.DbQueryOk															'��: ��ȸ�� ���� 
	End With
	
	With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 

	LngMaxRow = .frm1.vspdData.MaxRows										'Save previous Maxrow                                                
<%      
  	For LngRow = 1 To GroupCount
%>
        strData = strData & Chr(11) & "<%=LngRow%>"	'1  C_AdjustNo
        strData = strData & Chr(11) & "<%=UNIDateClientFormat(pAr0071.ExportARcptAdjustAdjustDt(LngRow))%>" '2 C_AdjustDt
        strData = strData & Chr(11) & "<%=ConvSPChars(pAr0071.ExportAAcctAcctCd(LngRow))%>"			'3  C_AcctCd
        strData = strData & Chr(11) & ""													'4  C_AcctCdPopUp
        strData = strData & Chr(11) & "<%=ConvSPChars(pAr0071.ExportAAcctAcctNm(LngRow))%>"  		'5  C_AcctNm 
		
        strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(pAr0071.ExportARcptAdjustAdjustAmt(LngRow), lgCurrency,ggAmtOfMoneyNo, "X" , "X")%>"
		strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(pAr0071.ExportARcptAdjustAdjustLocAmt(LngRow), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")%>"
		
		strData = strData & Chr(11) & "<%=ConvSPChars(pAr0071.ExportARcptAdjustDocDur(LngRow))%>"	'8  C_DocCur
        strData = strData & Chr(11) & ""													'9  C_DocCurPopUp
		strData = strData & Chr(11) & "<%=ConvSPChars(pAr0071.ExportItemFPrpaymSttlGlNo(LngRow))%>"		'10 C_GlNo
		strData = strData & Chr(11) & "<%=ConvSPChars(pAr0071.ExportAdjustAGlGlNo(LngRow))%>"		'11 C_RefNo

        strData = strData & Chr(11) & LngMaxRow + <%=LngRow%>									
        strData = strData & Chr(11) & Chr(12)
<%      
    Next
%>    
	.ggoSpread.Source = .frm1.vspdData 
	.ggoSpread.SSShowData strData
		
	.lgStrPrevKey = "<%=StrNextKey%>"

	.frm1.hRcptNo.value = "<%=Request("txtRcptNo")%>"

'	.DbQueryOk
				
	End With
</Script>
<%
	 
    Set pAr0071 = Nothing															'��: Unload Comproxy

	Response.End																	'��: Process End

End Select
%>
