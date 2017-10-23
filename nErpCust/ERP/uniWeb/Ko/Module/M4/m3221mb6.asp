<%
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3221mb6.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Import Local L/C Amend��� Query Transaction ó���� ASP					*
'*  7. Modified date(First) : 2000/05/02																*
'*  8. Modified date(Last)  : 2000/05/02																*
'*  9. Modifier (First)     :																			*
'* 10. Modifier (Last)      :																			*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              : 1. 2000/04/03 : Coding Start												*
'********************************************************************************************************

Response.Expires = -1															'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True															'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/IncServer.asp" -->
<%
																				'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

Call HideStatusWnd

Dim strMode																		'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

strMode = Request("txtMode")	

Select Case strMode
	Case CStr(UID_M0001)														'��: ���� ��ȸ/Prev/Next ��û�� ���� 
		Dim M32219																' Master L/C Header ��ȸ�� Object
		Dim strTransportMajor
		Dim ExAtTransportNm
		Dim ExBeTransportNm

		'-------------------- Minor Code Name�� ��ȸ�ϱ� ���� Major Code Setting -------------------
		strTransportMajor = ""

		Err.Clear																'��: Protect system from crashing

		If Request("txtLCAmdNo") = "" Then											'��: ��ȸ�� ���� ���� ���Դ��� üũ 
			Call ServerMesgBox("��ȸ ���ǰ��� ����ֽ��ϴ�!", vbInformation, I_MKSCRIPT)
			Response.End
		End If
		
		'---------------------------------- L/C Amend Header Data Query ----------------------------------

		Set M32219 = Server.CreateObject("M32219.M32219LookupLcAmendHdrSvr")

		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			Set M32219 = Nothing												'��: ComProxy UnLoad
			Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)
			Response.End														'��: Process End
		End If

		'-----------------------
		'Data manipulate  area(import view match)
		'-----------------------
		M32219.ImportMLcAmendHdrLcAmdNo = Request("txtLCAmdNo")
		M32219.CommandSent = "LOOKUP"
		M32219.ServerLocation = ggServerIP
		
		'-----------------------
		'Com action area
		'-----------------------
		M32219.ComCfg = gConnectionString
		M32219.Execute

		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			Set M32219 = Nothing												'��: ComProxy UnLoad
			Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)
			Response.End														'��: Process End
		End If

		'-----------------------
		'Com action result check area(DB,internal)
		'-----------------------
		If Not (M32219.OperationStatusMessage = MSG_OK_STR) Then
			Call DisplayMsgBox(M32219.OperationStatusMessage, vbInformation, "", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
			Set M32219 = Nothing												'��: ComProxy UnLoad
			Response.End														'��: Process End
		End If

		'-----------------------
		'Result data display area
		'-----------------------
		Const strDefDate = "1899-12-30"
%>
<Script Language=VBScript>
	With parent.frm1
	
		.txtLCDocNo.value = "<%=ConvSPChars(M32219.ExportMLcAmendHdrLcDocNo)%>"
		.txtLCAmendSeq.value = "<%=ConvSPChars(M32219.ExportMLcAmendHdrLcAmendSeq)%>"
		.txtApplicant.value = "<%=ConvSPChars(M32219.ExportMLcHdrApplicant)%>"
		.txtApplicantNm.value = "<%=ConvSPChars(M32219.ExportApplicantBBizPartnerBpNm)%>"
		
		Dim strDefDate 
		
		strDefDate = "<%=UNIDateClientFormat(strDefDate)%>"
		
		strDt = "<%=UNIDateClientFormat(M32219.ExportMLcAmendHdrAmendDt)%>"

		If strDt <> strDefDate Then
			.txtAmendDt.value = strDt
		End If
		
		If "<%=UNINumClientFormat(M32219.ExportMLcAmendHdrIncAmt, ggAmtOfMoney, 0)%>" <> "" Then
			.rdoAtDocAmt1.Checked = True
			.txtAmendAmt.value = "<%=UNINumClientFormat(M32219.ExportMLcAmendHdrIncAmt, ggAmtOfMoney, 0)%>"
		ElseIf "<%=UNINumClientFormat(M32219.ExportMLcAmendHdrDecAmt, ggAmtOfMoney.DecPoint, 0)%>" <> "" Then
			.rdoAtDocAmt2.Checked = True
			.txtAmendAmt.value = "<%=UNINumClientFormat(M32219.ExportMLcAmendHdrDecAmt, ggAmtOfMoney.DecPoint, 0)%>"
		End If

		.txtAtDocAmt.value = "<%=UNINumClientFormat(M32219.ExportMLcAmendHdrAtDocAmt, ggAmtOfMoney.DecPoint, 0)%>"

		strDt = "<%=UNIDateClientFormat(M32219.ExportMLcAmendHdrAtExpiryDt)%>"

		If strDt <> strDefDate Then
			.txtAtExpiryDt.value = strDt
			.txtHExpiryDt.value = strDt
		End If
		
		strDt = "<%=UNIDateClientFormat(M32219.ExportMLcAmendHdrAmendReqDt)%>"

		If strDt <> strDefDate Then
			.txtAmendReqDt.value = strDt
		End If
		
		strDt = "<%=UNIDateClientFormat(M32219.ExportMLcAmendHdrBeExpiryDt)%>"

		If strDt <> strDefDate Then
			.txtBeExpiryDt.value = strDt
		End If

		strDt = "<%=UNIDateClientFormat(M32219.ExportMLcAmendHdrAtLatestShipDt)%>"

		If strDt <> strDefDate Then
			.txtAtLatestShipDt.value = strDt
			.txtHLatestShipDt.value = strDt			
		End If
		
		strDt = "<%=UNIDateClientFormat(M32219.ExportMLcAmendHdrBeLatestShipDt)%>"

		If strDt <> strDefDate Then
			.txtBeLatestShipDt.value = strDt
		End If
		
		If "<%=ConvSPChars(M32219.ExportMLcAmendHdrAtPartialShipAsString)%>" = "Y" Then
			.rdoAtPartialShip1.Checked = True
		ElseIf "<%=ConvSPChars(M32219.ExportMLcAmendHdrAtPartialShipAsString)%>" = "N" Then
			.rdoAtPartialShip2.Checked = True
		End If
		
		.txtHTranshipment.value = "<%=ConvSPChars(M32219.ExportMLcAmendHdrAtTranshipmentAsString)%>"

		.txtBePartialShip.value = "<%=ConvSPChars(M32219.ExportMLcAmendHdrBePartialShipAsString)%>"
		.txtRemark.value = "<%=ConvSPChars(M32219.ExportMLcAmendHdrRemark)%>"
		.txtCurrency.value = "<%=ConvSPChars(M32219.ExportMLcAmendHdrCurrency)%>"
		.txtBeDocAmt.value = "<%=UNINumClientFormat(M32219.ExportMLcAmendHdrBeDocAmt, ggAmtOfMoney, 0)%>"
		.txtBeLocAmt.value = "<%=UNINumClientFormat(M32219.ExportMLcAmendHdrBeLocAmt, ggAmtOfMoney, 0)%>"

		strDt = "<%=UNIDateClientFormat(M32219.ExportMLcAmendHdrOpenDt)%>"

		If strDt <> strDefDate Then
			.txtOpenDt.value = strDt
		End If
		
		.txtOpenBank.value = "<%=ConvSPChars(M32219.ExportMLcAmendHdrOpenBank)%>"
		.txtOpenBankNm.value = "<%=ConvSPChars(M32219.ExportIssueBankBBankBankShNm)%>"
		.txtAdvNo.value = "<%=ConvSPChars(M32219.ExportMLcAmendHdrAdvNo)%>"
		.txtBeneficiary.value = "<%=ConvSPChars(M32219.ExportMLcAmendHdrBeneficiary)%>"
		.txtBeneficiaryNm.value = "<%=ConvSPChars(M32219.ExportBeneficiaryBBizPartnerBpNm)%>"
		.txtPurGrp.value = "<%=ConvSPChars(M32219.ExportMLcAmendHdrPurGrp)%>"
		.txtPurGrpNm.value = "<%=ConvSPChars(M32219.ExportBPurchaseGroupPurGrpNm)%>"
		.txtPurOrg.value = "<%=ConvSPChars(M32219.ExportMLcAmendHdrPurOrg)%>"
		.txtPurOrgNm.value = "<%=ConvSPChars(M32219.ExportBPurchaseOrganizationPurOrgNm)%>"
		.txtLCNo.value = "<%=ConvSPChars(M32219.ExportMLcAmendHdrLcNo)%>"		 
		.txtPONo.value = "<%=ConvSPChars(M32219.ExportMLcHdrPoNo)%>"		 
		.txtIncAmt.value = "<%=UNINumClientFormat(M32219.ExportMLcAmendHdrIncAmt, ggAmtOfMoney, 0)%>"
		.txtDecAmt.value = "<%=UNINumClientFormat(M32219.ExportMLcAmendHdrDecAmt, ggAmtOfMoney, 0)%>"
		.txtXchRate.value = "<%=UNINumClientFormat(M32219.ExportMLcAmendHdrXchRate, ggExchRate, 0)%>"

		Call parent.DbQueryOk()														'��: ��ȸ�� ���� 

		.txtHLCAmdNo.value = "<%=ConvSPChars(Request("txtLCAmdNo"))%>"

		Call parent.DbQueryOk()														'��: ��ȸ�� ���� 
	End With
</Script>
<%

'		Set B1a029 = Nothing														'��: ComProxy UnLoad
		Set M32219 = Nothing														'��: Unload Comproxy
		Response.End																'��: Process End

	Case Else
		Response.End
End Select
%>