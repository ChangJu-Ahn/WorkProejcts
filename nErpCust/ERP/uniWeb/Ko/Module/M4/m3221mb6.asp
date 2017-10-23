<%
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3221mb6.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Import Local L/C Amend등록 Query Transaction 처리용 ASP					*
'*  7. Modified date(First) : 2000/05/02																*
'*  8. Modified date(Last)  : 2000/05/02																*
'*  9. Modifier (First)     :																			*
'* 10. Modifier (Last)      :																			*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/03 : Coding Start												*
'********************************************************************************************************

Response.Expires = -1															'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True															'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/IncServer.asp" -->
<%
																				'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Call HideStatusWnd

Dim strMode																		'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

strMode = Request("txtMode")	

Select Case strMode
	Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 
		Dim M32219																' Master L/C Header 조회용 Object
		Dim strTransportMajor
		Dim ExAtTransportNm
		Dim ExBeTransportNm

		'-------------------- Minor Code Name을 조회하기 위한 Major Code Setting -------------------
		strTransportMajor = ""

		Err.Clear																'☜: Protect system from crashing

		If Request("txtLCAmdNo") = "" Then											'⊙: 조회를 위한 값이 들어왔는지 체크 
			Call ServerMesgBox("조회 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)
			Response.End
		End If
		
		'---------------------------------- L/C Amend Header Data Query ----------------------------------

		Set M32219 = Server.CreateObject("M32219.M32219LookupLcAmendHdrSvr")

		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			Set M32219 = Nothing												'☜: ComProxy UnLoad
			Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)
			Response.End														'☜: Process End
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
			Set M32219 = Nothing												'☜: ComProxy UnLoad
			Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)
			Response.End														'☜: Process End
		End If

		'-----------------------
		'Com action result check area(DB,internal)
		'-----------------------
		If Not (M32219.OperationStatusMessage = MSG_OK_STR) Then
			Call DisplayMsgBox(M32219.OperationStatusMessage, vbInformation, "", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
			Set M32219 = Nothing												'☜: ComProxy UnLoad
			Response.End														'☜: Process End
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

		Call parent.DbQueryOk()														'☜: 조회가 성공 

		.txtHLCAmdNo.value = "<%=ConvSPChars(Request("txtLCAmdNo"))%>"

		Call parent.DbQueryOk()														'☜: 조회가 성공 
	End With
</Script>
<%

'		Set B1a029 = Nothing														'☜: ComProxy UnLoad
		Set M32219 = Nothing														'☜: Unload Comproxy
		Response.End																'☜: Process End

	Case Else
		Response.End
End Select
%>