<%
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m4211mb5.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : 수입통관 등록에서 사용하기 위한 Open L/C 등록 Query Transaction 처리용 ASP*
'*  7. Modified date(First) : 2000/03/22																*
'*  8. Modified date(Last)  : 2000/03/22																*
'*  9. Modifier (First)     :																			*
'* 10. Modifier (Last)      :																			*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/22 : Coding Start												*
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

strMode = Request("txtMode")													'☜ : 현재 상태를 받음 

Select Case strMode
	Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 
		Dim M42119																' Master L/C Header 조회용 Object
		Dim B1a029																' Minor Code 조회용 Object
		Dim strLCTypeMajor
		Dim strTransportMajor
		Dim strPayTermsMajor
		Dim strFreightMajor
		Dim ExLCTypeNm
		Dim ExTransportNm
		Dim ExPayTermsNm
		Dim ExFreightNm

		'-------------------- Minor Code Name을 조회하기 위한 Major Code Setting -------------------
		strLCTypeMajor = ""
		strTransportMajor = ""
		strPayTermsMajor = ""
		strFreightMajor = ""

		Err.Clear																'☜: Protect system from crashing

		If Request("txtCCNo") = "" Then											'⊙: 조회를 위한 값이 들어왔는지 체크 
			Call ServerMesgBox("조회 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)
			Response.End
		End If
		
		'---------------------------------- L/C Header Data Query ----------------------------------

		Set M42119 = Server.CreateObject("M42119.M42119LookupImportCcHdrSvr")
		
		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			Set M42119 = Nothing												'☜: ComProxy UnLoad
			Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)
			Response.End														'☜: Process End
		End If

		'-----------------------
		'Data manipulate  area(import view match)
		'-----------------------
		M42119.ImportMCcHdrCcNo = Request("txtCCNo")
		M42119.CommandSent = "LOOKUP"
		M42119.ServerLocation = ggServerIP

		'-----------------------
		'Com action area
		'-----------------------
		M42119.ComCfg = gConnectionString
		M42119.Execute

		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			Set M42119 = Nothing												'☜: ComProxy UnLoad
			Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)
			Response.End														'☜: Process End
		End If

		'-----------------------
		'Com action result check area(DB,internal)
		'-----------------------
		If Not (M42119.OperationStatusMessage = MSG_OK_STR) Then
			Call DisplayMsgBox(M42119.OperationStatusMessage, vbInformation, "", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
			Set M42119 = Nothing												'☜: ComProxy UnLoad
			Response.End														'☜: Process End
		End If
		   
		'-----------------------
		'Result data display area
		'-----------------------

%>
<Script Language=VBScript>
	With parent.frm1
		.txtCCNo.value = "<%=ConvSPChars(M42119.ExportMCcHdrCcNo))%>"
		.txtPONo.value = "<%=ConvSPChars(M42119.ExportMCcHdrPoNo))%>"
		
		.txtIDNo.value = "<%=ConvSPChars(M42119.ExportMCcHdrIdNo)%>"
		.txtIDDt.value = "<%=M42119.ExportMCcHdrIdDt%>"
		.txtIDReqDt.value = "<%=M42119.ExportMCcHdrIdReqDt%>"
		.txtLoadingDt.value = "<%=UNIDateClientFormat(M42119.ExportMCcHdrLoadingDt)%>"
		.txtLoadingPort.value = "<%=ConvSPChars(M42119.ExportMCcHdrLoadingPort)%>"
		.txtLoadingCntry.value = "<%=ConvSPChars(M42119.ExportMCcHdrLoadingCntry)%>"
		.txtDischgeDt.value = "<%=UNIDateClientFormat(M42119.ExportMCcHdrDischgeDt)%>"
		.txtDischgePort.value = "<%=ConvSPChars(M42119.ExportMCcHdrDischgePort)%>"
		.txtWeightUnit.value = "<%=ConvSPChars(M42119.ExportMCcHdrWeightUnit)%>"
		.txtGrossWeight.value = "<%=ConvSPChars(M42119.ExportMCcHdrGrossWeight)%>"
		.txtTotPackingCnt.Text = "<%=ConvSPChars(M42119.ExportMCcHdrPackingCnt)%>"
		.txtPackingType.value = "<%=ConvSPChars(M42119.ExportMCcHdrPackingType)%>"
		.txtPurGrp.value = "<%=ConvSPChars(M42119.ExportMCcHdrPurGrp)%>"
		.txtPurGrpNm.value = "<%=ConvSPChars(ExName)%>"   
		.txtTransport.value = "<%=ConvSPChars(M42119.ExportMCcHdrTransport)%>"  
		.txtTransportNm.value = "<%=ConvSPChars(ExName)%>"
		.txtCurrency.value = "<%=ConvSPChars(M42119.ExportMCcHdrCurrency)%>"
		.txtXchRate.value = "<%=ConvSPChars(M42119.ExportMCcHdrXchRate)%>"
		.txtIDType.value = "<%=ConvSPChars(M42119.ExportMCcHdrIdType)%>"   
		.txtIDTypeNm.value = "<%=ConvSPChars(ExName)%>"     
		.txtVesselNm.value = "<%=ConvSPChars(ExName)%>"     
		.txtVesselCntry.value = "<%=ConvSPChars(M42119.ExportMCcHdrVesselCntry)%>"
		.txtPackingNo.value = "<%=ConvSPChars(M42119.ExportMCcHdrPackingNo)%>"
		.txtOrigin.value = "<%=ConvSPChars(M42119.ExportMCcHdrOrigin)%>"
		.txtOriginCntry.value = "<%=ConvSPChars(M42119.ExportMCcHdrOriginCntry)%>"
		.txtIPType.value = "<%=ConvSPChars(M42119.ExportMCcHdrIpType)%>"
		.txtIPTypeNm.value = "<%=ConvSPChars(ExName)%>"
		.txtExamTxt.value = "<%=ConvSPChars(M42119.ExportMCcHdrExamTxt)%>"
		.txtImportType.value = "<%=ConvSPChars(M42119.ExportMCcHdrImportType)%>"
		.txtBeneficiary.value = "<%=ConvSPChars(M42119.ExportMCcHdrBeneficiary)%>"
		.txtBeneficiaryNm.value = "<%=ConvSPChars(ExName)%>"
		.txtApplicant.value = "<%=ConvSPChars(M42119.ExportMCcHdrApplicant)%>"
		.txtApplicantNm.value = "<%=ConvSPChars(ExName)%>"
		.txtIPNo.value = "<%=ConvSPChars(M42119.ExportMCcHdrIpNo)%>"
		.txtIPDt.value = "<%=UNIDateClientFormat(M42119.ExportMCcHdrIpDt)%>"
		.txtCIFDocAmt.value = "<%=UNINumClientFormat(M42119.ExportMCcHdrCifDocAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.txtUSDXchRate.value = "<%=ConvSPChars(M42119.ExportMCcHdrUsdXchRate)%>"
		.txtCIFLocAmt.value = "<%=UNINumClientFormat(M42119.ExportMCcHdrCifLocAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.txtCustoms.value = "<%=ConvSPChars(M42119.ExportMCcHdrCustoms)%>"
		.txtCustomsNm.value = "<%=ConvSPChars(ExName)%>"
		.txtTariffTax.value = "<%=ConvSPChars(M42119.ExportMCcHdrTariffTax)%>"
		.txtTariffRate.value = "<%=ConvSPChars(M42119.ExportMCcHdrTariffRate)%>"
		.txtSpecialTax.value = "<%=ConvSPChars(M42119.ExportMCcHdrSpecialTax)%>"
		.txtEducTax.value = "<%=ConvSPChars(M42119.ExportMCcHdrEducTax)%>"
		.txtWineTax.value = "<%=ConvSPChars(M42119.ExportMCcHdrWineTax)%>"
		.txtArgriTax.value = "<%=ConvSPChars(M42119.ExportMCcHdrArgriTax)%>"
		.txtTrafTax.value = "<%=ConvSPChars(M42119.ExportMCcHdrTrafTax)%>"
		.txtVatAmt.value = "<%=UNINumClientFormat(M42119.ExportMCcHdrVatAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.txtVatRate.value = "<%=ConvSPChars(M42119.ExportMCcHdrVatRate)%>"
		.txtDeviceNo.value = "<%=ConvSPChars(M42119.ExportMCcHdrDeviceNo)%>"
		.txtDevicePlce.value = "<%=ConvSPChars(M42119.ExportMCcHdrDevicePlce)%>"
		.txtInputNo.value = "<%=ConvSPChars(M42119.ExportMCcHdrInputNo)%>"
		.txtInputDt.value = "<%=UNIDateClientFormat(M42119.ExportMCcHdrInputDt)%>"
		.txtOutputDt.value = "<%=UNIDateClientFormat(M42119.ExportMCcHdrOutputDt)%>"
		.txtCollectType.value = "<%=ConvSPChars(M42119.ExportMCcHdrCollectType)%>"
		.txtCollectTypeNm.value = "<%=ConvSPChars(ExName)%>"
		.txtCustomsExpDt.value = "<%=UNIDateClientFormat(M42119.ExportMCcHdrCustomsExpDt)%>"
		.txtPaymentNo.value = "<%=ConvSPChars(M42119.ExportMCcHdrPaymentNo)%>"
		.txtPaymentDt.value = "<%=UNIDateClientFormat(M42119.ExportMCcHdrPaymentDt)%>"
		.txtDvryDt.value = "<%=UNIDateClientFormat(M42119.ExportMCcHdrDvryDt)%>"
		.txtTaxBillNo.value = "<%=ConvSPChars(M42119.ExportMCcHdrTaxBillNo)%>"
		.txtTaxBillDt.value = "<%=UNIDateClientFormat(M42119.ExportMCcHdrTaxBillDt)%>"

		Call parent.DbQueryOk()														'☜: 조회가 성공 
	End With
</SCRIPT>
<%	
	Set M42119 = Nothing														'☜: Unload Comproxy
		Response.End																'☜: Process End

	Case Else
		Response.End
End Select
%>


