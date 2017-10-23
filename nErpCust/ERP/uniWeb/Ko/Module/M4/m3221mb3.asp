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
		
		Err.Clear																'☜: Protect system from crashing

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

%>
<Script Language=VBScript>
	With parent.frm1
	
		.txtLCDocNo.value = "<%=ConvSPChars(M32219.ExportMLcAmendHdrLcDocNo)%>"
		.txtLCAmendSeq.value = "<%=ConvSPChars(M32219.ExportMLcAmendHdrLcAmendSeq)%>"
		.txtBeneficiary.value = "<%=ConvSPChars(M32219.ExportBeneficiaryBBizPartnerBpCd)%>"
		.txtBeneficiaryNm.value = "<%=ConvSPChars(M32219.ExportBeneficiaryBBizPartnerBpNm)%>"		
		.txtAmendReqDt.text = "<%=UNIDateClientFormat(M32219.ExportMLcAmendHdrAmendReqDt)%>"				
		.txtAmendDt.text = "<%=UNIDateClientFormat(M32219.ExportMLcAmendHdrAmendDt)%>"
		
		If "<%=UNINumClientFormat(M32219.ExportMLcAmendHdrIncAmt, ggAmtOfMoney.DecPoint, 0)%>" <> "" Then
			.rdoAtDocAmt1.Checked = True
			.txtAmendAmt.text = "<%=UNINumClientFormat(M32219.ExportMLcAmendHdrIncAmt, ggAmtOfMoney.DecPoint, 0)%>"
		ElseIf "<%=UNINumClientFormat(M32219.ExportMLcAmendHdrDecAmt, ggAmtOfMoney.DecPoint, 0)%>" <> "" Then
			.rdoAtDocAmt2.Checked = True
			.txtAmendAmt.text = "<%=UNINumClientFormat(M32219.ExportMLcAmendHdrDecAmt, ggAmtOfMoney.DecPoint, 0)%>"
		End If

		.txtAtXchRate.text = "<%=UNINumClientFormat(M32219.ExportMLcAmendHdrAtXchRate, ggExchRate.DecPoint, 0)%>"
		.txtCurrency.value = "<%=ConvSPChars(M32219.ExportMLcAmendHdrCurrency)%>"
		.txtAtDocAmt.text = "<%=UNINumClientFormat(M32219.ExportMLcAmendHdrAtDocAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.txtAtLocAmt.text = "<%=UNINumClientFormat(M32219.ExportMLcAmendHdrAtLocAmt, ggAmtOfMoney.DecPoint, 0)%>"
		
		strDt = "<%=UNIDateClientFormat(M32219.ExportMLcAmendHdrAtExpiryDt)%>"

		
		.txtAtExpireDt.text = strDt
		.txtHExpiryDt.value = strDt
		
		
		strDt = "<%=UNIDateClientFormat(M32219.ExportMLcAmendHdrBeExpiryDt)%>"

		.txtBeExpireDt.text = strDt
		

		strDt = "<%=UNIDateClientFormat(M32219.ExportMLcAmendHdrAtLatestShipDt)%>"

		
		.txtAtLatestShipDt.text = strDt
		.txtHLatestShipDt.value = strDt			
		
		
		strDt = "<%=UNIDateClientFormat(M32219.ExportMLcAmendHdrBeLatestShipDt)%>"

		
		.txtBeLatestShipDt.text = strDt
		
		
		If "<%=ConvSPChars(M32219.ExportMLcAmendHdrAtPartialShip)%>" = "Y" Then
			.rdoAtPartialShip1.Checked = True
		ElseIf "<%=ConvSPChars(M32219.ExportMLcAmendHdrAtPartialShip)%>" = "N" Then
			.rdoAtPartialShip2.Checked = True
		End If
		
		.txtBePartialShip.value = "<%=ConvSPChars(M32219.ExportMLcAmendHdrBePartialShip)%>"
		.txtOpenBank.value = "<%=ConvSPChars(M32219.ExportIssueBankBBankBankCd)%>"
		.txtOpenBankNm.value = "<%=ConvSPChars(M32219.ExportIssueBankBBankBankNm)%>"
		.txtAdvBank.value = "<%=ConvSPChars(M32219.ExportAdviseBankBBankBankCd)%>"
		.txtAdvBankNm.value = "<%=ConvSPChars(M32219.ExportAdviseBankBBankBankNm)%>"
		
		strDt = "<%=UNIDateClientFormat(M32219.ExportMLcAmendHdrOpenDt)%>"
		.txtOpenDt.text = strDt
		
		
		.txtApplicant.value = "<%=ConvSPChars(M32219.ExportApplicantBBizPartnerBpCd)%>"
		.txtApplicantNm.value = "<%=ConvSPChars(M32219.ExportApplicantBBizPartnerBpNm)%>"
		.txtPurGrp.value = "<%=ConvSPChars(M32219.ExportBPurGrpPurGrp)%>"
		.txtPurGrpNm.value = "<%=ConvSPChars(M32219.ExportBPurGrpPurGrpNm)%>"
		.txtPurOrg.value = "<%=ConvSPChars(M32219.ExportBPurOrgPurOrg)%>"
		.txtPurOrgNm.value = "<%=ConvSPChars(M32219.ExportBPurOrgPurOrgNm)%>"
		.txtRemark.value = "<%=ConvSPChars(M32219.ExportMLcAmendHdrRemark)%>"
		.txtAdvNo.value = "<%=ConvSPChars(M32219.ExportMLcAmendHdrAdvNo)%>"
		.txtPreAdvRef.value = "<%=ConvSPChars(M32219.ExportMLcAmendHdrPreAdvRef)%>"
		.txtLCAmt.value = "<%=UNINumClientFormat(M32219.ExportMLcHdrDocAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.txtHLCNo.value = "<%=ConvSPChars(M32219.ExportMLcHdrLcNo)%>"
		
		Call parent.DbQueryOk()														'☜: 조회가 성공 

		.txtHLCAmdNo.value = "<%=ConvSPChars(Request("txtLCAmdNo"))%>"

		Call parent.DbQueryOk()														'☜: 조회가 성공 
	End With
</Script>
<%

		Set M32219 = Nothing														'☜: Unload Comproxy

		Response.End																'☜: Process End

	Case CStr(UID_M0002)														'☜: 현재 Save 요청을 받음 
		Dim M32211																' Master L/C Amend Header Save용 Object
		
		Err.Clear																'☜: Protect system from crashing

		lgIntFlgMode = CInt(Request("txtFlgMode"))								'☜: 저장시 Create/Update 판별 
	
		'⊙: 각 화면당 Relation이 되어 있지 않는 Field들에 대해서는 Lookup을 행한다.

		'⊙: Lookup Pad 동작후 정상적인 데이타 이면, 저장 로직 시작 
		Set M32211 = Server.CreateObject("M32211.M32211MaintLcAmendHdrSvr")

		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			Set M32211 = Nothing												'☜: ComProxy UnLoad
			Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)
			Response.End														'☜: Process End
		End If

		'-----------------------
		'Data manipulate  area(import view match)
		'-----------------------
		M32211.ImportMLcAmendHdrLcAmdNo = Trim(Request("txtLCAmdNo"))
		M32211.ImportMLcAmendHdrLcDocNo = UCase(Trim(Request("txtLCDocNo")))
		M32211.ImportMLcAmendHdrLcAmendSeq = Trim(Request("txtLCAmendSeq"))
		M32211.ImportBeneficiaryBBizPartnerBpCd = UCase(Trim(Request("txtBeneficiary")))
		
		If Len(Trim(Request("txtAmendReqDt"))) Then
			M32211.ImportMLcAmendHdrAmendReqDt = UNIConvDate(Request("txtAmendReqDt"))
		End If		
		
		If Len(Trim(Request("txtAmendDt"))) Then
			M32211.ImportMLcAmendHdrAmendDt = UNIConvDate(Request("txtAmendDt"))
		End If

		If Request("rdoAtDocAmt") = "I" Then
			If Len(Trim(Request("txtAmendAmt"))) Then
				M32211.ImportMLcAmendHdrIncAmt = UNIConvNum(Request("txtAmendAmt"),0)
			End If
		ElseIf Request("rdoAtDocAmt") = "D" Then
			If Len(Trim(Request("txtAmendAmt"))) Then
				M32211.ImportMLcAmendHdrDecAmt = UNIConvNum(Request("txtAmendAmt"),0)
			End If
		End If
		
		M32211.ImportMLcAmendHdrCurrency = UCase(Trim(Request("txtCurrency")))
		M32211.ImportMLcAmendHdrAtDocAmt = UNIConvNum(Request("txtAtDocAmt"),0)
		M32211.ImportMLcAmendHdrAtXchRate = UNIConvNum(Request("txtAtXchRate"),0)	

		If Len(Trim(Request("txtAtLocAmt"))) Then
			M32211.ImportMLcAmendHdrAtLocAmt = UNIConvNum(Request("txtBeDocAmt"),0)
		End If
		
		If Len(Trim(Request("txtAtExpireDt"))) Then
			M32211.ImportMLcAmendHdrAtExpiryDt = Request("txtAtExpireDt")
		End If
		
		If Len(Trim(Request("txtBeExpireDt"))) Then
			M32211.ImportMLcAmendHdrBeExpiryDt = Request("txtBeExpireDt")
		End If
		
		If Len(Trim(Request("txtAtLatestShipDt"))) Then
			M32211.ImportMLcAmendHdrAtLatestShipDt = Request("txtatLatestShipDt")
		End If

		If Len(Trim(Request("txtBeLatestShipDt"))) Then
			M32211.ImportMLcAmendHdrBeLatestShipDt = Request("txtBeLatestShipDt")
		End If

		If Not ISEMPTY(Request("chkAtPartialShip")) Then
			M32211.ImportMLcAmendHdrAtPartialShip = Request("rdoAtPartialShip")
		End If
		
		M32211.ImportMLcAmendHdrBePartialShip = Trim(Request("txtBePartialShip"))
		M32211.ImportMLcAmendHdrOpenBank = Trim(Request("txtOpenBank"))
		M32211.ImportMLcAmendHdrAdvBank = Trim(Request("txtAdvBank"))
		
		
		If Len(Trim(Request("txtBeLatestShipDt"))) Then
			M32211.ImportMLcAmendHdrOpenDt = UNIConvDate(Request("txtOpenDt"))
		End If
		
		M32211.ImportApplicantBBizPartnerBpCd = UCase(Trim(Request("txtApplicant")))
		M32211.ImportBPurGrpPurGrp = UCase(Trim(Request("txtPurGrp")))
		M32211.ImportMLcAmendHdrRemark = Trim(Request("txtRemark"))
		M32211.ImportMLcAmendHdrPreAdvRef = Trim(Request("txtPreAdvRef"))   

		M32211.ImportMLcAmendHdrLcKindAsString = "L"
		M32211.ImportMLcAmendHdrLcNo = Trim(Request("txtHLCNo"))
		M32211.ImportSWksUserUserId = UCase(Trim(Request("txtInsrtUserId")))
		
		If lgIntFlgMode = OPMD_CMODE Then
			M32211.CommandSent = "CREATE"
		ElseIf lgIntFlgMode = OPMD_UMODE Then
			M32211.CommandSent = "UPDATE"
		End If

		M32211.ServerLocation = ggServerIP

		'-----------------------
		'Com action area
		'-----------------------
		M32211.ComCfg = gConnectionString
		M32211.Execute

		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			Set M32211 = Nothing												'☜: ComProxy UnLoad
			Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)
			Response.End														'☜: Process End
		End If

		'-----------------------
		'Com action result check area(DB,internal)
		'-----------------------
		If Not (M32211.OperationStatusMessage = MSG_OK_STR) Then
			Call DisplayMsgBox(M32211.OperationStatusMessage, vbInformation, "", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
			Set M32211 = Nothing												'☜: ComProxy UnLoad
			Response.End														'☜: Process End
		End If

		'-----------------------
		'Result data display area
		'-----------------------
%>
<Script Language=VBScript>
	With parent
		.frm1.txtLCAmdNo.value = "<%=ConvSPChars(M32211.ExportMLcAmendHdrLcAmdNo)%>"
		.DbSaveOk
	End With
</Script>
<%
		Set M32211 = Nothing														'☜: Unload Comproxy

		Response.End																'☜: Process End

	Case CStr(UID_M0003)														'☜: 삭제 요청 
		
		Err.Clear																'☜: Protect system from crashing
        
            DIM Command 
    DIM I1_b_biz_partner_bp_cd 
    DIM I2_b_biz_partner_bp_cd 
    DIM I3_m_lc_amend_hdr 
    DIM I4_s_wks_user 
    DIM I5_b_pur_grp 
    dim strConvDt
    
     Const M435_I3_lc_amd_no = 0
     Const M435_I3_lc_no = 1
     Const M435_I3_lc_doc_no = 2
     Const M435_I3_lc_amend_seq = 3
     Const M435_I3_adv_no = 4
     Const M435_I3_pre_adv_ref = 5
     Const M435_I3_open_dt = 6
     Const M435_I3_be_expiry_dt = 7
     Const M435_I3_at_expiry_dt = 8
     Const M435_I3_manufacturer = 9
     Const M435_I3_agent = 10
     Const M435_I3_amend_dt = 11
     Const M435_I3_amend_req_dt = 12
     Const M435_I3_currency = 13
     Const M435_I3_be_doc_amt = 14
     Const M435_I3_at_doc_amt = 15
     Const M435_I3_at_xch_rate = 16
     Const M435_I3_inc_amt = 17
     Const M435_I3_dec_amt = 18
     Const M435_I3_be_loc_amt = 19
     Const M435_I3_at_loc_amt = 20
     Const M435_I3_be_partial_ship = 21
     Const M435_I3_at_partial_ship = 22
     Const M435_I3_be_latest_ship_dt = 23
     Const M435_I3_at_latest_ship_dt = 24
     Const M435_I3_open_bank = 25
     Const M435_I3_be_xch_rate = 26
     Const M435_I3_ext1_amt = 27
     Const M435_I3_ext1_cd = 28
     Const M435_I3_remark = 29
     Const M435_I3_lc_kind = 30
     Const M435_I3_remark2 = 31
     Const M435_I3_be_transhipment = 32
     Const M435_I3_at_transhipment = 33
     Const M435_I3_be_transfer = 34
     Const M435_I3_at_transfer = 35
     Const M435_I3_be_loading_port = 36
     Const M435_I3_at_loading_port = 37
     Const M435_I3_be_dischge_port = 38
     Const M435_I3_at_dischge_port = 39
     Const M435_I3_be_transport = 40
     Const M435_I3_at_transport = 41
     Const M435_I3_biz_area = 42
     Const M435_I3_charge_flg = 43
     Const M435_I3_adv_bank = 44
     Const M435_I3_ext1_qty = 45
     Const M435_I3_ext2_qty = 46
     Const M435_I3_ext3_qty = 47
     Const M435_I3_ext2_amt = 48
     Const M435_I3_ext3_amt = 49
     Const M435_I3_ext2_cd = 50
     Const M435_I3_ext3_cd = 51
     Const M435_I3_ext1_rt = 52
     Const M435_I3_ext2_rt = 53
     Const M435_I3_ext3_rt = 54
     Const M435_I3_ext1_dt = 55
     Const M435_I3_ext2_dt = 56
     Const M435_I3_ext3_dt = 57
     

    
    DIM  lgIntFlgMode
    DIM PM4G211
    Set PM4G211 = Server.CreateObject("PM4G211.cMMaintLcAmendHdrS")
    
		If Request("txtLCAmdNo") = "" Then										'⊙: 삭제를 위한 값이 들어왔는지 체크 
			Call ServerMesgBox("삭제 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
			Response.End 
		End If

		Set PM4G211 = Server.CreateObject("PM4G211.cMMaintLcAmendHdrS")
	

		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		 If CheckSYSTEMError(Err,True) = True Then
			Set PM4G211 = Nothing
			Exit Sub
		 End If
		
        
        I3_m_lc_amend_hdr(M435_I3_lc_amd_no) = UCase(Trim(Request("txtLCAmdNo")))
        Command = "Delete" 
   
    	call  PM4G211.MAINT_LC_AMEND_HDR_SVR(gStrGlobalCollection , Command, I1_b_biz_partner_bp_cd,I2_b_biz_partner_bp_cd,I3_m_lc_amend_hdr,I4_s_wks_user,I5_b_pur_grp) 
	
			    
		if CheckSYSTEMError2(Err,True,"","","","","") = true then 		
			Set PM4G211 = Nothing												'☜: ComProxy Unload
			Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)						'⊙:
			Response.End														'☜: 비지니스 로직 처리를 종료함 
 		end if

		'Data manipulate  area(import view match)
		'-----------------------
		
	
				             
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "Parent.DBSaveOK "           & vbCr
		Response.Write "</Script>"                  & vbCr 
                                    '☜: Clear Error status
		

		'-----------------------
		'Result data display area
		'-----------------------
%>
<Script Language=VBScript>
	With parent
		.DbDeleteOk
	End With
</Script>
<%
		Set M32211 = Nothing														'☜: Unload Comproxy

		Response.End																'☜: Process End

	Case Else
		Response.End
End Select
%>
