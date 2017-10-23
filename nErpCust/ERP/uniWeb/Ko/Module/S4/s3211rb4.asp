<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s3211ma4.asp																*
'*  4. Program Name         : L/C 상세정보(L/C등록에서)													*
'*  5. Program Desc         : L/C 상세정보(L/C등록에서)													*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/07/12																*
'*  8. Modified date(Last)  : 2000/08/29																*
'*  9. Modifier (First)     : An ChangHwan 																*
'* 10. Modifier (Last)      : Son bum Yeol																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/20 : 화면 design												*
'*							  2. 2000/07/12 : Coding ReStart											*
'*																										*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
																				
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "*", "NOCOOKIE", "RB")   
Call LoadBNumericFormatB("Q","*","NOCOOKIE","RB")

On Error Resume Next


Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0			                   '☜ : DBAgent Parameter 선언 
Dim lgSelectList1
Dim lgSelectList2
Dim lgSelectList3
Dim lgSelectList4
Dim lgSelectFrom

Dim arrRsVal(81)								'☜ : QueryData()실행시 레코드셋을 배열로 받을때 사용 
    
Call HideStatusWnd

'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
lgSelectList1= ""
lgSelectList2= ""
lgSelectList3= ""
lgSelectList4= ""
lgSelectFrom = ""

lgSelectList1 = lgSelectList1 & "a.lc_no, a.lc_doc_no, a.lc_amend_seq, a.adv_no, "
lgSelectList1 = lgSelectList1 & "a.lc_type, lc.minor_nm, a.adv_dt, a.advise_bank_cd, adv.bank_nm,"
lgSelectList1 = lgSelectList1 & "a.expiry_dt, a.issue_bank_cd, iss.bank_nm, a.open_dt, a.cur, a.lc_amt, "
lgSelectList1 = lgSelectList1 & "a.lc_loc_amt, a.xch_rate, a.applicant, app.bp_nm, a.beneficiary, ben.bp_nm, "
lgSelectList1 = lgSelectList1 & "a.amt_tolerance, a.incoterms, it.minor_nm, b.sales_grp, b.sales_grp_nm, "

lgSelectList2 = lgSelectList2 & "a.pay_meth, pm.minor_nm, a.pay_dur, a.latest_ship_dt, "
lgSelectList2 = lgSelectList2 & "a.transport, tp.minor_nm, a.trnshp_flag, a.partial_ship_flag, "
lgSelectList2 = lgSelectList2 & "a.loading_port, lp.minor_nm, a.dischge_port, dp.minor_nm,  "
lgSelectList2 = lgSelectList2 & "a.delivery_plce, a.file_dt, a.file_dt_txt, a.inv_cnt, a.pack_list, "
lgSelectList2 = lgSelectList2 & "a.cert_origin_flag, a.bl_awb_flg, a.freight, fr.minor_nm, "

lgSelectList3 = lgSelectList3 & "a.notify_party, notify.bp_nm, a.consignee, a.insur_policy, "
lgSelectList3 = lgSelectList3 & "a.doc1, a.doc2, a.doc3, a.doc4, a.doc5, a.pay_bank_cd, pay.bank_nm, "
lgSelectList3 = lgSelectList3 & "a.renego_bank_cd, ren.bank_nm, a.confirm_bank_cd, con.bank_nm, a.bank_txt, "
lgSelectList3 = lgSelectList3 & "a.transfer_flag, a.credit_core, cc.minor_nm, a.charge_cd, "
lgSelectList3 = lgSelectList3 & "a.charge_txt, a.payment_txt, a.shipment, a.pre_adv_ref, a.transport_comp, "

lgSelectList4 = lgSelectList4 & "a.origin, og.minor_nm,  a.origin_cntry, cntry.country_nm, "
lgSelectList4 = lgSelectList4 & "a.agent, agent.bp_nm, a.manufacturer, man.bp_nm, "
lgSelectList4 = lgSelectList4 & "a.remark, a.amend_dt "

lgSelectFrom = lgSelectFrom & "LEFT OUTER JOIN dbo.b_minor lc   ON (  a.lc_type=lc.minor_cd AND  lc.major_cd = " & FilterVar("S9000", "''", "S") & ") "
lgSelectFrom = lgSelectFrom & "LEFT OUTER JOIN dbo.b_minor it   ON (  a.incoterms=it.minor_cd AND  it.major_cd = " & FilterVar("B9006", "''", "S") & ") "
lgSelectFrom = lgSelectFrom & "LEFT OUTER JOIN dbo.b_minor pm   ON (  a.pay_meth=pm.minor_cd AND  pm.major_cd = " & FilterVar("B9004", "''", "S") & ") "
lgSelectFrom = lgSelectFrom & "LEFT OUTER JOIN dbo.b_minor tp   ON (  a.transport=tp.minor_cd AND  tp.major_cd = " & FilterVar("B9009", "''", "S") & ") " 
lgSelectFrom = lgSelectFrom & "LEFT OUTER JOIN dbo.b_minor lp   ON (  a.loading_port=lp.minor_cd AND  lp.major_cd = " & FilterVar("B9092", "''", "S") & ") "
lgSelectFrom = lgSelectFrom & "LEFT OUTER JOIN dbo.b_minor dp   ON (  a.dischge_port=dp.minor_cd AND  dp.major_cd = " & FilterVar("B9092", "''", "S") & ") "
lgSelectFrom = lgSelectFrom & "LEFT OUTER JOIN dbo.b_minor fr   ON (  a.freight=fr.minor_cd AND  fr.major_cd = " & FilterVar("S9007", "''", "S") & ") "
lgSelectFrom = lgSelectFrom & "LEFT OUTER JOIN dbo.b_minor cc   ON (  a.credit_core=cc.minor_cd AND  cc.major_cd = " & FilterVar("S9003", "''", "S") & ") "
lgSelectFrom = lgSelectFrom & "LEFT OUTER JOIN dbo.b_minor og   ON (  a.origin=og.minor_cd AND  og.major_cd = " & FilterVar("B9094", "''", "S") & ") "


Call FixUNISQLData()
Call QueryData()

'=============================================================================================================
Sub FixUNISQLData()	
    
    Dim strMode	
    Dim strVal		
    	
    Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
																		  '조회화면에서 필요한 query조건문들의 영역(Statements table에 있음)
    Redim UNIValue(0,5)													  '☜: 각각의 SQL ID와 입력될 where 조건의 쌍으로 된 2차원 배열 

    UNISqlId(0) = "S3211RA401"  ' main query(spread sheet에 뿌려지는 query statement)
   
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList1         '☜: Select list
    UNIValue(0,1) = lgSelectList2         '☜: Select list
    UNIValue(0,2) = lgSelectList3         '☜: Select list
    UNIValue(0,3) = lgSelectList4         '☜: Select list
    UNIValue(0,4) = lgSelectFrom
																		  
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	strVal = ""
	strMode = Request("txtMode")														'☜ : 현재 상태를 받음	
	if strMode =CStr(UID_M0001) then										
		Err.Clear															
		If Trim(Request("txtLCNo")) = "" Then											
			Call ServerMesgBox("조회 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)
			Response.End			
		End If	
	end if
	
	If Len(Trim(Request("txtLCNo"))) Then
		strVal= strVal & " " & FilterVar(Request("txtLCNo"), "''", "S") & " "
	END If
	
	UNIValue(0,5) = strVal    '	UNISqlId(0)의 두번째 ?에 입력됨	
        
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------    
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode 
End Sub

'=============================================================================================================
Sub QueryData()
	
	Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
	Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
	Dim iStr
	Dim iCnt
	
	'BlankchkFlg = False
	Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")	
	lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	
	iStr = Split(lgstrRetMsg,gColSep)    
    
	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
	
	If  rs0.EOF And rs0.BOF Then
	    Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
	    rs0.Close
	    Set rs0 = Nothing
    
	' 이 위치에 있는 Response.End 를 삭제하여야 함. Client 단에서 Name을 모두 뿌려준 후에 Response.End 를 기술함.
	Else    
		rs0.MoveFirst
		iCnt =0
		
		For iCnt= 0 to 81
		
			arrRsVal(iCnt) = rs0(iCnt)
						
		Next

		rs0.Close
		Set rs0 = Nothing
		Set lgADF   = Nothing
		
	End If
	
	
    
End Sub

%>

<Script Language=VBScript>
	With parent.frm1
		Dim strDt
	
		' Tab 1 : L/C 금액정보 
		
		.txtLCNo.value = "<%=ConvSPChars(arrRsVal(0))%>"
		.txtLCDocNo.value = "<%=ConvSPChars(arrRsVal(1))%>"
		.txtLCAmendSeq.value = "<%=ConvSPChars(arrRsVal(2))%>"
		.txtAdvNo.value = "<%=ConvSPChars(arrRsVal(3))%>"
		.txtLCType.value = "<%=ConvSPChars(arrRsVal(4))%>"
		.txtLCTypeNm.value = "<%=ConvSPChars(arrRsVal(5))%>"
		.txtAdvDt.text = "<%=UNIDateClientFormat(arrRsVal(6))%>"
		.txtAdvBank.value = "<%=ConvSPChars(arrRsVal(7))%>"
		.txtAdvBankNm.value = "<%=ConvSPChars(arrRsVal(8))%>"
		.txtExpireDt.text = "<%=UNIDateClientFormat(arrRsVal(9))%>"
		.txtOpenBank.value = "<%=ConvSPChars(arrRsVal(10))%>"
		.txtOpenBankNm.value = "<%=ConvSPChars(arrRsVal(11))%>"
		.txtOpenDt.text = "<%=UNIDateClientFormat(arrRsVal(12))%>"
		.txtCurrency.value = "<%=ConvSPChars(arrRsVal(13))%>"
		.txtDocAmt.text = "<%=UNINumClientFormat(arrRsVal(14), ggAmtOfMoney.DecPoint, 0)%>"
		.txtLocAmt.text = "<%=UNINumClientFormat(arrRsVal(15), ggAmtOfMoney.DecPoint, 0)%>"
		.txtXchRate.text = "<%=UNINumClientFormat(arrRsVal(16), ggExchRate.DecPoint, 0)%>"
		.txtApplicant.value = "<%=ConvSPChars(arrRsVal(17))%>"
		.txtApplicantNm.value = "<%=ConvSPChars(arrRsVal(18))%>"
		.txtBeneficiary.value = "<%=ConvSPChars(arrRsVal(19))%>"
		.txtBeneficiaryNm.value = "<%=ConvSPChars(arrRsVal(20))%>"
		.txttolerance.value = "<%=ConvSPChars(arrRsVal(21))%>"
		.txtIncoterms.value = "<%=ConvSPChars(arrRsVal(22))%>"
		.txtIncotermsNm.value = "<%=ConvSPChars(arrRsVal(23))%>"
		.txtSalesGroup.value = "<%=ConvSPChars(arrRsVal(24))%>"
		.txtSalesGroupNm.value = "<%=ConvSPChars(arrRsVal(25))%>"
		.txtPayTerms.value = "<%=ConvSPChars(arrRsVal(26))%>"
		.txtPayTermsNm.value = "<%=ConvSPChars(arrRsVal(27))%>"
		.txtPayDur.value = "<%=ConvSPChars(arrRsVal(28))%>"
		
		'Tab 2 : L/C 선적정보 
		
		.txtLatestShipDt.text = "<%=UNIDateClientFormat(arrRsVal(29))%>"
		.txtTransport.value = "<%=ConvSPChars(arrRsVal(30))%>"
		.txtTransportNm.value = "<%=ConvSPChars(arrRsVal(31))%>"
		
		If "<%=arrRsVal(32)%>" = "Y" Then
			.rdoTranshipment1.Checked = True
		ElseIf "<%=arrRsVal(32)%>" = "N" Then
			.rdoTranshipment2.Checked = True
		End If
		
		If "<%=arrRsVal(33)%>" = "Y" Then
			.rdoPartailShip1.Checked = True
		ElseIf "<%=arrRsVal(33)%>" = "N" Then
			.rdoPartailShip2.Checked = True
		End If
		
		.txtLoadingPort.value = "<%=ConvSPChars(arrRsVal(34))%>"
		.txtLoadingPortNm.value = "<%=ConvSPChars(arrRsVal(35))%>"
		.txtDischgePort.value = "<%=ConvSPChars(arrRsVal(36))%>"
		.txtDischgePortNm.value = "<%=ConvSPChars(arrRsVal(37))%>"
		.txtDeliveryPlce.value = "<%=ConvSPChars(arrRsVal(38))%>"
		
		'Tab 3 : 구비서류 
		
		.txtFileDt.value = "<%=ConvSPChars(arrRsVal(39))%>"
		.txtFileDtTxt.value = "<%=ConvSPChars(arrRsVal(40))%>"
		.txtInvCnt.value = "<%=ConvSPChars(arrRsVal(41))%>"
		.txtPackList.value = "<%=ConvSPChars(arrRsVal(42))%>"
		
		If "<%=arrRsVal(43)%>" = "Y" Then
			.chkCertOriginFlg.Checked = True
		ElseIf "<%=arrRsVal(43)%>" = "N" Then
			.chkCertOriginFlg.Checked = False
		End If

		
		If "<%=arrRsVal(44)%>" = "Y" Then
			.rdoBLAwFlg1.Checked = True
		ElseIf "<%=arrRsVal(44)%>" = "N" Then
			.rdoBLAwFlg2.Checked = True
		End If
		
		.txtFreight.value = "<%=ConvSPChars(arrRsVal(45))%>"
		.txtFreightNm.value = "<%=ConvSPChars(arrRsVal(46))%>"
		.txtNotifyParty.value = "<%=ConvSPChars(arrRsVal(47))%>"
		.txtNotifyPartyNm.value = "<%=ConvSPChars(arrRsVal(48))%>"
		.txtConsignee.value = "<%=ConvSPChars(arrRsVal(49))%>"
		.txtInsurPolicy.value = "<%=ConvSPChars( arrRsVal(50))%>"
		.txtDoc1.value = "<%=ConvSPChars(arrRsVal(51))%>"
		.txtDoc2.value = "<%=ConvSPChars(arrRsVal(52))%>"
		.txtDoc3.value = "<%=ConvSPChars(arrRsVal(53))%>"
		.txtDoc4.value = "<%=ConvSPChars(arrRsVal(54))%>"
		.txtDoc5.value = "<%=ConvSPChars(arrRsVal(55))%>"
		
		'Tab 4 : 은행 및 기타 
		
		.txtPayBank.value = "<%=ConvSPChars(arrRsVal(56))%>"
		.txtPayBankNm.value = "<%=ConvSPChars(arrRsVal(57))%>"
		.txtRenegoBank.value = "<%=ConvSPChars(arrRsVal(58))%>"
		.txtRenegoBankNm.value = "<%=ConvSPChars(arrRsVal(59))%>"
		.txtConfirmBank.value = "<%=ConvSPChars(arrRsVal(60))%>"
		.txtConfirmBankNm.value = "<%=ConvSPChars(arrRsVal(61))%>"
		.txtBankTxt.value = "<%=ConvSPChars(arrRsVal(62))%>"
		
		If "<%=arrRsVal(63)%>" = "Y" Then
			.rdoTransfer1.Checked = True
		ElseIf "<%=arrRsVal(63)%>" = "N" Then
			.rdoTransfer2.Checked = True
		End If
		
		.txtCreditCore.value = "<%=ConvSPChars(arrRsVal(64))%>"
		.txtCreditCoreNm.value = "<%=ConvSPChars(arrRsVal(65))%>"
		
		If "<%=arrRsVal(66)%>" = "AP" Then
			.rdoChargeCd1.Checked = True
		ElseIf "<%=arrRsVal(66)%>" = "BE" Then
			.rdoChargeCd2.Checked = True
		End If
		
		.txtChargeTxt.value = "<%=ConvSPChars(arrRsVal(67))%>"
		.txtPaymentTxt.value = "<%=ConvSPChars(arrRsVal(68))%>"
		.txtShipment.value = "<%=ConvSPChars(arrRsVal(69))%>"
		.txtPreAdvRef.value = "<%=ConvSPChars(arrRsVal(70))%>"
		.txtTransportComp.value = "<%=ConvSPChars(arrRsVal(71))%>"
		.txtOrigin.value = "<%=ConvSPChars(arrRsVal(72))%>"
		.txtOriginNm.value = "<%=ConvSPChars(arrRsVal(73))%>"
		.txtOriginCntry.value = "<%=ConvSPChars(arrRsVal(74))%>"
		.txtOriginCntryNm.value = "<%=ConvSPChars(arrRsVal(75))%>"
		.txtAgent.value = "<%=ConvSPChars(arrRsVal(76))%>"
		.txtAgentNm.value = "<%=ConvSPChars(arrRsVal(77))%>"
		.txtManufacturer.value = "<%=ConvSPChars(arrRsVal(78))%>"
		.txtManufacturerNm.value = "<%=ConvSPChars(arrRsVal(79))%>"
		.txtRemark.value = "<%=ConvSPChars(arrRsVal(80))%>"

		strDt = "<%=UNIDateClientFormat(arrRsVal(81))%>"

		If strDt <> "1899-12-30" Then
			.txtAmendDt.value = strDt
		End If

		Call parent.DbQueryOk()														'☜: 조회가 성공 

		.txtHLCNo.value = "<%=ConvSPChars(Request("txtLCNo"))%>"
	End With
</Script>

<%
	Response.End																'☜: Process End
%>


















	
	
		
