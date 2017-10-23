<%
'********************************************************************************************************
'*  1. Module Name          : Sales & Distribution														*
'*  2. Function Name        :																			*
'*  3. Program ID           : S5113RB9																	*
'*  4. Program Name         : B/L 상세정보																*
'*  5. Program Desc         : 수출 B/L Query Transaction 처리용 ASP										*
'*  7. Modified date(First) : 2000/04/21																*
'*  8. Modified date(Last)  : 2002/08/12																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Ahn TaeHee																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/18 : 화면 design												*
'*							  2. 2000/04/18 : Coding Start												*
'*							  2. 2002/08/12 : Ado														*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "S","NOCOOKIE","PB") 
Call LoadBNumericFormatB("Q","S","NOCOOKIE","PB") %>

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%																				'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

	Dim strMode
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message

    strMode      = Request("txtMode")                                               '☜: Read Operation Mode (CRUD)
    Select Case strMode
        Case CStr(UID_M0001)                                                        '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)
             Call SubBizSave()
        Case CStr(UID_M0003)                                                        '☜: Delete
             Call SubBizDelete()
    End Select
'============================================================================================================
Sub SubBizQuery()
	
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear 
    Call HideStatusWnd
    Dim lgCurrency
    Dim lgArrGlFlag   
    Dim lgStrGlFlag
    Err.Clear                                                               '☜: Protect system from crashing
    
	Call SubOpenDB(lgObjConn)
	call SubMakeSQLStatements

	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                 'If data not exists
	    lgObjRs.Close
	    lgObjConn.Close
	    Set lgObjRs = Nothing
	    Set lgObjConn = Nothing
	    If Request("txtPrevNext") <> "Q" OR Request("txtPrevNext") <> "P" Then
			'B/L정보가 없습니다.
		    Call DisplayMsgBox("205300", vbInformation, "", "", I_MKSCRIPT)             '☜: No data is found. 
		End If
	Exit sub
	End If
	
	lgCurrency = ConvSPChars(lgObjRs("Cur"))	
	'-----------------------
	'Result data display area
	'----------------------- 

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent.frm1"           & vbCr


	' Tab 1: 선적정보 1

	Response.Write ".txtCurrency.value	    = """ & lgCurrency         & """" & vbCr
	Response.Write " parent.CurFormatNumericOCX " & vbCr 		

    Response.Write ".txtSONo.value		   = """ & ConvSPChars(Trim(lgObjRs("so_no")))              & """" & vbCr
	Response.Write ".txtLCDocNo.value	   = """ & ConvSPChars(lgObjRs("lc_doc_no"))                & """" & vbCr

	If Trim(ConvSPChars(lgObjRs("lc_doc_no"))) = ""Then
		Response.Write".txtLCAmendSeq.value = """"" & vbCr
	Else
		Response.Write".txtLCAmendSeq.value = """   & lgObjRs("lc_amend_seq")          & """" & vbCr
	End If
	Response.Write ".txtBLIssueDt.text		= """ & UNIDateClientFormat(lgObjRs("bl_issue_dt"))   & """" & vbCr		

	Response.Write ".txtDocAmt.Text			= """ & UNINumClientFormatByCurrency(lgObjRs("bill_amt"), lgCurrency, ggAmtOfMoneyNo)   & """" & vbCr

	Response.Write ".txtXchRate.text		= """ & UNINumClientFormat(lgObjRs("Xchg_rate"), ggExchRate.DecPoint, 0)   & """" & vbCr

	Response.Write ".txtLocAmt.Text			= """ & UniConvNumberDBToCompany(lgObjRs("bill_amt_loc"), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)   & """" & vbCr

	Response.Write ".txtTransport.value		= """ & ConvSPChars(lgObjRs("transport"))   & """" & vbCr
	Response.Write ".txtTransportNm.value	= """ & ConvSPChars(lgObjRs("transport_nm"))   & """" & vbCr
	Response.Write ".txtApplicant.value		= """ & ConvSPChars(lgObjRs("applicant"))   & """" & vbCr
	Response.Write ".txtApplicantNm.value	= """ & ConvSPChars(lgObjRs("applicant_nm"))   & """" & vbCr
	Response.Write ".txtLoadingPort.value	= """ & ConvSPChars(lgObjRs("loading_port"))   & """" & vbCr
	Response.Write ".txtLoadingPortNm.value = """ & ConvSPChars(lgObjRs("loading_port_nm"))   & """" & vbCr
	Response.Write ".txtIncoterms.value		= """ & ConvSPChars(lgObjRs("incoterms"))   & """" & vbCr
	Response.Write ".txtIncotermsNm.value	= """ & ConvSPChars(lgObjRs("incoterms_nm"))   & """" & vbCr
	Response.Write ".txtDischgePort.value	= """ & ConvSPChars(lgObjRs("dischge_port"))   & """" & vbCr
	Response.Write ".txtDischgePortNm.value = """ & ConvSPChars(lgObjRs("dischge_port_nm"))   & """" & vbCr
	Response.Write ".txtSalesGroup.value	= """ & ConvSPChars(lgObjRs("sales_grp"))   & """" & vbCr
	Response.Write ".txtSalesGroupNm.value	= """ & ConvSPChars(lgObjRs("sales_grp_nm"))   & """" & vbCr
		
	Response.Write ".txtLoadingDt.text		= """ & UNIDateClientFormat(lgObjRs("loading_dt"))   & """" & vbCr
	
	Response.Write ".txtBeneficiary.value	= """ & ConvSPChars(lgObjRs("beneficiary"))   & """" & vbCr
	Response.Write ".txtBeneficiaryNm.value = """ & ConvSPChars(lgObjRs("beneficiary_nm"))   & """" & vbCr
	Response.Write ".txtFreight.value		= """ & ConvSPChars(lgObjRs("freight"))   & """" & vbCr
	Response.Write ".txtFreightNm.value		= """ & ConvSPChars(lgObjRs("freight_nm"))   & """" & vbCr
	Response.Write ".txtBLIssueCnt.text		= """ & ConvSPChars(lgObjRs("bl_issue_cnt"))   & """" & vbCr
	Response.Write ".txtBLIssuePlce.value	= """ & ConvSPChars(lgObjRs("bl_issue_plce"))   & """" & vbCr
		
	' Tab 2 : 선적정보 2
	
	Response.Write ".txtAgent.value				= """ & ConvSPChars(lgObjRs("agent"))   & """" & vbCr
	Response.Write ".txtAgentNm.value			= """ & ConvSPChars(lgObjRs("agent_nm"))   & """" & vbCr
	Response.Write ".txtManufacturer.value		= """ & ConvSPChars(lgObjRs("manufacturer"))   & """" & vbCr
	Response.Write ".txtManufacturerNm.value	= """ & ConvSPChars(lgObjRs("manufacturer_nm"))   & """" & vbCr
	Response.Write ".txtVesselNm.value			= """ & ConvSPChars(lgObjRs("vessel_nm"))   & """" & vbCr
	Response.Write ".txtVoyageNo.value			= """ & ConvSPChars(lgObjRs("voyage_no"))   & """" & vbCr
	Response.Write ".txtForwarder.value			= """ & ConvSPChars(lgObjRs("forwarder"))   & """" & vbCr
	Response.Write ".txtForwarderNm.value		= """ & ConvSPChars(lgObjRs("forwarder_nm"))   & """" & vbCr
	Response.Write ".txtVesselCntry.value		= """ & ConvSPChars(lgObjRs("vessel_cntry"))   & """" & vbCr
	Response.Write ".txtVesselCntryNm.value		= """ & ConvSPChars(lgObjRs("vessel_cntry_nm"))   & """" & vbCr
	Response.Write ".txtReceiptPlce.value		= """ & ConvSPChars(lgObjRs("receipt_plce"))   & """" & vbCr
	Response.Write ".txtDeliveryPlce.value		= """ & ConvSPChars(lgObjRs("delivery_plce"))   & """" & vbCr
	Response.Write ".txtFinalDest.value			= """ & ConvSPChars(lgObjRs("final_dest"))   & """" & vbCr
	Response.Write ".txtDischgeDt.text			= """ & UNIDateClientFormat(lgObjRs("dischge_plan_dt"))   & """" & vbCr
	Response.Write ".txtTranshipCntry.value		= """ & ConvSPChars(lgObjRs("tranship_cntry"))   & """" & vbCr
	Response.Write ".txtTranshipCntryNm.value	= """ & ConvSPChars(lgObjRs("tranship_cntry_nm"))   & """" & vbCr
	Response.Write ".txtTranshipDt.text			= """ & UNIDateClientFormat(lgObjRs("tranship_dt"))   & """" & vbCr
	Response.Write ".txtPackingType.value		= """ & ConvSPChars(lgObjRs("packing_type"))   & """" & vbCr
	Response.Write ".txtPackingTypeNm.value		= """ & ConvSPChars(lgObjRs("packing_type_nm"))   & """" & vbCr
	Response.Write ".txtTotPackingCnt.text		= """ & ConvSPChars(lgObjRs("tot_packing_cnt"))   & """" & vbCr
	Response.Write ".txtPackingTxt.value		= """ & ConvSPChars(lgObjRs("packing_txt"))   & """" & vbCr
	Response.Write ".txtContainerCnt.text		= """ & ConvSPChars(lgObjRs("container_cnt"))   & """" & vbCr
	Response.Write ".txtGrossWeight.text		= """ & UNINumClientFormat(lgObjRs("gross_weight"), ggQty.DecPoint, 0)   & """" & vbCr
	Response.Write ".txtWeightUnit.value		= """ & ConvSPChars(lgObjRs("weight_unit"))   & """" & vbCr
	Response.Write ".txtGrossVolumn.text		= """ & UNINumClientFormat(lgObjRs("gross_volumn"), ggQty.DecPoint, 0)   & """" & vbCr
	Response.Write ".txtVolumnUnit.value		= """ & ConvSPChars(lgObjRs("volumn_unit"))   & """" & vbCr
	Response.Write ".txtOrigin.value			= """ & ConvSPChars(lgObjRs("origin"))   & """" & vbCr
	Response.Write ".txtOriginNm.value			= """ & ConvSPChars(lgObjRs("origin_nm"))   & """" & vbCr
	Response.Write ".txtOriginCntry.value		= """ & ConvSPChars(lgObjRs("origin_cntry"))   & """" & vbCr
	Response.Write ".txtOriginCntryNm.value		= """ & ConvSPChars(lgObjRs("origin_cntry_nm"))   & """" & vbCr
	Response.Write ".txtFreightPlce.value		= """ & ConvSPChars(lgObjRs("freight_plce"))   & """" & vbCr
		
		'Tab 3 : 매출정보 

	Response.Write ".txtTaxBizArea.value		= """ & ConvSPChars(lgObjRs("tax_biz_area"))   & """" & vbCr
	Response.Write ".txtTaxBizAreaNm.value		= """ & ConvSPChars(lgObjRs("tax_biz_area_nm"))   & """" & vbCr
	Response.Write ".txtBillType.value			= """ & ConvSPChars(lgObjRs("bill_type"))   & """" & vbCr
	Response.Write ".txtBillTypeNm.value		= """ & ConvSPChars(lgObjRs("bill_type_nm"))   & """" & vbCr
	Response.Write ".txtPayer.value				= """ & ConvSPChars(lgObjRs("payer"))   & """" & vbCr
	Response.Write ".txtPayerNm.value			= """ & ConvSPChars(lgObjRs("payer_nm"))   & """" & vbCr
	Response.Write ".txtBilltoParty.value		= """ & ConvSPChars(lgObjRs("bill_to_party"))   & """" & vbCr
	Response.Write ".txtBilltoPartyNm.value		= """ & ConvSPChars(lgObjRs("bill_to_party_nm"))   & """" & vbCr
	Response.Write ".txtToSalesGroup.value		= """ & ConvSPChars(lgObjRs("to_sales_grp"))   & """" & vbCr 
	Response.Write ".txtToSalesGroupNm.value	= """ & ConvSPChars(lgObjRs("to_sales_grp_nm"))   & """" & vbCr 
	Response.Write ".txtCurrency1.value			= """ & lgCurrency   & """" & vbCr
	Response.Write ".txtDocAmt1.Text			= """ & UNINumClientFormatByCurrency(lgObjRs("bill_amt"), lgCurrency, ggAmtOfMoneyNo)   & """" & vbCr
		
	Response.Write ".txtPayDt.text				= """ & UNIDateClientFormat(lgObjRs("income_plan_dt"))   & """" & vbCr

	Response.Write ".txtLocAmt1.Text			= """ & UniConvNumberDBToCompany(lgObjRs("bill_amt_loc"), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)   & """" & vbCr

	Response.Write ".txtPayType.value			= """ & ConvSPChars(lgObjRs("pay_type"))   & """" & vbCr
	Response.Write ".txtPayTypeNm.value			= """ & ConvSPChars(lgObjRs("pay_type_nm"))   & """" & vbCr
	Response.Write ".txtPayTerms.value			= """ & ConvSPChars(lgObjRs("pay_meth"))   & """" & vbCr
	Response.Write ".txtPayTermsNm.value		= """ & ConvSPChars(lgObjRs("pay_meth_nm"))   & """" & vbCr
	Response.Write ".txtMoney.Text				= """ & UNINumClientFormatByCurrency(lgObjRs("collect_amt"), lgCurrency, ggAmtOfMoneyNo)   & """" & vbCr
		
	Response.Write ".txtPayDur.value			= """ & ConvSPChars(lgObjRs("pay_dur"))   & """" & vbCr
	Response.Write ".txtPayTermstxt.value		= """ & ConvSPChars(lgObjRs("pay_terms_txt"))   & """" & vbCr
	Response.Write ".txtRemark.value			= """ & ConvSPChars(lgObjRs("remark"))   & """" & vbCr		
	Response.Write ".txtRefFlg.value			= """ & ConvSPChars(lgObjRs("ref_flag"))   & """" & vbCr

    Response.Write "End With"          & vbCr
    Response.Write "</Script>"         & vbCr
	
	lgObjRs.Close
	lgObjConn.Close
	Set lgObjRs = Nothing
	Set lgObjConn = Nothing
										
End Sub	
%>

<SCRIPT LANGUAGE=VBSCRIPT RUNAT=SERVER>
'============================================================================================================
Sub SubMakeSQLStatements()
	Dim strSelectList, strFromList
	
	strSelectList = "SELECT * "
	strFromList = " FROM dbo.ufn_s_GetBLInfo ( " & FilterVar(Request("txtBlNo"), "''", "S") & ", " & FilterVar("N", "''", "S") & " , " & FilterVar("Y", "''", "S") & " , " & FilterVar("Y", "''", "S") & " , " & FilterVar("Y", "''", "S") & " ,  " & FilterVar(Request("txtPrevNext"), "''", "S") & ")"
	lgstrsql = strSelectList & strFromList
	
	'call svrmsgbox(lgstrsql, vbinformation, i_mkscript)
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
</SCRIPT>
