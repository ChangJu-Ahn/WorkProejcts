<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출채권관리 
'*  3. Program ID           : S5112MB2
'*  4. Program Name         : 예외매출채권내역등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : PS7G128.cSListBillDtlSvr,PS7G121.cSBillDtlSvr,PS7G115.cSPostOpenArSvr,PB3C104.cBLkUpItem
'*  7. Modified date(First) : 2002/11/14
'*  8. Modified date(Last)  : 2003/06/20
'*  9. Modifier (First)     : AHN TAE HEE
'* 10. Modifier (Last)      : Hwang Seongbae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/19 : 3rd 화면 Layout & ASP Coding
'*                            -2000/08/11 : 4th 화면 Layout
'*                            -2001/12/18 : Date 표준적용 
'*                            -2001/12/26 : VAT 개별통합 추가 
'*							  -2002/11/14 : UI성능 적용	
'**********************************************************************************************

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%Call LoadBasisGlobalInf()
Call loadInfTB19029B("I", "*","NOCOOKIE","MB") 
Call LoadBNumericFormatB("I","*","NOCOOKIE","MB") %>

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%													

On Error Resume Next														

Call HideStatusWnd
'for 구주Tax
Dim pvCB 

Dim iPS7G121												'☆ : 매출내역등록입력/수정/삭제용 ComProxy Dll 사용 변수 
Dim iPS7G128												'☆ : 매출내역등록 조회용 ComProxy Dll 사용 변수 
Dim iPS7G115											    '☆ : 매출내역확정용 ComProxy Dll 사용 변수					
Dim iPB3C104												'☆ : Item별 HS코드 조회용 ComProxy Dll 사용 변수 
		
Dim strMode		
Dim iStrNextKey											' 다음 값 
Dim lgStrPrevKey										' 이전 값 
Dim LngMaxRow											' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount															
Dim lgCurrency
Dim lgArrGlFlag
Dim lgStrGlFlag
Dim lgStrPostFlag
Dim lgStrGlNo

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

Select Case strMode

Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 
	Dim iStrSvrData

    Err.Clear                                                                '☜: Protect system from crashing

	If Request("txtHQuery") = "T" Then
		'-----------------------
		' 매출헤더를 읽어온다.
		'-----------------------
		Call SubOpenDB(lgObjConn)
		call SubMakeSQLStatements

		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                 'If data not exists
		    lgObjRs.Close
		    lgObjConn.Close
		    Set lgObjRs = Nothing
		    Set lgObjConn = Nothing
			'매출채권정보가 없습니다.
			Call DisplayMsgBox("205100", vbInformation, "", "", I_MKSCRIPT)             '☜: No data is found. 

			Response.Write "<Script Language=vbscript>" & vbcr
			Response.Write "parent.SetDefaultVal" & vbcr
			Response.Write "Call parent.SetToolbar(""11000000000011"")" & vbcr
			Response.Write "</Script>" & vbcr
		    Response.End
		End If
%>
		<Script Language=vbscript>
			With parent.frm1
				'-----------------------
				' 매출헤더의 내용을 표시한다.
				'----------------------- 
			
				'##### Rounding Logic #####
				'항상 거래화폐가 우선 
				<%
				lgCurrency = ConvSPChars(lgObjRs("Cur"))
				%>

				.txtCurrency.value			= "<%=lgCurrency%>"
				parent.CurFormatNumericOCX
				'##########################

				.txtSoldtoParty.value		= "<%=ConvSPChars(lgObjRs("Sold_to_Party"))%>"
				.txtSoldtoPartyNm.value		= "<%=ConvSPChars(lgObjRs("Sold_to_Party_Nm"))%>"
				.txtPayTermsCd.value		= "<%=ConvSPChars(lgObjRs("Pay_Meth"))%>"
				.txtPayTermsNm.value		= "<%=ConvSPChars(lgObjRs("Pay_Meth_Nm"))%>"
				.txtSalesGrpCd.value		= "<%=ConvSPChars(lgObjRs("sales_grp"))%>"
				.txtSalesGrpNm.value		= "<%=ConvSPChars(lgObjRs("sales_grp_nm"))%>"
				
				.txtOriginBillAmt.Text		= "<%=UNIConvNumDBToCompanyByCurrency(lgObjRs("bill_amt"),lgCurrency,ggAmtOfMoneyNo, "X" , "X")%>"
				.txtVatAmt.Text				= "<%=UNIConvNumDBToCompanyByCurrency(lgObjRs("vat_amt"),lgCurrency,ggAmtOfMoneyNo, gTaxRndPolicyNo, "X")%>"
				.txtVatRate.Text			= "<%=UNIConvNumDBToCompanyByCurrency(lgObjRs("vat_rate"), gCurrency, ggExchRateNo, "X" , "X")%>"
				.txtVatType.value			= "<%=ConvSPChars(lgObjRs("vat_type"))%>"
				.txtVatTypeNm.value			= "<%=ConvSPChars(lgObjRs("vat_type_nm"))%>"
				.txtHBillType.value			= "<%=ConvSPChars(lgObjRs("bill_type"))%>"
				.txtHBillTypeNm.value		= "<%=ConvSPChars(lgObjRs("bill_type_nm"))%>"
				.txtHRefFlag.value			= "<%=Trim(ConvSPChars(lgObjRs("ref_flag")))%>"
				
				<%'품목별 vat유형 관리 여부 Header에 vat유형이 등록되어 있는 경우 'N'%>
				If "<%=Trim(ConvSPChars(lgObjRs("vat_type")))%>" <> "" Then
					Parent.lgStrVatFlag = "N"
				Else
					Parent.lgStrVatFlag = "Y"
				End If
				
				<% '매출일 %>
				.txtHBillDt.Value			= "<%=UNIDateClientFormat(lgObjRs("bill_dt"))%>"				

				<% '수주진행상태 %>
				.txtSts.value				= "<%=lgObjRs("sts")%>"

				'VAT적용기준 
				If "<%=Trim(lgObjRs("vat_calc_type"))%>" = "1" Then
					.rdoVatCalcType1.checked = True
				Else
					.rdoVatCalcType2.checked = True
				End If

				'VAT 포함구분 
				If "<%=Trim(lgObjRs("vat_inc_flag"))%>" = "1" Then
					.rdoVatIncFlag1.checked = True
				Else
					.rdoVatIncFlag2.checked = True
				End If

				.txtHBillNo.value = "<%=ConvSPChars(lgObjRs("bill_no"))%>"
				
				'이전 매출번호 
				.txtRefBillNo.Value = "<%=Trim(ConvSPChars(lgObjRs("so_no")))%>"
				
				.txtXchgOp.value = "<%=lgObjRs("xchg_rate_op")%>"
				.txtXchgRate.Text = "<%=UNIConvNumDBToCompanyByCurrency(lgObjRs("xchg_rate"), gCurrency, ggExchRateNo, "X" , "X")%>"
								
				.txtHPostflag.value = "<%=lgObjRs("post_flag")%>"
				<%lgStrPostFlag = lgObjRs("post_flag")
				  lgStrGlNo = Trim(lgObjRs("gl_no"))
				  If lgStrPostFlag = "Y" AND Len(lgStrGlNo) Then
					lgArrGlFlag = Split(lgStrGlNo, Chr(11)) 
					lgStrGlFlag = lgArrGlFlag(0)%>
					
					If "<%=lgArrGlFlag(0)%>" = "G" Then	
						'회계전표번호 
						.txtGLNo.value	= "<%=lgArrGlFlag(1)%>"	
					ElseIf "<%=lgArrGlFlag(0)%>" = "T" Then
						'결의전표번호 
						.txtTempGLNo.value	= "<%=lgArrGlFlag(1)%>"	
					Else
						'Batch번호 
						.txtBatchNo.value	= "<%=lgArrGlFlag(1)%>"	
					End If
				<%Else%>
						.txtGLNo.value	= ""	
						.txtTempGLNo.value	= ""	
						.txtBatchNo.value	= ""	
				<% End If %>

				If "<%=lgStrPostFlag%>" = "Y" Then
					.btnPostFlag.value = "확정취소"
					If "<%=lgStrGlFlag%>" = "G" Or "<%=lgStrGlFlag%>" = "T" Then
						.btnGLView.disabled = False
					Else
						.btnGLView.disabled = True
					End If
				ELSE
					.btnPostFlag.value = "확정"
					.btnGLView.disabled = True
				End If

				<% '선수금 현황 버튼 Enable %>
				IF "<%=lgObjRs("PreRcpt_flag")%>" = "Y" Then
					.btnPreRcptView.disabled = False
				Else
					.btnPreRcptView.disabled = True
				End If

				'적립금관리여부 설정 
				<%If Trim(lgObjRs("deposit_flag")) <> "" Then %>
					parent.lgstrDepositFlag = "<%=Trim(lgObjRs("deposit_flag"))%>"
				<%Else%>
					parent.lgstrDepositFlag = "2"
				<%End If%>

				<%
				lgObjRs.Close
				lgObjConn.Close
				Set lgObjRs = Nothing
				Set lgObjConn = Nothing
				%>
				parent.SetSpreadHidden
				'-----------------------
				' Rounding Column Set
				'----------------------- 
				parent.CurFormatNumSprSheet
				
				'parent.GetDepositFlag
				
				parent.DbQueryOk														'☜: 조회가 성공 

			End With
		</Script>		
<%
	ElseIf Request("txtHQuery") = "F" Then
		lgCurrency = Request("txtCurrency")			
	End If		' End of Header Query
	
	'-----------------------
    ' 매출내역을 읽어온다.
    '-----------------------
    '--------------
	'Interface 정의 
	'--------------
    'View Name : imp_next s_bill_dtl
    Const S526_I1_bill_seq = 0
    'View Name : imp s_bill_hdr
    Const S526_I2_bill_no = 0
    'View Name : exp_next s_bill_dtl
    Const S526_E1_bill_seq = 0

    'Group Name : exp_grp
    Const S526_EG1_E1_minor_nm = 0    'View Name : exp_item b_minor
    Const S526_EG1_E2_cc_seq = 1    'View Name : exp_item s_cc_dtl
    Const S526_EG1_E3_cc_no = 2    'View Name : exp_item s_cc_hdr
    Const S526_EG1_E4_lc_seq = 3    'View Name : exp_item s_lc_dtl
    Const S526_EG1_E5_lc_no = 4    'View Name : exp_item s_lc_hdr
    Const S526_EG1_E6_bill_seq = 5    'View Name : exp_item s_bill_dtl
    Const S526_EG1_E6_bill_price = 6
    Const S526_EG1_E6_bill_amt = 7
    Const S526_EG1_E6_vat_amt = 8
    Const S526_EG1_E6_bill_qty = 9
    Const S526_EG1_E6_bill_unit = 10
    Const S526_EG1_E6_remark = 11
    Const S526_EG1_E6_item_acct = 12
    Const S526_EG1_E6_tracking_no = 13
    Const S526_EG1_E6_plant_biz_area = 14
    Const S526_EG1_E6_cost_cd = 15
    Const S526_EG1_E6_hs_no = 16
    Const S526_EG1_E6_cust_item_cd = 17
    Const S526_EG1_E6_bill_amt_loc = 18
    Const S526_EG1_E6_vat_type = 19
    Const S526_EG1_E6_vat_rate = 20
    Const S526_EG1_E6_vat_amt_loc = 21
    Const S526_EG1_E6_cust_po_no = 22
    Const S526_EG1_E6_cust_po_seq = 23
    Const S526_EG1_E6_gross_weight = 24
    Const S526_EG1_E6_net_weight = 25
    Const S526_EG1_E6_volume_size = 26
    Const S526_EG1_E6_ext1_qty = 27
    Const S526_EG1_E6_ext2_qty = 28
    Const S526_EG1_E6_ext3_qty = 29
    Const S526_EG1_E6_ext1_amt = 30
    Const S526_EG1_E6_ext2_amt = 31
    Const S526_EG1_E6_ext3_amt = 32
    Const S526_EG1_E6_ext1_cd = 33
    Const S526_EG1_E6_ext2_cd = 34
    Const S526_EG1_E6_ext3_cd = 35
    Const S526_EG1_E6_vat_inc_flag = 36
    Const S526_EG1_E6_deposit_price = 37
    Const S526_EG1_E6_deposit_amt = 38
    Const S526_EG1_E6_ret_item_flag = 39
    Const S526_EG1_E9_plant_cd = 42    'View Name : exp_item b_plant
    Const S526_EG1_E10_item_cd = 43    'View Name : exp_item b_item
    Const S526_EG1_E10_item_nm = 44
    Const S526_EG1_E10_spec = 45
    
    Const C_SHEETMAXROWS_D  = 100
    
    '--------
	'View선언 
	'--------    
    Dim I2_s_bill_hdr
    Dim I1_s_bill_dtl
    Dim EG1_exp_grp
    Dim E1_s_bill_dtl 
    
     '---------------------------------------
    'Data manipulate  area(import view match)
    '----------------------------------------
    redim I2_s_bill_hdr(0)
    
    I2_s_bill_hdr(S526_I2_bill_no) = Trim(Request("txtConBillNo"))
    
    redim I1_s_bill_dtl(0)
    
    If Trim(Request("lgStrPrevKey")) = "" then
		I1_s_bill_dtl(S526_I1_bill_seq) = 0
    Else
		I1_s_bill_dtl(S526_I1_bill_seq) = cdbl(Request("lgStrPrevKey"))
	End if	
    
    Set iPS7G128 = Server.CreateObject("PS7G128.cSListBillDtlSvr")    
  
    If CheckSYSTEMError(Err,True) = True Then
		Response.End		
    End If
	
	Call iPS7G128.S_LIST_BILL_DTL_SVR(gStrGlobalCollection , C_SHEETMAXROWS_D , _
	                                 I1_s_bill_dtl ,I2_s_bill_hdr, EG1_exp_grp, E1_s_bill_dtl)

	If CheckSYSTEMError(Err,True) = True Then
       Set iPS7G128 = Nothing		                                                 '☜: Unload Comproxy DLL
		Response.Write "<Script language=vbs> " & vbCr   
		Response.Write " Parent.frm1.txtConBillNo.Focus" & vbCr   
		Response.Write "</Script> " & vbCr          
       Response.End 
    End If   

    Set iPS7G128 = Nothing	
   
	'----------------------------
	' 매출내역의 내용을 표시한다.
	'---------------------------- 
	Dim iLngSheetMaxRows
	Dim iArrCols, iArrRows
	
	' Set Next key
	If Ubound(EG1_exp_grp,1) = C_SHEETMAXROWS_D Then
		'매출번호 
		iStrNextKey = E1_s_bill_dtl(S526_E1_bill_seq)
		iLngSheetMaxRows  = C_SHEETMAXROWS_D - 1
	Else
		iStrNextKey = ""
		iLngSheetMaxRows = Ubound(EG1_exp_grp,1)
	End If

	ReDim iArrCols(36)						' Column 수 
	Redim iArrRows(iLngSheetMaxRows)		' 조회된 Row 수만큼 배열 재정의 

	LngMaxRow = CInt(Request("txtMaxRows")) + 1

	' 변경되지 않는 값 설정 
	iArrCols(0)  = ""		' Row Header
   	iArrCols(2)  = ""		' 공장팝업 
   	iArrCols(4)  = ""		' 품목팝업 
   	iArrCols(8)  = ""		' 단위팝업 
   	iArrCols(14) = ""		' VAT유형 팝업 
	iArrCols(26) = ""		' 이전매출번호 
	iArrCols(27) = "0"		' 이전매출순번 
	iArrCols(35) = ""		' Tracking 번호팝업 
	
	For LngRow = 0 To iLngSheetMaxRows
   		iArrCols(1) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E9_plant_cd))		' 공장코드 
   		iArrCols(3) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E10_item_cd))		' 품목코드 
   		iArrCols(5) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E10_item_nm))		' 품목명 
   		iArrCols(6) = UNINumClientFormat(EG1_exp_grp(LngRow, S526_EG1_E6_bill_qty), ggQty.DecPoint, 0)	' 수량 
   		iArrCols(7) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E6_bill_unit))	' 단위 
  		iArrCols(10) = EG1_exp_grp(LngRow,S526_EG1_E6_vat_inc_flag)				' Vat포함구분 
		if iArrCols(10) = "1" Then
   			iArrCols(9) = "별도"	
		Else	
   			iArrCols(9) = "포함"
		End if

   		iArrCols(11) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,S526_EG1_E6_bill_price),lgCurrency,ggUnitCostNo, "X" , "X")	' 단가 
   		iArrCols(12) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,S526_EG1_E6_bill_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")	' 금액 
   		iArrCols(13) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E6_vat_type))		' VAT유형 
   		iArrCols(15) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E1_minor_nm))		' VAT유형명 
   		iArrCols(16) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,S526_EG1_E6_vat_rate), gCurrency, ggExchRateNo, "X" , "X")	' VAT율 
   		iArrCols(17) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,S526_EG1_E6_vat_amt),lgCurrency,ggAmtOfMoneyNo, gTaxRndPolicyNo  , "X")		' VAT 금액 
   		iArrCols(18) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,S526_EG1_E6_bill_amt_loc),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo  , "X")	' 자국금액 
   		iArrCols(19) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,S526_EG1_E6_vat_amt_loc),gCurrency,ggAmtOfMoneyNo, gTaxRndPolicyNo  , "X")	' VAT 자국금액 
   		iArrCols(20) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,S526_EG1_E6_deposit_price ),lgCurrency,ggUnitCostNo, "X" , "X")				' 적립단가 
   		iArrCols(21) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,S526_EG1_E6_deposit_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")				' 적립금액 
   		iArrCols(22) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E6_remark))	' 비고 
   		iArrCols(23) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E10_spec))	' 규격 
   		iArrCols(24) = UNINumClientFormat(EG1_exp_grp(LngRow,S526_EG1_E6_bill_seq), 0, 0)	' 매출순번 
		iArrCols(25) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E6_ret_item_flag))			' 반품여부 
		iArrCols(28) = iArrCols(12)		' 변경전금액 
		iArrCols(29) = iArrCols(10)		' 변경전 VAT 포함여부 
		iArrCols(30) = iArrCols(17)		' 변경전부가세액 
		iArrCols(31) = iArrCols(12)		' 변경전금액 
		iArrCols(32) = iArrCols(10)		' 변경전 VAT 포함여부 
		iArrCols(33) = iArrCols(17)		' 변경전부가세액 
		iArrCols(34) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E6_tracking_no))		'Tracking No 		' tracking no
		iArrCols(36) = LngMaxRow + LngRow
		
   		iArrRows(LngRow) = Join(iArrCols, gColSep)
	Next
        
	Response.Write "<Script language=vbs> " & vbCr   
	Response.Write "With parent " & vbCr   
    Response.Write " .ggoSpread.Source = .frm1.vspdData" & vbCr
    Response.Write " .frm1.vspdData.Redraw = False  "      & vbCr      
    Response.Write " .ggoSpread.SSShowDataByClip   """ & Join(iArrRows, gColSep & gRowSep) & gColSep & gRowSep & """ ,""F""" & vbCr
    Response.Write " .lgStrPrevKey = """ & iStrNextKey  & """" & vbCr  
    Response.Write " .DbQueryOk " & vbCr   
    Response.Write " If .frm1.txtHPostFlag.Value = """ & "Y" & """ Then  " & vbCr	         
	Response.Write " .SetPostYesSpreadColor(" & LngMaxRow & ")  " & vbCr	
	Response.Write " Else  " & vbCr
	Response.Write " .SetQuerySpreadColor(" & LngMaxRow & ") " & vbCr
	Response.Write " End If	 " & vbCr
	Response.Write ".frm1.vspdData.Redraw = True  "       & vbCr
	Response.Write "End With " & vbCr   
	Response.Write "</Script> " & vbCr          

Case CStr(UID_M0002)																'☜: 저장 요청을 받음 
									
    Err.Clear																		'☜: Protect system from crashing
    
    Dim iErrorPosition
    Dim iCUCount, iDCount, iIntIndex
    Dim itxtSpread, itxtHBillNo
	Dim itxtSpreadArrCount, itxtSpreadArr
    
    itxtSpread = ""
             
    iCUCount = Request.Form("txtCUSpread").Count
    iDCount  = Request.Form("txtDSpread").Count

    itxtSpreadArrCount = -1
             
    ReDim itxtSpreadArr(iCUCount + iDCount)
             
    For iIntIndex = 1 To iDCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtDSpread")(iIntIndex)
    Next
    
    For iIntIndex = 1 To iCUCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(iIntIndex)
    Next
    
    itxtSpread = Join(itxtSpreadArr,"")
    
	itxtHBillNo = Trim(Request("txtHBillNo"))
	
     pvCB = "F"			'for 구주Tax

    Set iPS7G121 = Server.CreateObject("PS7G121.cSBillDtlSvr")  
    
    If CheckSYSTEMError(Err,True) = True Then
		Response.Write "<Script language=vbs> " & vbCr   
		Response.Write "Call parent.RemovedivTextArea " & vbCr   
		Response.Write "</Script> "																				         & vbCr          
		Response.End		
    End If

'	Response.Write replace(replace(itxtSpread, chr(12), vbcrlf & vbcrlf), chr(11), vbcrlf)
'	Response.End 

	Call iPS7G121.S_MAINT_BILL_DTL_SVR(pvCB,gStrGlobalCollection ,itxtSpread, iErrorPosition, _
		                                itxtHBillNo)
    
	If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
       Set iPS7G121 = Nothing
		Response.Write "<Script language=vbs> " & vbCr   
		Response.Write "Call parent.RemovedivTextArea " & vbCr   
		If iErrorPosition > 0 Then
			Response.Write " Call Parent.SubSetErrPos(" & iErrorPosition & ")" & vbCr
		End If
		Response.Write "</Script> "	 & vbCr          
	   Response.End 
	End If

    Set iPS7G121 = Nothing	
    
    Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> " 													'☜: Row 의 상태 
				

Case CStr("PostFlag")																'☜: 확정 요청 

    Err.Clear					'☜: Protect system from crashing
    
    Dim itxtHBillNoPost

	Set iPS7G115 = Server.CreateObject("PS7G115.cSPostOpenArSvr")  
    
    If CheckSYSTEMError(Err,True) = True Then
		Response.End		
    End If
	
	itxtHBillNoPost = Trim(Request("txtHBillNo"))
	
	'for 구주Tax
    '@@@@@@@@@@@
     pvCB = "F"
    '@@@@@@@@@@@
	Call iPS7G115.S_POST_OPEN_AR_SVR(pvCB,gStrGlobalCollection ,itxtHBillNoPost)
    
	If CheckSYSTEMError(Err,True) = True Then
		Set iPS7G115 = Nothing
		Response.End		
    End If
	
	Set iPS7G115 = Nothing	
	
	Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "    

End Select

'==============================================================================
' 사용자 정의 서버 함수 
'==============================================================================
%>
<SCRIPT LANGUAGE=VBSCRIPT RUNAT=SERVER>
'============================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================
Sub SubMakeSQLStatements()
	Dim strSelectList, strFromList
	
	strSelectList = "SELECT * "
	strFromList = " FROM dbo.ufn_s_GetBillHdrInfo( " & FilterVar(Request("txtConBillNo"), "''", "S") & " , " & FilterVar("Y", "''", "S") & " , " & FilterVar("%", "''", "S") & ", " & FilterVar("Y", "''", "S") & " , " & FilterVar("Y", "''", "S") & " , " & FilterVar("Q", "''", "S") & " , " & FilterVar("N", "''", "S") & " )"
	lgstrsql = strSelectList & strFromList
	'call svrmsgbox(lgstrsql, vbinformation, i_mkscript)
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
</SCRIPT>
