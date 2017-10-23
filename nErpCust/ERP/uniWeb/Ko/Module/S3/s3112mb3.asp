<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 수주관리 
'*  3. Program ID           : S3112Mb3
'*  4. Program Name         : 수주내역Amend
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/12/04
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : son
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                          
'**********************************************************************************************


%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%													

On Error Resume Next
														
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
Call LoadBNumericFormatB("I","*","NOCOOKIE","MB")

Call HideStatusWnd

Dim I1_s_so_hdr	
Dim I2_s_so_hdr	
Const l2_s_so_no = 0
Const l2_s_cfm_flag = 1	
Redim I2_s_so_hdr(1)

Dim iLngRow	
Dim iLngMaxRow
Dim istrData
Dim iStrPrevKey
Dim iStrNextKey
Dim I1_command
Dim istrMode
Dim iPS3G152
Const C_SHEETMAXROWS_D  = 100	
Dim C_s_so_no
Dim C_s_so_seq

'--------------
' 수주내역정보 
'--------------
Dim EG1_exp_grp													
Const EG1_so_seq         = 0   '---수주순번 
Const EG1_hs_no          = 1   'HS부호 
Const EG1_so_price       = 2   '단가 
Const EG1_net_amt        = 3   '금액(수주순금액[거래화폐])
Const EG1_so_qty         = 4   '수량(수주량)
Const EG1_bonus_qty      = 5   '덤수량(할증수량[덤]) 
Const EG1_req_qty        = 6   '출고요청량 
Const EG1_req_bonus_qty  = 7
Const EG1_bill_qty       = 8   '---매출수량 
Const EG1_so_unit        = 9   '단위 
Const EG1_lc_qty         = 10
Const EG1_tol_more_rate  = 11  '과부족허용율(+)
Const EG1_tol_less_rate  = 12  '과부족허용율(-)
Const EG1_close_flag     = 13
Const EG1_so_status      = 14  '---수주진행상태 
Const EG1_remark         = 15  '비고 
Const EG1_cust_item_cd   = 16
Const EG1_gi_qty         = 17
Const EG1_gi_bonus_qty   = 18
Const EG1_pre_doc_seq    = 19   '---이전수주순번 
Const EG1_vat_amt        = 20   'VAT금액 
Const EG1_dlvy_dt        = 21   '납기일 
Const EG1_tracking_no    = 22   'Tracking No
Const EG1_so_base_qty    = 23   '---재고수량 
Const EG1_bonus_base_qty = 24   '---덤재고수량 
Const EG1_dn_seq         = 25   '이전출하순번 
Const EG1_cust_po_seq    = 26
Const EG1_bom_num        = 27
Const EG1_price_flag     = 28   '단가구분(가단가:N, 진단가:Y)
Const EG1_ctp_times      = 29   '---CTP Time
Const EG1_pur_qty        = 30
Const EG1_atp_flag       = 31
Const EG1_pre_doc_no     = 32   '---이전수주번호 
Const EG1_ext1_qty       = 33
Const EG1_ext2_qty       = 34
Const EG1_ext3_qty       = 35
Const EG1_ext1_amt       = 36
Const EG1_ext2_amt       = 37
Const EG1_ext3_amt       = 38
Const EG1_ext1_cd        = 39
Const EG1_ext2_cd        = 40
Const EG1_ext3_cd        = 41
Const EG1_net_amt_loc    = 42
Const EG1_vat_amt_loc    = 43
Const EG1_dn_no          = 44   '---이전출하번호 
Const EG1_lot_seq        = 45   'Lot seq
Const EG1_lot_no         = 46   'Lot no
Const EG1_ret_type       = 47   '반품유형(반품처리구분)
Const EG1_vat_type       = 48   'VAT유형 
Const EG1_vat_rate       = 49   'VAT율 
Const EG1_vat_inc_flag   = 50   '---VAT포함구분(1:별도, 2:포함)
Const EG1_sl_cd          = 51   '창고 
Const EG1_sl_nm          = 52   '---창고명 
Const EG1_plant_cd       = 53   '공장 
Const EG1_plant_nm       = 54   '---공장명 
Const EG1_item_cd        = 55   '품목 
Const EG1_item_nm        = 56   '품목명  
Const EG1_spec           = 57   '규격 
Const EG1_bp_cd          = 58   '납품처 
Const EG1_bp_nm          = 59   '---납품처명 

Const EG1_promise_dt      = 60   '출하요청일(출하예정일자)
Const EG1_vat_type_nm     = 61   'VAT유형명 
Const EG1_ret_type_nm     = 62   '반품유형명(반품처리구분명)
Const EG1_vat_inc_flag_nm = 63   'VAT포함구분명(1:별도, 2:포함)
Const EG1_aps_host        = 64   '---APSHost 
Const EG1_aps_port        = 65   '---APSPort
Const EG1_flag            = 66   '---CTPCheckFlag 

'--------------
' 수주헤더정보 
'--------------
Dim EG2_exp_grp 
Const EG2_so_dt           = 0
Const EG2_req_dlvy_dt     = 1
Const EG2_cfm_flag        = 2
Const EG2_price_flag      = 3
Const EG2_cur             = 4
Const EG2_net_amt         = 5
Const EG2_cust_po_no      = 6
Const EG2_deal_type       = 7
Const EG2_pay_meth        = 8
Const EG2_vat_inc_flag    = 9
Const EG2_vat_type        = 10
Const EG2_vat_rate        = 11
Const EG2_vat_amt         = 12
Const EG2_pre_doc_no      = 13
Const EG2_ret_item_flag   = 14
Const EG2_export_flag     = 15
Const EG2_so_sts          = 16
Const EG2_maint_no        = 17
Const EG2_auto_dn_flag    = 18
Const EG2_so_type         = 19
Const EG2_bp_cd2          = 20
Const EG2_bp_cd3          = 21
Const EG2_bp_nm3          = 22
Const EG2_vat_inc_flag_nm = 23

Const lsConfirm = "CONFIRM"													' 확정처리시 
istrMode = Request("txtMode")												' 현재 상태를 받음 

Select Case istrMode

Case CStr(UID_M0001)														' 현재 조회/Prev/Next 요청을 받음 

    Err.Clear     
    
    '--------------------------------------------------------------------------------------------------------
    ' 수주 HDR와 DTL를 읽어온다.
    '--------------------------------------------------------------------------------------------------------

    C_s_so_no = Trim(Request("txtConSoNo"))    
    iStrPrevKey = Trim(Request("lgStrPrevKey"))    
    If iStrPrevKey <> "" then
		C_s_so_seq = iStrPrevKey
    Else
		C_s_so_seq = 0
    End If    
	
    Set iPS3G152 = Server.CreateObject("PS3G152.cSListSoDtl")      
    
    If CheckSYSTEMError(Err,True) = True Then
       Response.End
    End If

    Call iPS3G152.S_LIST_SO_DTL_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, C_s_so_no, C_s_so_seq, _
                                    EG1_exp_grp, EG2_exp_grp)
    
    '-------------------------
    ' 수주헤더정보가 없을때.
    '-------------------------
    If cStr(Err.Description) = "B_MESSAGE" & Chr(11) & "203100" Then     
        If CheckSYSTEMError(Err,True) = True Then        
%>     
        <Script Language=vbscript>
 			parent.SetDefaultVal
 			Call parent.SetToolbar("11000000000011")
 			parent.frm1.txtConSoNo.focus
        </Script>
<%
        End If
        Response.End		                                                      
    End If    

    '----------------------------
	' 수주헤더정보를 표시한다.
	'----------------------------
%>
<Script Language=vbscript>
	With parent		
<%
		Dim lgCurrency																			'항상 거래화폐가 우선 
		lgCurrency = ConvSPChars(EG2_exp_grp(EG2_cur))
%>
		.frm1.txtCurrency.value = "<%=lgCurrency%>"
		parent.CurFormatNumericOCX	
		.frm1.txtSoldToParty.value		= "<%=ConvSPChars(EG2_exp_grp(EG2_bp_cd3))%>"           '주문처 
		.frm1.txtSoldToPartyNm.value	= "<%=ConvSPChars(EG2_exp_grp(EG2_bp_nm3))%>"           '주문처명 
		.frm1.txtCustPoNo.value			= "<%=ConvSPChars(EG2_exp_grp(EG2_cust_po_no))%>"	   '고객주문번호 
		.frm1.HReqDlvyDt.value			=          "<%=UNIDateClientFormat(EG2_exp_grp(EG2_req_dlvy_dt))%>"	 '---납기일	
		.frm1.txtNetAmt.Text			= "<%=UNINumClientFormatByCurrency(EG2_exp_grp(EG2_net_amt),lgCurrency,ggAmtOfMoneyNo)%>" '수주순금액 
		.frm1.txtHNetAmt.value			= "<%=UNINumClientFormatByCurrency(EG2_exp_grp(EG2_net_amt),lgCurrency,ggAmtOfMoneyNo)%>" '---수주순금액 
		.frm1.txtHVATAmt.value			=      "<%=UNINumClientFormatByTax(EG2_exp_grp(EG2_vat_amt),lgCurrency,ggAmtOfMoneyNo)%>" '---부과세액 
		.frm1.txtCurrency.value			= "<%=lgCurrency%>"
		.frm1.txtShipToParty.value		= "<%=ConvSPChars(EG2_exp_grp(EG2_bp_cd2))%>"           '---납품처 
		.frm1.txtVatIncFlag.value		= "<%=ConvSPChars(EG2_exp_grp(EG2_vat_inc_flag))%>"     '부과세구분 
		.frm1.txtVatIncFlagNm.value		= "<%=ConvSPChars(EG2_exp_grp(EG2_vat_inc_flag_nm))%>"  '부과세구분명 
		.frm1.txtHVatRate.value			= "<%=ConvSPChars(EG2_exp_grp(EG2_vat_rate))%>"         '---
		.frm1.txtHMaintNo.value 		= "<%=ConvSPChars(EG2_exp_grp(EG2_maint_no))%>"         '---
		.frm1.HPriceFlag.value 			= "<%=ConvSPChars(EG2_exp_grp(EG2_price_flag))%>"       '---
		.frm1.txtHConfirmFlg.value		= "<%=ConvSPChars(EG2_exp_grp(EG2_cfm_flag))%>"         '---
		.frm1.txtSoType.value			= "<%=ConvSPChars(EG2_exp_grp(EG2_so_type))%>"		    '---		
		.frm1.txtHPreSONo.value			= "<%=ConvSPChars(EG2_exp_grp(EG2_pre_doc_no))%>"       '---이전수주번호 
		.frm1.HRetItemFlag.value		= "<%=ConvSPChars(EG2_exp_grp(EG2_ret_item_flag))%>"    '---
		.frm1.txtHVATType.value			= "<%=ConvSPChars(EG2_exp_grp(EG2_vat_type))%>"		    '---
		.frm1.txtHVATIncFlag.value		= "<%=ConvSPChars(EG2_exp_grp(EG2_vat_inc_flag))%>"	    '---	
		.frm1.txtHVATIncFlagNm.value	= "<%=ConvSPChars(EG2_exp_grp(EG2_vat_inc_flag_nm))%>"  '---
		.frm1.txtHPayTermsCd.value		= "<%=ConvSPChars(EG2_exp_grp(EG2_pay_meth))%>"		    '---
		.frm1.txtHDealType.value		= "<%=ConvSPChars(EG2_exp_grp(EG2_deal_type))%>"	    '---
	
'		If "<%=EG2_exp_grp(EG2_cfm_flag)%>" = "Y" Then
'			.frm1.RdoConfirm.value = "N"
'			.frm1.btnConfirm.value = "확정취소"
'		ElseIf "<%=EG2_exp_grp(EG2_cfm_flag)%>" = "N" Then
'			.frm1.RdoConfirm.value = "Y"
'			.frm1.btnConfirm.value = "확정처리"
'		Else
'			.frm1.RdoConfirm.value = "Y"
'			.frm1.btnConfirm.value = "확정처리"
'		End IF
		
		.frm1.txtHSoNo.value = "<%=ConvSPChars(Request("txtConSoNo"))%>"		
		.frm1.HExportFlag.value = "<%=ConvSPChars(EG2_exp_grp(EG2_export_flag))%>"			'수출여부		
		.frm1.txtSoDt.value = "<%=UNIDateClientFormat(EG2_exp_grp(EG2_so_dt))%>"			'수주일 
 		
		If parent.lgIntFlgMode = parent.parent.OPMD_CMODE Then parent.CurFormatNumSprSheet		
		.lgIntFlgMode = OPMD_UMODE				
		parent.HideLotRetField
		
	End With
</Script>		
<%    
  
    '-------------------------
    ' 수주내역정보가 없을때.
    '-------------------------
    If CheckSYSTEMError(Err,True) = True Then
        Set iPS3G152 = Nothing		       
%>
		<Script Language=vbscript>
		With parent.frm1
'			If .vspdData.MaxRows = 0 Then
'				.btnConfirm.disabled   = True
'				.btnConfirm.value      = "확정처리"
'				.btnDNCheck.disabled   = True
'				.btnATPCheck.disabled  = True
'				.btnCTPCheck.disabled  = True
'				.btnAvlStkRef.disabled = True
'			End If

			If Trim(.txtHPreSONo.value) <> "" And UCase(Trim(.HRetItemFlag.value)) = "Y" Then
				parent.SetToolbar "11101011000111"
			ElseIf Trim(.txtHPreSONo.value) = "" And UCase(Trim(.HRetItemFlag.value)) = "Y" Then
				parent.SetToolbar "11101111001111"
			ElseIf UCase(Trim(.HRetItemFlag.value)) <> "Y" Then
				parent.SetToolbar "11101111001111"
			Else
				parent.SetToolbar "11101111001111"
			End If
			
			parent.SetDefaultPlant
			parent.ChangePlantColor
            .txtConSoNo.focus
		End With
		</Script>
<%

        Response.End                                              
        'Exit Sub
    End If   

    Set iPS3G152 = Nothing
		
	'----------------------------
	' 수주내역정보를 표시한다.
	'----------------------------
%>
<Script Language=vbscript>
    Dim LngLastRow      
    Dim iLngMaxRow       
    Dim iLngRow          
    Dim strTemp
    Dim istrData
	    
	With parent
		iLngMaxRow = .frm1.vspdData.MaxRows
<%        
		For iLngRow = 0 To UBound(EG1_exp_grp,1)
		    If iLngRow < C_SHEETMAXROWS_D  Then
		    Else
		       iStrNextKey = ConvSPChars(EG1_exp_grp(iLngRow, EG1_so_seq)) 
               Exit For
            End If 	           
%>           
			istrData = istrData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_item_cd))%>"     '품목코드			
			istrData = istrData & Chr(11)                                                             '품목코드팝업			
			istrData = istrData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_item_nm))%>"     '품목명			
			istrData = istrData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_spec))%>"        '규격			
			istrData = istrData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_so_unit))%>"     '단위 
			istrData = istrData & Chr(11)                                                             '단위팝업			
			istrData = istrData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_tracking_no))%>" 'Tracking No.
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(cdbl(EG1_exp_grp(iLngRow, EG1_so_qty)) - cdbl(EG1_exp_grp(iLngRow, EG1_req_qty)), ggQty.DecPoint, 0)%>" '수주잔량 
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_so_qty), ggQty.DecPoint, 0)%>"          '수량 
			istrData = istrData & Chr(11) & "<%=UNINumClientFormatByCurrency(EG1_exp_grp(iLngRow, EG1_so_price), lgCurrency, ggUnitCostNo)%>" '단가 
		    istrData = istrData & Chr(11) & "0"                                                       '단가체크버튼			
		    
			Select Case "<%=EG1_exp_grp(iLngRow, EG1_price_flag)%>"                                   '단가구분 (진단가/가단가)
			Case "Y"
				istrData = istrData & Chr(11) & "진단가"
			Case "N"
				istrData = istrData & Chr(11) & "가단가"
			Case Else
				istrData = istrData & Chr(11)
			End Select

        	If "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_vat_inc_flag))%>" =  "2" Then   	    	  
		     	istrData = istrData & Chr(11) & "<%=UNINumClientFormatByCurrency(cdbl(EG1_exp_grp(iLngRow, EG1_net_amt)) + cdbl(EG1_exp_grp(iLngRow, EG1_vat_amt)), lgCurrency, ggAmtOfMoneyNo)%>"
			Else
				istrData = istrData & Chr(11) & "<%=UNINumClientFormatByCurrency(EG1_exp_grp(iLngRow, EG1_net_amt), lgCurrency, ggAmtOfMoneyNo)%>"							
			End If		

			istrData = istrData & Chr(11) & "<%=UNINumClientFormatByCurrency(EG1_exp_grp(iLngRow, EG1_net_amt), lgCurrency, ggAmtOfMoneyNo)%>" '수주순금액(거래화폐)
			istrData = istrData & Chr(11) & "<%=UNINumClientFormatByTax(EG1_exp_grp(iLngRow, EG1_vat_amt),lgCurrency,ggAmtOfMoneyNo)%>"   'VAT 금액 
    		istrData = istrData & Chr(11) & "<%=UNINumClientFormatByCurrency(EG1_exp_grp(iLngRow, EG1_net_amt_loc), lgCurrency, ggAmtOfMoneyNo)%>" '수주순금액(자국금액)
    		istrData = istrData & Chr(11) & "<%=UNINumClientFormatByTax(EG1_exp_grp(iLngRow, EG1_vat_amt_loc),lgCurrency,ggAmtOfMoneyNo)%>"   'VAT 금액(자국금액)
			
    		istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_plant_cd))%>"     '공장코드 
			istrData = istrData & Chr(11)                                                                               '공장팝업			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_plant_nm))%>"     '공장명									
			istrData = istrData & Chr(11) & "<%=UNIDateClientFormat(EG1_exp_grp(iLngRow, EG1_dlvy_dt))%>"      '납기일			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_bp_cd))%>"        '납품처			
			istrData = istrData & Chr(11)                                                                               '납품처팝업			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_bp_nm))%>"        '납품처명			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_hs_no))%>"        'HS Code			
			istrData = istrData & Chr(11)                                                                               'Hs Code Popup
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_tol_more_rate), ggExchRate.DecPoint, 0)%>" '과부족허용율(+)
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_tol_less_rate), ggExchRate.DecPoint, 0)%>" '과부족허용율(-)
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_vat_type))%>"        'VAT Type
			istrData = istrData & Chr(11)                                                                                  'VAT Type Popup
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_vat_type_nm))%>"     'VAT Name
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_vat_rate), ggExchRate.DecPoint, 0)%>" 'VAT Rate
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_vat_inc_flag))%>"    'VAT포함구분 			
			'istrData = istrData & Chr(11) &                  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_vat_inc_flag_nm))%>" 'VAT포함구분명 
			Select Case "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_vat_inc_flag))%>"
			Case "1"
				istrData = istrData & Chr(11) & "별도"
			Case "2"
				istrData = istrData & Chr(11) & "포함"
			Case Else
				istrData = istrData & Chr(11)
			End Select
			
     		istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_ret_type))%>"        '반품유형 
			istrData = istrData & Chr(11)                                                                                  '반품유형팝업 
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_ret_type_nm))%>"     '반품유형명			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_lot_no))%>"          'Lot No			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_lot_seq))%>"         'Lot Seq			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_dn_no))%>"           '이전출하번호			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_dn_seq))%>"          '이전출하순번			
			istrData = istrData & Chr(11) & "<%=UNIDateClientFormat(EG1_exp_grp(iLngRow, EG1_promise_dt))%>"      '출하요청일											
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_bonus_qty), ggQty.DecPoint, 0)%>" '할증수량(덤)			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_sl_cd))%>"           '창고코드			
			istrData = istrData & Chr(11)                                                                                  '창고팝업			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_sl_nm))%>"           '창고명			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_remark))%>"          '비고			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_so_status))%>"       '수주진행상태			
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_bill_qty), ggQty.DecPoint, 0)%>"      '매출수량			
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_so_base_qty), ggQty.DecPoint, 0)%>"   '재고수량			
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_bonus_base_qty), ggQty.DecPoint, 0)%>"'덤재고수량			
			istrData = istrData & Chr(11) & ""                                                                             '관리순번			
			istrData = istrData & Chr(11) & ""                                                                             '주문서순번			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_aps_host))%>"        'APSHost			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_aps_port))%>"        'APSPort			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_ctp_times))%>"       'CTPTimes			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_flag))%>"            'CTPCheckFlag			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_so_seq))%>"          '수주순번			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_pre_doc_no))%>"      '이전수주번호			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_pre_doc_seq))%>"     '이전수주순번 
			istrData = istrData & Chr(11) & iLngMaxRow + <%=iLngRow%>			
			istrData = istrData & Chr(11) & Chr(12)			
<%      
		Next
%>    

		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowDataByClip istrData
	
		.lgStrPrevKey = "<%=iStrNextKey%>"
		
    	.frm1.txtHSoNo.value = "<%=ConvSPChars(Request("txtConSoNo"))%>"   ' Request값을 hidden input으로 넘겨줌 
					
'		If "<%=EG2_exp_grp(EG2_cfm_flag)%>" = "Y" Then         '확정버튼 처리				
'			If "<%=EG2_exp_grp(EG2_auto_dn_flag)%>" = "N" Or "<%=EG2_exp_grp(EG2_so_sts)%>" = 1 Then  
'				.frm1.btnDNCheck.disabled = True              '출하버튼 처리 
'			ElseIf "<%=EG2_exp_grp(EG2_auto_dn_flag)%>" = "Y" And "<%=EG2_exp_grp(EG2_so_sts)%>" <> 1 Then
'				.frm1.btnDNCheck.disabled = False
'			End IF								
'			.frm1.btnATPCheck.disabled = True                 'ATP CHECK버튼 처리 
'		Else				
'			.frm1.btnDNCheck.disabled = True                  '출하버튼 처리 							
'			.frm1.btnATPCheck.disabled = False                'ATP CHECK버튼 처리	
'		End IF
		.frm1.txtPlant.value	= "<%=ConvSPChars(EG1_exp_grp(0, EG1_plant_cd))%>"  '맨처음 품목의 공장을 입력한다.
	    .frm1.txtPlantNm.value	= "<%=ConvSPChars(EG1_exp_grp(0, EG1_plant_nm))%>"
		.DbQueryOk
	
	End With

</Script>
<%																			

Case CStr(UID_M0002)																'☜: 저장 요청을 받음 
	
	Dim orPosition
									
    Err.Clear																		'☜: Protect system from crashing
    
	If Request("txtMaxRows") = "" Then
		Call ServerMesgBox("MaxRows 조건값이 비어있습니다!",vbInformation, I_MKSCRIPT)              
		Response.End 
	End If

	'iLngMaxRow = CInt(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수 
    I1_command  = "SAVE"
    I2_s_so_hdr(l2_s_so_no) = Trim(Request("txtHSoNo"))
    'I2_s_so_hdr(l2_s_cfm_flag) = ""
    	
    Set iPS3G127 = Server.CreateObject("PS3G127.cSAmendSalesOrder")      
    
    If CheckSYSTEMError(Err,True) = True Then
       Response.End
       'Exit Sub
    End If    

    Call iPS3G127.S_AMEND_SALES_ORDER(gStrGlobalCollection, I1_command, I2_s_so_hdr, _
                              Trim(Request("txtSpread")), iErrorPosition)   

    If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
       Set iPS3G127 = Nothing
       Response.End
       'Exit Sub
	End If

    Set iPS3G127 = Nothing 			

%>
<Script Language=vbscript>
	With parent																			
		.DbSaveOk
	End With
</Script>
<%																

Case "DNCheck"																'☜: 현재 출하요청처리 받음 

    Err.Clear                                                               '☜: Protect system from crashing
    Dim iPS3G117
    
    I1_s_so_hdr = Trim(Request("txtSoNo"))
    
    Set iPS3G117 = Server.CreateObject("PS3G117.cSCreateDnBySoSvr")          
    
    If CheckSYSTEMError(Err,True) = True Then
       Response.End
       'Exit Sub
    End If    

    Call iPS3G117.S_CREATE_DN_BY_SO_SVR(gStrGlobalCollection, I1_s_so_hdr)    
    
	If CheckSYSTEMError(Err,True) = True Then
       Set iPS3G117 = Nothing	
       Response.End		                                               
       'Exit Sub
    End If
	
	Set iPS3G117 = Nothing	

%>
<Script Language=vbscript>
'=	parent.FncQuery
	parent.DbSaveOk()
</Script>		
<%

Case "PRICE"																'☜: 현재 출하요청처리 받음 

    Err.Clear                                                               '☜: Protect system from crashing
    
    Dim I4_s_so_dtl
    Const S321_I4_so_unit = 0
    Const S321_I4_so_qty = 1
    ReDim I4_s_so_dtl(1)
    
    DIm E1_s_so_dt
    Public Const S321_E1_so_price = 0
    Public Const S321_E1_bonus_qty = 1
    
    Dim pS31121PR
    
    Dim I1_ief_supplied_select_char    
    Dim I2_b_item_item_cd    
    Dim I3_s_so_hdr_so_no 
    
    I1_ief_supplied_select_char = Trim(Request("lsPriceQty"))   
    
    I2_b_item_item_cd = Trim(Request("lsItemCode"))
    
    I3_s_so_hdr_so_no = Trim(Request("txtHSoNo"))
    
    I4_s_so_dtl(S321_I4_so_unit) = Trim(Request("lsSoUnit"))
    I4_s_so_dtl(S321_I4_so_qty) = UNIConvNum(Request("lsSoQty"),0)	

    Set pS31121PR = Server.CreateObject("PS3G112.cSGetSoPriceSvr")    

    If CheckSYSTEMError(Err,True) = True Then
		Response.End		
    End If
    
    E1_s_so_dtl = pS31121PR.S_GET_SO_PRICE_SVR (gStrGlobalCollection,I1_ief_supplied_select_char, I2_b_item_item_cd, _
												I3_s_so_hdr_so_no,I4_s_so_dtl)
       
	If CheckSYSTEMError(Err,True) = True Then
       Set pS31121PR = Nothing		                                                 '☜: Unload Comproxy DLL
       Response.End 
    End If   

    Set pS31121PR = Nothing	
  
  'test용 
   'if iSEMPTY(E1_s_so_dtl(S321_E1_so_price)) THEN E1_s_so_dtl(S321_E1_so_price) = 0 END IF
   ' Response.Write "Price= " & E1_s_so_dtl(S321_E1_so_price)
   %><BR><% 
    
  '  if iSEMPTY(E1_s_so_dtl(S321_E1_bonus_qty)) THEN E1_s_so_dtl(S321_E1_bonus_qty) = 0 END IF
   ' Response.Write "QTY= " & E1_s_so_dtl(S321_E1_bonus_qty)
%>
<Script Language=vbscript>

	With parent																'☜: 화면 처리 ASP 를 지칭함 

		.frm1.vspdData.Row = <%=Request("PRow")%>

		Select Case "<%=Trim(Request("lsPriceQty"))%>"
		Case "A"
			.frm1.vspdData.Col = .C_SoPrice
			.frm1.vspdData.Text = "<%=UNINumClientFormatByCurrency(E1_s_so_dtl(S321_E1_so_price), Trim(Request("txtCurrency")), ggUnitCostNo)%>"
			.frm1.vspdData.Col = .C_BonusQty
			.frm1.vspdData.Text = "<%=UNINumClientFormat(E1_s_so_dtl(S321_E1_bonus_qty), ggQty.DecPoint, 0)%>"
		Case "P"
			.frm1.vspdData.Col = .C_SoPrice
			.frm1.vspdData.Text = "<%=UNINumClientFormatByCurrency(E1_s_so_dtl(S321_E1_so_price), Trim(Request("txtCurrency")), ggUnitCostNo)%>"
		Case "Q"
			.frm1.vspdData.Col = .C_BonusQty
			.frm1.vspdData.Text = "<%=UNINumClientFormat(E1_s_so_dtl(S321_E1_bonus_qty), ggQty.DecPoint, 0)%>"
		End Select

	End With

</Script>		
<%

Case "btnCONFIRM"																	'☜: 확정처리 요청을 받음 
									
    Err.Clear																		'☜: Protect system from crashing
    Dim iPS3G150
    Redim I1_s_so_hdr(1)
    
	iLngMaxRow = CInt(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수 
	
	I1_s_so_hdr(0) = Trim(Request("txtHSoNo"))
	I1_s_so_hdr(1) = Trim(Request("RdoConfirm"))
   
    Set iPS3G150 = Server.CreateObject("PS3G150.cSConfirmSalesOrderSvr")      
    
    If CheckSYSTEMError(Err,True) = True Then
       Response.End
       'Exit Sub
    End If    

    Call iPS3G150.S_CONFIRM_SALES_ORDER_SVR(gStrGlobalCollection, I1_s_so_hdr)   
    
    If CheckSYSTEMError(Err,True) = True Then
       Set iPS3G121 = Nothing 
       Response.End
       'Exit Sub
    End If 

    Set iPS3G121 = Nothing 		  
%>
<Script Language=vbscript>
	parent.DbSaveOk()	
</Script>
<%					


'************ 옛날 확정처리 로직 (지우지 말자) ***************
Case "--OLD--btnCONFIRM"																	'☜: 확정처리 요청을 받음 
									
    Err.Clear																		'☜: Protect system from crashing

	iLngMaxRow = CInt(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수 

    Set PS3G121 = Server.CreateObject("S31121.cSSoDtlSvr")

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If Err.Number <> 0 Then
		Set PS3G121 = Nothing																'☜: ComProxy UnLoad
		Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)
		Err.Clear						'⊙:
		Response.End																		'☜: Process End
	End If

    '-----------------------
    'Data manipulate area
    '-----------------------
    arrTemp = Split(Request("txtSpread"), gRowSep)
    
    lGrpCnt = 0
    
    For iLngRow = 1 To iLngMaxRow

		lGrpCnt = lGrpCnt + 1														'☜: Group Count    

		arrVal = Split(arrTemp(iLngRow-1), gColSep)

		'수주순번 
		If Trim(arrVal(2)) <> "" then PS3G121.ImpItemSSoDtlSoSeq(lGrpCnt) = Trim(arrVal(2))
		'품목코드		
		PS3G121.ImpItemBItemItemCd(lGrpCnt) = Trim(arrVal(3))
		'수량 
		PS3G121.ImpItemSSoDtlSoQty(lGrpCnt) = UNICDbl(arrVal(4),0)
		'단위 
		PS3G121.ImpItemSSoDtlSoUnit(lGrpCnt) = Trim(arrVal(5))
		'납기일						
		PS3G121.ImpItemSSoDtlDlvyDt(lGrpCnt) = Trim(arrVal(6))
		'납품처 
		PS3G121.ImpItemBBizPartnerBpCd(lGrpCnt) = Trim(arrVal(7))
		'단가 
		PS3G121.ImpItemSSoDtlSoPrice(lGrpCnt) = UNICDbl(arrVal(8),0)
		'단가구분(진단가/가단가)
		PS3G121.ImpItemSSoDtlPriceFlag(lGrpCnt) = Trim(arrVal(9))
		'할증수량(덤)				
		PS3G121.ImpItemSSoDtlBonusQty(lGrpCnt) = UNICDbl(arrVal(10),0)
		'HS Code
		PS3G121.ImpItemSSoDtlHsNo(lGrpCnt) = Trim(arrVal(11))
		'과부족허용율(+)
		PS3G121.ImpItemSSoDtlTolMoreRate(lGrpCnt) = UNICDbl(arrVal(12),0)
		'과부족허용율(-)
		PS3G121.ImpItemSSoDtlTolLessRate(lGrpCnt) = UNICDbl(arrVal(13),0)
		'공장코드 
		PS3G121.ImpItemBPlantPlantCd(lGrpCnt) = Trim(arrVal(14))
		'창고코드 
		PS3G121.ImpItemBStorageLocationSlCd(lGrpCnt) = Trim(arrVal(15))
		'비고  		                
		PS3G121.ImpItemSSoDtlRemark(lGrpCnt) = Trim(arrVal(16))
		'관리순번 
		PS3G121.ImpItemSSoDtlMaintSeq(lGrpCnt) = UNICDbl(arrVal(17),0)      
		'주문순번 
		PS3G121.ImpItemSSoDtlCustPoSeq(lGrpCnt) = UNICDbl(arrVal(18),0)    
				  			                	
		If lGrpCnt >= 40 Or iLngRow = iLngMaxRow Then									'⊙: 5개를 Group으로, 나머지 일때 

			PS3G121.ImpSSoHdrSoNo = Trim(Request("txtHSoNo"))
			PS3G121.ImpSSoHdrCfmFlag = Trim(Request("RdoConfirm"))
			PS3G121.ImpSWksUserUserId = Trim(Request("txtInsrtUserId"))
			PS3G121.CommandSent = lsConfirm

		    PS3G121.ComCfg = gConnectionString
			PS3G121.Execute

		    '-----------------------
		    'Com action result check area(OS,internal)
		    '-----------------------
		    If Err.Number <> 0 Then
				Set PS3G121 = Nothing
				Call ServerMesgBox(Err.Description, vbCritical, I_MKSCRIPT)              
				Response.End
				Err.Clear  
		    End If

		    '-----------------------
		    'Com action result check area(DB,internal)
		    '-----------------------
			If Not (PS3G121.OperationStatusMessage = MSG_OK_STR) Then
				Select Case PS3G121.OperationStatusMessage
					Case MSG_DEADLOCK_STR
						Call DisplayMsgBox2("999999", "25", "deadlock or timeout" , I_MKSCRIPT)
					Case MSG_DBERROR_STR
						Call DisplayMsgBox2(PS3G121.ExportErrEabSqlCodeSqlcode, _
								    PS3G121.ExportErrEabSqlCodeSeverity, _
								    PS3G121.ExportErrEabSqlCodeErrorMsg, I_MKSCRIPT)
					Case Else
						Call DisplayMsgBox(PS3G121.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
				End Select
				Set PS3G121 = Nothing
				Response.End 
			End If
							           
		    PS3G121.Clear
		    lGrpCnt = 0
		End If
	Next		
   

    Set PS3G121 = Nothing                                                   '☜: Unload Comproxy
	Response.End															'☜: Process End


Case "ItemByHsCode"															'☜: 품목별에 따른 HS CODE Change

	Dim iPB3C104
	
    Dim I1_b_item    
    Dim prE1_b_item
    Const prE1_item_cd = 0
    Const prE1_item_nm = 1
    Const prE1_formal_nm = 2
    Const prE1_spec = 3
    Const prE1_basic_unit = 4
    Const prE1_item_acct = 5
    Const prE1_item_class = 6
    Const prE1_phantom_flg = 7
    Const prE1_hs_cd = 8
    Const prE1_hs_unit = 9
    Const prE1_unit_weight = 10
    Const prE1_unit_of_weight = 11
    Const prE1_draw_no = 12
    Const prE1_item_image_flg = 13
    Const prE1_blanket_pur_flg = 14
    Const prE1_base_item_cd = 15
    Const prE1_proportion_rate = 16
    Const prE1_valid_flg = 17
    Const prE1_valid_from_dt = 18
    Const prE1_valid_to_dt = 19
    Const prE1_vat_type = 20
    Const prE1_vat_rate = 21
    
    
	Err.Clear
	
	I1_b_item = Trim(Request("ItemCd"))
	
    Set iPB3C104 = Server.CreateObject("PB3C104.cBLkUpItem")     
    
    If CheckSYSTEMError(Err,True) = True Then
       Response.End
       'Exit Sub
    End If
    
    Call iPB3C104.B_LOOK_UP_ITEM(gStrGlobalCollection, I1_b_item, , , , , prE1_b_item)	
	
	If CheckSYSTEMError(Err,True) = True Then
       Set iPB3C104 = Nothing	
       Response.End		                                               
       'Exit Sub
    End If	

%>

<Script Language="vbscript">
		With parent.frm1.vspdData
			.Row 	= "<%=Request("CRow")%>"
			.Col 	= parent.C_ItemName
			.text	= "<%=ConvSPChars(prE1_b_item(prE1_item_nm))%>"
			.Col 	= parent.C_ItemSpec
			.text	= "<%=ConvSPChars(prE1_b_item(prE1_spec))%>"
			.Col 	= parent.C_HsNo
			.text	= "<%=ConvSPChars(prE1_b_item(prE1_hs_cd))%>"
			.Col 	= parent.C_SoUnit
			.text	= "<%=ConvSPChars(prE1_b_item(prE1_basic_unit))%>"

			.Col	= parent.C_VatType

			If .text = "" Then
				If Len(parent.frm1.txtHVATType.value) Then
					.text	= parent.frm1.txtHVATType.value
				Else
					.text	= "<%=ConvSPChars(prE1_b_item(prE1_vat_type))%>"
				End If
			End If
			
			.Col	= parent.C_VatIncFlag

			If .text = "" Then 
				If Len(parent.frm1.txtHVATIncFlag.value) Then
					.Col	= parent.C_VatIncFlag
					.text	= parent.frm1.txtHVATIncFlag.value

					.Col	= parent.C_VatIncFlagNm
					Select Case parent.frm1.txtHVATIncFlag.value
					Case "1"
						parent.frm1.vspdData.Text = "별도"
					Case "2"
						parent.frm1.vspdData.Text = "포함"
					End Select
				End If			
			End If
			Call parent.SetVatType(<%=Request("CRow")%>)
			
			parent.lsPriceQty = "Q"
			Call parent.GetItemPrice(<%=Request("CRow")%>)
			Call parent.PricePadChange(<%=Request("CRow")%>)
		End With	
</Script>
<%
    Set iPB3C104 = Nothing

Case "CheckCreditlimit"															'☜: 여신한도 체크 

	Dim iPS3G113
	Dim BalanceAmt
	    
    'VB를 String으로 수정한뒤 전부수정(성미정) 
	
    Dim I1_b_currency
    Const S324_I1_currency = 0
    Redim I1_b_currency(S324_I1_currency)
    
    Dim I2_ief_supplied
    Const S324_I2_total_currency = 0
    Redim I2_ief_supplied(S324_I2_total_currency)
    
    Dim I3_s_so_hdr
    Const S324_I3_so_no = 0
    Redim I3_s_so_hdr(S324_I3_so_no)
    
    Dim E1_exchange_variable
    Const S324_E1_num_value_15_2 = 0
    
    Dim E2_ief_supplied
    Const S324_E2_command = 0
    Const S324_E2_select_char = 1
    
    I3_s_so_hdr(S324_I3_so_no) = Trim(Request("txtHSONo"))
    I2_ief_supplied(S324_I2_total_currency) = UNIConvNum(Request("txtNetAmt"),0)
    I1_b_currency(S324_I1_currency) = Trim(Request("txtCurrency"))  
        
    Set iPS3G113 = Server.CreateObject("PS3G113.cChkSoCreditLimitSvr")  
    
    If CheckSYSTEMError(Err,True) = True Then
       Response.End
       'Exit Sub
    End If
 
    Call iPS3G113.S_CHK_SO_CREDIT_LIMIT_SVR(gStrGlobalCollection, I1_b_currency, I2_ief_supplied, _
                                            I3_s_so_hdr, , E1_exchange_variable, E2_ief_supplied)

    If cStr(Err.Description) = "B_MESSAGE" & Chr(11) & "201727" then    

           BalanceAmt = UNINumClientFormat(E1_exchange_variable(S324_E1_num_value_15_2), ggAmtOfMoney.DecPoint, 0)          
%>     
<Script Language=vbscript>
           Dim msgCreditlimit
		   Dim BalanceAmt		

		   BalanceAmt = FormatNumber(Parent.UNICDbl("<%=BalanceAmt%>"), parent.ggAmtofMoney.DecPoint, -2)	
		   
		   msgCreditlimit = parent.parent.DisplayMsgBox("201929", parent.VB_YES_NO, parent.gCurrency, BalanceAmt)		   
		   
		   If msgCreditlimit = vbYes Then parent.DbSave
		      

</Script>
<% 
    ElseIf cStr(Err.Description) = "B_MESSAGE" & Chr(11) & "201722" then
  
           BalanceAmt = UNINumClientFormat(E1_exchange_variable(S324_E1_num_value_15_2), parent.parent.ggAmtofMoney.DecPoint, 0)
%>     
<Script Language=vbscript>      
           Dim BalanceAmt	
           			
		   BalanceAmt = FormatNumber(Parent.parent.UNICDbl("<%=BalanceAmt%>"), parent.parent.ggAmtofMoney.DecPoint, -2)
		   
		   Call parent.parent.DisplayMsgBox("201722", "X", parent.parent.gCurrency, BalanceAmt)   	
</Script>
<%      
    Else
           If CheckSYSTEMError(Err,True) = True Then    
               Set iPS3G113 = Nothing	
               Response.End		                                               
               'Exit Sub	
           End if    
%>     
<Script Language=vbscript>     
           Call parent.DbSave()
</Script>
<%             
    END IF     	                                               
     
    Set iPS3G113 = Nothing  
   	
    
'자동출하생성에서 체크함. 출하에서는 여신체크안함 
Case "CheckADNCreditlimit"														

	Dim pS14115

    Set pS14115 = Server.CreateObject("S14115.S14115ChkAdnCreditLimitSvr")
    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If Err.Number <> 0 Then
		Set pS14115 = Nothing																'☜: ComProxy UnLoad
		Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)
		Err.Clear						'⊙:
		Response.End																		'☜: Process End
	End If
    
    pS14115.ImpSSoHdrSoNo = Trim(Request("txtHSONo"))
    pS14115.ServerLocation = ggServerIP
    
    '-----------------------
    'Com action area
    '-----------------------
    pS14115.ComCfg = gConnectionString
    pS14115.Execute

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
	   Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)                                 '⊙:
	   Set pS14115 = Nothing																	'☜: ComProxy UnLoad
	   Response.End																				'☜: Process End
	End If
	
	'-----------------------
	'Com action result check area(DB,internal)
	'-----------------------
%>
<Script Language=vbscript>
<%
		If Not (pS14115.OperationStatusMessage = MSG_OK_STR) Then
			If pS14115.OperationStatusMessage = 201727 Then
%>
				Dim msgCreditlimit
				msgCreditlimit = parent.parent.DisplayMsgBox("201727", parent.VB_YES_NO, "X", "X")
				If msgCreditlimit = vbYes Then parent.RunAutoDN
<%
			ElseIf pS14115.OperationStatusMessage = 201722 Then
%>
				Call parent.parent.DisplayMsgBox("201722", "X", "X", "X")
<%
			Else
%>
				Call parent.parent.DisplayMsgBox("<%=pS14115.OperationStatusMessage%>", "X", "X", "X")
<%
			End If
		Else
%>
			Call parent.RunAutoDN()
<%	
		End If
%>

</Script>
<%					

	Set pS14115 = Nothing
	Response.End

End Select

'==============================================================================
' 사용자 정의 서버 함수 
'==============================================================================
%>
