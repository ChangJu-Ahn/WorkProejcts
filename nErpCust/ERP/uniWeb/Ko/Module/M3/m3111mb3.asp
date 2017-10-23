<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : m3111mb3
'*  4. Program Name         : 발주일괄확정등록 
'*  5. Program Desc         : 발주일괄확정등록 
'*  6. Component List       : PM3G118.cMListPurOrdHdrS / PM3G1R1.cMReleasePurOrdS
'*  7. Modified date(First) : 2000/05/11
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Shin Jin Hyun
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%	
call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 

    Dim lgOpModeCRUD
  
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgOpModeCRUD  = Request("txtMode") 
										                                              '☜: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call  SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
    End Select

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	On Error Resume Next
    Err.Clear			
	
	Dim iM31118rd																	'☆ : 조회용 ComProxy Dll 사용 변수 
	Dim iMax
	Dim PvArr
	
	Dim iStrData
	Dim StrNextKey		' 다음 값 
	Dim lgStrPrevKey	' 이전 값 
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
	Dim GroupCount          
	Const C_SHEETMAXROWS_D  = 100
	
	
    Dim I1_b_pur_grp_pur_grp
    Dim I2_m_pur_ord_hdr		'  View Name : imp_fr m_pur_ord_hdr
    Dim I3_m_pur_ord_hdr_po_dt	'  View Name : imp_to m_pur_ord_hdr
    Dim I4_b_biz_partner_bp_cd	
    Dim I5_m_pur_ord_hdr_po_no	'  View Name : imp_next m_pur_ord_hdr
    Dim I6_m_config_process_po_type_cd
    Dim E1_b_biz_partner
    Dim E2_b_pur_grp
    Dim E3_m_pur_ord_hdr
    Dim EG1_exp_group
    Dim E4_m_config_process
  
    Const M237_I2_po_dt = 0    '  View Name : imp_fr m_pur_ord_hdr
    Const M237_I2_release_flg = 1
    Const M237_I2_ret_flg = 2
    Redim I2_m_pur_ord_hdr(M237_I2_ret_flg)
    
    Const M237_E1_bp_cd = 0    '  View Name : exp_cond b_biz_partner
    Const M237_E1_bp_nm = 1


    Const M237_E2_pur_grp = 0    '  View Name : exp_cond b_pur_grp
    Const M237_E2_pur_grp_nm = 1

    Const M237_E3_po_no = 0    '  View Name : exp_next m_pur_ord_hdr

    Const M237_EG1_E1_po_type_cd = 0    '  View Name : exp_item m_config_process
    Const M237_EG1_E1_po_type_nm = 1
    Const M237_EG1_E2_pur_grp = 2    '  View Name : exp_item b_pur_grp
    Const M237_EG1_E2_pur_grp_nm = 3
    Const M237_EG1_E3_bp_cd = 4    '  View Name : exp_item b_biz_partner
    Const M237_EG1_E3_bp_nm = 5
    Const M237_EG1_E4_po_no = 6    '  View Name : exp_item m_pur_ord_hdr
    Const M237_EG1_E4_po_dt = 7
    Const M237_EG1_E4_po_cur = 8
    Const M237_EG1_E4_tot_po_doc_amt = 9
    Const M237_EG1_E4_pay_meth = 10
    Const M237_EG1_E4_pay_dur = 11
    Const M237_EG1_E4_vat_type = 12
    Const M237_EG1_E4_vat_rt = 13
    Const M237_EG1_E4_merg_pur_flg = 14
    Const M237_EG1_E4_pur_org = 15
    Const M237_EG1_E4_pur_biz_area = 16
    Const M237_EG1_E4_pur_cost_cd = 17
    Const M237_EG1_E4_xch_rt = 18
    Const M237_EG1_E4_pay_terms_txt = 19
    Const M237_EG1_E4_pay_type = 20
    Const M237_EG1_E4_tot_vat_doc_amt = 21
    Const M237_EG1_E4_tot_vat_loc_amt = 22
    Const M237_EG1_E4_tot_po_loc_amt = 23
    Const M237_EG1_E4_sppl_sales_prsn = 24
    Const M237_EG1_E4_sppl_tel_no = 25
    Const M237_EG1_E4_release_flg = 26
    Const M237_EG1_E4_cls_flg = 27
    Const M237_EG1_E4_import_flg = 28
    Const M237_EG1_E4_lc_flg = 29
    Const M237_EG1_E4_bl_flg = 30
    Const M237_EG1_E4_cc_flg = 31
    Const M237_EG1_E4_rcpt_flg = 32
    Const M237_EG1_E4_subcontra_flg = 33
    Const M237_EG1_E4_ret_flg = 34
    Const M237_EG1_E4_iv_flg = 35
    Const M237_EG1_E4_rcpt_type = 36
    Const M237_EG1_E4_issue_type = 37
    Const M237_EG1_E4_iv_type = 38
    Const M237_EG1_E4_sppl_cd = 39
    Const M237_EG1_E4_payee_cd = 40
    Const M237_EG1_E4_build_cd = 41
    Const M237_EG1_E4_remark = 42
    Const M237_EG1_E4_manufacturer = 43
    Const M237_EG1_E4_agent = 44
    Const M237_EG1_E4_applicant = 45
    Const M237_EG1_E4_offer_dt = 46
    Const M237_EG1_E4_expiry_dt = 47
    Const M237_EG1_E4_transport = 48
    Const M237_EG1_E4_incoterms = 49
    Const M237_EG1_E4_delivery_plce = 50
    Const M237_EG1_E4_packing_cond = 51
    Const M237_EG1_E4_inspect_means = 52
    Const M237_EG1_E4_dischge_city = 53
    Const M237_EG1_E4_dischge_port = 54
    Const M237_EG1_E4_loading_port = 55
    Const M237_EG1_E4_origin = 56
    Const M237_EG1_E4_sending_bank = 57
    Const M237_EG1_E4_invoice_no = 58
    Const M237_EG1_E4_fore_dvry_dt = 59
    Const M237_EG1_E4_shipment = 60
    Const M237_EG1_E4_charge_flg = 61
    Const M237_EG1_E4_tracking_no = 62
    Const M237_EG1_E4_so_no = 63

    Const M237_E4_po_type_cd = 0    '  View Name : exp_cond m_config_process
    Const M237_E4_po_type_nm = 1
    
	If Len(Trim(Request("txtFrDt"))) Then
		If UNIConvDate(Request("txtFrDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txtFrDt", 0, I_MKSCRIPT)
		    Exit Sub	
		End If
	End If
	If Len(Trim(Request("txtToDt"))) Then
		If UNIConvDate(Request("txtToDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txtToDt", 0, I_MKSCRIPT)
		    Exit Sub	
		End If
	End If

	lgStrPrevKey = Request("lgStrPrevKey")
  
    Set iM31118rd = Server.CreateObject("PM3G118.cMListPurOrdHdrS")    

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true Then 		
		Set iM31118rd = Nothing												'☜: ComProxy Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함	
	End if
    
    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    I5_m_pur_ord_hdr_po_no				= Request("lgStrPrevKey")
    If Request("txtFrDt") = "" then
    	I2_m_pur_ord_hdr(M237_I2_po_dt)			= "1900-01-01"
    Else
    	I2_m_pur_ord_hdr(M237_I2_po_dt)			= UniConvDate(Request("txtFrDt"))
    End if 
    If Request("txtToDt") = "" then
    	I3_m_pur_ord_hdr_po_dt			= "2999-12-31"
    Else
    	I3_m_pur_ord_hdr_po_dt			= UniConvDate(Request("txtToDt"))
    End if
    I4_b_biz_partner_bp_cd					= Trim(Request("txtSupplier"))
    I1_b_pur_grp_pur_grp					= Trim(Request("txtGroup"))
    I2_m_pur_ord_hdr(M237_I2_release_flg)	= Request("txtCfmflg")
    I2_m_pur_ord_hdr(M237_I2_ret_flg)		= "N"
    
    Call iM31118rd.M_LIST_PUR_ORD_HDR_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_b_pur_grp_pur_grp, _
										I2_m_pur_ord_hdr, I3_m_pur_ord_hdr_po_dt, I4_b_biz_partner_bp_cd, _
										I5_m_pur_ord_hdr_po_no, I6_m_config_process_po_type_cd, _
										E1_b_biz_partner, E2_b_pur_grp, E3_m_pur_ord_hdr, _
										EG1_exp_group, E4_m_config_process)

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err,True) = true Then 		
		Set iM31118rd = Nothing												'☜: ComProxy Unload
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "with parent" & vbCr
		Response.Write "	.frm1.txtSupplierNm.value = """ & ConvSPChars(E1_b_biz_partner(M237_E1_bp_nm))   & """" & vbCr
		Response.Write "	.frm1.txtPur_Grp_Nm.value = """ & ConvSPChars(E2_b_pur_grp(M237_E2_pur_grp_nm))  & """" & vbCr
		Response.Write "	.frm1.txtPur_Grp.focus"   & vbCr
		Response.Write "End With "   & vbCr
		Response.Write "</Script>"                  & vbCr
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	End if
	
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "with parent" & vbCr
	Response.Write "	.frm1.txtSupplierNm.value = """ & ConvSPChars(E1_b_biz_partner(M237_E1_bp_nm))   & """" & vbCr
	Response.Write "	.frm1.txtPur_Grp_Nm.value = """ & ConvSPChars(E2_b_pur_grp(M237_E2_pur_grp_nm))  & """" & vbCr
	Response.Write "End With "   & vbCr
    Response.Write "</Script>"                  & vbCr
    

	'-----------------------
	'Com action result check area(DB,internal)
	'-----------------------
	If CheckSYSTEMError(Err,True) = true Then 		
		Set iM31118rd = Nothing												'☜: ComProxy Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	
	End if

	iLngMaxRow = Request("txtMaxRows")											'Save previous Maxrow                                                
    GroupCount = UBound(EG1_exp_group,1)
    
	If ConvSPChars(EG1_exp_group(GroupCount, M237_EG1_E4_po_no)) = ConvSPChars(E3_m_pur_ord_hdr) Then  'next값..체크요!!
		StrNextKey = ""
	Else
		StrNextKey = E3_m_pur_ord_hdr
	End If
	'-----------------------
	'Result data display area
	'----------------------- 
	iMax = UBound(EG1_exp_group,1)
	ReDim PvArr(iMax)
	For iLngRow = 0 To UBound(EG1_exp_group,1)
		If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
		   StrNextKey = E3_m_pur_ord_hdr  'next값...
           Exit For
        End If  
		If EG1_exp_group(iLngRow, M237_EG1_E4_release_flg) = "Y" then
			istrData = istrData & Chr(11) & "1"	
		Else
			istrData = istrData & Chr(11) & "0"	
		End if
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M237_EG1_E4_po_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M237_EG1_E1_po_type_cd))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M237_EG1_E1_po_type_nm))
        istrData = istrData & Chr(11) & UniDateClientFormat(EG1_exp_group(iLngRow, M237_EG1_E4_po_dt))			
        istrData = istrData & Chr(11) & UniNumClientFormat(EG1_exp_group(iLngRow, M237_EG1_E4_tot_po_doc_amt),ggAmtOfMoney.DecPoint,0)	
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M237_EG1_E4_po_cur))	       
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M237_EG1_E3_bp_cd))	
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M237_EG1_E3_bp_nm))															    
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow                             
        istrData = istrData & Chr(11) & Chr(12)       

		PvArr(iLngRow) = istrData
		istrData=""
    Next  
    istrData = Join(PvArr, "")

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr												'☜: 화면 처리 ASP 를 지칭함 
	Response.Write ".frm1.vspdData.redraw = false" & vbCr
    Response.Write "	.ggoSpread.Source          =  .frm1.vspdData "         & vbCr
    Response.Write "	.ggoSpread.SSShowData        """ & istrData	    & """,""" & "F" & """" & vbCr	
    Response.Write "	.lgStrPrevKey              = """ & StrNextKey   & """" & vbCr 
%>
	Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,-1,-1,.C_Curr,.C_PoAmt,"A","I","X","X")
<%    
    Response.Write " .frm1.hdnFrDt.value     = """ & Request("txtFrDt")                  & """" & vbCr
	Response.Write " .frm1.hdnToDt.value     = """ & Request("txtToDt")                  & """" & vbCr
	Response.Write " .frm1.hdnSupplier.value = """ & ConvSPChars(Request("txtSupplier")) & """" & vbCr
	Response.Write " .frm1.hdnGroup.value    = """ & ConvSPChars(Request("txtGroup"))    & """" & vbCr
	Response.Write " .frm1.hdnCfmflg.value   = """ & ConvSPChars(Request("txtCfmflg"))   & """" & vbCr
    Response.Write " .DbQueryOk "		    	  & vbCr 
	Response.Write ".frm1.vspdData.redraw = True" & vbCr
    Response.Write " .frm1.vspdData.focus "		  & vbCr 
    Response.Write "End With" & vbCr
    Response.Write "</Script>" & vbCr
    

    Set iM31118rd = Nothing
		
End Sub    
'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	
	On Error Resume Next
    Err.Clear																		'☜: Protect system from crashing
    
	Dim iM31211																	'☆ : 입력/수정용 ComProxy Dll 사용 변수 

	Dim iErrorPosition
	Dim itxtSpread
    Dim itxtSpreadArr
    Dim itxtSpreadArrCount

    Dim iCUCount
    Dim ii
             
    itxtSpread = ""
             
    iCUCount = Request.Form("txtCUSpread").Count
             
    itxtSpreadArrCount = -1
             
    ReDim itxtSpreadArr(iCUCount)
             
    For ii = 1 To iCUCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
    Next
    itxtSpread = Join(itxtSpreadArr,"")

    Response.Write "<Script language=vbs> " & vbCr   
    Response.Write "Parent.RemovedivTextArea "      & vbCr   
    Response.Write "</Script> "      & vbCr   
	
    Set iM31211 = Server.CreateObject("PM3G1R1.cMReleasePurOrdS")    

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
	If CheckSYSTEMError(Err,True) = true Then 		
		Set iM31211 = Nothing												'☜: ComProxy Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	End if

	Call iM31211.M_RELEASE_PUR_ORD_SVR(gStrGlobalCollection, gUsrId, itxtSpread, iErrorPosition) 
                   
	If CheckSYSTEMError2(Err,True,iErrorPosition & "행:","","","","") = True Then
		Set iM31211 = Nothing
		Call SheetFocus(iErrorPosition, 1, I_MKSCRIPT)
		Exit Sub
	End If
 
    Set iM31211 = Nothing                                                   '☜: Unload Comproxy
                               
	Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DbSaveOk "      & vbCr							'☜: 화면 처리 ASP 를 지칭함 
    Response.Write "</Script> "             & vbCr
End Sub    

'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)

	If Trim(lRow) = "" Then Exit Function
	If iLoc = I_INSCRIPT Then
		strHTML = "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		Response.Write strHTML
	ElseIf iLoc = I_MKSCRIPT Then
		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write strHTML
	End If
End Function

%>
