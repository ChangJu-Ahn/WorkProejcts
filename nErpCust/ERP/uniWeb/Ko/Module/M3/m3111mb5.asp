<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->


<%	

call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
'call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m3111mb5
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/05/11
'*  8. Modified date(Last)  : 2000/05/11
'*  9. Modifier (First)     : Shin Jin Hyun
'* 10. Modifier (Last)      : Min, HJ
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'* 14. Business Logic of m3111ma5(발주일괄마감)
'**********************************************************************************************
    Dim lgOpModeCRUD
 
    On Error Resume Next                                                            '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    lgOpModeCRUD  = Request("txtMode") 
										                                              '☜: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call  SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
    '         Call SubBizDelete()
	
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    On Error Resume Next                                                             '☜: Protect system from crashing
    	
	Dim iPM3G18C																	'☆ : 조회용 ComProxy Dll 사용 변수 
	
	Dim StrNextKey		' 다음 값 
	Dim lgStrPrevKey	' 이전 값 
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
	Dim GroupCount      
	Dim arrValue    
	Dim istrData
	 
	Const C_SHEETMAXROWS_D  = 100
	
    Dim I1_m_pur_ord_dtl_next
    Dim I2_m_pur_ord_hdr_next
    Dim I3_m_pur_ord_hdr_po_dt
    Dim I4_b_biz_partner_bp_cd
    Dim I5_m_pur_ord_hdr
		Const M100_I5_po_no = 0
		Const M100_I5_po_dt = 1
		Const M100_I5_po_type_cd = 2
		Const M100_I5_dlvy_start_dt = 3
		Const M100_I5_dlvy_end_dt = 4
    Redim I5_m_pur_ord_hdr(M100_I5_dlvy_end_dt)
    '************* 발주형태 조회조건 추가 LSW 2006-08-17 ************
    
    Dim I6_b_pur_grp
    Dim I7_ief_supplied
    Dim I8_plant_cd
    Dim I9_item_cd
    
    Dim E1_b_biz_partner
    Const M100_E1_bp_cd = 0
    Const M100_E1_bp_nm = 1

    Dim E2_b_pur_grp
    Const M100_E2_pur_grp = 0
    Const M100_E2_pur_grp_nm = 1

    Dim EG1_exp_group
    Const M100_EG1_E1_m_pur_ord_hdr_po_no = 0
    Const M100_EG1_E1_m_pur_ord_hdr_po_dt = 1
    Const M100_EG1_E2_b_plant_plant_cd = 2
    Const M100_EG1_E2_b_plant_plant_nm = 3
    Const M100_EG1_E3_b_item_item_cd = 4
    Const M100_EG1_E3_b_item_item_nm = 5
    Const M100_EG1_E4_m_pur_ord_dtl_po_seq_no = 6
    Const M100_EG1_E4_m_pur_ord_dtl_dlvy_dt = 7
    Const M100_EG1_E4_m_pur_ord_dtl_po_qty = 8
    Const M100_EG1_E4_m_pur_ord_dtl_po_unit = 9
    Const M100_EG1_E4_m_pur_ord_dtl_po_base_qty = 10
    Const M100_EG1_E4_m_pur_ord_dtl_po_base_unit = 11
    Const M100_EG1_E4_m_pur_ord_dtl_fr_trans_coef = 12
    Const M100_EG1_E4_m_pur_ord_dtl_to_trans_coef = 13
    Const M100_EG1_E4_m_pur_ord_dtl_po_prc = 14
    Const M100_EG1_E4_m_pur_ord_dtl_po_prc_flg = 15
    Const M100_EG1_E4_m_pur_ord_dtl_po_doc_amt = 16
    Const M100_EG1_E4_m_pur_ord_dtl_po_loc_amt = 17
    Const M100_EG1_E4_m_pur_ord_dtl_rcpt_qty = 18
    Const M100_EG1_E4_m_pur_ord_dtl_iv_qty = 19
    Const M100_EG1_E4_m_pur_ord_dtl_lc_qty = 20
    Const M100_EG1_E4_m_pur_ord_dtl_bl_qty = 21
    Const M100_EG1_E4_m_pur_ord_dtl_cc_qty = 22
    Const M100_EG1_E4_m_pur_ord_dtl_po_sts = 23
    Const M100_EG1_E4_m_pur_ord_dtl_cls_flg = 24
    Const M100_EG1_E4_m_pur_ord_dtl_tracking_no = 25
    Const M100_EG1_E4_m_pur_ord_dtl_so_no = 26
    Const M100_EG1_E4_m_pur_ord_dtl_so_seq_no = 27
    Const M100_EG1_E4_m_pur_ord_dtl_sl_cd = 28
    Const M100_EG1_E4_m_pur_ord_dtl_rcpt_biz_area = 29
    Const M100_EG1_E4_m_pur_ord_dtl_ref_po_no = 30
    Const M100_EG1_E4_m_pur_ord_dtl_ref_po_seq_no = 31
    Const M100_EG1_E4_m_pur_ord_dtl_hs_cd = 32
    Const M100_EG1_E4_m_pur_ord_dtl_over_tol = 33
    Const M100_EG1_E4_m_pur_ord_dtl_under_tol = 34
    Const M100_EG1_E4_m_pur_ord_dtl_inspect_qty = 35
    Const M100_EG1_E5_b_biz_partner_bp_cd = 36
    Const M100_EG1_E5_b_biz_partner_bp_nm = 37
    Const M100_EG1_E3_b_item_spec = 38

    Dim E3_m_pur_ord_hdr
    Dim E4_m_pur_ord_dtl		

    Dim E5_b_plant
    Const M100_E5_plant_cd = 0
    Const M100_E5_plant_nm = 1
    
    Dim E6_m_config_process
    Const M100_E6_po_type_cd = 0
    Const M100_E6_po_type_nm = 1
    
    Dim E7_item 
    Const M100_E7_item_cd = 0
    Const M100_E7_itme_nm = 1

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
    If lgStrPrevKey <> "" then	
        arrValue = Split(lgStrPrevKey, gColSep)		
		I2_m_pur_ord_hdr_next = arrValue(0)
		I1_m_pur_ord_dtl_next = arrValue(1)
	else			
		I2_m_pur_ord_hdr_next = ""
		I1_m_pur_ord_dtl_next = ""
	End If	
  
    Set iPM3G18C = Server.CreateObject("PM3G18C.cMListPoDtlForClsS")    

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true Then 		
		Set iPM3G18C = Nothing												'☜: ComProxy Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	End if
    
    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    if Request("txtFrDt") = "" then
    	I5_m_pur_ord_hdr(M100_I5_po_dt)			= "1900-01-01"
    else
    	I5_m_pur_ord_hdr(M100_I5_po_dt)			= UniConvDate(Request("txtFrDt"))
    End if 
    if Request("txtToDt") = "" then
    	I3_m_pur_ord_hdr_po_dt			= "2999-12-31"
    else
    	I3_m_pur_ord_hdr_po_dt			= UniConvDate(Request("txtToDt"))
    End if
    if Request("txtStartDt") = "" then
    	I5_m_pur_ord_hdr(M100_I5_dlvy_start_dt)			= "1900-01-01"
    else
    	I5_m_pur_ord_hdr(M100_I5_dlvy_start_dt)			= UniConvDate(Request("txtStartDt"))
    End if 
    if Request("txtEndDt") = "" then
    	I5_m_pur_ord_hdr(M100_I5_dlvy_end_dt)			= "2999-12-31"
    else
    	I5_m_pur_ord_hdr(M100_I5_dlvy_end_dt)			= UniConvDate(Request("txtEndDt"))
    End if
    
    I4_b_biz_partner_bp_cd					= Request("txtSupplier")
    I5_m_pur_ord_hdr(M100_I5_po_no)			= Request("txtPoNo")'발주번호 
    '************* 발주형태 조회조건 추가 LSW 2006-08-17 ************
    I5_m_pur_ord_hdr(M100_I5_po_type_cd)	= Request("txtPotypeCd")
    I6_b_pur_grp							= Request("txtGroup")
    I7_ief_supplied							= Request("txtClsFlag")
    I8_plant_cd								= Request("txtPlantCd")
    I9_item_cd								= Request("txtItemCd")
							
    Call iPM3G18C.M_LIST_PO_DTL_FOR_CLS_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_m_pur_ord_dtl_next, I2_m_pur_ord_hdr_next, _
                    I3_m_pur_ord_hdr_po_dt, I4_b_biz_partner_bp_cd, I5_m_pur_ord_hdr, UCase(I6_b_pur_grp), I7_ief_supplied, I9_item_cd, E1_b_biz_partner, E2_b_pur_grp, _
                    EG1_exp_group, E3_m_pur_ord_hdr, E4_m_pur_ord_dtl, I8_plant_cd, E5_b_plant, E6_m_config_process, E7_item)			
	
	If CheckSYSTEMError(Err,True) = true Then 		
		Set iPM3G18C = Nothing												'☜: ComProxy Unload
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "with parent" & vbCr
		Response.Write "	.frm1.txtSupplierNm.value = """ & ConvSPChars(E1_b_biz_partner(M100_E1_bp_nm))      & """" & vbCr
		Response.Write "	.frm1.txtPur_Grp_Nm.value = """ & ConvSPChars(E2_b_pur_grp(M100_E2_pur_grp_nm))     & """" & vbCr
		Response.Write "	.frm1.txtPlantNm.value    = """ & ConvSPChars(E5_b_plant(M100_E5_plant_nm))         & """" & vbCr
		Response.Write "	.frm1.txtPotypeNm.value   = """ & ConvSPChars(E6_m_config_process(M100_E6_po_type_nm)) & """" & vbCr
		Response.Write "	.frm1.txtItemNm.value   = """ & ConvSPChars(E7_item(M100_E7_itme_nm)) & """" & vbCr
		Response.Write "End With "   & vbCr
		Response.Write "</Script>"                  & vbCr
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	End if
	
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "with parent" & vbCr
	Response.Write "	.frm1.txtSupplierNm.value = """ & ConvSPChars(E1_b_biz_partner(M100_E1_bp_nm))      & """" & vbCr
	Response.Write "	.frm1.txtPur_Grp_Nm.value = """ & ConvSPChars(E2_b_pur_grp(M100_E2_pur_grp_nm))     & """" & vbCr
	Response.Write "	.frm1.txtPlantNm.value    = """ & ConvSPChars(E5_b_plant(M100_E5_plant_nm))         & """" & vbCr
	Response.Write "	.frm1.txtPotypeNm.value   = """ & ConvSPChars(E6_m_config_process(M100_E6_po_type_nm)) & """" & vbCr
	Response.Write "	.frm1.txtItemNm.value	  = """ & ConvSPChars(E7_item(M100_E7_itme_nm)) & """" & vbCr
	Response.Write "End With "   & vbCr
    Response.Write "</Script>"                  & vbCr
    

	'-----------------------
	'Com action result check area(DB,internal)
	'-----------------------
	If CheckSYSTEMError(Err,True) = true Then 		
		Set iPM3G18C = Nothing												'☜: ComProxy Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	
	End if

	iLngMaxRow = Request("txtMaxRows")											'Save previous Maxrow                                                
    GroupCount = UBound(EG1_exp_group,1) 
    
	'-----------------------
	'Result data display area
	'----------------------- 
	For iLngRow = 0 To UBound(EG1_exp_group,1)
		If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
		   StrNextKey = ConvSPChars(E3_m_pur_ord_hdr) & gColSep & ConvSPChars(E4_m_pur_ord_dtl)
           Exit For
        End If  
		
		istrData = istrData & Chr(11) & "0"
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M100_EG1_E4_m_pur_ord_dtl_cls_flg))		'2003.05 정기 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M100_EG1_E1_m_pur_ord_hdr_po_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M100_EG1_E4_m_pur_ord_dtl_po_seq_no))		'발주순번 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M100_EG1_E2_b_plant_plant_cd))			'공장 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M100_EG1_E2_b_plant_plant_nm))			'공장명 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M100_EG1_E3_b_item_item_cd))			'품목 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M100_EG1_E3_b_item_item_nm))			'품목명 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M100_EG1_E3_b_item_spec))			'규격 
        istrData = istrData & Chr(11) & UNIDateClientFormat(EG1_exp_group(iLngRow,M100_EG1_E4_m_pur_ord_dtl_dlvy_dt))
        istrData = istrData & Chr(11) & UNIDateClientFormat(EG1_exp_group(iLngRow,M100_EG1_E1_m_pur_ord_hdr_po_dt))
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow,M100_EG1_E4_m_pur_ord_dtl_po_qty),ggQty.DecPoint,0)	'발주수량 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M100_EG1_E4_m_pur_ord_dtl_po_unit))	'단위 
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow,M100_EG1_E4_m_pur_ord_dtl_rcpt_qty),ggQty.DecPoint,0)	'입고수량 
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow,M100_EG1_E4_m_pur_ord_dtl_iv_qty),ggQty.DecPoint,0)	'매입수량 
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow,M100_EG1_E4_m_pur_ord_dtl_lc_qty),ggQty.DecPoint,0)	'Lc수량 
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow,M100_EG1_E4_m_pur_ord_dtl_bl_qty),ggQty.DecPoint,0)	'Bl수량 
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow,M100_EG1_E4_m_pur_ord_dtl_cc_qty),ggQty.DecPoint,0)	'Cc수량 
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow,M100_EG1_E4_m_pur_ord_dtl_inspect_qty),ggQty.DecPoint,0) 'Inspect Qty(검사중수량)
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M100_EG1_E5_b_biz_partner_bp_cd))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M100_EG1_E5_b_biz_partner_bp_nm))
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow                             
        istrData = istrData & Chr(11) & Chr(12)       
        
    Next  
    
    Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr												'☜: 화면 처리 ASP 를 지칭함 
    Response.Write "	.ggoSpread.Source          =  .frm1.vspdData "         & vbCr
    Response.Write "	.ggoSpread.SSShowData        """ & istrData	    & """" & vbCr	
    Response.Write "	.lgStrPrevKey              = """ & ConvSPChars(StrNextKey)   & """" & vbCr 
     
    Response.Write " .frm1.hdnFrDt.value     = """ & ConvSPChars(Request("txtFrDt"))     & """" & vbCr
	Response.Write " .frm1.hdnToDt.value     = """ & ConvSPChars(Request("txtToDt"))     & """" & vbCr
	Response.Write " .frm1.hdnStartDt.value     = """ & ConvSPChars(Request("txtStartDt"))     & """" & vbCr
	Response.Write " .frm1.hdnEndDt.value     = """ & ConvSPChars(Request("txtEndDt"))     & """" & vbCr
	Response.Write " .frm1.hdnSupplier.value = """ & ConvSPChars(Request("txtSupplier")) & """" & vbCr
	Response.Write " .frm1.hdnGroup.value    = """ & ConvSPChars(Request("txtGroup"))    & """" & vbCr
	Response.Write " .frm1.hdnPoNo.value     = """ & ConvSPChars(Request("txtPoNo"))     & """" & vbCr
	Response.Write " .frm1.txtHCfmFlag.value = """ & ConvSPChars(Request("txtClsFlag"))	 & """" & vbCr	'2003.05 정기 
	Response.Write " .frm1.hdnPlantCd.value  = """ & ConvSPChars(Request("txtPlantCd"))  & """" & vbCr	'2005.12
	Response.Write " .frm1.hdnPoTypeCd.value = """ & ConvSPChars(Request("txtPotypeCd"))  & """" & vbCr
	Response.Write " .frm1.hdnItemCd.value   = """ & ConvSPChars(Request("txtItemCd"))  & """" & vbCr
    Response.Write " .DbQueryOk "		    	  & vbCr 
    Response.Write " .frm1.vspdData.focus "		  & vbCr 
    Response.Write "End With" & vbCr
    Response.Write "</Script>" & vbCr


    Set iPM3G18C = Nothing
		
End Sub    
'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	
	Dim iPM3G16C																	'☆ : 입력/수정용 ComProxy Dll 사용 변수 
    Dim iErrorPosition
    Dim iStrSpread 
    
	On Error Resume Next
    Err.Clear																		'☜: Protect system from crashing
                                                
    Set iPM3G16C = Server.CreateObject("PM3G16C.cMClsPurOrdS")    

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true Then 		
		Set iPM3G16C = Nothing												'☜: ComProxy Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	
	End if
	
	iStrSpread = Trim(Request("txtSpread"))
	Call iPM3G16C.M_CLS_PUR_ORD_SVR(gStrGlobalCollection, iStrSpread, iErrorPosition) 
                   
    If CheckSYSTEMError2(Err,True,iErrorPosition & "행:","","","","") = True Then
       Set iPM3G16C = Nothing
	   Call SheetFocus(iErrorPosition, 1, I_MKSCRIPT)
       Exit Sub
	End If

    Set iPM3G16C = Nothing                                                   '☜: Unload Comproxy
    
	Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DbSaveOk "      & vbCr							'☜: 화면 처리 ASP 를 지칭함 
    Response.Write "</Script> "           

End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,arrColVal)
    Dim iSelCount
    
    On Error Resume Next

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status


End Sub

'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)
	dim strHTML

	If Trim(lRow) = "" Then Exit Function
	If iLoc = I_INSCRIPT Then
		strHTML = "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		'strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		Response.Write strHTML
	ElseIf iLoc = I_MKSCRIPT Then
		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		'strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write strHTML
	End If
End Function
%>
