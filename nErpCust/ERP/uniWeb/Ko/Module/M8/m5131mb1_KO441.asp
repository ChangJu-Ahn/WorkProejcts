<%@ LANGUAGE=VBSCript%>
<% Option explicit%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m5112mb1
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/05/12
'*  8. Modified date(Last)  : 2003/06/09
'*  9. Modifier (First)     : Shin Jin Hyun
'* 10. Modifier (Last)      : Lee Eun Hee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/05/05 : ..........
'* 14. Business Logic of m5131ma1(매입일괄확정)
'**********************************************************************************************
    Dim lgOpModeCRUD
 
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
	Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")
    Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")
	
    lgOpModeCRUD  = Request("txtMode") 
	

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)      
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
Sub SubBizQueryMulti()
	On Error Resume Next
    Err.Clear	
	
	Dim PvArr
	Dim lGrpCnt
	Dim iTotstrData
	
	Dim iM51211																	'☆ : 입력/수정용 ComProxy Dll 사용 변수 
	Dim iM51118																	'☆ : 조회용 ComProxy Dll 사용 변수 

	Dim iCommandSent
	Dim istrData
	Dim StrNextKey		' 다음 값 
	Dim lgStrPrevKey	' 이전 값 
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
	Dim GroupCount  
	Dim lgCurrency        
	Dim index,Count     ' 저장 후 Return 해줄 값을 넣을때 쓴는 변수 
	Const C_SHEETMAXROWS_D  = 100
	
	'import view
	Dim i1_pur_grp
	Dim i2_bp_cd
	Dim i3_m_iv_hdr
	Dim i4_iv_Dt
	Dim i5_iv_no
	Dim i6_iv_type_cd

	Const M592_I3_iv_dt = 0    '  View Name : imp_fr m_iv_hdr
    Const M592_I3_tax_biz_area = 1
    Const M592_I3_posted_flg = 2
	Redim i3_m_iv_hdr(M592_I3_posted_flg)
	
	'export view
	Dim E1_b_biz_area
	Dim E2_m_iv_hdr
	Dim E3_b_biz_partner
	Dim E4_b_pur_grp
	Dim E5_m_iv_type
	Dim EG1_exp_group
	
	'export group view
	Const M592_E1_biz_area_nm = 0		'  View Name : exp_cond b_biz_area
    Const M592_E1_biz_area_cd = 1
	
	Const M592_E2_iv_no = 0
	 
	Const M592_E3_bp_cd = 0    '  View Name : exp_cond b_biz_partner
    Const M592_E3_bp_nm = 1

	Const M592_E4_pur_grp = 0			'  View Name : exp_cond b_pur_grp
    Const M592_E4_pur_grp_nm = 1
	
	Const M592_E5_iv_type_cd = 0		'  View Name : exp_cond m_iv_type
    Const M592_E5_iv_type_nm = 1
    Const M592_E5_import_flg = 2
    Const M592_E5_except_flg = 3
    Const M592_E5_ret_flg = 4


	Const M592_EG1_E1_iv_type_nm = 0 
    Const M592_EG1_E1_iv_type_cd = 1
    Const M592_EG1_E2_biz_area_nm = 2
    Const M592_EG1_E3_iv_no = 3    
    Const M592_EG1_E3_iv_dt = 4
    Const M592_EG1_E3_ap_post_dt = 5
    Const M592_EG1_E3_pay_dt = 6
    Const M592_EG1_E3_posted_flg = 7
    Const M592_EG1_E3_sppl_iv_no = 8
    Const M592_EG1_E3_payee_cd = 9
    Const M592_EG1_E3_build_cd = 10
    Const M592_EG1_E3_pur_org = 11
    Const M592_EG1_E3_iv_biz_area = 12
    Const M592_EG1_E3_tax_biz_area = 13
    Const M592_EG1_E3_iv_cost_cd = 14
    Const M592_EG1_E3_pay_meth = 15
    Const M592_EG1_E3_pay_dur = 16
    Const M592_EG1_E3_pay_terms_txt = 17
    Const M592_EG1_E3_pay_type = 18
    Const M592_EG1_E3_gross_doc_amt = 19
    Const M592_EG1_E3_gross_loc_amt = 20
    Const M592_EG1_E3_net_doc_amt = 21
    Const M592_EG1_E3_net_loc_amt = 22
    Const M592_EG1_E3_cash_doc_amt = 23
    Const M592_EG1_E3_cash_loc_amt = 24
    Const M592_EG1_E3_iv_cur = 25
    Const M592_EG1_E3_xch_rt = 26
    Const M592_EG1_E3_vat_type = 27
    Const M592_EG1_E3_vat_rt = 28
    Const M592_EG1_E3_tot_vat_doc_amt = 29
    Const M592_EG1_E3_tot_vat_loc_amt = 30
    Const M592_EG1_E3_tot_diff_doc_amt = 31
    Const M592_EG1_E3_tot_diff_loc_amt = 32
    Const M592_EG1_E3_pay_bank_cd = 33
    Const M592_EG1_E3_pay_acct_cd = 34
    Const M592_EG1_E3_pp_no = 35
    Const M592_EG1_E3_pp_doc_amt = 36
    Const M592_EG1_E3_pp_loc_amt = 37
    Const M592_EG1_E3_remark = 38
    Const M592_EG1_E3_loan_no = 39
    Const M592_EG1_E3_loan_doc_amt = 40
    Const M592_EG1_E3_loan_loc_amt = 41
    Const M592_EG1_E3_bl_no = 42
    Const M592_EG1_E3_bl_doc_no = 43
    Const M592_EG1_E3_lc_doc_no = 44
    Const M592_EG1_E3_ext1_cd = 45
    Const M592_EG1_E3_gl_no = 46
    Const M592_EG1_E4_bp_cd = 47    '  View Name : exp_item b_biz_partner
    Const M592_EG1_E4_bp_nm = 48
    Const M592_EG1_E5_pur_grp = 49    '  View Name : exp_item b_pur_grp
    Const M592_EG1_E5_pur_grp_nm = 50
	
                                                             
    If Len(Trim(Request("txtFrIvDt"))) Then
		If UNIConvDate(Request("txtFrIvDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txtFrIvDt", 0, I_MKSCRIPT)
		    Exit Sub	
		End If
	End If
	
	If Len(Trim(Request("txtToIvDt"))) Then
		If UNIConvDate(Request("txtToIvDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txtToIvDt", 0, I_MKSCRIPT)
		    Exit Sub	
		End If
	End If
		
	lgStrPrevKey = Request("lgStrPrevKey")
	Set iM51118 = Server.CreateObject("PM8G118.cMListIvHdrS")    

	If CheckSYSTEMError(Err,True) = true then 
		Set iM51118 = Nothing			
		Exit Sub
	End If
	
    If Request("txtFrIvDt") <> "" then
		i3_m_iv_hdr(M592_I3_iv_dt)			= UNIConvDate(Request("txtFrIvDt"))
	Else
		i3_m_iv_hdr(M592_I3_iv_dt)			= "1900-01-01"
	End If

	If Request("txtToIvDt") <> "" then
		i4_iv_Dt							= UNIConvDate(Request("txtToIvDt"))
	Else
		i4_iv_Dt							= "1900-01-01"
	End If

    i6_iv_type_cd								= Request("txtIvType")
    i2_bp_cd									= Request("txtSppl")
    i3_m_iv_hdr(M592_I3_tax_biz_area)			= Request("txtBizArea")
    i3_m_iv_hdr(M592_I3_posted_flg)				= Request("txtApPost")
	i1_pur_grp									= Request("txtGrp")
    i5_iv_no 									= lgStrPrevKey
   
  
    
    Call iM51118.M_LIST_IV_HDR_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, i1_pur_grp, _
								 i2_bp_cd, i3_m_iv_hdr, i4_iv_Dt, i5_iv_no, _
								 i6_iv_type_cd, E1_b_biz_area, EG1_exp_group, E2_m_iv_hdr, _
								 E3_b_biz_partner, E4_b_pur_grp, E5_m_iv_type)				

	If CheckSYSTEMError(Err,True) = true then 
		Set iM51118 = Nothing			
		Exit Sub
	End If
	
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "with parent" & vbCr
	Response.Write "	.frm1.txtIvTypeNm.value  = """ & ConvSPChars(E5_m_iv_type(M592_E5_iv_type_nm))      & """" & vbCr
	Response.Write "	.frm1.txtSpplNm.value    = """ & ConvSPChars(E3_b_biz_partner(M592_E3_bp_nm))      & """" & vbCr
	Response.Write "	.frm1.txtBizAreaNm.value = """ & ConvSPChars(E1_b_biz_area(M592_EG1_E3_tax_biz_area))    & """" & vbCr
	Response.Write "	.frm1.txtGrpNm.value     = """ & ConvSPChars(E4_b_pur_grp(M592_E4_pur_grp_nm))      & """" & vbCr
	Response.Write "End With "   & vbCr
    Response.Write "</Script>"                  & vbCr

	iLngMaxRow = CInt(Request("txtMaxRows"))											'Save previous Maxrow                                          
	GroupCount = UBound(EG1_exp_group,1)
	ReDim PvArr(GroupCount)
	
	 IF GroupCount <> 0 then
		IF EG1_exp_group(GroupCount,M592_EG1_E3_iv_no) =  E2_m_iv_hdr(M592_E2_iv_no) then
				StrNextKey = ""
		Else
				StrNextKey = EG1_exp_group(GroupCount,M592_EG1_E3_iv_no )
		End If
	End if
	
	'============멀티 처리===
	
	For iLngRow = 0 To UBound(EG1_exp_group,1)
		If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
		   StrNextKey = EG1_exp_group(GroupCount,M592_EG1_E3_iv_no)
		   Exit For
		End If  
				
		istrData = istrData & Chr(11) & "0"
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M592_EG1_E3_posted_flg ) )
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M592_EG1_E3_iv_no ))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M592_EG1_E4_bp_cd ))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M592_EG1_E4_bp_nm ))
        istrData = istrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(EG1_exp_group(iLngRow,M592_EG1_E3_gross_doc_amt),0)
        istrData = istrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(EG1_exp_group(iLngRow,M592_EG1_E3_tot_vat_doc_amt),0)
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M592_EG1_E3_iv_cur ))
        istrData = istrData & Chr(11) & UNIDateClientFormat(EG1_exp_group(iLngRow,M592_EG1_E3_iv_dt ))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M592_EG1_E5_pur_grp ))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M592_EG1_E5_pur_grp_nm) )
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M592_EG1_E3_tax_biz_area))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M592_EG1_E2_biz_area_nm ))
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M592_EG1_E3_gl_no))
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
        istrData = istrData & Chr(11) & Chr(12)
		PvArr(lGrpCnt) = istrData
        lGrpCnt = lGrpCnt + 1
        istrData = ""
    Next
    
    iTotstrData = Join(PvArr, "")
    
	Set iM51118 = Nothing
	
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr												'☜: 화면 처리 ASP 를 지칭함 
    Response.Write "	.ggoSpread.Source    = .frm1.vspdData " & vbCr
    Response.Write  "    .frm1.vspdData.Redraw = False   "                     & vbCr    
    Response.Write "	.ggoSpread.SSShowData """ & iTotstrData	& """ , ""F""" & vbCr	
    
    Response.Write  "    Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData," & iLngMaxRow + 1 & "," & iLngMaxRow + iLngRow & ",Parent.C_Currency,Parent.C_IvAmt,""A"" ,""I"",""X"",""X"")" & vbCr
    Response.Write  "    Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData," & iLngMaxRow + 1 & "," & iLngMaxRow + iLngRow & ",Parent.C_Currency,Parent.C_VatAmt,""A"" ,""I"",""X"",""X"")" & vbCr
    
    Response.Write "	.lgStrPrevKey        = """ & StrNextKey & """" & vbCr  
    Response.Write "	.frm1.hdnFrDt.value     = """ & ConvSPChars(Request("txtFrIvDt"))      & """" & vbCr
	Response.Write "	.frm1.hdnToDt.value     = """ & ConvSPChars(Request("txtToIvDt"))      & """" & vbCr
	Response.Write "	.frm1.hdnIvType.value   = """ & ConvSPChars(Request("txtIvType"))      & """" & vbCr
	Response.Write "	.frm1.hdnSppl.value     = """ & ConvSPChars(Request("txtSppl"))        & """" & vbCr
	Response.Write "	.frm1.hdnBizArea.value  = """ & ConvSPChars(Request("txtBizArea"))     & """" & vbCr
	Response.Write "	.frm1.hdnGrp.value      = """ & ConvSPChars(Request("txtGrp"))         & """" & vbCr
	Response.Write "	.frm1.hdnApFlg.value    = """ & ConvSPChars(Request("txtApPost"))      & """" & vbCr
    Response.Write "	.DbQueryOk "		    	   & vbCr 
    Response.Write "	.frm1.vspdData.focus "		   & vbCr 
    Response.Write "    .frm1.vspdData.Redraw = True   "                      & vbCr   
    Response.Write "End With"   & vbCr
    Response.Write "</Script>" & vbCr    

		
End Sub    
'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	
	On Error Resume next
	Err.Clear	
	
	Dim iRowsData
	Dim iColsData
	Dim iPM8G211
	Dim L_SelectChar
	Dim I3_m_batch_ap_post_wks
	Dim IG1_imp_dtl_group				'☜: Protect system from crashing
	Dim pvCB
	Dim itxtSpread
	Dim itxtSpreadArr
    Dim itxtSpreadArrCount, ii
	Dim iErrorPosition
	Dim i
    Dim iCUCount
    
    Const M557_I3_ap_dt_type = 0
    Const M557_I3_ap_dt = 1
    Const M557_I3_import_flg = 2
    
    Const M557_IG1_I1_count = 0
    Const M557_IG1_I2_iv_no = 1
    Const M557_IG1_I3_ap_dt = 2
    
	
	Redim I3_m_batch_ap_post_wks(2)
	             
    itxtSpread = ""
             
    iCUCount = Request.Form("txtCUSpread").Count
             
    itxtSpreadArrCount = -1
             
    ReDim itxtSpreadArr(iCUCount)
             
    For ii = 1 To iCUCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
    Next
    itxtSpread = Join(itxtSpreadArr,"")
             
	Set iPM8G211 = server.CreateObject("PM8G211.cMPostApS")    
 
	If CheckSYSTEMError(Err,True) = true Then 		
		Set iPM8G211 = Nothing												'☜: ComPlus Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	End if
	
	iRowsData = Split(itxtSpread,gRowSep)
	
	I3_m_batch_ap_post_wks(M557_I3_ap_dt_type)		= Trim(Request("hdnApDateFlg"))
	I3_m_batch_ap_post_wks(M557_I3_import_flg)		= Trim(Request("hdnImportFlg"))
		
	L_SelectChar		= Trim(Request("hdnApFlg"))
	
	pvCB = "F"
	ReDim IG1_imp_dtl_group(ubound(iRowsData) - 1, 2)
	
	For i = 0 To ubound(iRowsData) - 1
		iColsData = Split(iRowsData(i),gColSep)
			
		IG1_imp_dtl_group(i, M557_IG1_I1_count)			=	iColsData(2)	'ROW NO.
		IG1_imp_dtl_group(i, M557_IG1_I2_iv_no)			=	iColsData(0)
		IG1_imp_dtl_group(i, M557_IG1_I3_ap_dt)			=	iColsData(1)
	Next
			
	Call iPM8G211.M_POST_AP_SVR(pvCB,gStrGlobalCollection, L_SelectChar, IG1_imp_dtl_group, I3_m_batch_ap_post_wks, iErrorPosition)

	If CheckSYSTEMError2(Err, True, iErrorPosition & "행:","","","","") = True Then
	  	Set iPM8G211 = Nothing
	  	Response.Write "<Script language=vbs> " & vbCr  
		Response.Write " Parent.RemovedivTextArea "      & vbCr							'☜: 화면 처리 ASP 를 지칭함 
		Response.Write "</Script> "
	  	Exit Sub
	End If
		
	Set iPM8G211 = Nothing
                       

    Response.Write "<Script language=vbs> " & vbCr  
    Response.Write " Parent.DbSaveOk "      & vbCr							'☜: 화면 처리 ASP 를 지칭함 
    Response.Write "</Script> "           
        
End Sub 


  
%>                   
