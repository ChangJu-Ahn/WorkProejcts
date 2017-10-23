<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "A", "NOCOOKIE", "MB")
	Call LoadBNumericFormatB("I", "A","NOCOOKIE","MB")    
    '---------------------------------------Common-----------------------------------------------------------
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
    lgOpModeCRUD      = Request("txtMode")                                          '☜: Read Operation Mode (CRUD)

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
          '  Call SubBizQuery()
			Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection                    
            Call SubBizQueryMulti()
			Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection    
        Case CStr(UID_M0002)                                                         '☜: Save
          '  Call SubBizSave()
            Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
            Call SubBizDelete()
        Case CStr(UID_M0004)                                                         '☜: Update
            Call SubBizSaveMultiUpdate()

        Case CStr(UID_M0002)                                                         '☜: Save
          '  Call SubBizSave()
            Call ubBizSaveMultiUpdate()

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
    '-------------------------------------------------------------------------------------------------
	'IMPORTS View 상수 
	'-------------------------------------------------------------------------------------------------
	Const A390_I3_org_change_id = 0    '[CONVERSION INFORMATION]  View Name : in_dept_cd b_acct_dept
	Const A390_I3_dept_cd = 1

	Const A390_I4_temp_gl_no		= 0												'View Name : in_temp_gl a_temp_gl
	Const A390_I4_temp_gl_dt		= 1
	Const A390_I4_gl_type			= 2
	Const A390_I4_insrt_user_id		= 3
	Const A390_I4_updt_user_id		= 4
	Const A390_I4_issued_dt			= 5
	Const A390_I4_gl_input_type		= 6
	Const A390_I4_cr_amt			= 7
	Const A390_I4_cr_loc_amt		= 8
	Const A390_I4_dr_amt			= 9
	Const A390_I4_dr_loc_amt		= 10
	Const A390_I4_conf_fg			= 11
	Const A390_I4_insrt_dt			= 12
	Const A390_I4_updt_dt			= 13
	Const A390_I4_hq_brch_fg		= 14
	Const A390_I4_internal_cd		= 15

	Dim obj1PADG010
	Dim iCommandSent
	Dim I3_b_acct_dept
	Dim I4_a_temp_gl

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	Redim I3_b_acct_dept(A390_I3_dept_cd)
	Redim I4_a_temp_gl(A390_I4_internal_cd)

	iCommandSent							= "DELETE"
	I4_a_temp_gl(A390_I4_temp_gl_no)		= UCase(Request("txtTempGlNo"))
	I4_a_temp_gl(A390_I4_temp_gl_dt)		= UNIConvDateCompanyToDB(Request("txtTempGlDt"),NULL)
	I3_b_acct_dept(A390_I3_org_change_id)	= Request("txtOrgChangeId")
	I3_b_acct_dept(A390_I3_dept_cd)			= UCase(Request("txtDeptCd"))
	

    Set obj1PADG010 = CreateObject("PADG010.cAMngTmpGlHqSvr")

    If CheckSYSTEMError(Err, True) = True Then
       Exit Sub
    End If


	


	Call obj1PADG010.A_MANAGE_TEMP_GL_HQ_SVR(gStrGloBalCollection, iCommandSent, , , I3_b_acct_dept, I4_a_temp_gl)

	If CheckSYSTEMError(Err, True) = True Then
       Set obj1PADG010 = Nothing
		Exit Sub
    End If

	Set obj1PADG010 = nothing

	Response.Write "<Script Language=vbscript>	" & vbcr
	Response.Write " With Parent				" & vbCr
	Response.Write " 	.DbDeleteOK				" & vbCr
	Response.Write " End With					" & vbCr
	Response.Write "</Script>					" & vbCr

End Sub
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

	Const C_GLINPUTTYPE = "HQ"
	'-------------------------------------------------------------------------------------------------
	'IMPORTS View 상수 
	'-------------------------------------------------------------------------------------------------
	Const A352_I1_a_temp_gl_gl_no = 0
	Const A352_I1_a_temp_gl_item_seq_nk1 = 1
	'-------------------------------------------------------------------------------------------------
	'EXPORTS View 상수 
	'-------------------------------------------------------------------------------------------------
	Const A352_E1_a_temp_gl_temp_gl_no = 0
	Const A352_E1_a_temp_gl_temp_gl_dt = 1
	Const A352_E1_a_temp_gl_type = 2
	Const A352_E1_a_temp_gl_input_type = 3
	Const A352_E1_a_temp_gl_cr_amt = 4
	Const A352_E1_a_temp_gl_cr_loc_amt = 5
	Const A352_E1_a_temp_gl_dr_amt = 6
	Const A352_E1_a_temp_gl_dr_loc_amt = 7
	Const A352_E1_a_temp_gl_conf_fg = 8
	Const A352_E1_a_temp_gl_temp_gl_desc = 9
	Const A352_E1_a_temp_gl_project_no = 10
	Const A352_E1_a_temp_gl_org_change_id = 11
	Const A352_E1_a_temp_gl_dept_cd = 12
	Const A352_E1_a_temp_gl_dept_nm = 13
'	Const A347_E1_a_temp_gl_hq_brch_fg = 14
'	Const A347_E1_a_temp_gl_hq_brch_no = 15
	Const A352_E1_a_temp_gl_hq_brch_fg = 14
	Const A352_E1_a_temp_gl_hq_brch_no = 15

	Const A352_EG1_a_temp_gl_item_item_seq = 0
	Const A352_EG1_a_temp_gl_item_dept_cd = 1
	Const A352_EG1_a_temp_gl_item_dept_nm = 2
	Const A352_EG1_a_temp_gl_item_acct_cd = 3
	Const A352_EG1_a_temp_gl_item_acct_nm = 4
	Const A352_EG1_a_temp_gl_item_dr_cr_fg = 5
	Const A352_EG1_a_temp_gl_item_item_amt = 6
	Const A352_EG1_a_temp_gl_item_item_loc_amt = 7
	Const A352_EG1_a_temp_gl_item_vat_type = 8
	Const A352_EG1_a_temp_gl_item_item_desc = 9
	Const A352_EG1_a_temp_gl_item_xch_rate = 10
	Const A352_EG1_a_temp_gl_item_doc_cur = 11
	Const A352_EG1_a_temp_gl_item_project_no = 12
	Const A352_EG1_a_temp_gl_item_gl_no = 13
	Const A352_EG1_a_temp_gl_item_org_change_id = 14
	Const A352_EG1_a_temp_gl_item_acct_type = 15
'	Const A347_EG1_biz_area_cd = 16
'	Const A347_EG1_biz_area_nm = 17
	Const A352_EG1_biz_area_cd = 16
	Const A352_EG1_biz_area_nm = 17

	Dim obj2PAGG005
	Dim I1_a_temp_gl
	Dim E1_a_temp_gl
	Dim EG1_a_temp_gl_item
	Dim LngRow
	DIm LngMaxRow
	Dim iStrData
    DIM Strtempglno
    Dim lgStrSQL

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear

	LngMaxRow	= UNIConvNum(Request("txtMaxRows"),0)

	Redim I1_a_temp_gl(A352_I1_a_temp_gl_item_seq_nk1)

	I1_a_temp_gl(A352_I1_a_temp_gl_gl_no) = UCase(Request("txttempglno"))

	If Request("lgStrPrevKey") <> "" Then
		I1_a_temp_gl(A352_I1_a_temp_gl_item_seq_nk1) = Request("lgStrPrevKey")
	End If

	Strtempglno = Request("txttempglno")
		                           	
	lgStrSQL = ""
	lgStrSQL = lgStrSQL & " SELECT HQ_BRCH_NO = ISNULL(HQ_BRCH_NO,'') FROM A_TEMP_GL WHERE TEMP_GL_NO  = " & FilterVar(Strtempglno,"","S")
	
	If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)					'☜ : No data is found.
		lgErrorStatus  = "YES"
        Exit Sub
	Else
	    IF ConvSPChars(lgObjRs("HQ_BRCH_NO")) = "" then
    			Call DisplayMsgBox("114410", vbInformation, "", "", I_MKSCRIPT)					'☜ : No data is found.
				lgErrorStatus  = "YES"
			Exit Sub
	    
	    END IF
	End If	
	
	Set obj2PAGG005 = CreateObject("PAGG005.cALkUpTmpGlSvr")

    If CheckSYSTEMError(Err, True) = True Then
       Exit Sub
    End If        

	Call obj2PAGG005.A_LOOKUP_TEMP_GL_SVR (gStrGloBalCollection,I1_a_temp_gl,E1_a_temp_gl,EG1_a_temp_gl_item)

	If CheckSYSTEMError(Err, True) = True Then
       Set obj2PAGG005 = Nothing
       Exit Sub
    End If

	Set obj2PAGG005 = Nothing

	If isarray(EG1_a_temp_gl_item) then
		For LngRow = 0 To Ubound(EG1_a_temp_gl_item)

		    iStrData = iStrData & Chr(11) & EG1_a_temp_gl_item(LngRow, A352_EG1_a_temp_gl_item_item_seq)				'1  C_ItemSeq
	        iStrData = iStrData & Chr(11) & ConvSPChars(EG1_a_temp_gl_item(LngRow, A352_EG1_biz_area_cd))				'2  C_bizareacd
	        iStrData = iStrData & Chr(11) & ""																			'3  C_bizareapopup
	        iStrData = iStrData & Chr(11) & ConvSPChars(EG1_a_temp_gl_item(LngRow, A352_EG1_biz_area_nm))				'4  C_bizareanm
	        iStrData = iStrData & Chr(11) & ConvSPChars(EG1_a_temp_gl_item(LngRow, A352_EG1_a_temp_gl_item_dept_cd))	'5  C_deptcd
	        iStrData = iStrData & Chr(11) & ""																			'6  C_deptpopup
	        iStrData = iStrData & Chr(11) & ConvSPChars(EG1_a_temp_gl_item(LngRow, A352_EG1_a_temp_gl_item_dept_nm))	'7  C_deptnm
	        iStrData = iStrData & Chr(11) & ConvSPChars(EG1_a_temp_gl_item(LngRow, A352_EG1_a_temp_gl_item_acct_cd))	'8  C_accttcd
	        iStrData = iStrData & Chr(11) & ""																			'9  C_Acctpopup
	        iStrData = iStrData & Chr(11) & ConvSPChars(EG1_a_temp_gl_item(LngRow, A352_EG1_a_temp_gl_item_acct_nm))	'10 C_Acctnm
	        iStrData = iStrData & Chr(11) & ConvSPChars(EG1_a_temp_gl_item(LngRow, A352_EG1_a_temp_gl_item_dr_cr_fg))	'11 C_DrCrFg
	        iStrData = iStrData & Chr(11) & ""     	                   													'12 C_DrCrNm
		    iStrData = iStrData & Chr(11) & UNINumClientFormat(EG1_a_temp_gl_item(LngRow, A352_EG1_a_temp_gl_item_item_amt), ggAmtOfMoney.DecPoint, 0)		'12 C_ItemAmt
		    iStrData = iStrData & Chr(11) & UNINumClientFormat(EG1_a_temp_gl_item(LngRow, A352_EG1_a_temp_gl_item_item_loc_amt), ggAmtOfMoney.DecPoint, 0)	'14 C_ItemLocAmt
		    iStrData = iStrData & Chr(11) & ConvSPChars(EG1_a_temp_gl_item(LngRow, A352_EG1_a_temp_gl_item_item_desc))	'15 C_ItemDesc
		    iStrData = iStrData & Chr(11) & UNINumClientFormat(EG1_a_temp_gl_item(LngRow, A352_EG1_a_temp_gl_item_xch_rate), ggAmtOfMoney.DecPoint, 0) 		'16 C_ExchRate

		    iStrData = iStrData & Chr(11) & EG1_a_temp_gl_item(LngRow, A352_EG1_a_temp_gl_item_item_seq)				'17  C_ItemSeq2

	        iStrData = iStrData & Chr(11) & LngMaxRow + LngRow + 1
	        iStrData = iStrData & Chr(11) & Chr(12)

	    Next
	End If

	Response.Write "<Script Language=vbscript>	" & vbcr
	Response.Write " With parent				" & vbCr

	If isarray(E1_a_temp_gl) then
		Response.Write "	.frm1.hCongFg.value     = """ & ConvSPChars(E1_a_temp_gl(A352_E1_a_temp_gl_conf_fg))										& """" & vbCr
		Response.Write "	.frm1.txttempglno.value = """ & UCase(Trim(ConvSPChars(E1_a_temp_gl(A352_E1_a_temp_gl_temp_gl_no))))						& """" & vbCr
		Response.Write "	.frm1.txtDeptCd.value   = """ & UCase(Trim(ConvSPChars(E1_a_temp_gl(A352_E1_a_temp_gl_dept_cd))))							& """" & vbCr
		Response.Write "	.frm1.txtDeptNm.value   = """ & ConvSPChars(E1_a_temp_gl(A352_E1_a_temp_gl_dept_nm))										& """" & vbCr
		Response.Write "	.frm1.txtGLDt.Text      = """ & UNIDateClientFormat(E1_a_temp_gl(A352_E1_a_temp_gl_temp_gl_dt))								& """" & vbCr
		Response.Write "	.frm1.txtCrAmt.Text     = """ & UNINumClientFormat(E1_a_temp_gl(A352_E1_a_temp_gl_cr_amt), ggAmtOfMoney.DecPoint, 0)		& """" & vbCr
		Response.Write "	.frm1.txtDrAmt.Text     = """ & UNINumClientFormat(E1_a_temp_gl(A352_E1_a_temp_gl_dr_amt), ggAmtOfMoney.DecPoint, 0)		& """" & vbCr
		Response.Write "	.frm1.txtCrAmt.Text     = """ & UNINumClientFormat(E1_a_temp_gl(A352_E1_a_temp_gl_cr_loc_amt), ggAmtOfMoney.DecPoint, 0)	& """" & vbCr
		Response.Write "	.frm1.txtDrAmt.Text     = """ & UNINumClientFormat(E1_a_temp_gl(A352_E1_a_temp_gl_dr_loc_amt), ggAmtOfMoney.DecPoint, 0)	& """" & vbCr
		Response.Write "	.frm1.txtHqBrchNo.value = """ & ConvSPChars(E1_a_temp_gl(A352_E1_a_temp_gl_hq_brch_no))										& """" & vbCr
		Response.Write "	.frm1.txtDesc.value = """ & ConvSPChars(E1_a_temp_gl(A352_E1_a_temp_gl_temp_gl_desc))										& """" & vbCr
				
	End If

	If isarray(EG1_a_temp_gl_item) then
		Response.Write "	.frm1.txtDocCur.value   = """ & UCase(Trim(ConvSPChars(EG1_a_temp_gl_item(0,A352_EG1_a_temp_gl_item_doc_cur))))				& """" & vbCr
'		Response.Write "	.frm1.txtDocCur.value   = """ & UCase(Trim(ConvSPChars(EG1_a_temp_gl_item(0,A352_EG1_a_temp_gl_item_doc_cur))))				& """" & vbCr
	End If

	Response.Write "	.frm1.htxttempglno.value	= """ & UCase(Trim(ConvSPChars(Request("txttempglno")))) & """" & vbCr
	Response.Write " 	.ggoSpread.Source = .frm1.vspdData	      " & vbCr
	Response.Write " 	.ggoSpread.SSShowData """ & iStrData & """" & vbCr
	Response.Write " 	.DbQueryOk				" & vbCr
	Response.Write " End With					" & vbCr
	Response.Write "</Script>					" & vbCr

End Sub
'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data
'============================================================================================================
Sub SubBizSaveMulti()

	'-------------------------------------------------------------------------------------------------
	'IMPORTS View 상수 
	'-------------------------------------------------------------------------------------------------
	Const A390_I3_org_change_id		= 0												'View Name : in_dept_cd b_acct_dept
	Const A390_I3_dept_cd			= 1

	Const A390_I4_temp_gl_no		= 0												'View Name : in_temp_gl a_temp_gl
	Const A390_I4_temp_gl_dt		= 1
	Const A390_I4_gl_type			= 2
	Const A390_I4_insrt_user_id		= 3
	Const A390_I4_updt_user_id		= 4
	Const A390_I4_issued_dt			= 5
	Const A390_I4_gl_input_type		= 6
	Const A390_I4_cr_amt			= 7
	Const A390_I4_cr_loc_amt		= 8
	Const A390_I4_dr_amt			= 9
	Const A390_I4_dr_loc_amt		= 10
	Const A390_I4_conf_fg			= 11
	Const A390_I4_insrt_dt			= 12
	Const A390_I4_updt_dt			= 13
	Const A390_I4_hq_brch_fg		= 14
	Const A390_I4_internal_cd		= 15
	Const A390_I4_temp_gl_desc		= 16 

	'Group Name : in_grp_temp_gl_item
	Const A390_IG1_I1_biz_area_cd	= 0												'View Name : in_grp_item b_biz_area
	Const A390_IG1_I2_select_char	= 1												'View Name : in_grp_item ief_supplied
	Const A390_IG1_I3_org_change_id = 2												'View Name : in_grp_item b_acct_dept
	Const A390_IG1_I3_dept_cd		= 3
	Const A390_IG1_I4_acct_cd		= 4												'View Name : in_grp_item a_acct
	Const A390_IG1_I5_item_seq		= 5												'View Name : in_grp_item a_temp_gl_item
	Const A390_IG1_I5_dr_cr_fg		= 6
	Const A390_IG1_I5_doc_cur		= 7
	Const A390_IG1_I5_xch_rate		= 8
	Const A390_IG1_I5_vat_type		= 9
	Const A390_IG1_I5_item_amt		= 10
	Const A390_IG1_I5_item_loc_amt	= 11
	Const A390_IG1_I5_item_desc		= 12
	Const A390_IG1_I5_insrt_user_id = 13
	Const A390_IG1_I5_updt_user_id	= 14

	'Group Name : in_grp_temp_gl_dtl
	Const A390_IG2_I1_select_char	= 0												'View Name : in_grp_dtl ief_supplied
	Const A390_IG2_I2_item_seq		= 1												'View Name : in_grp_dtl a_temp_gl_item
	Const A390_IG2_I3_ctrl_cd		= 2												'View Name : in_grp_dtl a_ctrl_item
	Const A390_IG2_I4_dtl_seq		= 3												'View Name : in_grp_dtl a_temp_gl_dtl
	Const A390_IG2_I4_ctrl_val		= 4
	Const A390_IG2_I4_insrt_user_id = 5
	Const A390_IG2_I4_insrt_dt		= 6
	Const A390_IG2_I4_updt_user_id	= 7
	Const A390_IG2_I4_updt_dt		= 8

	Dim obj3PADG010
	Dim iCommandSent
'	Dim I1_b_biz_area_biz_area_cd
	Dim I2_b_currency_currency
	Dim I3_b_acct_dept
	Dim I4_a_temp_gl
	Dim IG1_in_grp_temp_gl_item
	Dim IG2_in_grp_temp_gl_dtl
	Dim E1_b_auto_numbering_auto_no
	Dim E2_b_auto_numbering_auto_no
	Dim strDoc_cur

	Dim LngRow
	Dim LngMaxRow
	Dim LngMaxRow3
	Dim arrTemp
	Dim arrVal
	Dim strStatus
	Dim strTempGlNo
	Dim temp

	On Error Resume Next
	Err.Clear

	LngMaxRow	= UNIConvNum(Request("txtMaxRows"),0)
	LngMaxRow3	= UNIConvNum(Request("txtMaxRows3"),0)

	Redim I3_b_acct_dept(A390_I3_dept_cd)
	Redim I4_a_temp_gl(A390_I4_temp_gl_desc)
	Redim IG1_in_grp_temp_gl_item(LngMaxRow - 1, A390_IG1_I5_updt_user_id)
	Redim IG2_in_grp_temp_gl_dtl(LngMaxRow3 - 1, A390_IG2_I4_updt_dt)

	iCommandSent							= Request("txtCommandMode")
	I2_b_currency_currency					= Request("txtgCurrency")
	I3_b_acct_dept(A390_I3_org_change_id)	= Request("txtOrgChangeId")
	I3_b_acct_dept(A390_I3_dept_cd)			= UCase(Request("txtDeptCd"))
	I4_a_temp_gl(A390_I4_temp_gl_no)		= UCase(Request("txttempglno"))
	I4_a_temp_gl(A390_I4_temp_gl_dt)		= UNIConvDateCompanyToDB(Request("txtGLDt"),NULL)
	I4_a_temp_gl(A390_I4_gl_type)			= "03"
	I4_a_temp_gl(A390_I4_insrt_user_id)		= Request("txtInsrtUserId")
	I4_a_temp_gl(A390_I4_updt_user_id)		= Request("txtUpdtUserId")
	I4_a_temp_gl(A390_I4_temp_gl_desc)		= Request("txtDesc")
	
	strDoc_cur		= Request("txtDocCur")
	
	arrTemp = Split(Request("txtSpread"), gRowSep)													'ITEM SPREAD

    For LngRow = 0 to LngMaxRow - 1

		arrVal = Split(arrTemp(LngRow), gColSep)

		strStatus = arrVal(0)

		IG1_in_grp_temp_gl_item(LngRow, A390_IG1_I2_select_char) = strStatus						'CRUD 구분 

		If  UCase(Trim(Request("txtCommandMode"))) = "UPDATE" Then
			strStatus = "C"
	    End If

		Select Case UCase(Trim(strStatus))
			Case "C","U"
		    	IG1_in_grp_temp_gl_item(LngRow, A390_IG1_I5_item_seq)			= Cint(arrVal(2))	'ItemSEQ  * Key
		    	IG1_in_grp_temp_gl_item(LngRow, A390_IG1_I1_biz_area_cd)		= arrVal(3)
		    	IG1_in_grp_temp_gl_item(LngRow, A390_IG1_I3_dept_cd)			= arrVal(4)			'부서	
		    	IG1_in_grp_temp_gl_item(LngRow, A390_IG1_I4_acct_cd)			= arrVal(5)			'계정코드 
			    IG1_in_grp_temp_gl_item(LngRow, A390_IG1_I5_dr_cr_fg)			= arrVal(6)			'차대구분 
		    	IG1_in_grp_temp_gl_item(LngRow, A390_IG1_I5_doc_cur)			= strDoc_cur  ' 거래 통화 

				If Trim(arrVal(7)) <> "" Then
		    		IG1_in_grp_temp_gl_item(LngRow, A390_IG1_I5_item_amt)		= UNIConvNum(arrVal(7),0)	'거래금액 
				End If
				If Trim(arrVal(8)) <> "" Then
					IG1_in_grp_temp_gl_item(LngRow, A390_IG1_I5_item_loc_amt)	= UNIConvNum(arrVal(8),0)	'자국금액 
				End If

                               
			        IG1_in_grp_temp_gl_item(LngRow, A390_IG1_I5_item_desc)			= arrVal(9)			'비고 


				If Trim(arrVal(10)) <> "" Then
					IG1_in_grp_temp_gl_item(LngRow, A390_IG1_I5_xch_rate)		= UNIConvNum(arrVal(10),0)	'환율 
				Else 
					IG1_in_grp_temp_gl_item(LngRow, A390_IG1_I5_xch_rate)		= 0                        	'환율 
				End If

			Case "D"
				IG1_in_grp_temp_gl_item(LngRow, A390_IG1_I5_item_seq)			= Cint(arrVal(2))	'ItemSEQ  * Key
		End Select

		IG1_in_grp_temp_gl_item(LngRow, A390_IG1_I5_doc_cur)					= UCase(Request("txtDocCur"))
   		IG1_in_grp_temp_gl_item(LngRow, A390_IG1_I3_org_change_id)				= Request("txtOrgChangeId")
		IG1_in_grp_temp_gl_item(LngRow, A390_IG1_I5_insrt_user_id)				= Request("txtInsrtUserId")
		IG1_in_grp_temp_gl_item(LngRow, A390_IG1_I5_updt_user_id)				= Request("txtUpdtUserId")

		Erase arrVal

   	Next


   	arrTemp = Split(Request("txtSpread3"), gRowSep)													'DTL SPREAD3
    
	For LngRow = 0 to LngMaxRow3 - 1

		arrVal = Split(arrTemp(LngRow), gColSep)

			strStatus = arrVal(0)																	'Row 의 상태 
	 		IG2_in_grp_temp_gl_dtl(LngRow, A390_IG2_I1_select_char) = strStatus						'CRUD 구분 
			
			Select Case strStatus
				Case "C","U"
				    IG2_in_grp_temp_gl_dtl(LngRow, A390_IG2_I2_item_seq)		= Cint(arrVal(1))
				    IG2_in_grp_temp_gl_dtl(LngRow, A390_IG2_I4_dtl_seq)			= Cint(arrVal(2))
				    IG2_in_grp_temp_gl_dtl(LngRow, A390_IG2_I3_ctrl_cd)			= arrVal(3)
					IG2_in_grp_temp_gl_dtl(LngRow, A390_IG2_I4_ctrl_val)		= arrVal(4)
					IG2_in_grp_temp_gl_dtl(LngRow, A390_IG2_I4_insrt_user_id)	= Request("txtInsrtUserId")
				    IG2_in_grp_temp_gl_dtl(LngRow, A390_IG2_I4_updt_user_id)	= Request("txtUpdtUserId")
				Case "D"
					IG2_in_grp_temp_gl_dtl(LngRow, A390_IG2_I2_item_seq)		= Cint(arrVal(1))
   Call SvrMsgBox(IG2_in_grp_temp_gl_dtl(LngRow, A390_IG2_I2_item_seq), vbInformation, I_MKSCRIPT)
					
					IG2_in_grp_temp_gl_dtl(LngRow, A390_IG2_I4_dtl_seq)			= Cint(arrVal(2))
					IG2_in_grp_temp_gl_dtl(LngRow, A390_IG2_I4_insrt_user_id)	= Request("txtInsrtUserId")
				   	IG2_in_grp_temp_gl_dtl(LngRow, A390_IG2_I4_updt_user_id)	= Request("txtUpdtUserId")
			End Select

			Erase arrVal
   	Next

	Set obj3PADG010 = CreateObject("PADG010.cAMngTmpGlHqSvr")

    If CheckSYSTEMError(Err, True) = True Then
       Exit Sub
    End If


	
	

	Call obj3PADG010.A_MANAGE_TEMP_GL_HQ_SVR(gStrGloBalCollection, iCommandSent, , I2_b_currency_currency, _
										 I3_b_acct_dept, I4_a_temp_gl, IG1_in_grp_temp_gl_item, IG2_in_grp_temp_gl_dtl, _
										 E1_b_auto_numbering_auto_no, E2_b_auto_numbering_auto_no)

	If CheckSYSTEMError(Err, True) = True Then
       Set obj3PADG010 = Nothing
		Exit Sub
    End If

	Set obj3PADG010 = nothing

	strTempGlNo = E1_b_auto_numbering_auto_no

	Response.Write "<Script Language=vbscript>				" & vbcr
	Response.Write " With Parent							" & vbCr
	Response.Write " 	.DbSaveOk """ & strTempGlNo & """	" & vbCr
	Response.Write " End With								" & vbCr
	Response.Write "</Script>								" & vbCr

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
Sub SubBizSaveMultiUpdate()

	Const A394_I2_a_temp_gl_temp_gl_no = 0
    Const A394_I2_a_temp_gl_temp_gl_dt = 1
    Const A394_I2_a_temp_gl_org_change_id = 2
    Const A394_I2_a_temp_gl_dept_cd = 3
    Const A394_I2_a_temp_gl_gl_type = 4
    Const A394_I2_a_temp_gl_gl_input_type = 5
    Const A394_I2_a_temp_gl_temp_gl_desc = 6
    Const A394_I2_a_temp_gl_project_no = 7


	Dim PAGG125_cAMngTmpGlUpdSvr
	Dim iCommandSent
	Dim I1_b_currency
	Dim I2_a_gl
	Dim txtSpread
	Dim txtSpread3
	Dim iStrRetTempGlNo
	
	Dim iLngMaxRow
	Dim iLngMaxRow3
	Dim iLngRow
	Dim iArrTemp1
	Dim iArrTemp2
	
	Dim zDataAuth
	
	Err.Clear
	On Error Resume Next
	
	
	iLngMaxRow  = CInt(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수 
    iLngMaxRow3 = CInt(Request("txtMaxRows3"))
	
	'--------------------------------------------------------------------
	'A_TEP_GL에 대한 정보  Setting
	'--------------------------------------------------------------------
	iCommandSent = Request("txtCommandMode")	'Spread Sheet 내용을 담고 있는 Element명 
	I1_b_currency = Request("txtgCurrency")

    ReDim I2_a_gl(7)
	I2_a_gl(A394_I2_a_temp_gl_temp_gl_no)		= UCase(Trim(Request("txtTempGlNo")))
	I2_a_gl(A394_I2_a_temp_gl_temp_gl_dt)		= UNIConvDateCompanyToDB(Request("txtGLDt"),NULL) 
	I2_a_gl(A394_I2_a_temp_gl_org_change_id)	= Trim(Request("hOrgChangeId"))
	I2_a_gl(A394_I2_a_temp_gl_dept_cd)			= UCase(Trim(Request("txtDeptCd")))
	I2_a_gl(A394_I2_a_temp_gl_gl_type)			= "03"  ' Trim(Request("cboGlType")) 
	I2_a_gl(A394_I2_a_temp_gl_gl_input_type)	= "HQ"  'Trim(Request("cboGlInputType"))      
	I2_a_gl(A394_I2_a_temp_gl_temp_gl_desc)		= Request("txtDesc")
	
	'--------------------------------------------------------------------
	'A_TEMP_GL_ITEM에 대한 정보  Setting
	'--------------------------------------------------------------------
	
    iArrTemp1 = Split(Request("txtSpread"), gRowSep)	'ITEM SPREAD
    
   	For iLngRow = 1 To iLngMaxRow    	
   	   	 
		iArrTemp2 = Split(iArrTemp1(iLngRow-1), gColSep)
				
        txtSpread = txtSpread & "C" & gColSep
		' 전체 삭제후 생성될 내용(Update, Insert) 만 전달		

		If iArrTemp2(0) <> "D" Then

		    txtSpread = txtSpread & Cint(iArrTemp2(1))													& gColSep					'current row
		    txtSpread = txtSpread & Cint(iArrTemp2(2))													& gColSep					'ItemSEQ  * Key
			txtSpread = txtSpread & iArrTemp2(5)														& gColSep					'계정코드 
			txtSpread = txtSpread & iArrTemp2(6)														& gColSep					'차대구분 
			txtSpread = txtSpread & Request("hOrgChangeId")												& gColSep			
			txtSpread = txtSpread & iArrTemp2(4)						& gColSep					'부서	
			txtSpread = txtSpread & UCase(Trim(Request("txtDocCur")))	& gColSep					'거래통화 

			If Trim(iArrTemp2(10)) = "" then
				txtSpread = txtSpread & ""																& gColSep					'환율 
			Else
				txtSpread = txtSpread & UNIConvNum(iArrTemp2(10),0)												& gColSep				
			End if
			
			txtSpread = txtSpread & ""																	& gColSep					'부가세 type			
    
     	    If Trim(iArrTemp2(7)) = "" then																				'거래금액 
				txtSpread = txtSpread & ""																& gColSep								
			Else
				txtSpread = txtSpread & UNIConvNum(iArrTemp2(7),0)												& gColSep				
			End if		
				
			If Trim(iArrTemp2(7)) = "" then																				'자국금액 
				txtSpread = txtSpread & ""																& gColSep								
			Else
				txtSpread = txtSpread & UNIConvNum(iArrTemp2(8),0)											& gColSep				
			End if
			
                '        If iArrTemp2(6) = "DR" then
                  '      txtSpread = txtSpread & Trim(Request("txtDesc"))                                                       & gRowSep
                 '       Else
			  txtSpread = txtSpread & iArrTemp2(9)									& gRowSep	' 비고
                 '       End If		


		
 
		End If

		
	Next   
		
	If CheckSYSTEMError(Err, True) = True Then
       Exit Sub
    End If



    '--------------------------------------------------------------------
	'A_TEMP_GL_DTL에 대한 정보  Setting
	'--------------------------------------------------------------------	
	
	iArrTemp1 = Split(Request("txtSpread3"), gRowSep)
	For iLngRow = 1 to iLngMaxRow3
	
		iArrTemp2 = Split(iArrTemp1(iLngRow-1), gColSep)
		txtSpread3 = txtSpread3 & "C" & gColSep
		If iArrTemp2(0) <> "D" Then
		        txtSpread3 = txtSpread3 & Cint(iArrTemp2(1))	& gColSep
		        txtSpread3 = txtSpread3 & Cint(iArrTemp2(2))	& gColSep
		        txtSpread3 = txtSpread3 & Trim(iArrTemp2(3))	& gColSep
		    	txtSpread3 = txtSpread3 & UCase(iArrTemp2(4))	& gRowSep		    
		End If
		
   	Next
	
	'--------------------------------------------------------------------
	'실행하기.
	'--------------------------------------------------------------------
	
	Set PAGG125_cAMngTmpGlUpdSvr = CreateObject("PAGG125.cAMngTmpGlUpdSvr")
		
	If CheckSYSTEMError(Err, True) = True Then
       Exit Sub
    End If
				
	iStrRetTempGlNo = PAGG125_cAMngTmpGlUpdSvr.A_MANAGE_TEMP_GL_UPDATE_SVR(gStrGlobalCollection, iCommandSent, I1_b_currency, I2_a_gl, txtSpread, txtSpread3) 
	
'	if err.number <> 0 then
'		Response.Write "xx  "
'		Response.Write err.description & " :: " & err.source
'		Set PAGG125_cAMngTmpGlUpdSvr  = Nothing
'		Response.End
'	end if
	

	Response.Write "  err  : " & err.Description

	If CheckSYSTEMError(Err, True) = True Then		
       Set PAGG125_cAMngTmpGlUpdSvr = Nothing
       Exit Sub
    End If
    
    Set PAGG125_cAMngTmpGlUpdSvr  = Nothing

	Response.Write " <Script Language=vbscript>										    " & vbCr
	Response.Write " With parent														" & vbCr
    Response.Write "	.DbSaveOk """ & ConvSPChars(iStrRetTempGlNo)	&			 """" & vbCr    
    Response.Write " End With															" & vbCr
    Response.Write " </Script>															" & vbCr


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
%>
