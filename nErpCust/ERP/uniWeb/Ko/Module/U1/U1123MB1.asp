<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m4141mb1
'*  4. Program Name         : 구매반품등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2003-06-04
'*  9. Modifier (First)     : Shin Jin-hyun
'* 10. Modifier (Last)      : Kim Jin Ha
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'* 14. Business Logic of m4111ma1(구매반품등록)
'**********************************************************************************************
	
On Error Resume Next
Err.Clear
	
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE", "MB")
Call LoadBNumericFormatB("I", "*","NOCOOKIE", "MB")
Call HideStatusWnd

'----
Dim strMaxRows
strMaxRows = Request("txtMaxRows")
Dim strMvmtNo
'----

lgOpModeCRUD	=	Request("txtMode")	'☜: Read Operation Mode (CRUD)

Call SubOpenDB(lgObjConn)
			
Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)
             Call SubBizSaveMulti()
        Case CStr ("changeMvmtType")
			 Call SubchangeMvmtType()
		Case "changeSpplCd" 
			 Call DisplaySupplierNm(Request("txtSupplierCd"))	 
		Case "changeGroupCd" 
			 Call DisplayGroupNm(Request("txtGroupCd"))	 				 
End Select

Call SubCLOSEDB(lgObjConn)
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()														'☜: 현재 조회/Prev/Next 요청을 받음 
 
	Dim iPM7G428
	
	Dim istrData
	Dim StrNextKey		' 다음 값 
	Dim lgStrPrevKey	' 이전 값 
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
	Dim GroupCount          
	Dim index,Count     ' 저장 후 Return 해줄 값을 넣을때 쓴는 변수 
	 
	Dim TmpBuffer
	Dim iMax
	Dim iIntLoopCount
	Dim iTotalStr
	    
	Dim I1_m_pur_goods_mvmt_no
	Dim I2_m_pur_goods_mvmt_rcpt_no
	 
	Dim EG1_export_group
	Const M744_EG1_E1_minor_cd = 0
	Const M744_EG1_E1_minor_nm = 1
	Const M744_EG1_E2_lc_no = 2
	Const M744_EG1_E3_lc_seq = 3 
	Const M744_EG1_E4_cc_no = 4
	Const M744_EG1_E5_cc_seq = 5
	Const M744_EG1_E6_minor_nm = 6  
	Const M744_EG1_E7_minor_nm = 7 
	Const M744_EG1_E8_plant_cd = 8
	Const M744_EG1_E8_plant_nm = 9
	Const M744_EG1_E9_po_seq_no = 10 
	Const M744_EG1_E10_po_no = 11
	Const M744_EG1_E11_mvmt_no = 12 
	Const M744_EG1_E11_mvmt_dt = 13
	Const M744_EG1_E11_mvmt_sl_cd = 14
	Const M744_EG1_E11_lot_no = 15
	Const M744_EG1_E11_lot_sub_no = 16
	Const M744_EG1_E11_mvmt_unit = 17
	Const M744_EG1_E11_mvmt_qty = 18
	Const M744_EG1_E11_gm_no = 19
	Const M744_EG1_E11_gm_seq_no = 20
	Const M744_EG1_E11_gm_sub_seq_no = 21
	Const M744_EG1_E11_gm_year = 22
	Const M744_EG1_E11_mvmt_biz_area = 23
	Const M744_EG1_E11_mvmt_base_qty = 24
	Const M744_EG1_E11_mvmt_base_unit = 25
	Const M744_EG1_E11_fr_trans_coef = 26
	Const M744_EG1_E11_to_trans_coef = 27
	Const M744_EG1_E11_xch_rt = 28
	Const M744_EG1_E11_iv_qty = 29
	Const M744_EG1_E11_maker_lot_no = 30
	Const M744_EG1_E11_maker_lot_sub_no = 31
	Const M744_EG1_E11_distribt_amt = 32
	Const M744_EG1_E11_distribt_flg = 33
	Const M744_EG1_E11_mvmt_prc = 34
	Const M744_EG1_E11_mvmt_cur = 35
	Const M744_EG1_E11_mvmt_doc_amt = 36
	Const M744_EG1_E11_mvmt_loc_amt = 37
	Const M744_EG1_E11_applied_xch_type = 38
	Const M744_EG1_E11_cc_no_for_sales = 39
	Const M744_EG1_E11_cc_seq_for_sales = 40
	Const M744_EG1_E11_pur_org = 41
	Const M744_EG1_E11_mvmt_rcpt_no = 42
	Const M744_EG1_E11_mvmt_rcpt_dt = 43
	Const M744_EG1_E11_mvmt_rcpt_qty = 44
	Const M744_EG1_E11_mvmt_rcpt_unit = 45
	Const M744_EG1_E11_mvmt_rcpt_sl_cd = 46
	Const M744_EG1_E11_mvmt_method = 47
	Const M744_EG1_E11_inspect_req_no = 48
	Const M744_EG1_E11_inspect_sts = 49
	Const M744_EG1_E11_inspect_flg = 50
	Const M744_EG1_E11_inspect_good_qty = 51
	Const M744_EG1_E11_inspect_bad_qty = 52
	Const M744_EG1_E11_inspect_good_base_qty = 53
	Const M744_EG1_E11_inspect_bad_base_qty = 54
	Const M744_EG1_E11_inspect_result_no = 55
	Const M744_EG1_E11_inspect_result_dt = 56
	Const M744_EG1_E11_to_good_sl_cd = 57
	Const M744_EG1_E11_to_bad_sl_cd = 58
	Const M744_EG1_E11_ext1_cd = 59
	Const M744_EG1_E11_ext1_qty = 60
	Const M744_EG1_E11_ext1_amt = 61
	Const M744_EG1_E11_ext1_rt = 62
	Const M744_EG1_E11_ext2_cd = 63
	Const M744_EG1_E11_ext2_qty = 64
	Const M744_EG1_E11_ext2_amt = 65
	Const M744_EG1_E11_ext2_rt = 66
	Const M744_EG1_E11_ext3_cd = 67
	Const M744_EG1_E11_ext3_qty = 68
	Const M744_EG1_E11_ext3_amt = 69
	Const M744_EG1_E11_ext3_rt = 70
	Const M744_EG1_E11_ret_type = 71
	Const M744_EG1_E11_tracking_no = 72
	Const M744_EG1_E11_tot_ret_qty = 73
	Const M744_EG1_E11_ap_flg = 74
	Const M744_EG1_E11_iv_no = 75
	Const M744_EG1_E11_iv_seq_no = 76
	Const M744_EG1_E12_item_cd = 77
	Const M744_EG1_E12_item_nm = 78
	Const M744_EG1_E12_spec = 79
	Const M744_EG1_E13_sl_cd = 80
	Const M744_EG1_E13_sl_nm = 81
	'add 2002-11-14 JYYOON
	Const M743_EG1_E14_io_type_cd = 82
	Const M743_EG1_E14_io_type_nm = 83
	Const M743_EG1_E15_mvmt_no = 84
	Const M743_EG1_E16_bp_cd = 85
	Const M743_EG1_E16_bp_nm = 86
	Const M743_EG1_E17_pur_grp = 87
	Const M743_EG1_E17_pur_grp_nm = 88
	Const M743_EG1_E17_ret_ord_qty = 89
	 
	Dim strGmNo, strTempGlNo, strGlNo 
    
	Dim I4_scm_flag
	
    Const C_SHEETMAXROWS_D  = 100
    
    On Error Resume Next 
	Err.Clear                                                               '☜: Protect system from crashing
    
	lgStrPrevKey = Request("lgStrPrevKey")
  
    Set iPM7G428 = Server.CreateObject("PM7G428.cMListPurRcptS")    

    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
	End If
	
    I2_m_pur_goods_mvmt_rcpt_no  		= Request("txtGrNo")
    
    if trim(lgStrPrevKey) <> "" then
		I1_m_pur_goods_mvmt_no  	= lgStrPrevKey
	End If
	
	I4_scm_flag = "Y"
    
    Call iPM7G428.M_LIST_PUR_RCPT_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_m_pur_goods_mvmt_no, I2_m_pur_goods_mvmt_rcpt_no, EG1_export_group, "", I4_scm_flag)
	
	If CheckSYSTEMError2(Err,True,"","","","","") = true then 
		Set iM41218 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "	Parent.FncNew() " & vbCr
		Response.Write "	parent.frm1.txtGrNo.Value 	= """ & ConvSPChars(Request("txtGrNo")) & """" & vbCr
		Response.Write "</Script>" & vbCr
		Exit Sub
	End If

	iLngMaxRow = Request("txtMaxRows")											'Save previous Maxrow                                                
    GroupCount = UBound(EG1_export_group,1)
    
	
	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "With parent" & vbCr
    Response.Write ".frm1.txtMvmtType.Value      = """ & ConvSPChars(EG1_export_group(GroupCount,M743_EG1_E14_io_type_cd)) & """"  & vbCr
    Response.Write ".frm1.txtMvmtTypeNm.Value    = """ & ConvSPChars(EG1_export_group(GroupCount,M743_EG1_E14_io_type_nm)) & """"  & vbCr       	   		
    Response.Write ".frm1.txtGmDt.text           = """ & UNIDateClientFormat(EG1_export_group(GroupCount,M744_EG1_E11_mvmt_rcpt_dt)) & """" & vbCr
    Response.Write ".frm1.txtGroupCd.Value       = """ & ConvSPChars(EG1_export_group(GroupCount,M743_EG1_E17_pur_grp)) & """"       & vbCr
    Response.Write ".frm1.txtGroupNm.Value       = """ & ConvSPChars(EG1_export_group(GroupCount,M743_EG1_E17_pur_grp_nm)) & """"    & vbCr
    Response.Write ".frm1.txtSupplierCd.Value    = """ & ConvSPChars(EG1_export_group(GroupCount,M743_EG1_E16_bp_cd)) & """"     & vbCr
    Response.Write ".frm1.txtSupplierNm.Value    = """ & ConvSPChars(EG1_export_group(GroupCount,M743_EG1_E16_bp_nm)) & """"     & vbCr
    Response.Write ".frm1.txtGrNo.Value          = """ & ConvSPChars(I2_m_pur_goods_mvmt_rcpt_no) & """"         & vbCr
	Response.Write "End With"  & vbCr
	Response.Write "</Script>"	   & vbCr
     
    iIntLoopCount = 0
	iMax = UBound(EG1_export_group,1)
	ReDim TmpBuffer(iMax)
		
	For iLngRow = 0 To iMax
		If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
		   StrNextKey = ConvSPChars(EG1_export_group(iLngRow, M744_EG1_E11_mvmt_no)) 
           Exit For
        End If  
		
		istrData = ""
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E8_plant_cd))				'입고Seq
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E8_plant_nm))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E12_item_cd))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E12_item_nm))	
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E12_spec))	
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_tracking_no))	
                	
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,M744_EG1_E11_mvmt_rcpt_qty),ggQty.DecPoint,0)
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_mvmt_rcpt_unit))  
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E13_sl_cd))
        istrData = istrData & Chr(11) & "" 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E13_sl_nm))

        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_lot_no))
        istrData = istrData & Chr(11) & "" 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_lot_sub_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E1_minor_cd))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E1_minor_nm))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E10_po_no)) 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E9_po_seq_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_gm_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_gm_seq_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_mvmt_no))
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,M743_EG1_E17_ret_ord_qty),ggQty.DecPoint,0)
                 
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
        istrData = istrData & Chr(11) & Chr(12)
        
        If strGmNo = "" Then
            strGmNo = Trim(EG1_export_group(iLngRow,M744_EG1_E11_gm_no))
        End if
        
        TmpBuffer(iIntLoopCount) = istrData
        iIntLoopCount = iIntLoopCount + 1
    Next
	
	iTotalStr = Join(TmpBuffer, "")    

	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "With parent" & vbCr
	Response.Write "	.ggoSpread.Source 		= .frm1.vspdData" & vbCr
    Response.Write "	.ggoSpread.SSShowData        """ & iTotalStr	& """" & vbCr	   
	Response.Write "	.lgStrPrevKey              = """ & StrNextKey   & """" & vbCr 
    Response.Write "	.frm1.hdnGrNo.value = """ & ConvSPChars(Request("txtGrNo")) & """" & vbCr
    Response.Write "	.frm1.txtGrNo1.value = """ & ConvSPChars(Request("txtGrNo")) & """" & vbCr
    Response.write "	.DbQueryOk " & vbCr
    Response.Write "End With"  & vbCr
	Response.Write "</Script>"	   & vbCr
	
	set iPM7G428 = Nothing
	
	'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	'******* 전표No. 만들기. *********
	'*************************************************************************************************************

    If  strGmNo <> "" Or strGmNo <> Null then  

		Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
				
		lgStrSQL = "SELECT document_year FROM i_goods_movement_header " 
		lgStrSQL = lgStrSQL & " WHERE item_document_no = '" & strGmNo & "'"		
		If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False then
			Response.Write "<Script Language=VBScript>" & vbCr
			Response.Write "parent.frm1.hdnGlType.value	=	""B"" " & vbCr
			Response.Write "</Script>"					& vbCr
			
			Call SubCloseRs(lgObjRs)  
			Call SubCloseDB(lgObjConn)                                                        '☜: Make a DB Connection
			Exit Sub
		End if
		'A_GL.Ref_no
		strGmNo	=	strGmNo & "-" & lgObjRs("document_year")
							
		'수정(전표조회 추가)
		Response.Write "<Script Language=VBScript>" & vbCr
		
		lgStrSQL = "select temp_gl_no,gl_no from ufn_a_GetGlNo( '" & strGmNo & "') " 		
				
		IF FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False then
			Response.Write "parent.frm1.hdnGlType.value	=	""B""	  " & vbCr
			Response.Write "parent.frm1.hdnGlNo.value	=	""""      " & vbCr  
		Else
			strTempGlNo = lgObjRs("temp_gl_no")
			strGlNo = lgObjRs("gl_no")
			
			If ConvSPChars(Trim(strGlNo)) = "" And ConvSPChars(Trim(strTempGlNo)) <> "" Then
				Response.Write "parent.frm1.hdnGlType.value	=	""T""	  " & vbCr
				Response.Write "parent.frm1.hdnGlNo.value	=	""" & lgObjRs("temp_gl_no")& """" & vbCr  
			Else
				Response.Write "parent.frm1.hdnGlType.value	=	""A""	  " & vbCr
				Response.Write "parent.frm1.hdnGlNo.value	=	""" & lgObjRs("gl_no")& """" & vbCr  
			End If
		End If
		Response.Write "</Script>"					& vbCr
		Call SubCloseRs(lgObjRs)
		Call SubCloseDB(lgObjConn)
	Else
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "parent.frm1.hdnGlType.value	=	""B""	  " & vbCr
		Response.Write "</Script>"					& vbCr
	End If
	
End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data into Db
'============================================================================================================
Sub SubBizSaveMulti()														'☜: 저장 요청을 받음 

	Dim iPM7G421R
	Dim iCommandSent
	Dim iErrorPosition
	
	Dim itxtSpread , itxtSpread2
	Dim itxtSpreadArr
    Dim itxtSpreadArrCount
    Dim iCUCount
    Dim iDCount
    Dim ii
     
	Dim arrVal, arrTemp, strStatus		
	 
	Dim I1_ief_supplied_select_char	

	Dim I2_m_pur_goods_mvmt
	Const M686_I2_mvmt_rcpt_no = 0
	Const M686_I2_mvmt_rcpt_dt = 1
	Const M686_I2_gm_no = 2
	Const M686_I2_gm_seq_no = 3
	Redim  I2_m_pur_goods_mvmt(M686_I2_gm_seq_no)
	 
	Dim I3_m_mvmt_type_io_type_cd
	Dim I4_b_biz_partner_bp_cd
	Dim I5_b_pur_grp

	Dim E5_b_auto_numbering
	Dim iStrSpread
	Dim iLngMaxRow
  
    On Error Resume Next 
    Err.Clear																		'☜: Protect system from crashing

	Call RemovedivTextArea()
	
	If Len(Request("txtGmDt")) Then
		If UNIConvDate(Request("txtGmDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Exit Sub	
		End If
	End If
	
	itxtSpread = ""
	itxtSpread2 = REQUEST("txtSpread2")
    
    iCUCount = Request.Form("txtCUSpread").Count
    iDCount  = Request.Form("txtDSpread").Count
    
    itxtSpreadArrCount = -1
    
    ReDim itxtSpreadArr(iCUCount + iDCount)
             
    For ii = 1 To iDCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtDSpread")(ii)
    Next
    
    For ii = 1 To iCUCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
    Next
   
    itxtSpread = Join(itxtSpreadArr,"")
    
	lgIntFlgMode = CInt(Request("txtFlgMode"))	
	iLngMaxRow = CInt(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수 
	
	I2_m_pur_goods_mvmt(M686_I2_mvmt_rcpt_no) = Request("txtGrNo1")
    I2_m_pur_goods_mvmt(M686_I2_mvmt_rcpt_dt) = UNIConvDate(Request("txtGmDt"))
    I2_m_pur_goods_mvmt(M686_I2_gm_no)        = ""
    I3_m_mvmt_type_io_type_cd                 = UCase(Request("txtMvmtType"))
    I4_b_biz_partner_bp_cd                    = Trim(Request("txtSupplierCd"))
    I5_b_pur_grp                              = Trim(Request("txtGroupCd"))
    
    arrTemp = Split(itxtSpread, gRowSep)	
    arrVal = Split(arrTemp(0), gColSep)
	    	
    strStatus = arrVal(0)
    
    if strStatus = "C" then
       iCommandSent								= "CREATE"
    else 
       iCommandSent								= "DELETE"		
    end if
    
    Set iPM7G421R = Server.CreateObject("PM7G41R.cMMaintPurReturnS")    

    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
	End If

    Call iPM7G421R.M_MAINT_PUR_RETURN_SVR(gStrGlobalCollection, iCommandSent,I2_m_pur_goods_mvmt, _
                                Ucase(I3_m_mvmt_type_io_type_cd), Ucase(I4_b_biz_partner_bp_cd), Ucase(I5_b_pur_grp), _
                                itxtSpread, E5_b_auto_numbering, iErrorPosition)
	
	If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
	   Set iPM7G421R = Nothing
	   Exit Sub
	End If		

    Set iPM7G421R = Nothing                                                   '☜: Unload Comproxy  
    
    strMvmtNo = UCase(ConvSPChars(E5_b_auto_numbering))
    
    Dim lsfield , lsfrom , lswhere
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
    lsfield = " MVMT_NO "  
    lsfrom	= " M_PUR_GOODS_MVMT "
    lswhere	= " MVMT_RCPT_NO = '" & strMvmtNo & "'"
    
    call CommonQueryRs(lsfield , lsfrom , lswhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    Dim strREALMVMTNO
    strREALMVMTNO = SPLIT(lgF0,chr(11))
    
    if iCommandSent = "CREATE" then
    	Call SubBizSaveMulti2(itxtSpread, strREALMVMTNO , itxtSpread2)
    end if
    
   	Response.Write "<Script language=vbs> " & vbCr 
	Response.Write "With parent " & vbCr
	Response.Write "	If """ & lgIntFlgMode & """ = """ & OPMD_CMODE & """ Then " & vbCr
	Response.Write "		.frm1.txtGrNo.value    = """ & UCase(ConvSPChars(E5_b_auto_numbering))  & """" & vbCr'
	Response.Write "	End If"				& vbCr	
    Response.Write " .DbSaveOk "      & vbCr						
    Response.Write "End With " & vbCr
    Response.Write "</Script> "    
  
End Sub


Sub SubBizSaveMulti2(itxtSpread, strREALMVMTNO , itxtSpread2)

    Dim arrRowVal
    Dim arrColVal
    Dim arrRowVal2
    Dim arrColVal2
    Dim iDx
	Dim strXX
	
    On Error Resume Next
    Err.Clear
    
	arrRowVal  = Split(itxtSpread, gRowSep)
	arrRowVal2 = Split(itxtSpread2, gRowSep)
	
    For iDx = 1 To Cint(strMaxRows)     'lgLngMaxRow
        arrColVal  = Split(arrRowVal(iDx-1), gColSep)
		arrColVal2 = Split(arrRowVal2(iDx-1), gColSep)
		strXX = strrealMVMTNO(iDx-1)
		
        Select Case arrColVal(0)
            Case "C"                            '☜: Create
                Call SubBizSaveMultiUpdate(arrColVal, strXX , arrColVal2)
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next

End Sub    

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal, strXX , arrColVal2)
    On Error Resume Next
    Err.Clear
    
    Dim lgStrSQL
    
    lgStrSQL = "UPDATE  M_SCM_FIRM_PUR_RCPT" & _
			   "   SET  MVMT_NO			= " & FilterVar(Trim(strXX),"","S")   & "," & _
			   "		RCPT_QTY		= " & FilterVar(Trim(UCase(arrColVal(07))),"","D") & "," & _
			   "		RCPT_DT			= " & FilterVar(Request("txtGmDt"),"","S") & "," & _
			   "		UPDT_USER_ID	= " & FilterVar(gUsrId,"","S")                      & "," & _
			   "		UPDT_DT			= GetDate() "  & _
			   " WHERE  PO_NO			= " & FilterVar(Trim(UCase(arrColVal(17))),"","S") & _
			   "   AND  PO_SEQ_NO		= " & FilterVar(Trim(UCase(arrColVal(18))),"","D") & _	
			   "   AND  SPLIT_SEQ_NO	= " & FilterVar(Trim(UCase(arrColVal2(0))),"","D")
			   
    'CALL SVRMSGBOX(lgStrSQL,0,1)
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	'Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
    
End Sub

'============================================================================================================
' Name : SubchangeMvmtType
' Desc : 
'============================================================================================================
Sub SubchangeMvmtType()	
  
  Dim iPM1G439
   
  Dim I1_m_mvmt_type_io_type_cd
  Dim I2_ief_supplied_SELECT_CHAR
  
  Dim E1_m_mvmt_type
  Const EA_m_mvmt_type_io_type_cd1 = 0
  Const EA_m_mvmt_type_io_type_nm1 = 1
  Const EA_m_mvmt_type_mvmt_cd1 = 2
  Const EA_m_mvmt_type_rcpt_flg1 = 3
  Const EA_m_mvmt_type_import_flg1 = 4
  Const EA_m_mvmt_type_ret_flg1 = 5
  Const EA_m_mvmt_type_subcontra_flg1 = 6
  Const EA_m_mvmt_type_usage_flg1 = 7
  Const EA_m_mvmt_type_ext1_cd1 = 8
  Const EA_m_mvmt_type_ext2_cd1 = 9
  Const EA_m_mvmt_type_ext3_cd1 = 10
  Const EA_m_mvmt_type_ext4_cd1 = 11

  Dim iErrorPosition
  
    On Error Resume Next
    Err.Clear								    '☜: Protect system	from crashing

    Set	iPM1G439 = CreateObject("PM1G439.cMLookupMvmtTypeS")

    If CheckSYSTEMError(Err,True) = True Then
			Exit Sub
	End If
    
    I1_m_mvmt_type_io_type_cd 	= ucase(Request("txtMvmtType"))
    I2_ief_supplied_SELECT_CHAR = "2"

    E1_m_mvmt_type = iPM1G439.M_LOOKUP_MVMT_TYPE_SVR(gStrGlobalCollection,I1_m_mvmt_type_io_type_cd,I2_ief_supplied_SELECT_CHAR)

    If CheckSYSTEMError2(Err, True, "","","","","") = True Then
			Set iPM1G439 = Nothing
			Response.Write "<Script Language=VBScript>" & vbCr
					Response.Write "With parent" & vbCr
					Response.Write ".frm1.txtMvmtTypeNm.Value 	= """"	" & vbCr
					Response.Write ".frm1.hdnImportflg.Value 	= """"	" & vbCr
					Response.Write ".frm1.hdnRcptflg.Value 		= """"	" & vbCr
					Response.Write ".frm1.hdnRetflg.Value 		= """"	" & vbCr
					Response.Write ".frm1.hdnSubcontraflg.Value = """"	" & vbCr
					Response.Write ".frm1.txtMvmtType.focus 	        " & vbCr
					Response.Write "End With"				& vbCr
				Response.Write "</Script>"					& vbCr
			Exit Sub
	End If

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr												'☜: 화면 처리 ASP 를 지칭함 
    Response.Write ".frm1.txtMvmtType.Value 	= """ & ConvSPChars(E1_m_mvmt_type(EA_m_mvmt_type_io_type_cd1))& """" & vbCr
	Response.Write ".frm1.txtMvmtTypeNm.Value 	= """ & ConvSPChars(E1_m_mvmt_type(EA_m_mvmt_type_io_type_nm1))& """" & vbCr
	Response.Write ".frm1.hdnImportflg.Value 	= """ & ConvSPChars(E1_m_mvmt_type(EA_m_mvmt_type_import_flg1))& """" & vbCr
	Response.Write ".frm1.hdnRcptflg.Value 		= """ & ConvSPChars(E1_m_mvmt_type(EA_m_mvmt_type_rcpt_flg1))& """" & vbCr
	Response.Write ".frm1.hdnRetflg.Value 		= """ & ConvSPChars(E1_m_mvmt_type(EA_m_mvmt_type_ret_flg1))& """" & vbCr
	Response.Write ".frm1.hdnSubcontraflg.Value = """ & ConvSPChars(E1_m_mvmt_type(EA_m_mvmt_type_subcontra_flg1))& """" & vbCr
						
	Response.Write "If """ & Request("txtRefABCflag")& """ = ""A"" then " & vbCr			
	Response.Write ".lgOpenFlag	= True" & vbCr
	Response.Write ".lgRefABCflag = """" " & vbCr
	Response.Write "Call parent.OpenPORef()	" & vbCr		
	Response.Write "Elseif """ & Request("txtRefABCflag")& """ = ""B"" Then " & vbCr
	Response.Write ".lgOpenFlag	= True" & vbCr
	Response.Write ".lgRefABCflag = """" " & vbCr
	Response.Write "Call parent.OpenCCRef()	" & vbCr		
	Response.Write "Elseif """ & Request("txtRefABCflag")& """ = ""C"" Then " & vbCr
	Response.Write ".lgOpenFlag	= True" & vbCr
	Response.Write ".lgRefABCflag = """" " & vbCr
	Response.Write "Call parent.OpenLLCRef() " & vbCr
	Response.Write "Elseif """ & Request("txtRefABCflag")& """ = ""D"" Then " & vbCr
	Response.Write ".lgOpenFlag	= True" & vbCr
	Response.Write ".lgRefABCflag = """" " & vbCr
	Response.Write "Call parent.OpenIvRef()	" & vbCr					
	Response.Write "End if" & vbCr
				
	Response.Write "End With"				& vbCr
	Response.Write "</Script>"					& vbCr

End Sub

'-----------------------
'Display CodeName
'2002/08/19 Kim Jin Ha
'-----------------------
Sub DisplaySupplierNm(inCode)
	
	On Error Resume Next						
    Err.Clear   
    
	Call SubOpenDB(lgObjConn)
	
	lgStrSQL = ""
	lgStrSQL = lgStrSQL & " SELECT BP_NM FROM B_BIZ_PARTNER " 
	lgStrSQL = lgStrSQL & " WHERE BP_CD = '" & inCode & "' AND Bp_Type in ('S','CS') AND usage_flag='Y' AND  in_out_flag = 'O' "											'사외거래처만"		
	
	If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X")  then
	
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "With parent.frm1" & vbCr
		Response.Write "	.txtSupplierNm.value	=	""" & lgObjRs("BP_NM") & """ " & vbCr
		Response.Write "End With"  & vbCr
		Response.Write "</Script>"  & vbCr
			
		Call SubCloseRs(lgObjRs)  
	Else
		Call DisplayMsgBox("179020", vbInformation, "", "", I_MKSCRIPT)
		
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "With parent.frm1" & vbCr
		Response.Write "	.txtSupplierNm.value	=	"""" " & vbCr
		Response.Write "	.txtSupplierCd.focus			 " & vbCr
		Response.Write "End With"  & vbCr
		Response.Write "</Script>"  & vbCr
		
	End if
	
	Call SubCloseDB(lgObjConn)
End Sub 

'-----------------------
'Display DisplayGroupNm
'2002/08/19 Kim Jin Ha
'-----------------------
Sub DisplayGroupNm(inCode)
	
	On Error Resume Next						
    Err.Clear   
    
	Call SubOpenDB(lgObjConn)
	
	lgStrSQL = ""
	lgStrSQL = lgStrSQL & " SELECT PUR_GRP_NM FROM B_PUR_GRP " 
	lgStrSQL = lgStrSQL & " WHERE PUR_GRP = '" & inCode & "' AND USAGE_FLG='Y'"		
	
	If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X")  then
	
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "With parent.frm1" & vbCr
		Response.Write "	.txtGroupNm.value	=	""" & lgObjRs("PUR_GRP_NM") & """ " & vbCr
		Response.Write "End With"  & vbCr
		Response.Write "</Script>"  & vbCr
			
		Call SubCloseRs(lgObjRs)  
	Else
		Call DisplayMsgBox("125100", vbInformation, "", "", I_MKSCRIPT)
		
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "With parent.frm1" & vbCr
		Response.Write "	.txtGroupNm.value	=	"""" " & vbCr
		Response.Write "	.txtGroupCd.focus			 " & vbCr
		Response.Write "End With"  & vbCr
		Response.Write "</Script>"  & vbCr
		
	End if
	
	Call SubCloseDB(lgObjConn)
End Sub 

'============================================================================================================
' Name : RemovedivTextArea
' Desc : 
'============================================================================================================
Sub RemovedivTextArea()
    On Error Resume Next                                                             
    Err.Clear                                                                        
	
	Response.Write "<Script language=vbs> " & vbCr   
    Response.Write "Parent.RemovedivTextArea "      & vbCr   
    Response.Write "</Script> "      & vbCr   
End Sub
%>
