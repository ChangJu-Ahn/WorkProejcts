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
'*  3. Program ID           : m4111mb1
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : iM41211(Maint)
'							  iPM7G428(List)
'*  7. Modified date(First) : 2003/06/03
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Kim Jin Ha
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  :
'**********************************************************************************************

On Error Resume Next
Err.Clear
	
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE", "MB")
Call LoadBNumericFormatB("I", "*","NOCOOKIE", "MB")
Call HideStatusWnd

lgOpModeCRUD	=	Request("txtMode")	'☜: Read Operation Mode (CRUD)

Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)
             Call SubBizSaveMulti()
End Select

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
 
	Dim iPM7G428
	Dim istrData
	Dim StrNextKey		' 다음 값 
	Dim lgStrPrevKey	' 이전 값 
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
	Dim GroupCount          
		
	Const C_SHEETMAXROWS_D  = 100
	
	Dim FlgData
	Dim I1_m_pur_goods_mvmt_no
	Dim I2_m_pur_goods_mvmt_rcpt_no
	' === 2005.07.14 Tracking No.9894 by MJG ======================================================
	Dim I3_m_pur_goods_flag
	' === 2005.07.14 Tracking No.9894 by MJG ======================================================
	 
	Dim EG1_export_group
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
	Const M744_EG1_E11_lot_no = 15
	Const M744_EG1_E11_lot_sub_no = 16
	Const M744_EG1_E11_mvmt_qty = 18
	Const M744_EG1_E11_gm_no = 19
	Const M744_EG1_E11_gm_seq_no = 20
	Const M744_EG1_E11_maker_lot_no = 30
	Const M744_EG1_E11_maker_lot_sub_no = 31
	Const M744_EG1_E11_mvmt_prc = 34
	Const M744_EG1_E11_mvmt_cur = 35
	Const M744_EG1_E11_mvmt_doc_amt = 36
	Const M744_EG1_E11_mvmt_loc_amt = 37
	Const M744_EG1_E11_mvmt_rcpt_dt = 43
	Const M744_EG1_E11_mvmt_rcpt_qty = 44
	Const M744_EG1_E11_mvmt_rcpt_unit = 45
	Const M744_EG1_E11_inspect_req_no = 48
	Const M744_EG1_E11_inspect_flg = 50
	Const M744_EG1_E11_inspect_result_no = 55
  Const M743_EG1_E11_ext2_cd = 63
	Const M744_EG1_E11_tracking_no = 72
	Const M744_EG1_E11_iv_no = 75
	Const M744_EG1_E11_iv_seq_no = 76
	Const M744_EG1_E12_item_cd = 77
	Const M744_EG1_E12_item_nm = 78
	Const M744_EG1_E12_spec = 79
	Const M744_EG1_E13_sl_cd = 80
	Const M744_EG1_E13_sl_nm = 81
	
	Const M743_EG1_E14_io_type_cd = 82
	Const M743_EG1_E14_io_type_nm = 83
	Const M743_EG1_E16_bp_cd = 85
	Const M743_EG1_E16_bp_nm = 86
	Const M743_EG1_E17_pur_grp = 87
	Const M743_EG1_E17_pur_grp_nm = 88
	Const M743_EG1_E17_ret_ord_qty = 89
	Const M743_EG1_E17_remark_hdr = 90
	Const M744_EG1_E11_remark_dtl = 91

	Dim strGmNo, strTempGlNo, strGlNo 
	Dim istrMvmtCur	
	
	Dim TmpBuffer
    Dim iMax
    Dim iIntLoopCount
    Dim iTotalStr
    
    On Error Resume Next 
    Err.Clear                                                               '☜: Protect system from crashing
    
	lgStrPrevKey = Request("lgStrPrevKey")
  
    Set iPM7G428 = Server.CreateObject("PM7G428.cMListPurRcptS")    

    If CheckSYSTEMError(Err,True) = True Then
       Set iPM7G428 = Nothing
	   Exit Sub
	End If
	
    I2_m_pur_goods_mvmt_rcpt_no  		= FilterVar(UCase(Trim(Request("txtMvmtNo"))),"","SNM")
    
    if Trim(lgStrPrevKey) <> "" then
	   I1_m_pur_goods_mvmt_no  	= lgStrPrevKey
	End if

	' === 2005.07.14 Tracking No.9894 by MJG ======================================================
	I3_m_pur_goods_flag = "M"
	
    Call iPM7G428.M_LIST_PUR_RCPT_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_m_pur_goods_mvmt_no, I2_m_pur_goods_mvmt_rcpt_no, EG1_export_group, I3_m_pur_goods_flag)
	' === 2005.07.14 Tracking No.9894 by MJG ======================================================
			
	If CheckSYSTEMError2(Err,True,"","","","","") = true then 
		Response.Write "<Script Language=VBScript>" & vbCr		
		Response.Write " call parent.ggoOper.ClearField(parent.Document, ""2"")   "                 & vbCr
		Response.Write "</Script>" & vbCr
		Set iPM7G428 = Nothing
		Exit Sub
	End If
			
    iLngMaxRow = Request("txtMaxRows")											'Save previous Maxrow                                                
    GroupCount = UBound(EG1_export_group,1)
    
	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "With parent"                 & vbCr
    Response.Write ".frm1.txtMvmtType.Value      = """ & ConvSPChars(EG1_export_group(GroupCount,M743_EG1_E14_io_type_cd)) & """"			& vbCr
    Response.Write ".frm1.txtMvmtTypeNm.Value    = """ & ConvSPChars(EG1_export_group(GroupCount,M743_EG1_E14_io_type_nm)) & """"			& vbCr       	   		
    Response.Write ".frm1.txtGmDt.text           = """ & UNIDateClientFormat(EG1_export_group(GroupCount,M744_EG1_E11_mvmt_rcpt_dt)) & """" & vbCr
    Response.Write ".frm1.txtGroupCd.Value       = """ & ConvSPChars(EG1_export_group(GroupCount,M743_EG1_E17_pur_grp)) & """"				& vbCr
    Response.Write ".frm1.txtGroupNm.Value       = """ & ConvSPChars(EG1_export_group(GroupCount,M743_EG1_E17_pur_grp_nm)) & """"			& vbCr
    Response.Write ".frm1.txtSupplierCd.Value    = """ & ConvSPChars(EG1_export_group(GroupCount,M743_EG1_E16_bp_cd)) & """"				& vbCr
    Response.Write ".frm1.txtSupplierNm.Value    = """ & ConvSPChars(EG1_export_group(GroupCount,M743_EG1_E16_bp_nm)) & """"				& vbCr
    Response.Write ".frm1.txtMvmtNo1.Value       = """ & ConvSPChars(I2_m_pur_goods_mvmt_rcpt_no) & """"									& vbCr
    Response.Write ".frm1.txtRemark.Value       = """ & ConvSPChars(EG1_export_group(GroupCount,M743_EG1_E17_remark_hdr)) & """"									& vbCr
	Response.Write "End With"                    & vbCr
	Response.Write "</Script>"	                & vbCr
     	
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
                	
		If ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_inspect_flg)) = "Y" Then
			FlgData	= "1"
		Else
			FlgData	= "0"
		End if
		istrData = istrData & Chr(11) & FlgData   
        istrData = istrData & Chr(11) & ""
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,M744_EG1_E11_mvmt_rcpt_qty),ggQty.DecPoint,0)
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,M744_EG1_E11_mvmt_qty),ggQty.DecPoint,0)
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_mvmt_rcpt_unit))  
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_mvmt_cur))
        
        istrData = istrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(EG1_export_group(iLngRow,M744_EG1_E11_mvmt_prc), 0)        
        istrData = istrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(EG1_export_group(iLngRow,M744_EG1_E11_mvmt_doc_amt), 0)   
        istrData = istrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(EG1_export_group(iLngRow,M744_EG1_E11_mvmt_loc_amt), 0)   
        
        istrMvmtCur = EG1_export_group(iLngRow,M744_EG1_E11_mvmt_cur)
        
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E13_sl_cd))
        istrData = istrData & Chr(11) & "" 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E13_sl_nm))

        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E6_minor_nm))
        
        If ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_inspect_flg)) = "Y" Then
			istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E7_minor_nm))
		Else
			istrData	= istrData & Chr(11) & ""
		End if
		        
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_lot_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_lot_sub_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_maker_lot_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_maker_lot_sub_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_remark_dtl))			'비고란 추가 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_gm_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_gm_seq_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_inspect_req_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_inspect_result_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E10_po_no)) 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E9_po_seq_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E4_cc_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E5_cc_seq))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E2_lc_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E3_lc_seq))
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_mvmt_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_iv_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_iv_seq_no))
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,M743_EG1_E17_ret_ord_qty),ggQty.DecPoint,0)
        '2008-03-29 7:38오후 :: hanc
        '아래 4건은 if때문에 만든것이다. 저장됐을때 apply_flag를 update 하기 위해서... hidden으로 4개 pk field를 만들어 놓았다.
        'mes참조popup에서 받아올때 아래 4개 pk field에 넣었다가... 저장시.. 활용하기 위함.
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_iv_seq_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_iv_seq_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_iv_seq_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_iv_seq_no))

        '2008-04-28 8:43오후 :: hanc
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M743_EG1_E11_ext2_cd))
        istrData = istrData & Chr(11) & "" 
        
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
        istrData = istrData & Chr(11) & Chr(12)
        
        TmpBuffer(iIntLoopCount) = istrData
        iIntLoopCount = iIntLoopCount + 1
        
        If strGmNo = "" Then
            strGmNo = Trim(EG1_export_group(iLngRow,M744_EG1_E11_gm_no))
        End if
    Next
    
    iTotalStr = Join(TmpBuffer, "")
    
    Response.Write "<Script Language=VBScript>"      & vbCr
	Response.Write "With parent"                     & vbCr
	Response.Write "	.frm1.hdnMvmtCur.value = """ & ConvSPChars(istrMvmtCur) & """" & vbCr
	Response.Write "	.ggoSpread.Source = .frm1.vspdData"       & vbCr
	Response.Write "    .frm1.vspdData.Redraw = False   "                      & vbCr   
    Response.Write "	.ggoSpread.SSShowData        """ & iTotalStr	    & """ ,""F""" & vbCr
    
    Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData," & lgLngMaxRow + 1 & "," & lgLngMaxRow + iLngRow & "	,.C_Cur		,.C_MvmtPrc		,""C"" ,""I"",""X"",""X"")" & vbCr
    Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData," & lgLngMaxRow + 1 & "," & lgLngMaxRow + iLngRow & "	,.C_Cur		,.C_DocAmt		,""A"" ,""I"",""X"",""X"")" & vbCr
    Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData," & lgLngMaxRow + 1 & "," & lgLngMaxRow + iLngRow & "	,.parent.gCurrency	,.C_LocAmt		,""A"" ,""I"",""X"",""X"")" & vbCr
       
    Response.Write "	.lgStrPrevKey     = """ & StrNextKey   & """" & vbCr 
	Response.Write "	.frm1.hdnMvmtNo.value = """ & ConvSPChars(Request("txtMvmtNo")) & """" & vbCr
    Response.Write "	.frm1.hdnRcptNo.value = """ & ConvSPChars(StrNextKey) & """" & vbCr
    Response.write "	.DbQueryOk "                    & vbCr
    Response.Write "    .frm1.vspdData.Redraw = True   "                      & vbCr   	
    Response.Write "    if """ & strGmNo & """ = """" then " & vbCr
    Response.Write "    .frm1.btnGlSel.disabled = true" & vbCr
    Response.Write "    End If " & vbCr
	Response.Write "End With"                        & vbCr
	Response.Write "</Script>"	                    & vbCr
	
	Set iPM7G428 = Nothing 
	
	
	'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	'******* 전표No. 만들기. *********
	'*************************************************************************************************************
	If  strGmNo <> "" Or strGmNo <> Null then  

		Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
				
		lgStrSQL = "SELECT document_year FROM i_goods_movement_header " 
		lgStrSQL = lgStrSQL & " WHERE item_document_no =  " & FilterVar(strGmNo , "''", "S") & ""		
		
		IF FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False then
			Response.Write "<Script Language=VBScript>" & vbCr
			Response.Write "parent.frm1.hdnGlType.value	=	""B"" " & vbCr
			Response.Write "</Script>"					& vbCr
		
			Call SubCloseRs(lgObjRs)  
			Call SubCloseDB(lgObjConn)
			Exit Sub
		End if
		'A_GL.Ref_no
		strGmNo	=	strGmNo & "-" & lgObjRs("document_year")

		'수정(전표조회 추가)
		Response.Write "<Script Language=VBScript>" & vbCr
		lgStrSQL = "select temp_gl_no,gl_no from ufn_a_GetGlNo(  " & FilterVar(strGmNo, "''", "S") & ") " 		
				
		IF FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False then
			Response.Write "parent.frm1.hdnGlType.value	=	""B""	  " & vbCr
			Response.Write "parent.frm1.hdnGlNo.value	=	""""      " & vbCr  
		Else
			strTempGlNo = lgObjRs("temp_gl_no")
			strGlNo = lgObjRs("gl_no")
			If ConvSPChars(Trim(strGlNo)) = "" And ConvSPChars(Trim(strTempGlNo)) <> "" Then
				Response.Write "parent.frm1.hdnGlType.value	=	""T""	  " & vbCr
				Response.Write "parent.frm1.hdnGlNo.value	=	""" & lgObjRs("temp_gl_no") & """" & vbCr  
			Else
				Response.Write "parent.frm1.hdnGlType.value	=	""A""	  " & vbCr
				Response.Write "parent.frm1.hdnGlNo.value	=	""" & lgObjRs("gl_no") & """" & vbCr  
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
 Sub subBizSaveMulti()															'☜: 저장 요청을 받음 
 
    Dim iPM7G421
    Dim iPM7G452
    Dim iCommandSent
    Dim iErrorPosition
 
    Dim itxtSpread
	Dim itxtSpreadArr
    Dim itxtSpreadArrCount
    Dim iCUCount
    Dim iDCount
    Dim ii
    
    Dim I3_b_biz_partner_bp_cd
    DIm I4_b_pur_grp
    DIm I5_m_mvmt_type_io_type_cd
    Dim I6_m_pur_goods_mvmt
		Const M745_I6_mvmt_rcpt_no = 0  
		Const M745_I6_mvmt_rcpt_dt = 1
		Const M745_I6_gm_no = 2
		Const M745_I6_gm_seq_no = 3
    Redim I6_m_pur_goods_mvmt(M745_I6_gm_seq_no)
    
    Dim I7_remark_hdr
    
    '2006-03-13 구매속도 개선용 변수 선언 
	Dim I8_pur_rcpt_hdr
		Const M7B1_I1_io_type_cd = 0
        Const M7B1_I1_pur_org = 1
        Const M7B1_I1_pur_grp = 2
        Const M7B1_I1_bp_cd = 3
        Const M7B1_I1_mvmt_rcpt_dt = 4
        Const M7B1_I1_mvmt_rcpt_no = 5
        Const M7B1_I1_subcon_flg = 6
        Const M7B1_I1_remark_hdr = 7
        Const M7B1_I1_gm_no = 8
        Const M7B1_I1_gm_year = 9
        Const M7B1_I1_child_settle_flg = 10
        Const M7B1_I1_ret_flg = 11
        Const M7B1_I1_cost_cd = 12
        Const M7B1_I1_biz_area_cd = 13
        Const M7B1_I1_mvmt_cd = 14
        Const M7B1_I1_rcpt_flg = 15
    Redim I8_pur_rcpt_hdr(M7B1_I1_rcpt_flg)
    '2006-03-13 구매속도 개선용 변수 선언 끝 

    Dim E1_m_pur_goods_mvmt
		Const M745_E1_mvmt_rcpt_no = 0
		Const M745_E1_gm_no = 1
    
    '2006-03-13 구매속도 개선용 변수 선언 
    Dim E2_err_info
		Const M7B1_E2_pur_ord_no = 0
        Const M7B1_E2_pur_ord_seq = 1
        Const M7B1_E2_qty = 2
        Const M7B1_E2_trns_item_cd = 3
        Const M7B1_E2_base_unit = 4
        
        
    Dim msgStr1, msgStr2
	'2006-03-13 구매속도 개선용 변수 선언 끝 

	
	On Error Resume Next 		
    Err.Clear																		'☜: Protect system from crashing
	
	itxtSpread = ""
             
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
    
    Call RemovedivTextArea()
    
	If Len(Request("txtGmDt")) Then
		If UNIConvDate(Request("txtGmDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Response.End	
		End If
	End If
	
	lgIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 
	
	If lgIntFlgMode = OPMD_CMODE Then
	   iCommandSent								= "CREATE"
	Else 
	   iCommandSent								= "DELETE"		
	End if
	
	If iCommandSent = "DELETE" Then
		I6_m_pur_goods_mvmt(M745_I6_mvmt_rcpt_no)	= Request("txtMvmtNo1")
		I6_m_pur_goods_mvmt(M745_I6_mvmt_rcpt_dt)	= UNIConvDate(Request("txtGmDt"))
		I6_m_pur_goods_mvmt(M745_I6_gm_no)		    = ""							
		I3_b_biz_partner_bp_cd                  	= Trim(Request("txtSupplierCd"))
		I5_m_mvmt_type_io_type_cd                	= Trim(Request("txtMvmtType"))
		I4_b_pur_grp                            	= Trim(Request("txtGroupCd"))
		I7_remark_hdr								= Trim(Request("txtRemark"))
    
		Set iPM7G421 = Server.CreateObject("PM7G421.cMMaintPurRcptS")    
    
		If CheckSYSTEMError(Err,True) = True Then
			Exit Sub
		End If
	
		Call iPM7G421.M_MAINT_PUR_RCPT_SVR(gStrGlobalCollection, iCommandSent, UCase(I3_b_biz_partner_bp_cd), _
		            UCase(I4_b_pur_grp), UCase(I5_m_mvmt_type_io_type_cd), I6_m_pur_goods_mvmt, itxtSpread, _
		            E1_m_pur_goods_mvmt, iErrorPosition, I7_remark_hdr)

		If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
		   Set iPM7G421 = Nothing
		   Exit Sub
		End If		
    
		Set iPM7G421 = Nothing                                                   '☜: Unload Comproxy  
		
		
	'2006-03-13 구매속도 개선용 시작	
	ElseIf iCommandSent = "CREATE" Then


		Set iPM7G452 = Server.CreateObject("PM7G452.cMMaintPurRcptS")
		
		If CheckSYSTEMError(Err,True) = True Then
			Exit Sub
		End If
		
		I8_pur_rcpt_hdr(M7B1_I1_io_type_cd) = UCase(Trim(Request("txtMvmtType")))
		I8_pur_rcpt_hdr(M7B1_I1_pur_grp) = UCase(Trim(Request("txtGroupCd")))
		I8_pur_rcpt_hdr(M7B1_I1_bp_cd) = UCase(Trim(Request("txtSupplierCd")))
		I8_pur_rcpt_hdr(M7B1_I1_mvmt_rcpt_dt) = UNIConvDate(Request("txtGmDt"))
		I8_pur_rcpt_hdr(M7B1_I1_mvmt_rcpt_no) = UCase(Trim(Request("txtMvmtNo1")))
		I8_pur_rcpt_hdr(M7B1_I1_remark_hdr) = Trim(Request("txtRemark"))
		
		Call iPM7G452.M_REGISTER_PUR_RCPT_SVR(gStrGlobalCollection, _
									I8_pur_rcpt_hdr, _
									itxtSpread, _
									E1_m_pur_goods_mvmt, _
									E2_err_info, _
									iErrorPosition)
		
		Select Case Trim(Cstr(Err.Description))
			
			Case "B_MESSAGE" & Chr(11) & "174136"
				msgStr1 = "발주번호 : " & E2_err_info(M7B1_E2_pur_ord_no) & "   " & _
						  "발주순번 : " & E2_err_info(M7B1_E2_pur_ord_seq) & VbCrLf
				msgStr2 = "자품목 : " & E2_err_info(M7B1_E2_trns_item_cd) & "  " & _
						   UniNumClientFormat(E2_err_info(M7B1_E2_qty),ggQty.DecPoint,0) & " " & E2_err_info(M7B1_E2_base_unit)		   
						
				If CheckSYSTEMError2(Err,True,msgStr1,msgStr2,"","","") = True  Then
					Set iPM7G452 = Nothing
					If iErrorPosition <> 0 Then
						Response.Write "<Script Language=VBScript>" & vbCrLF
						Response.Write "Call parent.SheetFocus(" & iErrorPosition & ", 1)" & vbCrLF
						Response.Write "</Script>" & vbCrLF
					End If
					Response.End
				End If
				
			Case Else
				If CheckSYSTEMError(Err,True) = True Then
					Set iPM7G452 = Nothing
					If iErrorPosition <> 0 Then
						Response.Write "<Script Language=VBScript>" & vbCrLF
						Response.Write "Call parent.SheetFocus(" & iErrorPosition & ", 1)" & vbCrLF
						Response.Write "</Script>" & vbCrLF
					End If
					Response.End
				End If
		End Select
		

        '2008-03-29 5:55오후 :: hanc   begin===============================================================
        '이부분 추가로 코딩한 이유 : 
        '    --> 입고등록시 if 테이블 T_IF_RCV_PART_IN_KO441의 MVMT_RCPT_NO, ERP_APPLY_FLAG에 입고번호와 적용FLAG을 
        '        UPDATE하기 위함.
        Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
        Dim arrRowVal
        Dim arrColVal
        Dim iDx
        Dim lgStrSQL
        Dim s_mvmt_rcpt_no
        Dim s_gm_no
        
            s_mvmt_rcpt_no  =    UCase(ConvSPChars(E1_m_pur_goods_mvmt(M745_E1_mvmt_rcpt_no)))
            s_gm_no         =    UCase(ConvSPChars(E1_m_pur_goods_mvmt(M745_E1_gm_no)))
    
        	arrRowVal = Split(itxtSpread, gRowSep)                                 '☜: Split Row    data
        	
            For iDx = 0 To UBound(arrRowVal,1) - 1
                arrColVal = Split(arrRowVal(iDx), gColSep)                                 '☜: Split Column data
                call SubBizSaveMultiUpdate(arrColVal, s_mvmt_rcpt_no)
    
            Next
    		
        Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection
        '2008-03-29 5:55오후 :: hanc   end  ===============================================================
		
		
		
		Set iPM7G452 = Nothing
		
	End If	
    '2006-03-13 구매속도 개선 끝 
    
    Response.Write "<Script language=vbs> " & vbCr 
	Response.Write "With parent " & vbCr
	Response.Write "If """ & lgIntFlgMode & """ = """ & OPMD_CMODE & """ Then " & vbCr
	Response.Write ".frm1.txtMvmtNo.Value = """ & UCase(ConvSPChars(E1_m_pur_goods_mvmt(M745_E1_mvmt_rcpt_no))) & """ " & vbCr
	Response.Write ".frm1.hPoNo.Value = """ & UCase(ConvSPChars(E1_m_pur_goods_mvmt(M745_E1_mvmt_rcpt_no))) & """ " & vbCr
	Response.Write "End If"				& vbCr	
	Response.Write " .DbSaveOk "      & vbCr						
	Response.Write "End With " & vbCr
	Response.Write "</Script> "    
    
End Sub	

'2008-03-29 8:29오후 :: hanc
Sub SubBizSaveMultiUpdate(arrColVal, s_mvmt_rcpt_no)
    Dim lgStrSQL

    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
            lgStrSQL = "UPDATE T_IF_RCV_PART_IN_KO441 SET "
            lgStrSQL = lgStrSQL & " ERP_APPLY_FLAG   = 'Y',  "
            lgStrSQL = lgStrSQL & " UPDT_USER_ID     = " & FilterVar(gUsrId,"","S")                        & ","  
            lgStrSQL = lgStrSQL & " UPDT_DT          = GetDate()," 
            lgStrSQL = lgStrSQL & " MVMT_RCPT_NO     = " &  FilterVar(Trim(UCase(s_mvmt_rcpt_no)),"''","S")
            lgStrSQL = lgStrSQL & " WHERE TRANS_TIME    = " &  FilterVar(Trim(UCase(arrColVal(56))),"''","S")
            lgStrSQL = lgStrSQL & " AND   MAIN_LOT      = " &  FilterVar(Trim(UCase(arrColVal(57))),"''","S")
            lgStrSQL = lgStrSQL & " AND   IMPORT_TIME   = " &  FilterVar(Trim(UCase(arrColVal(58))),"''","S")
            lgStrSQL = lgStrSQL & " AND   CREATE_TYPE   = " &  FilterVar(Trim(UCase(arrColVal(59))),"''","S")

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords

    If CheckSYSTEMError(Err,True) = True Then
       lgErrorStatus    = "YES"
       ObjectContext.SetAbort
    End If
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
