<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : mc600mb1
'*  4. Program Name         : 납입지시입고등록 
'*  5. Program Desc         : 납입지시입고등록 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003-02-25
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Ryu Sung Won
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%
	On Error Resume Next
	
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("*", "M","NOCOOKIE", "MB")
	Call LoadBNumericFormatB("*", "M","NOCOOKIE", "MB")
	Call HideStatusWnd

	Dim istrData
	Dim PvArr
	Dim StrNextKey		' 다음 값 
	Dim lgStrPrevKey	' 이전 값 
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
	Dim GroupCount          
	Dim index,Count     ' 저장 후 Return 해줄 값을 넣을때 쓴는 변수 
	
	Const C_SHEETMAXROWS_D  = 100

	Call SubBizQueryMulti()

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
 
 Dim iPM7G428
 
 Dim FlgData
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
 
 Dim strGmNo   
 Dim istrMvmtCur	
 
    On Error Resume Next 
    Err.Clear                                                               '☜: Protect system from crashing
    
	lgStrPrevKey = Request("lgStrPrevKey")
  
    Set iPM7G428 = Server.CreateObject("PM7G428.cMListPurRcptS")    

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = True Then
       Set iPM7G428 = Nothing
	   Exit Sub
	End If
	
    '-----------------------
    'Data manipulate  area(import view match)    
    '-----------------------
    I2_m_pur_goods_mvmt_rcpt_no  = FilterVar(UCase(Trim(Request("txtMvmtNo"))),"","SNM")
    
    if Trim(lgStrPrevKey) <> "" then
	   I1_m_pur_goods_mvmt_no  	= lgStrPrevKey
	End if
    
    '-----------------------
    'Com action area
    '-----------------------
	Call iPM7G428.M_LIST_PUR_RCPT_SVR (gStrGlobalCollection, _
										C_SHEETMAXROWS_D, _ 
										I1_m_pur_goods_mvmt_no, _
										I2_m_pur_goods_mvmt_rcpt_no, _
										EG1_export_group)
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError2(Err,True,"","","","","") = true then 
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "	Parent.SetDefaultVal " & vbCr
		Response.Write "</Script>" & vbCr
		Set iM41218 = Nothing
		Exit Sub
	End If
		
    iLngMaxRow = Request("txtMaxRows")											'Save previous Maxrow                                                
    GroupCount = UBound(EG1_export_group,1)
    
 	
	   Response.Write "<Script Language=VBScript>" & vbCr
	   Response.Write "With parent"                 & vbCr
       Response.Write ".frm1.cboMvmtType.Value      = """ & ConvSPChars(EG1_export_group(GroupCount,M743_EG1_E14_io_type_cd)) & """"                              & vbCr
       Response.Write ".frm1.txtGmDt.text           = """ & UNIDateClientFormat(EG1_export_group(GroupCount,M744_EG1_E11_mvmt_rcpt_dt)) & """" & vbCr
       Response.Write ".frm1.txtGroupCd.Value       = """ & ConvSPChars(EG1_export_group(GroupCount,M743_EG1_E17_pur_grp)) & """"                                   & vbCr
       Response.Write ".frm1.txtGroupNm.Value       = """ & ConvSPChars(EG1_export_group(GroupCount,M743_EG1_E17_pur_grp_nm)) & """"                                & vbCr
       Response.Write ".frm1.txtSupplierCd.Value    = """ & ConvSPChars(EG1_export_group(GroupCount,M743_EG1_E16_bp_cd)) & """"                                 & vbCr
       Response.Write ".frm1.txtSupplierNm.Value    = """ & ConvSPChars(EG1_export_group(GroupCount,M743_EG1_E16_bp_nm)) & """"                                 & vbCr
       Response.Write ".frm1.txtMvmtNo1.Value       = """ & ConvSPChars(I2_m_pur_goods_mvmt_rcpt_no) & """"                                     & vbCr
	   Response.Write "End With"                    & vbCr
	   Response.Write "</Script>"	                & vbCr
     	
	'-----------------------
	'Result data display area
	'----------------------- 
	ReDim PvArr(UBound(EG1_export_group,1))
	
	For iLngRow = 0 To UBound(EG1_export_group,1)
		If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
		   StrNextKey = ConvSPChars(EG1_export_group(iLngRow, M744_EG1_E11_mvmt_no)) 
           Exit For
        End If  
		
		istrData = ""
		
        istrData = Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E8_plant_cd))				'입고Seq
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
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,M744_EG1_E11_mvmt_rcpt_qty),ggQty.DecPoint,0)
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_mvmt_rcpt_unit))  
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E13_sl_cd))
        istrData = istrData & Chr(11) & "" 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E13_sl_nm))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_lot_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_lot_sub_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_maker_lot_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_maker_lot_sub_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E10_po_no)) 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E9_po_seq_no))
        istrData = istrData & Chr(11) & Chr(11) & Chr(11) & Chr(11)
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_mvmt_no))
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow & Chr(11)
        
        If strGmNo = "" Then
            strGmNo = Trim(EG1_export_group(iLngRow,M744_EG1_E11_gm_no))
        End if
        
        PvArr(iLngRow) = istrData
        
    Next
    
    istrData = Join(PvArr, Chr(12)) & Chr(12)
    
    Response.Write "<Script Language=VBScript>"      & vbCr
	Response.Write "With parent"                     & vbCr
	Response.Write "	.frm1.hdnMvmtCur.value = """ & ConvSPChars(istrMvmtCur) & """" & vbCr
	Response.Write "	.frm1.vspdData.Redraw = False " & vbCr

	Response.Write "	.ggoSpread.Source = .frm1.vspdData"       & vbCr
	Response.Write "	.ggoSpread.SpreadLock -1,	-1    "       & vbCr

    Response.Write "	.ggoSpread.SSShowData        """ & istrData	    & """" & vbCr	   
    Response.Write "	.lgStrPrevKey     = """ & StrNextKey   & """" & vbCr 
	
    Response.Write "	.frm1.hdnMvmtNo.value = """ & ConvSPChars(Request("txtMvmtNo")) & """" & vbCr
    Response.Write "	.frm1.hdnRcptNo.value = """ & ConvSPChars(StrNextKey) & """" & vbCr
    Response.write "	.DbQueryOk "                    & vbCr
	Response.Write "	.frm1.vspdData.Redraw = True " & vbCr
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
		
		lgStrSQL = "SELECT gl_no FROM a_gl " 
		lgStrSQL = lgStrSQL & " WHERE ref_no =  " & FilterVar(strGmNo , "''", "S") & ""		
				
		IF FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False then
			
			lgStrSQL = "SELECT temp_gl_no FROM a_temp_gl " 
			lgStrSQL = lgStrSQL & " WHERE ref_no =  " & FilterVar(strGmNo , "''", "S") & ""		
						
			IF FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False then
				Response.Write "<Script Language=VBScript>" & vbCr
				Response.Write "parent.frm1.hdnGlType.value	=	""B""	  " & vbCr
				Response.Write "parent.frm1.hdnGlNo.value	=	""""      " & vbCr  
				Response.Write "</Script>"					& vbCr
		        Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSet		
	    	Else
				Response.Write "<Script Language=VBScript>" & vbCr
				Response.Write "parent.frm1.hdnGlType.value	=	""T""	  " & vbCr
				Response.Write "parent.frm1.hdnGlNo.value	=	""" & lgObjRs("temp_gl_no") & """" & vbCr  
				Response.Write "</Script>"					& vbCr
			End if
		    Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSet		
	    Else
			Response.Write "<Script Language=VBScript>" & vbCr
			Response.Write "parent.frm1.hdnGlType.value	=	""A""	  " & vbCr
			Response.Write "parent.frm1.hdnGlNo.value	=	""" & lgObjRs("gl_no") & """" & vbCr  
			Response.Write "</Script>"					& vbCr
		End if	
	    
	    Call SubCloseDB(lgObjConn)	                                                '☜ : Release DB Connection		
	Else
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "parent.frm1.hdnGlType.value	=	""B""	  " & vbCr
		Response.Write "</Script>"					& vbCr
	End if
	
 End Sub
        
%>
