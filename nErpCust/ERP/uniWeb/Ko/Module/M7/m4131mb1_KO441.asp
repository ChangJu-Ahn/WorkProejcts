<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
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
'*  3. Program ID           : m4131mb1
'*  4. Program Name         : 검사결과 
'*  5. Program Desc         :
'*  6. Comproxy List        : im41311/im41318
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2003/06/04
'*  9. Modifier (First)     : eVerfOrever
'* 10. Modifier (Last)      : Kim Jin Ha
'* 11. Comment              :
'* 12. Common Coding Guide  : 
'* 13. History              :
'**********************************************************************************************
													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
On Error Resume Next

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
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    On Error Resume Next
    Err.Clear                                                               '☜: Protect system from crashing
    
    Dim OBJ_PM7GR38	
    
    Dim E1_m_pur_goods_mvmt
    Const M706_E1_mvmt_no = 0    
    
    Redim E1_m_pur_goods_mvmt(0) 	
        
    Dim EG1_exp_group
    Const M706_EG1_E1_sl_nm = 0			'양품창고 
	Const M706_EG1_E2_sl_nm = 1			'불량품창고 
	Const M706_EG1_E3_minor_nm = 2		'검사상태 
	Const M706_EG1_E4_minor_nm = 3		'검사방법 
	Const M706_EG1_E5_plant_cd = 4    
	Const M706_EG1_E5_plant_nm = 5
	Const M706_EG1_E6_item_cd = 6    
	Const M706_EG1_E6_item_nm = 7
	Const M706_EG1_E6_spec = 8
	Const M706_EG1_E7_mvmt_no = 9  
	Const M706_EG1_E7_mvmt_dt = 10
	Const M706_EG1_E7_gm_no = 11
	Const M706_EG1_E7_gm_year = 12
	Const M706_EG1_E7_gm_seq_no = 13
	Const M706_EG1_E7_gm_sub_seq_no = 14
	Const M706_EG1_E7_mvmt_qty = 15
	Const M706_EG1_E7_mvmt_unit = 16
	Const M706_EG1_E7_mvmt_base_qty = 17
	Const M706_EG1_E7_mvmt_base_unit = 18
	Const M706_EG1_E7_mvmt_prc = 19
	Const M706_EG1_E7_mvmt_cur = 20
	Const M706_EG1_E7_mvmt_doc_amt = 21
	Const M706_EG1_E7_mvmt_loc_amt = 22
	Const M706_EG1_E7_xch_rt = 23
	Const M706_EG1_E7_applied_xch_type = 24
	Const M706_EG1_E7_fr_trans_coef = 25
	Const M706_EG1_E7_to_trans_coef = 26
	Const M706_EG1_E7_cc_no_for_sales = 27
	Const M706_EG1_E7_cc_seq_for_sales = 28
	Const M706_EG1_E7_cc_qty_for_sales = 29
	Const M706_EG1_E7_iv_qty = 30
	Const M706_EG1_E7_pur_org = 31
	Const M706_EG1_E7_mvmt_sl_cd = 32
	Const M706_EG1_E7_mvmt_biz_area = 33
	Const M706_EG1_E7_lot_no = 34
	Const M706_EG1_E7_lot_sub_no = 35
	Const M706_EG1_E7_maker_lot_no = 36
	Const M706_EG1_E7_maker_lot_sub_no = 37
	Const M706_EG1_E7_distribt_amt = 38
	Const M706_EG1_E7_distribt_flg = 39
	Const M706_EG1_E7_mvmt_rcpt_no = 40
	Const M706_EG1_E7_mvmt_rcpt_dt = 41
	Const M706_EG1_E7_mvmt_rcpt_qty = 42
	Const M706_EG1_E7_mvmt_rcpt_unit = 43
	Const M706_EG1_E7_mvmt_rcpt_sl_cd = 44
	Const M706_EG1_E7_mvmt_method = 45
	Const M706_EG1_E7_inspect_req_no = 46
	Const M706_EG1_E7_inspect_sts = 47
	Const M706_EG1_E7_inspect_flg = 48
	Const M706_EG1_E7_inspect_good_qty = 49
	Const M706_EG1_E7_inspect_bad_qty = 50
	Const M706_EG1_E7_inspect_good_base_qty = 51
	Const M706_EG1_E7_inspect_bad_base_qty = 52
	Const M706_EG1_E7_inspect_result_no = 53
	Const M706_EG1_E7_inspect_result_dt = 54
	Const M706_EG1_E7_to_good_sl_cd = 55
	Const M706_EG1_E7_to_bad_sl_cd = 56
	Const M706_EG1_E7_inspect_result_gqty = 57
	Const M706_EG1_E7_inspect_result_bqty = 58
	Const M706_EG1_E7_ext1_cd = 59
	Const M706_EG1_E7_ext1_qty = 60
	Const M706_EG1_E7_ext1_amt = 61
	Const M706_EG1_E7_ext1_rt = 62
	Const M706_EG1_E7_ext2_cd = 63
	Const M706_EG1_E7_ext2_qty = 64
	Const M706_EG1_E7_ext2_amt = 65
	Const M706_EG1_E7_ext2_rt = 66
	Const M706_EG1_E7_ext3_cd = 67
	Const M706_EG1_E7_ext3_qty = 68
	Const M706_EG1_E7_ext3_amt = 69
	Const M706_EG1_E7_ext3_rt = 70
	Const M706_EG1_E7_ap_flg = 71
	Const M706_EG1_E7_iv_no = 72
	Const M706_EG1_E7_iv_seq_no = 73

	Const M706_EG1_E7_ret_ord_qty = 74

	Dim StrNextKey		' 다음 값 
	Dim lgStrPrevKey	' 이전 값 
	Dim istrData
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
	Dim GroupCount          
	Dim index,Count     ' 저장 후 Return 해줄 값을 넣을때 쓴는 변수 
	
	Dim strGmNo, strTempGlNo, strGlNo 
	Dim iStrRsRegNo
	
	Dim TmpBuffer
    Dim iMax
    Dim iIntLoopCount
    Dim iTotalStr
    
	Const C_SHEETMAXROWS_D  = 100
	
	lgStrPrevKey = Request("lgStrPrevKey")
     
    Set OBJ_PM7GR38 = Server.CreateObject("PM7GR38.cMListInspResultS")    

    If CheckSYSTEMError(Err,True) = True Then
			Exit Sub
	End If
	
	iStrRsRegNo = Request("txtRsRegNo")
	
    Call OBJ_PM7GR38.M_LIST_INSP_RESULT_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, lgStrPrevKey, _
												iStrRsRegNo, E1_m_pur_goods_mvmt, EG1_exp_group)					

	
	If CheckSYSTEMError2(Err,True,"","","","","") = True Then
			Set OBJ_PM7GR38 = Nothing
			Response.Write "<Script Language=VBScript>" & vbCr
			Response.Write "	Parent.FncNew() " & vbCr
			Response.Write "	parent.frm1.txtRsRegNo.Value 	= """ & ConvSPChars(Request("txtRsRegNo")) & """" & vbCr
			Response.Write "</Script>" & vbCr
			Exit Sub
	End If
	
	Set OBJ_PM7GR38 = Nothing

	iLngMaxRow = Request("txtMaxRows")											'Save previous Maxrow                                                
    GroupCount = UBound(EG1_exp_group,1)
 
    IF GroupCount <> 0 then
		If EG1_exp_group(GroupCount, M706_EG1_E7_mvmt_no) = E1_m_pur_goods_mvmt(M706_E1_mvmt_no) Then
			StrNextKey = ""
		Else
			StrNextKey = EG1_exp_group(GroupCount, M706_EG1_E7_mvmt_no)
		End If
	End if
		
	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "With parent "		& vbCr
	Response.Write ".frm1.txtRegNo.Value 		= """ & ConvSPChars(EG1_exp_group(0 , M706_EG1_E7_mvmt_rcpt_no)) & """" & vbCr
	Response.Write ".frm1.txtReDt.text    		= """ & UNIDateClientFormat(EG1_exp_group(0, M706_EG1_E7_inspect_result_dt)) & """" & vbCr
	Response.Write ".frm1.txtRsRegNo1.Value 	= """ & ConvSPChars(Request("txtRsRegNo")) & """" & vbCr
	Response.Write "End With" & vbCr
	Response.Write "</Script>" & vbCr	
	
	iIntLoopCount = 0
	iMax = UBound(EG1_exp_group,1)
	ReDim TmpBuffer(iMax)
	
	For iLngRow = 0 To iMax
		If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
		   StrNextKey = ConvSPChars(EG1_exp_group(iLngRow, M706_EG1_E7_mvmt_no)) 
           Exit For
        End If  
		
		istrData = ""
		istrData = istrData & Chr(11) & ConvSPChars( EG1_exp_group(iLngRow, M706_EG1_E5_plant_cd ))
        istrData = istrData & Chr(11) & ConvSPChars( EG1_exp_group(iLngRow, M706_EG1_E5_plant_nm ))
		istrData = istrData & Chr(11) & ConvSPChars( EG1_exp_group(iLngRow, M706_EG1_E6_item_cd ))
        istrData = istrData & Chr(11) & ConvSPChars( EG1_exp_group(iLngRow, M706_EG1_E6_item_nm ))
        istrData = istrData & Chr(11) & ConvSPChars( EG1_exp_group(iLngRow, M706_EG1_E6_spec))
        
        istrData = istrData & Chr(11) & UNINumClientFormat( EG1_exp_group(iLngRow, M706_EG1_E7_mvmt_rcpt_qty),ggQty.DecPoint,0)
        istrData = istrData & Chr(11) & UNINumClientFormat( EG1_exp_group(iLngRow,M706_EG1_E7_inspect_result_gqty),ggQty.DecPoint,0)
        istrData = istrData & Chr(11) & UNINumClientFormat( EG1_exp_group(iLngRow,M706_EG1_E7_inspect_result_bqty),ggQty.DecPoint,0)
        istrData = istrData & Chr(11) & ConvSPChars( EG1_exp_group(iLngRow, M706_EG1_E7_mvmt_rcpt_unit))
        istrData = istrData & Chr(11) & ConvSPChars( EG1_exp_group(iLngRow, M706_EG1_E7_inspect_sts))
        istrData = istrData & Chr(11) & ConvSPChars( EG1_exp_group(iLngRow,M706_EG1_E3_minor_nm ))
        istrData = istrData & Chr(11) & ConvSPChars( EG1_exp_group(iLngRow, M706_EG1_E7_mvmt_method ))  
        istrData = istrData & Chr(11) & ConvSPChars( EG1_exp_group(iLngRow, M706_EG1_E4_minor_nm ))		
		istrData = istrData & Chr(11) & ConvSPChars( EG1_exp_group(iLngRow, M706_EG1_E7_to_good_sl_cd))
		istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars( EG1_exp_group(iLngRow,M706_EG1_E1_sl_nm))
		istrData = istrData & Chr(11) & ConvSPChars( EG1_exp_group(iLngRow,M706_EG1_E7_to_bad_sl_cd))    
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars( EG1_exp_group(iLngRow, M706_EG1_E2_sl_nm ))		
        istrData = istrData & Chr(11) & ConvSPChars( EG1_exp_group(iLngRow, M706_EG1_E7_gm_no))
        istrData = istrData & Chr(11) & ConvSPChars( EG1_exp_group(iLngRow, M706_EG1_E7_gm_seq_no))
        istrData = istrData & Chr(11) & ConvSPChars( EG1_exp_group(iLngRow, M706_EG1_E7_inspect_req_no))
        istrData = istrData & Chr(11) & ConvSPChars( EG1_exp_group(iLngRow, M706_EG1_E7_mvmt_no))
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow,M706_EG1_E7_ret_ord_qty),ggQty.DecPoint,0)
        
        istrData = istrData & Chr(11) & ConvSPChars( EG1_exp_group(iLngRow,M706_EG1_E7_lot_no))
        istrData = istrData & Chr(11) & ConvSPChars( EG1_exp_group(iLngRow,M706_EG1_E7_lot_sub_no))

        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
        istrData = istrData & Chr(11) & Chr(12)
        
        If strGmNo = "" Then
            strGmNo = Trim(EG1_exp_group(iLngRow, M706_EG1_E7_gm_no))
        End if
        
        TmpBuffer(iIntLoopCount) = istrData
        iIntLoopCount = iIntLoopCount + 1
    Next
	
	iTotalStr = Join(TmpBuffer, "")
		
	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "With parent"		& vbCr
	Response.Write "	.ggoSpread.Source 		= .frm1.vspdData" & vbCr
	Response.Write "	.ggoSpread.SSShowData """ & iTotalStr & """" & vbCr
	Response.Write "	.lgStrPrevKey 			= """ & StrNextKey & """" & vbCr
	Response.Write "	.frm1.hdnRsRegNo.value 	= """ & ConvSPChars(Request("txtRsRegNo")) & """ " & vbCr
	Response.Write "	.DBQueryOK									" & vbCr
	Response.Write "End With" & vbCr
	Response.Write "</Script>" & vbCr

    
	
	'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	'******* 전표No. 만들기. *********
	'*************************************************************************************************************	
	If  EG1_exp_group(0, M706_EG1_E7_gm_no) <> "" Or EG1_exp_group(0, M706_EG1_E7_gm_no)  <> Null then 
		
		Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
		
		lgStrSQL = "SELECT document_year FROM i_goods_movement_header " 
		lgStrSQL = lgStrSQL & " WHERE item_document_no =  " & FilterVar(EG1_exp_group(0, M706_EG1_E7_gm_no), "''", "S") & ""		
		
		If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False then
			Response.Write "<Script Language=VBScript>" & vbCr
			Response.Write "With parent.frm1" & vbCr
			Response.Write "parent.frm1.hdnGlType.value	=	""B"" " & vbCr
			Response.Write "End With"
			Response.Write "</Script>" & vbCr
			
			Call SubCloseRs(lgObjRs)
			Call SubCloseDB(lgObjConn)  
			Exit Sub
		End if
		
		strGmNo	= strGmNo & "-" & lgObjRs("document_year")
							
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
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
    Dim iObjPM7G311
    Dim iObjPM7G651
	Dim iErrorPosition 												
	Dim LngMaxRow																			' 현재 그리드의 최대Row
	Dim lgIntFlgMode
    Const iFlgMode = 1000
    
    Dim itxtSpread
	Dim itxtSpreadArr
    Dim itxtSpreadArrCount
    Dim iCUCount
    Dim iDCount
    Dim ii
    
    Dim m_pur_goods_mvmt_hdr
    Const pgm_hdr_Result_No = 0
    Const pgm_hdr_Reg_Dt = 1
    
    Dim iStrExpMPurGoodMvmtGmNo
    Dim iStrBAutoNumberingAutoNo
    
    Dim E3_err_info
	
	ReDim m_pur_goods_mvmt_hdr(1)
	
    On Error Resume Next																	'☜: Protect system from crashing
	Err.Clear 																				'☜: Clear Error status																		'☜: Protect system from crashing
	
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
    
	lgIntFlgMode = Trim(Request("txtFlgMode"))	
	iStrRegDt	 = Trim(Request("txtReDt"))
	
	m_pur_goods_mvmt_hdr(pgm_hdr_Result_No) = Trim(Request("txtRsRegNo1"))
	m_pur_goods_mvmt_hdr(pgm_hdr_Reg_Dt) = UNIConvDate(Trim(Request("txtReDt")))
	
	LngMaxRow = CInt(Request("txtMaxRows"))													'☜: 최대 업데이트된 갯수 
	
	Call RemovedivTextArea()
	
	Set iObjPM7G311 = Server.CreateObject("PM7G311.cMntPurInspResultS")
	
	If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
	End If
	
	Set iObjPM7G651 = Server.CreateObject("PM7G651.cMMngInspResultS")

    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
	End If
    
    If lgIntFlgMode = CStr(OPMD_UMODE) Then
		Call iObjPM7G311.M_MAINT_PUR_INSP_RESULT_SVR(gStrglobalcollection, itxtSpread, m_pur_goods_mvmt_hdr, _
												CLng(iErrorPosition), iStrExpMPurGoodMvmtGmNo, iStrBAutoNumberingAutoNo)
    Else
		Call iObjPM7G651.M_MNG_PUR_INSP_RSLT_SVR(gStrglobalcollection, itxtSpread, m_pur_goods_mvmt_hdr, _
											CLng(iErrorPosition), iStrExpMPurGoodMvmtGmNo, iStrBAutoNumberingAutoNo, E3_err_info)
	End If
	
	If CheckSYSTEMError2(Err,True, "","","","","") = True Then
	'If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
		Set iObjPM7G311 = Nothing															'☜: Unload Comproxy
		Set iObjPM7G651 = Nothing
		Exit Sub
	End If
	
	Set iObjPM7G311 = Nothing																'☜: Unload Comproxy
	Set iObjPM7G651 = Nothing
    Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "With parent " & vbCr																	
	Response.Write "	If """ & lgIntFlgMode & """ = """ & Cstr(OPMD_CMODE) & """ Then " & vbCr
	Response.Write "		.frm1.txtRsRegNo.Value = """ & iStrBAutoNumberingAutoNo & """ " & vbCr
	Response.Write "	End If " & vbCr
	Response.Write "	.DbSaveOk	" & vbCr
	Response.Write "End With" & vbCr
	Response.Write "</Script>" & vbCr			
	'Response.End																				'☜: Process End
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


