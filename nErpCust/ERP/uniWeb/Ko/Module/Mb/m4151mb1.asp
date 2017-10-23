<%@ LANGUAGE=VBSCript%>
<%Option Explicit    
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m4151mb1
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : M4111(Maint)
'							  iM41218(List)
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2003-06-02
'*  9. Modifier (First)     : Shin Jin-hyun
'* 10. Modifier (Last)      : Kim Jin Ha
'* 11. Comment              :
'* 12. History              :
'**********************************************************************************************
%>	
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%
	Dim istrData
	Dim StrNextKey																' 다음 값 
	Dim lgStrPrevKey															' 이전 값 
	Dim iLngMaxRow																' 현재 그리드의 최대Row
	Dim iLngRow
	Dim GroupCount          
	Dim index,Count																' 저장 후 Return 해줄 값을 넣을때 쓴는 변수 
	
	On Error Resume Next														'☜: Protect system from crashing
	Err.Clear 																	'☜: Clear Error status
				
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "M","NOCOOKIE", "MB")
	Call LoadBNumericFormatB("I", "M","NOCOOKIE", "MB")
	Call HideStatusWnd

	lgOpModeCRUD = Request("txtMode")											'☜: Read Operation Mode (CRUD)			

	Select Case CSTR(Trim(lgOpModeCRUD))
	        Case CStr(UID_M0001)                                                         '☜: Query
	             Call SubBizQueryMulti()
	        Case CStr(UID_M0002)
	             Call SubBizSaveMulti()
	        Case "changeMvmtType" 
				 Call SubchangeMvmtType()
			Case "changeSpplCd" 
				 Call DisplaySupplierNm(Request("txtSupplierCd"))	 
			Case "changeGroupCd" 
				 Call DisplayGroupNm(Request("txtGroupCd"))	 		
	End Select
	
'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	Dim iM41218																	'☆ : 조회용 ComProxy Dll 사용 변수 
	
	Const C_SHEETMAXROWS_D = 100
	
	Dim TmpBuffer
	Dim iMax
	Dim iIntLoopCount
	Dim iTotalStr
	
	Dim FlgData
	Dim I1_m_pur_goods_mvmt_no
	Dim I2_m_pur_goods_mvmt_rcpt_no
	Dim EG1_export_group

    Const M743_E1_pur_grp = 0	
	Const M743_E1_pur_grp_nm = 1
	 
	Const M743_E2_bp_cd = 0    
	Const M743_E2_bp_nm = 1

	Const M743_E4_io_type_cd = 0   
	Const M743_E4_io_type_nm = 1
    
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
    
    On Error Resume Next
    Err.Clear																		  '☜: Protect system from crashing
    
	lgStrPrevKey = Request("lgStrPrevKey")
												
    Set iM41218 = Server.CreateObject("PM7G428.cMListPurRcptS")    
   
    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
	End If
 	
    I2_m_pur_goods_mvmt_rcpt_no	= Trim(Request("txtGrNo"))
    
    if Trim(lgStrPrevKey) <> "" then
		I1_m_pur_goods_mvmt_no = lgStrPrevKey
	End if
	
	Call iM41218.M_LIST_PUR_RCPT_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_m_pur_goods_mvmt_no, _
                                                           I2_m_pur_goods_mvmt_rcpt_no, EG1_export_group)
    
    If CheckSYSTEMError2(Err,True,"","","","","") = true then 
		Set iM41218 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "	parent.fncNew()				" & vbCr
		Response.Write "	parent.frm1.txtGrNo.Value   = """ & ConvSPChars(Request("txtGrNo")) & """" & vbCr
		Response.Write "	parent.frm1.txtGrNo.focus	" & vbCr
		Response.Write "</Script>" & vbCr	
		Exit Sub
	End If
	
	Set iM41218 = Nothing
	
	iLngMaxRow = Request("txtMaxRows")											'Save previous Maxrow                                                
    GroupCount = UBound(EG1_export_group,1)

	'-----------------------
	'Result data display area
	'----------------------- 
	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write " Dim LngMaxRow " & vbCr
	Response.Write " With parent " & vbCr
	Response.Write "	LngMaxRow = .frm1.vspdData.MaxRows " & vbCr
	Response.Write "	.frm1.txtMvmtType.Value 	= """ & ConvSPChars(EG1_export_group(GroupCount,M743_EG1_E14_io_type_cd)) & """" & vbCr
	Response.Write "	.frm1.txtMvmtTypeNm.Value 	= """ & ConvSPChars(EG1_export_group(GroupCount,M743_EG1_E14_io_type_nm)) & """" & vbCr
	Response.Write "	.frm1.txtGmDt.text    		= """ & UNIDateClientFormat(EG1_export_group(0,M744_EG1_E11_mvmt_dt)) & """" & vbCr
	Response.Write "	.frm1.txtSupplierCd.Value 	= """ & ConvSPChars(EG1_export_group(GroupCount,M743_EG1_E16_bp_cd)) & """" & vbCr
	Response.Write "	.frm1.txtSupplierNm.Value 	= """ & ConvSPChars(EG1_export_group(GroupCount,M743_EG1_E16_bp_nm)) & """" & vbCr
	Response.Write "	.frm1.txtGroupCd.Value	  	= """ & ConvSPChars(EG1_export_group(GroupCount,M743_EG1_E17_pur_grp)) & """" & vbCr
	Response.Write "	.frm1.txtGroupNm.Value    	= """ & ConvSPChars(EG1_export_group(GroupCount,M743_EG1_E17_pur_grp_nm)) & """" & vbCr
	Response.Write "	.frm1.txtGrNo1.Value    	= """ & ConvSPChars(Request("txtGrNo")) & """" & vbCr
	Response.Write " End With " & vbCr
	Response.Write "</Script>" & vbCr
	
	iIntLoopCount = 0
	iMax = UBound(EG1_export_group,1)
	ReDim TmpBuffer(iMax)
	
	
	For iLngRow = 0 To iMax
		
		If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
		   StrNextKey = ConvSPChars(EG1_export_group(iLngRow, M744_EG1_E11_mvmt_no)) 
           Exit For
	    End If  
		istrData=""
		istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E8_plant_cd))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E8_plant_nm))
		istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E12_item_cd))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E12_item_nm))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E12_spec))
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,M744_EG1_E11_mvmt_rcpt_qty),ggQty.DecPoint,0)
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_mvmt_unit))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E13_sl_cd))
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E13_sl_nm))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_lot_no))
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_lot_sub_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E10_po_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E9_po_seq_no))
		istrData = istrData & Chr(11) & ""'pr_no
		istrData = istrData & Chr(11) & ""'resvdSeq
		istrData = istrData & Chr(11) & ""'HstySubSeqNo
		istrData = istrData & Chr(11) & ""'LotFlg
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_mvmt_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M744_EG1_E11_tracking_no))
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
        istrData = istrData & Chr(11) & Chr(12)
		
		TmpBuffer(iIntLoopCount) = istrData
        iIntLoopCount = iIntLoopCount + 1
	Next
	
	iTotalStr = Join(TmpBuffer, "")    
	
	Response.Write "<Script Language=VBScript>"		& vbCr
	Response.Write "With parent"					& vbCr
	Response.Write "	.ggoSpread.Source 	= .frm1.vspdData"				& vbCr
	Response.Write "	.ggoSpread.SSShowData """ & iTotalStr & """ "		& vbCr
	Response.Write "	.lgStrPrevKey 		= """ & StrNextKey & """ "		& vbCr
	Response.Write "		.frm1.hdnMvmtNo.value 	= """ & ConvSPChars(Request("txtGrNo")) & """ " & vbCr
	Response.Write "		.DBQueryOK	"									& vbCr
	Response.Write "End With"						& vbCr
	Response.Write "</Script>"						& vbCr
  
End Sub
'============================================================================================================
Sub SubBizSaveMulti()
	On Error Resume Next							
    Err.Clear		
	
	Dim iM41511																	'☆ : 입력/수정용 ComProxy Dll 사용 변수 
	
	Dim arrVal, arrTemp																'☜: Spread Sheet 의 값을 받을 Array 변수 
	Dim strStatus																	'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
	Dim	lGrpCnt																		'☜: Group Count
	Dim lgIntFlgMode
	Dim iErrorPosition
	
	Dim itxtSpread
	Dim itxtSpreadArr
    Dim itxtSpreadArrCount
    Dim iCUCount
    Dim iDCount
    Dim ii
    
	Const C_PlantCd		= 1
	Const C_PlantNm		= 2
	Const C_ItemCd		= 3
	Const C_ItemNm		= 4
	Const C_Spec		= 5
	Const C_SubQty		= 6
	Const C_Unit		= 7
	Const C_SlCd		= 8
	Const C_SlPop		= 9
	Const C_SlNm		= 10
	Const C_LotNo		= 11
	Const C_LotNoPop	= 12
	Const C_LotNoSeq	= 13
	Const C_PoNo		= 14
	Const C_PoSeq		= 15
	Const C_PrNo		= 16
	Const C_ResvdSeq	= 17
	Const C_HstySubSeqNo= 18
	Const C_LotFlg		= 19
	Const C_MvmtNo		= 20
	
    Dim iCommandSent
    Dim I2_m_mvmt_type_io_type_cd
    Dim I3_b_pur_grp_pur_grp
    Dim I4_b_biz_partner_bp_cd 
    Dim I5_m_pur_goods_mvmt
    Dim IG1_imp_group

    Dim  E6_b_auto_numbering 

    'M41511_MAINT_PUR_CHILD_ISSUE_SVR    =>    M674
	Const M674_I5_mvmt_rcpt_no = 0   
    Const M674_I5_mvmt_rcpt_dt = 1
    Const M674_I5_gm_no = 2
    Const M674_I5_gm_seq_no = 3
    ReDim I5_m_pur_goods_mvmt(M674_I5_gm_seq_no)
    
    Const M674_IG1_I1_po_no = 0  
    Const M674_IG1_I2_po_seq_no = 1  
    Const M674_IG1_I3_count = 2   
    Const M674_IG1_I4_plant_cd = 3 
    Const M674_IG1_I5_item_cd = 4  
    Const M674_IG1_I6_pr_no = 5    
    Const M674_IG1_I7_resvd_seq_no = 6 
    Const M674_IG1_I8_sub_seq_no = 7   
    Const M674_IG1_I9_mvmt_no = 8   
    Const M674_IG1_I9_lot_no = 9
    Const M674_IG1_I9_lot_sub_no = 10
    Const M674_IG1_I9_mvmt_rcpt_qty = 11
    Const M674_IG1_I9_mvmt_rcpt_unit = 12
    Const M674_IG1_I9_mvmt_rcpt_sl_cd = 13
    'Tracking No 추가(2003.06.13)
    Const M674_IG1_I9_tracking_no = 14
    Const M674_IG1_I9_ext1_cd = 15
    Const M674_IG1_I9_ext1_qty = 16
    Const M674_IG1_I9_ext1_amt = 17
    Const M674_IG1_I9_ext1_rt = 18
    Const M674_IG1_I9_ext2_cd = 19
    Const M674_IG1_I9_ext2_qty = 20
    Const M674_IG1_I9_ext2_amt = 21
    Const M674_IG1_I9_ext2_rt = 22
    Const M674_IG1_I9_ext3_cd = 23
    Const M674_IG1_I9_ext3_qty = 24
    Const M674_IG1_I9_ext3_amt = 25
    Const M674_IG1_I9_ext3_rt = 26
    
    Const M674_E3_io_type_cd = 0 
    Const M674_E3_io_type_nm = 1
    Const M674_E3_rcpt_flg = 2
    Const M674_E3_ret_flg = 3
    
    Const M674_E4_pur_grp = 0    
    Const M674_E4_pur_grp_nm = 1
    
    Const M674_E5_bp_cd = 0    
    Const M674_E5_bp_nm = 1
    
	Dim iLngRowCnt ', iLngRowCnt_D
	Dim iLngTempCnt '_C, iLngTempCnt_D
	

	Call RemovedivTextArea()
	
	If Len(Request("txtGmDt")) Then
		If UNIConvDate(Request("txtGmDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Response.End	
		End If
	End If
	
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
    
	lgIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 
	'iLngMaxRow = CInt(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수 
'    iLngMaxRow = iCUCount + iDCount											'☜: 최대 업데이트된 갯수 
    
'    arrTemp = Split(itxtSpread, gRowSep)									'☆: Spread Sheet 내용을 담고 있는 Element명 
	    
    lGrpCnt = 0
    '-------****
    '##################################################################
    If lgIntFlgMode = OPMD_CMODE Then
	   iCommandSent								= "CREATE"
	Else 
	   iCommandSent								= "DELETE"		
	End if
	'----------**
	
	I5_m_pur_goods_mvmt(M674_I5_mvmt_rcpt_no)	= Trim(Request("txtGrNo1"))
	I5_m_pur_goods_mvmt(M674_I5_mvmt_rcpt_dt)	= UNIConvDate(Request("txtGmDt"))
	I2_m_mvmt_type_io_type_cd 					= Trim(Request("txtMvmtType"))
	I4_b_biz_partner_bp_cd						= Trim(Request("txtSupplierCd"))
	I3_b_pur_grp_pur_grp 						= Trim(Request("txtGroupCd"))
				
	I5_m_pur_goods_mvmt(M674_I5_gm_seq_no) = 0
	
	Set iM41511 = Server.CreateObject("PM7G511.cMMntPurChldIssueS")    

   	If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
	End If
		
	Call iM41511.M_MAINT_PUR_CHILD_ISSUE_SVR(gStrGlobalCollection, iCommandSent, _
		             I2_m_mvmt_type_io_type_cd, I3_b_pur_grp_pur_grp,I4_b_biz_partner_bp_cd,  _
		             I5_m_pur_goods_mvmt, itxtSpread, E6_b_auto_numbering,iErrorPosition)
    
	'If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
	'   Set iM41511 = Nothing
	'   Exit Sub
	'End If 
	
	'If CheckSYSTEMError2(Err, True, ,"","","","") = True Then
	If CheckSYSTEMError(Err,True) = True Then
	   Set iM41511 = Nothing
	   Exit Sub
	End If 
	                 
	Set iM41511 = Nothing
	
	'##################################################################
'    For iLngRow = 1 To iLngMaxRow
'
 '    	arrVal = Split(arrTemp(iLngRow-1), gColSep)
'		strStatus = arrVal(0)														'☜: Row 의 상태 
'		iLngRowCnt = CLng(iLngRowCnt) + 1
'
'	Next 
'
'	ReDim IG1_imp_group(iLngRowCnt -1, M674_IG1_I9_ext3_rt)
'		
 '   For iLngRow = 1 To iLngMaxRow
    
'	 	arrVal = Split(arrTemp(iLngRow-1), gColSep)
'		strStatus = arrVal(0)														'☜: Row 의 상태 
 '       iLngTempCnt = CLng(iLngTempCnt) + 1				

'		IF strStatus = "D" Then
'			
'			iCommandSent = "DELETE"
			
'			I5_m_pur_goods_mvmt(M674_I5_mvmt_rcpt_no)	= Trim(Request("txtGrNo1"))
'				
'			IG1_imp_group(iLngTempCnt-1, M674_IG1_I9_mvmt_no)	= Trim(arrVal(C_MvmtNo))
'			
'			I5_m_pur_goods_mvmt(M674_I5_mvmt_rcpt_dt)	= UNIConvDate(Request("txtGmDt"))
'			I2_m_mvmt_type_io_type_cd 					= Trim(Request("txtMvmtType"))
'			I4_b_biz_partner_bp_cd						= Trim(Request("txtSupplierCd"))
'			I3_b_pur_grp_pur_grp 						= Trim(Request("txtGroupCd"))
				
'			I5_m_pur_goods_mvmt(M674_I5_gm_seq_no) = 0
'			
'		Else
'		
'			iCommandSent = "CREATE"
'			
'			I5_m_pur_goods_mvmt(M674_I5_mvmt_rcpt_no)	= Trim(Request("txtGrNo1"))
'				
'			IG1_imp_group(iLngTempCnt-1, M674_IG1_I4_plant_cd)			= Trim(arrVal(C_PlantCd))
'			IG1_imp_group(iLngTempCnt-1, M674_IG1_I5_item_cd)			= Trim(arrVal(C_ItemCd))
'			IG1_imp_group(iLngTempCnt-1, M674_IG1_I9_mvmt_rcpt_qty)		= UNIConvNum(arrVal(C_SubQty),0)
'			IG1_imp_group(iLngTempCnt-1, M674_IG1_I9_mvmt_rcpt_unit)	= Trim(arrVal(C_Unit))
'			IG1_imp_group(iLngTempCnt-1, M674_IG1_I9_mvmt_rcpt_sl_cd)	= Trim(arrVal(C_SLCd))
'		
'			if arrVal(C_LotNo) <> "" then
'				IG1_imp_group(iLngTempCnt-1, M674_IG1_I9_lot_no)		= Trim(arrVal(C_LotNo))
'			else
'				IG1_imp_group(iLngTempCnt-1, M674_IG1_I9_lot_no)		= "*"
'			End if
'		
'			IG1_imp_group(iLngTempCnt-1, M674_IG1_I9_lot_sub_no)		= UNIConvNum(arrVal(C_LotNoSeq),0)
'			IG1_imp_group(iLngTempCnt-1, M674_IG1_I1_po_no)				= Trim(arrVal(C_PoNo))
'			IG1_imp_group(iLngTempCnt-1, M674_IG1_I2_po_seq_no)			= UNIConvNum(arrVal(C_PoSeq),0)
'			IG1_imp_group(iLngTempCnt-1, M674_IG1_I6_pr_no)				= Trim(arrVal(C_PrNo))
'			IG1_imp_group(iLngTempCnt-1, M674_IG1_I7_resvd_seq_no)		= UNIConvNum(arrVal(C_ResvdSeq),0)
'			IG1_imp_group(iLngTempCnt-1, M674_IG1_I8_sub_seq_no)		= UNIConvNum(arrVal(C_HstySubSeqNo),0)
'			IG1_imp_group(iLngTempCnt-1, M674_IG1_I9_mvmt_no)			= Trim(arrVal(C_MvmtNo))
'			IG1_imp_group(iLngTempCnt-1, M674_IG1_I3_count)				= arrVal(C_MvmtNo+1)
'
			              
 '   		I5_m_pur_goods_mvmt(M674_I5_mvmt_rcpt_dt)	= UNIConvDate(Request("txtGmDt"))
'			I2_m_mvmt_type_io_type_cd 					= Trim(Request("txtMvmtType"))
'			I4_b_biz_partner_bp_cd						= Trim(Request("txtSupplierCd"))
'			I3_b_pur_grp_pur_grp 						= Trim(Request("txtGroupCd"))
'				
'			I5_m_pur_goods_mvmt(M674_I5_gm_seq_no) = 0
'			
'		End IF		
'	Next
'	
'	Set iM41511 = Server.CreateObject("PM7G511.cMMntPurChldIssueS")    
'
 '  	If CheckSYSTEMError(Err,True) = True Then
'		Exit Sub
'	End If
'		
'	Call iM41511.M_MAINT_PUR_CHILD_ISSUE_SVR(gStrGlobalCollection, iCommandSent, _
'		             I2_m_mvmt_type_io_type_cd, I3_b_pur_grp_pur_grp,I4_b_biz_partner_bp_cd,  _
'		             I5_m_pur_goods_mvmt, IG1_imp_group, E6_b_auto_numbering,iErrorPosition)
 '   
'	
'	If CheckSYSTEMError2(Err,True, "","","","","") = True Then
'		Set iM41511 = Nothing
'		Exit Sub
'	End If	                  
'	Set iM41511 = Nothing
	

	Response.Write "<Script Language=VBScript>"		& vbCr
	Response.Write "With parent"					& vbCr
	Response.Write "	If """ & Trim(lgIntFlgMode) & """ = """ & Trim(OPMD_CMODE) & """ Then " & vbCr
	Response.Write "		.frm1.txtGrNo.Value = """ & E6_b_auto_numbering & """ "				& vbCr
	Response.Write "	End If"																	& vbCr
	Response.Write "	.DbSaveOk()"														& vbCr
	Response.Write "End With"						& vbCr
	Response.Write "</Script>"						& vbCr
	
End Sub
'============================================================================================================
Sub SubchangeMvmtType()
	Dim iM14139
	
	Dim I1_m_mvmt_type_io_type_cd
    Dim I2_ief_supplied_SELECT_CHAR
    Dim E1_M_LOOKUP_MVMT_TYPE_SVR
    
    Const M367_E1_io_type_cd = 0    
    Const M367_E1_io_type_nm = 1
    Const M367_E1_mvmt_cd = 2
    Const M367_E1_rcpt_flg = 3
    Const M367_E1_ret_flg = 4
    Const M367_E1_import_flg = 5
    Const M367_E1_subcontra_flg = 6
    Const M367_E1_usage_flg = 7
    Const M367_E1_ext1_cd = 8
    Const M367_E1_ext2_cd = 9
    Const M367_E1_ext3_cd = 10
    Const M367_E1_ext4_cd = 11
    
    On Error Resume Next
    Err.Clear																			'☜: Protect system	from crashing
											
    Set	iM14139 = CreateObject("PM1G439.cMLookupMvmtTypeS")

    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
	End If
    
    I1_m_mvmt_type_io_type_cd = Request("txtMvmtType")
    I2_ief_supplied_SELECT_CHAR = "3"
    
    E1_M_LOOKUP_MVMT_TYPE_SVR = iM14139.M_LOOKUP_MVMT_TYPE_SVR(gStrGlobalCollection, _
				I1_m_mvmt_type_io_type_cd, I2_ief_supplied_SELECT_CHAR)

    If CheckSYSTEMError2(Err,True,"","","","","") = True Then
		Set iM14139 = Nothing
	
		Response.Write "<Script Language=VBScript>" & vbCr
			Response.Write "With parent" & vbCr
			Response.Write "	.frm1.txtMvmtTypeNm.Value 	= """"	" & vbCr
			Response.Write "	.frm1.hdnImportflg.Value 	= """"	" & vbCr
			Response.Write "	.frm1.hdnRcptflg.Value 		= """"	" & vbCr
			Response.Write "	.frm1.hdnRetflg.Value 		= """"	" & vbCr
			Response.Write "	.frm1.hdnSubcontraflg.Value = """"	" & vbCr
			Response.Write "	.frm1.txtMvmtType.focus         	" & vbCr
			Response.Write "End With"	& vbCr
			Response.Write "</Script>"	& vbCr
		Exit Sub
	End If
	
	Set iM14139 = Nothing
	
	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "With parent" & vbCr
	Response.Write "	.frm1.txtMvmtType.Value 	= """ & ConvSPChars(E1_M_LOOKUP_MVMT_TYPE_SVR(M367_E1_io_type_cd)) & """" & vbCr
	Response.Write "	.frm1.txtMvmtTypeNm.Value 	= """ & ConvSPChars(E1_M_LOOKUP_MVMT_TYPE_SVR(M367_E1_io_type_nm)) & """" & vbCr
	Response.Write "	.frm1.hdnImportflg.Value 	= """ & ConvSPChars(E1_M_LOOKUP_MVMT_TYPE_SVR(M367_E1_import_flg)) & """" & vbCr
	Response.Write "	.frm1.hdnRcptflg.Value 		= """ & ConvSPChars(E1_M_LOOKUP_MVMT_TYPE_SVR(M367_E1_rcpt_flg)) & """" & vbCr
	Response.Write "	.frm1.hdnRetflg.Value 		= """ & ConvSPChars(E1_M_LOOKUP_MVMT_TYPE_SVR(M367_E1_ret_flg)) & """" & vbCr
	Response.Write "	.frm1.hdnSubcontraflg.Value = """ & ConvSPChars(E1_M_LOOKUP_MVMT_TYPE_SVR(M367_E1_subcontra_flg)) & """" & vbCr
	Response.Write "End With"	& vbCr
	Response.Write "</Script>"		& vbCr

End Sub
'============================================================================================================
'Display CodeName
'============================================================================================================
Sub DisplaySupplierNm(inCode)
	
	On Error Resume Next						
    Err.Clear   
    
	Call SubOpenDB(lgObjConn)
	
	lgStrSQL = ""
	lgStrSQL = lgStrSQL & " SELECT BP_NM FROM B_BIZ_PARTNER " 
	lgStrSQL = lgStrSQL & " WHERE BP_CD =  " & FilterVar(inCode , "''", "S") & " AND Bp_Type in (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") AND usage_flag=" & FilterVar("Y", "''", "S") & "  AND  in_out_flag = " & FilterVar("O", "''", "S") & "  "											'사외거래처만		
	
	If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X")  then
	
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "With parent.frm1" & vbCr
		Response.Write "	.txtSupplierNm.value	=	""" & lgObjRs("BP_NM") & """ " & vbCr
		Response.Write "End With" & vbCr
		Response.Write "</Script>" & vbCr
			
		Call SubCloseRs(lgObjRs) 
	Else
		Call DisplayMsgBox("179020", vbInformation, "", "", I_MKSCRIPT)
		
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "With parent.frm1" & vbCr
		Response.Write "	.txtSupplierNm.value	=	"""" " & vbCr
		Response.Write "	.txtSupplierCd.focus			 " & vbCr
		Response.Write "End With" & vbCr
		Response.Write "</Script>" & vbCr
		
	End if
	Call SubCloseDB(lgObjConn) 
End Sub 

'============================================================================================================
'Display DisplayGroupNm
'============================================================================================================
Sub DisplayGroupNm(inCode)
	
	On Error Resume Next						
    Err.Clear   
    
	Call SubOpenDB(lgObjConn)
	
	lgStrSQL = ""
	lgStrSQL = lgStrSQL & " SELECT PUR_GRP_NM FROM B_PUR_GRP " 
	lgStrSQL = lgStrSQL & " WHERE PUR_GRP =  " & FilterVar(inCode , "''", "S") & " AND USAGE_FLG=" & FilterVar("Y", "''", "S") & " "		
	
	If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X")  then
	
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "With parent.frm1" & vbCr
		Response.Write "	.txtGroupNm.value	=	""" & lgObjRs("PUR_GRP_NM") & """ " & vbCr
		Response.Write "End With" & vbCr
		Response.Write "</Script>" & vbCr
			
		Call SubCloseRs(lgObjRs)  
	Else
		Call DisplayMsgBox("125100", vbInformation, "", "", I_MKSCRIPT)
		
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "With parent.frm1" & vbCr
		Response.Write "	.txtGroupNm.value	=	"""" " & vbCr
		Response.Write "	.txtGroupCd.focus			 " & vbCr
		Response.Write "End With" & vbCr
		Response.Write "</Script>" & vbCr
		
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
