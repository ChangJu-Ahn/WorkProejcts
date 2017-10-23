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
'*  3. Program ID           : m1311mb1
'*  4. Program Name         : PL등록 
'*  5. Program Desc         :
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2000/08/29
'*  8. Modified date(Last)  : 2003/06/12
'*  9. Modifier (First)     : Shin Jin Hyun
'* 10. Modifier (Last)      : Kim Jin Ha
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :	
'**********************************************************************************************

	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("*", "M","NOCOOKIE", "MB")
	Call HideStatusWnd                                                               '☜: Hide Processing message
     
	Dim lgOpModeCRUD
	
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgOpModeCRUD  = Request("txtMode")                                               '☜: Read Operation Mode (CRUD)
	
	Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call  SubBizQueryMulti()
        Case CStr(UID_M0002),CStr(UID_M0005)                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
        Case "changeItemPlant"                                                       '☜: Delete
             Call SubChangeItemPlant()
        Case "changePlntCd" 
			 Call DisplayPlntNm(Request("FLG"),Request("txtPlantCd"))	
		Case "changeItemCd" 
			 Call DisplayItemNm(Request("FLG"),Request("txtPlantCd"),Request("txtItemCd"))
		Case "changeSpplCd" 
			 Call DisplaySupplierNm(Request("FLG"),Request("txtSupplierCd"))	 
	End Select

'============================================================================================================
Sub SubBizQueryMulti()

	Const C_SHEETMAXROWS_D  = 100
	
	Dim TmpBuffer
	Dim iMax
	Dim iIntLoopCount
	Dim iTotalStr
	
	Const M325_pl_no = 0
    Const M325_ref_bom_no = 1
    Const M325_usage_flg = 2
    Const M325_valid_from_dt = 3
    Const M325_valid_to_dt = 4
    Const M325_ext1_cd = 5
    Const M325_ext2_cd = 6

    Const M325_plant_cd = 0
    Const M325_plant_nm = 1

    Const M325_item_cd = 0
    Const M325_item_nm = 1

    Const M325_bp_cd = 0
    Const M325_bp_nm = 1
	
	Const M081_EG1_b_item_item_cd = 0
 	Const M081_EG1_b_item_item_nm = 1
 	Const M081_EG1_m_pl_dtl_par_item_qty = 2
 	Const M081_EG1_m_pl_dtl_par_item_unit = 3
 	Const M081_EG1_m_pl_dtl_child_item_qty = 4
 	Const M081_EG1_m_pl_dtl_child_item_unit = 5
 	Const M081_EG1_m_pl_dtl_sppl_type = 6
 	Const M081_EG1_b_minor_sppl_type_nm = 7
 	Const M081_EG1_m_pl_dtl_pl_seq_no = 8
 	Const M081_EG1_m_pl_dtl_loss_rt = 9
	Const M081_EG1_m_pl_dtl_ext1_cd = 10
	Const M081_EG1_m_pl_dtl_ext1_qty = 11
	Const M081_EG1_m_pl_dtl_ext1_amt = 12
	Const M081_EG1_m_pl_dtl_ext1_rt = 13
	Const M081_EG1_m_pl_dtl_ext2_cd = 14
	Const M081_EG1_m_pl_dtl_ext2_qty = 15
	Const M081_EG1_m_pl_dtl_ext2_amt = 16
	Const M081_EG1_m_pl_dtl_ext2_rt = 17
	Const M081_EG1_m_pl_dtl_ext3_cd = 18
	Const M081_EG1_m_pl_dtl_ext3_qty = 19
	Const M081_EG1_m_pl_dtl_ext3_amt = 20
	Const M081_EG1_m_pl_dtl_ext3_rt = 21

	Const ExportMPlHdrValidFromDt = "1901-01-01"
	Const ExportMPlHdrValidToDt   = "2999-12-31"

	Dim iM13119																	'☆ : 조회용 ComProxy Dll 사용 변수 
	Dim iM13128

	Dim istrMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
	Dim istrPlNo
	Dim istrData
	    
	Dim StrNextKey		' 다음 값 
	Dim lgStrPrevKey	' 이전 값 

	Dim e1_b_biz_partner
	Dim e2_b_item
	Dim e3_b_plant
	Dim e4_m_pl_hdr
	
	Dim E1_m_pl_dtl_pl_seq_no
	Dim export_group
	
	Dim iStrValidFrDt
	Dim iStrValidToDt

	Dim iDesc
	
	Dim Sppl_CD
	Dim Item_CD
	Dim Plant_CD
	Dim GroupCount
	
   On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
	lgStrPrevKey = Request("lgStrPrevKey")

	Set iM13119 = Server.CreateObject("PM1G319.cMLookupPlHdrS")

	If CheckSYSTEMError(Err,True) = true then 		
		Exit Sub
	End if
	
	Sppl_CD = Trim(UCase(Request("txtSpplCd")))
	Item_CD = Trim(UCase(Request("txtitemCd")))
	Plant_CD = Trim(UCase(Request("txtPlantCd")))
	
	Call iM13119.M_LOOKUP_PL_HDR_SVR(gStrGlobalCollection, Sppl_Cd, item_Cd, Plant_Cd, _
									e1_b_biz_partner, e2_b_item, e3_b_plant, e4_m_pl_hdr)

	If Err.Number <> 0 Then
	   iDesc = Split(Err.Description, Chr(11))
	End If

	If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
		Set iM13119 = Nothing												'☜: ComProxy Unload
	    If ubound(iDesc,1) > 0 Then
	      If iDesc(1) = "171500" Then
				Response.Write "<Script Language=vbscript>" & vbCr
				Response.Write "	parent.changeItem()" 	& vbCr
				Response.Write "	parent.dbQueryOkhdr()" 	& vbCr
				Response.Write "</Script>"					& vbCr
				Exit Sub
	      End If	
		End If
		Exit Sub															'☜: 비지니스 로직 처리를 종료함 
	 End If
	 	
	istrPlNo = Trim(e4_m_pl_hdr(M325_pl_no))
	iLngMaxRow = CLng(Request("txtMaxRows"))
	
	'시작일은 "1901-01-01"보다 크야함.
	If cDate(ExportMPlHdrValidFromDt) > cDate(uniConvDate(e4_m_pl_hdr(M325_valid_from_dt))) Then
		iStrValidFrDt = ExportMPlHdrValidFromDt	
	Else		
		iStrValidFrDt = uniConvDate(e4_m_pl_hdr(M325_valid_from_dt))	
	End If

	'종료일은 "2999-12-31"보다 작아야함.	
	If cDate(ExportMPlHdrValidToDt) < cDate(uniConvDate(e4_m_pl_hdr(M325_valid_to_dt))) Then
		iStrValidToDt = ExportMPlHdrValidToDt	
	Else		
		iStrValidToDt = uniConvDate(e4_m_pl_hdr(M325_valid_to_dt))	
	End If

	Response.Write "<Script Language=vbscript>"					& vbCr
	Response.Write "With parent"								& vbCr
	Response.Write "	.frm1.txtPlantNm.value	= """ & ConvSPChars(e3_b_plant(M325_plant_nm))		& """" & vbCr
	Response.Write "	.frm1.txtItemNm.value	= """ & ConvSPChars(e2_b_item(M325_item_nm))		& """" & vbCr
	Response.Write "	.frm1.txtSpplNm.value	= """ & ConvSPChars(e1_b_biz_partner(M325_bp_nm))  	& """" & vbCr
    Response.Write "	.frm1.txtPlantCd.value	= """ & ConvSPChars(e3_b_plant(M325_plant_cd))		& """" & vbCr
	Response.Write "	.frm1.txtPlantNm.value	= """ & ConvSPChars(e3_b_plant(M325_plant_nm))		& """" & vbCr
	Response.Write "	.frm1.txtItemCd.value	= """ & ConvSPChars(e2_b_item(M325_item_cd))		& """" & vbCr
	Response.Write "	.frm1.txtItemNm.value	= """ & ConvSPChars(e2_b_item(M325_item_nm))		& """" & vbCr
	Response.Write "	.frm1.txtSpplCd.value	= """ & ConvSPChars(e1_b_biz_partner(M325_bp_cd))	& """" & vbCr
	Response.Write "	.frm1.txtSpplNm.value	= """ & ConvSPChars(e1_b_biz_partner(M325_bp_nm)) 	& """" & vbCr
	Response.Write "	.frm1.txtPlantCd2.value	= """ & ConvSPChars(e3_b_plant(M325_plant_cd))		& """" & vbCr
	Response.Write "	.frm1.txtPlantNm2.value	= """ & ConvSPChars(e3_b_plant(M325_plant_nm))		& """" & vbCr
	Response.Write "	.frm1.txtItemCd2.value	= """ & ConvSPChars(e2_b_item(M325_item_cd))		& """" & vbCr
	Response.Write "	.frm1.txtItemNm2.value	= """ & ConvSPChars(e2_b_item(M325_item_nm))		& """" & vbCr
	Response.Write "	.frm1.txtSpplCd2.value	= """ & ConvSPChars(e1_b_biz_partner(M325_bp_cd))	& """" & vbCr
	Response.Write "	.frm1.txtSpplNm2.value	= """ & ConvSPChars(e1_b_biz_partner(M325_bp_nm)) 	& """" & vbCr
	Response.Write "	.frm1.txtFrDt.text		= """ & UNIDateClientFormat(iStrValidFrDt)			& """" & vbCr
	Response.Write "	.frm1.txtToDt.text		= """ & UNIDateClientFormat(iStrValidToDt)			& """" & vbCr
	Response.Write "	.frm1.txtBomNo.Value	= """ & ConvSPChars(e4_m_pl_hdr(M325_ref_bom_no)) 	& """" & vbCr
    Response.Write "End With" 	& vbCr
    Response.Write "</Script>" 	& vbCr
     
    Set iM13119 = Nothing
		
	'========================================================================
	'M13119 호출 후 리턴된 값을 가지고 M13128을 다시 호출 List 값을 가져온다.
	'========================================================================
	Set iM13128 = Server.CreateObject("PM1G328.cMListPlDtlS")							'PL
	   
	Call iM13128.M_LIST_PL_DTL_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, istrPlNo, lgStrPrevKey, _
								 E1_m_pl_dtl_pl_seq_no, export_group)	

	If Err.Number <> 0 Then
	   iDesc = Split(Err.Description, Chr(11))
	End If

	If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
		Set iM13128 = Nothing												'☜: ComProxy Unload
	    If ubound(iDesc,1) > 0 Then
	      If iDesc(1) = "171600" Then
				response.write "<Script Language=vbscript>" & vbCr
				response.write "	parent.frm1.hdnPlant.Value = """ & ConvSPChars(UCase(Request("txtPlantCd"))) & """" & vbCr
				response.write "	parent.frm1.hdnItem.Value  = """ & ConvSPChars(UCase(Request("txtitemCd"))) & """" & vbCr
				response.write "	parent.frm1.hdnSppl.Value  = """ & ConvSPChars(UCase(Request("txtSpplCd"))) & """" & vbCr
				response.write "	parent.frm1.hdnPLNo.Value  = """ & ConvSPChars(UCase(istrPlNo)) & """" & vbCr
				response.write "	parent.changeItem()" & vbCr
				response.write "	parent.dbQueryok()" & vbCr
				response.write "</Script>" & vbCr
				Exit Sub
	      End If	
		End If	
		Exit Sub															'☜: 비지니스 로직 처리를 종료함 
	 End If
	
	
	iLngMaxRow = CLng(Request("txtMaxRows"))
	GroupCount = UBound(export_group,1)

	 IF GroupCount <> 0 then
		IF export_group(GroupCount,M081_EG1_m_pl_dtl_pl_seq_no) =  E1_m_pl_dtl_pl_seq_no then
				StrNextKey = ""
		Else
				StrNextKey = E1_m_pl_dtl_pl_seq_no
		End If
	End if
	
	iIntLoopCount = 0
	iMax = UBound(export_group,1)
	ReDim TmpBuffer(iMax)
	
	For iLngRow = 0 To iMax
		If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
		   StrNextKey = E1_m_pl_dtl_pl_seq_no
           Exit For
        End If 
		istrData = ""
		istrData = istrData & Chr(11) & ConvSPChars(export_group(iLngRow, M081_EG1_b_item_item_cd))
	    istrData = istrData & Chr(11) & ""
	    istrData = istrData & Chr(11) & ConvSPChars(export_group(iLngRow, M081_EG1_b_item_item_nm))
	    istrData = istrData & Chr(11) & UNINumClientFormat(export_group(iLngRow, M081_EG1_m_pl_dtl_par_item_qty), 4, 0)
	    istrData = istrData & Chr(11) & ConvSPChars(export_group(iLngRow, M081_EG1_m_pl_dtl_par_item_unit))
	    istrData = istrData & Chr(11) & ""
	    istrData = istrData & Chr(11) & UNINumClientFormat(export_group(iLngRow, M081_EG1_m_pl_dtl_child_item_qty), 4, 0)
	    istrData = istrData & Chr(11) & ConvSPChars(export_group(iLngRow, M081_EG1_m_pl_dtl_child_item_unit))
	    istrData = istrData & Chr(11) & ""
	    istrData = istrData & Chr(11) & ConvSPChars(export_group(iLngRow, M081_EG1_m_pl_dtl_sppl_type))
	    istrData = istrData & Chr(11) & ConvSPChars(export_group(iLngRow, M081_EG1_b_minor_sppl_type_nm))
	    istrData = istrData & Chr(11) & ConvSPChars(export_group(iLngRow, M081_EG1_m_pl_dtl_pl_seq_no))
	    istrData = istrData & Chr(11) & UNINumClientFormat(export_group(iLngRow, M081_EG1_m_pl_dtl_loss_rt), "", 0)
	    istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
	    istrData = istrData & Chr(11) & Chr(12)
	    
	    TmpBuffer(iIntLoopCount) = istrData
        iIntLoopCount = iIntLoopCount + 1
	Next
	
	iTotalStr = Join(TmpBuffer, "")    
	
	Response.Write "<Script Language=vbscript>"					& vbCr
	
	Response.Write "With Parent "								& vbCr
	Response.Write " .ggoSpread.Source = .frm1.vspdData"  		& vbCr
	Response.Write " .ggoSpread.SSShowData     """ & iTotalStr									& """"	& vbCr
	Response.Write " .lgStrPrevKey           = """ & StrNextKey									& """"	& vbCr
	Response.Write " .frm1.hdnPlant.value   = """ & ConvSPChars(UCase(Request("txtPlantCd")))	& """"	& vbCr
	Response.Write " .frm1.hdnItem.value    = """ & ConvSPChars(UCase(Request("txtitemCd")))	& """"	& vbCr
	Response.Write " .frm1.hdnSppl.value    = """ & ConvSPChars(UCase(Request("txtSpplCd")))	& """"	& vbCr
	Response.Write " .frm1.hdnPLNo.Value    = """ & ConvSPChars(UCase(istrPlNo))					& """"	& vbCr
    Response.Write " .changeItem()			"	    			& vbCr
    Response.Write " .DbQueryOk " & vbCr
    Response.Write " .frm1.vspdData.focus	"					& vbCr
    Response.Write "End With"									& vbCr
    Response.Write "</Script>"									& vbCr    
    
	Set iM13128 = Nothing

End Sub    

'============================================================================================================
Sub SubBizSaveMulti()
    On Error Resume Next  				
    Err.Clear																		'☜: Protect system from crashing
    
    Dim iM13111														'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
	Dim iErrorPosition
	Dim lgIntFlgMode
	Dim E1_m_pl_hdr_pl_no
	
	Dim itxtSpreadArr
    Dim itxtSpreadArrCount
    Dim iCUCount
    Dim iDCount
    Dim ii
    Dim itxtSpread
    
	Dim Sppl_CD
	Dim Item_CD
	Dim Plant_CD
	
	Dim iOppFlg
	
	Dim I5_m_pl_hdr
    Const M314_I5_pl_no = 0					'Part List번호 
    Const M314_I5_ref_bom_no = 1			'Bom번호 
    Const M314_I5_usage_flg = 2				'사용여부 
    Const M314_I5_valid_from_dt = 3			'유효시작일 
    Const M314_I5_valid_to_dt = 4			'유효종료일 
    Const M314_I5_ext1_cd = 5				'
    Const M314_I5_ext2_cd = 6

	lgIntFlgMode = CInt(Request("txtFlgMode"))									'☜: 저장시 Create/Update 판별 

	Redim I5_m_pl_hdr(M314_I5_ext2_cd)
	I5_m_pl_hdr(M314_I5_pl_no) 			= Trim(Request("hdnPLNo"))
	I5_m_pl_hdr(M314_I5_ref_bom_no) 	= Trim(Request("txtBomNo"))
	I5_m_pl_hdr(M314_I5_usage_flg) 		= "Y"    	
	I5_m_pl_hdr(M314_I5_valid_from_dt) 	= UNIConvDate(Request("txtFrDt"))
	I5_m_pl_hdr(M314_I5_valid_to_dt) 	= UNIConvDate(Request("txtToDt"))
	
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
    
	Sppl_CD = Trim(UCase(Request("txtSpplCd2")))
	Item_CD = Trim(UCase(Request("txtItemCd2")))
	Plant_CD = Trim(UCase(Request("txtPlantCd2")))
	
	If lgOpModeCRUD=CStr(UID_M0002) Then
		iOppFlg = "CREATE"
	Else
		iOppFlg = "UPDATE"
	End If
	
	Call RemovedivTextArea()
	
	Set iM13111 = Server.CreateObject("PM1G311.cMMaintPartListS")    

    If CheckSYSTEMError(Err,True) = true Then 
		Exit Sub
	End If
	
	Call iM13111.M_MAINT_PART_LIST_SVR(gStrGlobalCollection, iOppFlg, Sppl_Cd, Item_Cd, Plant_Cd, gUsrID, _											
												I5_m_pl_hdr, itxtSpread, E1_m_pl_hdr_pl_no, iErrorPosition)
	
	If CheckSYSTEMError2(Err,True,"","","","","") = true Then 	
		Set iM13111 = Nothing
		Exit Sub														    '☜: 비지니스 로직 처리를 종료함 
	End If

	Set iM13111 = Nothing

	'-----------------------
	'Result data display area
	'----------------------- 
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent " & vbCr 
	Response.Write "	.frm1.txtPlantCd.Value		= """ & UCase(Request("txtPlantCd2")) & """" & vbCr 
	Response.Write "	.frm1.txtItemCd.Value		= """ & UCase(Request("txtItemCd2")) & """" & vbCr 
	Response.Write "	.frm1.txtSpplCd.Value 		= """ & UCase(Request("txtSpplCd2")) & """" & vbCr 
	Response.Write "	.DbSaveOk" & vbCr 
	Response.Write "End With" & vbCr 
	Response.Write "</Script>" & vbCr 
				
End Sub    
'============================================================================================================

Sub SubBizDelete()

	On Error Resume Next  				
    Err.Clear																		'☜: Protect system from crashing
	
	'import P/L Head - Usage Flag를 Yes -> No로 변경함으로 삭제를 대신함..
    Const M314_I5_pl_no = 0					'Part List번호 
    Const M314_I5_ref_bom_no = 1			'Bom번호 
    Const M314_I5_usage_flg = 2				'사용여부 
    Const M314_I5_valid_from_dt = 3			'유효시작일 
    Const M314_I5_valid_to_dt = 4			'유효종료일 
    Const M314_I5_ext1_cd = 5				'
    Const M314_I5_ext2_cd = 6

    Dim iM13111																		'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
	Dim I5_m_pl_hdr
	Dim E1_m_pl_hdr_pl_no, iErrorPosition
	
	Dim Sppl_CD
	Dim Item_CD
	Dim Plant_CD
	Dim TmpSpread
	
	Redim I5_m_pl_hdr(M314_I5_ext2_cd)
	I5_m_pl_hdr(M314_I5_pl_no) 			= Trim(Request("txtPLNo"))
	I5_m_pl_hdr(M314_I5_ref_bom_no) 	= Trim(Request("txtBomNo"))
	I5_m_pl_hdr(M314_I5_usage_flg) 		= "Y"    	
	I5_m_pl_hdr(M314_I5_valid_from_dt) 	= UNIConvDate(Request("txtFrDt"))
	I5_m_pl_hdr(M314_I5_valid_to_dt) 	= UNIConvDate(Request("txtToDt"))
	
    Set iM13111 = Server.CreateObject("PM1G311.cMMaintPartListS")    

    if CheckSYSTEMError(Err,True) = true then 
		Exit Sub
	End If
	
	Sppl_CD = Trim(UCase(Request("txtSpplCd2")))
	Item_CD = Trim(UCase(Request("txtItemCd2")))
	Plant_CD = Trim(UCase(Request("txtPlantCd2")))
	TmpSpread = Request("txtSpread")
	
	Call iM13111.M_MAINT_PART_LIST_SVR(gStrGlobalCollection, _
											"DELETE", _
											Sppl_Cd, _
											Item_Cd, _
											Plant_Cd, _
											"", _
											I5_m_pl_hdr, _
											TmpSpread, _
											E1_m_pl_hdr_pl_no, _
											iErrorPosition)

	If CheckSYSTEMError(Err,True) = true then 	
		set iM13111 = nothing
		Exit Sub														    '☜: 비지니스 로직 처리를 종료함 
	End If
	
	set iM13111 = nothing

	'-----------------------
	'Result data display area
	'----------------------- 
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "	Call parent.DbDeleteOk()" & vbCr
	Response.Write "</Script>" & vbCr 

End Sub

'============================================================================================================

Sub SubChangeItemPlant()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Dim pB1b019
    Dim E5_b_item
    Const P020_E5_item_cd = 0
    Const P020_E5_item_nm = 1
    Const P020_E5_formal_nm = 2
    Const P020_E5_spec = 3
    Const P020_E5_basic_unit = 4
    Const P020_E5_item_acct = 5
    Const P020_E5_item_class = 6
    Const P020_E5_phantom_flg = 7
    Const P020_E5_hs_cd = 8
    Const P020_E5_hs_unit = 9
    Const P020_E5_unit_weight = 10
    Const P020_E5_unit_of_weight = 11
    Const P020_E5_draw_no = 12
    Const P020_E5_item_image_flg = 13
    Const P020_E5_blanket_pur_flg = 14
    Const P020_E5_base_item_cd = 15
    Const P020_E5_proportion_rate = 16
    Const P020_E5_valid_flg = 17
    Const P020_E5_valid_from_dt = 18
    Const P020_E5_valid_to_dt = 19
    Const P020_E5_vat_type = 20
    Const P020_E5_vat_rate = 21
    
    Dim Item_CD
    Set	pB1b019	 = Server.CreateObject("PB3C104.cBLkUpItem")

    If CheckSYSTEMError(Err,True) = true Then 
		Exit Sub
	End If
	
	Item_CD = Trim(UCase(Request("txtItemCd")))
	 
	Call pB1b019.B_LOOK_UP_ITEM(gStrGlobalCollection, Item_Cd, , , , , E5_b_item)
    
    If CheckSYSTEMError(Err,True) = true Then 
		Set pB1b019 = Nothing                                                   '☜: Unload Comproxy
		Exit Sub
	End If

	Response.Write "<Script	Language=vbscript>" & vbCr
	Response.Write "	with parent" & vbCr
	Response.Write "	    .frm1.vspdData.Row = .frm1.vspdData.ActiveRow" & vbCr
	Response.Write "        .frm1.vspdData.Col = .C_ChdItemUnit" & vbCr
	Response.Write "        .frm1.vspdData.text= """ & ConvSPChars(UCase(E5_b_item(P020_E5_basic_unit))) & """" & vbCr
	Response.Write "	end with " & vbCr
	Response.Write "</Script>" & vbCr

	Set pB1b019 = Nothing                                                   '☜: Unload Comproxy

End Sub
'-----------------------
'Display DisplayGroupNm
'-----------------------
Sub DisplayPlntNm(FLG,inCode)
	
	On Error Resume Next						
    Err.Clear   
    
	Call SubOpenDB(lgObjConn)
	
	lgStrSQL = ""
	lgStrSQL = lgStrSQL & " SELECT PLANT_NM FROM B_PLANT " 
	lgStrSQL = lgStrSQL & " WHERE PLANT_CD = " & FilterVar(inCode, "''", "S") & ""
	
	IF FLG = "1" THEN
		If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X")  then
	
			Response.Write "<Script Language=VBScript>" & vbCr
			Response.Write "With parent.frm1" & vbCr
			Response.Write "	.txtPlantNm.value	=	""" & lgObjRs("PLANT_NM") & """ " & vbCr
			Response.Write "End With" & vbCr
			Response.Write "</Script>" & vbCr
				
			Call SubCloseRs(lgObjRs)  
		Else
			Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)
			
			Response.Write "<Script Language=VBScript>" & vbCr
			Response.Write "With parent.frm1" & vbCr
			Response.Write "	.txtPlantCd.value	=	"""" " & vbCr
			Response.Write "	.txtPlantNm.value	=	"""" " & vbCr
			Response.Write "	.txtPlantCd.focus" & vbCr
			Response.Write "End With" & vbCr
			Response.Write "</Script>" & vbCr
		End if
	
	ELSE
		If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X")  then
	
			Response.Write "<Script Language=VBScript>" & vbCr
			Response.Write "With parent.frm1" & vbCr
			Response.Write "	.txtPlantNm2.value	=	""" & lgObjRs("PLANT_NM") & """ " & vbCr
			Response.Write "End With" & vbCr
			Response.Write "</Script>" & vbCr
				
			Call SubCloseRs(lgObjRs)  
		Else
			Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)
			
			Response.Write "<Script Language=VBScript>" & vbCr
			Response.Write "With parent.frm1" & vbCr
			Response.Write "	.txtPlantCd2.value	=	"""" " & vbCr
			Response.Write "	.txtPlantNm2.value	=	"""" " & vbCr
			Response.Write "	.txtPlantCd2.focus" & vbCr
			Response.Write "End With" & vbCr
			Response.Write "</Script>" & vbCr
		End if
	End if
		

End Sub 

'-----------------------
'Display CodeName
'-----------------------
Sub DisplayItemNm(FLG,inCode1, incode2)
	
	On Error Resume Next						
    Err.Clear   
    
	Call SubOpenDB(lgObjConn)
	
	lgStrSQL = ""
	lgStrSQL = lgStrSQL & " SELECT B.ITEM_NM FROM  B_ITEM_BY_PLANT A, B_ITEM B " 
	lgStrSQL = lgStrSQL & " WHERE A.PLANT_CD = " & FilterVar(inCode1, "''", "S") & ""
	lgStrSQL = lgStrSQL & " AND B. ITEM_CD  = " & FilterVar(inCode2, "''", "S") & ""
	
	IF FLG = "1" THEN
		If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X")  then
	
			Response.Write "<Script Language=VBScript>" & vbCr
			Response.Write "With parent.frm1" & vbCr
			Response.Write "	.txtitemNm.value	=	""" & lgObjRs("ITEM_NM") & """ " & vbCr
			Response.Write "End With" & vbCr
			Response.Write "</Script>" & vbCr
				
			Call SubCloseRs(lgObjRs)  
		Else
			Call DisplayMsgBox("122700", vbInformation, "", "", I_MKSCRIPT)
			
			Response.Write "<Script Language=VBScript>" & vbCr
			Response.Write "With parent.frm1" & vbCr
			Response.Write "	.txtItemCd.value	=	"""" " & vbCr
			Response.Write "	.txtItemNm.value	=	"""" " & vbCr
			Response.Write "	.txtItemCd.focus" & vbCr
			Response.Write "End With" & vbCr
			Response.Write "</Script>" & vbCr
			
		End if
	Else 
		If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X")  then
	
			Response.Write "<Script Language=VBScript>" & vbCr
			Response.Write "With parent.frm1" & vbCr
			Response.Write "	.txtitemNm2.value	=	""" & lgObjRs("ITEM_NM") & """ " & vbCr
			Response.Write "End With" & vbCr
			Response.Write "</Script>" & vbCr
				
			Call SubCloseRs(lgObjRs)  
		Else
			Call DisplayMsgBox("122700", vbInformation, "", "", I_MKSCRIPT)
			
			Response.Write "<Script Language=VBScript>" & vbCr
			Response.Write "With parent.frm1" & vbCr
			Response.Write "	.txtItemCd2.value	=	"""" " & vbCr
			Response.Write "	.txtItemNm2.value	=	"""" " & vbCr
			Response.Write "	.txtItemCd2.focus" & vbCr
			Response.Write "End With" & vbCr
			Response.Write "</Script>" & vbCr
			
		End if
	End if
End Sub 

'-----------------------
'Display CodeName
'-----------------------
Sub DisplaySupplierNm(FLG,inCode)
	
	On Error Resume Next						
    Err.Clear   
    
	Call SubOpenDB(lgObjConn)
	
	lgStrSQL = ""
	lgStrSQL = lgStrSQL & " SELECT BP_NM FROM B_BIZ_PARTNER " 
	lgStrSQL = lgStrSQL & " WHERE BP_CD =  " & FilterVar(inCode , "''", "S") & " AND Bp_Type in (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") AND usage_flag=" & FilterVar("Y", "''", "S") & " "		
	
	if FLG = "1" then
		If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X")  then
	
			Response.Write "<Script Language=VBScript>" & vbCr
			Response.Write "With parent.frm1" & vbCr
			Response.Write "	.txtSpplNm.value	=	""" & lgObjRs("BP_NM") & """ " & vbCr
			Response.Write "End With" & vbCr
			Response.Write "</Script>" & vbCr
				
			Call SubCloseRs(lgObjRs)  
		Else
			Call DisplayMsgBox("179020", vbInformation, "", "", I_MKSCRIPT)
			
			Response.Write "<Script Language=VBScript>" & vbCr
			Response.Write "With parent.frm1" & vbCr
			Response.Write "	.txtSpplCd.value	=	"""" " & vbCr
			Response.Write "	.txtSpplNm.value	=	"""" " & vbCr
			Response.Write "	.txtSpplCd.focus" & vbCr
			Response.Write "End With" & vbCr
			Response.Write "</Script>" & vbCr
		End if
	ELSE
		If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X")  then
	
			Response.Write "<Script Language=VBScript>" & vbCr
			Response.Write "With parent.frm1" & vbCr
			Response.Write "	.txtSpplNm2.value	=	""" & lgObjRs("BP_NM") & """ " & vbCr
			Response.Write "End With" & vbCr
			Response.Write "</Script>" & vbCr
				
			Call SubCloseRs(lgObjRs)  
		Else
			Call DisplayMsgBox("179020", vbInformation, "", "", I_MKSCRIPT)
			
			Response.Write "<Script Language=VBScript>" & vbCr
			Response.Write "With parent.frm1" & vbCr
			Response.Write "	.txtSpplCd2.value	=	"""" " & vbCr
			Response.Write "	.txtSpplNm2.value	=	"""" " & vbCr
			Response.Write "	.txtSpplCd2.focus" & vbCr
			Response.Write "End With" & vbCr
			Response.Write "</Script>" & vbCr
		End if
	END if
		
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
