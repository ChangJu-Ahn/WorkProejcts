<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : Bacth Posting Data List
'*  3. Program ID           : I1711mb1.asp
'*  4. Program Name         : Batch Posting 항목조회 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'                             
'*  7. Modified date(First) : 2001/05/08
'*  8. Modified date(Last)  : 2001/05/08
'*  9. Modifier (First)     : lee hae ryong
'* 10. Modifier (Last)      : lee hae ryong
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<%	
Call LoadBasisGlobalInf()	

On Error Resume Next											
Call HideStatusWnd
Dim pi17118
Dim strMode

Dim StrNextKey1
Dim StrNextKey2
Dim LngMaxRow
Dim LngRow
Dim strData
Dim PvArr
Dim SetComboList, ComboRow, ComboName

Const C_SHEETMAXROWS_D = 1000
'Const C_SHEETMAXROWS_D = 30000

    Dim I1_good_mvmt_workset_document_dt
    Dim I2_good_mvmt_workset_trns_type
    Dim I3_ief_supplied_select_char
    Dim I4_i_goods_movement_header
		Const I133_I4_item_document_no	= 0
		Const I133_I4_mov_type			= 1
		Const I133_I4_document_dt		= 2
		Const I133_I4_biz_area_cd		= 3    
	reDim I4_i_goods_movement_header(I133_I4_biz_area_cd)

    Dim E1_b_biz_area_nm
    Dim E2_b_minor_nm
    Dim E3_i_goods_movement_header
		Const I133_E3_item_document_no	= 0
		Const I133_E3_trns_type			= 1
    	Const I133_E3_mov_type			= 2
		Const I133_E3_document_dt		= 3
		Const I133_E3_biz_area_cd		= 4    
    Dim EG1_group_export
		Const I133_EG1_E1_i_goods_movement_header_item_document_no	= 0
		Const I133_EG1_E1_i_goods_movement_header_trns_type			= 1
		Const I133_EG1_E1_i_goods_movement_header_document_dt		= 2
		Const I133_EG1_E1_i_goods_movement_header_pos_dt			= 3
		Const I133_EG1_E2_b_minor_minor_nm							= 4
		Const I133_EG1_E1_i_goods_movement_header_gl_no				= 5
	
	StrNextKey1		= Request("lgStrPrevKey")
	StrNextKey2		= Request("lgStrPrevKey2")
	SetComboList	= SetComboSplit(Request("SetComboList"))

	I2_good_mvmt_workset_trns_type                       = UCase(Request("cboTrnsType"))
	I3_ief_supplied_select_char                          = "C"
	I4_i_goods_movement_header(I133_I4_item_document_no) = ""
	I4_i_goods_movement_header(I133_I4_mov_type)		 = UCase(Request("txtMovType"))
	I1_good_mvmt_workset_document_dt                     = UNIConvDate(Request("txtDocumentFrDt"))
	I4_i_goods_movement_header(I133_I4_document_dt)      = UNIConvDate(Request("txtDocumentToDt"))
	I4_i_goods_movement_header(I133_I4_biz_area_cd)      = Request("txtBizCd")


	If StrNextKey1 <> "" and StrNextKey2 <> "" then
    	I4_i_goods_movement_header(I133_I4_item_document_no) = StrNextKey1
    	I1_good_mvmt_workset_document_dt                     = StrNextKey2
    End if

	Set pi17118 = Server.CreateObject("PI1G190.cILstGoodMvmtBchPst")
    
	If CheckSYSTEMError(Err, True) = True Then
		Response.End
	End If    
	
	Call pi17118.I_LIST_GOODS_MVMT_BCH_POST(gStrGlobalCollection, C_SHEETMAXROWS_D, _
										I1_good_mvmt_workset_document_dt, _
										I2_good_mvmt_workset_trns_type, _
										I3_ief_supplied_select_char, _
										I4_i_goods_movement_header, _
										E1_b_biz_area_nm, _
										E2_b_minor_nm, _
										E3_i_goods_movement_header, _
										EG1_group_export)

    If CheckSYSTEMError(Err, True) = True Then
    	Set pi17118 = Nothing				
		Response.End						
	End If

	Set pi17118 = Nothing

	If isEmpty(EG1_group_export) Then
		Response.End					
	End If

	strData = ""
	LngMaxRow = CLng(Request("txtMaxRows")) + 1
	ReDim PvArr(ubound(EG1_group_export,1))
	
	For LngRow = 0 To ubound(EG1_group_export,1)

		For ComboRow = 0 To Ubound(SetComboList, 2)
			If UCase(Trim(SetComboList(0, ComboRow))) = UCase(Trim(EG1_group_export(LngRow, I133_EG1_E1_i_goods_movement_header_trns_type)))  Then
				ComboName = Trim(SetComboList(1, ComboRow))
				Exit For
			End If
		Next
			
		strData = Chr(11) & "0" & _
				  Chr(11) & ConvSPChars(EG1_group_export(LngRow, I133_EG1_E1_i_goods_movement_header_item_document_no)) & _
				  Chr(11) & ComboName & _
				  Chr(11) & ConvSPChars(EG1_group_export(LngRow, I133_EG1_E2_b_minor_minor_nm)) & _
     			  Chr(11) & UniDateClientFormat(EG1_group_export(LngRow, I133_EG1_E1_i_goods_movement_header_document_dt)) & _
     			  Chr(11) & UniDateClientFormat(EG1_group_export(LngRow, I133_EG1_E1_i_goods_movement_header_pos_dt)) & _
     			  Chr(11) & ConvSPChars(EG1_group_export(LngRow, I133_EG1_E1_i_goods_movement_header_gl_no)) & _
				  Chr(11) & LngMaxRow + LngRow & Chr(11) & Chr(12) 
		
		PvArr(LngRow) = strData
	Next
    
    strData = Join(PvArr, "")
    
	If EG1_group_export(ubound(EG1_group_export,1), I133_EG1_E1_i_goods_movement_header_item_document_no) = E3_i_goods_movement_header(I133_E3_item_document_no) and _
	   EG1_group_export(ubound(EG1_group_export,1), I133_EG1_E1_i_goods_movement_header_document_dt) = E3_i_goods_movement_header(I133_E3_document_dt) Then	 
	   
		StrNextKey1 = ""
		StrNextKey2 = ""
	Else
		StrNextKey1 = E3_i_goods_movement_header(I133_E3_item_document_no)
		StrNextKey2 = E3_i_goods_movement_header(I133_E3_document_dt)
	End if

    Response.Write "<Script Language=vbscript>" & vbCr
    Response.Write "With parent" & vbCr
    Response.Write "	.frm1.txtBizNm.value    = """ & ConvSPChars(E1_b_biz_area_nm) & """" & vbCr
	Response.Write "	.frm1.txtMovTypeNm.value  = """ & ConvSPChars(E2_b_minor_nm) & """" & vbCr
	Response.Write "   .ggoSpread.Source          = .frm1.vspdData	" & vbCr
    Response.Write "   .ggoSpread.SSShowData        """ & strData & """" & vbCr
    Response.Write "   .lgStrPrevKey              = """ & ConvSPChars(StrNextKey1)	   & """" & vbCr  
    Response.Write "   .lgStrPrevKey2             = """ & ConvSPChars(StrNextKey2)	   & """" & vbCr  
    
   	Response.Write "	If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> """" Then "	& vbCr
  	Response.Write "		.DbQuery								"				& vbCr
  	Response.Write "    Else								"				& vbCr
  	Response.Write "		.DbQueryOK								"				& vbCr
	Response.Write "    End If								"				& vbCr

	Response.Write "End with " & vbcr
    Response.Write "</Script>	" & vbCr
	Response.End

Function SetComboSplit(ByVal InitCombo)
	Dim ComboList
	Dim InitCode, InitName
	Dim iArrR
	
	ComboList = Split(Initcombo,Chr(12))
	InitCode  = Split(ComboList(0),Chr(11))
	InitName  = Split(ComboList(1),Chr(11))
	
	ReDim ComboList(1, Ubound(InitCode) - 1)
	
	For iArrR = 0 To Ubound(InitCode) - 1
		ComboList(0, iArrR) = InitCode(iArrR)
		ComboList(1, iArrR) = InitName(iArrR)
	Next
	SetComboSplit = ComboList
End Function
%>	


