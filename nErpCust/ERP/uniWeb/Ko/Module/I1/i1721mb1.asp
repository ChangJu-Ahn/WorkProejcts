<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : Bacth Posting Data List
'*  3. Program ID           : I1721mb1.asp
'*  4. Program Name         : Batch Posting Cancel항목조회 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'                             
'*  7. Modified date(First) : 2001/05/14
'*  8. Modified date(Last)  : 2001/05/14
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

Err.Clear
On Error Resume Next											
Call HideStatusWnd
Dim iPI1G190	
Dim strMode		
Dim strData 
Dim PvArr
Dim SetComboList, ComboRow, ComboName

Dim StrNextKey	
Dim StrNextKey2	
Dim lgStrPrevKey
Dim lgStrPrevKey2
Dim LngMaxRow	
Dim LngRow

Const C_SHEETMAXROWS_D = 100

Dim I1_good_mvmt_workset_document_dt
Dim I2_good_mvmt_workset_trns_type
Dim I3_ief_supplied_select_char
Dim I4_i_goods_movement_header
    Const I134_I4_item_document_no = 0
    Const I134_I4_mov_type = 1
    Const I134_I4_document_dt = 2
    Const I134_I4_biz_area_cd = 3
ReDim I4_i_goods_movement_header(I134_I4_biz_area_cd)

Dim E1_b_biz_area_nm
Dim E2_b_minor_nm
Dim E3_i_goods_movement_header
    Const I134_E3_item_document_no = 0
    Const I134_E3_trns_type = 1
    Const I134_E3_mov_type = 2
    Const I134_E3_document_dt = 3
    Const I134_E3_biz_area_cd = 4
Dim EG1_group_export
    Const I134_EG1_E1_i_goods_movement_header_item_document_no = 0
    Const I134_EG1_E1_i_goods_movement_header_trns_type        = 1
    Const I134_EG1_E1_i_goods_movement_header_document_dt      = 2
    Const I134_EG1_E1_i_goods_movement_header_pos_dt           = 3
    Const I134_EG1_E2_b_minor_minor_nm                         = 4
	Const I134_EG1_E1_i_goods_movement_header_gl_no			   = 5    

	lgStrPrevKey = Request("lgStrPrevKey")
	lgStrPrevKey2 = UNIConvDate(Request("lgStrPrevKey2"))
	SetComboList  = SetComboSplit(Request("SetComboList"))

	'-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    I3_ief_supplied_select_char  = "D"
    I4_i_goods_movement_header(I134_I4_biz_area_cd) = Request("txtBizCd")
    I2_good_mvmt_workset_trns_type                  = Request("cboTrnsType")
    I4_i_goods_movement_header(I134_I4_mov_type)    = Request("txtMovType")
    I4_i_goods_movement_header(I134_I4_document_dt) = UNIConvDate(Request("txtDocumentToDt"))
    I1_good_mvmt_workset_document_dt                = UNIConvDate(Request("txtDocumentFrDt"))
    
    If lgStrPrevKey <> "" And lgStrPrevKey2 <> "" Then
		I4_i_goods_movement_header(I134_I4_item_document_no) = lgStrPrevKey
		I1_good_mvmt_workset_document_dt = lgStrPrevKey2
    End If

    Set iPI1G190 = Server.CreateObject("PI1G190.cILstGoodMvmtBchPst")

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = True Then
        Set iPI1G190 = Nothing			
		Response.End						
    End If
      
    '-----------------------
    'Com Action Area
    '-----------------------
    Call iPI1G190.I_LIST_GOODS_MVMT_BCH_POST (gStrGlobalCollection, C_SHEETMAXROWS_D, _
											I1_good_mvmt_workset_document_dt, _
											I2_good_mvmt_workset_trns_type, _
											I3_ief_supplied_select_char, _
											I4_i_goods_movement_header, _
											E1_b_biz_area_nm, _
											E2_b_minor_nm, _
											E3_i_goods_movement_header, _
											EG1_group_export)
	
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = True Then
		Set iPI1G190 = Nothing						
		Response.End								
    End If
	
	Set iPI1G190 = Nothing							
    
 	If EG1_group_export(Ubound(EG1_group_export, 1), I134_EG1_E1_i_goods_movement_header_item_document_no) = E3_i_goods_movement_header(I134_E3_item_document_no) and _
	   EG1_group_export(Ubound(EG1_group_export, 1), I134_EG1_E1_i_goods_movement_header_document_dt) = E3_i_goods_movement_header(I134_E3_document_dt) Then
	   
		StrNextKey = ""
		StrNextKey2 = ""
	else
		StrNextKey = E3_i_goods_movement_header(I134_E3_item_document_no)
		StrNextKey2 = UNIDateClientFormat(E3_i_goods_movement_header(I134_E3_document_dt))
	End If

	LngMaxRow = CLng(Request("txtMaxRows")) + 1
	ReDim PvArr(Ubound(EG1_group_export, 1))
  	
  	For LngRow = 0 To Ubound(EG1_group_export, 1)
		
		For ComboRow = 0 To Ubound(SetComboList, 2)
			If UCase(Trim(SetComboList(0, ComboRow))) = UCase(Trim(EG1_group_export(LngRow, I134_EG1_E1_i_goods_movement_header_trns_type)))  Then
				ComboName = Trim(SetComboList(1, ComboRow))
				Exit For
			End If
		Next
		
		strData = Chr(11) & "0" & _
				  Chr(11) & ConvSPChars(EG1_group_export(LngRow, I134_EG1_E1_i_goods_movement_header_item_document_no)) & _
		          Chr(11) & ComboName & _
				  Chr(11) & ConvSPChars(EG1_group_export(LngRow, I134_EG1_E2_b_minor_minor_nm)) & _
				  Chr(11) & UniDateClientFormat(EG1_group_export(LngRow, I134_EG1_E1_i_goods_movement_header_document_dt)) & _
				  Chr(11) & UniDateClientFormat(EG1_group_export(LngRow, I134_EG1_E1_i_goods_movement_header_pos_dt)) & _
				  Chr(11) & ConvSPChars(EG1_group_export(LngRow, I134_EG1_E1_i_goods_movement_header_gl_no)) & _				  
				  Chr(11) & LngMaxRow + LngRow & Chr(11) & Chr(12)
		
		PvArr(LngRow) = strData
    Next
	strData = Join(PvArr, "")


    Response.Write "<Script Language=vbscript>" & vbCr
    Response.Write "With parent" & vbCr
    Response.Write "	.frm1.txtBizNm.value    = """ & ConvSPChars(E1_b_biz_area_nm) & """" & vbCr
	Response.Write "	.frm1.txtMovTypeNm.value  = """ & ConvSPChars(E2_b_minor_nm) & """" & vbCr
	Response.Write "   .ggoSpread.Source          = .frm1.vspdData	" & vbCr
    Response.Write "   .ggoSpread.SSShowData        """ & strData & """" & vbCr
    Response.Write "   .lgStrPrevKey              = """ & ConvSPChars(StrNextKey)	   & """" & vbCr  
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

