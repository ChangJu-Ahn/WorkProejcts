<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : Bacth Posting Data List
'*  3. Program ID           : I1711mb2.asp
'*  4. Program Name         : Batch Posting 항목상세조회 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'                             
'*  7. Modified date(First) : 2001/05/08
'*  8. Modified date(Last)  : 2004/10/26
'*  9. Modifier (First)     : Lee Seung Wook
'* 10. Modifier (Last)      : Lee Seung Wook
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
Dim PI1G191

Dim LngMaxRow
Dim LngRow
Dim strData
Dim PvArr

Dim StrNextKey	
Dim lgStrPrevKey

Const C_SHEETMAXROWS_D = 300

Dim I1_inv_close_yyyymm
Dim I2_ief_supplied_select_char

Dim I3_i_goods_mvmt_header
	Const I134_I3_item_document_no = 0
	Const I134_I3_biz_area_cd = 1
Redim I3_i_goods_mvmt_header(I134_I3_biz_area_cd)
	
Dim I4_cost_document_fg
Dim I5_mov_type

Dim EG1_group_export
	Const I134_EG1_E1_i_goods_movement_header_item_document_no = 0
	Const I134_EG1_E1_i_goods_movement_header_document_dt = 1
	Const I134_EG1_E1_i_goods_movement_header_pos_dt = 2
	Const I134_EG1_E1_i_goods_movement_header_gl_no = 3
	
Dim E1_item_document_no
	
	
	lgStrPrevKey = Request("lgStrPrevKey")
	
	I1_inv_close_yyyymm									 = Request("txtDocumentDt")
	I2_ief_supplied_select_char                          = "D"
	I3_i_goods_mvmt_header(I134_I3_biz_area_cd)			 = Request("txtBizCd")
	I4_cost_document_fg									 = Request("txtFlag")
	I5_mov_type											 = Request("txtMoveType")
	
	if lgStrPrevKey <> "" then
       I3_i_goods_mvmt_header(I134_I3_item_document_no) = lgStrPrevKey
    End if
	

	Set PI1G191 = Server.CreateObject("PI1G191.cILstGoodMvmtBchPst")
    
	If CheckSYSTEMError(Err, True) = True Then
		Set PI1G191 = Nothing
		Response.End
	End If
	    
	Call PI1G191.I_LIST_GOODS_BCH_POST_DETAIL(gStrGlobalCollection, _
											C_SHEETMAXROWS_D, _
											I1_inv_close_yyyymm, _
											I2_ief_supplied_select_char, _
											I3_i_goods_mvmt_header, _
											I4_cost_document_fg, _
											I5_mov_type, _
											E1_item_document_no, _
											EG1_group_export)

    If CheckSYSTEMError(Err, True) = True Then
    	Set PI1G191 = Nothing				
		Response.End						
	End If
	
	Set PI1G191 = Nothing
	
	If EG1_group_export(Ubound(EG1_group_export, 1),I134_EG1_E1_i_goods_movement_header_item_document_no) = E1_item_document_no Then
		StrNextKey = ""
	Else
		StrNextKey = E1_item_document_no
	End If

	LngMaxRow = CLng(Request("txtMaxRows")) + 1
	ReDim PvArr(ubound(EG1_group_export,1))
	
	For LngRow = 0 To ubound(EG1_group_export,1)
		
	    if LngRow > C_SHEETMAXROWS_D Then
		Exit For
	    End If 
	
		strData = Chr(11) & ConvSPChars(EG1_group_export(LngRow, I134_EG1_E1_i_goods_movement_header_item_document_no)) & _
				  Chr(11) & UniDateClientFormat(EG1_group_export(LngRow, I134_EG1_E1_i_goods_movement_header_document_dt)) & _
				  Chr(11) & UniDateClientFormat(EG1_group_export(LngRow, I134_EG1_E1_i_goods_movement_header_pos_dt)) & _
				  Chr(11) & ConvSPChars(EG1_group_export(LngRow, I134_EG1_E1_i_goods_movement_header_gl_no)) & _
				  Chr(11) & LngMaxRow + LngRow & Chr(11) & Chr(12) 
				  	
		PvArr(LngRow) = strData
	Next
    
    strData = Join(PvArr, "")
    

    Response.Write "<Script Language=vbscript>" & vbCr
    Response.Write "With parent" & vbCr
	Response.Write "   .ggoSpread.Source          = .frm1.vspdData2	" & vbCr
    Response.Write "   .ggoSpread.SSShowData        """ & strData & """" & vbCr
    Response.Write "   .lgStrPrevKey              = """ & ConvSPChars(StrNextKey)	   & """" & vbCr
   	Response.Write "	If .frm1.vspdData2.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData2, 0) And .lgStrPrevKey <> """" Then "	& vbCr
  	Response.Write "		.DbDtlQuery								"				& vbCr
  	Response.Write "    Else								"				& vbCr
  	Response.Write "		.DbDtlQueryOK								"				& vbCr
	Response.Write "    End If								"				& vbCr

	Response.Write "End with " & vbcr
    Response.Write "</Script>	" & vbCr
	Response.End

%>	


