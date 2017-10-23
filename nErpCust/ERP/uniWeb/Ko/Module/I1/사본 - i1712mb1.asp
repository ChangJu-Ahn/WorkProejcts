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
'*  8. Modified date(Last)  : 2004/10/06
'*  9. Modifier (First)     : lee hae ryong
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
Dim strMode

Dim LngMaxRow
Dim LngRow
Dim strData
Dim PvArr
Dim SetComboList, ComboRow, ComboName


Dim I1_inv_close_yyyymm
Dim I2_good_mvmt_workset_trns_type
Dim I3_ief_supplied_select_char
Dim I4_i_biz_area_cd
Dim I5_cost_document_fg

Dim E1_b_biz_area_nm
Dim E3_i_goods_movement_header
	Const I134_E3_mov_type = 0
	Const I134_E3_mov_nm = 1
	Const I134_E3_trns_type = 2
	Const I134_E3_trns_nm = 3
	
	SetComboList	= SetComboSplit(Request("SetComboList"))

	I2_good_mvmt_workset_trns_type                       = UCase(Request("cboTrnsType"))
	I3_ief_supplied_select_char                          = "C"
	I5_cost_document_fg									 = Request("txtFlag")
	I1_inv_close_yyyymm									 = Request("txtDocumentDt")
	I4_i_biz_area_cd									 = Request("txtBizCd")
	

	Set PI1G191 = Server.CreateObject("PI1G191.cILstGoodMvmtBchPst")
    
	If CheckSYSTEMError(Err, True) = True Then
		Response.End
	End If    
	
	Call PI1G191.I_LIST_GOODS_MVMT_BCH_POST(gStrGlobalCollection, _
										I1_inv_close_yyyymm, _
										I2_good_mvmt_workset_trns_type, _
										I3_ief_supplied_select_char, _
										I4_i_biz_area_cd, _
										I5_cost_document_fg, _
										E1_b_biz_area_nm, _
										E3_i_goods_movement_header)


    If CheckSYSTEMError(Err, True) = True Then
    	Set PI1G191 = Nothing				
		Response.End						
	End If
	
	Set PI1G191 = Nothing


	if isEmpty(E3_i_goods_movement_header) then
		Response.End					
	end if    
		
	strData = ""
	LngMaxRow = CLng(Request("txtMaxRows")) + 1
	ReDim PvArr(ubound(E3_i_goods_movement_header,1))
	
	For LngRow = 0 To ubound(E3_i_goods_movement_header,1)

		For ComboRow = 0 To Ubound(SetComboList, 2)
			If UCase(Trim(SetComboList(0, ComboRow))) = UCase(Trim(E3_i_goods_movement_header(LngRow, I134_E3_trns_nm)))  Then
				ComboName = Trim(SetComboList(1, ComboRow))
				Exit For
			End If
		Next
			
		strData = Chr(11) & "0" & _
				  Chr(11) & ConvSPChars(E3_i_goods_movement_header(LngRow, I134_E3_mov_type)) & _
				  Chr(11) & ConvSPChars(E3_i_goods_movement_header(LngRow, I134_E3_mov_nm)) & _
				  Chr(11) & ConvSPChars(E3_i_goods_movement_header(LngRow, I134_E3_trns_nm)) & _
				  Chr(11) & LngMaxRow + LngRow & Chr(11) & Chr(12) 
		
		PvArr(LngRow) = strData
	Next
    
    strData = Join(PvArr, "")
    

    Response.Write "<Script Language=vbscript>" & vbCr
    Response.Write "With parent" & vbCr
    Response.Write "	.frm1.txtBizNm.value    = """ & ConvSPChars(E1_b_biz_area_nm) & """" & vbCr
	Response.Write "   .ggoSpread.Source          = .frm1.vspdData	" & vbCr
    Response.Write "   .ggoSpread.SSShowData        """ & strData & """" & vbCr
    
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


