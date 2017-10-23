<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : VMI 검사결과등록 조회 
'*  3. Program ID           : I1522MB1
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/01/06
'*  8. Modified date(Last)  : 2003/04/25
'*  9. Modifier (First)     : Choi Sung Jae
'* 10. Modifier (Last)      : Ahn Jung Je
'* 11. Comment              : VB Conversion
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")  
Call HideStatusWnd 

Err.Clear
On Error Resume Next

Dim subPI5G115							

Const C_SHEETMAXROWS_D = 100

Dim LngMaxRow		
Dim LngRow
Dim strData
Dim PvArr

Dim I1_I_VMI_GOODS_MVMT_HDR
	Const I505_I1_plant_cd = 0
	Const I505_I1_trns_type = 1
	Const I505_I1_from_document_dt = 2
	Const I505_I1_to_document_dt = 3
ReDim I1_I_VMI_GOODS_MVMT_HDR(I505_I1_to_document_dt)
Dim I2_I_GOODS_MOVEMENT_NEXT
	Const I505_I2_item_document_no = 0
	Const I505_I2_document_year = 1
	Const I505_I2_seq_no = 2
	Const I505_I2_sub_seq_no = 3
ReDim I2_I_GOODS_MOVEMENT_NEXT(I505_I2_sub_seq_no)	

Dim E1_I_VMI_GOODS_MOVEMENT_NEXT
	Const I505_E1_item_document_no = 0
	Const I505_E1_document_year = 1
	Const I505_E1_seq_no = 2
	Const I505_E1_sub_seq_no = 3
Dim EG1_I_VMI_GOODS_MVMT_DTL
	Const I505_EG1_seq_no = 0
	Const I505_EG1_sub_seq_no = 1
	Const I505_EG1_item_cd = 2
	Const I505_EG1_item_nm = 3
	Const I505_EG1_good_qty = 4
	Const I505_EG1_bad_qty = 5
	Const I505_EG1_entry_qty = 6
	Const I505_EG1_entry_unit = 7
	Const I505_EG1_bp_cd = 8
	Const I505_EG1_bp_nm = 9
	Const I505_EG1_sl_cd = 10
	Const I505_EG1_sl_nm = 11
	Const I505_EG1_document_dt = 12
	Const I505_EG1_insp_flag = 13
	Const I505_EG1_tracking_no = 14
	Const I505_EG1_lot_no = 15
	Const I505_EG1_lot_sub_no = 16
	Const I505_EG1_b_item_spec = 17
	Const I505_EG1_base_unit = 18
	Const I505_EG1_item_document_no = 19
	Const I505_EG1_document_year = 20
	Const I505_EG1_insp_req_no = 21

    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
	I1_I_VMI_GOODS_MVMT_HDR(I505_I1_plant_cd)         = Request("txtPlantCd")
	I1_I_VMI_GOODS_MVMT_HDR(I505_I1_trns_type)        = "VR"
	I1_I_VMI_GOODS_MVMT_HDR(I505_I1_from_document_dt) = UNIConvDate(Request("txtDocumentFrDt"))
	I1_I_VMI_GOODS_MVMT_HDR(I505_I1_to_document_dt)   = UNIConvDate(Request("txtDocumentToDt"))
	
	I2_I_GOODS_MOVEMENT_NEXT(I505_I2_item_document_no) = Request("lgStrPrevKey")
	I2_I_GOODS_MOVEMENT_NEXT(I505_E1_document_year)    = Request("lgStrPrevKey1")
	I2_I_GOODS_MOVEMENT_NEXT(I505_I2_seq_no)           = Request("lgStrPrevKey2")
	I2_I_GOODS_MOVEMENT_NEXT(I505_I2_sub_seq_no)       = Request("lgStrPrevKey3")
	    
    LngMaxRow = CLng(Request("txtMaxRows"))

    Set subPI5G115 = Server.CreateObject("PI5G115.cIVMILookUpInspectMvmt")
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
	If CheckSYSTEMError(Err, True) = True Then
		Response.End											
	End If    
    '-----------------------
    'Com Action Area
    '-----------------------
    Call subPI5G115.I_VMI_LOOK_UP_INSPECT_MVMT(gStrGlobalCollection, C_SHEETMAXROWS_D, _
												I1_I_VMI_GOODS_MVMT_HDR, _
												I2_I_GOODS_MOVEMENT_NEXT, _
												E1_I_VMI_GOODS_MOVEMENT_NEXT, _
												EG1_I_VMI_GOODS_MVMT_DTL) 	
    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err, True) = True Then
    	Set subPI5G115 = Nothing											
		Response.End														
	End If

  	Set subPI5G115 = Nothing											

	if isEmpty(EG1_I_VMI_GOODS_MVMT_DTL) then
		Response.End
	end if 

	ReDim Preserve PvArr(Ubound(EG1_I_VMI_GOODS_MVMT_DTL,1))

	For LngRow = 0 To Ubound(EG1_I_VMI_GOODS_MVMT_DTL,1)
	    strData = Chr(11) & ConvSPChars(EG1_I_VMI_GOODS_MVMT_DTL(LngRow, I505_EG1_item_cd)) & _
				  Chr(11) & ConvSPChars(EG1_I_VMI_GOODS_MVMT_DTL(LngRow, I505_EG1_item_nm)) & _
				  Chr(11) & UniConvNumberDBToCompany("0", ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
				  Chr(11) & UniConvNumberDBToCompany(EG1_I_VMI_GOODS_MVMT_DTL(LngRow, I505_EG1_good_qty), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
				  Chr(11) & UniConvNumberDBToCompany(EG1_I_VMI_GOODS_MVMT_DTL(LngRow, I505_EG1_bad_qty), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
				  Chr(11) & UniConvNumberDBToCompany(EG1_I_VMI_GOODS_MVMT_DTL(LngRow, I505_EG1_entry_qty), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
				  Chr(11) & ConvSPChars(EG1_I_VMI_GOODS_MVMT_DTL(LngRow, I505_EG1_entry_unit)) & _
				  Chr(11) & ConvSPChars(EG1_I_VMI_GOODS_MVMT_DTL(LngRow, I505_EG1_bp_cd)) & _
				  Chr(11) & ConvSPChars(EG1_I_VMI_GOODS_MVMT_DTL(LngRow, I505_EG1_bp_nm)) & _
				  Chr(11) & ConvSPChars(EG1_I_VMI_GOODS_MVMT_DTL(LngRow, I505_EG1_sl_cd)) & _
				  Chr(11) & ConvSPChars(EG1_I_VMI_GOODS_MVMT_DTL(LngRow, I505_EG1_sl_nm)) & _
				  Chr(11) & UNIDateClientFormat(EG1_I_VMI_GOODS_MVMT_DTL(LngRow, I505_EG1_document_dt))
		
		If Trim(EG1_I_VMI_GOODS_MVMT_DTL(LngRow, I505_EG1_insp_flag)) = "Y" Then                                      'C_InspFlg
			strData = strData & Chr(11) & "1"
		Else
			strData = strData & Chr(11) & "0"
	    End If
	    
	    strData = strData & Chr(11) & ConvSPChars(EG1_I_VMI_GOODS_MVMT_DTL(LngRow, I505_EG1_tracking_no)) & _
							Chr(11) & ConvSPChars(EG1_I_VMI_GOODS_MVMT_DTL(LngRow, I505_EG1_lot_no)) & _
							Chr(11) & ConvSPChars(EG1_I_VMI_GOODS_MVMT_DTL(LngRow, I505_EG1_lot_sub_no)) & _
							Chr(11) & ConvSPChars(EG1_I_VMI_GOODS_MVMT_DTL(LngRow, I505_EG1_b_item_spec)) & _
							Chr(11) & ConvSPChars(EG1_I_VMI_GOODS_MVMT_DTL(LngRow, I505_EG1_base_unit)) & _
							Chr(11) & ConvSPChars(EG1_I_VMI_GOODS_MVMT_DTL(LngRow, I505_EG1_item_document_no)) & _
							Chr(11) & ConvSPChars(EG1_I_VMI_GOODS_MVMT_DTL(LngRow, I505_EG1_document_year)) & _
							Chr(11) & ConvSPChars(EG1_I_VMI_GOODS_MVMT_DTL(LngRow, I505_EG1_seq_no)) & _
							Chr(11) & ConvSPChars(EG1_I_VMI_GOODS_MVMT_DTL(LngRow, I505_EG1_sub_seq_no)) & _
							Chr(11) & ConvSPChars(EG1_I_VMI_GOODS_MVMT_DTL(LngRow, I505_EG1_insp_req_no)) & _
							Chr(11) & LngMaxRow + LngRow & Chr(11) & Chr(12)
		PvArr(LngRow) = strData
	Next
	
	strData = Join(PvArr, "")
	
	Response.Write "<Script Language=vbscript>" & vbcr
	Response.Write "With Parent" & vbcr
	Response.Write "    .ggoSpread.Source = .frm1.vspdData" & vbcr
	Response.Write "    .ggoSpread.SSShowData """ & strData & """" & vbcr

	Response.Write "    .lgStrPrevKey  = """ & ConvSPChars(E1_i_vmi_goods_movement_Next(I505_E1_item_document_no))  & """" & vbcr
	Response.Write "    .lgStrPrevKey1 = """ & ConvSPChars(E1_i_vmi_goods_movement_Next(I505_E1_document_year))     & """" & vbcr
	Response.Write "    .lgStrPrevKey2 = """ & ConvSPChars(E1_i_vmi_goods_movement_Next(I505_E1_seq_no))            & """" & vbcr
	Response.Write "    .lgStrPrevKey3 = """ & ConvSPChars(E1_i_vmi_goods_movement_Next(I505_E1_sub_seq_no))        & """" & vbcr
	Response.Write "	If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> """" Then "	& vbCr
	Response.Write "		.DbQuery								"				& vbCr
	Response.Write "    Else								"				& vbCr
	Response.Write "		.DbQueryOK								"				& vbCr
	Response.Write "    End If								"				& vbCr
	Response.Write "End With"       & vbcr
	Response.Write "</Script>"      & vbcr
   	Response.End	

%>

