<%@LANGUAGE = VBScript%>
<%Option Explicit
'======================================================================================================
'*  1. Module Name          : Inventory
'*  2. Function Name        : Vendor Managed Inventory
'*  3. Program ID           : i1531mb2.asp
'*  4. Program Name         : List Purchase Order Informationr for VMI
'*  5. Program Desc         : List Purchase Order Informationr for VMI
'*  6. Modified date(First) : 2003-01-07
'*  7. Modified date(Last)  : 2003-01-07
'*  8. Modifier (First)     : Ahn, Jung Je
'*  9. Modifier (Last)      : Ahn, Jung Je
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "I", "NOCOOKIE","MB")

Err.Clear
On Error Resume Next													
Call HideStatusWnd 
Dim pPI5S220													'☆ : 조회용 ComProxy Dll 사용 변수 

Dim strData
Dim PvArr
Dim GroupCount
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow    

Dim RemainQty

Const C_SHEETMAXROWS_D = 100
   '-----------------------
    'IMPORTS View
    '-----------------------
	Dim I1_m_pur_ord_dtl
		Const I512_I1_plant_cd = 0
		Const I512_I1_rcpt_type = 1
		Const I512_I1_item_cd = 2
		Const I512_I1_sl_cd = 3
		Const I512_I1_tracking_no = 4
		Const I512_I1_bp_cd = 5
		Const I512_I1_po_no = 6
		Const I512_I1_po_seq_no = 7
	ReDim I1_m_pur_ord_dtl(I512_I1_po_seq_no)

	'-----------------------
	'EXPORTS View
	'-----------------------
	Dim	EG1_m_pur_ord_dtl
		Const I512_EG1_bp_cd = 0
		Const I512_EG1_bp_nm = 1
		Const I512_EG1_rcpt_type = 2
		Const I512_EG1_po_no = 3
		Const I512_EG1_po_seq_no = 4
		Const I512_EG1_po_unit = 5
		Const I512_EG1_po_remain_qty = 6
		Const I512_EG1_po_trans_qty = 7

	Dim	EG2_i_vmi_onhand_stock
		Const I512_EG2_sl_cd = 0
		Const I512_EG2_sl_nm = 1
		Const I512_EG2_good_onhand_qty = 2
		Const I512_EG2_lot_no = 3
		Const I512_EG2_lot_sub_no = 4
	
	Dim EG3_next_list
		Const I512_EG3_bp_cd = 0
		Const I512_EG3_po_no = 1
		Const I512_EG3_po_seq_no = 2

    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
	I1_m_pur_ord_dtl(I512_I1_plant_cd)		= Request("txtPlantCd")
	I1_m_pur_ord_dtl(I512_I1_rcpt_type)		= Request("cboMvmtType")
    I1_m_pur_ord_dtl(I512_I1_item_cd)		= Request("txtItemCd")
    I1_m_pur_ord_dtl(I512_I1_sl_cd)			= Request("txtSLCd")
    I1_m_pur_ord_dtl(I512_I1_tracking_no)   = Request("txtTrackingNo")

    If Request("lgStrPrevKey21") <> "" Then
		I1_m_pur_ord_dtl(I512_I1_bp_cd)		= Request("lgStrPrevKey21")
		I1_m_pur_ord_dtl(I512_I1_po_no)		= Request("lgStrPrevKey22")
		I1_m_pur_ord_dtl(I512_I1_po_seq_no)	= Request("lgStrPrevKey23")
    End If

	Set pPI5S220 = Server.CreateObject("PI5S220.cIVMIListPurchaseItem")
    
	If CheckSYSTEMError(Err, True) = True Then
		Response.End												'☜: 비지니스 로직 처리를 종료함 
	End If    
	
	Call pPI5S220.I_VMI_LIST_PURCHASE_ITEM(gStrGlobalCollection, C_SHEETMAXROWS_D, _
										I1_m_pur_ord_dtl, _
										EG1_m_pur_ord_dtl, _
										EG2_i_vmi_onhand_stock, _
										EG3_next_list)

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err, True) = True Then
    	Set pPI5S220 = Nothing												'☜: ComProxy Unload
	'Call ServerMesgBox("Call 실패 ERR_number :" & Err.number & " ERR_description :" & Err.description , vbCritical, I_MKSCRIPT)  
		Response.Write "<Script Language=VBScript>" & vbCrLF
		Response.Write "Call parent.SetFocusToDocument(""M"")" & vbCrLF
		Response.Write "parent.frm1.vspdData1.focus" & vbCrLF
		Response.Write "</Script>" & vbCrLF
		Response.End
	End If

	
	Set pPI5S220 = Nothing
	
	if isEmpty(EG1_m_pur_ord_dtl) then
		Response.End														'☜: 비지니스 로직 처리를 종료함 
	end if

	
	LngMaxRow = CLng(Request("txtMaxRows")) + 1
	GroupCount = ubound(EG1_m_pur_ord_dtl,1)
	
	ReDim PvArr(GroupCount)
	
	For LngRow = 0 To GroupCount

		strData = Chr(11) & ConvSPChars(EG1_m_pur_ord_dtl(LngRow, I512_EG1_bp_cd)) & _
				  Chr(11) & ConvSPChars(EG1_m_pur_ord_dtl(LngRow ,I512_EG1_bp_nm)) & _
				  Chr(11) & ConvSPChars(EG1_m_pur_ord_dtl(LngRow, I512_EG1_po_no)) & _
				  Chr(11) & ConvSPChars(EG1_m_pur_ord_dtl(LngRow, I512_EG1_po_seq_no)) & _
				  Chr(11) & ConvSPChars(EG1_m_pur_ord_dtl(LngRow, I512_EG1_rcpt_type)) & _
				  Chr(11) & ConvSPChars(EG1_m_pur_ord_dtl(LngRow, I512_EG1_po_unit)) & _
				  Chr(11) & UniConvNumberDBToCompany(EG1_m_pur_ord_dtl(LngRow, I512_EG1_po_trans_qty), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _
				  Chr(11) & "" & _
				  Chr(11) & ConvSPChars(EG2_i_vmi_onhand_stock(LngRow, I512_EG2_sl_cd)) & _
				  Chr(11) & "" & _
				  Chr(11) & ConvSPChars(EG2_i_vmi_onhand_stock(LngRow, I512_EG2_sl_nm)) & _
				  Chr(11) & UniConvNumberDBToCompany(EG2_i_vmi_onhand_stock(LngRow, I512_EG2_good_onhand_qty), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _
				  Chr(11) & "" & _
				  Chr(11) & UniConvNumberDBToCompany(EG2_i_vmi_onhand_stock(LngRow, I512_EG2_good_onhand_qty), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _ 
				  Chr(11) & ConvSPChars(EG2_i_vmi_onhand_stock(LngRow, I512_EG2_lot_no)) & _
				  Chr(11) & ConvSPChars(EG2_i_vmi_onhand_stock(LngRow, I512_EG2_lot_sub_no)) & _
				  Chr(11) & "" & Chr(11) & "" & Chr(11) & "" & Chr(11) & LngMaxRow + LngRow & _
				  Chr(11) & LngMaxRow + LngRow & Chr(11) & Chr(12) 

        PvArr(LngRow) = strData

	Next

	strData = Join(PvArr, "")

	Response.Write "<Script Language=vbscript> " & vbCr   
    Response.Write " With Parent "               & vbCr
	Response.Write "	.frm1.cbohMvmtType.value			= """ & ConvSPChars(Request("cboMvmtType")) & """" & vbCr

	Response.Write "	.ggoSpread.Source	= .frm1.vspdData2 "				& vbCr
	Response.Write "	.ggoSpread.SSShowData  """ & strData  & """"        & vbCr

	Response.Write "    .lgStrPrevKey21  = """ & ConvSPChars(EG3_next_list(I512_EG3_bp_cd)) & """" & vbCr  
    Response.Write "    .lgStrPrevKey22  = """ & ConvSPChars(EG3_next_list(I512_EG3_po_no)) & """" & vbCr  
	Response.Write "    .lgStrPrevKey23  = """ & ConvSPChars(EG3_next_list(I512_EG3_po_seq_no)) & """" & vbCr  
	
	Response.Write "	If .frm1.vspdData2.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData2, 0) And .lgStrPrevKey21 <> """" Then "	& vbCr
	Response.Write "		.DbdtlQuery1 """ & LngMaxRow + LngRow  & """" 	& vbCr
	Response.Write "    Else								"				& vbCr
  	Response.Write "		.DbdtlQuery1Ok	""" & LngMaxRow  & """ , """ & LngRow  & """" 	& vbCr
	Response.Write "    End If								"				& vbCr
  	
  	Response.Write " 	.frm1.vspdData2.focus	"				& vbCr
	
	Response.Write " End with	" & vbCr
    Response.Write "</Script>      " & vbCr   
	Response.End     
%>
