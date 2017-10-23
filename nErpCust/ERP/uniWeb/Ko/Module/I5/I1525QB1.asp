<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : VMI 재고현황조회 
'*  3. Program ID           : I1525QB1
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/01/06
'*  8. Modified date(Last)  : 2003/04/25
'*  9. Modifier (First)     : Choi Sung Jae
'* 10. Modifier (Last)      : Ahn Jung Je
'* 11. Comment              : VB Conversion
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "I","NOCOOKIE","QB")   
Call HideStatusWnd 

Err.Clear
On Error Resume Next
	
Dim iPI5G160								
Dim LngMaxRow		
Dim LngRow
Dim strData
Dim PvArr

Const C_SHEETMAXROWS_D = 100

Dim I1_i_vmi_onhand_stock
	Const I509_I1_plant_cd = 0
	Const I509_I1_sl_cd = 1
	Const I509_I1_bp_cd = 2
	Const I509_I1_item_cd = 3
	Const I509_I1_tracking_no = 4
	Const I509_I1_lot_no = 5
	Const I509_I1_lot_sub_no = 6
ReDim I1_i_vmi_onhand_stock(I509_I1_lot_sub_no)

Dim EG1_i_vmi_onhand_stock
	Const I509_EG1_bp_cd = 0
	Const I509_EG1_bp_nm = 1
	Const I509_EG1_item_cd = 2
	Const I509_EG1_item_nm = 3
	Const I509_EG1_good_on_hand_qty = 4
	Const I509_EG1_basic_unit = 5
	Const I509_EG1_tracking_no = 6
	Const I509_EG1_lot_no = 7
	Const I509_EG1_lot_sub_no = 8
	Const I509_EG1_spec = 9
Dim EG2_next_list
	Const I509_EG2_bp_cd = 0
	Const I509_EG2_item_cd = 1
	Const I509_EG2_tracking_no = 2
	Const I509_EG2_lot_no = 3
	Const I509_EG2_lot_sub_no = 4
    
	
    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
	I1_i_vmi_onhand_stock(I509_I1_plant_cd)    = Request("txtPlantCd")
	I1_i_vmi_onhand_stock(I509_I1_sl_cd)       = Request("txtSlCd")
	If Trim(Request("lgStrPrevKey")) <> "" Then
		I1_i_vmi_onhand_stock(I509_I1_bp_cd)   = Request("lgStrPrevKey")
	Else
		I1_i_vmi_onhand_stock(I509_I1_bp_cd)   = Request("txtBpCd")
	End if
	If Trim(Request("lgStrPrevKey1")) <> "" Then
		I1_i_vmi_onhand_stock(I509_I1_item_cd) = Request("lgStrPrevKey1")
	Else
		I1_i_vmi_onhand_stock(I509_I1_item_cd) = Request("txtItemCd")
	End if
	If Trim(Request("lgStrPrevKey2")) <> "" Then
		I1_i_vmi_onhand_stock(I509_I1_tracking_no) = Request("lgStrPrevKey2")
	Else
		I1_i_vmi_onhand_stock(I509_I1_tracking_no) = ""
	End if
	If Trim(Request("lgStrPrevKey3")) <> "" Then
		I1_i_vmi_onhand_stock(I509_I1_lot_no)  = Request("lgStrPrevKey3")
	Else
		I1_i_vmi_onhand_stock(I509_I1_lot_no)  = ""
	End if
	If Trim(Request("lgStrPrevKey4")) <> "" Then
		I1_i_vmi_onhand_stock(I509_I1_lot_sub_no) = Request("lgStrPrevKey4")
	Else
		I1_i_vmi_onhand_stock(I509_I1_lot_sub_no) = ""
	End if

    LngMaxRow = CLng(Request("txtMaxRows"))


    Set iPI5G160 = Server.CreateObject("PI5G160.clVMIListOnhandStk")
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
	If CheckSYSTEMError(Err, True) = True Then
		Response.End											
	End If    

    '-----------------------
    'Com Action Area
    '-----------------------
    Call iPI5G160.I_VMI_LIST_ONHAND_STK(gStrGlobalCollection, C_SHEETMAXROWS_D, _
                           I1_i_vmi_onhand_stock, _
                           EG1_i_vmi_onhand_stock, _
                           EG2_next_list)

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err, True) = True Then
    	Set iPI5G160 = Nothing												
		Response.End												
	End If

  	Set iPI5G160 = Nothing											

	if isEmpty(EG1_i_vmi_onhand_stock) then
		Response.End
	end if 
	
	ReDim PvArr(Ubound(EG1_i_vmi_onhand_stock,1))
	
	For LngRow = 0 To Ubound(EG1_i_vmi_onhand_stock,1)
		strData = Chr(11) & ConvSPChars(EG1_i_vmi_onhand_stock(LngRow, I509_EG1_bp_cd)) & _
				  Chr(11) & ConvSPChars(EG1_i_vmi_onhand_stock(LngRow, I509_EG1_bp_nm)) & _
				  Chr(11) & ConvSPChars(EG1_i_vmi_onhand_stock(LngRow, I509_EG1_item_cd)) & _
				  Chr(11) & ConvSPChars(EG1_i_vmi_onhand_stock(LngRow, I509_EG1_item_nm)) & _
				  Chr(11) & UniConvNumberDBToCompany(EG1_i_vmi_onhand_stock(LngRow, I509_EG1_good_on_hand_qty), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
				  Chr(11) & ConvSPChars(EG1_i_vmi_onhand_stock(LngRow, I509_EG1_basic_unit)) & _
				  Chr(11) & ConvSPChars(EG1_i_vmi_onhand_stock(LngRow, I509_EG1_tracking_no)) & _
				  Chr(11) & ConvSPChars(EG1_i_vmi_onhand_stock(LngRow, I509_EG1_lot_no)) & _
				  Chr(11) & ConvSPChars(EG1_i_vmi_onhand_stock(LngRow, I509_EG1_lot_sub_no)) & _
				  Chr(11) & ConvSPChars(EG1_i_vmi_onhand_stock(LngRow, I509_EG1_spec)) & _
				  Chr(11) & LngMaxRow + LngRow & Chr(11) & Chr(12)
		PvArr(LngRow) = strData
	Next
	
	strData = Join(PvArr, "")

	Response.Write "<Script Language=vbscript>" & vbcr
	Response.Write "With Parent" & vbcr
	Response.Write "    .ggoSpread.Source = .frm1.vspdData" & vbcr
	Response.Write "    .ggoSpread.SSShowData """ & strData & """" & vbcr

	Response.Write "    .lgStrPrevKey  = """ & ConvSPChars(EG2_next_list(I509_EG2_bp_cd))       & """" & vbcr
	Response.Write "    .lgStrPrevKey1 = """ & ConvSPChars(EG2_next_list(I509_EG2_item_cd))     & """" & vbcr
	Response.Write "    .lgStrPrevKey2 = """ & ConvSPChars(EG2_next_list(I509_EG2_tracking_no)) & """" & vbcr
	Response.Write "    .lgStrPrevKey3 = """ & ConvSPChars(EG2_next_list(I509_EG2_lot_no))      & """" & vbcr
	Response.Write "    .lgStrPrevKey4 = """ & ConvSPChars(EG2_next_list(I509_EG2_lot_sub_no))  & """" & vbcr
	
	Response.Write "	If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> """" Then "	& vbCr
	Response.Write "		.DbQuery								"				& vbCr
	Response.Write "    Else								"				& vbCr
	Response.Write "		.DbQueryOK								"				& vbCr
	Response.Write "    End If								"				& vbCr
	Response.Write "End With"       & vbcr
	Response.Write "</Script>"      & vbcr
%>

