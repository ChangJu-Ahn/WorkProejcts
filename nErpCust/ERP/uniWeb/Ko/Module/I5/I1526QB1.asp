<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : VMI 수불현황조회 
'*  3. Program ID           : I1526QB1
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
	
Dim iPI5G170		
Dim LngMaxRow		
Dim LngRow
Dim strData
Dim SetComboList
Dim ComboRow
Dim PvArr

Const C_SHEETMAXROWS_D = 100

Dim I1_i_vmi_goods_mvmt
	Const I510_I1_plant_cd = 0
	Const I510_I1_start_document_dt = 1
	Const I510_I1_end_document_dt = 2
	Const I510_I2_sl_cd = 3
	Const I510_I2_bp_cd = 4
	Const I510_I2_item_cd = 5
	Const I510_I1_trns_type = 6
ReDim I1_i_vmi_goods_mvmt(I510_I1_trns_type)
Dim I2_i_vmi_goods_mvmt_Next
	Const I510_I2_item_document_no = 0
	Const I510_I2_document_Year = 1
	Const I510_I2_seq_no = 2
	Const I510_I2_Sub_seq_no = 3
ReDim I2_i_vmi_goods_mvmt_Next(I510_I2_Sub_seq_no)

Dim E1_i_vmi_goods_mvmt_Next
 Const I510_E1_item_document_no = 0
 Const I510_E1_document_Year = 1
 Const I510_E1_seq_no = 2
 Const I510_E1_Sub_seq_no = 3
Dim EG1_i_vmi_goods_mvmt
 Const I510_EG1_item_cd = 0
 Const I510_EG1_item_nm = 1
 Const I510_EG1_document_dt = 2
 Const I510_EG1_sl_cd = 3
 Const I510_EG1_trns_sl_cd = 4
 Const I510_EG1_bp_cd = 5
 Const I510_EG1_bp_nm = 6
 Const I510_EG1_trns_type = 7
 Const I510_EG1_qty = 8
 Const I510_EG1_basic_unit = 9
 Const I510_EG1_tracking_no = 10
 Const I510_EG1_lot_no = 11
 Const I510_EG1_lot_sub_no = 12
 Const I510_EG1_spec = 13
 Const I510_EG1_item_document_no = 14
 Const I510_EG1_document_Year = 15
 Const I510_EG1_seq_no = 16
 Const I510_EG1_Sub_seq_no = 17
 Const I510_EG1_document_text = 18

	I1_i_vmi_goods_mvmt(I510_I1_plant_cd)          = Trim(Request("txtPlantCd"))
	I1_i_vmi_goods_mvmt(I510_I1_start_document_dt) = UNIConvDate(Request("txtTrnsFrDt"))
	I1_i_vmi_goods_mvmt(I510_I1_end_document_dt)   = UNIConvDate(Request("txtTrnsToDt"))
	I1_i_vmi_goods_mvmt(I510_I2_sl_cd)             = Trim(Request("txtSlCd"))
	I1_i_vmi_goods_mvmt(I510_I2_bp_cd)             = Trim(Request("txtBpCd"))
	I1_i_vmi_goods_mvmt(I510_I2_item_cd)           = Trim(Request("txtItemCd"))
	I1_i_vmi_goods_mvmt(I510_I1_trns_type)         = Trim(Request("cboTrnsType"))
		
	I2_i_vmi_goods_mvmt_Next(I510_I2_item_document_no) = Trim(Request("lgStrPrevKey"))
	I2_i_vmi_goods_mvmt_Next(I510_I2_document_Year)    = Trim(Request("lgStrPrevKey1"))
	I2_i_vmi_goods_mvmt_Next(I510_I2_seq_no)           = Trim(Request("lgStrPrevKey2"))
	I2_i_vmi_goods_mvmt_Next(I510_I2_Sub_seq_no)       = Trim(Request("lgStrPrevKey3"))
	    
    LngMaxRow      = CLng(Request("txtMaxRows"))
    SetComboList   = SetComboSplit(Request("SetComboList"))

    Set iPI5G170 = Server.CreateObject("PI5G170.cIVMIListGoodsMvmt")

	If CheckSYSTEMError(Err, True) = True Then
		Response.End												
	End If    

    '-----------------------
    'Com Action Area
    '-----------------------
    Call iPI5G170.I_VMI_LIST_GOODS_MVMT(gStrGlobalCollection, C_SHEETMAXROWS_D, _
                           I1_i_vmi_goods_mvmt, _
                           I2_i_vmi_goods_mvmt_Next, _
                           E1_i_vmi_goods_mvmt_Next, _
                           EG1_i_vmi_goods_mvmt) 	

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err, True) = True Then
    	Set iPI5G170 = Nothing												
		Response.End												
	End If

  	Set iPI5G170 = Nothing											

	if isEmpty(EG1_i_vmi_goods_mvmt) then
		Response.End
	end if 
	
	ReDim PvArr(Ubound(EG1_i_vmi_goods_mvmt,1))
	
	For LngRow = 0 To Ubound(EG1_i_vmi_goods_mvmt,1)
	    For ComboRow = 0 To Ubound(SetComboList, 2)
			If UCase(Trim(SetComboList(0, ComboRow))) = UCase(Trim(EG1_i_vmi_goods_mvmt(LngRow, I510_EG1_trns_type)))  Then
				EG1_i_vmi_goods_mvmt(LngRow, I510_EG1_trns_type) = Trim(SetComboList(1, ComboRow))
				Exit For
			End If
		Next
		
	    strData = Chr(11) & ConvSPChars(EG1_i_vmi_goods_mvmt(LngRow, I510_EG1_item_cd)) & _
				  Chr(11) & ConvSPChars(EG1_i_vmi_goods_mvmt(LngRow, I510_EG1_item_nm)) & _
				  Chr(11) & UniDateClientFormat(EG1_i_vmi_goods_mvmt(LngRow, I510_EG1_document_dt)) & _
				  Chr(11) & ConvSPChars(EG1_i_vmi_goods_mvmt(LngRow, I510_EG1_sl_cd)) & _
				  Chr(11) & ConvSPChars(EG1_i_vmi_goods_mvmt(LngRow, I510_EG1_trns_sl_cd)) & _
				  Chr(11) & ConvSPChars(EG1_i_vmi_goods_mvmt(LngRow, I510_EG1_bp_cd)) & _
				  Chr(11) & ConvSPChars(EG1_i_vmi_goods_mvmt(LngRow, I510_EG1_bp_nm)) & _
				  Chr(11) & ConvSPChars(EG1_i_vmi_goods_mvmt(LngRow, I510_EG1_trns_type)) & _
				  Chr(11) & UniConvNumberDBToCompany(EG1_i_vmi_goods_mvmt(LngRow, I510_EG1_qty), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
				  Chr(11) & ConvSPChars(EG1_i_vmi_goods_mvmt(LngRow, I510_EG1_basic_unit)) & _
				  Chr(11) & ConvSPChars(EG1_i_vmi_goods_mvmt(LngRow, I510_EG1_tracking_no)) & _
				  Chr(11) & ConvSPChars(EG1_i_vmi_goods_mvmt(LngRow, I510_EG1_lot_no)) & _
				  Chr(11) & ConvSPChars(EG1_i_vmi_goods_mvmt(LngRow, I510_EG1_lot_sub_no)) & _
				  Chr(11) & ConvSPChars(EG1_i_vmi_goods_mvmt(LngRow, I510_EG1_spec)) & _
				  Chr(11) & ConvSPChars(EG1_i_vmi_goods_mvmt(LngRow, I510_EG1_item_document_no)) & _
				  Chr(11) & ConvSPChars(EG1_i_vmi_goods_mvmt(LngRow, I510_EG1_document_Year)) & _
				  Chr(11) & ConvSPChars(EG1_i_vmi_goods_mvmt(LngRow, I510_EG1_seq_no)) & _
				  Chr(11) & ConvSPChars(EG1_i_vmi_goods_mvmt(LngRow, I510_EG1_Sub_seq_no)) & _
				  Chr(11) & LngMaxRow + LngRow & Chr(11) & Chr(12)
		PvArr(LngRow) = strData   
	Next
	strData = Join(PvArr, "")
	
	Response.Write "<Script Language=vbscript>" & vbcr
	Response.Write "With Parent" & vbcr
	Response.Write "    .ggoSpread.Source = .frm1.vspdData" & vbcr
	Response.Write "    .ggoSpread.SSShowData """ & strData & """" & vbcr

	Response.Write "    .lgStrPrevKey  = """ & ConvSPChars(E1_i_vmi_goods_mvmt_Next(I510_E1_item_document_no)) & """" & vbcr
	Response.Write "    .lgStrPrevKey1 = """ & ConvSPChars(E1_i_vmi_goods_mvmt_Next(I510_E1_document_Year))    & """" & vbcr
	Response.Write "    .lgStrPrevKey2 = """ & ConvSPChars(E1_i_vmi_goods_mvmt_Next(I510_E1_seq_no))           & """" & vbcr
	Response.Write "    .lgStrPrevKey3 = """ & ConvSPChars(E1_i_vmi_goods_mvmt_Next(I510_E1_Sub_seq_no))       & """" & vbcr
	Response.Write "	If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> """" Then "	& vbCr
	Response.Write "		.DbQuery								"				& vbCr
	Response.Write "    Else								"				& vbCr
	Response.Write "		.DbQueryOK								"				& vbCr
	Response.Write "    End If								"				& vbCr
	Response.Write "End With"       & vbcr
	Response.Write "</Script>"      & vbcr

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

