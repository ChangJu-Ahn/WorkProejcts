'**********************************************************************************************
'*  1. Module Name          : SCM
'*  2. Function Name        : u2116ma1
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         : 입고예정일등록 (Manage Planned Delivery Date)
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2004/07/27
'*  8. Modified date(Last)  : 2004/08/12
'*  9. Modifier (First)     : Park, BumSoo
'* 10. Modifier (Last)      : Park, BumSoo
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            
'**********************************************************************************************
Const BIZ_PGM_ID	= "u2116mb1.asp"			'☆: List & Manage SCM Orders
Const BIZ_PGM_ID2	= "u2116mb2.asp"			'☆: List & Manage SCM Orders

Dim C_OrderDt
Dim C_ItemCode
Dim C_ItemName
Dim C_Spec
Dim C_RetFlag
Dim C_PlantCd
Dim C_PlantNm
Dim C_OrderUnit
Dim C_OrderNo
Dim C_OrderSeq
Dim C_OrderQty
Dim C_DvryDt
Dim	C_RcptQty
Dim	C_UnRcptQty
Dim	C_InspQty
Dim	C_FirmDvryQty
Dim C_RemainQty
Dim C_DvryPlanDt
Dim C_DvryQty
'Dim C_SLYN
Dim C_SLCD
Dim C_SLPOP
Dim C_SLNM
Dim C_ClsFlg
Dim C_RcptFlg

Dim	C_LotNo
Dim	C_LotSubNo
Dim	C_LotFlg

Dim C_Title
Dim C_DlvyPlanDt
Dim C_DlvyQty
'Dim	C_SerialNo
Dim C_RcptDt
Dim	C_ReceiptQty
Dim	C_RcptRemainQty

Dim lgOldRow

'================================================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE
    lgIntGrpCount = 0
    lgStrPrevKey = ""
    lgStrPrevKey1 = ""
    lgLngCurRows = 0
    lgSortKey1 = 1
	lgSortKey2 = 1
	lgOldRow = 0
End Sub

'================================================================================================================================
Sub InitSpreadComboBox()

End Sub

'================================================================================================================================
Sub InitData()

	Dim intRow
    Dim intIndex
    
End Sub

'================================================================================================================================
Sub SetDefaultVal()
	frm1.txtDvFrDt.text = UniConvDateAToB(UNIDateAdd ("M", -1, LocSvrDate, parent.gServerDateFormat), parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtDvToDt.text = UniConvDateAToB(UNIDateAdd ("M", 3, LocSvrDate, parent.gServerDateFormat), parent.gServerDateFormat, parent.gDateFormat)
	Call SetBPCD()
End Sub

'================================================================================================================================
Sub SetBPCD()

	If 	CommonQueryRs2by2(" BP_NM ", " B_BIZ_PARTNER ", " BP_CD = " & FilterVar(parent.gUsrId, "", "S"), lgF0) = False Then
		Call ggoOper.SetReqAttr(frm1.txtPlantCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtItemCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtDvFrDt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtDvToDt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtPoFrDt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtPoToDt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtSLCD,"Q")
		Call ggoOper.SetReqAttr(frm1.txtTRACKINGNO,"Q")
		Call DisplayMsgBox("210033","X","X","X")
		Call SetToolBar("10000000000011")
		Exit Sub
	Else
	    Call SetToolBar("11000000000011")								'⊙: 버튼 툴바 제어 
	End If

	lgF0 = Split(lgF0, Chr(11))
	frm1.txtBpCd.value = parent.gUsrId
	frm1.txtBpNm.value = lgF0(1)

End Sub

'================================================================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)

	Call InitSpreadPosVariables(pvSpdNo)

	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1
			ggoSpread.Source = .vspdData1
			ggoSpread.Spreadinit "V20070124", , Parent.gAllowDragDropSpread
			.vspdData1.ReDraw = False
	
			.vspdData1.MaxCols = C_RcptFlg + 1
			.vspdData1.MaxRows = 0
			
			Call GetSpreadColumnPos("A")

			ggoSpread.SSSetDate 	C_OrderDt,		"수주일자", 10, 2, parent.gDateFormat		 
			ggoSpread.SSSetEdit		C_ItemCode,		"품목"    , 18,,,18,2
			ggoSpread.SSSetEdit		C_ItemName,		"품목명"  , 18
			ggoSpread.SSSetEdit		C_Spec,			"규격"    , 15
			ggoSpread.SSSetEdit		C_RetFlag,		"구분"	  , 8 ,2
			ggoSpread.SSSetEdit		C_PlantCd,		"납품공장"  , 10
			ggoSpread.SSSetEdit		C_PlantNm,		"납품공장명",	12
			ggoSpread.SSSetEdit		C_OrderUnit,	"단위"    ,  7,,,3,2
			ggoSpread.SSSetEdit		C_OrderNo,		"수주번호", 15
			ggoSpread.SSSetEdit		C_OrderSeq,		"행번"        , 7
			ggoSpread.SSSetFloat	C_OrderQty,		"수주량"      ,10,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetDate 	C_DvryDt,		"납기일"      ,10, 2, parent.gDateFormat
			ggoSpread.SSSetFloat	C_RcptQty,		"납품량"      ,10,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_UnRcptQty,	"미납품량"    ,10,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_InspQty,		"검사중수량"    ,10,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_FirmDvryQty,	"납품대기량"  ,10,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_RemainQty,	"납품잔량"    ,10,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetDate 	C_DvryPlanDt,	"납품예정일자",12, 2, parent.gDateFormat
			ggoSpread.SSSetFloat	C_DvryQty,		"납품예정수량",12,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
            ggoSpread.SSSetEdit		C_SLCD,		    "납품창고"      ,10,,,10
			ggoSpread.SSSetButton   C_SLPOP
			ggoSpread.SSSetEdit		C_SLNM,		    "납품창고명"  ,14
			ggoSpread.SSSetEdit		C_ClsFlg,		"발주마감"    ,10,2
			ggoSpread.SSSetEdit		C_RcptFlg,		"입출고구분"    ,10,2

			'Call ggoSpread.SSSetColHidden( C_ClsFlg, C_ClsFlg , True)			
			Call ggoSpread.SSSetColHidden( .vspdData1.MaxCols, .vspdData1.MaxCols , True)
			
			.vspdData1.ReDraw = True
   
			ggoSpread.SSSetSplit2(3)
    
			Call SetSpreadLock("A")
			
			.vspdData1.ReDraw = true    
    
		End With
	
    End If

    If pvSpdNo = "B" Or pvSpdNo = "*" Then
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1
			ggoSpread.Source = .vspdData2
			ggoSpread.Spreadinit "V20050420", , Parent.gAllowDragDropSpread
			.vspdData2.ReDraw = False
	
			.vspdData2.MaxCols = C_RcptRemainQty + 1
			.vspdData2.MaxRows = 0
			
			Call GetSpreadColumnPos("B")
			ggoSpread.SSSetEdit		C_Title,		"Title"       , 10,2,,18,2
			ggoSpread.SSSetDate 	C_DlvyPlanDt,	"납품예정일자", 12, 2, parent.gDateFormat
			ggoSpread.SSSetFloat	C_DlvyQty,		"납품예정수량", 12,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetDate 	C_RcptDt,		"납품일"      , 10, 2, parent.gDateFormat
			ggoSpread.SSSetFloat	C_ReceiptQty,	"납품수량"    , 10,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_RcptRemainQty,"납품대기량"  , 12,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
    
			Call ggoSpread.SSSetColHidden( .vspdData2.MaxCols, .vspdData2.MaxCols, True)
			
			ggoSpread.SSSetSplit2(2)
			
			Call SetSpreadLock("B")
			
			.vspdData2.ReDraw = true
    
		End With
    End If
    
End Sub

'================================================================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)

	If pvSpdNo = "A" Then
		'--------------------------------
		'Grid 1
		'--------------------------------
		With frm1
			ggoSpread.Source = .vspdData1
	
			.vspdData1.ReDraw = False
   			ggoSpread.SpreadLock	 C_OrderDt, -1, C_OrderDt
   			ggoSpread.SpreadLock	 C_ItemCode, -1, C_ItemCode
			ggoSpread.SpreadLock	 C_ItemName, -1, C_ItemName
			ggoSpread.SpreadLock	 C_Spec, -1, C_Spec
			ggoSpread.SpreadLock	 C_retflag, -1, C_retflag
			ggoSpread.SpreadLock	 C_PlantCd, -1, C_PlantCd
			ggoSpread.SpreadLock	 C_PlantNm, -1, C_PlantNm
			ggoSpread.SpreadLock	 C_OrderUnit, -1, C_OrderUnit
			ggoSpread.SpreadLock	 C_OrderNo, -1, C_OrderNo
			ggoSpread.SpreadLock	 C_OrderSeq, -1, C_OrderSeq
			ggoSpread.SpreadLock	 C_OrderQty, -1, C_OrderQty
			ggoSpread.SpreadLock	 C_DvryDt, -1, C_DvryDt
			ggoSpread.SpreadLock	 C_RcptQty, -1, C_RcptQty
			ggoSpread.SpreadLock	 C_UnRcptQty, -1, C_UnRcptQty
			ggoSpread.SpreadLock	 C_InspQty, -1, C_InspQty
			ggoSpread.SpreadLock	 C_FirmDvryQty, -1, C_FirmDvryQty
			ggoSpread.SpreadLock	 C_RemainQty, -1, C_RemainQty
			ggoSpread.SpreadLock	 C_ClsFlg, -1, C_ClsFlg
			ggoSpread.SpreadLock	 C_RcptFlg, -1, C_RcptFlg
			ggoSpread.SSSetRequired  C_DvryPlanDt, -1
			ggoSpread.SSSetRequired  C_DvryQty, -1
			ggoSpread.SSSetRequired  C_SLCD, -1
			
			ggoSpread.SpreadLock	 C_SLNM, -1, C_SLNM
			.vspdData1.ReDraw = True
	
		End With
	End If
		
	If pvSpdNo = "B" Then 
		'--------------------------------
		'Grid 2
		'--------------------------------
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End If	

End Sub

'================================================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    
    With frm1.vspdData1 
    
    .Redraw = False

    ggoSpread.Source = frm1.vspdData1
    ggoSpread.SSSetProtected C_OrderDt,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ItemCode,	pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ItemName,	pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_Spec,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_PlantCd,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_PlantNm,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_OrderUnit,	pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_OrderNo,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_OrderSeq,	pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_OrderQty,	pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_DvryDt,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_RcptQty,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_UnRcptQty,	pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_InspQty,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_FirmDvryQty,	pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_RemainQty,	pvStartRow, pvEndRow
	
	.Col = C_ClsFlg
	If .Text = "Y" Then
		ggoSpread.SSSetProtected  C_DvryPlanDt,	pvStartRow, pvEndRow
		ggoSpread.SSSetProtected  C_DvryQty,		pvStartRow, pvEndRow
		
		ggoSpread.SSSetProtected  C_SLCD,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_SLNM,		pvStartRow, pvEndRow
	Else
		ggoSpread.SSSetRequired  C_DvryPlanDt,	pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_DvryQty,		pvStartRow, pvEndRow
		
		ggoSpread.SSSetRequired  C_SLCD,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_SLNM,		pvStartRow, pvEndRow
	End If

	ggoSpread.SSSetProtected C_SLNM,		pvStartRow, pvEndRow
	
    .Col = 1
    .Row = .ActiveRow
    .Action = 0                         'parent.SS_ACTION_ACTIVE_CELL
    .EditMode = True
    
    .Redraw = True
    
    End With
End Sub

'================================================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)	
	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		' Grid 1(vspdData1)
		C_OrderDt		= 1
		C_ItemCode		= 2
		C_ItemName		= 3
		C_Spec			= 4
		C_retflag		= 5
		C_PlantCd		= 6
		C_PlantNm		= 7
		C_OrderUnit		= 8
		C_OrderNo		= 9
		C_OrderSeq		= 10
		C_OrderQty		= 11
		C_DvryDt		= 12
		C_RcptQty		= 13
		C_UnRcptQty		= 14
		C_InspQty		= 15
		C_FirmDvryQty	= 16
		C_RemainQty		= 17
		C_DvryPlanDt	= 18
		C_DvryQty		= 19
		C_SLCD			= 20
		C_SLPOP			= 21
		C_SLNM			= 22
		C_ClsFlg		= 23
		C_RcptFlg		= 24

	End If	
	
	If pvSpdNo = "B" Or pvSpdNo = "*" Then
		' Grid 2(vspdData2)
		C_Title			= 1
		C_DlvyPlanDt	= 2
		C_DlvyQty		= 3
		C_RcptDt		= 4
		C_ReceiptQty	= 5
		C_RcptRemainQty	= 6
	End If

End Sub
 
'================================================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case Ucase(pvSpdNo)
 		Case "A"
 			ggoSpread.Source = frm1.vspdData1 
 			
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_OrderDt		= iCurColumnPos(1)
			C_ItemCode		= iCurColumnPos(2)
			C_ItemName		= iCurColumnPos(3)
			C_Spec			= iCurColumnPos(4)
			C_retflag		= iCurColumnPos(5)
			C_PlantCd		= iCurColumnPos(6)
			C_PlantNm		= iCurColumnPos(7)
			C_OrderUnit		= iCurColumnPos(8)
			C_OrderNo		= iCurColumnPos(9)
			C_OrderSeq		= iCurColumnPos(10)
			C_OrderQty		= iCurColumnPos(11)
			C_DvryDt		= iCurColumnPos(12)
			C_RcptQty		= iCurColumnPos(13)
			C_UnRcptQty		= iCurColumnPos(14)
			C_InspQty		= iCurColumnPos(15)
			C_FirmDvryQty	= iCurColumnPos(16)
			C_RemainQty		= iCurColumnPos(17)
			C_DvryPlanDt	= iCurColumnPos(18)
			C_DvryQty		= iCurColumnPos(19)
			C_SLCD			= iCurColumnPos(20)
			C_SLPOP			= iCurColumnPos(21)
			C_SLNM			= iCurColumnPos(22)
			C_ClsFlg		= iCurColumnPos(23)
			C_RcptFlg		= iCurColumnPos(24)
			
		Case "B"
			
			ggoSpread.Source = frm1.vspdData2
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_Title			= iCurColumnPos(1)
			C_DlvyPlanDt	= iCurColumnPos(2)
			C_DlvyQty		= iCurColumnPos(3)
			C_RcptDt		= iCurColumnPos(4)
			C_ReceiptQty	= iCurColumnPos(5)
			C_RcptRemainQty	= iCurColumnPos(6) 
			
 	End Select
 
End Sub

'================================================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "납품공장"
	arrParam(1) = "(			SELECT	DISTINCT B.PLANT_CD FROM M_SCM_PLAN_PUR_RCPT A, M_PUR_ORD_DTL B, M_PUR_ORD_HDR C "
	arrParam(1) = arrParam(1) & "WHERE	A.PO_NO = B.PO_NO AND A.PO_SEQ_NO = B.PO_SEQ_NO AND A.SPLIT_SEQ_NO = 0 "
	arrParam(1) = arrParam(1) & "AND	A.PO_NO = C.PO_NO AND C.BP_CD = '" & frm1.txtBpCd.value & "') A, B_PLANT B"
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = "A.PLANT_CD = B.PLANT_CD"			
	arrParam(5) = "납품공장"			
	
    arrField(0) = "A.PLANT_CD"	
    arrField(1) = "B.PLANT_NM"	
    
    arrHeader(0) = "납품공장"		
    arrHeader(1) = "납품공장명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'================================================================================================================================
Function OpenItemInfo(Byval strCode)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtItemCd.className) = "PROTECTED" Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "품목팝업"
	arrParam(1) = "(			SELECT	DISTINCT ITEM_CD FROM M_SCM_PLAN_PUR_RCPT A, M_PUR_ORD_HDR B "
	arrParam(1) = arrParam(1) & "WHERE	A.PO_NO = B.PO_NO AND A.SPLIT_SEQ_NO = 0 AND B.BP_CD = '" & frm1.txtBpCd.value & "') A, B_ITEM B"
	arrParam(2) = Trim(frm1.txtItemCd.Value)								' Code Condition
	arrParam(3) = ""														' Name Cindition
	arrParam(4) = "A.ITEM_CD = B.ITEM_CD "
	arrParam(5) = "품목"
	 
    arrField(0) = "A.ITEM_CD"												' Field명(0)
    arrField(1) = "B.ITEM_NM"												' Field명(1)
    
    arrHeader(0) = "품목"													' Header명(0)
    arrHeader(1) = "품목명"													' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus

End Function

'================================================================================================================================
Function OpenSlInfo(Byval strCode)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtItemCd.className) = "PROTECTED" Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "창고팝업"
	arrParam(1) = "B_STORAGE_LOCATION "
	arrParam(2) = Trim(frm1.txtSlCd.Value)								' Code Condition
	arrParam(3) = ""														' Name Cindition
	arrParam(4) = ""
	arrParam(5) = "창고"
	 
    arrField(0) = "SL_CD"												' Field명(0)
    arrField(1) = "SL_NM"												' Field명(1)
    
    arrHeader(0) = "창고"													' Header명(0)
    arrHeader(1) = "창고명"													' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetSlInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtSlCd.focus

End Function

'================================================================================================================================
Function OpenBpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtBpCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"										' 팝업 명칭 
	arrParam(1) = "B_Biz_Partner"								' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtBpCd.Value)						' Code Condition
	arrParam(3) = ""
	arrParam(4) = "BP_TYPE In ('S','CS') And usage_flag='Y'"	' Where Condition
	arrParam(5) = "공급처"										' TextBox 명칭 
	
    arrField(0) = "BP_CD"										' Field명(0)
    arrField(1) = "BP_NM"										' Field명(1)
    
    arrHeader(0) = "공급처"										' Header명(0)
    arrHeader(1) = "공급처명"									' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Exit Function
	Else
		frm1.txtBpCd.Value = arrRet(0)
		frm1.txtBpNm.Value = arrRet(1)
		frm1.txtBpCd.focus
	End If	
End Function

Function OpenSLCD(byval strCon)  
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "납품창고"     
	arrParam(1) = "B_STORAGE_LOCATION"   
	 
	frm1.vspdData1.Row = frm1.vspdData1.ActiveRow 
	frm1.vspdData1.Col = C_SLCD
	arrParam(2) = Trim(frm1.vspdData1.text) 
	frm1.vspdData1.Col = C_PlantCd
	arrParam(4) = " PLANT_CD = '" & Trim(frm1.vspdData1.text) & "' "      
	arrParam(5) = "납품창고"    
	 
	arrField(0) = "SL_CD"     
	arrField(1) = "SL_NM"     
	    
	arrHeader(0) = "납품창고"   
	arrHeader(1) = "납품창고명"   
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.vspdData1.Row = frm1.vspdData1.ActiveRow 
		frm1.vspdData1.Col = C_SLCD
		frm1.vspdData1.text = arrRet(0) 
		frm1.vspdData1.Col = C_SLNM
		frm1.vspdData1.text = arrRet(1) 
		ggoSpread.UpdateRow Row
	End If 
End Function


'================================================================================================================================
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtPlantCd.focus()		
End Function

'================================================================================================================================
Function SetItemInfo(Byval arrRet)
    With frm1
		.txtItemCd.value = arrRet(0)
		.txtItemNm.value = arrRet(1)
    End With
End Function

'================================================================================================================================
Function SetSlInfo(Byval arrRet)
    With frm1
		.txtSlCd.value = arrRet(0)
		.txtSlNm.value = arrRet(1)
    End With
End Function

'--------------------------------------  OpenTrackingInfo()  ------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTracking Info PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackingInfo()
	
	Dim arrRet
	Dim arrParam(4)
	Dim iCalledAspName
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	If IsOpenPop = True Or UCase(frm1.txtTrackingNo.className) = "PROTECTED" Then Exit Function
	
	iCalledAspName = AskPRAspName("P4600PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4600PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtTrackingNo.value)
	arrParam(2) = Trim(frm1.txtItemCd.value)
	arrParam(3) = frm1.txtPoFrDt.Text
	arrParam(4) = frm1.txtPoToDt.Text
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetTrackingNo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtTrackingNo.focus
	
End Function

'------------------------------------------  SetTrackingNo()  -----------------------------------------
'	Name : SetTrackingNo()
'	Description : Tracking No. Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetTrackingNo(Byval arrRet)
    frm1.txtTrackingNo.Value = arrRet(0)
End Function


'================================================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub
 
'================================================================================================================================
Sub vspdData1_Change(ByVal Col , ByVal Row)
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.UpdateRow Row
End Sub

'================================================================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row )

	If lgIntFlgMode = Parent.OPMD_UMODE Then
  		Call SetPopupMenuItemInf("1101111111")
  	Else
  		Call SetPopupMenuItemInf("1001111111")
  	End If

	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData1

	If frm1.vspdData1.MaxRows = 0 Then
		Exit Sub
	End If
   	
   	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData1
	        If lgSortKey1 = 1 Then
        	    ggoSpread.SSSort Col
	            lgSortKey1 = 2
	        Else
        	    ggoSpread.SSSort Col, lgSortKey1
	            lgSortKey1 = 1
        	End If
	End If

'''' JSA 2006-9-11 삭제 
''''	If lgOldRow <> Row Then
''''				
''''		frm1.vspdData2.MaxRows = 0 
''''		lgStrPrevKey1 = ""
''''		If DbDtlQuery = False Then	
''''			Call RestoreToolBar()
''''			Exit Sub
''''		End If
''''		
''''		lgOldRow = frm1.vspdData1.ActiveRow
''''			
''''	End If

End Sub

'================================================================================================================================
Sub vspdData2_Click(ByVal Col , ByVal Row )

	If lgIntFlgMode = Parent.OPMD_UMODE Then
  		Call SetPopupMenuItemInf("1101111111")
  	Else
  		Call SetPopupMenuItemInf("1001111111")
  	End If

	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData2

	If frm1.vspdData2.MaxRows = 0 Then
		Exit Sub
	End If
   	
   	If Row <= 0 Then
   		If Col = C_DlvyPlanDt or Col = C_SerialNo or Col = C_RcptDt Then
			ggoSpread.Source = frm1.vspdData2
			If lgSortKey2 = 1 Then
			    ggoSpread.SSSort Col
			    lgSortKey2 = 2
			Else
			    ggoSpread.SSSort Col, lgSortKey2
			    lgSortKey2 = 1
			End If
		End If
	End If

End Sub

'================================================================================================================================
Sub vspdData1_MouseDown(Button,Shift,x,y)
		
	If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'================================================================================================================================
Sub vspdData2_MouseDown(Button,Shift,x,y)
		
	If Button = 2 And gMouseClickStatus = "SP2C" Then
       gMouseClickStatus = "SP2CR"
    End If

End Sub


'================================================================================================================================
Sub vspdData1_DblClick(ByVal Col, ByVal Row)
   
 	If Row <= 0 Then
		Exit Sub
 	End If
    
  	If frm1.vspdData1.MaxRows = 0 Then
		Exit Sub
 	End If

End Sub
 
'================================================================================================================================
Sub vspdData1_KeyPress(index , KeyAscii )
    On Error Resume Next                                                    '☜: Protect system from crashing
End Sub


'================================================================================================================================
'   Event Name : vspdData1_ScriptLeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'================================================================================================================================
Sub vspdData1_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If NewRow * Row <= 0 Or Row = NewRow Then
		Exit Sub
	End If
	
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData
	
	lgStrPrevKey1 = ""

	Call SetActiveCell(frm1.vspdData1,NewCol,NewRow,"M","X","X")
	If DbDtlQuery = False Then	
		Exit Sub
	End If
	
End Sub

'================================================================================================================================
Sub vspdData1_ComboSelChange(ByVal Col, ByVal Row)
    On Error Resume Next                                                    '☜: Protect system from crashing
End Sub

'================================================================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	
	If CheckRunningBizProcess = True Then Exit Sub
	
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
		If lgStrPrevKey <> "" Then
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If	
		End If
    End if
    
End Sub

'================================================================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then
        Exit Sub
	End If
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then
		If lgStrPrevKey1 <> "" Then
			If DbDtlQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub

'================================================================================================================================
Sub vspdData1_ButtonClicked(Col, Row, ButtonDown)
	With frm1.vspdData1
		 ggoSpread.Source = frm1.vspdData1
		 .Row = Row
         .Col = Col
		If Row > 0 Then
			Select Case Col
				
				Case C_SLPOP
					.Col = Col - 1
			    	.Row = Row
					
					Call OpenSLCD(.text)
				
					
			End Select
		End If
    
	End With	
End Sub


'================================================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
 
'================================================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub
 
'================================================================================================================================
Sub vspdData1_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub 

'================================================================================================================================
Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
End Sub
 
'================================================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub 
 
'================================================================================================================================
Sub PopRestoreSpreadColumnInf()
	Dim LngRow

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    
    Call InitSpreadSheet("*")
    
    Call InitSpreadComboBox
    
	Call ggoSpread.ReOrderingSpreadData()

End Sub 

'================================================================================================================================
Sub txtDvFrDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDvFrDt.Action = 7
        SetFocusToDocument("M")
		Frm1.txtDvFrDt.Focus
    End If
End Sub

'================================================================================================================================
Sub txtDvToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDvToDt.Action = 7
        SetFocusToDocument("M")
		Frm1.txtDvToDt.Focus
    End If
End Sub

'================================================================================================================================
Sub txtDvFrDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'================================================================================================================================
Sub txtDvToDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'================================================================================================================================
Sub txtPoFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPoFrDt.Action = 7
		Call SetFocusToDocument("M") 
		frm1.txtPoFrDt.focus
	End If
End Sub

'================================================================================================================================
Sub txtPoToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPoToDt.Action = 7
		Call SetFocusToDocument("M") 
		frm1.txtPoToDt.focus
	End If
End Sub

'================================================================================================================================
Sub txtPoFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call FncQuery()
End Sub

'================================================================================================================================
Sub txtPoToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call FncQuery()
End Sub

'================================================================================================================================
Function FncQuery()
 
    Dim IntRetCD 
    
    FncQuery = False
    
    Err.Clear

    ggoSpread.Source = frm1.vspdData1
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")
	If IntRetCD = vbNo Then
	    Exit Function
	End If
    End If

    If ValidDateCheck(frm1.txtDvFrDt, frm1.txtDvToDt) = False Then Exit Function
	If ValidDateCheck(frm1.txtPoFrDt, frm1.txtPoToDt) = False Then Exit Function
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData

    Call InitVariables

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function
       
    FncQuery = True
   
End Function

'================================================================================================================================
Function FncNew() 
	On Error Resume Next	
End Function

'================================================================================================================================
Function FncDelete() 
	On Error Resume Next   
End Function

'================================================================================================================================
Function FncSave()
    Dim IntRetCD 
         
    FncSave = False 
    Err.Clear
    
    ggoSpread.Source = frm1.vspdData1
    If  ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData1
    If Not ggoSpread.SSDefaultCheck Then
       Exit Function
    End If
    
    If Not chkField(Document, "2") Then
       Exit Function
    End If

	Call DisableToolBar( parent.TBC_SAVE)
	If DbSave = False Then
		Call  RestoreToolBar()
		Exit Function
	End If
    
    FncSave = True
    
End Function

'================================================================================================================================
Function FncCopy() 
        
    If frm1.vspdData1.MaxRows < 1 Then Exit Function	
        
    frm1.vspdData1.focus
    Set gActiveElement = document.activeElement 
    frm1.vspdData1.EditMode = True
	    
    frm1.vspdData1.ReDraw = False
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.CopyRow
    frm1.vspdData1.ReDraw = True
    SetSpreadColor frm1.vspdData1.ActiveRow, frm1.vspdData1.ActiveRow

End Function

'================================================================================================================================
Function FncPaste() 
     ggoSpread.SpreadPaste
End Function

'================================================================================================================================
Function FncCancel() 
    If frm1.vspdData1.MaxRows < 1 Then Exit Function	
    ggoSpread.EditUndo
    Call initData(frm1.vspdData1.ActiveRow)
End Function

'================================================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
Dim IntRetCD
Dim imRow
Dim pvRow
	
On Error Resume Next
	
	FncInsertRow = False
    
    If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)

	Else
		imRow = AskSpdSheetAddRowCount()

		If imRow = "" Then
			Exit Function
		End If
	End If
    
    With frm1
	.vspdData1.focus
	Set gActiveElement = document.activeElement 
	ggoSpread.Source = .vspdData1
	.vspdData1.ReDraw = False
	ggoSpread.InsertRow .vspdData1.ActiveRow, imRow
    SetSpreadColor .vspdData1.ActiveRow, .vspdData1.ActiveRow + imRow -1
	.vspdData1.ReDraw = True
    End With
    
    lgLngCurRows = imRow + lgLngCurRows

	Set gActiveElement = document.ActiveElement
	If Err.number = 0 Then FncInsertRow = True
End Function


'================================================================================================================================
Function FncDeleteRow() 

    Dim lDelRows
    
    If frm1.vspdData1.MaxRows < 1 Then Exit Function

    lDelRows = ggoSpread.DeleteRow
    lgLngCurRows = lDelRows + lgLngCurRows

End Function

'================================================================================================================================
Function FncPrint() 
	Call parent.FncPrint()
End Function

'================================================================================================================================
Function FncPrev() 
    On Error Resume Next
End Function

'================================================================================================================================
Function FncNext() 
    On Error Resume Next
End Function

'================================================================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)												
End Function

'================================================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                         
End Function

'================================================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'================================================================================================================================
Function FncExit()

    Dim IntRetCD
    FncExit = False
    
    ggoSpread.Source = frm1.vspdData1							'⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'⊙: Check If data is chaged
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")	'⊙: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
    
End Function

'================================================================================================================================
Function FncScreenSave() 
    Call ggoSpread.SaveLayout
End Function

'================================================================================================================================
Function FncScreenRestore() 
    If ggoSpread.AllClear = True Then
       ggoSpread.LoadLayout
    End If
End Function

'================================================================================================================================
Sub MakeKeyStream(pOpt)

	Dim strPoNo
	Dim strPoSeqNo

   Select Case pOpt
       Case "M"
			With frm1
				If lgIntFlgMode = parent.OPMD_UMODE Then
					lgKeyStream = UCase(Trim(.hPlantCd.value))  & Parent.gColSep
					lgKeyStream = lgKeyStream & UCase(Trim(.hItemCd.value))  & Parent.gColSep
					lgKeyStream = lgKeyStream & UCase(Trim(.hBPCd.value))  & Parent.gColSep
					lgKeyStream = lgKeyStream & Trim(.hDvFrDt.value)  & Parent.gColSep
					lgKeyStream = lgKeyStream & Trim(.hDvToDt.value)  & Parent.gColSep
					lgKeyStream = lgKeyStream & Trim(.hPoFrDt.value)  & Parent.gColSep
					lgKeyStream = lgKeyStream & Trim(.hPoToDt.value)  & Parent.gColSep
					lgKeyStream = lgKeyStream & Trim(.hSlCD.Value)  & Parent.gColSep
					lgKeyStream = lgKeyStream & Trim(.hTRACKINGNO.Value)  & Parent.gColSep
				Else
					lgKeyStream = UCase(Trim(.txtPlantCd.value))  & Parent.gColSep
					lgKeyStream = lgKeyStream & UCase(Trim(.txtItemCd.value))  & Parent.gColSep
					lgKeyStream = lgKeyStream & UCase(Trim(.txtBPCd.value))  & Parent.gColSep
					lgKeyStream = lgKeyStream & Trim(.txtDvFrDt.Text)  & Parent.gColSep
					lgKeyStream = lgKeyStream & Trim(.txtDvToDt.Text)  & Parent.gColSep
					lgKeyStream = lgKeyStream & Trim(.txtPoFrDt.Text)  & Parent.gColSep
					lgKeyStream = lgKeyStream & Trim(.txtPoToDt.Text)  & Parent.gColSep
					lgKeyStream = lgKeyStream & Trim(.txtSlCD.Value)  & Parent.gColSep
					lgKeyStream = lgKeyStream & Trim(.txtTRACKINGNO.Value)  & Parent.gColSep
					
					.hPlantCd.value		= .txtPlantCd.value
					.hItemCd.value		= .txtItemCd.value
					.hBPCd.value		= .txtBPCd.value
					.hDvFrDt.value		= .txtDvFrDt.Text
					.hDvToDt.value		= .txtDvToDt.Text
					.hPoFrDt.value		= .txtPoFrDt.Text
					.hPoToDt.value		= .txtPoToDt.Text
					.hSlCd.value		= .txtSlCd.value
					.hTrackingNo.value	= .txtTrackingNo.value
					
				End If
			End With

       Case "S"
			With frm1
				.vspdData1.Row = .vspdData1.ActiveRow
				.vspdData1.Col = C_OrderNo
				strPoNo = .vspdData1.text
				.vspdData1.Col = C_OrderSeq
				strPoSeqNo = .vspdData1.text
					
				lgKeyStream = lgKeyStream & UCase(Trim(strPoNo))  & Parent.gColSep
				lgKeyStream = lgKeyStream & UCase(Trim(strPoSeqNo))  & Parent.gColSep

			End With

	End Select

	   
End Sub

'================================================================================================================================
Function DbQuery() 

    Dim strVal
    
    Err.Clear

    DbQuery = False
    
    Call LayerShowHide(1)
 
    Call MakeKeyStream("M")
    
	lgLngCurRows = frm1.vspdData1.MaxRows

	strVal = BIZ_PGM_ID & "?txtMode="	& parent.UID_M0001
	strVal = strVal & "&txtKeyStream="  & lgKeyStream
	strVal = strVal & "&lgStrPrevKey="  & lgStrPrevKey
 
    Call RunMyBizASP(MyBizASP, strVal)

    DbQuery = True
    
End Function

'================================================================================================================================
Function DbQueryOk(ByVal LngMaxRow)

 	Dim lRow
 	Dim LngRow    
	
    Call ggoOper.LockField(Document, "Q")
    Call SetToolBar("11001001000111")
    
    
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		If DbDtlQuery = False Then	
			Call RestoreToolBar()
			Exit Function
		End If
	End If
	
	'=======================================
	' 마감된 경우 
	'=======================================
	frm1.vspdData1.Redraw = False

	For LngRow = lgLngCurRows + 1 To frm1.vspdData1.MaxRows
		frm1.vspdData1.Col = C_ClsFlg 
		frm1.vspdData1.Row = LngRow
		If frm1.vspdData1.Text = "Y" Then 
			Call SetSpreadColor(LngRow, LngRow)
		End If 
	Next

	frm1.vspdData1.Redraw = True
	'=======================================

	lgLngCurRows = frm1.vspdData1.MaxRows

	Call initdata()

	Frm1.vspdData1.Focus
	lgIntFlgMode = parent.OPMD_UMODE
	
End Function

'================================================================================================================================
Function DbQueryNotOk()	

'	Call SetToolBar("11000000000011")
'    '-----------------------
'    'Reset variables area
'    '-----------------------
'    lgIntFlgMode = parent.OPMD_CMODE

End Function

'================================================================================================================================
Function DbDtlQuery() 
    Dim strVal
	
    DbDtlQuery = False

	Call LayerShowHide(1)

	lgKeyStream = ""
    Call MakeKeyStream("S")
    
	strVal = BIZ_PGM_ID2 & "?txtMode="	& parent.UID_M0001
	strVal = strVal & "&txtKeyStream="  & lgKeyStream
	strVal = strVal & "&lgStrPrevKey="  & lgStrPrevKey1

    Call RunMyBizASP(MyBizASP, strVal)													'☜: 비지니스 ASP 를 가동 
    
    DbDtlQuery = True
    
End Function

'================================================================================================================================
Function DbDtlQueryOk()
	Call SetQuerySpreadColor
End Function

Sub SetQuerySpreadColor()

	Dim iArrColor1, iArrColor2
	Dim iLoopCnt
	
	iArrColor1 = Split(lgStrColorFlag,Parent.gRowSep)

	For iLoopCnt=0 to ubound(iArrColor1,1) - 1
		iArrColor2 = Split(iArrColor1(iLoopCnt),Parent.gColSep)
		
		With frm1.vspdData2	
			.Col = -1
			.Row =  iArrColor2(0)
		
			Select Case iArrColor2(1)
				Case "1"
					.BackColor = RGB(176,234,244) '하늘색 
					.ForeColor = vbBlue
			End Select
		End With
	Next

End Sub

'================================================================================================================================
Function DbSave() 

    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal, strDel
	
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
	exit Function
	end if

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
    
       For lRow = 1 To .vspdData1.MaxRows
    
           .vspdData1.Row = lRow
           .vspdData1.Col = 0
        
           Select Case .vspdData1.Text

               Case  ggoSpread.UpdateFlag                                      '☜: Update
               
					.vspdData1.Col = C_DvryQty
					If .vspdData1.Value > 0 Then
														  strVal = strVal & "C"  &  parent.gColSep					
														  strVal = strVal & lRow &  parent.gColSep
						.vspdData1.Col = C_OrderNo	    : strVal = strVal & Trim(.vspdData1.Text) &  parent.gColSep	'2
						.vspdData1.Col = C_OrderSeq     : strVal = strVal & Trim(.vspdData1.Value)&  parent.gColSep	'3
						.vspdData1.Col = C_DvryPlanDt	: strVal = strVal & Trim(.vspdData1.Text) &  parent.gColSep	'4
						.vspdData1.Col = C_DvryQty		: strVal = strVal & Trim(.vspdData1.Value) & parent.gColSep	'5
						.vspdData1.Col = C_SLCD			: strVal = strVal & Trim(.vspdData1.Text) &  parent.gRowSep	'6
                   
						lGrpCnt = lGrpCnt + 1
					End If                                        
           End Select
       Next
	
       .txtMode.value        =  parent.UID_M0002
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal

	End With
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
    DbSave = True
    
End Function

'================================================================================================================================
Function DbSaveOk()

	Call InitVariables
	ggoSpread.source = frm1.vspdData1
    frm1.vspdData1.MaxRows = 0
    Call RemovedivTextArea
    Call MainQuery()

End Function

'================================================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr, parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData1.MaxCols - 1
           Frm1.vspdData1.Col = iDx
           Frm1.vspdData1.Row = iRow
           If Frm1.vspdData1.ColHidden <> True And Frm1.vspdData1.BackColor <>  parent.UC_PROTECTED Then
              Frm1.vspdData1.Col = iDx
              Frm1.vspdData1.Row = iRow
              Frm1.vspdData1.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub

'================================================================================================================================
Function RemovedivTextArea()

	Dim ii
		
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next

End Function

'================================================================================================================================
Function SheetFocus(lRow, lCol)
	frm1.vspdData1.focus
	frm1.vspdData1.Row = lRow
	frm1.vspdData1.Col = lCol
	frm1.vspdData1.Action = 0
	frm1.vspdData1.SelStart = 0
	frm1.vspdData1.SelLength = len(frm1.vspdData1.Text)
End Function
 