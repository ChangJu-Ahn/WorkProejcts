'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_QRY1_ID	= "i2221mb1.asp"								
Const BIZ_PGM_QRY2_ID	= "i2221mb2.asp"								

'==========================================  1.2.1 Global 상수 선언  ======================================
 ' Grid 1(vspdData1) - Operation 
Dim C_PlantCd
Dim C_PlantNm
Dim C_TrackingNo2
Dim C_Location
Dim C_TotQty
Dim C_TotAmt
Dim C_Price 
Dim C_PriceFlag
Dim C_PrevTotQty
Dim C_PrevTotAmt
Dim C_PrevPrice 
Dim C_PrevPriceFlag

 ' Grid 2(vspdData2) - Operation 
Dim C_SlCd
Dim C_SlNm
Dim C_TrackingNo
Dim C_GoodQty
Dim C_BadQty
Dim C_InspQty
Dim C_TransQty
Dim C_SchdRcptQty
Dim C_SchdIssueQty
Dim C_PrevGoodQty
Dim C_PrevBadQty 
Dim C_PrevInspQty
Dim C_PrevTrnsQty
Dim C_AllocationQty
Dim C_PickingQty



'==========================================  2.1.1 InitVariables()  =====================================
'	Name : InitVariables()																				=
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)				=
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE
	lgIntGrpCount     = 0						
	lgBlnFlgChgValue  = False
	lgStrPrevKeyIndex = ""                      
	lgLngCurRows      = 0
    lgOldRow		  = 0
	
End Sub

'==========================================  2.2.1 SetDefaultVal()  =====================================
'=	Name : SetDefaultVal()																				=
'=	Description : 화면 초기화(수량 Field나 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)		=
'========================================================================================================
Sub SetDefaultVal()
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End if
	frm1.txtItemCd.focus
End Sub

'============================= 2.2.3 InitSpreadSheet() ==================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'======================================================================================== 
Sub InitSpreadSheet(ByVal pvSpdNo)
	If pvSpdNo = "" Or pvSpdNo = "A" Then  
	
		Call InitSpreadPosVariables(pvSpdNo)

		ggoSpread.Source = frm1.vspdData1
		ggoSpread.Spreadinit "V20050217", , Parent.gAllowDragDropSpread
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1.vspdData1

			.ReDraw = false
				 
			.MaxCols = C_PrevPriceFlag+1											
			.MaxRows = 0
				
			Call GetSpreadColumnPos("A")
				
			ggoSpread.SSSetEdit     C_PlantCd,       "공장",			10,,,,2
			ggoSpread.SSSetEdit     C_PlantNm,       "공장명",			25
			ggoSpread.SSSetEdit     C_TrackingNo2,    "Tracking No.",	20
			ggoSpread.SSSetEdit     C_Location,      "Location",		20
			ggoSpread.SSSetFloat    C_TotQty,        "현재고수량",		15, parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
			ggoSpread.SSSetFloat    C_TotAmt,        "현재고금액",		15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000, parent.gComNumDec
			ggoSpread.SSSetFloat    C_Price,         "단가",			15, parent.ggUnitCostNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec 
			ggoSpread.SSSetEdit     C_PriceFlag,     "단가구분",		10,2
			ggoSpread.SSSetFloat    C_PrevTotQty,    "전월재고수량",	15, parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
			ggoSpread.SSSetFloat    C_PrevTotAmt,    "전월재고금액",	15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000, parent.gComNumDec
			ggoSpread.SSSetFloat    C_PrevPrice,     "전월단가",		15, parent.ggUnitCostNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec 
			ggoSpread.SSSetEdit     C_PrevPriceFlag, "전월단가구분",	13,2		

 			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			ggoSpread.SpreadLockWithOddEvenRowColor()
			ggoSpread.SSSetSplit2(2)
				
			.ReDraw = true
    
		End With
	End If
		
	If pvSpdNo = "" Or pvSpdNo = "B" Then    

		Call InitSpreadPosVariables(pvSpdNo)

		ggoSpread.Source = frm1.vspdData2
		ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread

		'------------------------------------------
		' Grid 2 - Component Spread Setting
		'------------------------------------------
	
		With frm1.vspdData2

			.ReDraw = false

			.MaxCols = C_PickingQty +1										
			.MaxRows = 0

			Call GetSpreadColumnPos("B")

			ggoSpread.SSSetEdit     C_SlCd,         "창고",				10,,,,2
			ggoSpread.SSSetEdit		C_SlNm,         "창고명",			18
			ggoSpread.SSSetEdit		C_TrackingNo,   "Tracking No.",		20
			ggoSpread.SSSetFloat	C_GoodQty,      "양품재고량",		15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_BadQty,       "불량재고량",		15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_InspQty,      "검사중수량",		15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_TransQty,     "이동중수량",		15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_SchdRcptQty,  "입고예정량",		15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_SchdIssueQty, "출고예정량",		15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_PrevGoodQty,  "전월양품재고량",	15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_PrevBadQty,   "전월불량재고량",	15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_PrevInspQty,  "전월검사중수량",	15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_PrevTrnsQty,	"전월이동중수량",	15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_AllocationQty,"재고할당량",		15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	
			ggoSpread.SSSetFloat	C_PickingQty,	"PICKING수량",		15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			
 			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			ggoSpread.SpreadLockWithOddEvenRowColor()
			ggoSpread.SSSetSplit2(2)
			.ReDraw = true
		End With
	End If
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
	If pvSpdNo = "" Or pvSpdNo = "A" Then						
	' Grid 1(vspdData1) - Operation 
		C_PlantCd             = 1
		C_PlantNm             = 2
		C_TrackingNo2		  = 3
		C_Location            = 4
		C_TotQty              = 5
		C_TotAmt              = 6
		C_Price               = 7
		C_PriceFlag           = 8
		C_PrevTotQty          = 9
		C_PrevTotAmt          = 10
		C_PrevPrice           = 11
		C_PrevPriceFlag       = 12
	End If
	If pvSpdNo = "" Or pvSpdNo = "B"  Then							
		' Grid 2(vspdData2) - Operation 
		C_SlCd                = 1
		C_SlNm                = 2
		C_TrackingNo		  = 3
		C_GoodQty		      = 4
		C_BadQty		      = 5
		C_InspQty             = 6
		C_TransQty            = 7
		C_SchdRcptQty         = 8
		C_SchdIssueQty        = 9
		C_PrevGoodQty         = 10
		C_PrevBadQty          = 11
		C_PrevInspQty         = 12
		C_PrevTrnsQty		  = 13	
		C_AllocationQty       = 14
		C_PickingQty		  = 15
	End If
End Sub


'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData1 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

		' Grid 1(vspdData1) - Operation 
 		C_PlantCd		= iCurColumnPos(1)
		C_PlantNm		= iCurColumnPos(2)
		C_TrackingNo2	= iCurColumnPos(3)
		C_Location		= iCurColumnPos(4)
		C_TotQty		= iCurColumnPos(5)
		C_TotAmt		= iCurColumnPos(6)
		C_Price			= iCurColumnPos(7)
		C_PriceFlag		= iCurColumnPos(8)
		C_PrevTotQty	= iCurColumnPos(9)
		C_PrevTotAmt	= iCurColumnPos(10)
		C_PrevPrice		= iCurColumnPos(11)
		C_PrevPriceFlag	= iCurColumnPos(12)

	Case "B"
 		ggoSpread.Source = frm1.vspdData2 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

 		' Grid 2(vspdData2) - Operation 
		C_SlCd          = iCurColumnPos(1)
		C_SlNm          = iCurColumnPos(2)
		C_TrackingNo	= iCurColumnPos(3)
		C_GoodQty		= iCurColumnPos(4)
		C_BadQty		= iCurColumnPos(5)
		C_InspQty       = iCurColumnPos(6)
		C_TransQty      = iCurColumnPos(7)
		C_SchdRcptQty   = iCurColumnPos(8)
		C_SchdIssueQty  = iCurColumnPos(9)
		C_PrevGoodQty   = iCurColumnPos(10)
		C_PrevBadQty    = iCurColumnPos(11)
		C_PrevInspQty   = iCurColumnPos(12)
		C_PrevTrnsQty	= iCurColumnPos(13)	
		C_AllocationQty = iCurColumnPos(14)
		C_PickingQty	= iCurColumnPos(15)

 	End Select
End Sub

'------------------------------------------  OpenItemCd()  --------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item Cd
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "품목 팝업"			
	arrParam(1) = "B_ITEM A"
	arrParam(2) = Trim(frm1.txtItemCd.Value)		
	arrParam(3) = ""																			
	arrParam(4) = "" 
	arrParam(5) = "품목"			
	
	arrField(0) = "A.ITEM_CD"															
	arrField(1) = "A.ITEM_NM"															
	arrField(2) = "A.SPEC"
	
	arrHeader(0) = "품목"															
	arrHeader(1) = "품목명"															
	arrHeader(2) = "규격"
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		Call SetItemCd(arrRet)
	End If	
	
End Function

'------------------------------------------  SetItemCd()  ----------------------------------------------
'	Name : SetItemCd()
'	Description : Item Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetItemCd(byRef arrRet)
	frm1.txtItemCd.Value    = arrRet(0)		
	frm1.txtItemNm.Value    = arrRet(1)
	frm1.txtSpec.value      = arrRet(2)
	frm1.txtItemCd.focus		
End Function

'------------------------------------------  OpenOnhandDtlRef()  -------------------------------------------------
'	Name : OpenOnhandDtlRefCode()
'	Description : OnahndStock detail Reference
'--------------------------------------------------------------------------------------------------------- 

Function OpenOnhandDtlRef()
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim Param1 
	Dim Param2
	Dim Param3
	Dim Param4	
	Dim Param5	
	Dim Param6
	Dim Param7 
	Dim Param8 
	Dim Param9
	
	If IsOpenPop = True Then Exit Function
	
	If frm1.vspdData1.MaxRows = 0 and frm1.vspdData2.MaxRows = 0 Then
		Call DisplayMsgBOX("900002","X","X","X")
		frm1.txtItemCd.focus
		Exit Function
	End if

	Param2 = Trim(frm1.txtItemCd.value)
	Param8 = Trim(frm1.txtItemNm.value)
	Param9 = Trim(frm1.txtBasicUnit.value)
	Param5 = "I"				
	
	if Param2 = "" then
		Call DisplayMsgBox("169903","X", "X", "X")  
		frm1.txtItemCd.focus
		Exit Function
	End If
	ggoSpread.Source = frm1.vspdData2    

	With frm1.vspdData2	    
		If .MaxRows = 0 Then
		    Call DisplayMsgBox("169902","X", "X", "X")  
			Exit Function
		else
			.Col = C_SlCd
			.Row = .ActiveRow
			 Param1 = Trim(.Text )
			.Col = C_SlNm
			.Row = .ActiveRow
			 Param7 = Trim(.Text )
			
		End If	
    End With
    
	ggoSpread.Source = frm1.vspdData1
	
	With frm1.vspdData1	
		If .MaxRows = 0 Then
			Exit Function
		else
			.Col = C_PlantCd
			.Row = .ActiveRow
			 Param4 = Trim(.Text )
			 .Col = C_TrackingNo2
			.Row = .ActiveRow
			 Param3 = Trim(.Text )
		End If	
	
		if Param4 = "" then
			Call DisplayMsgBox("169901","X", "X", "X")   
			Exit Function
		End IF
		
	End With
	
	IsOpenPop = True

	iCalledAspName = AskPRAspName("I2212RA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I2212RA1","x")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9), _
		 "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")		    
    	
	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.vspdData1.focus
		Exit Function
	End If	
	
End Function

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("0000111111")        

 	gMouseClickStatus = "SPC"   
    
 	Set gActiveSpdSheet = frm1.vspdData1
    
 	If frm1.vspdData1.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData1 
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col				
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey		
 			lgSortKey = 1
 		End If
 		Exit Sub
 	End If
 	
	
End Sub

Sub vspdData2_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")       

 	gMouseClickStatus = "SP2C"   
    
 	Set gActiveSpdSheet = frm1.vspdData2
    
 	If frm1.vspdData2.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData2
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col					
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey		
 			lgSortKey = 1
 		End If
 		Exit Sub
 	End If
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData1_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub 

Sub vspdData2_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SP2C" Then
       gMouseClickStatus = "SP2CR"
    End If
End Sub  

'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
  
'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData1_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub 
 
Sub vspdData2_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("B")
End Sub 
 '========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub 
 
'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : 그리드를 예전 상태로 복원한다.
'========================================================================================
Sub PopRestoreSpreadColumnInf()

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf() 
    
    Select case gActiveSpdSheet.id
    case "vaSpread1"
		Call InitSpreadSheet("A")	
	case "vaSpread2"
		Call InitSpreadSheet("B")
	End Select

    
    Call ggoSpread.ReOrderingSpreadData
End Sub 

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData1_GotFocus()
    ggoSpread.Source = frm1.vspdData1
End Sub

Sub vspdData2_GotFocus()
    ggoSpread.Source = frm1.vspdData2
End Sub

'==========================================================================================
'   Event Name : vspdData1_LeaveCell
'   Event Desc :
'==========================================================================================
Sub vspdData1_ScriptLeaveCell(ByVal Col, ByVal Row, Byval NewCol, Byval NewRow, Byval Cancel)
	
	If NewRow <= 0 Or Row = NewRow Then
		Exit Sub
	End If
	
	'frm1.vspdData2.MaxRows = 0
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData
	
	lgStrPrevKeyIndex2 = ""
	If DbDtlQuery(NewRow) = False Then	
		Exit Sub
	End If	
	
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	If frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1, NewTop) Then
		If lgStrPrevKeyIndex <> "" Then
			If DbQuery = False Then
				Exit Sub
			End if
		End if
	End if
End Sub

Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop)
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2, NewTop) Then
		If lgStrPrevKeyIndex2 <> "" Then
			If DbDtlQuery(frm1.vspdData1.ActiveRow) = False Then
				Exit Sub
			End if
		End if
	End if
End Sub

'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'********************************************************************************************************* 
Function FncQuery

	FncQuery = False
	
	on Error resume next
	Err.Clear
	
	Call InitVariables
	
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkFieldByCell(frm1.txtItemCd, "A",1) Then Exit Function
    
    Call SetToolbar("11000000000111")
	ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
	
	'-----------------------
	'Check Item CODE		
	'-----------------------
    If 	CommonQueryRs(" ITEM_NM, SPEC, BASIC_UNIT "," B_ITEM ", " ITEM_CD = " & FilterVar(frm1.txtItemCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		
		Call DisplayMsgBox("122600","X","X","X")
		frm1.txtItemNm.Value = ""
		frm1.txtBasicUnit.Value = ""
		frm1.txtSpec.Value = ""
		frm1.txtItemCd.focus
		Exit function
    End If
    lgF0 = Split(lgF0, Chr(11))
	lgF1 = Split(lgF1, Chr(11))
	lgF2 = Split(lgF2, Chr(11))
	frm1.txtItemNm.Value	= lgF0(0)
	frm1.txtSpec.Value		= lgF1(0)
	frm1.txtBasicUnit.Value = lgF2(0)

	If DbQuery = False Then	
		Exit Function
	End If
	
	FncQuery = False
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)								
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 화면 속성, Tab유무 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , True)      
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()	
    FncExit = True
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'********************************************  5.1 DbQuery()  *******************************************
' Function Name : DbQuery																				*
' Function Desc : This function is data query and display												*
'********************************************************************************************************
Function DbQuery()
	
    Err.Clear												                   
	    
    DbQuery = False											                   
	    
    Call LayerShowHide(1)
	    
    Dim strVal

    strVal =  BIZ_PGM_QRY1_ID & "?txtMode="				& parent.UID_M0001				& _
								"&txtItemCd="			& Trim(frm1.txtItemCd.value)	& _
								"&lgStrPrevKeyIndex="	& lgStrPrevKeyIndex				& _
								"&txtMaxRows="			& frm1.vspdData1.MaxRows

    Call RunMyBizASP(MyBizASP, strVal)						                    

    DbQuery = True                                                              

End Function
'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()															
    '-----------------------
    'Reset variables area
    '-----------------------
    Call ggoOper.LockField(Document, "Q")									
    lgStrPrevKeyIndex2 = ""
    ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData
	lgOldRow = 1
	frm1.vspdData1.focus
	
	Call DbDtlQuery(1)

End Function

'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtlQuery(ByRef Row) 

	Dim strVal			
   
	DbDtlQuery = False
			
	frm1.vspdData1.Row = Row
	frm1.vspdData1.Col = C_PlantCd
			
	If Trim(frm1.txtItemCd.value) = "" Then
		Call DisplayMsgBox("169903","X", "X", "X")
		frm1.txtItemCd.focus 
		Exit Function
	End if 

	Call LayerShowHide(1)

	strVal = BIZ_PGM_QRY2_ID	& "?txtMode="			 & parent.UID_M0001				& _
								  "&txtPlantCd="         & Trim(frm1.vspdData1.Text)	& _
								  "&txtItemCd="          & Trim(frm1.txtItemCd.value)	& _
								  "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex2			& _
								  "&txtMaxRows="         & frm1.vspdData2.MaxRows

	Call RunMyBizASP(MyBizASP, strVal)										

	DbDtlQuery = True

End Function

Function DbDtlQueryOk()											
	frm1.vspdData1.focus					
End Function
