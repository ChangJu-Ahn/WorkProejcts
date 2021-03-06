Const BIZ_PGM_QRY_ID	= "p1712qb1.asp"											'☆: 비지니스 로직 ASP명 
Const EBOM_HISTORY_PGM_ID = "p1713qa1"												'☆: 설계BOM이력조회 ASP명 
Const PBOM_CREATE_PGM_ID = "p1713ma1"												'☆: 제조BOM작성(이관의뢰) ASP명 
Const EBOM_TO_PBOM_PGM_ID = "p1714ma1"												'☆: 제조BOM이관 ASP명 

Dim C_Level
Dim C_Seq
Dim C_ChildItemCd
Dim C_ChildItemNm
Dim C_Spec
Dim C_ChildItemUnit
Dim C_ItemAcctNm
Dim C_ProcTypeNm
Dim C_BomType
Dim C_ChildItemBaseQty
Dim C_ChildBasicUnit
Dim C_PrntItemBaseQty
Dim C_PrntBasicUnit
Dim C_SafetyLT
Dim C_LossRate
Dim C_SupplyFlgNm
Dim C_ValidFromDt
Dim C_ValidToDt
Dim	C_ECNNo
Dim	C_ECNDescription
Dim	C_ECNReasonCd
Dim	C_DrawingPath
Dim C_Remark
Dim C_HdrItemCd
Dim C_HdrBomNo
Dim C_Row	

Dim IsOpenPop

' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_Level					= 1
	C_Seq					= 2
	C_ChildItemCd			= 3
	C_ChildItemNm			= 4
	C_Spec					= 5
	C_ChildItemUnit			= 6
	C_ItemAcctNm			= 7
	C_ProcTypeNm			= 8
	C_BomType				= 9
	C_ChildItemBaseQty		= 10
	C_ChildBasicUnit		= 11
	C_PrntItemBaseQty		= 12
	C_PrntBasicUnit			= 13
	C_SafetyLT				= 14
	C_LossRate				= 15
	C_SupplyFlgNm			= 16
	C_ValidFromDt			= 17
	C_ValidToDt				= 18
	C_ECNNo					= 19	
	C_ECNDescription		= 20
	C_ECNReasonCd			= 21
	C_DrawingPath			= 22
	C_Remark				= 23
	C_HdrItemCd				= 24
	C_HdrBomNo				= 25
	C_Row					= 26
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKeyIndex = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1                                       '⊙: initializes sort direction
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtBaseDt.Text = StartDate
	frm1.txtBomNo.value = "E"
	frm1.cboItemAcct.value = ""
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
	
	Call initSpreadPosVariables()

	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20030109",, parent.gAllowDragDropSpread

		.ReDraw = False

		.MaxCols = C_Row												'☜: 최대 Columns의 항상 1개 증가시킴 
		.MaxRows = 0

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit 	C_Level, 				"레벨"		,	8
		ggoSpread.SSSetFloat	C_Seq,					"순서"		,	6,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,1,FALSE,"Z" 
		ggoSpread.SSSetEdit		C_ChildItemCd,			"자품목"	,	20,,,18,2
		ggoSpread.SSSetEdit 	C_ChildItemNm, 			"자품목명"	,	30
		ggoSpread.SSSetEdit 	C_Spec,	 				"규격"		,	30
		ggoSpread.SSSetEdit		C_ChildItemUnit,		"단위"		,	6,,,3,2
		ggoSpread.SSSetEdit		C_ItemAcctNm,			"품목계정"	,	10
		ggoSpread.SSSetEdit 	C_ProcTypeNm, 			"조달구분"	,	12
		ggoSpread.SSSetEdit		C_BomType,				"BOM Type"	,	10,,,1,2
		ggoSpread.SSSetFloat	C_ChildItemBaseQty,		"자품목기준수",	15, "8",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit 	C_ChildBasicUnit,		"단위"		,	6,,,3,2
		ggoSpread.SSSetFloat	C_PrntItemBaseQty,		"모품목기준수", 15, "8",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_PrntBasicUnit,		"단위"		,	6,,,3,2
		ggoSpread.SSSetFloat	C_SafetyLT, 			"안전L/T"	,	10,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,1,FALSE,"Z" 
		ggoSpread.SSSetFloat	C_LossRate,				"Loss율"	,	10,"7",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,1,FALSE,"Z" 
		ggoSpread.SSSetEdit		C_SupplyFlgNm,			"유무상구분",	8
		ggoSpread.SSSetDate		C_ValidFromDt,			"시작일"	,	11, 2, parent.gDateFormat
		ggoSpread.SSSetDate 	C_ValidToDt,			"종료일"	,	11, 2, parent.gDateFormat
		ggoSpread.SSSetEdit		C_ECNNo,				"설계변경번호", 18
		ggoSpread.SSSetEdit		C_ECNDescription,		"설계변경내용", 30
		ggoSpread.SSSetEdit		C_ECNReasonCd,			"설계변경근거", 10
		ggoSpread.SSSetEdit		C_DrawingPath,			"도면경로"	,	30
		ggoSpread.SSSetEdit 	C_Remark,	 			"비고"		,	30,,, 1000
		ggoSpread.SSSetEdit		C_HdrItemCd,			"Header품목",	5
		ggoSpread.SSSetEdit		C_HdrBomNo,				"header BOM No.", 5
		ggoSpread.SSSetEdit		C_Row,					"순서", 5

		ggoSpread.SSSetSplit2(3)											'frozen 기능 추가 

		Call ggoSpread.MakePairsColumn(C_Level, C_ChildItemCd)
		Call ggoSpread.MakePairsColumn(C_ChildItemBaseQty, C_ChildBasicUnit)
		Call ggoSpread.MakePairsColumn(C_PrntItemBaseQty, C_PrntBasicUnit)

		Call ggoSpread.SSSetColHidden(C_ChildItemUnit, C_ChildItemUnit, True)
		Call ggoSpread.SSSetColHidden(C_HdrItemCd, C_HdrBomNo, True)
		Call ggoSpread.SSSetColHidden(C_BomType, C_BomType, True)
		Call ggoSpread.SSSetColHidden(C_Row, C_Row, True)
    
		.ReDraw = True

		Call SetSpreadLock 

    End With
    
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_Level					= iCurColumnPos(1)
			C_Seq					= iCurColumnPos(2)
			C_ChildItemCd			= iCurColumnPos(3)
			C_ChildItemNm			= iCurColumnPos(4)
			C_Spec					= iCurColumnPos(5)
			C_ChildItemUnit			= iCurColumnPos(6)
			C_ItemAcctNm			= iCurColumnPos(7)
			C_ProcTypeNm			= iCurColumnPos(8)
			C_BomType				= iCurColumnPos(9)
			C_ChildItemBaseQty		= iCurColumnPos(10)
			C_ChildBasicUnit		= iCurColumnPos(11)
			C_PrntItemBaseQty		= iCurColumnPos(12)
			C_PrntBasicUnit			= iCurColumnPos(13)
			C_SafetyLT				= iCurColumnPos(14)
			C_LossRate				= iCurColumnPos(15)
			C_SupplyFlgNm			= iCurColumnPos(16)
			C_ValidFromDt			= iCurColumnPos(17)
			C_ValidToDt				= iCurColumnPos(18)
			C_ECNNo					= iCurColumnPos(19)
			C_ECNDescription		= iCurColumnPos(20)
			C_ECNReasonCd			= iCurColumnPos(21)
			C_DrawingPath			= iCurColumnPos(22)
			C_Remark				= iCurColumnPos(23)
			C_HdrItemCd				= iCurColumnPos(24)
			C_HdrBomNo				= iCurColumnPos(25)
			C_Row					= iCurColumnPos(26)
    End Select    
End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'=======================================================================================================
'   Event Name : txtBaseDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtBaseDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtBaseDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtBaseDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtBaseDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtBaseDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'------------------------------------------  OpenCondPlant()  -------------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"											' 팝업 명칭 
	arrParam(1) = "B_PLANT A, P_PLANT_CONFIGURATION B"					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.value)							' Code Condition
	arrParam(3) = ""													' Name Cindition
	arrParam(4) = "A.PLANT_CD = B.PLANT_CD AND B.ENG_BOM_FLAG = 'Y'"	' Where Condition
	arrParam(5) = "공장"												' TextBox 명칭 
	
    arrField(0) = "A.PLANT_CD"											' Field명(0)
    arrField(1) = "A.PLANT_NM"											' Field명(1)
    
    arrHeader(0) = "공장"											' Header명(0)
    arrHeader(1) = "공장명"											' Header명(1)
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenIremCd()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd(ByVal str, ByVal iPos)
	Dim arrRet
	Dim arrParam(5), arrField(11)
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If
	
	If frm1.txtPlantCd.value = "" or CheckPlant(frm1.txtPlantCd.value) = False Then
		Call DisplayMsgBox("971012", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)   ' Plant Code
	arrParam(1) = Trim(str)	' Item Code
	
	arrParam(2) = ""												' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrField(0) = 1							'ITEM_CD
    arrField(1) = 2 						'ITEM_NM											
    
	iCalledAspName = AskPRAspName("B1B11PA4")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA4", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemCd(arrRet, iPos)
	End If	
	  
	If iPos = 0 Then
		Call SetFocusToDocument("M")
		frm1.txtItemCd.focus			
	Else
		Call SetActiveCell(frm1.vspdData,C_ChildItemCd,frm1.vspdData.ActiveRow,"M","X","X")
		Set gActiveElement = document.activeElement
	End IF
	
End Function

'------------------------------------------  OpenBomNo()  -------------------------------------------------
'	Name : OpenBomNo()
'	Description : Condition BomNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBomNo(ByVal strItem, ByVal strBom)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then
		IsOpenPop = False
		Exit Function
	End If
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	If strItem = "" Then
		Call DisplayMsgBox("971012", "X", "품목", "X")
		
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	'---------------------------------------------
	 ' Parameter Setting
	 '--------------------------------------------- 

	IsOpenPop = True

	arrParam(0) = "BOM팝업"						' 팝업 명칭 
	arrParam(1) = "B_MINOR"							' TABLE 명칭 
	
	arrParam(2) = Trim(frm1.txtBomNo.value)		' Code Condition
	
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = "MAJOR_CD = " & FilterVar("P1401", "''", "S") & " "
	
	arrParam(5) = "BOM Type"						' TextBox 명칭 
	
    arrField(0) = "MINOR_CD"						' Field명(0)
    arrField(1) = "MINOR_NM"						' Field명(1)
        
    arrHeader(0) = "BOM Type"					' Header명(0)
    arrHeader(1) = "BOM 특성"					' Header명(1)
    
	arrRet = window.showModalDialog("../../ComAsp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetBomNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtBomNo.focus
	
End Function

'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)	
End Function

'------------------------------------------  SetItemCd()  --------------------------------------------------
'	Name : SetItemCd()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetItemCd(byval arrRet, ByVal iPos)
	
	If iPos = 0 Then
		frm1.txtItemCd.Value    = arrRet(0)		
		frm1.txtItemNm.Value    = arrRet(1)
		
	Else 
		With frm1.vspdData
			.Col = C_ChildItemCd
			.Row = .ActiveRow
			.Text = arrRet(0)		
			
			Call LookUpItemByPlant(arrRet(0),.Row)

		End With
		
	End IF
	
End Function

'------------------------------------------  SetBomNo()  --------------------------------------------------
'	Name : SetBomNo()
'	Description : Bom No Popup에서 return된 값 
'--------------------------------------------------------------------------------------------------------- 
Function SetBomNo(byval arrRet)
	frm1.txtBomNo.Value    = arrRet(0)		
	frm1.txtBomNm.Value    = arrRet(1)		
End Function


Sub SetCookieVal()
	
	If ReadCookie("txtItemCd") <> "" Then
		frm1.txtItemCd.Value = ReadCookie("txtItemCd")
		frm1.txtItemNm.value = ReadCookie("txtItemNm")
	End If	
	If ReadCookie("txtPlantCd") <> "" Then
		frm1.txtPlantCd.Value = ReadCookie("txtPlantCd")
		frm1.txtPlantNm.value = ReadCookie("txtPlantNm")
	End If	
	
	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtItemCd", ""
	WriteCookie "txtItemNm", ""

End Sub

'=============================================  2.5.1 LoadEBomHistory()  ======================================
'=	Event Name : LoadEBomHistory	Jump																			=
'=	Event Desc :																						=
'========================================================================================================
Function LoadEBomHistory()
	Dim IntRetCD
    
	 '------ Check previous data area ------ 
	If lgIntFlgMode = parent.OPMD_CMODE Then
		If lgBlnFlgChgValue = False Then
			Call DisplayMsgBox("900002", "X", "X", "X")
			Exit Function
		End If
	End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")
        If IntRetCd = vbNo Then
			Exit Function
		End If
	End If
	
	WriteCookie "txtPlantCd", Trim(frm1.txtPlantCd.value)
	WriteCookie "txtPlantNm", frm1.txtPlantNm.value 
	
	WriteCookie "txtItemCd", Trim(frm1.txtItemCd.value)
	WriteCookie "txtItemNm", frm1.txtItemNm.value 

	PgmJump(EBOM_HISTORY_PGM_ID)
End Function

'=============================================  2.5.1 LoadPBomCreate()  ======================================
'=	Event Name : LoadPBomCreate	Jump																			=
'=	Event Desc :																						=
'========================================================================================================
Function LoadPBomCreate()
	Dim IntRetCD
    
	 '------ Check previous data area ------ 
	If lgIntFlgMode = parent.OPMD_CMODE Then
		If lgBlnFlgChgValue = False Then
			Call DisplayMsgBox("900002", "X", "X", "X")
			Exit Function
		End If
	End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")
        If IntRetCd = vbNo Then
			Exit Function
		End If
	End If
	
	WriteCookie "txtPlantCd", Trim(frm1.txtPlantCd.value)
	WriteCookie "txtPlantNm", frm1.txtPlantNm.value 
	
	WriteCookie "txtItemCd", Trim(frm1.txtItemCd.value)
	WriteCookie "txtItemNm", frm1.txtItemNm.value 

	PgmJump(PBOM_CREATE_PGM_ID)
End Function

'=============================================  2.5.1 LoadEBomToPBom()  ======================================
'=	Event Name : LoadEBomToPBom	Jump																			=
'=	Event Desc :																						=
'========================================================================================================
Function LoadEBomToPBom()
	Dim IntRetCD
    
	 '------ Check previous data area ------ 
	If lgIntFlgMode = parent.OPMD_CMODE Then
		If lgBlnFlgChgValue = False Then
			Call DisplayMsgBox("900002", "X", "X", "X")
			Exit Function
		End If
	End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")
        If IntRetCd = vbNo Then
			Exit Function
		End If
	End If
	
	WriteCookie "txtPlantCd", Trim(frm1.txtPlantCd.value)
	WriteCookie "txtPlantNm", frm1.txtPlantNm.value 
	
	WriteCookie "txtItemCd", Trim(frm1.txtItemCd.value)
	WriteCookie "txtItemNm", frm1.txtItemNm.value 

	PgmJump(EBOM_TO_PBOM_PGM_ID)
End Function

'==========================================================================================
'   Function Name :SetFieldProp
'   Function Desc :여러 Case에 따른 Field들의 속성을 변경한다.
'==========================================================================================
Function SetFieldProp(ByVal lRow, ByVal Level, ByVal ProcType)
	
End Function

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
    gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
    Call SetPopupMenuItemInf("0000110111")

	If Row <= 0 Or Col < 0 Then
		Exit Sub
	End If
	
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData 

    If Row >= NewRow Then
        Exit Sub
    End If

	'----------  Coding part  -------------------------------------------------------------   

    End With

End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
    
		If lgStrPrevKeyIndex <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(parent.TBC_QUERY)  
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    
    '-----------------------
    'Erase contents area
    '-----------------------
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If
		
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If
		
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call InitVariables															'⊙: Initializes local global variables
    																			
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then   
		Exit Function           
    End If     																		'☜: Query db data

    FncQuery = True																'⊙: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	If frm1.vspdData.maxrows < 1 Then Exit Function
	frm1.vspdData.Redraw = False
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
	frm1.vspdData.Redraw = True
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
   Call parent.fncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)							'☜: 화면 유형 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI, False)	                   '☜:화면 유형, Tab 유무 
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
    
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 

    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      

    DbQuery = False
    
    LayerShowHide(1)
		
    Err.Clear                                                               '☜: Protect system from crashing

	Dim strVal

    With frm1
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001						'☜: 
		
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtBomNo=" & Trim(.txtBomNo.value) 
		strVal = strVal & "&txtBaseDt=" & Trim(.txtBaseDt.Text)
		
		If frm1.rdoSrchType1.checked = True Then
			strVal = strval & "&rdoSrchType=" & frm1.rdoSrchType1.value 
		ElseIf frm1.rdoSrchType2.checked = True Then
			strVal = strval & "&rdoSrchType=" & frm1.rdoSrchType2.value 
		End If       
		
		strVal = strVal & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex         '☜: Next key tag

		Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk(LngMaxRow)										'☆: 조회 성공후 실행로직 
	Call SetToolbar("11000000000111")								'⊙: 버튼 툴바 제어 
    If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
    End If
    lgIntFlgMode = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    
End Function


'========================================================================================
' Function Name : CheckPlant
' Function Desc : 생산Configuration에 설계공장으로 설정이 되었는지 Check
'========================================================================================
Function CheckPlant(ByVal sPlantCd)	
														
    Err.Clear																

    CheckPlant = False
    
	Dim arrVal, strWhere, strFrom

	If Trim(sPlantCd) <> "" Then
	
		strFrom = "B_PLANT A, P_PLANT_CONFIGURATION B"
		strWhere = 				" A.PLANT_CD = B.PLANT_CD AND B.ENG_BOM_FLAG = 'Y' AND"
		strWhere = strWhere & 	" A.PLANT_CD = " & FilterVar(sPlantCd, "''", "S")

		If Not CommonQueryRs("A.PLANT_NM", strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
    		Exit Function
		End If
	End If

	CheckPlant = True
	
End Function
