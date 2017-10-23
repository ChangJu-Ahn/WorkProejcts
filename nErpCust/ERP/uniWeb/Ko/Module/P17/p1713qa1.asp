<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : 설계BOM관리 
'*  2. Function Name        : 
'*  3. Program ID           : p1713qa1.asp
'*  4. Program Name         : 설계BOM변경이력 조회 
'*  5. Program Desc         :
'*  6. Component List        : 
'*  7. Modified date(First) : 2005/01/25
'*  8. Modified date(Last)  : 2005/01/25
'*  9. Modifier (First)     : Cho Yong Chill
'* 10. Modifier (Last)      : Cho Yong Chill
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/lgvariables.inc"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit

Const BIZ_PGM_ID = "p1713qb1.asp"											'☆: 비지니스 로직 ASP명 

Dim C_ActionType
Dim C_ModifiedDate
Dim C_Level
Dim C_Seq
Dim C_ChildItemCd
Dim C_ChildItemPopUp
Dim C_ChildItemNm
Dim C_Spec
Dim C_ChildItemUnit
Dim C_ItemAcct
Dim C_ItemAcctNm
Dim C_ProcType
Dim C_ProcTypeNm
Dim C_BomType
Dim C_BomTypePopup
Dim C_ChildItemBaseQty
Dim C_ChildBasicUnit
Dim C_ChildBasicUnitPopup
Dim C_PrntItemBaseQty
Dim C_PrntBasicUnit
Dim C_PrntBasicUnitPopup
Dim C_SafetyLT
Dim C_LossRate
Dim C_SupplyFlg
Dim C_SupplyFlgNm
Dim C_ValidFromDt
Dim C_ValidToDt
Dim C_ECNNo
Dim C_ECNNoPopup
Dim C_ECNDesc
Dim C_ReasonCd
Dim C_ReasonCdPopup
Dim C_ReasonNm
Dim C_DrawingPath
Dim C_Remark
Dim C_HdrItemCd
Dim C_HdrBomNo
Dim C_HdrProcType
Dim C_ItemValidFromDt
Dim C_ItemValidToDt
Dim C_ItemAcctGrp
Dim C_Row

<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim IsOpenPop
Dim StartDate

StartDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_ActionType			= 1 
	C_ModifiedDate			= 2 
	C_Level					= 3 
	C_Seq					= 4 
	C_ChildItemCd			= 5 
	C_ChildItemNm			= 6 
	C_Spec					= 7 
	C_ItemAcctNm			= 8 
	C_ProcTypeNm			= 9 
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
	C_ECNDesc				= 20
	C_ReasonNm				= 21
	C_DrawingPath			= 22
	C_Remark				= 23
	C_Row					= 24
End Sub                        
                               
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()     
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()            
                               
    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
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
	frm1.txtChgFromDt.Text = UNIDateAdd("D",-10, StartDate, parent.gDateFormat)	'☆: 초기화면에 뿌려지는 처음 날짜 
	frm1.txtChgToDt.Text = StartDate
	frm1.txtBomNo.value = "E"
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("Q", "P", "NOCOOKIE", "MA")%>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

Sub InitSpreadSheet()
	
	Call initSpreadPosVariables()

	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20050127",, parent.gAllowDragDropSpread

		.ReDraw = False

		.MaxCols = C_Row												'☜: 최대 Columns의 항상 1개 증가시킴 
		.MaxRows = 0

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit		C_ActionType,			"설계변경구분", 12
		ggoSpread.SSSetDate		C_ModifiedDate,			"설계변경일", 11, 2, parent.gDateFormat
		ggoSpread.SSSetEdit 	C_Level, 				"레벨", 8
		ggoSpread.SSSetFloat	C_Seq,					"순서", 6, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, 1, FALSE, "Z" 
		ggoSpread.SSSetEdit		C_ChildItemCd,			"자품목", 20,,, 18, 2
		ggoSpread.SSSetEdit 	C_ChildItemNm, 			"자품목명", 30
		ggoSpread.SSSetEdit 	C_Spec,	 				"규격", 30
		ggoSpread.SSSetEdit		C_ItemAcctNm,			"품목계정", 10
		ggoSpread.SSSetEdit 	C_ProcTypeNm, 			"조달구분", 12
		ggoSpread.SSSetFloat	C_ChildItemBaseQty,		"자품목기준수"	, 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,, "Z"
		ggoSpread.SSSetEdit 	C_ChildBasicUnit,		"단위"			, 6,,, 3, 2
		ggoSpread.SSSetFloat	C_PrntItemBaseQty,		"모품목기준수"	, 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,, "Z"
		ggoSpread.SSSetEdit		C_PrntBasicUnit,		"단위"			, 6,,, 3, 2
		ggoSpread.SSSetFloat 	C_SafetyLT, 			"안전L/T"	, 10, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, 1, FALSE, "Z" 
		ggoSpread.SSSetFloat	C_LossRate,				"Loss율"	, 10, "7", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, 1, FALSE, "Z" 
		ggoSpread.SSSetEdit		C_SupplyFlgNm,			"유무상구분", 10
		ggoSpread.SSSetDate		C_ValidFromDt,			"시작일"	, 11, 2, parent.gDateFormat
		ggoSpread.SSSetDate 	C_ValidToDt,			"종료일"	, 11, 2, parent.gDateFormat
		ggoSpread.SSSetEdit		C_ECNNo,				"설계변경번호", 18,,, 18, 2
 		ggoSpread.SSSetEdit		C_ECNDesc,				"설계변경내용", 30,,, 100
		ggoSpread.SSSetEdit		C_ReasonNm,				"설계변경근거명", 14
		ggoSpread.SSSetEdit		C_DrawingPath,			"도면경로", 30,,, 100
		ggoSpread.SSSetEdit 	C_Remark,	 			"비고"		, 30,,, 1000
		ggoSpread.SSSetEdit		C_Row,					"순서", 5
		
'		ggoSpread.SSSetSplit2(C_ChildItemNm)											'frozen 기능 추가 
		
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

			C_ActionType			= iCurColumnPos(1) 
			C_ModifiedDate			= iCurColumnPos(2)             
			C_Level					= iCurColumnPos(3) 
			C_Seq					= iCurColumnPos(4) 
			C_ChildItemCd			= iCurColumnPos(5) 
			C_ChildItemNm			= iCurColumnPos(6) 
			C_Spec					= iCurColumnPos(7) 
			C_ItemAcctNm			= iCurColumnPos(8)
			C_ProcTypeNm			= iCurColumnPos(9)
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
			C_ECNDesc				= iCurColumnPos(20)
			C_ReasonNm				= iCurColumnPos(21)
			C_DrawingPath			= iCurColumnPos(22)
			C_Remark				= iCurColumnPos(23)
			C_Row					= iCurColumnPos(24)
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
	arrParam(2) = Trim(frm1.txtPlantCd.Value)							' Code Condition
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
	Frm1.txtPlantCd.Focus
	
End Function

 '------------------------------------------  OpenECNNo()  -------------------------------------------------
 '	Name : OpenECNNo()
 '	Description : ECN PopUp
 '--------------------------------------------------------------------------------------------------------- 
 Function OpenECNNo()
 	Dim arrRet
 	Dim arrParam(4), arrField(10)
 	Dim iCalledAspName, IntRetCD
 
 	If IsOpenPop = True Then 
 		IsOpenPop = False
 		Exit Function
 	End If
 	
 	IsOpenPop = True
 	
 	arrParam(0) = frm1.txtECNNo.value   ' ECN No.
 
 	iCalledAspName = AskPRAspName("P1410PA1")
 	
 	If Trim(iCalledAspName) = "" Then
 		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P1410PA1", "X")
 		IsOpenPop = False
 		Exit Function
 	End If
 	
 	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
 		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
 	
 	IsOpenPop = False
 	
 	If arrRet(0) <> "" Then
 		Call SetECNNo(arrRet)
 	End If	
 	
 	Call SetFocusToDocument("M")
	Frm1.txtECNNo.Focus
 	
 End Function
 
'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6)
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
	arrParam(1) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(2) = ""						' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field명(0) : "ITEM_CD"	
    arrField(1) = 2 							' Field명(1) : "ITEM_NM"	
    
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
		Call SetItemCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	Frm1.txtItemCd.Focus
	
End Function

'------------------------------------------  OpenChildItemCd()  -------------------------------------------------
'	Name : OpenChildIremCd()
'	Description : Child Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenChildItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD
	
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
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)   ' Plant Code
	arrParam(1) = Trim(frm1.txtChildItemCd.value)	' Item Code
	arrParam(2) = ""					' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field명(0) : "CHILD_ITEM_CD"	
    arrField(1) = 2 							' Field명(1) : "CHILD_ITEM_NM"	
    
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
		Call SetChildItemCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	Frm1.txtChildItemCd.Focus
	
End Function

'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
	
	frm1.txtItemCd.focus
	Set gActiveElement = document.activeElement		
End Function

'------------------------------------------  SetECNNo()  --------------------------------------------------
'	Name : SetECNNo()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetECNNo(ByVal arrRet)
	frm1.txtECNNo.Value			= arrRet(0)		
	frm1.txtECNNoDesc.Value		= arrRet(1)
	
	frm1.txtItemCd.focus
	Set gActiveElement = document.activeElement	
End Function
'------------------------------------------  SetItemCd()  --------------------------------------------------
'	Name : SetItemCd()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemCd(byval arrRet)
	frm1.txtItemCd.Value    = arrRet(0)		
	frm1.txtItemNm.Value    = arrRet(1)
		
	frm1.txtChildItemCd.focus
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  SetChildItemCd()  --------------------------------------------------
'	Name : SetChildItemCd()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetChildItemCd(byval arrRet)
	frm1.txtChildItemCd.Value    = arrRet(0)		
	frm1.txtChildItemNm.Value    = arrRet(1)
	
	frm1.txtPlantCd.focus
	Set gActiveElement = document.activeElement
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

'=======================================================================================================
'   Event Name : txtChgFromDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtChgFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtChgFromDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtChgFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtChgToDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtChgToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtChgToDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtChgToDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtChgFromDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtChgFromDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtChgToDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtChgToDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
    
    Call LoadInfTB19029  
	
	Call AppendNumberPlace("6", "3", "0")
	Call AppendNumberPlace("7", "2", "2")   
	Call AppendNumberPlace("8", "11", "6")	
	Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec, FALSE,, ggStrMinPart, ggStrMaxPart)
	
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet

    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal

    Call InitVariables                                                      '⊙: Initializes local global variables
    
	Call SetToolbar("11000000000011")									'⊙: 버튼 툴바 제어 
    
    If parent.gPlant <> "" and CheckPlant(parent.gPlant) = True Then
		frm1.txtPlantCd.value = parent.gPlant
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement  
		Call txtPlantCd_OnChange()
	Else
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 	
	End If
	
	Call SetCookieVal
   
End Sub

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
Sub vspdData_Click(ByVal Col, ByVal Row )
    Set gActiveSpdSheet = frm1.vspdData
    
	gMouseClickStatus = "SPC"
    Call SetPopupMenuItemInf("1101110111")

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
	If Button = 2 And gMouseClickStatus = "SPC" Then
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
' Event Name : txtPlantCd_OnChange
' Event Desc : 공장코드 변경시 공장명 초기화 
'========================================================================================
Sub txtPlantCd_OnChange()
	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If	
End Sub
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 

    FncQuery = False															'⊙: Processing is NG
    Err.Clear	
    																'☜: Protect system from crashing
    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		'데이타가 변경되었습니다. 조회하시겠습니까?    
    	IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			'⊙: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
   
    '-----------------------
    'Erase contents area
    '-----------------------
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If
		
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If
	
	If frm1.txtChildItemCd.value = "" Then
		frm1.txtChildItemNm.value = ""
	End If
	
	If frm1.txtECNNo.value = "" Then
		frm1.txtECNNoDesc.value = ""
	End If			
    
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call ggoSpread.ClearSpreadData
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
    Dim iStrPlantCd, iStrItemCd, iStrChildItemCd
    Dim iStrEcnNo, iStrSrchType, iStrBOMType

    DbQuery = False

    LayerShowHide(1)
		
    Err.Clear                                                               '☜: Protect system from crashing

	Dim strVal
    
    With frm1
    
		iStrPlantCd = UCase(Trim(.txtPlantCd.value))
		iStrItemCd = UCase(Trim(.txtItemCd.value))
		iStrChildItemCd = UCase(Trim(.txtChildItemCd.value))
		iStrEcnNo = UCase(Trim(.txtECNNo.value))
		
		If .rdoSrchType1.checked = True Then
			iStrSrchType = .rdoSrchType1.value 
		ElseIf .rdoSrchType2.checked = True Then
			iStrSrchType = .rdoSrchType2.value 
		End If  

		iStrBOMType = UCase(Trim(.txtBomNo.value))
		    
		strVal = BIZ_PGM_ID & "?rdoSrchType=" & iStrSrchType
		strVal = strVal & "&txtPlantCd=" & iStrPlantCd						'☆: 조회 조건 데이타 
		strVal = strVal & "&txtItemCd=" & iStrItemCd						'☆: 조회 조건 데이타 
		strVal = strVal & "&txtChildItemCd=" & iStrChildItemCd				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtECNNo=" & iStrEcnNo							'☆: 조회 조건 데이타 
		strVal = strVal & "&txtBomNo=" & "E" 
		strVal = strVal & "&txtChgFromDt=" & Trim(.txtChgFromDt.Text)
		strVal = strVal & "&txtChgToDt=" & Trim(.txtChgToDt.Text)
		strVal = strVal & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex         '☜: Next key tag

		Call RunMyBizASP(MyBizASP, strVal)									'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk(LngMaxRow)										'☆: 조회 성공후 실행로직 
	
	
	Call ggoOper.LockField(Document, "Q")							'⊙: This function lock the suitable field
	
	Call SetToolbar("11000000000011")

	Call txtPlantCd_OnChange()	'2003-08-11
	
'	frm1.vspdData.focus
'	lgBlnFlgChgValue = False

End Function
	
Function DbQueryNotOk()
    lgIntFlgMode = parent.OPMD_CMODE								'⊙: Indicates that current mode is Update mode
	
	Call SetToolbar("11001111001001") 
	  
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
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<!--'#########################################################################################################
'       					6. Tag부 
'#########################################################################################################-->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>설계BOM이력조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=90%>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=28 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>변경일</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtChgFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="시작일" tag="11"> </OBJECT>');</SCRIPT>
										&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtChgToDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="종료일" tag="12"> </OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 ROWSPAN=3 NOWRAP>단계</TD>
									<TD CLASS=TD6 ROWSPAN=3 NOWRAP>
										<INPUT TYPE="RADIO" NAME="rdoStepType" ID="rdoSrchType1" CLASS="RADIO" tag="1X" Value="1"><LABEL FOR="rdoStepType1">단단계</LABEL>
										<BR>
										<INPUT TYPE="RADIO" NAME="rdoStepType" ID="rdoSrchType2" CLASS="RADIO" tag="1X" Value="2" CHECKED><LABEL FOR="rdoStepType2">다단계</LABEL>
									</TD>
									
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>모품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="12XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=16 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>설계변경번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtECNNo" SIZE=18 MAXLENGTH=18 tag="11XXXU" ALT="설계변경번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnEcnCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenECNNo()"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>자품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtChildItemCd" SIZE=18 MAXLENGTH=18 tag="11XXXU" ALT="자품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnChildItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenChildItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtChildItemNm" SIZE=16 tag="14"></TD>									
									<TD CLASS=TD5 NOWRAP>설계변경내용</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtECNNoDesc" SIZE=35 tag="14"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD HEIGHT="100%" COLSPAN = 4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hShiftCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hBomType" tag="24">
<INPUT TYPE=HIDDEN NAME="txtBomNo" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
