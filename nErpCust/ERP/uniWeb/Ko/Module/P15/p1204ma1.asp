<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : Standard Routing
'*  3. Program ID           : P1204ma1
'*  4. Program Name         : Standard Routing Entry
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/04/03
'*  8. Modified date(Last)  : 2002/12/03
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_QRY_ID = "p1204mb1.asp"											'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID= "p1204mb2.asp"											'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_DEL_ID = "p1204mb3.asp"											'☆: 비지니스 로직 ASP명 

Dim C_OprNo
Dim C_WcCd
Dim C_WcPopup
Dim C_WcNm
Dim C_JobCd
Dim C_JobNm
Dim C_InsideFlg
Dim C_InsideFlgDesc
Dim C_RoutOrder
Dim C_RoutOrderDesc

<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim IsOpenPop
Dim lgChgValidToDtFlg
          
Dim BaseDate, StartDate

BaseDate = "<%=GetSvrDate%>"
StartDate = UniConvDateAToB(BaseDate, parent.gServerDateFormat, parent.gDateFormat)

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_OprNo			= 1
	C_WcCd			= 2
	C_WcPopup		= 3
	C_WcNm			= 4
	C_JobCd			= 5
	C_JobNm			= 6
	C_InsideFlg		= 7
	C_InsideFlgDesc	= 8
	C_RoutOrder		= 9
	C_RoutOrderDesc	= 10
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    
    'lgChgValidToDtFlg = False
    
    lgIntGrpCount = 100                         'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey    = 1                                       '⊙: initializes sort direction
    
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtValidFromDt.text = StartDate
	frm1.txtValidToDt.text =  UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

Sub InitSpreadSheet()
	Call initSpreadPosVariables()

	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021125",, parent.gAllowDragDropSpread

		.ReDraw = False

		.MaxCols = C_RoutOrderDesc + 1												'☜: 최대 Columns의 항상 1개 증가시킴 
		.MaxRows = 0

		.Col = .MaxCols																'☜: 공통콘트롤 사용 Hidden Column
		.ColHidden = True
    
		.Col = C_InsideFlg
		.ColHidden = True
    
		.Col = C_RoutOrder
		.ColHidden = True
		   
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit		C_OprNo, "공정", 10,,,3,2
		ggoSpread.SSSetEdit		C_WcCd, "작업장", 12,,,7,2
		ggoSpread.SSSetButton	C_WcPopup
		ggoSpread.SSSetEdit		C_WcNm, "작업장명", 27
		ggoSpread.SSSetCombo	C_JobCd, "공정작업코드",15
		ggoSpread.SSSetCombo	C_JobNm, "공정작업명",26
		ggoSpread.SSSetEdit		C_InsideFlg, "타입", 12
		ggoSpread.SSSetEdit		C_InsideFlgDesc, "타입", 10
		ggoSpread.SSSetCombo	C_RoutOrder, "공정단계", 12      
		ggoSpread.SSSetCombo	C_RoutOrderDesc, "공정단계", 10
    
		Call ggoSpread.MakePairsColumn(C_WcCd, C_WcPopup)

		Call ggoSpread.SSSetColHidden(C_InsideFlg, C_InsideFlg, True)
		Call ggoSpread.SSSetColHidden(C_RoutOrder, C_RoutOrder, True)
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		ggoSpread.SSSetSplit2(3)										'frozen 기능추가 
    
		.ReDraw = True
	
		Call SetSpreadLock 
    
    End With
    
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SpreadLock	C_OprNo, -1,C_OprNo
    ggoSpread.SpreadLock	C_WcNm, -1,C_WcNm
    ggoSpread.spreadLock	C_InsideFlg, -1,C_InsideFlg
    ggoSpread.spreadLock	C_RoutOrder, -1,C_RoutOrder
    ggoSpread.spreadLock	C_InsideFlgDesc, -1,C_InsideFlgDesc    
    ggoSpread.spreadLock	C_RoutOrderDesc, -1,C_RoutOrderDesc
    
    ggoSpread.SSSetRequired	C_WcCd, -1
    ggoSpread.SSSetProtected .vspdData.MaxCols, -1
    
    .vspdData.ReDraw = True

    End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SSSetRequired	C_OprNo,		pvStartRow, pvEndRow
		ggoSpread.SSSetRequired	C_WcCd,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_WcNm,			pvStartRow, pvEndRow
'		ggoSpread.SSSetProtected	C_JobNm,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_InsideFlg,	pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_RoutOrder,	pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_InsideFlgDesc, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_RoutOrderDesc, pvStartRow, pvEndRow
		.vspdData.ReDraw = True
    End With
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
			C_OprNo			= iCurColumnPos(1)
			C_WcCd			= iCurColumnPos(2)
			C_WcPopup		= iCurColumnPos(3)
			C_WcNm			= iCurColumnPos(4)
			C_JobCd			= iCurColumnPos(5)
			C_JobNm			= iCurColumnPos(6)
			C_InsideFlg		= iCurColumnPos(7)
			C_InsideFlgDesc	= iCurColumnPos(8)
			C_RoutOrder		= iCurColumnPos(9)
			C_RoutOrderDesc	= iCurColumnPos(10)
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
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData(1)
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim iColSep
	
	iColSep = Parent.gColSep
	
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1006", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    ggoSpread.Source = frm1.vspdData
	lgF0 = "" & iColSep & lgF0
	lgF1 = "" & iColSep & lgF1
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_JobCd
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_JobNm
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1201", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    ggoSpread.Source = frm1.vspdData
	lgF0 = "" & iColSep & lgF0
	lgF1 = "" & iColSep & lgF1
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_RoutOrder
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_RoutOrderDesc
    
End Sub

'==========================================  2.2.6 InitData()  =======================================
'	Name : InitData()
'	Description : Combo Display
'========================================================================================================= 
Sub InitData(ByVal lngStartRow)
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = lngStartRow To .MaxRows
			.Row = intRow
			.Col = C_JobCd
			intIndex = .value
			.col = C_JobNm
			.value = intindex
	
			.Row = intRow
			.Col = C_RoutOrder
			intIndex = .value
			.col = C_RoutOrderDesc
			.value = intindex
			
		Next	
	End With
End Sub

'------------------------------------------  OpenWcPopup()  -------------------------------------------------
'	Name : OpenWcPopup()
'	Description : WcPopup
'--------------------------------------------------------------------------------------------------------- 
Function OpenWcPopup(ByVal str)
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then		
		Call DisplayMsgBox("971012", "X", "공장", "X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "작업장팝업"	
	arrParam(1) = "P_WORK_CENTER"				
	arrParam(2) = str
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " AND VALID_TO_DT >=  " & FilterVar(BaseDate , "''", "S") & "" 
	arrParam(5) = "작업장"			
	
    arrField(0) = "WC_CD"	
    arrField(1) = "WC_NM"	
    arrField(2) = "HH" & parent.gcolsep & "INSIDE_FLG"
    arrField(3) = "CASE WHEN INSIDE_FLG=" & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("사내", "''", "S") & " ELSE " & FilterVar("외주", "''", "S") & " END"
    arrField(4) = "dbo.ufn_GetCodeName(" & FilterVar("P1013", "''", "S") & ", WC_MGR)"
    
    arrHeader(0) = "작업장"		
    arrHeader(1) = "작업장명"		
    arrHeader(2) = "작업장구분"		
    arrHeader(3) = "작업장구분"		
    arrHeader(4) = "작업장담당자"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetWc(arrRet)
	End If	
	
	Call SetActiveCell(frm1.vspdData, C_WcCd, frm1.vspdData.ActiveRow,"M","X","X")
	Set gActiveElement = document.activeElement

End Function

'------------------------------------------  OpenRouting()  -------------------------------------------------
'	Name : OpenRouting()
'	Description : RoutingPopup
'--------------------------------------------------------------------------------------------------------- 

Function OpenRouting(ByVal str)
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtRoutNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "공장", "X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "표준라우팅 팝업"	
	arrParam(1) = "(SELECT DISTINCT ROUT_NO, PLANT_CD, DESCRIPTION FROM P_STANDARD_ROUTING) A"				
	arrParam(2) = str
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S")
	arrParam(5) = "표준라우팅"			
	
    arrField(0) = "ROUT_NO"
    arrField(1) = "DESCRIPTION"	
       
    arrHeader(0) = "표준라우팅"		
    arrHeader(1) = "표준라우팅명"		
        
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetRouting(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtRoutNo.focus
	
End Function


'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : OpenPlantPopup
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant(ByVal str)
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = str
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    arrField(2) = "CUR_CD"	
    
    arrHeader(0) = "공장"		
    arrHeader(1) = "공장명"		
    arrHeader(2) = "통화코드"		
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  SetWc()  --------------------------------------------------
'	Name : SetWc()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetWc(Byval arrRet)
	With frm1
		.vspdData.Col = C_WcCd
		.vspdData.Text = arrRet(0)
		.vspdData.Col = C_WcNm
		.vspdData.Text = arrRet(1)
		.vspdData.Col = C_InsideFlg
		.vspdData.Text = UCase(arrRet(2))
		
		If UCase(arrRet(2)) = "Y" then
			.vspdData.Col = C_InsideFlgDesc 
			.vspdData.Text = "사내"
		Else
			.vspdData.Col = C_InsideFlgDesc 
			.vspdData.Text = "외주"
		End if			
		
		Call vspdData_Change(.vspdData.Col, .vspdData.Row)		' 변경이 일어났다고 알려줌 
	
	End With
End Function

'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)	
	frm1.txtPlantNm.value	 = arrRet(1) 	
End Function

'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetRouting()
'	Description : Routing Popup에서 Routing NO setting
'--------------------------------------------------------------------------------------------------------- 
Function SetRouting(byval arrRet)
	frm1.txtRoutNo.Value    = arrRet(0)		
	frm1.txtRoutingNm.Value    = arrRet(1)		
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field    
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	
    Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet
    '----------  Coding part  -------------------------------------------------------------
	
    Call InitComboBox
    
    Call SetToolbar("11101101001011")										'⊙: 버튼 툴바 제어 
    
    Call SetDefaultVal
    
    Call InitVariables
    
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtRoutNo.focus 
		Set gActiveElement = document.activeElement 
	Else 
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
	End If
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
Sub vspdData_Click(ByVal Col , ByVal Row )
    gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData
    Call SetPopupMenuItemInf("1101110111")
	
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
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================

Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

    '----------  Coding part  -------------------------------------------------------------
End Sub


'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'==========================================================================================

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1

	'----------  Coding part  -------------------------------------------------------------   
	With frm1.vspdData 
	
    ggoSpread.Source = frm1.vspdData
   
    If Row > 0 And Col = C_WcPopUp Then
        .Col = C_WcCd
        .Row = Row
        
        Call OpenWcPopup(.Text)
        
        Call SetActiveCell(frm1.vspdData,C_WcCd,Row,"M","X","X")
		Set gActiveElement = document.activeElement
        
    End If
    End With
End Sub


'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================

Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

    If Row >= NewRow Then
        Exit Sub
    End If

	'----------  Coding part  -------------------------------------------------------------   

End Sub


'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
	With frm1.vspdData
		
		.Row = Row
    
		Select Case Col
			Case  C_JobCd
				.Col = Col
				intIndex = .Value
				.Col = C_JobNm
				.Value = intIndex
			Case  C_JobNm
				.Col = Col
				intIndex = .Value
				.Col = C_JobCd
				.Value = intIndex
		End Select
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
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(parent.TBC_QUERY)  
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub

'=======================================================================================================
'   Event Name : txtValidFromDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtValidFromDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtValidFromDt.Action = 7 
		Call SetFocusToDocument("M")
		Frm1.txtValidFromDt.Focus
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtValidFromDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtValidFromDt_Change() 
	lgBlnFlgChgValue = True 
End Sub

'=======================================================================================================
'   Event Name : txtValidToDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtValidToDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtValidToDt.Action = 7 
		Call SetFocusToDocument("M")
		Frm1.txtValidToDt.Focus
	End If 
End Sub
'=======================================================================================================
'   Event Name : txtValidToDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================

Sub txtValidToDt_Change() 
	lgBlnFlgChgValue = True 
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
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")				'⊙: "Will you destory previous data"
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
    
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call ggoSpread.ClearSpreadData
    Call SetDefaultVal															'⊙: Initializes local global variables
    Call InitVariables

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
    End If     
    
    FncQuery = True																'⊙: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")				'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    frm1.txtRoutNo.value = ""
    
    Call ggoOper.ClearField(Document, "2")											'⊙: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
    Call SetDefaultVal
	Call InitVariables																'⊙: Initializes local global variables
	
	Call SetToolbar("11101101001011")										'⊙: 버튼 툴바 제어 
    
    frm1.txtRoutingNo.focus 
    Set gActiveElement = document.activeElement 
     
    FncNew = True																	'⊙: Processing is OK

End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
    Dim IntRetCD 
    
    FncDelete = False                                                       '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                                  '☆: 아래 메세지를 DB화 해서 이 라인으로 대체 
        Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")		            '⊙: "Will you destory previous data"	
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If
    
	If DbDelete = False Then   
		Exit Function           
    End If         
    
    FncDelete = True                                                        '⊙: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False AND ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                            '⊙: No data changed!!
        Exit Function
    End If

    '-----------------------
    'Check content area
    '-----------------------
    
    ggoSpread.Source = frm1.vspdData
	
	If lgIntFlgMode = parent.OPMD_CMODE Then
		If frm1.vspdData.MaxRows = 0 Then
			Call DisplayMsgBox("971012", "X", "공정", "X")
			Exit Function
		End If	
	End If
    
    If Not chkField(Document, "2") Then
		Exit Function
	End If
	ggoSpread.Source = frm1.vspdData
	If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
       Exit Function
    End If
    
    '-----------------------
    'Save function call area
    '-----------------------
   
    If DbSave = False Then   
		Exit Function           
    End If          '☜: Save db data
    
    FncSave = True                                                          '⊙: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
	If frm1.vspdData.maxrows < 1 Then Exit Function
		
	frm1.vspdData.focus
	Set gActiveElement = document.activeElement 
	frm1.vspdData.EditMode = True
	frm1.vspdData.ReDraw = False
    ggoSpread.Source = frm1.vspdData
    ggoSpread.CopyRow
    Call SetSpreadColor(frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow)
    
    frm1.vspdData.Col = C_OprNo
    frm1.vspdData.Row = frm1.vspdData.ActiveRow
    
    frm1.vspdData.Text = ""
    
    frm1.vspdData.ReDraw = True
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	If frm1.vspdData.maxrows < 1 Then Exit Function
	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
	Call InitData(1)
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
    Dim iIntReqRows

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	If IsNumeric(Trim(pvRowCnt)) Then
		iIntReqRows = CInt(pvRowCnt)
	Else
		iIntReqRows = AskSpdSheetAddRowCount()
		If iIntReqRows = "" Then
		    Exit Function
		End If
	End If

	With frm1
		.vspdData.focus
		Set gActiveElement = document.activeElement 
		ggoSpread.Source = .vspdData
		.vspdData.EditMode = True
		.vspdData.ReDraw = False
    
        ggoSpread.InsertRow , iIntReqRows
        
		Call SetSpreadColor(.vspdData.ActiveRow, .vspdData.ActiveRow + iIntReqRows - 1)
    
		.vspdData.ReDraw = True
    
    End With
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
    
    '----------------------
    ' 데이터가 없는 경우 
    '----------------------
    If frm1.vspdData.maxrows < 1 Then Exit Function
    
    ggoSpread.Source = frm1.vspdData
    
	lDelRows = ggoSpread.DeleteRow

End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint()                                               '☜: Protect system from crashing
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)												'☜: 화면 유형 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI, False)                                         '☜:화면 유형, Tab 유무 
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
	ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")			'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 

    DbQuery = False
    
    LayerShowHide(1)
		
    Err.Clear                                                               '☜: Protect system from crashing

	Dim strVal
    
    With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001							'☜: 
			strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)				'☆: 조회 조건 데이타 
			strVal = strVal & "&txtRoutNo=" & Trim(.hRoutNo.value)				'☆: 조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgCurDt=" & UniConvYYYYMMDDToDate(parent.gDateFormat, "1900","01","01")
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001							'☜: 
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'☆: 조회 조건 데이타 
			strVal = strVal & "&txtRoutNo=" & Trim(.txtRoutNo.value)				'☆: 조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgCurDt=" & UniConvYYYYMMDDToDate(parent.gDateFormat, "1900","01","01")
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True
    
End Function


'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk(ByVal LngMaxRow)														'☆: 조회 성공후 실행로직 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgBlnFlgChgValue = false
    
    Call SetToolbar("11111111001111")
    
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
	
	Call InitData(LngMaxRow)												'⊙: Job Name Setting
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If	
    lgIntFlgMode = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
	
End Function


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbSave() 

    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
	Dim iColSep
	Dim TmpBufferVal, TmpBufferDel
	Dim iTotalStrVal, iTotalStrDel
	Dim iValCnt, iDelCnt
	
    DbSave = False                                                          '⊙: Processing is NG
	
	If ValidDateCheck(frm1.txtValidFromDt, frm1.txtValidToDt) = False Then Exit Function      
	     
    LayerShowHide(1)
		
	With frm1
		
		.txtMode.value = parent.UID_M0002
    
		'-----------------------
		'Data manipulate area
		'-----------------------
		iColSep = Parent.gColSep
		ReDim TmpBufferVal(0) : ReDim TmpBufferDel(0)
		iValCnt = 0 : iDelCnt = 0
		
		'-----------------------
		'Data manipulate area
		'-----------------------

		For lRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = lRow
		    .vspdData.Col = 0
 
		    Select Case .vspdData.Text

		        Case ggoSpread.InsertFlag											'☜: 신규 
		        
					strVal = ""
					
					strVal = strVal & "C" & iColSep & lRow & iColSep				'⊙: C=Create, Sheet가 2개 이므로 구별				                
		            
		            .vspdData.Col = C_OprNo			
			        strVal = strVal & Trim(.vspdData.Text) & iColSep
			            
			        .vspdData.Col = C_WCCd			
			        strVal = strVal & Trim(.vspdData.Text) & iColSep

			        .vspdData.Col = C_JobCd	
			        strVal = strVal & Trim(.vspdData.Text) & iColSep

			        .vspdData.Col = C_InsideFlg
			        strVal = strVal & Trim(.vspdData.Text) & iColSep
			       
			        .vspdData.Col = C_RoutOrder
			        strVal = strVal & Trim(.vspdData.Text) & iColSep
			        
			        strVal = strVal & UNIConvDate(.txtValidFromDt.Text) & iColSep
			        strVal = strVal & UNIConvDate(.txtValidToDt.Text) & iColSep
			        strVal = strVal & .txtRoutingNm1.value & parent.gRowSep
			        
			        ReDim Preserve TmpBufferVal(iValCnt)
			        TmpBufferVal(iValCnt) = strVal
			        iValCnt = iValCnt + 1

		        Case ggoSpread.UpdateFlag											'☜: 신규 
		        
					strVal = ""
					
					strVal = strVal & "U" & iColSep & lRow & iColSep				'⊙: C=Create, Sheet가 2개 이므로 구별				                
		            
		            .vspdData.Col = C_OprNo			
			        strVal = strVal & Trim(.vspdData.Text) & iColSep
			            
			        .vspdData.Col = C_WCCd			
			        strVal = strVal & Trim(.vspdData.Text) & iColSep

			        .vspdData.Col = C_JobCd	
			        strVal = strVal & Trim(.vspdData.Text) & iColSep

			        .vspdData.Col = C_InsideFlg
			        strVal = strVal & Trim(.vspdData.Text) & iColSep
			       
			        .vspdData.Col = C_RoutOrder
			        strVal = strVal & Trim(.vspdData.Text) & iColSep
			        
			        strVal = strVal & UNIConvDate(.txtValidFromDt.Text) & iColSep
			        strVal = strVal & UNIConvDate(.txtValidToDt.Text) & iColSep
			        strVal = strVal & .txtRoutingNm1.value & parent.gRowSep
			        
			        ReDim Preserve TmpBufferVal(iValCnt)
			        TmpBufferVal(iValCnt) = strVal
			        iValCnt = iValCnt + 1
		            
		        Case ggoSpread.DeleteFlag											'☜: 삭제 
					
					strDel = ""
					
					strDel = strDel & "D" & iColSep & lRow & iColSep				'⊙: D=Delete
					
		            .vspdData.Col = C_OprNo	'10
		            strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep									
		            
		            ReDim Preserve TmpBufferDel(iDelCnt)
		            TmpBufferDel(iDelCnt) = strDel
		            iDelCnt = iDelCnt + 1
		            
				Case Else
					If lgBlnFlgChgValue = True Then
						
						strVal = ""
						
						strVal = strVal & "U" & iColSep & lRow & iColSep				'⊙: U=Update		
			
			            .vspdData.Col = C_OprNo			
				        strVal = strVal & Trim(.vspdData.Text) & iColSep
			            
				        .vspdData.Col = C_WCCd			
					    strVal = strVal & Trim(.vspdData.Text) & iColSep

						.vspdData.Col = C_JobCd	
					    strVal = strVal & Trim(.vspdData.Text) & iColSep

						.vspdData.Col = C_InsideFlg
						strVal = strVal & Trim(.vspdData.Text) & iColSep
			       
						.vspdData.Col = C_RoutOrder
						strVal = strVal & Trim(.vspdData.Text) & iColSep
			        
						strVal = strVal & UNIConvDate(.txtValidFromDt.Text) & iColSep
						strVal = strVal & UNIConvDate(.txtValidToDt.Text) & iColSep
						strVal = strVal & .txtRoutingNm1.value & parent.gRowSep
						
						ReDim Preserve TmpBufferVal(iValCnt)
						TmpBufferVal(iValCnt) = strVal
						iValCnt = iValCnt + 1
			        
					End If
		    End Select
		            
		Next
		
		iTotalStrDel = Join(TmpBufferDel, "")
		iTotalStrVal = Join(TmpBufferVal, "")
		
		.txtSpread.value = iTotalStrDel & iTotalStrVal
		
		.txtMaxRows.value = .vspdData.MaxRows
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										'☜: 비지니스 ASP 를 가동 
	
	End With
	
    DbSave = True                                                           '⊙: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
   
	frm1.txtRoutNo.value = frm1.txtRoutingNo.value 
	
	Call InitVariables
	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.MaxRows = 0
    Call MainQuery()

End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================

Function DbDelete()
	Dim strVal
	
	DbDelete = False														'⊙: Processing is NG
	
	LayerShowHide(1)
		
    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)				'☜: 삭제 조건 데이타 
    strVal = strVal & "&txtRoutingNo=" & Trim(frm1.txtRoutingNo.value)				'☜: 삭제 조건 데이타 
    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    
    DbDelete = True                                                         '⊙: Processing is NG 
End Function
Function DbDeleteOk()
	Call InitVariables
	Call FncNew()
End Function

Function SheetFocus(lRow, lCol)
	frm1.vspdData.focus
	frm1.vspdData.Row = lRow
	frm1.vspdData.Col = lCol
	frm1.vspdData.Action = 0
	frm1.vspdData.SelStart = 0
	frm1.vspdData.SelLength = len(frm1.vspdData.Text)
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>표준라우팅등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
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
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12NXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant frm1.txtPlantCd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>표준라우팅</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtRoutNo" SIZE=15 MAXLENGTH=7 tag="12XXXU" ALT = "표준라우팅"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRoutNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenRouting frm1.txtRoutNo.value">&nbsp;<INPUT TYPE=TEXT NAME="txtRoutingNm" SIZE=30 MAXLENGTH=40 tag="14"></TD>
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
								<TD CLASS="TD5" NOWRAP>표준라우팅</TD>
								<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtRoutingNo" SIZE=15 MAXLENGTH=7 tag="23XXXU" ALT="표준라우팅">&nbsp;<INPUT TYPE=TEXT NAME="txtRoutingNm1" SIZE=30 MAXLENGTH=50 tag="21" ALT="표준라우팅명"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>유효기간</TD>
								<TD CLASS="TD656" NOWRAP>
									<script language =javascript src='./js/p1204ma1_I545670807_txtValidFromDt.js'></script>&nbsp;~&nbsp;
									<script language =javascript src='./js/p1204ma1_I342035514_txtValidToDt.js'></script>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT="100%" COLSPAN = 2>
								<script language =javascript src='./js/p1204ma1_vspdData_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hRoutNo" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
