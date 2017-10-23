<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2511MA1
'*  4. Program Name         : 검사의뢰조회 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_QRY_ID = "q2511mb1.asp"				'☆: 비지니스 로직 ASP명 

Const BIZ_PGM_JUMP1_ID = "Q2111MA1"
Const BIZ_PGM_JUMP2_ID = "Q2211MA1"
Const BIZ_PGM_JUMP3_ID = "Q2311MA1"
Const BIZ_PGM_JUMP4_ID = "Q2411MA1"

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim C_InspReqNo '= 1
Dim C_ItemCd '= 2
Dim C_ItemNm '= 3 
Dim C_BpCd '= 4
Dim C_BpNm '= 5
Dim C_WcCd '= 6
Dim C_WcNm '= 7
Dim C_InspReqDt '= 8
Dim C_LotNo '= 9
Dim C_LotSubNo '= 10
Dim C_LotSize '= 11
Dim C_InspStatusFlag '= 12

Dim IsOpenPop

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part-------------------------------------------------------------------
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.cboInspClassCd.value = "R"
	frm1.cboStatusFlag.value = "N"
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE","MA") %>
End Sub

'=============================================== 2.2.3 SpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021225", , Parent.gAllowDragDropSpread

		.ReDraw = false
		
		.MaxCols = C_InspStatusFlag + 1
		.MaxRows = 0
		
		Call GetSpreadColumnPos("A")

    	ggoSpread.SSSetEdit C_InspReqNo,	"검사의뢰번호", 13, 0, -1, 18
    	ggoSpread.SSSetEdit C_ItemCd,		"품목코드", 15, 0, -1, 18
    	ggoSpread.SSSetEdit C_ItemNm,		"품목명", 20, 0, -1, 40 
    	ggoSpread.SSSetEdit C_BpCd,			"공급처코드", 10, 0, -1, 10
    	ggoSpread.SSSetEdit C_BpNm,			"공급처명", 20, 0, -1, 40
    	ggoSpread.SSSetEdit C_WcCd,			"작업장코드", 10, 0, -1, 7
    	ggoSpread.SSSetEdit C_WcNm,			"작업장명", 20, 0, -1, 40
    	ggoSpread.SSSetEdit C_InspReqDt,	"검사의뢰일", 12, 2, -1, 10
    	ggoSpread.SSSetEdit C_LotNo,		"로트번호", 8, 0, -1, 12
    	ggoSpread.SSSetEdit C_LotSubNo,		"순번", 10, 1, -1, 3
    	ggoSpread.SSSetFloat C_LotSize,		"로트크기", 12, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
    	ggoSpread.SSSetEdit C_InspStatusFlag,  "검사진행현황", 14, 2, -1, 40
    		
   		'Column Frozen
 		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		.ReDraw = true

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

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("Q0001", "''", "S") & " ORDER BY MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
	Call SetCombo2(frm1.cboInspClassCd , lgF0, lgF1, Chr(11))

	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("Q0013", "''", "S") & " ORDER BY MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
	Call SetCombo2(frm1.cboStatusFlag , lgF0, lgF1, Chr(11))
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()	
	C_InspReqNo = 1
	C_ItemCd	= 2
	C_ItemNm	= 3 
	C_BpCd		= 4
	C_BpNm		= 5
	C_WcCd		= 6
	C_WcNm		= 7
	C_InspReqDt = 8
	C_LotNo		= 9
	C_LotSubNo	= 10
	C_LotSize	= 11
	C_InspStatusFlag = 12
End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 		Case "A"
 			ggoSpread.Source = frm1.vspdData
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_InspReqNo = iCurColumnPos(1)
			C_ItemCd	= iCurColumnPos(2)
			C_ItemNm	= iCurColumnPos(3) 
			C_BpCd		= iCurColumnPos(4)
			C_BpNm		= iCurColumnPos(5)
			C_WcCd		= iCurColumnPos(6)
			C_WcNm		= iCurColumnPos(7)
			C_InspReqDt = iCurColumnPos(8)
			C_LotNo		= iCurColumnPos(9)
			C_LotSubNo	= iCurColumnPos(10)
			C_LotSize	= iCurColumnPos(11)
			C_InspStatusFlag = iCurColumnPos(12)
 	End Select
End Sub

'------------------------------------------  OpenPlant() -------------------------------------------------
'	Name : OpenPlant()
'	Description :Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	OpenPlant = false
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			

    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	

    arrHeader(0) = "공장코드"		
    arrHeader(1) = "공장명"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam,arrField, arrHeader), _
		"dialogWIDTH=420px; dialogHEIGHT=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtPlantCd.Value    = arrRet(0)
		frm1.txtPlantNm.Value    = arrRet(1)
	End If	
	
	frm1.txtPlantCd.Focus
	Set gActiveElement = document.activeElement
	OpenPlant = true	
End Function

'------------------------------------------  OpenItem()  -------------------------------------------------
'	Name : OpenItem()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItem()
	OpenItem = false
	
	Dim arrRet
	Dim arrParam1, arrParam2, arrParam3, arrParam4, arrParam5
	Dim arrField(6)
	Dim iCalledAspName, IntRetCD
	
	'공장코드가 있는 지 체크 
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("220705","X","X","X") 		'공장정보가 필요합니다 
		frm1.txtPlantCd.Focus
		Set gActiveElement = document.activeElement
		Exit Function	
	End If
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	arrParam1 = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam2 = Trim(frm1.txtPlantNm.Value)	' Plant Name
	arrParam3 = Trim(frm1.txtItemCd.Value)	' Item Code
	arrParam4 = ""	'Trim(frm1.txtItemNm.Value)	' Item Name
	arrParam5 = Trim(frm1.cboInspClassCd.Value)
  
	iCalledAspName = AskPRAspName("q1211pa2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "q1211pa2", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam1, arrParam2, arrParam3, arrParam4, arrParam5, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		  
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtItemCd.Value    = arrRet(0)		
		frm1.txtItemNm.Value    = arrRet(1)		
	End If	
	
	frm1.txtItemCd.Focus
	Set gActiveElement = document.activeElement
	OpenItem = true
End Function

'------------------------------------------  OpenBp()  -------------------------------------------------
'	Name : OpenBp()
'	Description : Bp PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	Dim Param1, Param2, Param3, Param4
	Dim iCalledAspName, IntRetCD
	
	If UCase(frm1.txtBpCd.ClassName) = UCase(Parent.UCN_PROTECTED)  Then
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처 팝업"					' 팝업 명칭 
	arrParam(1) = "B_BIZ_PARTNER"					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtBpCd.Value)					' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = "(BP_TYPE = " & FilterVar("CS", "''", "S") & " Or BP_TYPE = " & FilterVar("S", "''", "S") & " )"			' Where Condition	
	arrParam(5) = "공급처"						' 조건필드의 라벨 명칭	
	
    arrField(0) = "BP_CD"								' Field명(0)
    arrField(1) = "BP_NM"								' Field명(1)
    
    arrHeader(0) = "공급처코드"					' Header명(0)
    arrHeader(1) = "공급처명"						' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtBpCd.Value = arrRet(0)
		frm1.txtBpNm.Value = arrRet(1)
	End If	
	
	frm1.txtBpCd.Focus
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  OpenCust()  -------------------------------------------------
'	Name : OpenCust()
'	Description : Cust PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenCust()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If UCase(frm1.txtCustCd.ClassName) = UCase(Parent.UCN_PROTECTED)  Then
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "거래처 팝업"					' 팝업 명칭 
	arrParam(1) = "B_BIZ_PARTNER"					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtCustCd.Value)					' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = "(BP_TYPE = " & FilterVar("CS", "''", "S") & " Or BP_TYPE = " & FilterVar("C", "''", "S") & " )"			' Where Condition	
	arrParam(5) = "거래처"						' 조건필드의 라벨 명칭	
	
    arrField(0) = "BP_CD"								' Field명(0)
    arrField(1) = "BP_NM"								' Field명(1)
    
    arrHeader(0) = "거래처코드"					' Header명(0)
    arrHeader(1) = "거래처명"						' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtCustCd.Value = arrRet(0)
		frm1.txtCustNm.Value = arrRet(1)
	End If	
	
	frm1.txtCustCd.Focus
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  OpenWc()  -------------------------------------------------
'	Name : OpenWc()
'	Description : Wc PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenWc()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	'공장코드가 있는 지 체크 
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("220705","X","X","X") 		'공장정보가 필요합니다 
		frm1.txtPlantCd.Focus
		Set gActiveElement = document.activeElement
		Exit Function	
	End If
	
	If UCase(frm1.txtWcCd.ClassName) = UCase(Parent.UCN_PROTECTED)  Then
		Exit Function
	End If
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "작업장 팝업"					' 팝업 명칭 
	arrParam(1) = "P_WORK_CENTER"					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtWcCd.Value)					' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " " 	' Where Condition
	arrParam(5) = "작업장"						' 조건필드의 라벨 명칭	
	
    arrField(0) = "Wc_CD"								' Field명(0)
    arrField(1) = "Wc_NM"								' Field명(1)
    
    arrHeader(0) = "작업장코드"					' Header명(0)
    arrHeader(1) = "작업장명"						' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtWcCd.Value = arrRet(0)
		frm1.txtWcNm.Value = arrRet(1)
	End If	
	
	frm1.txtWcCd.Focus
	Set gActiveElement = document.activeElement
End Function

'=============================================  2.5.1 LoadInspection()======================================
'=	Event Name : LoadInspection
'=	Event Desc :
'========================================================================================================
Function LoadInspection()
	Dim intRetCD
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If intRetCD = vbNo Then	Exit Function
	End If
	
	If lgIntFlgMode = Parent.OPMD_CMODE Then
		If lgBlnFlgChgValue = False Then
			Call DisplayMsgBox("900002", "X", "X", "X")
			Exit Function
		End If
	End If
	
	With frm1
		WriteCookie "txtPlantCd", UCase(Trim(.txtPlantCd.value))
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "cboInspClassCd", Trim(.cboInspClassCd.value)
		WriteCookie "txtInspReqNo", GetSpreadText(.vspdData,C_InspReqNo,.vspdData.ActiveRow,"X","X")
		
		If .hStatusFlag.Value =  "N" Then
			WriteCookie "IsInspectionRequest", "True"
		End If
		
		If .cboInspClassCd.Value = "R" Then
			PgmJump(BIZ_PGM_JUMP1_ID)
		ElseIf  .cboInspClassCd.Value = "P" Then
			PgmJump(BIZ_PGM_JUMP2_ID)
		ElseIf  .cboInspClassCd.Value = "F" Then
			PgmJump(BIZ_PGM_JUMP3_ID)
		ElseIf  .cboInspClassCd.Value = "S" Then
			PgmJump(BIZ_PGM_JUMP4_ID)
		End If	

	End With
End Function

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
 	gMouseClickStatus = "SPC"   
    
 	Set gActiveSpdSheet = frm1.vspdData
	Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
    
 	If frm1.vspdData.MaxRows = 0 Then Exit Sub
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData 
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey			'Sort in Descending
 			lgSortKey = 1
 		End If
	Else
 	End If
End Sub

'========================================================================================
' Function Name : vspdData_DblClick
' Function Desc : 그리드 해더 더블클릭시 네임 변경 
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
 	If Row <= 0 Then Exit Sub
    If frm1.vspdData.MaxRows = 0 Then Exit Sub
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call InitVariables                                                      '⊙: Initializes local global variables
	Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet
	Call InitComboBox
	Call SetDefaultVal
	Call SetToolbar("11000000000011")										'⊙: 버튼 툴바 제어								
	If frm1.txtPlantCd.value = "" Then
	   frm1.txtPlantCd.value = UCase(Parent.gPlant)
	   frm1.txtPlantNm.value = Parent.gPlantNm
	End IF
	frm1.txtPlantCd.focus
End Sub

'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
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
    Call InitSpreadSheet
    Call ggoSpread.ReOrderingSpreadData
 	'------ Developer Coding part (Start)
	Call DbQueryOk
 	'------ Developer Coding part (End) 	
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )    	
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'==========================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With frm1.vspdData
		If Row >= NewRow Then Exit Sub
	End With
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft, ByVal NewTop )
	If OldLeft <> NewLeft Then Exit Sub
	
	 '----------  Coding part -------------------------------------------------------------
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop)  Then	'☜: 재쿼리 체크 
		If lgStrPrevKey <> ""Then		'⊙:다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If CheckRunningBizProcess = True Then Exit Sub
		
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If
End Sub

'=======================================================================================================
'   Event Name : txtInspReqDtFr_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtInspReqDtFr_DblClick(Button)
    If Button = 1 Then
        frm1.txtInspReqDtFr.Action = 7
    End If
End Sub

'=======================================================================================================
'   Event Name : txtInspReqDtFr_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtInspReqDtFr_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtInspReqDtTo_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtInspReqDtTo_DblClick(Button)
    If Button = 1 Then
        frm1.txtInspReqDtTo.Action = 7
    End If
End Sub

'=======================================================================================================
'   Event Name : txtInspReqDtTo_Change
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtInspReqDtTo_Change()
    lgBlnFlgChgValue = True
End Sub

Function  txtInspReqDtFr_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Function

Function  txtInspReqDtTo_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Function

Function  txtPlantCd_onChange()
	If Trim(frm1.txtPlantCd.Value) = "" Then
			frm1.txtPlantNm.Value = ""
	End If
End Function

Function  txtItemCd_onChange()
	If Trim(frm1.txtItemCd.Value) = "" Then
			frm1.txtItemNm.Value = ""
	End If
End Function

Function  txtBpCd_onChange()
	If Trim(frm1.txtBpCd.Value) = "" Then
			frm1.txtBpNm.Value = ""
	End If
End Function

Function  txtWcCd_onChange()
	If Trim(frm1.txtWcCd.Value) = "" Then
			frm1.txtWcNm.Value = ""
	End If
End Function

Function  cboInspClassCd_onChange()
	With frm1
		Select Case .cboInspClassCd.Value
			Case "R"
				.txtWcCd.Value = ""
				.txtWcNm.Value = ""
				.txtCustCd.Value = ""
				.txtCustNm.Value = ""
				ProtectTag(.txtWcCd)		
				ProtectTag(.txtCustCd)		
				ReleaseTag(.txtBpCd)
			Case "P"
				.txtBpCd.Value = ""
				.txtBpNm.Value = ""
				.txtCustCd.Value = ""
				.txtCustNm.Value = ""
				ProtectTag(.txtBpCd)		
				ProtectTag(.txtCustCd)
				ReleaseTag(.txtWcCd)		
			Case "F"
				.txtBpCd.Value = ""
				.txtBpNm.Value = ""
				.txtCustCd.Value = ""
				.txtCustNm.Value = ""
				.txtWcCd.Value = ""
				.txtWcNm.Value = ""
				ProtectTag(.txtBpCd)		
				ProtectTag(.txtCustCd)
				ProtectTag(.txtWcCd)
			Case "S"
				.txtBpCd.Value = ""
				.txtBpNm.Value = ""
				.txtWcCd.Value = ""
				.txtWcNm.Value = ""
				ProtectTag(.txtBpCd)		
				ProtectTag(.txtWcCd)
				ReleaseTag(.txtCustCd)
		End Select 
	End With
End Function

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        						'⊙: Processing is NG
    
    Err.Clear                                                               						'☜: Protect system from crashing

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")							'⊙: Clear Contents  Field
    Call InitVariables										'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then	Exit Function
    	
    If ValidDateCheck(frm1.txtInspReqDtFr, frm1.txtInspReqDtTo) = False Then Exit Function
    	
    ggoSpread.Source = frm1.vspdData
    With frm1.vspdData
		If frm1.cboInspClassCd.Value = "R" Then
			.Col = C_BpCd		
			.ColHidden =False
			.Col = C_BpNm	
			.ColHidden = False
			.Col = C_WcCd		
			.ColHidden = True
			.Col = C_WcNm	
			.ColHidden = True
			ggoSpread.SSSetEdit C_BpCd, "공급처코드", 10, 0, -1, 10
			ggoSpread.SSSetEdit C_BpNm, "공급처명", 20, 0, -1, 40
		ElseIf frm1.cboInspClassCd.Value = "P" Then
			.Col = C_BpCd		
			.ColHidden = True
			.Col = C_BpNm	
			.ColHidden = True
			.Col = C_WcCd		
			.ColHidden = False
			.Col = C_WcNm	
			.ColHidden = False
		ElseIf frm1.cboInspClassCd.Value = "F" Then
			.Col = C_BpCd		
			.ColHidden = True
			.Col = C_BpNm	
			.ColHidden = True
			.Col = C_WcCd		
			.ColHidden = True
			.Col = C_WcNm	
			.ColHidden = True
		ElseIf frm1.cboInspClassCd.Value = "S" Then
			.Col = C_BpCd		
			.ColHidden =False
			.Col = C_BpNm	
			.ColHidden = False
			.Col = C_WcCd		
			.ColHidden = True
			.Col = C_WcNm	
			.ColHidden = True
			ggoSpread.SSSetEdit C_BpCd, "거래처코드", 10, 0, -1, 10
			ggoSpread.SSSetEdit C_BpNm, "거래처명", 20, 0, -1, 40
		End If
    End With
    '-----------------------
    'Query function call area
    '-----------------------
	If DbQuery = False then Exit Function		'☜: Query db data
       
    FncQuery = True
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    	Dim IntRetCD 
    
	FncNew = False                                                          '⊙: Processing is NG
	
	Err.Clear                                                               '☜: Protect system from crashing
	'On Error Resume Next                                                    '☜: Protect system from crashing
	ggoSpread.Source = frm1.vspdData
	'-----------------------
	'Check previous data area
	'-----------------------
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then	Exit Function
	
	End If
	
	'-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------
	
	Call ggoOper.ClearField(Document, "A")
	Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
	Call InitVariables                                                      '⊙: Initializes local global variables
	Call SetDefaultVal
	
	If frm1.txtPlantCd.value = "" Then
	   frm1.txtPlantCd.value = UCase(Parent.gPlant)
	   frm1.txtPlantNm.value = Parent.gPlantNm
	End IF
	
	frm1.txtPlantCd.focus
	FncNew = True
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	FncCopy = True
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
    	FncCancel = True
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow() 
	FncInsertRow = True
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
	FncDeleteRow = True
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	Call Parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    On Error Resume Next                                                    						'☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    On Error Resume Next                                                    						'☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)					'☜: 화면 유형 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind()
	Call parent.FncFind(Parent.C_MULTI, False)     
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit    

    iColumnLimit  =  C_InspStatusFlag    

    ACol = Frm1.vspdData.ActiveCol
    ARow = Frm1.vspdData.ActiveRow

    If ACol > iColumnLimit Then
       frm1.vspdData.col = iColumnLimit
       frm1.vspdData.row = 0
       iRet = DisplayMsgBox("900030", "X", Trim(frm1.vspdData.text), "X")
       Exit Function
    End If 
    
    Frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_NONE
    ggoSpread.Source = Frm1.vspdData
    ggoSpread.SSSetSplit(ACol)
    Frm1.vspdData.Col = ACol
    Frm1.vspdData.Row = ARow
    Frm1.vspdData.Action = Parent.SS_ACTION_ACTIVE_CELL 
    Frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_BOTH
End Function

'===============================================================================
' Function Name : FncExit
' Function Desc : This function is related to Excel
'===============================================================================
Function FncExit()
	FncExit = True
End Function

'===============================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'===============================================================================
Function DbQuery() 
	Dim strVal
	
	DbQuery = False
	
	Err.Clear                                                               					'☜: Protect system from crashing
	Call LayerShowHide(1)
		
	With frm1	
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001 _
									& "&txtPlantCd=" & .hPlantCd.value _
									& "&cboInspClassCd=" & .hInspClassCd.value _
									& "&txtInspReqNo=" & .hInspReqNo.value _
									& "&txtItemCd=" & .hItemCd.value _
									& "&txtBpCd=" & .hBpCd.value _
									& "&txtCustCd=" & .hCustCd.value _
									& "&txtWcCd=" & .hWcCd.value _
									& "&txtInspReqDtFr=" & .hInspReqDtFr.Value _
									& "&txtInspReqDtTo=" & .hInspReqDtTo.Value _
									& "&cboStatusFlag=" & .hStatusFlag.Value _
									& "&lgStrPrevKey=" & lgStrPrevKey _
									& "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001 _
									& "&txtPlantCd=" & Trim(.txtPlantCd.Value) _
									& "&cboInspClassCd=" & Trim(.cboInspClassCd.value) _
									& "&txtInspReqNo=" & Trim(.txtInspReqNo.value) _
									& "&txtItemCd=" & Trim(.txtItemCd.value) _
									& "&txtBpCd=" & Trim(.txtBpCd.value) _
									& "&txtCustCd=" & Trim(.txtCustCd.value) _
									& "&txtWcCd=" & Trim(.txtWcCd.value) _
									& "&txtInspReqDtFr=" & Trim(.txtInspReqDtFr.text) _
									& "&txtInspReqDtTo=" & Trim(.txtInspReqDtTo.text) _
									& "&cboStatusFlag=" & .cboStatusFlag.Value _
									& "&lgStrPrevKey=" & lgStrPrevKey _
									& "&txtMaxRows=" & .vspdData.MaxRows
		End If
		
		Call RunMyBizASP(MyBizASP, strVal)							'☜: 비지니스 ASP 를 가동 
		
		DbQuery = True                                                          					'⊙: Processing is NG
	End With    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()									'☆: 조회 성공후 실행로직 
	'-----------------------
	'Reset variables area
	'-----------------------
	lgIntFlgMode = Parent.OPMD_UMODE			'⊙: Indicates that current mode is Update mode
	Call SetToolbar("11000000000111")										'⊙: 버튼 툴바 제어								
	lgBlnFlgChgValue = False
	
	Call ggoOper.LockField(Document, "Q")		'⊙: This function lock the suitable field
	
	With frm1
		Select Case .cboInspClassCd.Value
			Case "R"
				.txtWcCd.Value = ""
				.txtWcNm.Value = ""
				.txtCustCd.Value = ""
				.txtCustNm.Value = ""
				ProtectTag(.txtWcCd)		
				ProtectTag(.txtCustCd)		
				ReleaseTag(.txtBpCd)
			Case "P"
				.txtBpCd.Value = ""
				.txtBpNm.Value = ""
				.txtCustCd.Value = ""
				.txtCustNm.Value = ""
				ProtectTag(.txtBpCd)		
				ProtectTag(.txtCustCd)
				ReleaseTag(.txtWcCd)		
			Case "F"
				.txtBpCd.Value = ""
				.txtBpNm.Value = ""
				.txtCustCd.Value = ""
				.txtCustNm.Value = ""
				.txtWcCd.Value = ""
				.txtWcNm.Value = ""
				ProtectTag(.txtBpCd)		
				ProtectTag(.txtCustCd)
				ProtectTag(.txtWcCd)		
			
			Case "S"
				.txtBpCd.Value = ""
				.txtBpNm.Value = ""
				.txtWcCd.Value = ""
				.txtWcNm.Value = ""
				ProtectTag(.txtBpCd)		
				ProtectTag(.txtWcCd)
				ReleaseTag(.txtCustCd)		
		
		End Select 
	End With
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
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
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif"><IMG SRC="../../../CShared/image/table/seltab_up_left.gif" WIDTH="9" HEIGHT="23"></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>검사의뢰 조회</font></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="right"><IMG SRC="../../../CShared/image/table/seltab_up_right.gif" WIDTH="10" HEIGHT="23"></TD>
						    	</TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD  WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%>></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>공장</TD>
        									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE="10" MAXLENGTH="4" ALT="공장" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
										<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE="20" MAXLENGTH=40 tag="14" ></TD>								
        									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>검사분류</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboInspClassCd" ALT="검사분류" STYLE="WIDTH: 150px" tag="12"></SELECT></TD>
									<TD CLASS="TD5" NOWRAP>검사진행현황</TD>
									<TD CLASS="TD6" NOWRAP><SELECT Name="cboStatusFlag" ALT="검사진행현황" STYLE="WIDTH: 100px" tag="12"></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>검사의뢰번호</TD>
        									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtInspReqNo" SIZE=20 MAXLENGTH=18 ALT="검사의뢰번호" tag="11XXXU"></TD>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=20 ALT="품목" tag="11XXXU"><IMG src="../../../CShared/image/btnPopup.gif" name=btnItemCd align=top  TYPE="BUTTON" width=16 height=20 onclick="vbscript:OpenItem()">
															<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>공급처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 ALT="공급처" tag="11XXXU"><IMG align=top height=20 name=btnBpCd onclick="vbscript:OpenBp()" src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">
															<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>
									<TD CLASS="TD5" NOWRAP>거래처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtCustCd" SIZE=10 MAXLENGTH=10 ALT="거래처" tag="11XXXU"><IMG align=top height=20 name=btnCustCd onclick="vbscript:OpenCust()" src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">
															<INPUT TYPE=TEXT NAME="txtCustNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>작업장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtWcCd" SIZE=10 MAXLENGTH=7 ALT="작업장" tag="11XXXU"><IMG align=top height=20 name=btnWcCd onclick="vbscript:OpenWc()" src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">
										<INPUT TYPE=TEXT NAME="txtWcNm" SIZE=20 MAXLENGTH=40 tag="14" ></TD>
									<TD CLASS="TD5" NOWRAP>검사의뢰일</TD>
									<TD CLASS="TD6">
										<script language =javascript src='./js/q2511ma1_fpDateTime5_txtInspReqDtFr.js'></script>&nbsp;~&nbsp;
										<script language =javascript src='./js/q2511ma1_fpDateTime6_txtInspReqDtTo.js'></script>
									</TD>
								</TR>
								
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=*  WIDTH=100% VALIGN=TOP>						
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD>
									<script language =javascript src='./js/q2511ma1_I780354118_vspdData.js'></script>
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
	<TR HEIGHT=20>
      		<TD WIDTH=100%>
      			<TABLE <%=LR_SPACE_TYPE_30%>>
        				<TR>        				
        					<TD WIDTH=10>&nbsp;</TD>			
        					<TD WIDTH=* ALIGN=RIGHT>
        					<A href="vbscript:LoadInspection">검사등록</A>
        					</TD>
					<TD WIDTH=10>&nbsp;</TD>	
       				</TR>
      			</TABLE>
      		</TD>
   	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="HIDDEN" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hInspClassCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hInspReqNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hBpCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hCustCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hWcCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hInspReqDtFr" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hInspReqDtTo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hStatusFlag" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>


