<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2115MA1
'*  4. Program Name         : 부적합처리 등록 
'*  5. Program Desc         : Quality Configuration
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

Const BIZ_PGM_QRY_ID = "Q2115MB1.asp"										 '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID = "Q2115MB2.asp"										 '☆: 비지니스 로직 ASP명 

Const BIZ_PGM_JUMP1_ID = "Q2111MA1"
Const BIZ_PGM_JUMP2_ID = "Q2116MA1"
Const BIZ_PGM_JUMP3_ID = "Q2117MA1"

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim C_DispositionCd '= 1
Dim C_DispositionPopup '= 2
Dim C_DispositionNm '= 3
Dim C_Qty '= 4
Dim C_Remark '= 5

Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim strInspClass

Dim IsOpenPop          

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                                               	'⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                	'⊙: Indicates that no value changed
    lgIntGrpCount = 0                                                     	  	'⊙: Initializes Group View Size
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False						'☆: 사용자 변수 초기화 
    lgStrPrevKey1 = ""
    lgStrPrevKey2 = ""
    
    '###검사분류별 변경부분 Start###
    strInspClass = "R"
	'###검사분류별 변경부분 End###
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE","MA") %>
End Sub

 '==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
	End If
		
	If ReadCookie("txtPlantCd") <> "" Then
		frm1.txtPlantCd.Value = ReadCookie("txtPlantCd")
	End If
	
	If ReadCookie("txtPlantNm") <> "" Then
		frm1.txtPlantNm.Value = ReadCookie("txtPlantNm")
	End If
		
	If ReadCookie("txtInspReqNo") <> "" Then
		frm1.txtInspReqNo.Value = ReadCookie("txtInspReqNo")
	End If
	
	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtInspReqNo", ""	
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021225", , Parent.gAllowDragDropSpread
		
		.ReDraw = false
   		
   		.MaxCols = C_Remark + 1
		.MaxRows = 0
		
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit C_DispositionCd, "처리코드", 15, 0, -1, 2, 2
		ggoSpread.SSSetButton C_DispositionPopup
		ggoSpread.SSSetEdit C_DispositionNm, "처리명", 29, 0, -1, 40
    	ggoSpread.SSSetFloat C_Qty, "수량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
    	ggoSpread.SSSetEdit C_Remark, "비고",55, 0, -1, 200
  		
		Call ggoSpread.MakePairsColumn(C_DispositionCd, C_DispositionPopup)
 		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)	    
	    
	    .ReDraw = true	
	    
	    Call SetSpreadLock
	End With
End Sub

'================================== 2.2.5 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
	Call ggoSpread.SpreadLock(frm1.vspdData.MaxCols, -1, frm1.vspdData.MaxCols)
End Sub

'================================== 2.2.7 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	With frm1
		.vspdData.ReDraw = False
		ggoSpread.SSSetRequired C_DispositionCd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_DispositionNm, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_Qty, pvStartRow, pvEndRow
		.vspdData.ReDraw = True
	End With    
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()
	C_DispositionCd = 1
	C_DispositionPopup = 2
	C_DispositionNm = 3
	C_Qty = 4
	C_Remark = 5	
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

			C_DispositionCd = iCurColumnPos(1)
			C_DispositionPopup = iCurColumnPos(2)
			C_DispositionNm = iCurColumnPos(3)
			C_Qty = iCurColumnPos(4)
			C_Remark = iCurColumnPos(5)
 	End Select
End Sub

'------------------------------------------  OpenPlant()  -------------------------------------------------
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
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWIDTH=420px; dialogHEIGHT=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtPlantCd.Focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtPlantCd.Value    = arrRet(0)		
		frm1.txtPlantNm.Value    = arrRet(1)		
		frm1.txtPlantCd.Focus		
	End If	
	
	Set gActiveElement = document.activeElement
	OpenPlant = true	
End Function

'------------------------------------------  OpenInspReqNo()  -------------------------------------------------
'	Name : OpenInspReqNo()
'	Description : InspReqNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenInspReqNo()        
	OpenInspReqNo = false
	Dim arrRet
	Dim Param1, Param2, Param3, Param4, Param5, Param6
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("220705","X","X","X") 		'공장정보가 필요합니다 
		frm1.txtPlantCd.Focus
		Set gActiveElement = document.activeElement
		Exit Function	
	End If
	
	IsOpenPop = True
	
	Param1 = Trim(frm1.txtPlantCd.value)		
	Param2 = Trim(frm1.txtPlantNm.Value)	
	Param3 = Trim(frm1.txtInspReqNo.Value)	
	'###검사분류별 변경부분 Start###	
	Param4 = strInspClass 		'검사분류 
	'###검사분류별 변경부분 End###
	Param5 = ""			'판정 
	Param6 = ""			'검사진행상태 
	
	iCalledAspName = AskPRAspName("Q4111pa1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "Q4111pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1, Param2, Param3, Param4, Param5, Param6), _
		"dialogWidth=820px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	frm1.txtInspReqNo.Focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtInspReqNo.Value    = arrRet(0)		
		frm1.txtInspReqNo.Focus		
	End If	
	
	Set gActiveElement = document.activeElement
	OpenInspReqNo = true
End Function

'------------------------------------------  OpenDisposition()  -------------------------------------------------
'	Name : OpenDisposition()
'	Description :Disposition PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenDisposition(Byval strCode)
	OpenDisposition = false
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "부적합처리팝업"	
	arrParam(1) = "Q_Disposition"				
	arrParam(2) = strCode
	arrParam(3) = ""
	'###검사분류별 변경부분 Start###
	arrParam(4) = "INSP_CLASS_CD= " & FilterVar(strInspClass, "''", "S") & ""		' Where Condition			
	'###검사분류별 변경부분 End###
	arrParam(5) = "부적합처리"			

    arrField(0) = "Disposition_CD"	
    arrField(1) = "Disposition_NM"	
    
    arrHeader(0) = "부적합처리코드"		
    arrHeader(1) = "부적합처리명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWIDTH=420px; dialogHEIGHT=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	Call SetActiveCell(frm1.vspdData,C_DispositionCd,frm1.vspdData.ActiveRow,"M","X","X")
	If arrRet(0) = "" Then
		Exit Function
	Else
		With frm1
			.vspdData.Col = C_DispositionCd
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_DispositionNm
			.vspdData.Text = arrRet(1)
		
			Call vspdData_Change(C_DispositionCd, .vspdData.ActiveRow)		 ' 변경이 읽어났다고 알려줌 
			Call SetActiveCell(.vspdData,C_DispositionCd,.vspdData.ActiveRow,"M","X","X")
		End With
	End If	
	Set gActiveElement = document.activeElement
	OpenDisposition = true
End Function

'=============================================  2.5.1 LoadInspection()  ======================================
'=	Event Name : LoadInspection
'=	Event Desc :
'========================================================================================================
Function LoadInspection()
	Dim intRetCD
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If intRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	With frm1
		WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "txtInspReqNo", Trim(.txtInspReqNo.value)
	End With
	
	PgmJump(BIZ_PGM_JUMP1_ID)
End Function

'=============================================  2.5.2 LoadRejectReport()  ======================================
'=	Event Name : LoadRejectReport
'=	Event Desc :
'========================================================================================================
Function LoadRejectReport()
	Dim intRetCD
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")
		If intRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	With frm1
		WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "txtInspReqNo", Trim(.txtInspReqNo.value)
	End With
	
	PgmJump(BIZ_PGM_JUMP2_ID)
End Function

'=============================================  2.5.3 LoadRelease()  ======================================
'=	Event Name : LoadRelease
'=	Event Desc :
'========================================================================================================
Function LoadRelease()
	Dim intRetCD
	ggoSpread.Source = frm1.vspdData
        If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")
		If intRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	With frm1
		WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "txtInspReqNo", Trim(.txtInspReqNo.value)
	End With
	
	PgmJump(BIZ_PGM_JUMP3_ID)
End Function

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
 	gMouseClickStatus = "SPC"   
    
 	Set gActiveSpdSheet = frm1.vspdData

	Call SetPopupMenuItemInf("1101111111")         '화면별 설정 
    
 	If frm1.vspdData.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData 
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
 			lgSortKey = 1
 		End If
	Else
 		'------ Developer Coding part (Start)
	 	'------ Developer Coding part (End)
 	End If
End Sub

'========================================================================================
' Function Name : vspdData_DblClick
' Function Desc : 그리드 해더 더블클릭시 네임 변경 
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
 	Dim iColumnName
    
 	If Row <= 0 Then
		Exit Sub
 	End If
    
  	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
 	End If
 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)
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
	Call LoadInfTB19029
	Call ggoOper.LockField(Document, "N")
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call InitSpreadSheet
	Call InitVariables
	'----------  Coding part  -------------------------------------------------------------
	Call SetDefaultVal
	Call SetToolBar("11101101000011")
	Set gActiveElement = document.activeElement
	
	If Trim(frm1.txtPlantCd.Value) = "" Then
		frm1.txtPlantCd.focus 
	Else
		frm1.txtInspReqNo.focus 
	End If
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
Sub Form_QueryUnload(Cancel , UnloadMode)
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 And Col = C_DispositionPopup Then
			.Col = C_DispositionCd
			.Row = Row
			Call OpenDisposition(.Text)
		End If
	End With
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	Call CheckMinNumSpread(frm1.vspdData, Col, Row)

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'==========================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
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
	'----------  Coding part  -----------------------------
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop)  Then
		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If 
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery()
	Dim IntRetCD 
	
	FncQuery = False                                                        							'⊙: Processing is NG
	
	Err.Clear                                                            		   					'☜: Protect system from crashing
	ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	'-----------------------
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "2")								'⊙: Clear Contents  Field
	Call InitVariables
	
	'-----------------------
	'Check condition area
	'-----------------------
	If Not chkField(Document, "1") Then
		Exit Function
	End If
	
	Call ggoOper.LockField(Document, "N")								'⊙: This function lock the suitable field
	'-----------------------
	'Query function call area
	'-----------------------
	
	If DbQuery = False then
		Exit Function
	End If											'☜: Query db data
	
	FncQuery = True		
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
	Dim IntRetCD 
	
	FncNew = False                                            					'⊙: Processing is NG
	Err.Clear                            							'☜: Protect system from crashing
	  '-----------------------
	'Check previous data area
	'-----------------------
	ggoSpread.Source = frm1.vspdData
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	  '-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "A")
	Call ggoOper.LockField(Document, "N")                                       		'⊙: Lock  Suitable  Field
	
	Call InitVariables																'⊙: Initializes local global variables
	Call SetDefaultVal
	Call SetToolBar("11100000000011")		'⊙: 버튼 툴바 제어 
	If Trim(frm1.txtPlantCd.Value) = "" Then
		frm1.txtPlantCd.focus 
	Else
		frm1.txtInspReqNo.focus 
	End If    	
	FncNew = True
End Function

'========================================================================================
' Function Name : Fnc
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
	Dim IntRetCD
	
	FncDelete = False									'⊙: Processing is NG
	
	  '-----------------------
	'Precheck area
	'-----------------------
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")  
		Exit Function
	End If

	IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X") 
	If IntRetCD = vbNo Then
		Exit Function
	End If

	'-----------------------
	'Delete function call area
	'-----------------------
	If DbDelete = False Then
		Exit Function
	End If
	
	FncDelete = True        
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	Dim IntRetCD 
	Dim lngSumOfDefectQty
	Dim i
	FncSave = False                                                         					'⊙: Processing is NG
	
	Err.Clear						                                                        '☜: Protect system from crashing
	
	'-----------------------
	'Precheck area
	'-----------------------
	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = False  Then
		IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
		Exit Function
	End If
	
	'-----------------------
	'Check content area
	'-----------------------
	If Trim(frm1.txtPlantCd.Value) = "" Then
    	Call DisplayMsgBox("970021", "X", frm1.txtPlantCd.Alt, "X")
    	Exit Function
    End If
    	
	If Not chkField(Document, "2") Then
    	Exit Function
    End If
    	
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSDefaultCheck = False Then    				'⊙: Check contents area
    	Exit Function
    End If
    	
    With frm1
		lngSumOfDefectQty = 0
		For i = 1 To .vspdData.MaxRows
			.vspdData.Row = i
			.vspdData.Col = 0
			If .vspdData.Text <> ggoSpread.DeleteFlag Then
				.vspdData.Col = C_Qty
				
				If UNICDbl(.vspdData.Text) = 0 Then
					Call DisplayMsgBox("225002", "X", "X", "X")
					.vspdData.Action = 0
					.vspdData.Focus
					Exit Function
				End If
				lngSumOfDefectQty = lngSumOfDefectQty + UNICDbl(.vspdData.Text)
			End If
		Next
	
		If .hDecisionCd.Value = "R" Then
	    	If lngSumOfDefectQty > UNICDbl(.txtLotSize.Text) Then
	    		'부적합 처리수는 Lot크기보다 클 수 없습니다.
	    		Call DisplayMsgBox("223620", "X", "X", "X")
	    		Exit Function
	    	End If
	    Else
	    	If lngSumOfDefectQty > UNICDbl(.txtDefectQty.Text) Then
				'부적합 처리수는 불량수보다 클 수 없습니다.
				Call DisplayMsgBox("223621", "X", "X", "X")
				Exit Function
	    	End If
	    End If
	End With
	
	'-----------------------
	'Save function call area
	'-----------------------
	If DbSave = False then	
		Exit Function
	End If			                                		                '☜: Save db data
	
	FncSave = True                                                        					  '⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	FncCopy = false
	With frm1
		If .vspdData.MaxRows < 1 then
	    	Exit function
    	End if
		.vspdData.ReDraw = False
		ggoSpread.Source = .vspdData	
		ggoSpread.CopyRow
		Call SetSpreadColor(.vspdData.ActiveRow, .vspdData.ActiveRow)
		.vspdData.Row = .vspdData.ActiveRow
		.vspdData.Col = C_DispositionCd
		.vspdData.Text = ""
		.vspdData.Col = C_DispositionNm
		.vspdData.Text = ""
		frm1.vspdData.ReDraw = True                                   					            '☜: Protect system from crashing
	End With

	Call SetActiveCell(frm1.vspdData,C_DispositionCd,frm1.vspdData.ActiveRow,"M","X","X")
	Set gActiveElement = document.ActiveElement		
	FncCopy = true
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	FncCancel = false
	If frm1.vspdData.MaxRows < 1 Then
		Exit Function
	End If

    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
    FncCancel = true
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt)
	Dim IntRetCD
	Dim imRow
	
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
    	.vspdData.ReDraw = False
		.vspdData.focus 
		ggoSpread.Source = .vspdData
		ggoSpread.InsertRow .vspdData.ActiveRow, imRow
		Call SetSpreadColor(.vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1)
		lgBlnFlgChgValue = True
	End With
	
	Call SetActiveCell(frm1.vspdData,C_DispositionCd,frm1.vspdData.ActiveRow,"M","X","X")
	Set gActiveElement = document.ActiveElement		
	FncInsertRow = true
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
	FncDeleteRow = false
	Dim lDelRows
    	Dim iDelRowCnt, i
    
    	With frm1.vspdData
    		If .MaxRows < 1 Then
    			Exit Function
    		End If 
			.focus
    		ggoSpread.Source = frm1.vspdData 
			lgBlnFlgChgValue = True
			lDelRows = ggoSpread.DeleteRow
    	End With
    FncDeleteRow = true
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
	FncPrev = false
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
	FncNext = false                                                    '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
 	Call parent.FncExport(Parent.C_MULTI)		
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
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub

'========================================================================================
' Function Name : FncExit
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExit()	
	Dim IntRetCD
	
	FncExit = False
	
	ggoSpread.Source = frm1.vspdData
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True  Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	FncExit = True
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
	Dim strVal	
	Err.Clear                                                               					'☜: Protect system from crashing
	
	DbDelete = False									'⊙: Processing is NG
	
	strVal = BIZ_PGM_DEL_ID & "?txtInspReqNo=" & Trim(frm1.txtInspReqNo.value)				'☆: 삭제 조건 데이타 
	
	Call RunMyBizASP(MyBizASP, strVal)				
	
	DbDelete = True
End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()
	Call MainNew()
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal
	
	DbQuery = False	
	Err.Clear                                                               					'☜: Protect system from crashing
	Call LayerShowHide(1)		
	With frm1	
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001  	
			strVal = strVal & "&txtPlantCd=" & .hPlantCd.value					'☆: 조회 조건 데이타 
			strVal = strVal & "&txtInspReqNo=" & .hInspReqNo.value
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey					
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001   
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.Value)		 	'☆: 조회 조건 데이타 
			strVal = strVal & "&txtInspReqNo=" & Trim(.txtInspReqNo.value)
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If
		
		Call RunMyBizASP(MyBizASP, strVal)							'☜: 비지니스 ASP 를 가동 
		
	End With
	DbQuery = True                                                          					'⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryOkOPEN

' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()
	lgIntFlgMode = Parent.OPMD_UMODE									'⊙: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False
    
    'Release완료인 경우에 삭제/저장/행추가/행삭제/취소/복사 Disable
	If frm1.hStatusFlag.Value = "R" Then
		Call SetToolBar("11100000000111")	
	Else
		Call SetToolBar("11101111001111")
	End If
	
	Call SetSpreadColor(1, frm1.vspdData.MaxRows)
	ggoSpread.SSSetProtected C_DispositionCd, 1, -1
	ggoSpread.SSSetProtected C_DispositionPopup, 1, -1
	ggoSpread.SSSetProtected C_DispositionNm, 1, -1
	
    Call ggoOper.LockField(Document, "Q")
End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================
Function DbSave() 
	Dim lRow        
	Dim lGrpCnt     
	Dim strVal 
	Dim strDel
	Dim strTotalInspDispositionQty
	strTotalInspDispositionQty = 0
	
	Call LayerShowHide(1)
	
	DbSave = False                                                          '⊙: Processing is NG
    
	On Error Resume Next                                                   '☜: Protect system from crashing

	With frm1
		.txtMode.value = Parent.UID_M0002
		.txtUpdtUserId.value = Parent.gUsrID
		.txtInsrtUserId.value = Parent.gUsrID
		.txtFlgMode.value = lgIntFlgMode
		    
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
    
		strVal = ""
    	strDel = ""

    	'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
    		.vspdData.Row = lRow
			.vspdData.Col = 0
			Select Case .vspdData.Text
				Case ggoSpread.InsertFlag					'☜: 신규 
					strVal = strVal & "C" & Parent.gColSep			'☜: C=Create
					.vspdData.Col = C_DispositionCd			'1
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					.vspdData.Col = C_Qty					'2
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					.vspdData.Col = C_Remark				'3
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					strVal = strVal & CStr(lRow) & Parent.gRowSep	'4
					lGrpCnt = lGrpCnt + 1
				Case ggoSpread.UpdateFlag					'☜: 수정 
					strVal = strVal & "U" & Parent.gColSep			'☜: U=Update
					.vspdData.Col = C_DispositionCd			'1
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					.vspdData.Col = C_Qty					'2
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					.vspdData.Col = C_Remark				'3
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					strVal = strVal & CStr(lRow) & Parent.gRowSep	'4
					lGrpCnt = lGrpCnt + 1
				Case ggoSpread.DeleteFlag					'☜: 삭제 
					strDel = strDel & "D" & Parent.gColSep			'☜: D=Delete
					.vspdData.Col = C_DispositionCd			'1
					strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
					strDel = strDel & CStr(lRow) & Parent.gRowSep	'2
					lGrpCnt = lGrpCnt + 1
			End Select
		Next
			
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strDel & strVal
	
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)
	End With
	
	DbSave = True 
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function
'========================================================================================
Function DbSaveOk()
	Call InitVariables
	frm1.vspdData.MaxRows = 0
	Call MainQuery()
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
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="center" CLASS="CLSMTAB"><FONT COLOR=white>수입검사 부적합처리</FONT></TD>
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
	     								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 ALT="공장" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
										<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>								
	     								<TD CLASS="TD5" NOWRAP>검사의뢰번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtInspReqNo" SIZE=20  MAXLENGTH=18 ALT="검사의뢰번호" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnInspReqNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenInspReqNo()"></TD>							
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>판정</TD>
                				<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDecision" SIZE=20 MAXLENGTH=20 ALT="판정" tag="24"></TD>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>품목</TD>
                							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=20 ALT="품목" tag="24">
									<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 MAXLENGTH=20 tag="24" ></TD>
								<TD CLASS="TD5" NOWRAP>공급처</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=4 ALT="공급처" tag="24">
									<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 MAXLENGTH=20 tag="24" ></TD>
							</TR>
			                			<TR>
			                				<TD CLASS="TD5" NOWRAP>로트번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtLotNo" SIZE=15 MAXLENGTH=12 ALT="LOT NO" tag="24">
			                					<INPUT TYPE=TEXT NAME="txtLotSubNo" SIZE=10 MAXLENGTH=5 tag="24" STYLE="Text-Align: Right"></TD>
			                				<TD CLASS="TD5" NOWRAP>로트크기</TD>     
			                				<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/q2115ma1_fpDoubleSingle1_txtLotSize.js'></script>
								</TD>       
							</TR>
			                			<TR>
			                				<TD CLASS="TD5" NOWRAP>검사수</TD>            
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/q2115ma1_fpDoubleSingle2_txtInspQty.js'></script>
								</TD>
								<TD CLASS="TD5" NOWRAP>불량수</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/q2115ma1_fpDoubleSingle3_txtDefectQty.js'></script>
								</TD>
			                			</TR>
							<TR>
								<TD HEIGHT=100% WIDTH=100% colspan=4>
									<script language =javascript src='./js/q2115ma1_I741776657_vspdData.js'></script>
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
        					<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:LoadInspection">수입검사</A>&nbsp;|&nbsp;<A href="vbscript:LoadRejectReport">불합격통지</A>&nbsp;|&nbsp;<A href="vbscript:LoadRelease">Release</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
       				</TR>
      			</TABLE>
      		</TD>
    	</TR>
    	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm"  tabindex=-1 WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noreSIZE framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" tabindex=-1></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtFlgMode" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hInspReqNo" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hPlantCd" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hDecisionCd" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hStatusFlag" TAG="24" tabindex=-1>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

