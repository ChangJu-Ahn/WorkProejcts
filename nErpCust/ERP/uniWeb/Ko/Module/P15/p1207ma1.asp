
<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : Standard Work
'*  3. Program ID           : p1207ma1
'*  4. Program Name         : Standard Manufacturing Instruction Management
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/03/22
'*  8. Modified date(Last)  : 2002/12/03
'*  9. Modifier (First)     : Hong Chang Ho
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
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_LOOKUP_ID	= "p1208mb0.asp"								' Lookup ASP
Const BIZ_PGM_QRY_ID    = "p1207mb1.asp"								'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID   = "p1207mb2.asp"								'☆: 비지니스 로직 ASP명 

Dim C_Seq
Dim C_InstrCd
Dim C_InstrPopUp
Dim C_InstrNm
Dim C_ValidFromDt
Dim C_ValidToDt

<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim lgBlnBtnClick           ' Variable is for Dirty flag
Dim lgBlnFlgStdChgValue     ' Variable is for Dirty flag
Dim lgBlnFlgSaveValue
Dim lgBlnFlgLookupValue
Dim lgBlnMqryMode
Dim IsOpenPop

Dim StartDate

StartDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_Seq			= 1
	C_InstrCd		= 2
	C_InstrPopUp	= 3
	C_InstrNm		= 4
	C_ValidFromDt	= 5
	C_ValidToDt		= 6
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
	lgBlnFlgStdChgValue = False					' Variable is for Dirty flag
	lgBlnFlgSaveValue = False
	lgBlnFlgLookupValue = False					'
	lgBlnBtnClick = False
	lgBlnMqryMode = False
    
    lgIntGrpCount = 100                         'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1                                       '⊙: initializes sort direction
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtStdDt.Text = StartDate
	frm1.txtValidFromDt.Text = StartDate
	frm1.txtValidToDt.Text =  UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")
	
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

		.MaxCols = C_ValidToDt + 1												'☜: 최대 Columns의 항상 1개 증가시킴 
		.MaxRows = 0

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit		C_Seq, "작업순서", 8,,,3,2
		ggoSpread.SSSetEdit		C_InstrCd, "단위작업", 18,,,10,2
		ggoSpread.SSSetButton	C_InstrPopUp
		ggoSpread.SSSetEdit		C_InstrNm, "단위작업내역", 64
		ggoSpread.SSSetDate		C_ValidFromDt, "유효시작일", 11, 2, parent.gDateFormat
		ggoSpread.SSSetDate		C_ValidToDt, "유효종료일", 11, 2, parent.gDateFormat
    
		Call ggoSpread.MakePairsColumn(C_InstrCd, C_InstrPopUp)
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		ggoSpread.SSSetSplit2(3)
    
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
		ggoSpread.SpreadLock	C_Seq, -1, C_Seq
		ggoSpread.SpreadLock	C_InstrNm, -1, C_InstrNm
		ggoSpread.spreadLock	C_ValidFromDt, -1, C_ValidFromDt
		ggoSpread.spreadLock	C_ValidToDt, -1, C_ValidToDt
    
		ggoSpread.SSSetRequired	C_InstrCd, -1
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
		ggoSpread.SSSetRequired	C_Seq,			pvStartRow, pvEndRow
		ggoSpread.SSSetRequired	C_InstrCd,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_InstrNm,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_ValidFromDt,	pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_ValidToDt,	pvStartRow, pvEndRow
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
			C_Seq			= iCurColumnPos(1)
			C_InstrCd		= iCurColumnPos(2)
			C_InstrPopUp	= iCurColumnPos(3)
			C_InstrNm		= iCurColumnPos(4)
			C_ValidFromDt	= iCurColumnPos(5)
			C_ValidToDt		= iCurColumnPos(6)
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

'------------------------------------------  OpenInstrPopUp()  -------------------------------------------------
'	Name : OpenInstrPopUp()
'	Description : OpenInstrPopup
'--------------------------------------------------------------------------------------------------------- 
Function OpenInstrPopUp(ByVal str)
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "단위작업팝업"	
	arrParam(1) = "P_MFG_INSTRUCTION_DETAIL"				
	arrParam(2) = str
	arrParam(3) = ""
	arrParam(4) = "VALID_START_DT <=  " & FilterVar(UNIConvDate(frm1.txtStdDt.Text), "''", "S") & " AND " & _ 
				  "VALID_END_DT >=  " & FilterVar(UNIConvDate(frm1.txtStdDt.Text), "''", "S") & ""
	arrParam(5) = "단위작업"			
	
    arrField(0) = "MFG_INSTRUCTION_DTL_CD"	
    arrField(1) = "MFG_INSTRUCTION_DTL_DESC"	
    arrField(2) = "DD" & parent.gColSep & "CONVERT(VARCHAR(40),VALID_START_DT)"
    arrField(3) = "DD" & parent.gColSep & "CONVERT(VARCHAR(40),VALID_END_DT)"
    
    arrHeader(0) = "단위작업"		
    arrHeader(1) = "단위작업내역"		
    arrHeader(2) = "유효시작일"		
    arrHeader(3) = "유효종료일"		
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetInstr(arrRet)
	End If	
	
	Call SetActiveCell(frm1.vspdData,C_InstrCd,frm1.vspdData.ActiveRow,"M","X","X")
	Set gActiveElement = document.activeElement
	
End Function

'------------------------------------------  OpenStdInstr()  -------------------------------------------------
'	Name : OpenStdInstrPopup()
'	Description : StdInstrPopup
'--------------------------------------------------------------------------------------------------------- 
Function OpenStdInstrPopUp(ByVal str)
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtStdInstrCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "표준작업지시팝업"	
	arrParam(1) = "P_MFG_INSTRUCTION_HEADER"
	arrParam(2) = str
	arrParam(3) = ""
	arrParam(4) = "VALID_FROM_DT <=  " & FilterVar(UNIConvDate(frm1.txtStdDt.Text), "''", "S") & " AND " & _ 
				  "VALID_TO_DT >=  " & FilterVar(UNIConvDate(frm1.txtStdDt.Text), "''", "S") & ""
	arrParam(5) = "표준작업지시"

    arrField(0) = "MFG_INSTRUCTION_CD"
    arrField(1) = "MFG_INSTRUCTION_NM"	
    arrField(2) = "DD" & parent.gColSep & "CONVERT(VARCHAR(40),VALID_FROM_DT)"
    arrField(3) = "DD" & parent.gColSep & "CONVERT(VARCHAR(40),VALID_TO_DT)"
        
    arrHeader(0) = "표준작업지시"		
    arrHeader(1) = "표준작업지시명"		
    arrHeader(2) = "유효시작일"		
    arrHeader(3) = "유효종료일"		
        
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetStdInstr(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtStdInstrCd.focus
	
End Function

'------------------------------------------  SetItemInfo()  --------------------------------------------------
'	Name : SetInstr()
'	Description : 단위작업내역에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetInstr(Byval arrRet)
	With frm1
		.vspdData.Col = C_InstrCd
		.vspdData.Text = arrRet(0)
		.vspdData.Col = C_InstrNm
		.vspdData.Text = arrRet(1)
		.vspdData.Col = C_ValidFromDt 
		.vspdData.Text = arrRet(2)
		.vspdData.Col = C_ValidToDt 
		.vspdData.Text = arrRet(3)
		
		Call vspdData_Change(.vspdData.Col, .vspdData.Row)		' 변경이 일어났다고 알려줌 
	End With
End Function

'------------------------------------------  SetStdInstr()  --------------------------------------------------
'	Name : SetStdInstr()
'	Description : StdInstr PopUp에서 Standard Instruction Code setting
'--------------------------------------------------------------------------------------------------------- 
Function SetStdInstr(byval arrRet)
	frm1.txtStdInstrCd.Value    = arrRet(0)		
	frm1.txtStdInstrNm.Value    = arrRet(1)		
End Function

'==========================================================================================
'   Event Name : LookUpInstr
'==========================================================================================
Function LookUpInstr(Byval StrInstrCd, Byval Row)
    
	Dim strVal
	
	If lgBlnBtnClick = True Then Exit Function
    Call LayerShowHide(1)
    
    strVal = BIZ_PGM_LOOKUP_ID & "?txtMode=" & parent.UID_M0001			'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtWICd=" & Trim(strInstrCd)				'☆: 조회 조건 데이타 
    strVal = strVal & "&txtStdDt=" & Trim(frm1.txtStdDt.Text)        '☜: 조회 조건 데이타 
    strVal = strVal & "&txtRow=" & Row								'☜: 조회 조건 데이타 
    Call RunMyBizASP(MyBizASP, strVal)								'☜: 비지니스 ASP 를 가동 
	
End Function

Function LookUpWIFail(ByRef Row)
    With frm1.vSpdData
		.Row = Row
		.Col = C_InstrCd
		.Text = ""
		.Col = C_InstrNm
		.Text = ""
		.Col = C_ValidFromDt
		.Text = ""
		.Col = C_ValidToDt
		.Text = ""
		.Focus
		.Row = Row
		.Col = C_InstrCd
		.Action = 0
	End With
	If lgBlnFlgSaveValue =True	Then
		lgBlnFlgSaveValue = False
	End If
End Function

Function LookUpWISuccess(ByRef strInstrCd, ByRef strInstrNm, ByRef strValidFromDt,ByRef strValidToDt, ByRef Row ) 
	With frm1.vspdData	
	
		.Row = Row
		.Col = C_InstrCd
		.Text = UCase(strInstrCd)
		.Col = C_InstrNm
		.Text = UCase(strInstrNm)
		.Col = C_ValidFromDt
		.Text = strValidFromDt
		.Col = C_ValidToDt
		.Text = strValidToDt
	End With
	
	If lgBlnFlgSaveValue =True	Then
		lgBlnFlgSaveValue = False
		lgBlnFlgLookupValue = True
		Call MainSave()
	End If
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet
    '----------  Coding part  -------------------------------------------------------------

    Call SetToolbar("11101101001011")										'⊙: 버튼 툴바 제어 
    Call SetDefaultVal
    Call InitVariables
    
	frm1.txtStdInstrCd.Focus 
	Set gActiveElement = document.activeElement 
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

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
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
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

	Dim strSeq, strInstrCd
	Dim i

	With frm1.vspdData
		Select Case Col

		    Case C_Seq
				.Row = Row
				.Col = C_Seq
				strSeq = .Text
				If strSeq = "" Then Exit Sub
				
				For i = 1 To .MaxRows
					If i <> Row Then
						.Row = i
						.Col = C_Seq
						If UCase(Trim(.Text)) = UCase(Trim(strSeq)) Then
							Call DisplayMsgBox("181416", "X", UCase(Trim(strSeq)), "X")
							.Row = Row
							.Text = ""
							Exit Sub
						End If
					End If						
				Next
				
		    Case C_InstrCd
				.Row = Row
				.Col = C_InstrCd
				strInstrCd = .Text
				
				If .Text <> "" Then	
					Call LookUpInstr(strInstrCd, Row)	
				End If
		End Select

	End With
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	
	lgBlnBtnClick = True
	'----------  Coding part  -------------------------------------------------------------   
	With frm1.vspdData 
	
		ggoSpread.Source = frm1.vspdData

		If Row > 0 And Col = C_InstrPopUp Then
		    .Col = C_InstrCd
		    .Row = Row
		    
		    Call OpenInstrPopUp(.Text)
		    
		    Call SetActiveCell(frm1.vspdData,C_InstrCd,Row,"M","X","X")
			Set gActiveElement = document.activeElement
		    
		End If
    End With
	lgBlnBtnClick = False
End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================

Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row >= NewRow Then
        Exit Sub
    End If
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
'   Event Name : txtStdDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtStdDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtStdDt.Action = 7 
		Call SetFocusToDocument("M")
		Frm1.txtStdDt.Focus
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtValidDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtStdDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		lgBlnMqryMode = True
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtStdDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtStdDt_Change() 
'	lgBlnFlgChgValue = True 
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
	lgBlnFlgStdChgValue = True 
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
	lgBlnFlgStdChgValue = True 
End Sub  

Sub txtStdInstrNm1_OnChange()
	lgBlnFlgChgValue = True 
	lgBlnFlgStdChgValue = True 
End Sub  

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
    Dim IntRetCD 

	If lgBlnMqryMode = False Then
		If CheckRunningBizProcess = True Then
			Exit Function
		End If
	End If    

	lgBlnMqryMode = False

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")				'⊙: "Will you destroy previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    If frm1.txtStdInstrCd.value = "" Then
		frm1.txtStdInstrNm.value = ""
	End If
    
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call ggoSpread.ClearSpreadData
'    Call SetDefaultVal															'⊙: Initializes local global variables
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
    frm1.txtStdInstrCd1.value = ""
    
    Call ggoOper.ClearField(Document, "A")											'⊙: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
    Call SetDefaultVal
	Call InitVariables																'⊙: Initializes local global variables
	
	Call SetToolbar("11101101001011")										'⊙: 버튼 툴바 제어 
    
    frm1.txtStdInstrCd1.Focus 
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
    
    If CheckRunningBizProcess = True And lgBlnFlgLookupValue = False Then
    	lgBlnFlgSaveValue = True
		Exit Function
    End If

    lgBlnFlgLookupValue = False
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

'   입력한 행이 하나도 없을때	
	If lgIntFlgMode = parent.OPMD_CMODE Then
		If frm1.vspdData.MaxRows = 0 Then
			Call DisplayMsgBox("971008", "X", "작업순서", "X")
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
	If frm1.vspdData.MaxRows < 1 Then Exit Function
		
	frm1.vspdData.focus
	Set gActiveElement = document.activeElement 
	frm1.vspdData.EditMode = True
	frm1.vspdData.ReDraw = False
    ggoSpread.Source = frm1.vspdData
    ggoSpread.CopyRow
    Call SetSpreadColor(frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow)
    
    frm1.vspdData.Col = C_Seq
    frm1.vspdData.Row = frm1.vspdData.ActiveRow
    
    frm1.vspdData.Text = ""
    
    frm1.vspdData.ReDraw = True
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
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
		.vspdData.Focus
		Set gActiveElement = document.activeElement 
		ggoSpread.Source = .vspdData
		.vspdData.EditMode = True
		.vspdData.ReDraw = False
    
		ggoSpread.InsertRow , iIntReqRows
    
		.vspdData.ReDraw = True
    
		Call SetSpreadColor(.vspdData.ActiveRow, .vspdData.ActiveRow + iIntReqRows - 1)
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
    If frm1.vspdData.MaxRows < 1 Then Exit Function
    
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
			strVal = strVal & "&txtStdInstrCd=" & Trim(.hStdInstrCd.value)				'☆: 조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgCurDt=" & UniConvYYYYMMDDToDate(parent.gDateFormat, "1900","01","01")
			strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001							'☜: 
			strVal = strVal & "&txtStdInstrCd=" & Trim(.txtStdInstrCd.value)				'☆: 조회 조건 데이타 
			strVal = strVal & "&txtStdDt=" & UNIConvDate(frm1.txtStdDt.Text)
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgCurDt=" & UniConvYYYYMMDDToDate(parent.gDateFormat, "1900","01","01")
			strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
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
    lgBlnFlgChgValue = False
    lgBlnFlgStdChgValue = False
    
    Call SetToolbar("11111111001111")
    
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If	
    lgIntFlgMode = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
	
End Function


'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================

Function DbSave() 

    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
	Dim TmpBufferVal, TmpBufferDel
	Dim iTotalStrVal, iTotalStrDel
	Dim iValCnt, iDelCnt
	
    DbSave = False                                                          '⊙: Processing is NG
	
	If ValidDateCheck(frm1.txtValidFromDt, frm1.txtValidToDt) = False Then Exit Function      
	     
    LayerShowHide(1)
		
	With frm1
		.txtMode.Value = parent.UID_M0002
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
    
    '-----------------------
    'Data manipulate area
    '-----------------------
    lGrpCnt = 1
    ReDim TmpBufferVal(0) : ReDim TmpBufferDel(0)
    iValCnt = 0 : iDelCnt = 0
    
    '-----------------------
    'Data manipulate area
    '-----------------------
    For lRow = 1 To .vspdData.MaxRows
    
        .vspdData.Row = lRow
        .vspdData.Col = 0
 
        Select Case .vspdData.Text

            Case ggoSpread.InsertFlag												'☜: 신규 
				
				strVal = ""
				
				strVal = strVal & "C" & parent.gColSep				'⊙: C=Create, Sheet가 2개 이므로 구별				                
                
		        strVal = strVal & UCase(Trim(.txtStdInstrCd1.Value)) & parent.gColSep

                .vspdData.Col = C_Seq			
		        strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		        .vspdData.Col = C_InstrCd			
				strVal = strVal & UCase(Trim(.vspdData.Text)) & parent.gRowSep					'⊙: 마지막 데이타는 Row 분리기호를 넣는다		        
                
                ReDim Preserve TmpBufferVal(iValCnt)
                
                TmpBufferVal(iValCnt) = strVal
                
                iValCnt = iValCnt + 1
                
                lGrpCnt = lGrpCnt + 1

            Case ggoSpread.UpdateFlag												'☜: 수정 
            
				strVal = ""
				
				strVal = strVal & "U" & parent.gColSep				'⊙: U=Update, Sheet가 2개 이므로 구별				                

		        strVal = strVal & UCase(Trim(.txtStdInstrCd1.Value)) & parent.gColSep
                
                .vspdData.Col = C_Seq			
				strVal = strVal & Trim(.vspdData.Text) &parent.gColSep						'⊙: 마지막 데이타는 Row 분리기호를 넣는다		        

		        .vspdData.Col = C_InstrCd			
				strVal = strVal & UCase(Trim(.vspdData.Text)) & parent.gRowSep					'⊙: 마지막 데이타는 Row 분리기호를 넣는다		        
				
				ReDim Preserve TmpBufferVal(iValCnt)
                
                TmpBufferVal(iValCnt) = strVal
                
                iValCnt = iValCnt + 1
				
                lGrpCnt = lGrpCnt + 1

            Case ggoSpread.DeleteFlag												'☜: 삭제 
            
				strDel = ""
				
				strDel = strDel & "D" & parent.gColSep				'⊙: D=Delete
				
		        strDel = strDel & UCase(Trim(.txtStdInstrCd1.Value)) & parent.gColSep
		        
                .vspdData.Col = C_Seq	'10
                strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep					'⊙: 마지막 데이타는 Row 분리기호를 넣는다 
                
                ReDim Preserve TmpBufferDel(iDelCnt)
                
                TmpBufferDel(iDelCnt) = strDel
                
                iDelCnt = iDelCnt + 1
                
                lGrpCnt = lGrpCnt + 1

	    End Select
                
    Next

	If lgIntFlgMode = parent.OPMD_CMODE Then
		frm1.txtStdMode.Value = "C"
	ElseIf lgBlnFlgStdChgValue = True Then
		frm1.txtStdMode.Value = "U"
	Else
		frm1.txtStdMode.Value = "N"
	End If
	
	iTotalStrDel = Join(TmpBufferDel, "")
	iTotalStrVal = Join(TmpBufferVal, "")
		
	.txtMaxRows.value = lGrpCnt-1
	.txtSpread.value = iTotalStrDel & iTotalStrVal

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										'☜: 비지니스 ASP 를 가동 
	
	End With
	
    DbSave = True                                                           '⊙: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
	frm1.txtStdInstrCd.value = frm1.txtStdInstrCd1.value 
	
	Call InitVariables
	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.MaxRows = 0

	lgBlnMqryMode = True
	Call MainQuery()
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete()
	Dim strVal0

    strVal0 = ""

	DbDelete = False														'⊙: Processing is NG
	
	LayerShowHide(1)

	frm1.txtStdMode.Value = "D"
	
	With frm1
		.txtMode.Value = parent.UID_M0003
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
		.txtMaxRows.value = 0
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)									'☜: 비지니스 ASP 를 가동 
    
    DbDelete = True                                                         '⊙: Processing is NG 
End Function

Function DbDeleteOk()
	Call InitVariables
	Call FncNew()
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>표준작업지시등록</font></td>
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
									<TD CLASS=TD5 NOWRAP>표준작업지시</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtStdInstrCd" SIZE=15 MAXLENGTH=6 tag="12XXXU" ALT = "표준작업지시"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnStdInstrCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenStdInstrPopUp frm1.txtStdInstrCd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtStdInstrNm" SIZE=50 MAXLENGTH=40 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>기준일</TD>
									<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/p1207ma1_I416225951_txtStdDt.js'></script>
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
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>표준작업지시</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtStdInstrCd1" SIZE=15 MAXLENGTH=6 tag="23XXXU" ALT="표준작업지시">&nbsp;<INPUT TYPE=TEXT NAME="txtStdInstrNm1" SIZE=40 MAXLENGTH=40 tag="22X1" ALT="표준작업지시명" ></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>유효기간</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p1207ma1_I648467004_txtValidFromDt.js'></script>
									&nbsp;~&nbsp;
									<script language =javascript src='./js/p1207ma1_I155155712_txtValidToDt.js'></script>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT="100%" COLSPAN = 2>
								<script language =javascript src='./js/p1207ma1_vspdData_vspdData.js'></script>
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
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtStdMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hStdInstrCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
