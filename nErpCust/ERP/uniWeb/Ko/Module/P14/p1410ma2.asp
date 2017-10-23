<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1410ma2.asp
'*  4. Program Name         : Query ECN Info.
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/03/07
'*  8. Modified date(Last)  :  
'*  9. Modifier (First)     : Ryu Sung Won
'* 10. Modifier (Last)      : 
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_QRY_ID = "p1410mb9.asp"

Dim C_EcnNo
Dim C_EcnDesc
Dim C_ReasonCd
Dim C_IssuedBy
Dim C_Status
Dim C_EBomFlg
Dim C_EBomDt
Dim C_MBomFlg
Dim C_MBomDt
Dim C_ValidFromDt
Dim C_ValidToDt
Dim C_InsrtId
Dim C_InsrtDt
Dim C_Remark

<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim lgBlnFlgConChg				'☜: Condition 변경 Flag
Dim lgOldRow
Dim iDBSYSDate

Dim IsOpenPop
'Dim lgStrPrevKey

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_EcnNo			= 1
	C_EcnDesc		= 2	
	C_ReasonCd		= 3
	C_IssuedBy		= 4
	C_Status		= 5
	C_EBomFlg		= 6
	C_EBomDt		= 7
	C_MBomFlg		= 8
	C_MBomDt		= 9
	C_ValidFromDt	= 10
	C_ValidToDt		= 11
	C_InsrtId		= 12
	C_InsrtDt		= 13
	C_Remark		= 14
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'==================================================================================================== 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE			
    lgBlnFlgChgValue = False			
    lgIntGrpCount = 0					

    IsOpenPop = False												
	lgStrPrevKey = ""
	lgSortKey = 1
	lgOldRow = 0
End Sub

'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("Q", "P", "NOCOOKIE", "MA")%>
End Sub

'========================= 2.2.3 InitSpreadSheet() ======================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()    
	
	With frm1.vspdData

		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021125",, parent.gAllowDragDropSpread

		.ReDraw = False

		.MaxCols = C_Remark + 1
		.MaxRows = 0

		Call GetSpreadColumnPos("A")
    
		ggoSpread.SSSetEdit		C_EcnNo,		"설계변경번호", 18
		ggoSpread.SSSetEdit		C_EcnDesc,		"설계변경내용", 30
		ggoSpread.SSSetEdit		C_ReasonCd,		"설계변경근거", 10
		ggoSpread.SSSetEdit		C_IssuedBy,		"설계변경근거명", 14
		ggoSpread.SSSetEdit		C_Status,		"설계변경상태", 12
		ggoSpread.SSSetEdit		C_EBomFlg,		"설계BOM반영여부", 14, 2
		ggoSpread.SSSetEdit		C_EBomDt,		"설계BOM반영일", 14, 2
		ggoSpread.SSSetEdit		C_MBomFlg,		"생산BOM반영여부", 14, 2
		ggoSpread.SSSetEdit		C_MBomDt,		"생산BOM반영일", 14, 2
		ggoSpread.SSSetEdit		C_ValidFromDt,	"시작일", 11, 2
		ggoSpread.SSSetEdit		C_ValidToDt,	"종료일", 11, 2
		ggoSpread.SSSetEdit		C_InsrtId,		"생성자", 13
		ggoSpread.SSSetEdit		C_InsrtDt,		"생성일", 11, 2
		ggoSpread.SSSetEdit		C_Remark,		"비고", 50
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
    
		ggoSpread.SSSetSplit2(1)										'frozen 기능추가 

		.ReDraw = True

		Call SetSpreadLock 

    End With
    
End Sub

'============================== 2.2.4 SetSpreadLock() ===================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'============================ 2.2.5 SetSpreadColor() ====================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
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
    
           	C_EcnNo			= iCurColumnPos(1)
           	C_EcnDesc		= iCurColumnPos(2)
			C_ReasonCd		= iCurColumnPos(3)
			C_IssuedBy		= iCurColumnPos(4)
			C_Status		= iCurColumnPos(5)
			C_EBomFlg		= iCurColumnPos(6)
			C_EBomDt		= iCurColumnPos(7)
			C_MBomFlg		= iCurColumnPos(8)
			C_MBomDt		= iCurColumnPos(9)
			C_ValidFromDt	= iCurColumnPos(10)
			C_ValidToDt		= iCurColumnPos(11)
			C_InsrtId		= iCurColumnPos(12)
			C_InsrtDt		= iCurColumnPos(13)
			C_Remark		= iCurColumnPos(14)
			
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

'========================================  2.2.1 SetDefaultVal()  ==================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'===================================================================================================
Sub SetDefaultVal()
	iDBSYSDate = "<%=GetSvrDate%>"
	frm1.txtValidDt.text = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
End Sub

Sub InitComboBox()

End Sub

'------------------------------------------  OpenECNInfo()  ----------------------------------------------
'	Name : OpenECNInfo()
'	Description : 
'---------------------------------------------------------------------------------------------------------
Function OpenECNInfo()

	Dim arrRet
	Dim arrParam(4), arrField(10)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtECNNo.value)	' ECNNo
	arrParam(1) = ""						' ReasonCd
	arrParam(2) = ""						' Status
	arrParam(3) = ""						' EBomFlg
	arrParam(4) = ""						' MBomFlg

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
		Call SetECNInfo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	Frm1.txtECNNo.Focus
	
End Function

'------------------------------------------  OpenReasonPopup()  ------------------------------------------
'	Name : OpenReasonPopup()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Function OpenReasonPopup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
   
	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If
  
	'---------------------------------------------
	' Parameter Setting
	'--------------------------------------------- 

	IsOpenPop = True

	arrParam(0) = "설계변경번호팝업"					' 팝업 명칭 
	arrParam(1) = "B_MINOR"								' TABLE 명칭 
	arrParam(2) = UCase(Trim(frm1.txtReasonCd.value))	' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = "MAJOR_CD = " & FilterVar("P1402", "''", "S") & ""
	
	arrParam(5) = "설계변경근거"						' TextBox 명칭 
	
    arrField(0) = "MINOR_CD"						' Field명(0)
    arrField(1) = "MINOR_NM"						' Field명(1)
        
    arrHeader(0) = "설계변경근거"					' Header명(0)
    arrHeader(1) = "설계변경근거명"					' Header명(1)
    
	arrRet = window.showModalDialog("../../ComAsp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetReasonInfo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	Frm1.txtReasonCd.Focus	
	
End Function


'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  SetClassCd()  ------------------------------------------------
'	Name : SetClassCd()
'	Description : Class Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetECNInfo(byval arrRet)
	frm1.txtECNNo.Value    = arrRet(0)		
	frm1.txtECNNoDesc.Value	= arrRet(1)
	
	frm1.txtECNNo.focus
	Set gActiveElement = document.activeElement 
End Function

'------------------------------------------  SetReasonInfo()  --------------------------------------------------
'	Name : SetReasonInfo()
'	Description : 
'---------------------------------------------------------------------------------------------------------
Function SetReasonInfo(byval arrRet)
	frm1.txtReasonCd.Value		= arrRet(0)
	frm1.txtReasonDesc.Value	= arrRet(1)	
		
	frm1.txtReasonCd.focus
	Set gActiveElement = document.activeElement
End Function


'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
'on error resume next
err.Clear
    Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)    

    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field

	Call SetDefaultVal
   	Call InitComboBox
    Call InitVariables		
    Call InitSpreadSheet	
	Call SetToolbar("11000000000011")

	frm1.txtEcnNo.focus
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
	Dim IntRetCD
	
	gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
	Call SetPopupMenuItemInf("0000111111")    
	
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
       
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
    End If
    
	If Row <= 0 Or Col < 0 Then
		ggoSpread.Source = frm1.vspdData
		Exit Sub
	End If
	
	frm1.vspdData.Row = Row
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'=======================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 1.실행시간(runtime)에 팝업메뉴를 통해서 동적으로 바꾸자.
'				 2.Mouse로 특정Cell을 선택("SPC")하고 오른쪽 버튼("SPCR")을 누르면 팝업이 보인다.
'				   팝업에서 특정 메뉴 item을 선택("SPCRP") 실제 칼럼을 freeze한다.
'=======================================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
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
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
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

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

End Sub

'=======================================================================================================
'   Event Name : txtValidDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtValidDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call FncQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtValidDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtValidDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtValidDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtValidDt.Focus
    End If
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
        
    FncQuery = False                                                       
    
    Err.Clear                                                              
        
	'-----------------------
    'Erase contents area
    '----------------------- 

    Call ggoOper.ClearField(Document, "2")		
    Call InitVariables							
        
	'-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then								
       Exit Function
    End If
    
	'-----------------------
    'Query function call area
    '----------------------- 
    If DbQuery = False Then
		Exit Function
    End If													
       
    FncQuery = True													

End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
	On Error Resume Next
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
	On Error Resume Next    
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	On Error Resume Next
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	On Error Resume Next	
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	On Error Resume Next	
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow() 
	On Error Resume Next	
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
	On Error Resume Next	
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)                                                   <%'☜: Protect system from crashing%>
    'Call parent.FncExport(parent.C_SINGLE)											
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)     
    'Call parent.FncFind(parent.C_SINGLE, False)                                         '☜:화면 유형, Tab 유무 
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
'Function FncSplitColumn()
'    
'    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
'       Exit Function
'    End If
'
'    ggoSpread.Source = gActiveSpdSheet
'    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
'    
'End Function

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
'Sub FncSplitColumn()
'
'    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
'       Exit Sub
'    End If
'
'    ggoSpread.Source = gActiveSpdSheet
'    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
'
'End Sub

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	
	Dim strcboStatus, strcboEBomFlg, strcboMBomFlg
	
	Err.Clear															

	DbQuery = False														

	LayerShowHide(1)
		
	Dim strVal
	
	If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&txtEcnNo="		& UCase(Trim(frm1.hECNNo.value))
		strVal = strVal & "&txtReasonCd="	& Trim(frm1.hReasonCd.value)
		strVal = strVal & "&txtValidDt="	& frm1.hValidDt.value
		strVal = strVal & "&cboStatus="		& Trim(frm1.hStatus.value)
		strVal = strVal & "&cboEBomFlg="	& Trim(frm1.hEBomFlg.value)
		strVal = strVal & "&cboMBomFlg="	& Trim(frm1.hMBomFlg.value)
		
		strVal = strVal & "&lgIntFlgMode="	& lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey="	& lgStrPrevKey
		strVal = strVal & "&txtMaxRows="	& frm1.vspdData.MaxRows
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&txtEcnNo="		& UCase(Trim(frm1.txtECNNo.value))
		strVal = strVal & "&txtReasonCd="	& Trim(frm1.txtReasonCd.value)
		strVal = strVal & "&txtValidDt="	& frm1.txtValidDt.text

		If frm1.cboStatus1.checked = True then
			strcboStatus = ""
		ElseIf frm1.cboStatus2.checked = True then
			strcboStatus = "1"
		Else			
			strcboStatus = "2"
		End IF
		
		If frm1.cboEBomFlg1.checked = True then
			strcboEBomFlg = ""
		ElseIf frm1.cboEBomFlg2.checked = True then
			strcboEBomFlg = "Y"
		Else			
			strcboEBomFlg = "N"
		End IF
		
		If frm1.cboMBomFlg1.checked = True then
			strcboMBomFlg = ""
		ElseIf frm1.cboMBomFlg2.checked = True then
			strcboMBomFlg = "Y"
		Else			
			strcboMBomFlg = "N"
		End IF

		strVal = strVal & "&cboStatus=" & strcboStatus
		strVal = strVal & "&cboEBomFlg=" & strcboEBomFlg
		strVal = strVal & "&cboMBomFlg=" & strcboMBomFlg

		strVal = strVal & "&lgIntFlgMode="	& lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey="	& lgStrPrevKey
		strVal = strVal & "&txtMaxRows="	& frm1.vspdData.MaxRows
	
	End If  

	Call RunMyBizASP(MyBizASP, strVal)									

	DbQuery = True																					
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()													

 '------ Reset variables area ------
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If	
	lgIntFlgMode = parent.OPMD_UMODE											

	Call ggoOper.LockField(Document, "Q")								
	Call SetToolbar("11000000000111")
	
	frm1.vspddata.Focus
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>설계변경정보조회</font></td>
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
									<TD CLASS=TD5 NOWRAP>설계변경번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtECNNo" SIZE=18 MAXLENGTH=18 tag="11XXXU" ALT="설계변경번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnECNNoPopup" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenECNInfo()">
														<INPUT TYPE=TEXT NAME="txtECNNoDesc" SIZE=18 tag="X4" ALT="설계변경내역"></TD>
									<TD CLASS=TD5 NOWRAP>설계변경근거</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtReasonCd" SIZE=10 MAXLENGTH=2 tag="11XXXU" ALT="설계변경근거"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnReasonPopup" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenReasonPopup()">
														 <INPUT TYPE=TEXT NAME="txtReasonDesc" SIZE=18 tag="X4" ALT="설계변경근거명"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>기준일</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/p1410ma2_I326236659_txtValidDt.js'></script>
									</TD>
									<TD CLASS=TD5 NOWRAP>설계변경상태</TD>
									<TD CLASS=TD6 NOWRAP>
											<INPUT TYPE="RADIO" CLASS="RADIO" NAME="cboStatus" tag="1X" CHECKED ID="cboStatus1" VALUE=""><LABEL FOR="cboStatus1">전체</LABEL>
											<INPUT TYPE="RADIO" CLASS="RADIO" NAME="cboStatus" tag="1X" ID="cboStatus2" VALUE="1"><LABEL FOR="cboStatus2">Active</LABEL>
											<INPUT TYPE="RADIO" CLASS="RADIO" NAME="cboStatus" tag="1X" ID="cboStatus3" VALUE="2"><LABEL FOR="cboStatus3">Inactive</LABEL>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>설계BOM반영여부</TD>
									<TD CLASS=TD6 NOWRAP>
											<INPUT TYPE="RADIO" CLASS="RADIO" NAME="cboEBomFlg" tag="1X" CHECKED ID="cboEBomFlg1" VALUE=""><LABEL FOR="cboEBomFlg1">전체</LABEL>
											<INPUT TYPE="RADIO" CLASS="RADIO" NAME="cboEBomFlg" tag="1X" ID="cboEBomFlg2" VALUE="Y"><LABEL FOR="cboEBomFlg2">예</LABEL>
											<INPUT TYPE="RADIO" CLASS="RADIO" NAME="cboEBomFlg" tag="1X" ID="cboEBomFlg3" VALUE="N"><LABEL FOR="cboEBomFlg3">아니오</LABEL>
								</TD>
									<TD CLASS=TD5 NOWRAP>생산BOM반영여부</TD>
									<TD CLASS=TD6 NOWRAP>
											<INPUT TYPE="RADIO" CLASS="RADIO" NAME="cboMBomFlg" tag="1X" CHECKED ID="cboMBomFlg1" VALUE=""><LABEL FOR="cboMBomFlg1">전체</LABEL>
											<INPUT TYPE="RADIO" CLASS="RADIO" NAME="cboMBomFlg" tag="1X" ID="cboMBomFlg2" VALUE="Y"><LABEL FOR="cboMBomFlg2">예</LABEL>
											<INPUT TYPE="RADIO" CLASS="RADIO" NAME="cboMBomFlg" tag="1X" ID="cboMBomFlg3" VALUE="N"><LABEL FOR="cboMBomFlg3">아니오</LABEL>
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
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT=* WIDTH=100%>
									<script language =javascript src='./js/p1410ma2_vspdData_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hEcnNo" tag="24"><INPUT TYPE=HIDDEN NAME="hReasonCd" tag="24"><INPUT TYPE=HIDDEN NAME="hValidDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hStatus" tag="24"><INPUT TYPE=HIDDEN NAME="hEBomFlg" tag="24"><INPUT TYPE=HIDDEN NAME="hMBomFlg" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
