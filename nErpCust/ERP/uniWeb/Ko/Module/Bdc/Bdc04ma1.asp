<!--
======================================================================================================
*  1. Module Name          : BDC
*  2. Function Name        : 
*  3. Program ID           : 
*  4. Program Name         : 
*  5. Program Desc         :
*  6. Comproxy List        : 
*  7. Modified date(First) : 2005/01/07
*  8. Modified date(Last)  : 2005/01/07
*  9. Modifier (First)     : Kweon, Soon Tae
* 10. Modifier (Last)      : Kweon, Soon Tae
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">        
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit
Const BIZ_PGM_ID		= "BDC04MB1.asp"
Const BIZ_RUN_ID		= "BDC04MB2.asp"
Const BIZ_PGM_JUMP_ID	= "BDC05MA1"

Dim  C_CHECK	 
Dim  C_PROCESS_ID 
Dim  C_PROCESS_NM 
Dim  C_JOB_NO    
Dim  C_JOB_NM	
Dim  C_STATE	 
Dim  C_RESULT	 
Dim  C_TOTAL	 
Dim  C_SUCESS	
Dim  C_FAIL	 
Dim  C_SCHEDULE  
Dim  C_START	
Dim  C_END

Dim IsOpenPop
Dim iColSep, iRowSep
Dim lgOldRow

iColSep = parent.gColSep
iRowSep = parent.gRowSep

<!-- #Include file="../../inc/lgvariables.inc" -->
'=========================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE
    lgIntGrpCount = 0
    lgStrPrevKey = ""
    lgLngCurRows = 0
    lgSortKey = 1    
    lgOldRow = 0
    frm1.txtSpread.value = ""
    frm1.btnRunJob.disabled = True
	frm1.btnCanJob.disabled = True
End Sub

'=========================================================================================================
Sub SetDefaultVal()
	Dim LocSvrDate
	LocSvrDate = "<%=GetSvrDate%>"
	
	frm1.txtTrnsFrDt.text = UniConvDateAToB(UNIDateAdd ("D", -1, LocSvrDate, Parent.gServerDateFormat), Parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtTrnsToDt.text   = UniConvDateAToB(UNIDateAdd ("D", 1, LocSvrDate, Parent.gServerDateFormat), Parent.gServerDateFormat, parent.gDateFormat)
	
	frm1.btnRunJob.disabled = True
	frm1.btnCanJob.disabled = True
	
End Sub

'========================================  2.2.1 SetCookieVal()  ======================================
'	Name : SetCookieVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=================================================================================================== 
Sub SetCookieVal()
   	
	frm1.txtJobID.value	= ReadCookie("txtJobId")

	WriteCookie "txtJobId", ""
		
End Sub

'=========================================================================================================
Sub LoadInfTB19029()
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
    <% Call loadInfTB19029A("I", "*", "NOCOOKIE","QA") %>
End Sub

'=========================================================================================================
Sub InitSpreadSheet()
	
	Call InitSpreadPosVariables()
	Call AppendNumberPlace("6", "5", "0")
	
    With frm1.vspdData
        ggoSpread.Source = frm1.vspdData
        ggoSpread.Spreadinit "V20050202", , Parent.gAllowDragDropSpread
		
        .ReDraw = False
        .MaxCols = C_END + 1
        .MaxRows = 0
        
        Call GetSpreadColumnPos()
		
        ' 인수 의미   
        'CHECK-------------------------------------------
        ' 1:  Index             Required
        ' 2:  Header            Required
        ' 3:  ColWidth          Required
        ' 4:  HAlign            Option(0,1,2)
        ' 5:  Checktext         Option(True, False)
        ' 6:  CheckCenter       Option(True, False)
        ' 7:  Row               Option(default: all Rows)

        'COMBO-------------------------------------------
        ' 1:  Index             Required
        ' 2:  Header            Required
        ' 3:  ColWidth          Required
        ' 4:  HAlign            Option(0,1,2)
        ' 5:  ComboBoxEditable  Option(True, False)
        ' 6:  Row               Option(default: all Rows)

        'EDIT--------------------------------------------
        ' 1:  Index             Required
        ' 2:  Header            Required
        ' 3:  ColWidth          Required
        ' 4:  HAlign            Option(0,1,2)
        ' 5:  Row               Option(default: all Rows)
        ' 6:  Length            Option(Max input character count)
        ' 7:  CharCase          Option(0:LowCase, 1:Default, 2:UpperCase)

        ggoSpread.SSSetCheck C_CHECK,      "", 3,,,True
        ggoSpread.SSSetEdit  C_PROCESS_ID, "업무코드", 8
		ggoSpread.SSSetEdit  C_PROCESS_NM, "업무코드명", 12
        ggoSpread.SSSetEdit  C_JOB_NO,     "작업번호", 15
        ggoSpread.SSSetEdit  C_JOB_NM,     "작업명", 20
        ggoSpread.SSSetEdit  C_STATE,      "상태", 6
        ggoSpread.SSSetEdit  C_RESULT,     "결과", 6
        ggoSpread.SSSetFloat C_TOTAL,		"등록", 6,"6", ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
        ggoSpread.SSSetFloat C_SUCESS,		"성공", 6,"6", ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
        ggoSpread.SSSetFloat C_FAIL,		"실패", 6,"6", ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
        ggoSpread.SSSetEdit  C_SCHEDULE,   "예약", 19
        ggoSpread.SSSetEdit  C_START,      "시작", 19
        ggoSpread.SSSetEdit  C_END,        "종료", 19
        
        ggoSpread.SSSetSplit2(2)
        .ReDraw = True
		
        Call SetSpreadLock()
        Call ggoSpread.MakePairsColumn(C_JOB_NO, C_JOB_NM, "1")
        Call ggoSpread.SSSetColHidden(C_SCHEDULE, C_SCHEDULE, True)
        Call ggoSpread.SSSetColHidden(C_PROCESS_ID, C_PROCESS_ID, True)
        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
    End With
End Sub

'=========================================================================================================
'Fpoint spreadsheet의 특정 영역을 입력불가 Cell로 Lock시킨다. 
'반환값 없다.
Sub SetSpreadLock()
  ggoSpread.SpreadLockWithOddEvenRowColor()  
    With frm1
        .vspdData.ReDraw = False
      ' ggoSpread.SpreadLock C_PROCESS_ID,   -1, C_END, -1
        .vspdData.ReDraw = True
    End With
  
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()	
	
	' Grid 1(vspdData) - Operation 
	C_CHECK			= 1
	C_PROCESS_ID	= 2
	C_PROCESS_NM	= 3
	C_JOB_NO		= 4
	C_JOB_NM		= 5
	C_STATE			= 6
	C_RESULT		= 7
	C_TOTAL			= 8
	C_SUCESS		= 9
	C_FAIL			= 10
	C_SCHEDULE		= 11
	C_START			= 12
	C_END			= 13

End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'==========================================================================================================
Sub GetSpreadColumnPos()
 	Dim iCurColumnPos

 	ggoSpread.Source = frm1.vspdData
 			
 	Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 	
 	C_CHECK			= iCurColumnPos(1)
	C_PROCESS_ID	= iCurColumnPos(2)
	C_PROCESS_NM	= iCurColumnPos(3)
	C_JOB_NO		= iCurColumnPos(4)
	C_JOB_NM		= iCurColumnPos(5)
	C_STATE			= iCurColumnPos(6)
	C_RESULT		= iCurColumnPos(7)
	C_TOTAL			= iCurColumnPos(8)
	C_SUCESS		= iCurColumnPos(9)
	C_FAIL			= iCurColumnPos(10)
	C_SCHEDULE		= iCurColumnPos(11)
	C_START			= iCurColumnPos(12)
	C_END			= iCurColumnPos(13)
	
End Sub				


'=========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
End Sub

'=========================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow

    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> UC_PROTECTED Then
              Frm1.vspdData.Action = 0 
              Exit For
           End If
       Next
    End If
End Sub

'=========================================================================================================
Sub InitComboBox()
    Call SetCombo(frm1.cboJobState, "W", "대기")
    Call SetCombo(frm1.cboJobState, "C", "취소")
    Call SetCombo(frm1.cboJobState, "R", "실행")
    Call SetCombo(frm1.cboJobState, "D", "완료")
End Sub

'=========================================================================================================
Sub InitSpreadComboBox()
    Dim strCboData1
    Dim strCboData2
    Dim IntRetCD

    ggoSpread.Source = frm1.vspdData

    strCboData1 = "W" & vbTab & "C" & vbTab & "R" & vbTab & "D"
    ggoSpread.SetCombo strCboData1, C_STATE
    
    strCboData2 = "S" & vbTab & "F"
    ggoSpread.SetCombo strCboData2, C_RESULT
End Sub

'=======================================================================================================
'   Event Name : txtTrnsFrDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtTrnsFrDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtTrnsFrDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtTrnsFrDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtTrnsToDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================0
Sub txtTrnsToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtTrnsToDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtTrnsToDt.Focus
    End If
End Sub

'------------------------------------------  txtTrnsFrDt_KeyDown ----------------------------------------
'	Name : txtTrnsFrDt_KeyDown
'	Description : Plant Popup에서 Return되는 값 setting
'----------------------------------------------------------------------------------------------------------
Sub txtTrnsFrDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	

'------------------------------------------  txtTrnsToDt_KeyDown ------------------------------------------
'	Name : txtTrnsToDt_KeyDown
'	Description : Plant Popup에서 Return되는 값 setting
'----------------------------------------------------------------------------------------------------------
Sub txtTrnsToDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	

'=========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")
    Call InitSpreadSheet
    Call InitVariables
    Call InitComboBox
    Call InitSpreadComboBox
    Call SetDefaultVal
    Call SetToolbar("11001011000111")
    If parent.ReadCookie("txtJobId") <> "" Then
		Call SetCookieVal
	End If
    frm1.txtProcessID.focus
	Set gActiveElement = document.activeElement
End Sub

'=========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'=========================================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row)
    'ggoSpread.Source = frm1.vspdData
    'ggoSpread.UpdateRow Row

End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
	
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData
    
 	If frm1.vspdData.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey			'Sort in Descending
 			lgSortKey = 1
 		End If
 		
 		lgOldRow = Row
 		
	Else
 		'------ Developer Coding part (Start)
 		If lgOldRow <> Row Then		
			frm1.vspdData.Row = row
		
			lgOldRow = Row
		
		End If		
	 	'------ Developer Coding part (End)
	
 	End If
	
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
'   Event Name : vspdData_DblClick
'   Event Desc : This event is spread sheet data changed
'==========================================================================================

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	If Row <= 0 Then Exit Sub
	
	ggoSpread.Source = frm1.vspdData
	
	With frm1.vspdData
		.Row = Row
		.Col = C_CHECK

		If .Value = "0" Then
			If ButtonDown = 1 Then
				ggoSpread.UpdateRow Row
			Else
				ggoSpread.SSDeleteFlag Row,Row
			End If	
		Else
			If ButtonDown = 1 Then
				ggoSpread.UpdateRow Row
			Else
				ggoSpread.SSDeleteFlag Row,Row
			End If			
		End If
	End With
	
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then			'⊙: 다른 비즈니스 로직이 진행 중이면 더 이상 업무로직ASP를 호출하지 않음 
         Exit Sub
	End If     
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
		    If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub

'=========================================================================================================
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
    Call GetSpreadColumnPos()
End Sub 

Function btnCanJob_onClick()
	
	Dim IntRetCD
	
	Err.Clear                                                 '☜: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData                          '⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then                   '⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001", "x", "x", "x")     '⊙: Display Message(There is no changed data.)
        Exit Function
    End If

	Call OpenCanJob
	
End Function

Function btnRunJob_onClick()
	
	Dim IntRetCD
	
	Err.Clear                                                 '☜: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData                          '⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then                   '⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001", "x", "x", "x")     '⊙: Display Message(There is no changed data.)
        Exit Function
    End If
	
	Call OpenRunJob
End Function

Function rdoCfmAll_OnClick()
	frm1.txtResultCD.value = frm1.rdoCfmAll.value
End Function

Function rdoCfmSuc_OnClick()
	frm1.txtResultCD.value = frm1.rdoCfmYes.value
End Function

Function rdoCfmFail_OnClick()
	frm1.txtResultCD.value = frm1.rdoCfmNo.value
End Function
 
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
    
	Dim LngRow
	 
    ggoSpread.Source = gActiveSpdSheet

    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()  
    
	Call ggoSpread.ReOrderingSpreadData
    
	ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
	
		.ReDraw = False
    
		For LngRow = 1 To  .MaxRows
		
			If GetSpreadText(frm1.vspdData, C_STATE, LngRow, "X", "X") <> "대기" Then
				ggoSpread.SpreadLock C_CHECK,	LngRow, C_CHECK, LngRow
			End If
		
		Next 
	
		.ReDraw = True
	
	End With

End Sub 


'=========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    FncQuery = False
    Err.Clear
	
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "x", "x")
        If IntRetCD = vbNo Then
          Exit Function
        End If
    End If
    
    If Not chkField(Document, "1") Then
		Exit Function
    End If

    Call ggoSpread.ClearSpreadData()
    Call InitVariables
	If ValidDateCheck(frm1.txtTrnsFrDt, frm1.txtTrnsToDt) = False Then Exit Function

    If DbQuery = False Then
       Exit Function
    End If
       
    FncQuery = True
End Function

'=========================================================================================================
Function FncNew() 
End Function

'=========================================================================================================
Function FncDelete() 
End Function

'=========================================================================================================
Function FncSave() 
    Dim IntRetCD
    FncSave = False

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "x", "x", "x")
        Exit Function
    End If

    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then
        Exit Function
    End If

    If DbSave = False Then
        Exit Function
    End If

    FncSave = True
End Function

'=========================================================================================================
Function FncCopy() 
	Dim nActiveRow
    With frm1.vspdData
        If .ActiveRow > 0 Then
            .focus
            .ReDraw = False
            
            ggoSpread.Source = frm1.vspdData 
            ggoSpread.CopyRow
            nActiveRow = frm1.vspdData.ActiveRow
            SetSpreadColor nActiveRow, nActiveRow
    
    		frm1.vspdData.SetText C_MnuID, nActiveRow, ""
            .ReDraw = True
        End If
    End With
End Function

'=========================================================================================================
Function FncCancel()
    ggoSpread.EditUndo
End Function

'=========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    
    On Error Resume Next                                                              
    Err.Clear                                                                     
    
    FncInsertRow = False                                                             

    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
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
        ggoSpread.InsertRow ,imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.ReDraw = True
    End With
    
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 

    '------ Developer Coding part (End )   -------------------------------------------------------------- 

    If Err.number = 0 Then
       FncInsertRow = True                                                              
    End If   
    
    Set gActiveElement = document.ActiveElement   
End Function

'=========================================================================================================
Function FncDeleteRow()
    Dim lDelRows

    With frm1.vspdData 
        .focus
        ggoSpread.Source = frm1.vspdData 
        lDelRows = ggoSpread.DeleteRow
    End With
End Function

'=========================================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'=========================================================================================================
Function FncPrev() 
End Function

'=========================================================================================================
Function FncNext() 
End Function

'=========================================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)
End Function

'=========================================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)
End Function

'=========================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'=========================================================================================================
Function FncExit()
    Dim IntRetCD
    FncExit = False
    
    ggoSpread.Source = frm1.vspdData    
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "x", "x")
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If
    FncExit = True
End Function

'=========================================================================================================
Function DbQuery() 
    Dim strVal    
    Dim IntRetCD

    DbQuery = False

    Call LayerShowHide(1)
    With frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
        strVal = BIZ_PGM_ID & _
                "?txtMode="         & Parent.UID_M0001          & _
                "&lgIntFlgMode="	& lgIntFlgMode				& _
                "&txtProcessID="    & Trim(.hProcessID.value) & _
                "&txtJobID="        & Trim(.hJobID.value)     & _
                "&txtRegisterID="   & Trim(.hRegisterID.value)& _
                "&txtTrnsFrDt="     & Trim(.hTrnsFrDt.value)  & _
                "&txtTrnsToDt="     & Trim(.hTrnsToDt.value)  & _
                "&cboJobState="     & Trim(.hJobState.value)  & _
                "&txtResultCD="     & Trim(.hResultCD.value)  & _
                "&txtMaxRows="      & .vspdData.MaxRows         & _
                "&lgStrPrevKey="    & lgStrPrevKey
        Else
			strVal = BIZ_PGM_ID & _
                "?txtMode="         & Parent.UID_M0001          & _
                "&lgIntFlgMode="	& lgIntFlgMode				& _
                "&txtProcessID="    & Trim(.txtProcessID.value) & _
                "&txtJobID="        & Trim(.txtJobID.value)     & _
                "&txtRegisterID="   & Trim(.txtRegisterID.value)& _
                "&txtTrnsFrDt="     & Trim(.txtTrnsFrDt.Text)  & _
                "&txtTrnsToDt="     & Trim(.txtTrnsToDt.Text)  & _
                "&cboJobState="     & Trim(.cboJobState.value)  & _
                "&txtResultCD="     & Trim(.txtResultCD.value)  & _
                "&txtMaxRows="      & .vspdData.MaxRows         & _
                "&lgStrPrevKey="    & lgStrPrevKey
        End If       
        Call RunMyBizASP(MyBizASP, strVal)
    End With

    DbQuery = True
End Function

'=========================================================================================================
Function DbQueryOk(ByVal LngMaxRow)

	Dim iLngRow
	
	lgOldRow = 1
    
    With frm1.vspdData
    
		.Redraw = False
    
		For iLngRow = LngMaxRow + 1 To  .MaxRows
		
			'If GetSpreadText(frm1.vspdData, C_STATE, iLngRow, "X", "X") <> "대기" or _
			if GetSpreadText(frm1.vspdData, C_SUCESS, iLngRow, "X", "X") = "0"  Then
				ggoSpread.SpreadunLock C_CHECK,			iLngRow, C_CHECK, iLngRow
			else
				ggoSpread.SpreadLock C_CHECK,			iLngRow, C_CHECK, iLngRow
			End If
		
		Next 
		
		.Redraw = True	
		
    End With
    
    lgIntFlgMode = Parent.OPMD_UMODE
	
	frm1.btnRunJob.disabled = False
	frm1.btnCanJob.disabled = False
	
    Call ggoOper.LockField(Document, "Q")
    Call SetToolbar("11001011000111")
    
    
End Function


'========================================================================================
' Function Name : DbSave
' Function Desc : 
'========================================================================================
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt, i, j, iCnt
	Dim strVal, strDel,strTxt
	Dim ColSep,RowSep

	
    DbSave = False
                                                              
    ColSep = parent.gColSep               
	RowSep = parent.gRowSep    
	
    Call LayerShowHide(1)
    On Error Resume Next                                                   <%'☜: Protect system from crashing%>

	With frm1
		.txtMode.value = parent.UID_M0002
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
    
<%  '-----------------------
    'Data manipulate area
    '----------------------- %>
    lGrpCnt = 1
    
    strVal = ""
    strDel = ""
   
<%  '-----------------------
    'Data manipulate area
    '----------------------- %>
    ' Data 연결 규칙 
    ' 0: Flag , 1: Row위치, 2~N: 각 데이타   

    For lRow = 1 To .vspdData.MaxRows

		
		'if GetSpreadText(frm1.vspdData,C_chk,lRow,"X","X")="1" then
		.vspdData.Row = lRow
		.vspdData.Col = 0
		
			Select Case .vspdData.Text
			    Case ggoSpread.InsertFlag
					strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep		'☜: C=Create
			    Case ggoSpread.UpdateFlag
					strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep		'☜: U=Update
					
				 Case ggoSpread.DeleteFlag
					strVal = strVal & "D" & parent.gColSep & lRow & parent.gColSep		'☜: D=delete
				
			End Select			

			Select Case .vspdData.Text
			    Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag		'☜: 신규, 수정 
					
                        '없음 
			    Case ggoSpread.DeleteFlag							'☜: 삭제 
						strVal = strVal & GetSpreadtext(frm1.vspdData,C_JOB_NO,lRow,"X","X") & ColSep  & RowSep
				
  			            
			End Select
			
	   ' end if	
		Next


	
	.txtMaxRows.value = lGrpCnt-1	
	.txtSpread.value = strDel & strVal
	.hAction.value = "D"

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										<%'☜: 비지니스 ASP 를 가동 %>
	
	End With
	
    DbSave = True                                                           
    
End Function

Function DbSaveOk()
	On Error Resume Next
End Function

'=========================================================================================================
Function OpenProcessID()
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "업무코드"
    arrParam(1) = "B_BDC_MASTER"
    arrParam(2) = Trim(frm1.txtProcessID.Value)
    arrParam(3) = ""
    arrParam(4) = " USE_FLAG= " & Filtervar("Y", "''", "S")
    arrParam(5) = "업무코드"
    
    arrField(0) = "PROCESS_ID"
    arrField(1) = "PROCESS_NAME"
    
    arrHeader(0) = "업무코드"
    arrHeader(1) = "업무명"
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", _
									Array(arrParam, arrField, arrHeader), _
									"dialogWidth=420px; dialogHeight=450px; center: Yes; " & _
									"help: No; resizable: No; status: No;")

    IsOpenPop = False

    If arrRet(0) <> "" Then
        frm1.txtProcessID.Value    = Trim(arrRet(0))
        frm1.txtProcessNm.value    = Trim(arrRet(1))
    End If    

    frm1.txtProcessID.focus
    Set gActiveElement = document.activeElement
End Function

Function OpenJobID()
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "작업코드"
    arrParam(1) = "B_BDC_JOBS"
    arrParam(2) = Trim(frm1.txtJobID.Value)
    arrParam(3) = ""
    arrParam(4) = ""
    arrParam(5) = "작업코드"
    
    arrField(0) = "JOB_ID"
    arrField(1) = "JOB_TITLE"
    
    arrHeader(0) = "작업코드"
    arrHeader(1) = "작 업 명"
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", _
									Array(arrParam, arrField, arrHeader), _
									"dialogWidth=420px; dialogHeight=450px; center: Yes;" & _
									" help: No; resizable: No; status: No;")
    
    IsOpenPop = False

    If arrRet(0) <> "" Then
        frm1.txtJobID.Value    = Trim(arrRet(0))
        frm1.txtJobNm.value    = Trim(arrRet(1))
    End If    

    frm1.txtJobID.focus
    Set gActiveElement = document.activeElement
End Function

Function OpenRegisterID()
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "등록자ID"
    arrParam(1) = "Z_USR_MAST_REC"
    arrParam(2) = Trim(frm1.txtProcessID.Value)
    arrParam(3) = ""
    arrParam(4) = ""
    arrParam(5) = "등록자ID"
    
    arrField(0) = "USR_ID"
    arrField(1) = "USR_NM"
    
    arrHeader(0) = "등록자ID"
    arrHeader(1) = "등록자성명"

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", _
									Array(arrParam, arrField, arrHeader), _
									"dialogWidth=420px; dialogHeight=450px; center: Yes;" & _
									"help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) <> "" Then
        frm1.txtRegisterID.Value    = Trim(arrRet(0))
        frm1.txtRegisterNm.value    = Trim(arrRet(1))
    End If    

    frm1.txtRegisterID.focus
    Set gActiveElement = document.activeElement
End Function

'=========================================================================================================
Function OpenAddJob()
    Dim iCalledAspName
	Dim arrRet
	Dim arrParam(2)
    
    iCalledAspName = AskPRAspName("BDC04PA1")
    
    If IsOpenPop = True Then Exit Function
    
    If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "BDC04PA1", "X")
		IsOpenPop = False
		Exit Function
	End If

    IsOpenPop = True
	
    arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=600px; dialogHeight=170px; center: Yes; help: No; resizable: No; status: No;")
		
    IsOpenPop = False
    
    If arrRet Then
		Call InitVariables
		frm1.vspdData.MaxRows = 0
		Call MainQuery()
	End If
End Function

'=========================================================================================================
Function OpenRunJob()
	Dim strVal
    Dim lRow
    Dim lGrpCnt

	Call LayerShowHide(1)

    With frm1
        .txtMode.value = Parent.UID_M0002
        .txtUpdtUserId.value = Parent.gUsrID
        .txtInsrtUserId.value = Parent.gUsrID

        lGrpCnt = 1

        strVal = ""

        For lRow = 1 To .vspdData.MaxRows
            Select Case GetSpreadText(.vspdData, 0, lRow, "X", "X")
                Case ggoSpread.UpdateFlag
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_JOB_NO,	lRow, "X", "X")) & iColSep
                    lGrpCnt = lGrpCnt + 1
            End Select
        Next

        .txtMaxRows.value = lGrpCnt-1
        .txtSpread.value = strVal
        
        Call ExecMyBizASP(frm1, BIZ_RUN_ID)
    End With
End Function

'=========================================================================================================
Function OpenRunJobOk()

	
	Call DisplayMsgBox("183114", "x", "x", "x")
    Call InitVariables
    frm1.vspdData.MaxRows = 0
    Call MainQuery()
End Function

'=========================================================================================================
Function OpenCanJob() 
    Dim lRow
    Dim lGrpCnt
    Dim retVal
    Dim strVal
    Dim iColSep, iRowSep

    iColSep = parent.gColSep
    iRowSep = parent.gRowSep

    OpenCanJob = False

    With frm1
        .txtMode.value = Parent.UID_M0002
        .txtUpdtUserId.value = Parent.gUsrID
        .txtInsrtUserId.value = Parent.gUsrID

        lGrpCnt = 1

        strVal = ""

        For lRow = 1 To .vspdData.MaxRows
            Select Case GetSpreadText(.vspdData, 0, lRow, "X", "X")
                Case ggoSpread.UpdateFlag
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_JOB_NO,   lRow, "X", "X")) & iColSep & lRow & iRowSep
                   '---------------------------------------------------------------------------------------
                    lGrpCnt = lGrpCnt + 1
            End Select
        Next
    
        .txtMaxRows.value = lGrpCnt-1
        .txtSpread.value = strVal
        
        Call ExecMyBizASP(frm1, BIZ_PGM_ID)
    End With
    
	OpenCanJob = True
End Function

Function OpenCanJobOk()
	Call InitVariables
    frm1.vspdData.MaxRows = 0
    Call MainQuery()
End Function

Function JumpJobResult()

    Dim IntRetCd, strVal
    
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

    ggoSpread.Source = frm1.vspdData                        '⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then					'⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("189217", "x", "x", "x")   '⊙: Display Message(There is no changed data.)
        Exit Function
    End If

   	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_STATE
	
	'If Trim(frm1.vspdData.Text) <> "완료" Then
	'	msgbox "@@작업상태가 완료인 작업만 상세조회를 할 수 있습니다."
		'Call DisplayMsgBox("189218", "x", "x", "x")
	'	Exit Function
	'End If
	
	frm1.vspdData.Col = C_JOB_NO
	WriteCookie "txtJobId", UCase(Trim(frm1.vspdData.Text))
	
	PgmJump(BIZ_PGM_JUMP_ID)
	
End Function



</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME="frm1" ACTION="" TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BasicTB" CELLSPACING=0>
    <TR>
        <TD HEIGHT=5>&nbsp;</TD>
    </TR>
    <TR HEIGHT=23>
        <TD WIDTH=100%>
            <TABLE CLASS="BasicTB" CELLSPACING=0>
                <TR>
                    <TD WIDTH=10>&nbsp;</TD>
                    <TD CLASS="CLSMTABP">
                        <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
                            <TR>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif">
                                    <img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23">
                                </td>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB">
                                    <font color=white>작업관리</font>
                                </td>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right">
                                    <img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23">
                                </td>
                            </TR>
                        </TABLE>
                    </TD>
                    <TD WIDTH=* align=right>
                        <a href="vbscript:OpenAddJob">작업추가</A>
                    </TD>
                    <TD WIDTH=10>&nbsp;</TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR HEIGHT=*>
        <TD WIDTH=100% CLASS="Tab11">
            <TABLE CLASS="BasicTB" CELLSPACING=0>
                <TR>
                    <TD HEIGHT=5 WIDTH=100%></TD>
                </TR>
                <TR>
                    <TD HEIGHT=20 WIDTH=100%>
                        <FIELDSET CLASS="CLSFLD">
                            <TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
                                <TR>
                                    <TD CLASS="TD5">업무코드</TD>
                                    <TD CLASS="TD6">
                                        <INPUT CLASS="clstxt" TYPE=TEXT NAME="txtProcessID" SIZE=15 MAXLENGTH=20 tag="11XXXU" ALT="업무코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProcessID"  align="top" TYPE="BUTTON" ONCLICK="vbscript:OpenProcessID()">&nbsp;
                                        <INPUT TYPE=TEXT NAME="txtProcessNm" CLASS=protected readonly=true TABINDEX="-1" SIZE=29 MAXLENGTH=40 tag="14">
                                    </TD>
                                    <TD CLASS="TD5">작업번호</TD>
                                    <TD CLASS="TD6">
                                        <INPUT CLASS="clstxt" TYPE=TEXT NAME="txtJobID" SIZE=18 MAXLENGTH=18 tag="11XXXU" ALT="작업번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProcessID"  align="top" TYPE="BUTTON" ONCLICK="vbscript:OpenJobID()">&nbsp;
                                        <INPUT TYPE=TEXT NAME="txtJobNm" CLASS=protected readonly=true TABINDEX="-1" SIZE=29 MAXLENGTH=40 tag="14">
                                    </TD>
                                </TR>
                                <TR>
                                    <TD CLASS="TD5">등록자</TD>
                                    <TD CLASS="TD6">
                                        <INPUT CLASS="clstxt" TYPE=TEXT NAME="txtRegisterID" SIZE=10 MAXLENGTH=13 tag="11XXXU" ALT="등록자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRegisterID"  align="top" TYPE="BUTTON" ONCLICK="vbscript:OpenRegisterID()">&nbsp;
                                        <INPUT TYPE=TEXT NAME="txtRegisterNm" CLASS=protected readonly=true TABINDEX="-1" SIZE=29 MAXLENGTH=40 tag="14">
                                    </TD>
                                    </TD>
									<TD CLASS="TD5" NOWRAP>등록기간</TD>
									<TD CLASS="TD6" NOWRAP>
									    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime1 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtTrnsFrDt CLASSID="CLSID:DD55D13A-EBF7-11D0-8810-0000C0E5948C" ALT="등록기간" tag="12X1"></OBJECT>');</SCRIPT>
									&nbsp;~&nbsp;
									    
									    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime1 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtTrnsToDt CLASSID="CLSID:DD55D13A-EBF7-11D0-8810-0000C0E5948C" ALT="등록기간" tag="12X1"></OBJECT>');</SCRIPT>
									</TD>
                                </TR>
                                <TR>
									<TD CLASS="TD5" NOWRAP HEIGHT=5>작업상태</TD>
									<TD CLASS="TD6" NOWRAP HEIGHT=5>
									    <SELECT Name="cboJobState" ALT="작업상태" STYLE="WIDTH: 133px" tag="11">
									        <OPTION Value=""></OPTION>
									    </SELECT>
									</TD>
                                    <TD CLASS="TD5">결과코드</TD>
                                    <TD CLASS="TD6">
										<INPUT TYPE=radio CLASS="RADIO" NAME="rdoResultFlag" ID="Radio3" VALUE="" TAG = "11X" CHECKED>
											<LABEL FOR="rdoCfmAll">전체</LABEL>&nbsp;&nbsp;
										<INPUT TYPE=radio CLASS="RADIO" NAME="rdoResultFlag" ID="Radio5" VALUE="" TAG = "11X">
											<LABEL FOR="rdoCfmSuc">성공</LABEL>&nbsp;&nbsp;
										<INPUT type=radio CLASS="RADIO" NAME="rdoResultFlag" ID="Radio4" VALUE="Y" TAG = "11X">
											<LABEL FOR="rdoCfmFail">실패</LABEL>&nbsp;&nbsp;
                                    </TD>
                                </TR>
                            </TABLE>
                        </FIELDSET>
                    </TD>
                </TR>
                <TR>
                    <TD WIDTH=100% HEIGHT=* valign=top>
                        <TABLE WIDTH="100%" HEIGHT="100%">
                            <TR>
                                <TD HEIGHT="100%">
                                    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>

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
                    <TD>	
						<TABLE <%=LR_SPACE_TYPE_30%>>
							<TR>
							  <TD WIDTH=10>&nbsp;</TD>
							  <TD align="left">
									<A><button name="btnRunJob" class="clsmbtn">작업실행</button></a>
								    <A><button name="btnCanJob" class="clsmbtn">작업취소</button></a>
							  </TD>
							  <TD WIDTH=* Align=right><A href="vbscript:JumpJobResult">작업상세조회</A> </TD>
							  <TD WIDTH=10>&nbsp;</TD>
							</TR>
						</TABLE>
					</TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>>
            <IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no NORESIZE FRAMESPACING=0></IFRAME>
        </TD>
    </TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtResultCD" tag="24">
<INPUT TYPE=HIDDEN NAME="hProcessID" tag="24">
<INPUT TYPE=HIDDEN NAME="hJobID" tag="24">
<INPUT TYPE=HIDDEN NAME="hRegisterID" tag="24">
<INPUT TYPE=HIDDEN NAME="hTrnsFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hTrnsToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hJobState" tag="24">
<INPUT TYPE=HIDDEN NAME="hResultCD" tag="24">
<INPUT TYPE=HIDDEN NAME="hAction" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
    <IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO NORESIZE FRAMESPACING=0 WIDTH=300 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
