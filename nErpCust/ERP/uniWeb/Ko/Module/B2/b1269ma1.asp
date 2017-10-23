<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : 세무서정보 등록 
'*  3. Program ID           : B1269MA1
'*  4. Program Name         : 세무서정보 등록 
'*  5. Program Desc         : 부가세 신고 사업장 정보와 해당 세무서코드 정보 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/04/18
'*  8. Modified date(Last)  : 2001/03/07
'*  9. Modifier (First)     : Kang Chang Goo
'* 10. Modifier (Last)      : Kim Hee Jung / Cho Ig Sung
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/18 : ..........
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit 

			'☜: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" -->			


'**********************************************************************************************************
Const BIZ_PGM_ID = "b1269mb1.asp"												'☆: 비지니스 로직 ASP명 


'========================================================================================================= 
Dim C_TaxCd 
Dim C_TaxNm	
Dim C_TaxEngNm 


'========================================================================================================= 
Dim IsOpenPop
      
'========================================================================================================
Sub initSpreadPosVariables()         '1.2 변수에 Constants 값을 할당 
	 C_TaxCd    = 1														'☆: Spread Sheet의 Column별 상수  
	 C_TaxNm	 = 2
	 C_TaxEngNm = 3
End Sub

'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode	 = Parent.OPMD_CMODE               'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount    = 0                        'initializes Group View Size
    
    lgStrPrevKey	= ""						    'initializes Previous Key
    lgLngCurRows	= 0                            'initializes Deleted Rows Count    
    lgSortKey		= 1
    lgPageNo        = "0"
End Sub


'========================================================================================================= 
Sub SetDefaultVal()
End Sub
'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE", "MA") %>
End Sub


'========================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()    
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021106",,Parent.gAllowDragDropSpread            '2.1 [Spreadinit]에 Source Version No. 와 DragDrop지원여부를 설정 
	With frm1.vspdData
	
		.MaxCols = C_TaxEngNm + 1									'☜: 최대 Columns의 항상 1개 증가시킴 
		.MaxRows = 0

		.ReDraw = false

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit C_TaxCd,    "세무서코드",   30,,,10,2
		ggoSpread.SSSetEdit C_TaxNm,    "세무서명",     44,,,20
		ggoSpread.SSSetEdit C_TaxEngNm, "세무서 영문명",44,,,50

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
    With frm1

    .vspdData.ReDraw = False
	ggoSpread.SpreadLock C_TaxCd, -1, C_TaxCd
	ggoSpread.SSSetRequired C_TaxNm, -1, C_TaxNm
	ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    .vspdData.ReDraw = True

    End With
End Sub


'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SSSetRequired		C_TaxCd, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_TaxNm, pvStartRow, pvEndRow
    .vspdData.ReDraw = True
    
    End With
End Sub

'------------------------------------------  OpenPopUp()  -------------------------------------------------
'	Name : OpenPopUp()
'	Description : PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
		Case 0
			arrParam(0) = "세무서 팝업"				' 팝업 명칭 
			arrParam(1) = "B_TAX_OFFICE"    			' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "세무서"					' 조건필드의 라벨 명칭 

			arrField(0) = "TAX_OFFICE_CD"				' Field명(0)
			arrField(1) = "TAX_OFFICE_NM"			  	' Field명(1)
   
			arrHeader(0) = "세무서코드"					' Header명(0)
			arrHeader(1) = "세무서명"				' Header명(1)
	End Select

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
	Select Case iWhere
		Case 0
			frm1.txtTaxCd.focus
	End Select
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function

'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0
				.txtTaxCd.focus
				.txtTaxCd.value = arrRet(0)
				.txtTaxNm.value = arrRet(1)
		End Select

	End With

End Function

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
			C_TaxCd          = iCurColumnPos(1)
			C_TaxNm          = iCurColumnPos(2)
			C_TaxEngNm       = iCurColumnPos(3)

    End Select    

End Sub

'========================================================================================================= 
Sub Form_Load()

	Call LoadInfTB19029 
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call InitSpreadSheet
    Call InitVariables
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call SetToolbar("1100110100101111")
    frm1.txtTaxCd.focus    
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
  
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


'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

    If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else
		Call SetPopupMenuItemInf("1101111111")
	End If

	gMouseClickStatus = "SPC"

	Set gActiveSpdSheet = frm1.vspdData

	If Row = 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort
			lgSortKey = 2
		Else
			ggoSpread.SSSort ,lgSortKey
			lgSortKey = 1
		End If
	End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData.Row = Row
End Sub


Sub vspdData_MouseDown(Button , Shift , x , y)

	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub


'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc : 
'==========================================================================================

Sub vspdData_Change(ByVal Col, ByVal Row)
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

'    If Frm1.vspdData.CellType = Parent.SS_CELL_TYPE_FLOAT Then
'      If CDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
'         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
'      End If
'    End If
    Call CheckMinNumSpread(frm1.vspdData, Col, Row) 

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
End Sub


'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : 
'==========================================================================================

Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData 
		If Row >= NewRow Then
		    Exit Sub
		End If

		'If NewRow = .MaxRows Then
		'    DbQuery
		'End if    
    End With
End Sub



'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================

Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub


'==========================================================================================
'   Event Name :vspdData_KeyPress
'   Event Desc :
'==========================================================================================

Sub vspdData_KeyPress(index , KeyAscii )
     lgBinFlgChgValue = True                                                 '⊙: Indicates that value changed
End Sub


'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
    	If lgPageNo <> "" Then
			Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End if
    	End If
    End if
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    'If Col <= C_SNm Or NewCol <= C_SNm Then
     '   Cancel = True
      '  Exit Sub
   ' End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub




'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")			    '데이타가 변경되었습니다. 조회하시겠습니까?
    	If IntRetCD = vbNo Then
	      	Exit Function
    	End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
	ggoSpread.Source= frm1.vspdData
	ggoSpread.ClearSpreadData
	Call InitVariables 															'⊙: Initializes local global variables
    '-----------------------
    'Check condition area
    '-----------------------

    if frm1.txtTaxCd.value = "" then
		frm1.txtTaxNm.value = ""
    end if

    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery																'☜: Query db data
       
    FncQuery = True																'⊙: Processing is OK
    
End Function


'========================================================================================
Function FncNew() 
	On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False
    Err.Clear
'    On Error Resume Next
    
    ggoSpread.Source = frm1.vspddata
    If ggoSpread.SSCheckChange = False  Then  '⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","x","x","x")            '⊙: Display Message(There is no changed data.)
        Exit Function
    End If

  '-----------------------
    'Check content area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
       Exit Function
    End If
    
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave				                                                  '☜: Save db data

    FncSave = True                                                          '⊙: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 

    if frm1.vspdData.MaxRows < 1 then Exit Function

    ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function  FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    
    On Error Resume Next
    
    FncInsertRow = False
   ' Err.Clear

    if IsNumeric(Trim(pvRowCnt)) Then
			imRow = CInt(pvRowCnt)
		else
			imRow = AskSpdSheetAddRowcount()

			If ImRow="" then
			Exit Function
			End If
	End If
	
	If imRow = "" Then
		Exit function
	End If
		
	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow -1
       .vspdData.ReDraw = True
    End With
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
    FncInsertRow = True   
    
    IF Err.number = 0 Then
	    FncInsertRow = True
	End If
	
	
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
	Dim lDeIRows

    if frm1.vspdData.MaxRows < 1 then Exit Function

    With frm1.vspdData 

    .focus
    ggoSpread.Source = frm1.vspdData 

	lDeIRows = ggoSpread.DeleteRow

    End With

End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 

	if frm1.vspdData.MaxRows < 1 then Exit Function

	frm1.vspdData.ReDraw = False

    ggoSpread.Source = frm1.vspdData
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow

    frm1.vspdData.Col = C_TaxCd
    frm1.vspdData.Text = ""

	frm1.vspdData.ReDraw = True

Set gActiveElement = document.ActiveElement

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
    Call parent.FncExport(Parent.C_MULTI)												'☜: 화면 유형 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                         '☜:화면 유형, Tab 유무 
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
' Function Desc : 
'========================================================================================

Function FncExit()
	Dim IntRetCD
	FncExit = False

    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")                '데이타가 변경되었습니다. 종료 하시겠습니까?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    FncExit = True
End Function
 '*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
	Dim strVal

    DbQuery = False
    Call LayerShowHide(1)

    Err.Clear                                                               '☜: Protect system from crashing

    With frm1

		strVal = BIZ_PGM_ID & "?txtMode="		& Parent.UID_M0001						'☜:조회표시 
		strVal = strVal		& "&lgStrPrevKey="	& lgStrPrevKey
		strVal = strVal		& "&txtTaxCd="		& Trim(.txtTaxCd.value)	 			    '☆: 조회 조건 데이타			
		'--------------------------------------------------------------------------
		strVal = strVal & "&lgPageNo="       & lgPageNo         
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows		
		Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

    End With

    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    
    Call ggoOper.LockField(Document, "Q")	'⊙: This function lock the suitable field
   	Call SetToolbar("1100111100111111")
	frm1.vspdData.focus
    Set gActiveElement = document.ActiveElement   
End Function


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbSave() 
    Dim IRow
    Dim lGrpCnt
	Dim strVal
	Dim strDel

    DbSave = False                                                          '⊙: Processing is NG

    'On Error Resume Next                                                   '☜: Protect system from crashing

    Call LayerShowHide(1)

	With frm1

		.txtMode.value = Parent.UID_M0002

		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
		strDel = ""

		'-----------------------
		'Data manipulate area
		'-----------------------
		For IRow = 1 To .vspdData.MaxRows

			.vspdData.Row = IRow
			.vspdData.Col = 0

			Select Case .vspdData.Text
	            Case ggoSpread.InsertFlag											'☜: 신규 

					strVal = strVal & "C" & Parent.gColSep & IRow & Parent.gColSep				'☜: U=Update
					.vspdData.Col = C_TaxCd

					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					.vspdData.Col = C_TaxNm

					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					.vspdData.Col = C_TaxEngNm

					strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
					lGrpCnt = lGrpCnt + 1

	            Case ggoSpread.UpdateFlag											'☜: 수정 
					strVal = strVal & "U" & Parent.gColSep & IRow & Parent.gColSep				'☜: U=Update
					.vspdData.Col = C_TaxCd

					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					.vspdData.Col = C_TaxNm

					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					.vspdData.Col = C_TaxEngNm

					strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
					lGrpCnt = lGrpCnt + 1

	            Case ggoSpread.DeleteFlag											'☜: 삭제 
					strDel = strDel & "D" & Parent.gColSep & IRow & Parent.gColSep				'☜: U=Update
					.vspdData.Col = C_TaxCd		'1

					strDel = strDel & Trim(.vspdData.Text) & Parent.gRowSep
					lGrpCnt = lGrpCnt + 1
	        End Select
		Next

		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strDel & strVal

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)									'☜: 비지니스 ASP 를 가동 

	End With
    DbSave = True
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================

Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
	Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field   
	ggoSpread.Source= frm1.vspdData
	ggoSpread.ClearSpreadData
	Call InitVariables
	Call Dbquery

End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->

</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>세무서정보등록</font></td>
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
									<TD CLASS="TD5">세무서</TD>
									<TD CLASS="TD656"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtTaxCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="세무서"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTaxCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtTaxCd.Value, 0)">
										 <INPUT TYPE=TEXT ID="txtTaxNm" NAME="txtTaxNm" SIZE=30 tag="14X">
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
								<TD HEIGHT="100%"><script language =javascript src='./js/b1269ma1_I840397459_vspdData.js'></script>
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
		<!--<TD WIDTH=100% HEIGHT=30%><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=100% FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>-->		
		<TD WIDTH=50% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 tabindex="-1"></IFRAME></TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" tabindex="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hTaxCd" tag="24" tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML> 
