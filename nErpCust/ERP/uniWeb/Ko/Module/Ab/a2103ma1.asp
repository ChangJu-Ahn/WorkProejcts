<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID		    : A2103MA1
'*  4. Program Name         : 계정분류형태 등록 
'*  5. Program Desc         : 계정분류형태 등록 수정 삭제 조회 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/09/08
'*  8. Modified date(Last)  : 2002/11/25
'*  9. Modifier (First)     : Jong Hwan, Kim
'* 10. Modifier (Last)      : Chang Joo, Kim
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'********************************************************************************************** -->
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
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'⊙: 비지니스 로직 ASP명 
Const BIZ_PGM_ID = "a2103mb1.asp"			'☆: 비지니스 로직 ASP명 
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
'⊙: Grid Columns
Dim C_ClassType 
Dim C_ClassTypeNM
Dim C_ClassTypeEngNM

Sub InitSpreadPosVariables()
	C_ClassType			= 1
	C_ClassTypeNM		= 2
	C_ClassTypeEngNM	= 3
End Sub

<%
StartDate = DateSerial(Year(Date),Month(Date),1)								'☆: 초기화면에 뿌려지는 시작 날짜 

StartDate= Year(StartDate) & "-" & Right("0" & Month(StartDate),2) & "-" & Right("0" & Day(StartDate),2)
EndDate= Year(Date) & "-" & Right("0" & Month(Date),2) & "-" & Right("0" & Day(Date),2)	'☆: 초기화면에 뿌려지는 마지막 날짜 
%>

<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================
'Dim lgStrPrevKey
Dim lgBlnStartFlag				' 메세지 관련하여 프로그램 시작시점 Check Flag


'========================================================================================

Dim IsOpenPop


'========================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
    lgPageNo     = "0"

End Sub

'========================================================================================
Sub SetDefaultVal()
End Sub


'========================================================================================
Sub  LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
End Sub


'========================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021203",,parent.gAllowDragDropSpread

	With frm1.vspdData

		.MaxCols = C_ClassTypeEngNM + 1
		.MaxRows = 0


		.ReDraw = False

        Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit C_ClassType, "계정분류형태코드", 28, , , 4, 2
		ggoSpread.SSSetEdit C_ClassTypeNM, "계정분류형태명", 45, , , 50
		ggoSpread.SSSetEdit C_ClassTypeEngNM, "계정분류형태영문명", 45, , , 50

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	    .ReDraw = True

		Call SetSpreadLock                                              '바뀐부분 

    End With

End Sub


'========================================================================================
Sub SetSpreadLock()

    With frm1.vspdData
		.ReDraw = False
		ggoSpread.SpreadLock C_ClassType, -1, C_ClassType		' 계정분류Type을 Lock
		ggoSpread.SSSetRequired C_ClassTypeNM, -1, C_ClassTypeNM
		ggoSpread.SSSetProtected	.MaxCols,-1,-1
		.ReDraw = True
    End With

End Sub

'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1

		.vspdData.ReDraw = False

		' 필수 입력 항목으로 설정 
		' SSSetRequired(ByVal Col, ByVal Row, Optional ByVal Row2 = -10)
		ggoSpread.SSSetRequired C_ClassType, pvStartRow, pvEndRow	' 계정분류형태코드 
		ggoSpread.SSSetRequired C_ClassTypeNM, pvStartRow, pvEndRow	' 계정분류형태명 

		.vspdData.ReDraw = True

    End With

End Sub

'========================================================================================
Sub InitComboBox()
End Sub

'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

				C_ClassType			= iCurColumnPos(1)
				C_ClassTypeNM			= iCurColumnPos(2)
				C_ClassTypeEngNM				= iCurColumnPos(3)
    End Select
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

	arrParam(0) = "계정분류형태 팝업"			' 팝업 명칭 
	arrParam(1) = "A_ACCT_CLASS_TYPE" 			' TABLE 명칭 
	arrParam(2) = strCode						' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "계정분류형태"				' 조건필드의 라벨 명칭 

    arrField(0) = "CLASS_TYPE"					' Field명(0)
    arrField(1) = "CLASS_TYPE_NM"				' Field명(1)
    arrField(2) = "CLASS_TYPE_ENG_NM"			' Field명(2)

    arrHeader(0) = "계정분류형태코드"				' Header명(0)
    arrHeader(1) = "계정분류형태명"				' Header명(1)
    arrHeader(2) = "계정분류형태영문명"			' Header명(2)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtClassType.focus
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
		If iWhere = 0 Then		' Condition
			.txtClassType.focus
			.txtClassType.value = Trim(arrRet(0))
			.txtClassTypeNM.value = arrRet(1)
		End If
	End With
End Function


'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N")
    Call InitSpreadSheet
    Call InitVariables
    Call SetToolbar("1100110100101111")
    frm1.txtClassType.focus 

End Sub

'========================================================================================================= 
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'========================================================================================================= 
Sub vspdData_Click(ByVal Col , ByVal Row)

	If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else
		Call SetPopupMenuItemInf("1101111111")
	End If

	gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData
    If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
       Exit Sub
   	End If

    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
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
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub


'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub


'========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'========================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With frm1.vspdData
		If Row >= NewRow Then
			Exit Sub
		End If
	End With
End Sub


'========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
		If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
    	If lgPageNo <> "" Then
           If DbQuery = False Then
              Exit Sub
           End if
    	End If
    End If
End Sub



'========================================================================================================= 
Function FncQuery()
Dim IntRetCD 

    FncQuery = False
    Err.Clear

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")			    '데이타가 변경되었습니다. 조회하시겠습니까?
    	If IntRetCD = vbNo Then
	      	Exit Function
    	End If
    End If

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")
    Call InitVariables

    '-----------------------
    If Not chkField(Document, "1") Then
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery

    FncQuery = True

End Function


'========================================================================================
Function FncNew() 
Dim IntRetCD 

    FncNew = False
    Err.Clear

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X","X") '☜ 바뀐부분 
         'IntRetCD = MsgBox("데이타가 변경되었습니다. 신규작업을 하시겠습니까?", vbYesNo)
         If IntRetCD = vbNo Then
             Exit Function
         End If
    End If

    '-----------------------
    Call ggoOper.ClearField(Document, "1")
    Call ggoOper.LockField(Document, "N")

    Call InitSpreadSheet
    Call InitVariables

    Call SetToolbar("1100110100101111")
    frm1.txtClassType.focus

    FncNew = True
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement

End Function


'========================================================================================================= 
Function FncDelete() 
Dim IntRetCD 

    FncDelete = False
    Err.Clear


    If lgIntFlgMode <> Parent.OPMD_UMODE Then
        IntRetCD = DisplayMsgBox("900002", "X","X","X")
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")
    If IntRetCD = vbNo Then
        Exit Function
    End If

    If DbDelete = False Then
       Exit Function
    End If


    Call ggoOper.ClearField(Document, "1")

    FncDelete = True

End Function


'========================================================================================================= 
Function FncSave()
Dim IntRetCD

    FncSave = False
    Err.Clear

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False  Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")
        Exit Function
    End If

    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then
       Exit Function
    End If

   '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave				                                                  '☜: Save db data
    FncSave = True
End Function


'========================================================================================
Function FncCopy()
	Dim IntRetCD

	if frm1.vspdData.maxrows < 1 then exit function 

	frm1.vspdData.ReDraw = False

    ggoSpread.Source = frm1.vspdData
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow

    frm1.vspdData.Col = C_ClassType
    frm1.vspdData.Text = ""

	frm1.vspdData.ReDraw = True

End Function


'========================================================================================
Function FncCancel()
    ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo
End Function


'========================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    On Error Resume Next
    Err.Clear

    FncInsertRow = False

    If IsNumeric(Trim(pvRowCnt)) then
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

    If Err.number = 0 Then
       FncInsertRow = True
    End If

    Set gActiveElement = document.ActiveElement
End Function




'========================================================================================
Function FncDeleteRow()
Dim lDelRows
Dim iDelRowCnt, i

    With frm1.vspdData 

		.focus
		ggoSpread.Source = frm1.vspdData 

		lDelRows = ggoSpread.DeleteRow

    End With
End Function


'========================================================================================
Function FncPrev()
    On Error Resume Next
End Function


'========================================================================================
Function FncNext()
    On Error Resume Next
End Function


'========================================================================================
Function FncPrint()
	Call parent.FncPrint()
End Function


'========================================================================================
Function FncExcel()
    Call parent.FncExport(Parent.C_MULTI)
End Function


'========================================================================================
Function FncFind()
    Call parent.FncFind(Parent.C_MULTI, False)
End Function

'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub

'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
End Sub


'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")   '☜ 바뀐부분 
		 If IntRetCD = vbNo Then
		     Exit Function
		End If
	End If

    FncExit = True
End Function

'========================================================================================================= 
Function DbQuery()
Dim strVal

    DbQuery = False
    Err.Clear

	Call LayerShowHide(1)

    With frm1

		'strVal = BIZ_PGM_ID & "?txtMode="		& Parent.UID_M0001						'☜:조회표시 
		'strVal = strVal		& "&lgStrPrevKey="	& lgStrPrevKey
		'strVal = strVal		& "&txtClassType="		& Trim(.txtClassType.value)	 			    '☆: 조회 조건 데이타			
		'--------------------------------------------------------------------------

	    If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtClassType=" & Trim(.hClassType.value)	'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtClassType=" & Trim(.txtClassType.value)	'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If
    		strVal = strVal & "&lgPageNo="       & lgPageNo

		Call RunMyBizASP(MyBizASP, strVal)		'☜: 비지니스 ASP 를 가동 

    End With

    DbQuery = True

End Function

'========================================================================================================= 
Function DbQueryOk()
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE
    Call ggoOper.LockField(Document, "Q")
    Call SetToolbar("1100111100111111")
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement
End Function


'========================================================================================================= 
Function DbSave() 
Dim lRow
Dim lGrpCnt
Dim strVal, strDel
	
    DbSave = False

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
		For lRow = 1 To .vspdData.MaxRows

		    .vspdData.Row = lRow
		    .vspdData.Col = 0

		    Select Case .vspdData.Text
		        Case ggoSpread.InsertFlag							'☜: 신규 
					strVal = strVal & "C" & Parent.gColSep & lRow & Parent.gColSep					'☜: C=Create
		            .vspdData.Col = C_ClassType
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_ClassTypeNM
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_ClassTypeEngNM
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
		            lGrpCnt = lGrpCnt + 1
				Case ggoSpread.UpdateFlag							'☜: 수정 
					strVal = strVal & "U" & Parent.gColSep & lRow & Parent.gColSep					'☜: U=Update
				    .vspdData.Col = C_ClassType
				    strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
				    .vspdData.Col = C_ClassTypeNM
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_ClassTypeEngNM
				    strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
				    lGrpCnt = lGrpCnt + 1
		        Case ggoSpread.DeleteFlag							'☜: 삭제 
					strVal = strVal & "D" & Parent.gColSep & lRow & Parent.gColSep					'☜: U=Update
		            .vspdData.Col = C_ClassType
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
		            lGrpCnt = lGrpCnt + 1
		    End Select
		Next

		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strDel & strVal

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	End With
    DbSave = True
End Function


'========================================================================================================= 
Function DbSaveOk()
	Call ggoOper.ClearField(Document, "2")
	Call InitVariables
	Call DbQuery
End Function


'========================================================================================================= 
Function DbDelete()
	On Error Resume Next
End Function


'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>계정분류형태등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>계정분류형태</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtClassType" MAXLENGTH="4" SIZE=20 ALT ="계정분류형태" tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtClassType.Value, 0)">&nbsp;
													<INPUT NAME="txtClassTypeNM" MAXLENGTH="50" SIZE=30 ALT ="계정분류형태명" tag="14X"></TD>
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
								<TD HEIGHT="100%">
								<script language =javascript src='./js/a2103ma1_I965589672_vspdData.js'></script>
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
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" src="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA class=hidden name=txtSpread tag="24" tabindex="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtMaxRows" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="hClassType" tag="24" tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

