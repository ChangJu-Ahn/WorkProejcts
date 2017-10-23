<%@ LANGUAGE="VBSCRIPT" %>
<!--
'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 회계기준정보관리 
'*  3. Program ID		    : A2106MA1
'*  4. Program Name         : 거래항목 등록 
'*  5. Program Desc         : 거래항목 등록 수정 삭제 조회 
'*  6. Component List       : +
'*  7. Modified date(First) : 2000/09/08
'*  8. Modified date(Last)  : 2001/03/05
'*  9. Modifier (First)     : Jong Hwan, Kim
'* 10. Modifier (Last)      : Hee Jung, Kim / Cho Ig Sung
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit 


Const BIZ_PGM_ID = "a2106mb1.asp"


Dim C_JnlCD
Dim C_JnlNM
Dim C_JnlEngNM
Dim C_JnlType
Dim C_JnlTypeNm
Dim C_SysFg
Dim C_TransTblNM
Dim C_TransColmNM

Dim lgBlnStartFlag
Dim  IsOpenPop
'========================================================================================================= 

<!-- #Include file="../../inc/lgvariables.inc" -->

'========================================================================================================= 
Sub InitSpreadPosVariables()
	 C_JnlCD		= 1
	 C_JnlNM		= 2
	 C_JnlEngNM		= 3
	 C_JnlType		= 4
	 C_JnlTypeNm	= 5
	 C_SysFg		= 6
	 C_TransTblNM	= 7
	 C_TransColmNM	= 8
End Sub



'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0
    lgStrPrevKey = ""
    lgLngCurRows = 0
    lgSortKey = 1

End Sub

'========================================================================================================= 
Sub SetDefaultVal()
End Sub

'========================================================================================================= 
Sub  LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
End Sub

Sub InitComboBox()

'========================================================================================================= 
Dim iCodeArr,iNameArr

	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A2006", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

	iCodeArr = vbTab & lgF0
    iNameArr = vbTab & lgF1

    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_JnlType
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_JnlTypeNm
End Sub
'========================================================================================================= 
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows
			.Row = intRow
			.Col	 = C_JnlType
			intIndex = .value

			.Col	= C_JnlTypeNM
			.value	= intindex
		Next
	End With
End Sub

'========================================================================================================= 
Sub InitSpreadSheet()

    Call initSpreadPosVariables()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021203",,parent.gAllowDragDropSpread

	With frm1.vspdData

		.MaxCols = C_TransColmNM + 1
		.MaxRows = 0

		.ReDraw = False
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit  C_JnlCD,	   "거래항목코드",  15, , , 20, 2
		ggoSpread.SSSetEdit  C_JnlNM,	   "거래항목명",    20, , , 50
		ggoSpread.SSSetEdit  C_JnlEngNM,   "거래항목영문명",30, , , 50

    	ggoSpread.SSSetCombo C_JnlType,    "",   2
		ggoSpread.SSSetCombo C_JnlTypeNm,  "사용처",   14
		Call InitComboBox()	   '모듈구분 
		ggoSpread.SSSetCheck C_SysFg,	   "시스템구분",    12,-10, "", True, -1
		ggoSpread.SSSetEdit  C_TransTblNM, "연결테이블명",  25, , , 32
		ggoSpread.SSSetEdit  C_TransColmNM,"연결컬럼명",    25, , , 32

		call ggoSpread.MakePairsColumn(C_JnlType,C_JnlTypeNm,"1")

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_JnlType,C_JnlType,True)

		.ReDraw = True

		Call SetSpreadLock

    End With
End Sub

Sub SetSpreadLock()
    With frm1.vspdData
		.ReDraw = False

		'SpreadLock(ByVal Col1, ByVal Row1, Optional ByVal Col2 = -10, Optional ByVal Row2 = -10)
		ggoSpread.SpreadLock C_JnlCD, -1, C_JnlCD
		ggoSpread.SSSetRequired C_JnlNM,-1, C_JnlNM
		ggoSpread.SpreadLock C_SysFg, -1, C_SysFg
		ggoSpread.SSSetProtected	.MaxCols,-1,-1

		.ReDraw = True
    End With
End Sub


'============================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
		.vspdData.ReDraw = False

		' SSSetRequired(ByVal Col, ByVal Row, Optional ByVal Row2 = -10)
		ggoSpread.SSSetRequired C_JnlCD,	 pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_JnlNM,	 pvStartRow, pvEndRow
'    	ggoSpread.SSSetRequired C_JnlTypeNM, lRow, lRow

		.vspdData.ReDraw = True
    End With
End Sub


'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
				C_JnlCD			= iCurColumnPos(1)
				C_JnlNM			= iCurColumnPos(2)
				C_JnlEngNM		= iCurColumnPos(3)
				C_JnlType		= iCurColumnPos(4)
				C_JnlTypeNm		= iCurColumnPos(5)
				C_SysFg			= iCurColumnPos(6)
				C_TransTblNM	= iCurColumnPos(7)
				C_TransColmNM	= iCurColumnPos(8)
	End Select
End Sub


Function OpenPopUp(Byval strCode, Byval iWhere)
Dim arrRet
Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "거래항목 팝업"
	arrParam(1) = "A_JNL_ITEM"
	arrParam(2) = strCode
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "거래항목"

    arrField(0) = "JNL_CD"
    arrField(1) = "JNL_NM"
    arrField(2) = "JNL_ENG_NM"

    arrHeader(0) = "거래항목코드"
    arrHeader(1) = "거래항목명"
    arrHeader(2) = "거래항목영문명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtJnlCD.focus
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If

End Function

Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		If iWhere = 0 Then
			.txtJnlCD.focus
			.txtJnlCD.value = arrRet(0)
			.txtJnlNM.value = arrRet(1)
		End If
	End With
End Function

Sub Form_Load()

    On Error Resume Next
    Err.Clear

    Call LoadInfTB19029
    Call ggoOper.ClearField(Document, "1")
    Call ggoOper.LockField(Document, "N")

    Call InitSpreadSheet
    Call InitVariables

'    Call SetDefaultVal
'	Call InitComboBox
    Call SetToolbar("110011010010111")

	frm1.txtJnlCD.focus 

End Sub

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub


Sub vspdData_Change(ByVal Col , ByVal Row )
	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col
	Call CheckMinNumSpread(frm1.vspdData,Col,Row)
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
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


'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub


Sub vspdData_MouseDown(Button , Shift , x , y)

	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

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

'========================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	With frm1.vspdData
		.Row = Row
		Select Case Col
			Case  C_JnlTypeNm
				.Col     = Col
				intIndex = .Value
				.Col     = C_JnlType
				.Value   = intIndex
		End Select
	End With
End Sub

'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then
			DbQuery
		End If
    End if
End Sub

'========================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData 

    If Row >= NewRow Then
        Exit Sub
    End If

    End With

End Sub

'========================================================================================================
Function FncQuery()
	Dim IntRetCD 

    FncQuery = False
    Err.Clear

	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    Call ggoOper.ClearField(Document, "2")

    If Not chkField(Document, "1") Then
       Exit Function
    End If

    Call InitVariables
    Call SetDefaultVal

    Call DbQuery
    FncQuery = True
End Function

'========================================================================================================
Function FncNew()
    Dim IntRetCD 

    FncNew = False
    Err.Clear

    FncNew = True
End Function

'========================================================================================================
Function FncDelete()
    Dim intRetCD
    FncDelete = False
    Err.Clear
    FncDelete = True
End Function

'========================================================================================================
Function FncSave()
Dim IntRetCD 

    FncSave = False
    Err.Clear

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")
        Exit Function
    End If

    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then
       Exit Function
    End If

    Call DbSave

    FncSave = True

End Function

'========================================================================================================
Function FncCopy()
Dim IntRetCD

	frm1.vspdData.ReDraw = False

    if frm1.vspdData.MaxRows < 1 then Exit Function

    ggoSpread.Source = frm1.vspdData
	ggoSpread.CopyRow
	SetSpreadColor frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow

	frm1.vspdData.Col = C_JnlCd
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Text = ""

	frm1.vspdData.ReDraw = True

End Function

'========================================================================================================
Function FncCancel()

    if frm1.vspdData.MaxRows < 1 then Exit Function

    ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo

	Call InitData
End Function


'========================================================================================================
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

'========================================================================================================
Function FncDeleteRow()
Dim lDelRows

    if frm1.vspdData.MaxRows < 1 then Exit Function

    With frm1.vspdData
		.focus
		ggoSpread.Source = frm1.vspdData
		lDelRows = ggoSpread.DeleteRow
    End With
End Function

Function FncPrint()
	Call parent.FncPrint()
End Function

Function FncPrev()
    On Error Resume Next
End Function

Function FncNext()
    On Error Resume Next
End Function

Function FncExcel()
    Call parent.FncExport(Parent.C_MULTI)
End Function

Function FncFind()
    Call parent.FncFind(Parent.C_MULTI, False)
End Function

'========================================================================================================
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
	Call InitData()
End Sub


'========================================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")   '☜ 바뀐부분 
		 If IntRetCD = vbNo Then
		     Exit Function
		End If
	End If

    FncExit = True
End Function

'========================================================================================================
Function DbQuery()

    DbQuery = False
    Err.Clear

	Call LayerShowHide(1)

	Dim strVal

    With frm1
	    If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID	& "?txtMode="		& Parent.UID_M0001
			strVal = strVal		& "&txtJnlCd="		& Trim(.hJnlCd.value)
			strVal = strVal		& "&lgStrPrevKey="	& lgStrPrevKey
			strVal = strVal		& "&txtMaxRows="	& .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_ID	& "?txtMode="		& Parent.UID_M0001
			strVal = strVal		& "&txtJnlCd="		& Trim(.txtJnlCd.value)
			strVal = strVal		& "&lgStrPrevKey="	& lgStrPrevKey
			strVal = strVal		& "&txtMaxRows="	& .vspdData.MaxRows
		End If
		strVal = strVal & "&lgPageNo="       & lgPageNo
    End With

		Call RunMyBizASP(MyBizASP, strVal)

    DbQuery = True

End Function

'========================================================================================================
Function DbQueryOk()

    lgIntFlgMode = Parent.OPMD_UMODE

    Call ggoOper.LockField(Document, "Q")
	Call InitData

	Call SetToolbar("110011110011111")
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function DbSave()
	Dim lRow
	Dim lGrpCnt
	Dim strVal, strDel

    DbSave = False

	Call LayerShowHide(1)

	lGrpCnt = 1
	strVal = ""
	strDel = ""

	With frm1
		For lRow = 1 To .vspdData.MaxRows

		    .vspdData.Row = lRow
		    .vspdData.Col = 0
		    Select Case .vspdData.Text
		        Case ggoSpread.InsertFlag
													  strVal = strVal & "C" & Parent.gColSep & lRow & Parent.gColSep
		            .vspdData.Col = C_JnlCD			: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_JnlNM			: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					.vspdData.Col = C_JnlEngNM		: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					.vspdData.Col = C_JnlType		: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					.vspdData.Col = C_SysFg
					If  IsNull(.vspdData.Text) Then 
					    .vspdData.Text  = ""
					End If

					If  Trim(CStr(.vspdData.Text)) = "" Or   Trim(CStr(.vspdData.Text)) = "0" Then 
						strVal = strVal & "N" & Parent.gColSep
					Else
						strVal = strVal & "Y" & Parent.gColSep
					END IF

					.vspdData.Col = C_TransTblNM	: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					.vspdData.Col = C_TransColmNM	: strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
					lGrpCnt = lGrpCnt + 1
				Case ggoSpread.UpdateFlag
													  strVal = strVal & "U" & Parent.gColSep & lRow & Parent.gColSep
		            .vspdData.Col = C_JnlCD			: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_JnlNM			: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					.vspdData.Col = C_JnlEngNM		: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					.vspdData.Col = C_JnlType		: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					.vspdData.Col = C_SysFg
					If  IsNull(.vspdData.Text) Then 
					    .vspdData.Text  = ""
					End If

					If  Trim(CStr(.vspdData.Text)) = "" Or   Trim(CStr(.vspdData.Text)) = "0" Then 
						strVal = strVal & "N" & Parent.gColSep
					Else
						strVal = strVal & "Y" & Parent.gColSep
					END IF
						
					.vspdData.Col = C_TransTblNM	: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					.vspdData.Col = C_TransColmNM	: strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
					lGrpCnt = lGrpCnt + 1

			    Case ggoSpread.DeleteFlag
													  strVal = strVal & "D" & Parent.gColSep & lRow & Parent.gColSep
					.vspdData.Col = C_JnlCD			: strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
					lGrpCnt = lGrpCnt + 1
			End Select
		Next

		.txtMode.value		= Parent.UID_M0002
		.txtMaxRows.value	= lGrpCnt-1
		.txtSpread.value	= strDel & strVal

	End With
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    DbSave = True

End Function

'========================================================================================================
Function DbSaveOk()
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call InitVariables															'⊙: Initializes local global variables
	call DBQuery()
End Function

'========================================================================================================
Function DbDelete()
	On Error Resume Next
End Function

'========================================================================================================
Function DbDeleteOk()
	On Error Resume Next
End Function


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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>거래항목등록</font></td>
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
									<TD CLASS="TD5" NOWRAP>거래항목</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtJnlCD" MAXLENGTH="20" SIZE=20 ALT ="거래항목코드" tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtJnlCD.Value, 0)">&nbsp;
													<INPUT NAME="txtJnlNM" MAXLENGTH="50" SIZE=30 ALT  ="거래항목명" tag="14X"></TD>
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
								<script language =javascript src='./js/a2106ma1_I284121578_vspdData.js'></script>
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
			<IFRAME NAME="MyBizASP"  src="../../blank.htm"  WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA class=hidden name=txtSpread tag="24" tabindex="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtMaxRows" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="hJnlCd" tag="24" tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
