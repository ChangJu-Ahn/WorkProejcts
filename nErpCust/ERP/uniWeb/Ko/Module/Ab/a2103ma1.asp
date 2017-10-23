<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID		    : A2103MA1
'*  4. Program Name         : �����з����� ��� 
'*  5. Program Desc         : �����з����� ��� ���� ���� ��ȸ 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/09/08
'*  8. Modified date(Last)  : 2002/11/25
'*  9. Modifier (First)     : Jong Hwan, Kim
'* 10. Modifier (Last)      : Chang Joo, Kim
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
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

Option Explicit																	'��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_ID = "a2103mb1.asp"			'��: �����Ͻ� ���� ASP�� 
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
'��: Grid Columns
Dim C_ClassType 
Dim C_ClassTypeNM
Dim C_ClassTypeEngNM

Sub InitSpreadPosVariables()
	C_ClassType			= 1
	C_ClassTypeNM		= 2
	C_ClassTypeEngNM	= 3
End Sub

<%
StartDate = DateSerial(Year(Date),Month(Date),1)								'��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ 

StartDate= Year(StartDate) & "-" & Right("0" & Month(StartDate),2) & "-" & Right("0" & Day(StartDate),2)
EndDate= Year(Date) & "-" & Right("0" & Month(Date),2) & "-" & Right("0" & Day(Date),2)	'��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ 
%>

<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================
'Dim lgStrPrevKey
Dim lgBlnStartFlag				' �޼��� �����Ͽ� ���α׷� ���۽��� Check Flag


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

		ggoSpread.SSSetEdit C_ClassType, "�����з������ڵ�", 28, , , 4, 2
		ggoSpread.SSSetEdit C_ClassTypeNM, "�����з����¸�", 45, , , 50
		ggoSpread.SSSetEdit C_ClassTypeEngNM, "�����з����¿�����", 45, , , 50

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	    .ReDraw = True

		Call SetSpreadLock                                              '�ٲ�κ� 

    End With

End Sub


'========================================================================================
Sub SetSpreadLock()

    With frm1.vspdData
		.ReDraw = False
		ggoSpread.SpreadLock C_ClassType, -1, C_ClassType		' �����з�Type�� Lock
		ggoSpread.SSSetRequired C_ClassTypeNM, -1, C_ClassTypeNM
		ggoSpread.SSSetProtected	.MaxCols,-1,-1
		.ReDraw = True
    End With

End Sub

'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1

		.vspdData.ReDraw = False

		' �ʼ� �Է� �׸����� ���� 
		' SSSetRequired(ByVal Col, ByVal Row, Optional ByVal Row2 = -10)
		ggoSpread.SSSetRequired C_ClassType, pvStartRow, pvEndRow	' �����з������ڵ� 
		ggoSpread.SSSetRequired C_ClassTypeNM, pvStartRow, pvEndRow	' �����з����¸� 

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

	arrParam(0) = "�����з����� �˾�"			' �˾� ��Ī 
	arrParam(1) = "A_ACCT_CLASS_TYPE" 			' TABLE ��Ī 
	arrParam(2) = strCode						' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "�����з�����"				' �����ʵ��� �� ��Ī 

    arrField(0) = "CLASS_TYPE"					' Field��(0)
    arrField(1) = "CLASS_TYPE_NM"				' Field��(1)
    arrField(2) = "CLASS_TYPE_ENG_NM"			' Field��(2)

    arrHeader(0) = "�����з������ڵ�"				' Header��(0)
    arrHeader(1) = "�����з����¸�"				' Header��(1)
    arrHeader(2) = "�����з����¿�����"			' Header��(2)

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
'	Description : CtrlItem Popup���� Return�Ǵ� �� setting
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
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")			    '����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?
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
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X","X") '�� �ٲ�κ� 
         'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. �ű��۾��� �Ͻðڽ��ϱ�?", vbYesNo)
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
    Call DbSave				                                                  '��: Save db data
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
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")   '�� �ٲ�κ� 
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

		'strVal = BIZ_PGM_ID & "?txtMode="		& Parent.UID_M0001						'��:��ȸǥ�� 
		'strVal = strVal		& "&lgStrPrevKey="	& lgStrPrevKey
		'strVal = strVal		& "&txtClassType="		& Trim(.txtClassType.value)	 			    '��: ��ȸ ���� ����Ÿ			
		'--------------------------------------------------------------------------

	    If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtClassType=" & Trim(.hClassType.value)	'��ȸ ���� ����Ÿ 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtClassType=" & Trim(.txtClassType.value)	'��ȸ ���� ����Ÿ 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If
    		strVal = strVal & "&lgPageNo="       & lgPageNo

		Call RunMyBizASP(MyBizASP, strVal)		'��: �����Ͻ� ASP �� ���� 

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
		        Case ggoSpread.InsertFlag							'��: �ű� 
					strVal = strVal & "C" & Parent.gColSep & lRow & Parent.gColSep					'��: C=Create
		            .vspdData.Col = C_ClassType
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_ClassTypeNM
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_ClassTypeEngNM
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
		            lGrpCnt = lGrpCnt + 1
				Case ggoSpread.UpdateFlag							'��: ���� 
					strVal = strVal & "U" & Parent.gColSep & lRow & Parent.gColSep					'��: U=Update
				    .vspdData.Col = C_ClassType
				    strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
				    .vspdData.Col = C_ClassTypeNM
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_ClassTypeEngNM
				    strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
				    lGrpCnt = lGrpCnt + 1
		        Case ggoSpread.DeleteFlag							'��: ���� 
					strVal = strVal & "D" & Parent.gColSep & lRow & Parent.gColSep					'��: U=Update
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


'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>�����з����µ��</font></td>
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
									<TD CLASS="TD5" NOWRAP>�����з�����</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtClassType" MAXLENGTH="4" SIZE=20 ALT ="�����з�����" tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtClassType.Value, 0)">&nbsp;
													<INPUT NAME="txtClassTypeNM" MAXLENGTH="50" SIZE=30 ALT ="�����з����¸�" tag="14X"></TD>
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

