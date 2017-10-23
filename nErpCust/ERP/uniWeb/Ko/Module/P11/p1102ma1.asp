<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1102ma1.asp
'*  4. Program Name         : Calendar Creation
'*  5. Program Desc         :
'*  6. Component List		:
'*  7. Modified date(First) : 2000/04/08
'*  8. Modified date(Last)  : 2002/12/16
'*  9. Modifier (First)     : Mr  KimGyoungDon
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'��: indicates that All variables must be declared in advance

Dim BaseDate
Dim strYear
Dim strMonth
DIm strDay
Dim lgMaxYear
DIm lgMinYear

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************

<!-- #Include file="../../inc/lgvariables.inc" -->	

BaseDate = "<%=GetSvrDate%>"
Call ExtractDateFrom(BaseDate, Parent.gServerDateFormat, Parent.gServerDateType, strYear, strMonth, strDay)

lgMaxYear = StrYear + 20
lgMinYear = StrYear - 10

Const BIZ_PGM_BATCH_ID  = "p1102mb2.asp"
Const BIZ_PGM_LOOKUP_ID  = "p1102mb4.asp"

Dim C_Date 
Dim C_Remark 

Dim IsOpenPop

'========================== 2.2.3 InitSpreadSheet() =====================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
  Sub InitSpreadSheet()

	Call InitSpreadPosVariables()

	With frm1.vspdData

		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021123", , Parent.gAllowDragDropSpread
	
		.ReDraw = False

		.MaxCols = C_Remark + 1													'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.MaxRows = 0
		
		Call GetSpreadColumnPos("A")
	
		ggoSpread.SSSetDate 	C_Date,		"��¥", 11, 2, Parent.gDateFormat	
		ggoSpread.SSSetEdit 	C_Remark,	"���", 42,,,20
	
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)

		.ReDraw = True
	
    Call SetSpreadLock 
    
    End With
    
End Sub


'========================================================================================
' Function Name : GetSpreadColumnPos
' Function Desc : 
'======================================================================================== 
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
	
		Case "A"
			ggoSpread.Source = frm1.vspdData 
		
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_Date	 = iCurColumnPos(1)
			C_Remark = iCurColumnPos(2)
			
	End Select

End Sub


'========================================================================================
' Function Name : InitSpreadPosVariables
' Function Desc : 
'======================================================================================== 
Sub InitSpreadPosVariables()
	
	C_Date = 1
	C_Remark = 2
	
End Sub


'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : �׸��� ������ 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)

   ggoSpread.Source = frm1.vspdData
   
   Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
   
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


'========================================================================================
' Function Name : vspdData_Click
' Function Desc : �׸��� ��� Ŭ���� ���� 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	Call SetPopupMenuItemInf("1001011111")
   
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
	End If
	
	'------ Developer Coding part (Start)
	'------ Developer Coding part (End)
	
	
End Sub


'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : �׸��� �����¸� �����Ѵ�.
'========================================================================================
Sub PopSaveSpreadColumnInf()
	
   ggoSpread.Source = gActiveSpdSheet
   Call ggoSpread.SaveSpreadColumnInf()
   
End Sub 


'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : �׸��� �����¸� �����Ѵ�.
'========================================================================================
Sub PopRestoreSpreadColumnInf()
	
   ggoSpread.Source = gActiveSpdSheet
   
   Call ggoSpread.RestoreSpreadInf()
   Call InitSpreadSheet
   Call ggoSpread.ReOrderingSpreadData
   
End Sub 



'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

	
   ggoSpread.Source = frm1.vspdData
   Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
   Call GetSpreadColumnPos("A")
   
End Sub 


'====================== 2.2.4 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadLock()

    ggoSpread.SpreadLock -1, -1

End Sub

'=========================== 2.2.5 SetSpreadColor() =====================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
		
    With frm1
    
	    .vspdData.ReDraw = False
		ggoSpread.SSSetRequired	C_Date, pvStartRow, pvEndRow
		.vspdData.ReDraw = True
    
    End With
End Sub

'------------------------------------------  OpenCalType()  ----------------------------------------------
'	Name : OpenCalType()
'	Description : Calendar Type Popup
'---------------------------------------------------------------------------------------------------------
Function OpenCalType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "Į���� Ÿ�� �˾�"			' �˾� ��Ī 
	arrParam(1) = "P_MFG_CALENDAR_TYPE"				' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtClnrType.Value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "Į���� Ÿ��"					' TextBox ��Ī 
	
    arrField(0) = "CAL_TYPE"						' Field��(0)
    arrField(1) = "CAL_TYPE_NM"						' Field��(1)
    
    arrHeader(0) = "Į���� Ÿ��"				' Header��(0)
    arrHeader(1) = "Į���� Ÿ�Ը�"				' Header��(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetCalType(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtClnrType.focus
    
End Function

'------------------------------------------  SetCalType()  -----------------------------------------------
'	Name : SetCalType()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetCalType(byval arrRet)
	frm1.txtClnrType.value = arrRet(0)
	frm1.txtClnrTypeNm.value = arrRet(1)
End Function

'#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'#########################################################################################################
'******************************************  3.1 Window ó��  ********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'*********************************************************************************************************
Sub chkSat_OnClick()
	If frm1.chkSat.checked = true  then
		Call ggoOper.SetReqAttr(frm1.ChkFirstWeek,"Q") 	
		Call ggoOper.SetReqAttr(frm1.ChkSecondWeek,"Q") 	
		Call ggoOper.SetReqAttr(frm1.ChkThirdWeek,"Q") 	
		Call ggoOper.SetReqAttr(frm1.ChkForthWeek,"Q") 	
		Call ggoOper.SetReqAttr(frm1.ChkFifthWeek,"Q") 	
	Else
		Call ggoOper.SetReqAttr(frm1.ChkFirstWeek,"D")
		Call ggoOper.SetReqAttr(frm1.ChkSecondWeek,"D") 	
		Call ggoOper.SetReqAttr(frm1.ChkThirdWeek,"D") 	
		Call ggoOper.SetReqAttr(frm1.ChkForthWeek,"D") 	
		Call ggoOper.SetReqAttr(frm1.ChkFifthWeek,"D") 	
		 	
	End if
End Sub 

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================

Sub vspdData_Change(ByVal Col , ByVal Row )
	Dim strDate
	Dim strYear1
	Dim strMonth1
	Dim strDay1
	
	
    ggoSpread.Source = frm1.vspdData
    
	With frm1.vspdData
	
    If Col = C_Date Then
        .Col = Col
        .Row = Row
        If Trim(.Text) <> "" Then
			Call ExtractDateFrom(.Text,Parent.gDateFormat,Parent.gComDateType,strYear1,strMonth1,strDay1)
			strDate = .Text
			
			If IsDate(strDate) Then
				If strYear1 < Trim(frm1.txtYear.value) or strYear1 > Trim(frm1.txtYear.value) Then
					strYear1 = Trim(frm1.txtYear.value)

					strDate = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear1, strMonth1, strDay1)
					.Text = strDate
				End If
			Else
				If Not ggoSpread.SSDefaultCheck Then                                  '��: Check contents area
						   
				End If
			
			End If
			
			
        End If
    End If
    End With
End Sub

Sub txtYear_Change()
	
	Dim LngRow
	Dim fpMaxDate
	Dim fpMinDate
	Dim strYear1
	Dim strMonth1
	Dim strDay1
	
	If Len(Trim(frm1.txtYear.value)) <> 4 Then
		Exit Sub
	End If
	
	For LngRow = 1 To frm1.vspdData.MaxRows 
		frm1.vspdData.Col = C_Date
		frm1.vspdData.Row = LngRow
		
		If IsDate(frm1.vspdData.Text) Then
			Call ExtractDateFrom(frm1.vspdData.Text,Parent.gDateFormat,Parent.gComDateType,strYear1,strMonth1,strDay1)
			frm1.vspdData.Text = UniConvYYYYMMDDToDate(Parent.gDateFormat, Trim(frm1.txtYear.value), strMonth1, strDay1)
		End If
	Next 
			
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'======================================================================================================
Sub Form_Load()
	Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
	Call AppendNumberPlace("6","4","0")
	Call AppendNumberRange("6",lgMinYear, lgMaxYear)
	
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,FALSE,,ggStrMinPart,ggStrMaxPart)
    
    Call InitSpreadSheet                                                    '��: Setup the Spread sheet
    Call SetDefaultVal

    '----------  Coding part  -------------------------------------------------------------
    Call SetToolbar("10000100000011")										'��: ��ư ���� ���� 
	frm1.txtClnrType.focus
	Set gActiveElement = document.activeElement 
End Sub

'==========================================  2.2.1 SetDefaultVal()  ======================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
	
	frm1.txtYear.text = StrYear
	
End Sub

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Save Button of Main ToolBar
'========================================================================================

Function FncSave() 

	Dim IntRetCD
	Dim IntRows
	Dim IntRow
	Dim strVal
	
	lgBlnFlgChgValue = False 
	
	If frm1.txtClnrType.value = "" Then
		frm1.txtClnrTypeNm.value = ""
	End If
	
	If Not chkField(Document, "2") Then
		Exit Function
	End If
	
	With frm1.vspdData
	For IntRows = 1 To .MaxRows - 1
		
		.Row = IntRows
		.Col = C_Date
		strVal = Trim(.Text)
		
		For IntRow = IntRows +1 To .MaxRows 
			
			.Row = IntRow
			
			If strVal = Trim(.Text) Then
				Call DisplayMsgBox("970001","X","������¥","X")
				.focus
				.Row = IntRows 
				.Action = 0
				.SelStart = 0
				.SelLength = len(frm1.vspdData.Text)
				Exit Function
			End if

		Next
		 
	Next
	End With
	
	ggoSpread.Source = frm1.vspdData 
	
    If Not ggoSpread.SSDefaultCheck Then                                  '��: Check contents area
       Exit Function
    End If
	
	If CInt(frm1.txtYear.Value) < CInt(lgMinYear) Then
		Call DisplayMsgBox("970023","X","�����⵵",lgMinYear)
		frm1.txtYear.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	If CInt(frm1.txtYear.Value) > CInt(lgMaxYear) Then
		Call DisplayMsgBox("972004","X","�����⵵", lgMaxYear)
		frm1.txtYear.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If
	
	Call LookUpCal()
                                                     '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
	
	If frm1.vspdData.maxrows < 1 Then Exit Function
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo                                                  '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow(ByVal pvRowCnt)

	Dim IntRetCD
	Dim imRow
		
	'On Error Resume Next

	FncInsertRow = False

	If frm1.vspdData.maxrows >= 0 Then 
		Call SetToolbar("10000101000011")										'��: ��ư ���� ���� 
	End if

	If IsNumeric(Trim(pvRowCnt)) Then
		
		imRow = CInt(pvRowCnt)
		
	Else
		
		imRow= AskSpdSheetAddRowCount()
		
		If imRow = "" Then
			Exit Function
		End If
		
	End If
	
    With frm1
 
	    .vspdData.ReDraw = False
		.vspdData.focus
		ggoSpread.Source = .vspdData
		ggoSpread.InsertRow .vspdData.ActiveRow, imRow				
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow -1		
		.vspdData.ReDraw = True
		.vspdData.EditMode = True
	
	End With
	
    Set gActiveElement = document.activeElement 
	
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
    Call parent.FncExport(Parent.C_MULTI)												'��: ȭ�� ���� 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                         '��:ȭ�� ����, Tab ���� 
End Function
'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function

'========================================================================================
' Function Name : LookUpCal
' Function Desc : ������ Į���ٰ� �ִ� �� üũ�Ѵ�.
'========================================================================================
Function LookUpCal()
	
	If Len(frm1.txtYear.value) <> 4 Then
		Call DisplayMsgBox("971012","X", "�����⵵","X")
		Exit function
	End If

	Dim strVal
    
    LayerShowHide(1)
		
    With frm1
    
    strVal = BIZ_PGM_LOOKUP_ID & "?txtMode=" & Parent.UID_M0001						'��: 
    strVal = strVal & "&txtYear=" & Trim(.txtYear.value)						'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtClnrType=" & Trim(.txtClnrType.value)				'��: ��ȸ ���� ����Ÿ 
	
	Call RunMyBizASP(MyBizASP, strVal)											'��: �����Ͻ� ASP �� ���� 
        
    End With

End Function

Function ClnrNO()
	
		Call LayerShowHide(0)
		Call BtnDisabled(0)
	
	
End Function
'========================================================================================
' Function Name : LookUpCal
' Function Desc : �̹� ������ Į���ٰ� ���� ��� ���� ���� 
'========================================================================================
Function ClnrLookUpOk()
	Dim rtnVal
	
	rtnVal = DisplayMsgBox("188100",Parent.VB_YES_NO,"X","X")
	
	If rtnVal = vbYes Then
		Call DbExecute
	Else
		Call LayerShowHide(0)
		Call BtnDisabled(0)
	End If
	
	
End Function

'========================================================================================
' Function Name : DbExecute
' Function Desc : Į���� ���� ��ư�� ������ Į���ٸ� �����Ѵ�.
'========================================================================================

Function DbExecute()

	Dim strVal
    Dim IntRows 
    Dim IntCols 
    Dim vbIntRet 
    Dim lStartRow 
    Dim lEndRow 
    Dim boolCheck 
    Dim lGrpcnt 
    
        	   
    With frm1
		.txtMode.value = Parent.UID_M0002											'��: ���� ���� 
		.txtUpdtUserId.value  = Parent.gUsrID
	End With

    '-----------------------
    'Data manipulate area
    '-----------------------
    lGrpCnt = 1

	With frm1.vspdData
	    
    For IntRows = 1 To .MaxRows
    
		.Row = IntRows
		.Col = 0
	 
		If .Text = ggoSpread.InsertFlag Then
			strVal = strVal & "Sheet1" & Parent.gColSep & "C" & Parent.gColSep				'��: C=Create, Sheet�� 2�� �̹Ƿ� ���� 
		End If
	
		Select Case .Text
	    
		    Case ggoSpread.InsertFlag
	 
		        .Col = C_Date			'1
		        strVal = strVal & Trim(.Text) & Parent.gColSep
		            
		        .Col = C_Remark			'2
		        strVal = strVal & Trim(.Text) & Parent.gRowSep
	 
		        lGrpCnt = lGrpCnt + 1
		End Select

    Next

	End With
	
	frm1.txtMaxRows.value = lGrpCnt-1										'��: Spread Sheet�� ����� �ִ밹�� 
	frm1.txtSpread.value = strVal											'��: Spread Sheet ������ ���� 
   
    Call ExecMyBizASP(frm1, BIZ_PGM_BATCH_ID)										'��: �����Ͻ� ASP �� ���� 
	
End Function

'========================================================================================
' Function Name : DbExecOk
' Function Desc : Į���� ���� ��ư�� ������ Į���ٸ� �����Ѵ�.
'========================================================================================
Function DbExecOk()

	Call DisplayMsgBox("183114","X","X","X")
	
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	

<!--'==========================================  1.1.1 Style Sheet  ====================================
'=====================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB3" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD HEIGHT=5>&nbsp;</TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE CELLSPACING=0 CELLPADDING=0 WIDTH=100% border=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>����Į���ٻ���</font></td>
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
			<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>
				<TR>
					<TD HEIGHT=5 WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET>
							<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD CLASS=TD5 NOWRAP>Į���� Ÿ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtClnrType" SIZE=5 MAXLENGTH=2 tag="22XXXU" ALT="Į���� Ÿ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCalType()">&nbsp;<INPUT TYPE=TEXT NAME="txtClnrTypeNm" SIZE=30 tag="24"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�����⵵</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/p1102ma1_I658303737_txtYear.js'></script>								
									</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE CLASS="BasicTB" CELLSPACING=0>
							<TR>
								<TD HEIGHT=5 WIDTH=100%></TD>
							</TR>
							<TR>
								<TD WIDTH=100%  valign=top>
									<FIELDSET>
										<LEGEND>���ϼ���</LEGEND>
										<TABLE CLASS="TB2" CELLSPACING=0>
											<TR>
												<TD>
													<INPUT TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid" TAG="2X" NAME="chkSun" VALUE="Y" CHECKED=TRUE>Sun
													<INPUT TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid" TAG="2X" NAME="chkMon" VALUE="Y">Mon
													<INPUT TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid" TAG="2X" NAME="chkTue" VALUE="Y">Tue
													<INPUT TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid" TAG="2X" NAME="chkWed" VALUE="Y">Wed
													<INPUT TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid" TAG="2X" NAME="chkThu" VALUE="Y">Thu
													<INPUT TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid" TAG="2X" NAME="chkFri" VALUE="Y">Fri
													<INPUT TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid" TAG="2X" NAME="chkSat" VALUE="Y">Sat
												</TD>
											</TR>
										</TABLE>	
									</FIELDSET>	
								</TD>	
							</TR>
							<TR>
								<TD HEIGHT=5 WIDTH=100%></TD>
							</TR>	
							<TR>
								<TD WIDTH=100%  valign=top>
									<FIELDSET>
										<LEGEND>����ϼ���</LEGEND>
										<TABLE CLASS="TB2" CELLSPACING=0>
											<TR>
												<TD>
													<INPUT TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid" TAG="2X" NAME="ChkFirstWeek"  VALUE="Y">1��
													<INPUT TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid" TAG="2X" NAME="ChkSecondWeek" VALUE="Y">2��
													<INPUT TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid" TAG="2X" NAME="ChkThirdWeek"  VALUE="Y">3��
													<INPUT TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid" TAG="2X" NAME="ChkForthWeek"  VALUE="Y">4��
													<INPUT TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid" TAG="2X" NAME="ChkFifthWeek"  VALUE="Y">5��													
												</TD>
											</TR>
										</TABLE>	
									</FIELDSET>	
								</TD>	
							</TR>
							<TR>
								<TD HEIGHT=5 WIDTH=100%></TD>
							</TR>	
							<TR>
								<TD WIDTH=100% HEIGHT=100% valign=top>&nbsp;���ϼ���
									<TABLE WIDTH=100% HEIGHT=96%>
										<TR>
											<TD>
												<script language =javascript src='./js/p1102ma1_I483030766_vspdData.js'></script>
											</TD>
										</TR>
									</TABLE>
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
					<TD><BUTTON NAME="btnBatch" CLASS="CLSMBTN" Flag=1 ONCLICK=FncSave>����</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMajorFlg" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
