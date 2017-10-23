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

Option Explicit																	'☜: indicates that All variables must be declared in advance

Dim BaseDate
Dim strYear
Dim strMonth
DIm strDay
Dim lgMaxYear
DIm lgMinYear

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
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

		.MaxCols = C_Remark + 1													'☜: 최대 Columns의 항상 1개 증가시킴 
		.MaxRows = 0
		
		Call GetSpreadColumnPos("A")
	
		ggoSpread.SSSetDate 	C_Date,		"날짜", 11, 2, Parent.gDateFormat	
		ggoSpread.SSSetEdit 	C_Remark,	"비고", 42,,,20
	
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
' Function Desc : 그리드 폭조정 
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
' Function Desc : 그리드 헤더 클릭시 정렬 
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
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopSaveSpreadColumnInf()
	
   ggoSpread.Source = gActiveSpdSheet
   Call ggoSpread.SaveSpreadColumnInf()
   
End Sub 


'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopRestoreSpreadColumnInf()
	
   ggoSpread.Source = gActiveSpdSheet
   
   Call ggoSpread.RestoreSpreadInf()
   Call InitSpreadSheet
   Call ggoSpread.ReOrderingSpreadData
   
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

	arrParam(0) = "칼렌다 타입 팝업"			' 팝업 명칭 
	arrParam(1) = "P_MFG_CALENDAR_TYPE"				' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtClnrType.Value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "칼렌다 타입"					' TextBox 명칭 
	
    arrField(0) = "CAL_TYPE"						' Field명(0)
    arrField(1) = "CAL_TYPE_NM"						' Field명(1)
    
    arrHeader(0) = "칼렌다 타입"				' Header명(0)
    arrHeader(1) = "칼렌다 타입명"				' Header명(1)

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
'	Description : Plant Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetCalType(byval arrRet)
	frm1.txtClnrType.value = arrRet(0)
	frm1.txtClnrTypeNm.value = arrRet(1)
End Function

'#########################################################################################################
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'#########################################################################################################
'******************************************  3.1 Window 처리  ********************************************
'	Window에 발생 하는 모든 Even 처리	
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
				If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
						   
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
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'======================================================================================================
Sub Form_Load()
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call AppendNumberPlace("6","4","0")
	Call AppendNumberRange("6",lgMinYear, lgMaxYear)
	
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,FALSE,,ggStrMinPart,ggStrMaxPart)
    
    Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet
    Call SetDefaultVal

    '----------  Coding part  -------------------------------------------------------------
    Call SetToolbar("10000100000011")										'⊙: 버튼 툴바 제어 
	frm1.txtClnrType.focus
	Set gActiveElement = document.activeElement 
End Sub

'==========================================  2.2.1 SetDefaultVal()  ======================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
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
				Call DisplayMsgBox("970001","X","같은날짜","X")
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
	
    If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
       Exit Function
    End If
	
	If CInt(frm1.txtYear.Value) < CInt(lgMinYear) Then
		Call DisplayMsgBox("970023","X","생성년도",lgMinYear)
		frm1.txtYear.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	If CInt(frm1.txtYear.Value) > CInt(lgMaxYear) Then
		Call DisplayMsgBox("972004","X","생성년도", lgMaxYear)
		frm1.txtYear.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If
	
	Call LookUpCal()
                                                     '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
	
	If frm1.vspdData.maxrows < 1 Then Exit Function
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
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
		Call SetToolbar("10000101000011")										'⊙: 버튼 툴바 제어 
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
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function

'========================================================================================
' Function Name : LookUpCal
' Function Desc : 생성된 칼렌다가 있는 지 체크한다.
'========================================================================================
Function LookUpCal()
	
	If Len(frm1.txtYear.value) <> 4 Then
		Call DisplayMsgBox("971012","X", "생성년도","X")
		Exit function
	End If

	Dim strVal
    
    LayerShowHide(1)
		
    With frm1
    
    strVal = BIZ_PGM_LOOKUP_ID & "?txtMode=" & Parent.UID_M0001						'☜: 
    strVal = strVal & "&txtYear=" & Trim(.txtYear.value)						'☆: 조회 조건 데이타 
    strVal = strVal & "&txtClnrType=" & Trim(.txtClnrType.value)				'☆: 조회 조건 데이타 
	
	Call RunMyBizASP(MyBizASP, strVal)											'☜: 비지니스 ASP 를 가동 
        
    End With

End Function

Function ClnrNO()
	
		Call LayerShowHide(0)
		Call BtnDisabled(0)
	
	
End Function
'========================================================================================
' Function Name : LookUpCal
' Function Desc : 이미 생성된 칼렌다가 있을 경우 생성 여부 
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
' Function Desc : 칼렌다 생성 버튼을 누르면 칼렌다를 생성한다.
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
		.txtMode.value = Parent.UID_M0002											'☜: 저장 상태 
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
			strVal = strVal & "Sheet1" & Parent.gColSep & "C" & Parent.gColSep				'⊙: C=Create, Sheet가 2개 이므로 구별 
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
	
	frm1.txtMaxRows.value = lGrpCnt-1										'☜: Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread.value = strVal											'☜: Spread Sheet 내용을 저장 
   
    Call ExecMyBizASP(frm1, BIZ_PGM_BATCH_ID)										'☜: 비지니스 ASP 를 가동 
	
End Function

'========================================================================================
' Function Name : DbExecOk
' Function Desc : 칼렌다 생성 버튼을 누르면 칼렌다를 생성한다.
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>생산칼렌다생성</font></td>
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
									<TD CLASS=TD5 NOWRAP>칼렌다 타입</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtClnrType" SIZE=5 MAXLENGTH=2 tag="22XXXU" ALT="칼렌다 타입"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCalType()">&nbsp;<INPUT TYPE=TEXT NAME="txtClnrTypeNm" SIZE=30 tag="24"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>생성년도</TD>
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
										<LEGEND>요일선택</LEGEND>
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
										<LEGEND>토요일선택</LEGEND>
										<TABLE CLASS="TB2" CELLSPACING=0>
											<TR>
												<TD>
													<INPUT TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid" TAG="2X" NAME="ChkFirstWeek"  VALUE="Y">1주
													<INPUT TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid" TAG="2X" NAME="ChkSecondWeek" VALUE="Y">2주
													<INPUT TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid" TAG="2X" NAME="ChkThirdWeek"  VALUE="Y">3주
													<INPUT TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid" TAG="2X" NAME="ChkForthWeek"  VALUE="Y">4주
													<INPUT TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid" TAG="2X" NAME="ChkFifthWeek"  VALUE="Y">5주													
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
								<TD WIDTH=100% HEIGHT=100% valign=top>&nbsp;휴일선택
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
					<TD><BUTTON NAME="btnBatch" CLASS="CLSMBTN" Flag=1 ONCLICK=FncSave>실행</BUTTON></TD>
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
