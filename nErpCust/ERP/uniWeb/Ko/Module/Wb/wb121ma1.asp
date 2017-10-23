<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<%'**********************************************************************************************
'*
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 
'*  3. Program ID           : VB121MA1
'*  4. Program Name         : 사용할 법인선택 
'*  5. Program Desc         : 사용할 법인 
'*  6. Component List       : 
'*  7. Modified date(First) : 2004/12/28
'*  8. Modified date(Last)  : 2004/12/28
'*  9. Modifier (First)     : LSHSAT
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'********************************************************************************************** 
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->				<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우  -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '☆: 해당 위치에 따라 달라짐, 상대 경로  -->

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                                                '☜: indicates that All variables must be declared in advance 


'********************************************  1.2 Global 변수/상수 선언  *********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->

'============================================  1.2.1 Global 상수 선언  ====================================

Const BIZ_PGM_ID = "wb121mb1.asp"											 '☆: 비지니스 로직 ASP명 

Dim C_CO_CD
Dim C_C0_NM
Dim C_FISC_YEAR
Dim C_REP_TYPE
Dim C_REP_TYPE_NM
Dim C_FISC_START_DT
Dim C_FISC_END_DT
Dim C_REVISION_YM2

Dim IsOpenPop          
Dim lgStrPrevKey2
Dim lgCurrGrid, lgOldCol, lgOldRow , lgChgFlg

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
sub InitSpreadPosVariables()
	
	C_CO_CD			= 1
	C_C0_NM			= 2
	C_FISC_YEAR		= 3
	C_REP_TYPE		= 4
	C_REP_TYPE_NM	= 5
	C_FISC_START_DT = 6
	C_FISC_END_DT	= 7
	C_REVISION_YM2	= 8
end sub

'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
End Sub


'========================================================================================================= 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub



'------------------------------------------  OpenCalType()  -------------------------------------------------
'	Name :InitComboBox_Five()
'	Description : 
'------------------------------------------------------------------------------------------------------------
Sub InitComboBox()

End Sub

 
'==========================================  2.4.3 Set???()  ===============================================
'	Name : OpenCompanyInfo()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 

Function OpenCompanyInfo(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "법인 팝업"						' 팝업 명칭 
	arrParam(1) = "TB_COMPANY_HISTORY"						' TABLE 명칭 
	arrParam(2) = strCode 							' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = " USE_FLG='Y' "								' Where Condition
	arrParam(5) = "법인"

    arrField(0) = "Upper(CO_CD)"					' Field명(0)
    arrField(1) = "CO_NM"						' Field명(1)

    arrHeader(0) = "법인코드"						' Header명(0)
    arrHeader(1) = "법인명"						' Header명(1)

	arrRet = window.showModalDialog("wb101ra1.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtCO_CD.focus
	    Exit Function
	Else
		Call SetCompanyInfo(arrRet,iWhere)
	End If	

End Function



'------------------------------------------  SetItemInfo()  -------------------------------------------------
'	Name : SetCostInfo()
'	Description : Popup에서 Return되는 값 setting
'------------------------------------------------------------------------------------------------------------
Function SetCompanyInfo(Byval arrRet,byval iWhere)'

	With frm1
		If iWhere = 0 Then
			.txtCO_CD.focus
			.txtCO_CD.value     = arrRet(0)
			.txtCO_FULLNM.value = arrRet(1)
			
			Call FncQuery
		End If
'		lgBlnFlgChgValue = False
	End With

End Function


Sub txtCO_CD_onChange()	' 법인코드 변경시 
	Dim arrVal
	
	If Len(frm1.txtCO_CD.Value) > 0 Then
		If CommonQueryRs("DISTINCT CO_NM", " TB_COMPANY_HISTORY " , " CO_CD = '" & frm1.txtCO_CD.Value &"' ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
	    	arrVal				= Split(lgF0, Chr(11))
			frm1.txtCO_FULLNM.Value	= arrVal(0)
		Else
			Call DisplayMsgBox("970000", "x",frm1.txtCO_CD.alt,"x")
			frm1.txtCO_FULLNM.Value	= ""
		End If
	Else
		frm1.txtCO_FULLNM.Value = ""
	End If

End Sub



'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	'frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	'frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	'frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	'frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"
	frm1.txtCO_CD.focus
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

			C_CO_CD			= iCurColumnPos(1)
			C_C0_NM			= iCurColumnPos(2)
			C_FISC_YEAR		= iCurColumnPos(3)
			C_REP_TYPE		= iCurColumnPos(4)
			C_REP_TYPE_NM	= iCurColumnPos(5)
			C_FISC_START_DT = iCurColumnPos(6)
			C_FISC_END_DT	= iCurColumnPos(7)
			C_REVISION_YM2	= iCurColumnPos(8)
    End Select    
End Sub
'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()  

	With frm1.vspdData
		
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread    

		.ReDraw = false
		.MaxCols   = C_REVISION_YM2 + 1                                          ' ☜:☜: Add 1 to Maxcols
		.Col       = .MaxCols                                                        ' ☜:☜: Hide maxcols
		.ColHidden = True                                                            ' ☜:☜:
		.MaxRows = 0

		ggoSpread.Source = Frm1.vspdData    
		ggoSpread.ClearSpreadData 

		Call GetSpreadColumnPos("A")  

		ggoSpread.SSSetEdit     C_CO_CD,				"법인코드",	 10 ,			20
		ggoSpread.SSSetEdit     C_C0_NM,				"법인명",	 20	,			100
		ggoSpread.SSSetEdit     C_FISC_YEAR,			"사업연도",	 10	,			2
		ggoSpread.SSSetEdit     C_REP_TYPE,				"신고구분",	 10	,			2
		ggoSpread.SSSetEdit     C_REP_TYPE_NM,			"신고구분",	 10	,			2
		ggoSpread.SSSetDate     C_FISC_START_DT,		"회기시작일", 10,		2,		Parent.gDateFormat,	-1
		ggoSpread.SSSetDate     C_FISC_END_DT,			"회기종료일", 10,		2,		Parent.gDateFormat,	-1
		ggoSpread.SSSetEdit     C_REVISION_YM2,			"개정판연월", 10,			2

		'Call ggoSpread.SSSetColHidden(C_REVISION_YM2,C_REVISION_YM2,True)	
		Call ggoSpread.SSSetColHidden(C_REP_TYPE,C_REP_TYPE,True)	

		Call SetSpreadLock 
		
		.OperationMode = 5
		.ReDraw = true
	    
	End With
    
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
      ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    .vspdData.ReDraw = True
    
    End With
End Sub

'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr, parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <>  parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub


'========================================================================================================= 
Sub Form_Load()
    Call InitVariables																'⊙: Initializes local global variables
    Call LoadInfTB19029																'⊙: Load table , B_numeric_format
	
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
    
    '----------  Coding part  -------------------------------------------------------------
	Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitComboBox

    Call SetToolBar("1100000000000111")

    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

	Call InitData
	
    lgBlnFlgChgValue = False 
    frm1.txtCO_CD.focus 

'	FncQuery

End Sub

'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub


'========================================================================================
Function FncQuery() 
    Dim IntRetCD

    FncQuery = False
    Err.Clear

  '-----------------------
    'Check previous data area
    '----------------------- 
'    If lgBlnFlgChgValue = True Then
'		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")				'⊙: "Will you destory previous data"
'		If IntRetCD = vbNo Then
'			Exit Function
'		End If
'    End If

  '-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call InitVariables															'⊙: Initializes local global variables

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
		Call RestoreToolBar()
        Exit Function
    End If

    FncQuery = True
End Function


'========================================================================================
Function FncNew() 
    Dim IntRetCD 

    FncNew = True

	Call InitData
End Function


'========================================================================================
Function FncDelete() 
    Dim IntRetCD

    FncDelete = True
End Function


'========================================================================================
Function FncSave() 
    Dim iRow, iMaxRows

	FncSave = False
	Err.Clear

    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If
    	
	With frm1.vspdData
		iMaxRows = .MaxRows
		If iMaxRows > 0 Then
			For iRow = 1 To iMaxRows
				.Row = iRow
				If .SelModeSelected Then
					.Col = C_REVISION_YM2 
					If .Text <> C_REVISION_YM Then
						'Call DisplayMsgBox("900002",parent.VB_INFORMATION, C_REVISION_YM, .Text)
					End If
				End If
			Next
		End If
	End With
	'-----------------------
	'Save function call area
	'-----------------------
	IF  DbSave	= False then
		Exit Function
	End If

	FncSave = True
End Function


'========================================================================================
Function FncCopy() 
   
End Function


'========================================================================================
Function FncCancel()
     On Error Resume Next
End Function


'========================================================================================
Function FncInsertRow()
     On Error Resume Next
End Function


'========================================================================================
Function FncDeleteRow()
     On Error Resume Next
End Function


'========================================================================================
Function FncPrint()
     On Error Resume Next
    'parent.FncPrint()
End Function


'========================================================================================
Function FncPrev()
End Function


'========================================================================================
Function FncNext()

End Function


'========================================================================================
Function FncExcel()
    Call parent.FncExport(parent.C_SINGLE)												'☜: 화면 유형 
End Function


'========================================================================================
Function FncFind()
    Call parent.FncFind(parent.C_SINGLE, False)                                         '☜:화면 유형, Tab 유무 
End Function


'========================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = True
End Function

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
	Call InitData()
End Sub


'=======================================================================================================
'   Event Name : 
'   Event Desc : 달력을 호출한다.
'=======================================================================================================

Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
    End If
End Sub




'=======================================================================================================
'   Event Name : Keypress(Key)
'   Event Desc : 3rd party control에서 Enter 키를 누르면 조회 실행 
'=======================================================================================================

Sub txtFISC_YEAR_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub
'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
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
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
		Exit Sub
	End if
	
	If frm1.vspdData.MaxRows = 0 then
		Exit Sub
	End if
	
	Call FncSave
End Sub
'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub


'========================================================================================
Function DbDelete() 
    Err.Clear

    DbDelete = True
End Function


'========================================================================================
Function DbDeleteOk()

End Function


'========================================================================================
Function DbQuery()

    Err.Clear

    DbQuery = False
    Call LayerShowHide(1)
    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtCO_Cd=" & Trim(frm1.txtCo_Cd.value)				'☆: 조회 조건 데이타 
    strVal = strVal & "&txtFISC_YEAR=" & Trim(frm1.txtFISC_YEAR.text)				'☆: 조회 조건 데이타 

	Call RunMyBizASP(MyBizASP, strVal)

    DbQuery = True

End Function

'========================================================================================
Function DbQueryOk()
	lgIntFlgMode      =  parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode

	'Call SetToolbar("1100000000011111")												'⊙: Set ToolBar
    'Call InitData()
    'Call  ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   

	With frm1.vspdData
	
		If .MaxRows > 0 Then
			.Row = 1
			.SelModeSelected = True
		End If
		.focus
	End With
	
End Function

'========================================================================================
Function DbSave() 

    Err.Clear
	DbSave = False

    Dim strVal

    Call LayerShowHide(1) 

	Dim iLngSelectedRows, lMaxRows, lMaxCols, iRow, lCol

	With frm1.vspdData
		iLngSelectedRows = .SelModeSelCount
		lMaxRows = .MaxRows : lMaxCols = .MaxCols
		
		If iLngSelectedRows > 0 Then 
			For iRow = 1 To lMaxRows
				.Row = iRow
				If .SelModeSelected Then
					strVal = "U" &  Parent.gColSep
					For lCol = 1 To lMaxCols
						.Col = lCol : strVal = strVal & Trim(.Value) &  Parent.gColSep
					Next
					Exit For
				End IF
			Next
		End if			
	End With

	With frm1
		.txtSpread.value	= strVal
		.txtMode.value		= parent.UID_M0002
		.txtFlgMode.value   = lgIntFlgMode

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	End With

    DbSave = True
End Function

'========================================================================================
Function DbSaveOk()

    lgBlnFlgChgValue = False
    Call DisplayMsgBox("183114", "X", "X" , "X")
End Function

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)

    Dim iDx
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

   	If Frm1.vspdData.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData.text) <  UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	 ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================

Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")       

    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
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
    
   	If lgOldRow <> Row Then
		
		frm1.vspdData.Col = 1
		frm1.vspdData.Row = row
	
		lgOldRow = Row
		  		
	End If
       frm1.vspdData.Row = Row
End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

       If Button = 2 And  gMouseClickStatus = "SPC" Then
           gMouseClickStatus = "SPCR"
        End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : This function is called when cursor leave cell
'========================================================================================================
Sub vspdData_ScriptLeaveCell(Col,Row,NewCol,NewRow,Cancel)
	If NewRow <= 0 Or NewCol < 0 Then
		Exit Sub
	End If
	
		frm1.vspdData.Col = 1
		frm1.vspdData.Row = NewRow
	
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
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
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP_BAK">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP_BAK"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>법인</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtCO_CD" MAXLENGTH="10" SIZE=10 ALT ="법인코드" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenCompanyInfo(frm1.txtco_cd.value,0)"> <INPUT NAME="txtCO_FULLNM" MAXLENGTH="30" SIZE=30 ALT ="법인명" tag="14X"></TD>
									<TD CLASS="TD5" NOWRAP>사업연도</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/wb121ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% VALIGN=TOP>
                        <TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" NOWRAP>
									<script language =javascript src='./js/wb121ma1_vaSpread1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=* WIDTH=100% VALIGN=TOP>
						<TABLE>
							<TR>
								<TD WIDTH=10>&nbsp;</TD>
								<TD><BUTTON NAME="btn1"  CLASS="CLSSBTN" ONCLICK="vbscript:FncSave()" Flag=1>법인 선택</BUTTON>&nbsp;
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
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 tabindex="-1"></IFRAME>		
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" style="display: 'none'"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>

