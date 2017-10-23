<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Item Account)
'*  3. Program ID           : b1b06ma1.asp
'*  4. Program Name         : Item Account Control
'*  5. Program Desc         :
'*  6. Modified date(First) : 2004/11/26
'*  7. Modified date(Last)  : 2004/11/26
'*  8. Modifier (First)     : Chen, Jae Hyun
'*  9. Modifier (Last)      : Chen, Jae Hyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
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
<SCRIPT LANGUAGE="VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	<%'☜: indicates that All variables must be declared in advance%>

Const BIZ_PGM_QRY_ID = "B1B06MB1.asp"												'비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID = "B1B06MB2.asp"												'비지니스 로직 ASP명 

Dim C_ItemAcct
Dim C_ItemAcctNm
Dim C_RepItemAcct
Dim C_RepItemAcctNm

Dim IsOpenPop          
Dim lsConcd

<!-- #Include file="../../inc/lgvariables.inc" -->

Sub initSpreadPosVariables()  
    C_ItemAcct  = 1
    C_ItemAcctNm  = 2
    C_RepItemAcct	= 3
    C_RepItemAcctNm	= 4
End Sub

Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
End Sub

'==========================================  2.2.6 InitData()  ========================================== 
'	Name : InitData()
'	Description : Combo Display
'======================================================================================================== 
Sub InitData(ByVal lngStartRow)
	Dim intRow
	Dim intIndex

	With frm1.vspdData
		For intRow = lngStartRow To .MaxRows
			
			.Row = intRow
			.col = C_ItemAcct
			intIndex = .value
			.Col = C_ItemAcctNm
			.value = intindex
			
			.Row = intRow
			.col = C_RepItemAcct
			intIndex = .value
			.Col = C_RepItemAcctNm
			.value = intindex
		Next	
	End With
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

Sub InitSpreadSheet()
    Call initSpreadPosVariables()  

	With frm1.vspdData
    ggoSpread.Source = frm1.vspdData

    'patch version
    ggoSpread.Spreadinit "V20041120",,parent.gAllowDragDropSpread    
    
	.ReDraw = false
	
    .MaxCols = C_RepItemAcctNm + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
	.Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
    .ColHidden = True
	
    .MaxRows = 0
    ggoSpread.ClearSpreadData
	
    Call GetSpreadColumnPos("A")    
	
	ggoSpread.SSSetCombo C_ItemAcct  , "품목계정"			,15
	ggoSpread.SSSetCombo C_ItemAcctNm  , "품목계정명"        ,35
	ggoSpread.SSSetCombo C_RepItemAcct  , "품목계정그룹"     ,15
	ggoSpread.SSSetCombo C_RepItemAcctNm  , "품목계정그룹명"   ,35
	
	Call ggoSpread.MakePairsColumn(C_ItemAcct, C_ItemAcctNm)
	Call ggoSpread.MakePairsColumn(C_RepItemAcct, C_RepItemAcctNm)
	Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
    
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub

Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SpreadLock    C_ItemAcct,  -1, C_ItemAcct
    ggoSpread.SpreadLock    C_ItemAcctNm,  -1, C_ItemAcctNm
    ggoSpread.SSSetRequired	C_RepItemAcct,  -1, -1
    ggoSpread.SSSetRequired	C_RepItemAcctNm,  -1, -1
    .vspdData.ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    Dim iRow
    
    With frm1
    
    .vspdData.ReDraw = False
    
    ggoSpread.SSSetRequired		C_ItemAcct,  pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_ItemAcctNm,  pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_RepItemAcct,  pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_RepItemAcctNm,  pvStartRow, pvEndRow
    
    .vspdData.ReDraw = True
    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_ItemAcct      = iCurColumnPos(1)
            C_ItemAcctNm      = iCurColumnPos(2)
            C_RepItemAcct      = iCurColumnPos(3)
            C_RepItemAcctNm    = iCurColumnPos(4)
    End Select    
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & Filtervar("P1001","''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboItemAcct, lgF0, lgF1, Chr(11))

	frm1.cboItemAcct.value = ""
    
End Sub


'========================== 2.2.6 InitSpreadComboBox()  =====================================
'	Name : InitSpreadComboBox()
'	Description : Combo Display
'=========================================================================================
Sub InitSpreadComboBox()
	
	'****************************
	'List Minor code(Item Acct)
	'****************************
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & Filtervar("P1001","''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_ItemAcct
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_ItemAcctNm
    
    
    '****************************
	'List Minor code(Item Acct)
	'****************************
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & Filtervar("P1000","''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_RepItemAcct
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_RepItemAcctNm
    
End Sub

Sub Form_Load()


    Call LoadInfTB19029                                                     <%'Load table , B_numeric_format%>
    
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                                                                                      <%'Format Numeric Contents Field%>                                                                            
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call InitComboBox()
    Call InitSpreadComboBox()
    
    Call SetToolBar("1100110100101111")										<%'버튼 툴바 제어 %>

    
    frm1.cboItemAcct.focus
    
End Sub

Sub vspdData_ButtonClicked(Col, Row, ButtonDown)

	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
    
	End With
End Sub

Sub vspdData_Change(ByVal Col , ByVal Row )
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'==========================================================================================
'   Event Name :vspddata_ComboSelChange                                                          
'   Event Desc :Combo Change Event                                                                           
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	With frm1.vspdData
	
		.Row = Row
		Select Case Col

			Case  C_ItemAcct
				.Col = Col
				intIndex = Trim(.Value)
				.Col = C_ItemAcctNm
				.Value = intIndex
				
				ggoSpread.Source = frm1.vspdData
				ggoSpread.UpdateRow Row
			
			Case C_ItemAcctNm	
				.Col = Col
				intIndex = .Value
				.Col = C_ItemAcct
				.Value = intIndex
				
				ggoSpread.Source = frm1.vspdData
				ggoSpread.UpdateRow Row
				
			Case  C_RepItemAcct
				.Col = Col
				intIndex = .Value
				.Col = C_RepItemAcctNm
				.Value = intIndex
				
				ggoSpread.Source = frm1.vspdData
				ggoSpread.UpdateRow Row
			
			Case C_RepItemAcctNm	
				.Col = Col
				intIndex = .Value
				.Col = C_RepItemAcct
				.Value = intIndex
				
				ggoSpread.Source = frm1.vspdData
				ggoSpread.UpdateRow Row
					
		End Select		
				
	End with			

End Sub

Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    Else
    	frm1.vspdData.Row = Row
		frm1.vspdData.Col = C_ItemAcct		
    End If

	frm1.vspdData.Row = Row
End Sub

Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
        Exit Sub
    End If

End Sub

Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    

Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
      	   Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreTooBar()
			    Exit Sub
			End If  				
    	End If
    End if
End Sub

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               <%'Protect system from crashing%>

	'-----------------------
    'Check previous data area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    
	'-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")									'Clear Contents  Field
    ggoSpread.ClearSpreadData
    Call InitVariables                                                      'Initializes local global variables
    															
	'-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then										'This function check indispensable field
       Exit Function
    End If
    
	'-----------------------
    'Query function call area
    '----------------------- 
    Call DisableToolBar(Parent.TBC_QUERY)
    If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function
    End If
       
    FncQuery = True															
    
End Function

Function FncSave() 
    
    FncSave = False                                                         
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
        Exit Function
    End If
    
<%  '-----------------------
    'Check content area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then     'Not chkField(Document, "2") OR     '⊙: Check contents area
       Exit Function
    End If
    
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          
    
End Function

Function FncCopy()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

End Function

Function FncCancel() 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo 
    Call InitData(1)
End Function

Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG
    
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

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
End Function

Function FncDeleteRow() 
    With frm1.vspdData
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
End Function

Function FncPrint()
    Call parent.FncPrint()                                                   '☜: Protect system from crashing
End Function

Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)											 <%'☜: 화면 유형 %>
End Function

Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                      <%'☜:화면 유형, Tab 유무 %>
End Function

Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 종료 하시겠습니까?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

Function FncSplitColumn()
    
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    iColumnLimit  =  5                 ' split 한계치  maxcol이 아님(5번째 칼럼이 split의 최고치)
                                       ' 5라는 값은 표준이 아닙니다.개발자가 업무에 맞게 수정요 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    
    ACol = Frm1.vspdData.ActiveCol
    ARow = Frm1.vspdData.ActiveRow

    If ACol > iColumnLimit Then
       iRet = DisplayMsgBox("900030", "X", iColumnLimit , "X")
       Exit Function  
    End If   
    
    Frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_NONE
    
    ggoSpread.Source = Frm1.vspdData
    ggoSpread.SSSetSplit(ACol)    
    Frm1.vspdData.Col = ACol
    Frm1.vspdData.Row = ARow
    
    Frm1.vspdData.Action = Parent.SS_ACTION_ACTIVE_CELL   
    
    Frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_BOTH
    
End Function

Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call InitSpreadComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData(1)
End Sub

Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>

	Call LayerShowHide(1)
	
	Dim strVal
    
    With frm1
       
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001						<%'현재 검색조건으로 Query%>
		strVal = strVal & "&txtItemAcct=" & .cboItemAcct.value				
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows

    
		Call RunMyBizASP(MyBizASP, strVal)										<%'☜: 비지니스 ASP 를 가동 %>
        
    End With
    
    DbQuery = True
    
End Function

Function DbQueryOk(ByVal LngMaxRow)													'조회 성공후 실행로직 
	
    '-----------------------
    'Reset variables area
    '-----------------------
	lgIntFlgMode = Parent.OPMD_UMODE
    Call InitData(LngMaxRow)
	Call SetToolBar("110011110001111")										<%'버튼 툴바 제어 %>
	
End Function

Function DbSave() 															<%'As New P21011ManageIndReqSvr%>
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal, strDel
	
    DbSave = False                                                          
    
    Call LayerShowHide(1)
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>

	With frm1
		.txtMode.value = Parent.UID_M0002

	'-----------------------
    'Data manipulate area
    '-----------------------
    lGrpCnt = 1
    
    strVal = ""
    strDel = ""

    For lRow = 1 To .vspdData.MaxRows
    
        .vspdData.Row = lRow
        .vspdData.Col = 0
        
        Select Case .vspdData.Text

            Case ggoSpread.InsertFlag									    <%'☜: 신규 %>
				
				strVal = strVal & "CREATE" & Parent.gColSep					  			<%'☜: C=Create, Row위치 정보 %>

                .vspdData.Col = C_ItemAcct	'1
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                
                .vspdData.Col = C_RepItemAcct	'2
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                
                strVal = strVal & lRow & Parent.gRowSep
                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.UpdateFlag
                
                strVal = strVal & "UPDATE" & Parent.gColSep								<%'☜: C=Create, Row위치 정보 %>

                .vspdData.Col = C_ItemAcct	'1
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                
                .vspdData.Col = C_RepItemAcct	'2
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

                strVal = strVal & lRow & Parent.gRowSep
                
                lGrpCnt = lGrpCnt + 1
                
                
            Case ggoSpread.DeleteFlag										<%'☜: 삭제 %>

				strDel = strDel & "DELETE" & Parent.gColSep								<%'☜: D=Update, Row위치 정보 %>
				
               .vspdData.Col = C_ItemAcct	'1
                strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
                
                .vspdData.Col = C_RepItemAcct	'2
                strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep

                strDel = strDel & lRow & Parent.gRowSep
                
                lGrpCnt = lGrpCnt + 1
                
        End Select
                
    Next
	
	.txtMaxRows.value = lGrpCnt-1	
	.txtSpread.value = strDel & strVal
	
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										<%'☜: 비지니스 ASP 를 가동 %>
	
	End With
	
    DbSave = True                                                           
    
End Function

Function DbSaveOk()													        <%' 저장 성공후 실행 로직 %>
	Call InitVariables
	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call MainQuery()
End Function

'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(lRow, lCol)
	frm1.vspdData.focus
	frm1.vspdData.Row = lRow
	frm1.vspdData.Col = lCol
	frm1.vspdData.Action = 0
	frm1.vspdData.SelStart = 0
	frm1.vspdData.SelLength = len(frm1.vspdData.Text)
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2KCM.inc" -->	
</HEAD>

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>품목계정설정</font></td>
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
									<TD CLASS="TD5">품목계정</TD>
									<TD CLASS="TD656">
										<SELECT NAME="cboItemAcct" ALT="품목계정" STYLE="Width: 98px;" tag="11"><OPTION VALUE=""></OPTION></SELECT>
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
								<TD HEIGHT="100%">
									<script language =javascript src='./js/b1b06ma1_I369410997_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24">
</TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

