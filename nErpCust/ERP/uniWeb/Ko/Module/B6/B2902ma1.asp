
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Organization(Table Reflection등록)
'*  3. Program ID           : B2902ma1.asp
'*  4. Program Name         : B2902ma1.asp
'*  5. Program Desc         : 내부부서코드반영 Table등록 
'*  6. Modified date(First) : 2000/09/25
'*  7. Modified date(Last)  : 2000/09/25
'*  8. Modifier (First)     : Hwang Jeong-won
'*  9. Modifier (Last)      : Kim Jeong Min
'* 10. Comment              : 2002/11/29 : Include Slim, Grid UI Upgrade
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
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	<%'☜: indicates that All variables must be declared in advance%>

Const BIZ_PGM_ID = "B2902mb1.asp"												<%'비지니스 로직 ASP명 %>
 
Dim C_ModuleNm 
Dim C_ModuleCd
Dim C_Table
Dim C_TablePopUp
Dim C_UseFlag
Dim C_ChangeDt
Dim C_ChangeId
Dim C_SuccessFlag

Dim IsOpenPop          
Dim lgStrPrevKey2

<!-- #Include file="../../inc/lgvariables.inc" -->

Sub InitSpreadPosVariables()
     C_ModuleNm = 1
     C_ModuleCd = 2
     C_Table = 3
     C_TablePopUp = 4
     C_UseFlag = 5
     C_ChangeDt = 6
     C_ChangeId = 7
     C_SuccessFlag = 8
End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevKey2 = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","B","NOCOOKIE" ,"MA") %>
End Sub

Sub InitSpreadSheet()
	Call InitSpreadPosVariables()
	
	With frm1.vspdData
	
    .MaxCols = C_SuccessFlag + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
	
	.Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
    .ColHidden = True
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20021129",,parent.gAllowDragDropSpread

    .MaxRows = 0
    ggoSpread.ClearSpreadData
	call GetSpreadColumnPos("A")

	.ReDraw = false
    	
    ggoSpread.SSSetCombo C_ModuleNm, "업무", 18 '1
	ggoSpread.SSSetCombo C_ModuleCd, " ", 10 '2
	ggoSpread.SSSetEdit C_Table, "Table", 32,,,32,2 '3
    ggoSpread.SSSetButton C_TablePopUp '4
    ggoSpread.SSSetCheck C_UseFlag, "사용여부", 12, 2, "사용", True '5
    ggoSpread.SSSetDate C_ChangeDt, "부서개편일", 20, 2, parent.gDateFormat '6    
    ggoSpread.SSSetEdit C_ChangeId, "부서개편ID", 18, 2,,5 '7
    ggoSpread.SSSetCheck C_SuccessFlag, "반영성공여부", 24, 2, "성공", True '8
	
    Call ggoSpread.MakePairsColumn(C_Table,C_TablePopUp)
	Call ggoSpread.SSSetColHidden(C_ModuleCd,C_ModuleCd,true)
	
	.ReDraw = true

    Call SetSpreadLock
    
    End With
    
End Sub

Sub SetSpreadLock()
    With frm1
	    .vspdData.ReDraw = False
		ggoSpread.SpreadLock C_ModuleCd, -1, C_ModuleCd, -1
		ggoSpread.SpreadLock C_Table, -1, C_Table
		ggoSpread.SpreadLock C_TablePopup, -1, C_TablePopup
		ggoSpread.SpreadLock C_ChangeDt, -1, C_ChangeDt
		ggoSpread.SpreadLock C_ChangeId, -1, C_ChangeId
		ggoSpread.SpreadLock C_SuccessFlag, -1, C_SuccessFlag
		ggoSpread.SSSetProtected .vspdData.MaxCols, -1, -1
	    .vspdData.ReDraw = True
    End With
End Sub

Sub InitData()
	Dim intRow
	Dim intIndex 

	With frm1.vspdData
		For intRow = 1 To .MaxRows			
			.Row = intRow
			.Col = C_ModuleCd  :  intIndex = .value             ' .Value means that it is index of cell,not value in combo cell type
			.Col = C_ModuleNm  :  .value = intindex					
		Next	
	End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
		ggoSpread.SSSetRequired C_ModuleNm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Table, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ChangeDt, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ChangeId, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_SuccessFlag, pvStartRow, pvEndRow
    .vspdData.ReDraw = True
    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_ModuleNm = iCurColumnPos(1)
            C_ModuleCd = iCurColumnPos(2)
            C_Table = iCurColumnPos(3)
            C_TablePopUp = iCurColumnPos(4)
            C_UseFlag = iCurColumnPos(5)
            C_ChangeDt = iCurColumnPos(6)
            C_ChangeId = iCurColumnPos(7)
            C_SuccessFlag = iCurColumnPos(8)
    End Select    
End Sub

Sub InitSpreadComboBox()
    Dim strCboData,strCboData2
    Dim i
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("B0001", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	
    
    if lgF0 <> "" then
        strCboData = replace(lgF0,chr(11),vbTab)
        strCboData2 = replace(lgF1,chr(11),vbTab)
        strCboData = left(strCboData,len(strCboData) - 1)
        strCboData2 = left(strCboData2,len(strCboData2) - 1)

		ggoSpread.SetCombo strCboData, C_ModuleCd
		ggoSpread.SetCombo strCboData2, C_ModuleNm
	end if
End Sub

Sub initComboBox()
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("B0001", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	
    
    if frm1.cboModuleCd.length < 2 then     ' cboModuleCD Combo have one item when form load
        Call SetCombo2(frm1.cboModuleCd ,lgF0  ,lgF1  ,Chr(11))
    end if
End Sub

<% '----------------------------------------  OpenTable()  ------------------------------------------
'	Name : OpenTable()
'	Description : Table PopUp
'--------------------------------------------------------------------------------------------------------- %>
Function OpenTable(iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim lsOpenPop

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "Table 팝업"					' 팝업 명칭 
	arrParam(1) = "z_table_info"					' TABLE 명칭 
	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(2) = frm1.txtTable.value			' Code Condition
	Else 'spread
		arrParam(2) = frm1.vspdData.Text			' Code Condition
	End If
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = " lang_cd =  " & FilterVar(parent.gLang , "''", "S") & ""	' Where Condition
	arrParam(5) = "Table"					' 조건필드의 라벨 명칭 
	
    arrField(0) = "table_id"						' Field명(0)
    arrField(1) = "table_nm"						' Field명(1)
    arrField(2) = ""								' Field명(2)
    
    arrHeader(0) = "Table"							' Header명(0)
    arrHeader(1) = "Table명"					' Header명(1)
    arrHeader(2) = ""								' Header명(2)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtTable.focus
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetTable(arrRet, iWhere)
	End If	
			
End Function

<% '------------------------------------------  SetTable()  --------------------------------------
'	Name : SetTable()
'	Description : Table Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- %>
Function SetTable(Byval arrRet, Byval iWhere)
		
	With frm1
		If iWhere = 0 Then 'TextBox(Condition)
			.txtTable.value = arrRet(0)
		Else 'spread
			.vspdData.Col = C_Table
			.vspdData.Text = arrRet(0)
		End If
	End With
End Function

Sub Form_Load()

    Call LoadInfTB19029                                                     'Load table , B_numeric_format
    
    Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field                   

    Call InitSpreadSheet                                                    'Setup the Spread sheet
    Call InitVariables                                                      'Initializes local global variables
    
    Call InitSpreadComboBox
    Call initComboBox
    Call SetToolbar("1100110100101111")										'버튼 툴바 제어 
    frm1.cboModuleCd.focus
    
End Sub

Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101011111") 
    
    gMouseClickStatus = "SPC" 
    Set gActiveSpdSheet = frm1.vspdData

	if frm1.vspddata.MaxRows <= 0 then
		exit sub
	end if
	
	if Row <=0 then
		ggoSpread.Source = frm1.vspdData
		if lgSortkey = 1 then
			ggoSpread.SSSort Col
			lgSortKey = 2
		else
			ggoSpread.SSSort Col, lgSortkey
			lgSortKey = 1
		end if
		Exit sub
	end if
	frm1.vspdData.Row = Row
    
End Sub

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	With frm1.vspdData 
	
    ggoSpread.Source = frm1.vspdData
   
    If Row > 0 And Col = C_TablePopUp Then
		    .Row = Row
		    .Col = C_Table

		    Call OpenTable(1)        
    End If
    
    End With
End Sub

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim index
	
	With frm1.vspdData
		If Col = C_ModuleNm And Row > 0 Then
			.Row = Row
			.Col = Col
			index = .TypeComboBoxCurSel
			
			.Col = C_ModuleCd
			.TypeComboBoxCurSel = index		
		End If
	End With
	
	ggoSpread.UpdateRow Row
End Sub

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then					'☜: 조회중이면 다음 조회 안하도록 체크 
        Exit Sub
	End If
	
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    <%'☜: 재쿼리 체크 %>
    	If lgStrPrevKey <> "" Then                  <%'다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 %>
      		Call DisableToolBar(parent.TBC_QUERY)					'☜ : Query 버튼을 disable 시킴.
			If DBQuery = False Then 
			   Call RestoreToolBar()
			   Exit Sub 
			End If 
    	End If

    End if
    
End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    

Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               ' Protect system from crashing

    '-----------------------
    'Check previous data area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    
    '-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")									<%'Clear Contents  Field%>
    ggoSpread.ClearSpreadData
    Call InitVariables                                                      <%'Initializes local global variables%>
    															
    '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '----------------------- 
    If DbQuery = False Then Exit Function							<%'Query db data%>
       
    FncQuery = True															
    
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
	Call InitData()
End Sub

Function FncSave() 
    
    FncSave = False                                                         
    
    '-----------------------
    'Precheck area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        Call DisplayMsgBox("900001", "X", "X", "X")                           'No data changed!!
        Exit Function
    End If
    
    '-----------------------
    'Check content area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then     'Not chkField(Document, "2") OR    '⊙: Check contents area
       Exit Function
    End If
    
    '-----------------------
    'Save function call area
    '----------------------- 
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          
    
End Function

Function FncCopy() 
	With frm1.vspdData
		If .ActiveRow > 0 Then
			.focus
			.ReDraw = False
			
			ggoSpread.Source = frm1.vspdData 
			ggoSpread.CopyRow
			SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
    
			'Key field clear
			.Col = C_Table
			.Text = ""
			
			.Col = C_ChangeDt
			.Text = ""
			
			.Col = C_ChangeId
			.Text = ""
			
			.Col = C_SuccessFlag
			.Text = ""
			
			.ReDraw = True
		End If
	End With
End Function

Function FncCancel() 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
End Function

Function FncInsertRow(ByVal pvRowCnt) 
    Dim imRow
    Dim iSetRow
    
    if IsNumeric(Trim(pvRowCnt)) Then
        imRow = Cint(pvRowCnt)
    Else
        imRow = AskSpdSheetAddRowCount()
        if imRow = "" Then
            Exit Function
        End if
    End if
    
	With frm1
	
		.vspdData.focus
		ggoSpread.Source = .vspdData
		
		.vspdData.ReDraw = False
		ggoSpread.InsertRow .vspdData.ActiveRow,imRow
		SetSpreadColor .vspdData.ActiveRow,.vspdData.ActiveRow + imRow - 1
		
        For iSetRow = .vspdData.ActiveRow -1  to .vspdData.ActiveRow + imRow - 1
            .vspdData.Row = iSetRow
       	    .vspdData.Col = C_UseFlag
		    .vspdData.Value = 1
		Next
		
		.vspdData.ReDraw = True
    
    End With
End Function

Function FncDeleteRow() 
    Dim lDelRows

    With frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
    
End Function

Function FncPrint()
    Call parent.FncPrint()                                                   '☜: Protect system from crashing
End Function

Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)											 '☜: 화면 유형 
End Function

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                      '☜:화면 유형, Tab 유무 
End Function

Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")                ' 데이타가 변경되었습니다. 종료 하시겠습니까?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

Function DbQuery() 

    DbQuery = False

	Call LayerShowHide(1)
	
	Dim strVal
    
    With frm1
    
     If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 
		strVal = strVal & "&txtTable=" & .hTable.value 			'☆: 조회 조건 데이타 
        strVal = strVal & "&cboModuleCd=" & .hModuleCd.value 
        strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
        strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	 Else
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 
        strVal = strVal & "&txtTable=" & Trim(.txtTable.value)			'☆: 조회 조건 데이타 
        strVal = strVal & "&cboModuleCd=" & Trim(.cboModuleCd.value)
        strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
        strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
     End If
    
	Call RunMyBizASP(MyBizASP, strVal)										' ☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True
    
End Function

Function DbQueryOk()													'조회 성공후 실행로직 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE

	Call SetToolbar("1100111100111111")										' 버튼 툴바 제어 
	
End Function

Function DbSave() 
    Dim lRow        
    Dim lGrpCnt  
	Dim strVal, strDel
	Dim a, b
	
    DbSave = False                                                          
    
    Call LayerShowHide(1)

	With frm1
		.txtMode.value = parent.UID_M0002
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
    
    '-----------------------
    'Data manipulate area
    '----------------------- 
    lGrpCnt = 1
    
    strVal = ""
    strDel = ""
    
    '-----------------------
    'Data manipulate area
    '----------------------- 
    ' Data 연결 규칙 
    ' 0: Flag , 1: Row위치, 2~N: 각 데이타 

    For lRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = lRow
		    .vspdData.Col = 0

		    Select Case .vspdData.Text
		        Case ggoSpread.InsertFlag								' ☜: 신규 
					strVal = strVal & "C" & parent.gColSep	& lRow & parent.gColSep	'☜: C=Create
		        Case ggoSpread.UpdateFlag								' ☜: 수정 
					strVal = strVal & "U" & parent.gColSep	& lRow & parent.gColSep	'☜: U=Update
			End Select
			
		    Select Case .vspdData.Text

		        Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag			' ☜: 수정, 신규 
					
		            .vspdData.Col = C_Table	'1
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ModuleCd		'2
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_UseFlag	'1
	                If .vspdData.Value = 1 Then
			            strVal = strVal & "Y" & parent.gRowSep
			        Else
			            strVal = strVal & "N" & parent.gRowSep
					End If
		            
		            lGrpCnt = lGrpCnt + 1
		            
		        Case ggoSpread.DeleteFlag								' ☜: 삭제 

					strDel = strDel & "D" & parent.gColSep & lRow & parent.gColSep	' ☜: U=Update

		            .vspdData.Col = C_Table	'1
		            strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
		            
		            lGrpCnt = lGrpCnt + 1
		    End Select
                
    Next
	
	.txtMaxRows.value = lGrpCnt-1	
	.txtSpread.value = strDel & strVal
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										' ☜: 비지니스 ASP 를 가동 
	
	End With
	
    DbSave = True                                                           
    
End Function

Function DbSaveOk()													        ' 저장 성공후 실행 로직 
	Call InitVariables
	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
	
    Call MainQuery()
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
									<TD CLASS="TD5">업무</TD>
									<TD CLASS="TD6" COLSPAN=3><SELECT NAME="cboModuleCd" tag="11X" STYLE="WIDTH: 160px;"><OPTION value=""></OPTION></SELECT></TD>						
									<TD CLASS="TD5">Table</TD>
									<TD CLASS="TD6">
										<INPUT TYPE=TEXT NAME="txtTable" SIZE=30 MAXLENGTH=32 tag="11XXXU"  ALT="Table"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTable" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenTable(0)">
										<DIV  style="display:none;"><input type="text" ID="txtDummy" NAME="txtDummy" TITLE="txtDummy"></div></TD>
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
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="B2902mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hTable" tag="24"><INPUT TYPE=HIDDEN NAME="hModuleCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

