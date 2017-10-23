
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Translation of Unit)
'*  3. Program ID           : B1f01ma1.asp
'*  4. Program Name         : B1f01ma1.asp
'*  5. Program Desc         :
'*  6. Modified date(First) : 2000/09/08
'*  7. Modified date(Last)  : 2002/06/21
'*  8. Modifier (First)     : Hwang Jeong-won
'*  9. Modifier (Last)      : Sim Hae Young
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
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit								

Const BIZ_PGM_ID = "B1f01mb1.asp"												<%'비지니스 로직 ASP명 %>
 
Dim C_Dimension
Dim C_DimensionCd
Dim C_Unit
Dim C_UnitPopUp
Dim C_Col1
Dim C_ToUnit
Dim C_ToUnitPopUp
Dim C_Col2
Dim C_Factor
Dim C_Col3
Dim C_ToFactor

Dim IsOpenPop          
Dim lgStrPrevKey2, lgStrPrevKey3

<!-- #Include file="../../inc/lgvariables.inc" -->

Sub InitSpreadPosVariables()
    C_Dimension     = 1
    C_DimensionCd   = 2
    C_Unit          = 3
    C_UnitPopUp     = 4
    C_Col1          = 5
    C_ToUnit        = 6
    C_ToUnitPopUp   = 7
    C_Col2          = 8
    C_Factor        = 9
    C_Col3          = 10
    C_ToFactor      = 11
End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevKey2 = ""
    lgStrPrevKey3 = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
End Sub

Sub LoadInfTB19029()
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
    <% Call loadInfTB19029A("I", "B","NOCOOKIE","MA") %>
End Sub

Sub InitSpreadSheet()
    Call initSpreadPosVariables()  

	With frm1.vspdData
	
	ggoSpread.Source = frm1.vspdData	
   'patch version
    ggoSpread.Spreadinit "V20021202",,parent.gAllowDragDropSpread    
    
	.ReDraw = false

    .MaxCols = C_ToFactor + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
	.Col = .MaxCols														'☆: 사용자 별 Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData
	
    Call GetSpreadColumnPos("A")  

	ggoSpread.SSSetCombo C_Dimension, "Dimension", 22
    ggoSpread.SSSetCombo C_DimensionCd, "Dimension코드", 23
    ggoSpread.SSSetEdit C_Unit, "기준단위", 20,,,3,2
    ggoSpread.SSSetButton C_UnitPopUp
    ggoSpread.SSSetEdit C_Col1, ":", 3, 2
    ggoSpread.SSSetEdit C_ToUnit, "변환단위", 20,,,3,2
    ggoSpread.SSSetButton C_ToUnitPopUp
    ggoSpread.SSSetEdit C_Col2, "=", 3, 2
    
    ggoSpread.SSSetFloat C_Factor,"기준계수",20,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetEdit C_Col3, ":", 3, 2
	ggoSpread.SSSetFloat C_ToFactor,"변환계수",20,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	
	call ggoSpread.MakePairsColumn(C_Unit,C_UnitPopUp)
	call ggoSpread.MakePairsColumn(C_ToUnit,C_ToUnitPopUp)

	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub

Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False

    ggoSpread.SpreadLock C_Dimension, -1, C_Dimension
    ggoSpread.SpreadLock C_DimensionCd, -1, C_DimensionCd
    ggoSpread.SpreadLock C_Unit, -1, C_Unit
    ggoSpread.SpreadLock C_UnitPopUp, -1, C_UnitPopUp
    ggoSpread.SpreadLock C_Col1, -1, C_Col1
    ggoSpread.SpreadLock C_ToUnit, -1, C_ToUnit
    ggoSpread.SpreadLock C_ToUnitPopUp, -1, C_ToUnitPopUp
    ggoSpread.SpreadLock C_Col2, -1, C_Col2
    ggoSpread.SpreadLock C_Col3, -1, C_Col3
	ggoSpread.SSSetProtected .vspdData.MaxCols, -1, -1
    .vspdData.ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SSSetRequired C_Dimension, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_Unit, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_ToUnit, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_Col1, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_Col2, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_Col3, pvStartRow, pvEndRow
    .vspdData.ReDraw = True
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_Dimension     = iCurColumnPos(1)
            C_DimensionCd   = iCurColumnPos(2)
            C_Unit          = iCurColumnPos(3)
            C_UnitPopUp     = iCurColumnPos(4)
            C_Col1          = iCurColumnPos(5)
            C_ToUnit        = iCurColumnPos(6)
            C_ToUnitPopUp   = iCurColumnPos(7)
            C_Col2          = iCurColumnPos(8)
            C_Factor        = iCurColumnPos(9)
            C_Col3          = iCurColumnPos(10)
            C_ToFactor      = iCurColumnPos(11)
    End Select    
End Sub

Sub cboDimension_OnChange()	
	frm1.txtUnit.value = ""
	frm1.txtUnitNm.value = ""
End Sub

Sub InitSpreadComboBox()
    Dim strCboData 
    Dim strCboData2
    Dim i 
    
    strCboData = ""
    strCboData2 = ""

	Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("B0005", "''", "S") & "  ", _
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
	
    strCboData = Replace(lgF0,chr(11),vbTab)
    strCboData2 = Replace(lgF1,chr(11),vbTab)  
    
	ggoSpread.SetCombo strCboData, C_DimensionCd
    ggoSpread.SetCombo strCboData2, C_Dimension
End Sub

Sub InitComboBox()
	Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("B0005", "''", "S") & "  ", _
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
	Call SetCombo2(frm1.cboDimension, lgF0, lgF1, Chr(11))
End Sub

Function OpenUnit(iSeq)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "단위 팝업"					<%' 팝업 명칭 %>
	arrParam(1) = "b_unit_of_measure"					<%' TABLE 명칭 %>
	arrParam(2) = frm1.txtUnit.value		<%' Code Condition%>
	arrParam(3) = ""							<%' Name Cindition%>
	
	If iSeq = 0 Then 
		If Trim(frm1.cboDimension.value) = "" Then
			arrParam(4) = ""
		Else
			arrParam(4) = "dimension =  " & FilterVar(frm1.cboDimension.value, "''", "S") & ""
		End If
	ElseIf iSeq = 2 Then	
		arrParam(4) = ""							<%' Where Condition%>
	Else 
		frm1.vspdData.Col = C_DimensionCd
		If frm1.vspdData.text = "" Then
			IntRetCD = DisplayMsgBox("970029", vbOKOnly, "Dimension", "X")			'⊙: "Will you destory previous data"
			IsOpenPop = False
			Exit Function
		End If
		arrParam(4) = "dimension =  " & FilterVar(frm1.vspdData.text , "''", "S") & ""	<%' Where Condition%>
	End If
	
	arrParam(5) = "단 위"						<%' 조건필드의 라벨 명칭 %>
	
    arrField(0) = "unit"					<%' Field명(0)%>
    arrField(1) = "unit_nm"					<%' Field명(1)%>
    
    arrHeader(0) = "단위코드"					<%' Header명(0)%>
    arrHeader(1) = "단위명"						<%' Header명(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtUnit.focus 
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetUnit(arrRet, iSeq)
	End If	
	
End Function

Function SetUnit(Byval arrRet, Byval iSeq)
	With frm1
		If iSeq = 0 Then 
			.txtUnit.value = arrRet(0)
			.txtUnitNm.value = arrRet(1)		
		ElseIf iSeq = 1 Then 
			.vspdData.Col = C_Unit
			.vspdData.Text = arrRet(0)

			lgBlnFlgChgValue = True
		Else 
			.vspdData.Col = C_ToUnit
			.vspdData.Text = arrRet(0)

			lgBlnFlgChgValue = True
		End If
	End With
End Function

Sub Form_Load()

    Call LoadInfTB19029                                                     <%'Load table , B_numeric_format%>
    
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
                                                                            <%'Format Numeric Contents Field%>                                                                            
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
           
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
   
    Call InitSpreadComboBox
    Call InitComboBox
    Call SetToolbar("1100110100101111")										<%'버튼 툴바 제어 %>
    frm1.cboDimension.focus 
    
End Sub

Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

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

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim index
	
	With frm1.vspdData
	
		.Row = Row
    
		Select Case Col
			Case C_Dimension
				.Col = Col
				index = .Value
				.Col = C_DimensionCd
				.Value = index
		End Select
	End With	
	ggoSpread.UpdateRow Row
End Sub

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 And Col = C_UnitPopUp Then
		    .Row = Row
		    .Col = C_Unit

		    Call OpenUnit(1)
		ElseIf Row > 0 And Col = C_ToUnitPopUp Then
		    .Row = Row
		    .Col = C_ToUnit

		    Call OpenUnit(2)
		End If
    End With
End Sub

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then					'☜: 조회중이면 다음 조회 안하도록 체크 
        Exit Sub
	End If
	
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    <%'☜: 재쿼리 체크 %>
    	If (lgStrPrevKey <> "" And lgStrPrevKey2 <> "" And lgStrPrevKey3 <> "") Then <%'다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 %>
      		Call DisableToolBar(parent.TBC_QUERY)					'☜ : Query 버튼을 disable 시킴.
			If DBQuery = False Then 
			   Call RestoreToolBar()
			   Exit Sub 
			End If 
    	End If

    End if
    
End Sub

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               <%'Protect system from crashing%>

<%  '-----------------------
    'Check previous data area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    
<%  '-----------------------
    'Erase contents area
    '----------------------- %>
    frm1.txtUnitNm.value = ""
    Call ggoOper.ClearField(Document, "2")									<%'Clear Contents  Field%>
    ggoSpread.ClearSpreadData
    Call InitVariables                                                      <%'Initializes local global variables%>
    															
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If
    
<%  '-----------------------
    'Query function call area
    '----------------------- %>
    If DbQuery = False Then Exit Function		  					<%'Query db data%>
       
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
    If Not ggoSpread.SSDefaultCheck Then     'Not chkField(Document, "2") OR    '⊙: Check contents area
       Exit Function
    End If
    
    ' 서로 대응되는 변환단위끼리 하나는 삭제이고 하나는 수정되는 것을 막음 
    if CheckDelOk = False then
        Call DisplayMsgBox("124014", "X", "X", "X")                          <%'No data changed!!%>
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

    FncCopy = False                                                               '☜: Processing is NG

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
 
    ggoSpread.Source = Frm1.vspdData

	With frm1.vspdData
		If .ActiveRow > 0 Then
			.focus
			.ReDraw = False
			
			ggoSpread.Source = frm1.vspdData 
			ggoSpread.CopyRow
			SetSpreadColor .ActiveRow, .ActiveRow
    
			'Key field clear
			.Col=C_DimensionCd
			.Text=""
			
			.Col=C_Dimension
			.Text=""
				
			.Col=C_Unit
			.Text=""
			
			.Col=C_ToUnit
			.Text=""

			.ReDraw = True
		End If
	End With

    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
End Function

Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim iRow 
    
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
	
		.vspdData.focus
		ggoSpread.Source = .vspdData
		
		.vspdData.ReDraw = False
        ggoSpread.InsertRow ,imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1
            .vspdData.Row = iRow
		    .vspdData.Col = C_Col1
		    .vspdData.value = " : "
		    .vspdData.Col = C_Col2
		    .vspdData.value = " = "
		    .vspdData.Col = C_Col3
		    .vspdData.value = " : "
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
    Call parent.FncExport(parent.C_MULTI)											 <%'☜: 화면 유형 %>
End Function

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                      <%'☜:화면 유형, Tab 유무 %>
End Function

Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 종료 하시겠습니까?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
	
	Call LayerShowHide(1)
	
	Dim strVal
    
    With frm1
    
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 
		strVal = strVal & "&txtDimensionCd=" & .hDimension.value 			'☆: 조회 조건 데이타 
		strVal = strVal & "&txtUnit=" & .hUnit.value 			'☆: 조회 조건 데이타 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&lgStrPrevKey3=" & lgStrPrevKey3
    Else
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 
		strVal = strVal & "&txtDimensionCd=" & Trim(.cboDimension.value)			'☆: 조회 조건 데이타 
		strVal = strVal & "&txtUnit=" & Trim(.txtUnit.value)			'☆: 조회 조건 데이타 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&lgStrPrevKey3=" & lgStrPrevKey3
    End If        
   
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True
    
End Function

Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE
   
	Call SetToolbar("1100111100111111")										<%'버튼 툴바 제어 %>
		
End Function

Function DbSave() 
    Dim lRow        
    Dim lGrpCnt  
	Dim strVal, strDel
	
    DbSave = False                                                          
    
    Call LayerShowHide(1)
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>

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
    
		    .vspdData.Row = lRow
		    .vspdData.Col = 0
		    
		    Select Case .vspdData.Text

		        Case ggoSpread.InsertFlag								'☜: 신규 
					
					strVal = strVal & "C" & parent.gColSep	& lRow & parent.gColSep	'☜: C=Create
										
		            .vspdData.Col = C_Unit		'3
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ToUnit	'6
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_Factor	'9
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gColSep
		            If unicdbl(.vspdData.Text) <= 0 Then
		            	.vspdData.Row = 0
						Call DisplayMsgBox("970022", "X", .vspdData.Text, "0")

						.vspdData.Row = lRow
						.vspdData.Action = 0 'ActionActiveCell
						.vspdData.EditMode = True
						Call LayerShowHide(0)
						Exit Function
					End If
		            		            
		            .vspdData.Col = C_ToFactor	'11
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gRowSep
		            If unicdbl(.vspdData.Text) <= 0 Then
		            	.vspdData.Row = 0
						Call DisplayMsgBox("970022", "X", .vspdData.Text, "0")

						.vspdData.Row = lRow
						.vspdData.Action = 0 'ActionActiveCell
						.vspdData.EditMode = True
						Call LayerShowHide(0)
						Exit Function
					End If

		            lGrpCnt = lGrpCnt + 1
		            
		        Case ggoSpread.UpdateFlag								'☜: 수정 
		
					strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep	'☜: U=Update

		            .vspdData.Col = C_Unit		'3
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ToUnit	'6
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_Factor	'9
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gColSep		            
		            If unicdbl(.vspdData.Text) <= 0 Then
		            	.vspdData.Row = 0
						Call DisplayMsgBox("970022", "X", .vspdData.Text, "0")

						.vspdData.Row = lRow
						.vspdData.Action = 0 'ActionActiveCell
						.vspdData.EditMode = True
						Call LayerShowHide(0)
						Exit Function
					End If
					
		            .vspdData.Col = C_ToFactor	'11
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gRowSep		            
		            If unicdbl(.vspdData.Text) <= 0 Then
		            	.vspdData.Row = 0
						Call DisplayMsgBox("970022", "X", .vspdData.Text, "0")

						.vspdData.Row = lRow
						.vspdData.Action = 0 'ActionActiveCell
						.vspdData.EditMode = True
						Call LayerShowHide(0)
						Exit Function
					End If

		            lGrpCnt = lGrpCnt + 1
		            
		            
		        Case ggoSpread.DeleteFlag								'☜: 삭제 

					strDel = strDel & "D" & parent.gColSep & lRow & parent.gColSep	'☜: U=Update

		            .vspdData.Col = C_Unit		'3
		            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ToUnit		'6
		            strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
		            
		            lGrpCnt = lGrpCnt + 1
		    End Select
		            
		Next
	
	.txtMaxRows.value = lGrpCnt-1	
	.txtSpread.value = strDel & strVal
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										<%'☜: 비지니스 ASP 를 가동 %>
	
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

Function CheckDelOk()
    Dim lRow,lRow2
    Dim lsDelUnit,lsDelToUnit,lsUpUnit,lsUpToUnit
    
	With frm1
         For lRow = 1 To .vspdData.MaxRows
             .vspdData.Row = lRow
             .vspdData.Col = 0
             if .vspdData.Text = ggoSpread.DeleteFlag then								'☜: 삭제 
                 .vspdData.Col = C_Unit		'3
                 lsDelUnit =  Trim(.vspdData.Text)
                 .vspdData.Col = C_ToUnit		'6
                 lsDelToUnit = Trim(.vspdData.Text) 
                 For lRow2 = 1 to  .vspdData.MaxRows
                     .vspdData.Row = lRow2
                     .vspdData.Col = 0
                     if .vspdData.Text = ggoSpread.UpdateFlag then						'☜: 수정 
                         .vspdData.Col = C_Unit		'3
                         lsUpUnit =  Trim(.vspdData.Text)
                         .vspdData.Col = C_ToUnit		'6
                         lsUpToUnit = Trim(.vspdData.Text) 
                         if lsDelUnit = lsUpToUnit and lsDelToUnit = lsUpUnit then
                             CheckDelOk = False
                             exit function
                         end if
                     end if
                 next
              end if
        next
	End with
	CheckDelOk = true
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>단위환산</font></td>
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
									<TD CLASS="TD5">Dimension</TD>
									<TD CLASS="TD6"><SELECT NAME="cboDimension" tag="11X" STYLE="WIDTH: 150px;"><OPTION value=""></OPTION></SELECT></TD>
									<TD CLASS="TD5">기준단위</TD>
									<TD CLASS="TD6">
										<INPUT TYPE=TEXT NAME="txtUnit" SIZE=10  MAXLENGTH=3 tag="11XXXU" ALT="단위코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnUnit" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenUnit(0)">
										<INPUT TYPE=TEXT NAME="txtUnitNm" tag="14X">
									</TD>
									<TD CLASS="TD6">&nbsp;</TD>
									<TD CLASS="TD6">&nbsp;</TD>
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
									<script language =javascript src='./js/b1f01ma1_vaSpread1_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="B1f01mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hDimension" tag="24">
<INPUT TYPE=HIDDEN NAME="hUnit" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

