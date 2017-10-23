
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Organization(horg_his 부서변경History)
'*  3. Program ID           : B2404ma1.asp
'*  4. Program Name         : B2404ma1.asp
'*  5. Program Desc         : 부서변경History등록 
'*  6. Modified date(First) : 2000/10/27
'*  7. Modified date(Last)  : 2005/10/17
'*  8. Modifier (First)     : Hwang Jeong-won
'*  9. Modifier (Last)      : Jeong Yong Kyun
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

Option Explicit																	<%'☜: indicates that All variables must be declared in advance%>

Const BIZ_PGM_ID = "B2404mb1.asp"												<%'비지니스 로직 ASP명 %> 

Dim C_OldDept       
Dim C_OldDeptPopup
Dim C_OldDeptNm
Dim C_Dept
Dim C_DeptPopup
Dim C_DeptNm
Dim C_ChggbnNm
Dim C_Chggbn

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop          

Dim lgStrPrevKey2, lgStrPrevKey3, lgStrPrevKey4

Sub InitSpreadPosVariables()
    C_OldDept       = 1
    C_OldDeptPopup  = 2
    C_OldDeptNm     = 3
    C_Dept          = 4
    C_DeptPopup     = 5
    C_DeptNm        = 6
    C_ChggbnNm      = 7
    C_Chggbn        = 8
End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevKey2 = ""                          'initializes Previous Key
    lgStrPrevKey3 = ""                          'initializes Previous Key
    lgStrPrevKey4 = ""                          'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "B","NOCOOKIE","MA") %>
End Sub

Sub InitSpreadSheet()
    Call initSpreadPosVariables()  

	With frm1.vspdData
	
	ggoSpread.Source = frm1.vspdData	
   'patch version
    ggoSpread.Spreadinit "V20021202",,parent.gAllowDragDropSpread    
    
	.ReDraw = false

    .MaxCols = C_Chggbn + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
	.Col = .MaxCols														'☆: 사용자 별 Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData
	
    Call GetSpreadColumnPos("A")  

    ggoSpread.SSSetEdit C_OldDept, "변경전 부서", 12,,,10,2
    ggoSpread.SSSetButton C_OldDeptPopUp
    ggoSpread.SSSetEdit C_OldDeptNm, "변경전 부서명", 35,,,200
    ggoSpread.SSSetEdit C_Dept, "변경후 부서", 12,,,10,2
    ggoSpread.SSSetButton C_DeptPopUp
    ggoSpread.SSSetEdit C_DeptNm, "변경후 부서명", 35,,,200    
    ggoSpread.SSSetCombo C_ChggbnNm, "변경사유", 17, 0, false, -1
    ggoSpread.SSSetCombo C_Chggbn, "", 16, 0, false, -1
	
    call ggoSpread.MakePairsColumn(C_OldDept,C_OldDeptPopUp)
    call ggoSpread.MakePairsColumn(C_Dept,C_DeptPopUp)

    Call ggoSpread.SSSetColHidden(C_Chggbn,C_Chggbn,True)
	
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub

Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False
    ggoSpread.SpreadLock C_OldDept, -1, C_OldDept
    ggoSpread.SpreadLock C_OldDeptPopUp, -1, C_OldDeptPopUp
    ggoSpread.SpreadLock C_OldDeptNm, -1, C_OldDeptNm
    
    ggoSpread.SpreadLock C_Dept, -1, C_Dept
    ggoSpread.SpreadLock C_DeptPopUp, -1, C_DeptPopUp
    ggoSpread.SpreadLock C_DeptNm, -1, C_DeptNm
	ggoSpread.SSSetProtected .vspdData.MaxCols, -1, -1
    .vspdData.ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False    
    ggoSpread.SSSetRequired C_OldDept, pvStartRow, pvEndRow    
    ggoSpread.SSSetProtected C_OldDeptNm, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_Dept, pvStartRow, pvEndRow    
    ggoSpread.SSSetProtected C_DeptNm, pvStartRow, pvEndRow
    .vspdData.ReDraw = True
    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_OldDept       = iCurColumnPos(1)
            C_OldDeptPopup  = iCurColumnPos(2)
            C_OldDeptNm     = iCurColumnPos(3)
            C_Dept          = iCurColumnPos(4)
            C_DeptPopup     = iCurColumnPos(5)
            C_DeptNm        = iCurColumnPos(6)
            C_ChggbnNm      = iCurColumnPos(7)
            C_Chggbn        = iCurColumnPos(8)
    End Select    
End Sub

Sub InitSpreadComboBox()
	Dim iCodeArr 
	Dim iNameArr
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6 

 Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("B0013", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = Replace(lgF0,chr(11),vbTab)
    iNameArr = Replace(lgF1,chr(11),vbTab)
    
    ggoSpread.SetCombo "" & vbTab & iNameArr, C_ChggbnNm
    ggoSpread.SetCombo "" & vbTab & iCodeArr, C_Chggbn
    
End Sub

Sub InitData()
	Dim intRow
	Dim intIndex 

	With frm1.vspdData
		For intRow = 1 To .MaxRows			
			.Row = intRow
			.Col = C_Chggbn	   :  intIndex = .Value             ' .Value means that it is index of cell,not value in combo cell type
			.Col = C_ChggbnNm  :  .Value = intindex					
		Next	
	End With
End Sub

Function OpenOrgId(iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "부서개편ID 팝업"				<%' 팝업 명칭 %>
	arrParam(1) = "horg_abs"						<%' TABLE 명칭 %>
	If iWhere = 0 Then
		arrParam(2) = frm1.txtOldOrgId.value		<%' Code Condition%>
	Else
		arrParam(2) = frm1.txtOrgId.value			<%' Code Condition%>
	End If
	
	arrParam(3) = ""								<%' Name Cindition%>
	arrParam(4) = ""								<%' Where Condition%>
	arrParam(5) = "부서개편ID"					<%' 조건필드의 라벨 명칭 %>
	
    arrField(0) = "orgid"							<%' Field명(0)%>
    arrField(1) = "orgnm"							<%' Field명(1)%>
    
    arrHeader(0) = "부서개편ID"					<%' Header명(0)%>
    arrHeader(1) = "부서개편명"					<%' Header명(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	IF iWhere = 0 Then 
		frm1.txtOldOrgId.focus 
	Else
		frm1.txtOrgId.focus 
	End if
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetOrgId(arrRet, iWhere)
	End If	
	
End Function

Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

		
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "부서 팝업"				<%' 팝업 명칭 %>
	arrParam(1) = "horg_mas"					<%' TABLE 명칭 %>
	arrParam(2) = frm1.vspdData.Text			<%' Code Condition%>	
	arrParam(3) = ""							<%' Name Cindition%>
			
	If iWhere = 0 Then
		arrParam(4) = " orgid =  " & FilterVar(frm1.txtOldOrgId.value, "''", "S") & " "		<%' Where Condition%>
	Else
		arrParam(4) = " orgid =  " & FilterVar(frm1.txtOrgId.value, "''", "S") & " " 		<%' Where Condition%>
	End If
	arrParam(5) = "부서코드"					<%' 조건필드의 라벨 명칭 %>
	
    arrField(0) = "dept"					<%' Field명(0)%>
    arrField(1) = "ldeptnm"					<%' Field명(1)%>
    
    arrHeader(0) = "부서코드"				<%' Header명(0)%>
    arrHeader(1) = "부서명"					<%' Header명(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetDept(arrRet, iWhere)
	End If	
	
End Function

Function SetOrgId(Byval arrRet, Byval iWhere)
		
	With frm1
		If iWhere = 0 Then 'TextBox(Condition)
			.txtOldOrgId.value = arrRet(0)
			.txtOldOrgNm.value = arrRet(1)
		Else 'spread
			.txtOrgId.value = arrRet(0)
			.txtOrgNm.value = arrRet(1)
		End If
	End With
End Function

Function SetDept(Byval arrRet, Byval iWhere)
	With frm1
		If iWhere = 0 Then 'Olddept값 
			.vspdData.Col = C_OldDept
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_OldDeptNm
			.vspdData.Text = arrRet(1)
			lgBlnFlgChgValue = True
		Else 'dept값 
			.vspdData.Col = C_Dept
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_DeptNm
			.vspdData.Text = arrRet(1)			
			lgBlnFlgChgValue = True
		End If
	End With
End Function

Sub Form_Load()

    Call LoadInfTB19029                                                     <%'Load table , B_numeric_format%>
    
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
                                                                            
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call InitSpreadComboBox
    Call SetToolbar("1100111100101111")										<%'버튼 툴바 제어 %>    
    frm1.txtOldOrgId.focus
    
End Sub

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

Sub vspdData_Change(ByVal Col , ByVal Row )

	Dim iDx
    Dim iCurrency
       
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    Select Case Col
        Case  C_ChggbnNm
            iDx = Frm1.vspdData.value
            Frm1.vspdData.Col = C_Chggbn
            Frm1.vspdData.value = iDx
    End Select    
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
             
    Call CheckMinNumSpread(frm1.vspdData, Col, Row)		
    
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
End Sub

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	If frm1.txtOldOrgId.value = "" Then
		Call DisplayMsgBox("970021", "X", "변경전 ID", "X")     '입력필수항목 Check
		Exit Sub
	ElseIf frm1.txtOrgId.value = "" Then
		Call DisplayMsgBox("970021", "X", "변경후 ID", "X")
		Exit Sub
	End If

	With frm1.vspdData 
	
    ggoSpread.Source = frm1.vspdData
   
    If Row > 0 And Col = C_OldDeptPopUp Then
		    .Row = Row
		    .Col = C_OldDept

		    Call OpenDept(0)
	ElseIf Row > 0 And Col = C_DeptPopUp Then
		    .Row = Row
		    .Col = C_Dept

		    Call OpenDept(1)
    End If
    
    End With
End Sub

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim index
	
	With frm1.vspdData
	
		.Row = Row
    
		Select Case Col
			Case C_ChggbnNm
				.Col = Col
				index = .Value
				.Col = C_Chggbn
				.Value = index
		End Select
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
    	If lgStrPrevKey <> "" And lgStrPrevKey2 <> "" And lgStrPrevKey3 <> "" And _
    	lgStrPrevKey4 <> "" Then                  <%'다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 %>
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

	
    If (RTrim(frm1.txtOldOrgId.value) = RTrim(frm1.txtOrgId.value)) AND _
       ((RTrim(frm1.txtOldOrgId.value) <> "") OR (RTrim(frm1.txtOrgId.value) <> ""))Then 
		Call DisplayMsgBox("124718", "X", "X", "X")     
        frm1.txtOldOrgId.focus
        Exit Function    
    End If
    
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
    frm1.txtOldOrgNm.value = ""
    frm1.txtOrgNm.value = ""
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
    If DbQuery = False Then Exit Function							<%'Query db data%>
       
    FncQuery = True															
    
End Function

Function FncSave() 
    
    If RTrim(frm1.txtOldOrgId.value) = RTrim(frm1.txtOrgId.value) Then 
		Call DisplayMsgBox("124718", "X", "X", "X")     
        frm1.txtOldOrgId.focus
        Exit Function    
    End If

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

    If RTrim(frm1.txtOldOrgId.value) = RTrim(frm1.txtOrgId.value) Then 
		Call DisplayMsgBox("124718", "X", "X", "X")     
        frm1.txtOldOrgId.focus
        Exit Function    
    End If

	With frm1
		If .vspdData.ActiveRow > 0 Then
			.vspdData.focus
			.vspdData.ReDraw = False
		
			ggoSpread.CopyRow
			 SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow
			
			.vspdData.Col = C_OldDept
			.vspdData.Text = ""
    
			.vspdData.Col = C_OldDeptNm
			.vspdData.Text = ""
			
			.vspdData.Col = C_Dept
			.vspdData.Text = ""
			
			.vspdData.Col = C_DeptNm
			.vspdData.Text = ""
				
			.vspdData.ReDraw = True
		End If
	End With
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
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

   

	If frm1.txtOldOrgId.value = "" Then                                      'Check if there is retrived data
		Call DisplayMsgBox("970021", "X", "변경전 ID", "X")
        Exit Function
    ElseIf frm1.txtOrgId.value = "" Then                                      'Check if there is retrived data
		Call DisplayMsgBox("970021", "X", "변경후 ID", "X")
        Exit Function
	End If
	
	If RTrim(frm1.txtOldOrgId.value) = RTrim(frm1.txtOrgId.value) Then 
		Call DisplayMsgBox("124718", "X", "X", "X")     
        frm1.txtOldOrgId.focus
        Exit Function    
    End If

	With frm1
		.vspdData.focus
		ggoSpread.Source = .vspdData
		'.vspdData.EditMode = True
		.vspdData.ReDraw = False
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

Function FncPrev() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

Function FncNext() 
    On Error Resume Next                                                    '☜: Protect system from crashing
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
    
    With Frm1
    
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&importoldorgid="     & .txtOldOrgId.Value         '☜: Query Key        
        strVal = strVal     & "&importorgid="        & .txtOrgId.Value         '☜: Query Key                
        
    Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '☜:  Run biz logic

    DbQuery = True  
End Function

Function DbQueryOk()													<%'조회 성공후 실행로직 %>	
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'부서개편 id nm값 리턴 
	
    call CommonQueryRs("ORGNM "," HORG_ABS "," ORGID =  " & FilterVar(frm1.txtOrgId.value, "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  
    frm1.txtOrgnm.value = Trim(Replace(lgF0,Chr(11),"")) 
    
    call CommonQueryRs("ORGNM "," HORG_ABS "," ORGID =  " & FilterVar(frm1.txtOldOrgId.value, "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  
    frm1.txtOldOrgnm.value = Trim(Replace(lgF0,Chr(11),""))       
	
	Call InitData()

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
				strVal = strVal & "C" & parent.gColSep						'☜: C=Create
		    Case ggoSpread.UpdateFlag								'☜: 수정 
				strVal = strVal & "U" & parent.gColSep						'☜: U=Update
		End Select
			
		Select Case .vspdData.Text
		    Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag			'☜: 수정, 신규 
					
				strVal = strVal & Trim(.txtOrgId.value) & parent.gColSep
					
				.vspdData.Col = C_Dept		'4
		        strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					
				strVal = strVal & Trim(.txtOldOrgId.value) & parent.gColSep					
					
		        .vspdData.Col = C_OldDept		'1
		        strVal = strVal & Trim(.vspdData.Text) & parent.gColSep		            
		            
		        .vspdData.Col = C_Chggbn		'7
		        strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep

		        lGrpCnt = lGrpCnt + 1
		            
		    Case ggoSpread.DeleteFlag								'☜: 삭제 

				strDel = strDel & "D" & parent.gColSep                     '☜: U=Update
					
				strDel = strDel & Trim(.txtOrgId.value) & parent.gColSep
					
				.vspdData.Col = C_Dept		'4
		        strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    
                strDel = strDel & Trim(.txtOldOrgId.value) & parent.gColSep
                     
		        .vspdData.Col = C_OldDept		'1
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
    
    Call MainQuery()
    
	Call SetSpreadLock 
End Function

Function btnBatch_OnClick()

	Dim strVal
	Dim IntRetCD
		
	If frm1.txtOldOrgId.value = "" Then                                      'Check if there is retrived data
		Call DisplayMsgBox("970021", "X", "변경전 ID", "X")
        Exit Function
	End If
	  
    IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO, "X", "X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If
	
    If RTrim(frm1.txtOldOrgId.value) = RTrim(frm1.txtOrgId.value) Then 
		Call DisplayMsgBox("124718", "X", "X", "X")     
        frm1.txtOldOrgId.focus
        Exit Function    
    End If

	Call LayerShowHide(1)
	
    strVal = BIZ_PGM_ID & "?txtMode=Gen"
	strVal = strVal     & "&txtOldOrgId=" & frm1.txtOldOrgId.value
	
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동	
	
	
End Function

Function Batch_OK()
	Dim i
	
	With frm1
		
		For i = 1 to .vspdData.MaxRows
			.vspdData.Col = 0
			.vspdData.Row = i
			.vspdData.Text = ggoSpread.Insertflag		
		Next
		
    ggoSpread.SpreadUnLock C_OldDept, -1, C_ChggbnNm
    .vspdData.ReDraw = False    
    ggoSpread.SSSetRequired C_OldDept, -1, -1
    ggoSpread.SSSetProtected C_OldDeptNm, -1, -1
    ggoSpread.SSSetRequired C_Dept, -1, -1    
    ggoSpread.SSSetProtected C_DeptNm, -1, -1
    '''ggoSpread.SSSetRequired C_ChggbnNm, -1, -1
    .vspdData.ReDraw = True
    
    End With
    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
    call CommonQueryRs("ORGNM "," HORG_ABS "," ORGID =  " & FilterVar(frm1.txtOrgId.value, "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  
    frm1.txtOrgnm.value = Trim(Replace(lgF0,Chr(11),""))     
    call CommonQueryRs("ORGNM "," HORG_ABS "," ORGID =  " & FilterVar(frm1.txtOldOrgId.value, "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  
    frm1.txtOldOrgnm.value = Trim(Replace(lgF0,Chr(11),""))       
    
    ggoSpread.ClearSpreadData "T"
	
End Function

Sub txtOldOrgId_onChange()
    Call Check_departmentID("txtOldOrgId")
End Sub

Sub txtOrgId_onChange()
    Call Check_departmentID("txtOrgId")
End Sub

Sub Check_departmentID(strID)
    If strID = "txtOldOrgId" Then
    
        If RTrim(frm1.txtOldOrgId.value) = "" Then 
            Exit Sub
        End If

        If  CommonQueryRs(" ORGNM "," HORG_ABS "," ORGID= " & FilterVar(frm1.txtOldOrgId.value, "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
	    	Call DisplayMsgBox("970000", "X", "부서개편ID", "X")     
            frm1.txtOldOrgNm.value = ""
            frm1.txtOldOrgId.focus
            Exit Sub    
        Else
            frm1.txtOldOrgNm.value = Replace(lgF0, Chr(11), "")
        End If
    Else
        If RTrim(frm1.txtOrgId.value) = "" Then 
            Exit Sub
        End If

        If  CommonQueryRs(" ORGNM "," HORG_ABS "," ORGID= " & FilterVar(frm1.txtOrgId.value, "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
	    	Call DisplayMsgBox("970000", "X", "부서개편ID", "X")     
            frm1.txtOrgNm.value = ""
            frm1.txtOrgId.focus
            Exit Sub    
        Else
            frm1.txtOrgNm.value = Replace(lgF0, Chr(11), "")
        End If
    End If

End Sub


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
									<TD CLASS="TD5">변경전ID</TD>
									<TD CLASS="TD6">
										<INPUT TYPE=TEXT NAME="txtOldOrgId" SIZE=10 MAXLENGTH=5 tag="12XXXU"  ALT="변경전 ID"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOrgId" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenOrgId(0)">
										<INPUT TYPE=TEXT NAME="txtOldOrgNm" SIZE=30 tag="14X">
									</TD>
									<TD CLASS="TD5">변경후ID</TD>
									<TD CLASS="TD6">
										<INPUT TYPE=TEXT NAME="txtOrgId" SIZE=10 MAXLENGTH=5 tag="12XXXU"  ALT="변경후 ID"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOrgId" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenOrgId(1)">
										<INPUT TYPE=TEXT NAME="txtOrgNm" SIZE=30 tag="14X">
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
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>								</TD>
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
					<TD><BUTTON NAME="btnBatch" CLASS="CLSMBTN" Flag=1>자동입력</BUTTON></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="B2404mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hOldOrgId" tag="24">
<INPUT TYPE=HIDDEN NAME="hOrgId" tag="24">
<INPUT TYPE=HIDDEN NAME="hOldDept" tag="24">
<INPUT TYPE=HIDDEN NAME="hDept" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>