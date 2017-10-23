
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Minor Code)
'*  3. Program ID           : b1a02ma1.asp
'*  4. Program Name         : b1a02ma1.asp
'*  5. Program Desc         :
'*  6. Modified date(First) : 1999/09/10
'*  7. Modified date(Last)  : 2002/12/10
'*  8. Modifier (First)     : Hwang Jeong-won
'*  9. Modifier (Last)      : Sim Hae Young
'* 11. Comment              :
'**********************************************************************************************
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
<SCRIPT LANGUAGE="vbscript" SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID = "B1a02mb1.asp"												'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_JUMP_ID = "B1a01ma1"												'☆: Jump시 호출 ASP명 

Const CookieSplit = 1233

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim C_Minor    
Dim C_MinorNm  
Dim C_MinorType

Dim lgStrQueryFlag			  ' "N":Next, "P":Prev, "Q":Query

Dim IsOpenPop          

Sub InitSpreadPosVariables()
    C_Minor     = 1
    C_MinorNm   = 2
    C_MinorType = 3
End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""
   
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgStrQueryFlag = "Q"
    
End Sub

Sub InitData()
	Dim intRow
	Dim intIndex 
	Dim pYesNo
	
    ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
	.ReDraw = false
		For intRow = 1 To .MaxRows
			.Row = intRow
			.Col = C_MinorType
			
 			If .Text = "시스템 정의" Then
 			    ggoSpread.SpreadLock	C_Minor,        -1, C_Minor,       -10
 			    ggoSpread.SpreadLock	C_MinorNm,  intRow, C_MinorNm,   intRow
 			    ggoSpread.SpreadLock	C_MinorType,    -1, C_MinorType,    -1
 			Else
 			    ggoSpread.SpreadLock C_MinorType,    -1, C_MinorType,   -1    
			End If

		Next	
	.ReDraw = True
	End With
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "B","NOCOOKIE","MA") %>
End Sub

Sub InitSpreadSheet()

    Call initSpreadPosVariables()  

	With frm1.vspdData  
	
	ggoSpread.Source = frm1.vspdData	
    ggoSpread.Spreadinit "V20021206",,parent.gAllowDragDropSpread    
    
	.ReDraw = false
	
	.MaxCols = C_MinorType + 1												'☜: 최대 Columns의 항상 1개 증가시킴 
	.Col = .MaxCols														'☆: 사용자 별 Hidden Column
    .ColHidden = True
     
    .MaxRows = 0
    ggoSpread.ClearSpreadData

    Call GetSpreadColumnPos("A")  

	ggoSpread.SSSetEdit C_Minor, "Minor코드", 26	, , ,15, 2			'1
	ggoSpread.SSSetEdit C_MinorNm, "Minor코드명", 50	, , ,50			'2
	ggoSpread.SSSetCombo C_MinorType, "Minor코드 정의형식", 40
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub

Sub SetSpreadLock()
    With frm1
    .vspdData.ReDraw = False
    ggoSpread.SpreadLock C_Minor, -1, C_Minor
	ggoSpread.SSSetRequired C_MinorNm, -1, -1
	ggoSpread.SSSetRequired C_MinorType, -1, -1
	ggoSpread.SSSetProtected .vspdData.MaxCols, -1, -1
    .vspdData.ReDraw = True
    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    Dim iRow
    
    With frm1
        .vspdData.ReDraw = False
    
        ggoSpread.SSSetRequired		C_Minor,        pvStartRow, pvEndRow
        ggoSpread.SSSetRequired		C_MinorNm,      pvStartRow, pvEndRow
        ggoSpread.SSSetProtected	C_MinorType,    pvStartRow, pvEndRow
    
        For iRow = pvStartRow to pvEndRow
            .vspdData.Col = C_MinorType
	        .vspdData.Row = iRow
	        .vspdData.Text = "사용자 정의"
        Next
        .vspdData.ReDraw = True
    End With
   
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_Minor     = iCurColumnPos(1)
            C_MinorNm   = iCurColumnPos(2)
            C_MinorType = iCurColumnPos(3)
    End Select    
End Sub

Sub InitSpreadComboBox()
    ggoSpread.SetCombo "시스템 정의" & vbtab & "사용자 정의", C_MinorType    
End Sub

Function OpenMajor()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "Major코드 팝업"			<%' 팝업 명칭 %>
	arrParam(1) = "B_MAJOR"				 		<%' TABLE 명칭 %>
	arrParam(2) = frm1.txtMajor.value			<%' Code Condition%>
	arrParam(3) = ""							<%' Name Cindition%>
	arrParam(4) = ""							<%' Where Condition%>
	arrParam(5) = "Major코드"			
	
    arrField(0) = "major_cd"					<%' Field명(0)%>
    arrField(1) = "major_nm"				<%' Field명(1)%>
    
    arrHeader(0) = "Major코드"						<%' Header명(0)%>
    arrHeader(1) = "Major코드명"						<%' Header명(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtMajor.focus 
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetMajor(arrRet)
	End If	

End Function

Function SetMajor(Byval arrRet)
	With frm1
		.txtMajor.value = arrRet(0)
		.txtMajorNm.value = arrRet(1)		
	End With
End Function

Function CookiePage(ByVal flgs)

	On Error Resume Next

	Const CookieSplit = 1233						<%'Cookie Split String : CookiePage Function Use%>

	Dim strTemp, arrVal
	
	If flgs = 1 Then
	
		WriteCookie CookieSplit , frm1.txtMajor.value 
		
	ElseIf flgs = 0 Then
	
		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function
		
		arrVal = Split(strTemp, parent.gRowSep)	

		If arrVal(0) = "" then Exit Function

		frm1.txtMajor.value =  arrVal(0)
		
		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit function
		End If
		
		WriteCookie CookieSplit , ""
		
		FncQuery()
			
	End If
			
End Function

Sub Form_Load()
 
    Call LoadInfTB19029                                                     <%'Load table , B_numeric_format%>
    
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
        
      
    Call InitSpreadSheet 
                                                       <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call InitSpreadComboBox
    Call SetToolbar("1100110100101111")									'⊙: 버튼 툴바 제어 
    Call CookiePage(0)
    
    frm1.txtMajor.focus 
    
End Sub

Sub vspdData_Change(ByVal Col, ByVal Row)
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

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then					'☜: 조회중이면 다음 조회 안하도록 체크 
        Exit Sub
	End If
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop)  And Not(lgStrPrevKey = "") Then
		Call DisableToolBar(parent.TBC_QUERY)					'☜ : Query 버튼을 disable 시킴.
		If DBQuery = False Then 
		   Call RestoreToolBar()
		   Exit Sub 
		End If 
    End if
    
End Sub

Function FncQuery() 
    Dim IntRetCD 

    lgStrQueryFlag = "Q"
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

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
    frm1.txtMajorNm.value = ""
    
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    ggoSpread.ClearSpreadData

    If lgStrQueryFlag = "Q" Then Call InitVariables							'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function		
       
    FncQuery = True																'⊙: Processing is OK
    
End Function

Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    'On Error Resume Next                                                    '☜: Protect system from crashing
    
    
    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")    
		If IntRetCD = vbNo Then
			Exit Function
		End If
       
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")                                         '⊙: Clear Condition Field
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    
    Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
    Call InitVariables                                                      '⊙: Initializes local global variables
    
    FncNew = True                                                           '⊙: Processing is OK

End Function

Function FncDelete() 
    Dim IntRetCD 
    
    FncDelete = False                                                       '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    'On Error Resume Next                                                    '☜: Protect system from crashing
    
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                                  '☆:
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")  
    If IntRetCD = vbNo Then
        Exit Function
    End If    
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")                                         '⊙: Clear Condition Field
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    
    FncDelete = True                                                        '⊙: Processing is OK
    
End Function

Function FncSave() 
    
    FncSave = False                                                         '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    'On Error Resume Next                                                    '☜: Protect system from crashing
    
    If Trim(frm1.txtMajor.value) = "" Then
		Call DisplayMsgBox("122204", "X", "X", "X")
		Exit Function
	End If
	
    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
        Exit Function
    End If
    
    '-----------------------
    'Check content area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then   'Not chkField(Document, "2") OR   '⊙: Check contents area
       Exit Function
    End If
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          '⊙: Processing is OK
    
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

Function FncCopy() 
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
 
    ggoSpread.Source = Frm1.vspdData

	If frm1.rdoChargeCd2.checked = True Then
		Call DisplayMsgBox("122504", "X", "X", "X")
		Exit Function
	End If

	With frm1.vspdData
		If .ActiveRow > 0 Then
			.focus
			.ReDraw = false
		
			ggoSpread.CopyRow
			SetSpreadColor .ActiveRow, .ActiveRow
    
			'key field clear
			.Col=C_Minor
			.Text=""
    				
			.ReDraw = true
		End If
    End with
   
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
    
	If Trim(frm1.txtMajor.value) = "" Then
		Call DisplayMsgBox("122204", "X", "X", "X")
		Exit Function
	End If
	
	If frm1.rdoChargeCd2.checked = True Then
		Call DisplayMsgBox("122504", "X", "X", "X")
		Exit Function
	End If

	
	With frm1
	    .vspdData.focus
        ggoSpread.Source = .vspdData
    
        .vspdData.ReDraw = False

        ggoSpread.InsertRow ,imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
    
        .vspdData.ReDraw = True

        lgBlnFlgChgValue = True
    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   

End Function

Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
        
    With frm1.vspdData 
    .focus
    .Row = .ActiveRow
    .Col = C_MinorType
    
    If .Text = "시스템 정의" Then
    	Call DisplayMsgBox("900031", "X", "X", "X")
    	Exit Function
    End If        
    ggoSpread.Source = frm1.vspdData 
    
	lDelRows = ggoSpread.DeleteRow

    lgBlnFlgChgValue = True
    
    End With
End Function

Function FncPrev() 	
	Dim IntRetCD 
	
	lgStrQueryFlag = "P"
	lgStrPrevKey = ""
	lgIntFlgMode = parent.OPMD_CMODE
	
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 조회 하시겠습니까?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    ggoSpread.ClearSpreadData

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function								'☜: Query db data
      
End Function

Function FncNext() 	
	Dim IntRetCD 
	
	lgStrQueryFlag = "N"
	lgStrPrevKey = ""
	lgIntFlgMode = parent.OPMD_CMODE
	
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 조회 하시겠습니까?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    ggoSpread.ClearSpreadData
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function
End Function

Function FncPrint() 
    Call parent.FncPrint()                                                   '☜: Protect system from crashing
End Function

Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)                                                   <%'☜: Protect system from crashing%>
End Function

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                         <%'☜:화면 유형, Tab 유무 %>
End Function

Function Clear()
	Call ggoOper.ClearField(Document, "2")
	ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData

End Function

Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim B1A028         'As New P21018ListIndReqSvr

    DbQuery = False
    
    Call LayerShowHide(1)
    
    Err.Clear                                                               '☜: Protect system from crashing

	Dim strVal
	    
    With frm1
        
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & lgStrQueryFlag							'☜: 
		strVal = strVal & "&txtMajor=" & .hMajor.value 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    Else
		strVal = BIZ_PGM_ID & "?txtMode=" & lgStrQueryFlag							'☜: 
		strVal = strVal & "&txtMajor=" & Trim(.txtMajor.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    End If
   
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True
    

End Function

Function DbQueryOk()														'☆: 조회 성공후 실행로직 
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode    
    
    Call InitData()
    Call SetToolbar("1100111111111111")										'⊙: 버튼 툴바 제어 
        
End Function

Function DbPrevNextOk()														'☆: 조회 성공후 실행로직 
	
	lgIntFlgMode = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
	
    Call MainQuery()
End Function

Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
	
    DbSave = False                                                          '⊙: Processing is NG
    
    Call LayerShowHide(1)
    
    'On Error Resume Next                                                   '☜: Protect system from crashing

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
		For lRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = lRow
		    .vspdData.Col = 0
		    
		    Select Case .vspdData.Text
		        Case ggoSpread.InsertFlag
					strVal = strVal & "C" & parent.gColSep	& lRow & parent.gColSep		'☜: C=Create
		        Case ggoSpread.UpdateFlag
					strVal = strVal & "U" & parent.gColSep	& lRow & parent.gColSep		'☜: U=Update
			End Select			

		    Select Case .vspdData.Text
		        Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag				'☜: 신규, 수정 
		            
		            If lgIntFlgMode = parent.OPMD_UMODE Then
						strVal = strVal & Trim(.hMajor.value) & parent.gColSep
		            Else
						strVal = strVal & Trim(.txtMajor.value) & parent.gColSep
					End if
					
		            .vspdData.Col = C_Minor	'1
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

		            .vspdData.Col = C_MinorNm	'2
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

		            .vspdData.Col = C_MinorType		
		            
		            Select Case Trim(.vspdData.Text)
						Case "시스템 정의"
							strVal = strVal & "S" & parent.gRowSep			'3
						Case "사용자 정의"								
							strVal = strVal & "U" & parent.gRowSep			'3
		            End Select
		            		           
		            lGrpCnt = lGrpCnt + 1

		        Case ggoSpread.DeleteFlag									'☜: 삭제 
		        
					strDel = strDel & "D" & parent.gColSep & lRow & parent.gColSep        '☜: U=Update

		            If lgIntFlgMode = parent.OPMD_UMODE Then
						strDel = strDel & Trim(.hMajor.value) & parent.gColSep
		            Else
						strDel = strDel & Trim(.txtMajor.value) & parent.gColSep
					End if

		            .vspdData.Col = C_Minor	'1
		            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_MinorNm	'2
		            strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
  
  		            lGrpCnt = lGrpCnt + 1
		    End Select
		Next
	
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strDel & strVal
	
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'☜: 비지니스 ASP 를 가동 
	
	End With
	
    DbSave = True                                                           '⊙: Processing is NG
End Function

Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
   
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>Minor코드등록</font></td>
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
						<TD CLASS="TD5">Major코드</TD>
						<TD CLASS="TD6">
										<INPUT TYPE=TEXT NAME="txtMajor" SIZE=10 MAXLENGTH=5 tag="12XXXU" ALT="Major코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCountryCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenMajor()">
										<INPUT TYPE=TEXT NAME="txtMajorNm" tag="14X"maxlength=30>
							</TABLE>
						</FIELDSET>
						</TD>
					</TR>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>						
						<TR>				
							<TD CLASS="TD5">Minor코드 길이</TD>
							<TD CLASS="TD6">
							<INPUT TYPE=TEXT NAME="txtMinorLen" SIZE=10 MAXLENGTH=2 tag="14" STYLE="Text-Align:Right" ALT="Minor길이"></TD>
						</TR>
						<TR>
							<TD CLASS="TD5">사용자정의 Minor코드 추가가능여부</TD>
							<TD CLASS="TD6">
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoChargeCd" TAG="2X" VALUE="Y" CHECKED ID="rdoChargeCd1" disabled>가능
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoChargeCd" TAG="2X" VALUE="N" ID="rdoChargeCd2" disabled>불가능
							</TD>							
						</TR>
						<TR>
							<TD WIDTH=100% HEIGHT=100% COLSPAN=2>
								<script language =javascript src='./js/b1a02ma1_OBJECT1_vspdData.js'></script>
							</TD>
						</TR>
						</TABLE>
					</TD>
				</TR>				
			</TABLE>
		</TD>
	</TR>
	<tr>
      <td <%=HEIGHT_TYPE_01%>></td>
    </tr>
    <tr HEIGHT="20">
      <td WIDTH="100%">
      <table WIDTH="100%">
        <tr>
          <td WIDTH="*" ALIGN=RIGHT><a href = "VBSCRIPT:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage 1">Major코드등록</a></td>
		  <TD WIDTH=10>&nbsp;</TD>
        </tr>
      </table>
      </td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="B1a02mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hMajor" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

