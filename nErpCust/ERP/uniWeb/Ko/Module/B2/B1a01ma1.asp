<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Major Code)
'*  3. Program ID           : b1a01ma1.asp
'*  4. Program Name         : b1a01ma1.asp
'*  5. Program Desc         :
'*  6. Modified date(First) : 1999/09/10
'*  7. Modified date(Last)  : 2002/12/02
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
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	<%'☜: indicates that All variables must be declared in advance%>

Const BIZ_PGM_ID = "B1a01mb1.asp"												<%'비지니스 로직 ASP명 %>
Const BIZ_PGM_JUMP_ID = "B1a02ma1"												'☆: Jump시 호출 ASP명 

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim C_MajorCd
Dim C_MajorNm
Dim C_MinorLen
Dim C_Type
Dim C_YesNo

Const CookieSplit = 1233

Dim IsOpenPop          
Dim lsConcd

Sub initSpreadPosVariables()  
    C_MajorCd  = 1
    C_MajorNm  = 2
    C_MinorLen = 3 
    C_Type     = 4
    C_YesNo    = 5
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

Sub InitData()
	Dim intRow
	Dim intIndex 
	Dim pYesNo
	
    ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
	.ReDraw = false
		For intRow = 1 To .MaxRows
			.Row = intRow
			.Col = C_YesNo : pYesNo = .Text 
			.Col = C_Type
 			If .Text = "시스템 정의"  Or pYesNo = "Y" Then
 			    ggoSpread.SpreadLock		-1, intRow, -1, intRow
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

    'patch version
    ggoSpread.Spreadinit "V20021119",,parent.gAllowDragDropSpread    
    
	.ReDraw = false
	
    .MaxCols = C_YesNo + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
	.Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
    .ColHidden = True
	
    .MaxRows = 0
    ggoSpread.ClearSpreadData
	
    Call GetSpreadColumnPos("A")    
   
    ggoSpread.SSSetEdit  C_MajorCd  , "Major코드"         ,20,  ,, 5,2
    ggoSpread.SSSetEdit  C_MajorNm  , "Major코드명"       ,50,  ,, 30    
    ggoSpread.SSSetFloat C_MinorLen ,"Minor코드 길이"     ,16,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"1","10"
    ggoSpread.SSSetCombo C_Type     , "Minor코드 정의형태",28                             
    ggoSpread.SSSetEdit  C_YesNo    , ""                      ,20,  ,, 5,2

    Call ggoSpread.SSSetColHidden(C_YesNo,C_YesNo,True)
	
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub

Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SpreadLock    C_MajorCd,  -1, C_MajorCd
    ggoSpread.SSSetRequired	C_MajorNm,  -1, -1
    ggoSpread.SSSetRequired	C_MinorLen, -1, -1
    ggoSpread.SpreadLock	C_Type,     -1, C_Type
	ggoSpread.SSSetProtected .vspdData.MaxCols, -1, -1
    .vspdData.ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    Dim iRow
    
    With frm1
    
    .vspdData.ReDraw = False
    
    ggoSpread.SSSetRequired		C_MajorCd,  pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_MajorNm,  pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_MinorLen, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_Type,     pvStartRow, pvEndRow
    
    For iRow = pvStartRow to pvEndRow
        .vspdData.Col = C_Type
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
            C_MajorCd      = iCurColumnPos(1)
            C_MajorNm      = iCurColumnPos(2)
            C_MinorLen     = iCurColumnPos(3)
            C_Type         = iCurColumnPos(4)
            C_YesNo        = iCurColumnPos(5)
            
    End Select    
End Sub

Sub InitSpreadComboBox()
    ggoSpread.SetCombo "시스템 정의" & vbtab & "사용자 정의", C_Type
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
    arrHeader(1) = "Major코드명"					<%' Header명(1)%>
    
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
	
	Call vspdData_Click(frm1.vspdData.ActiveCol, frm1.vspdData.ActiveRow)
	
	If flgs = 1 Then
		WriteCookie CookieSplit , lsConcd
		
	ElseIf flgs = 0 Then
	
		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function
		
		frm1.txtMajor.value =  strTemp			
		
		WriteCookie CookieSplit , ""
		
		FncQuery()
			
	End If

End Function

Sub Form_Load()


    Call LoadInfTB19029                                                     <%'Load table , B_numeric_format%>
    
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
                                                                            <%'Format Numeric Contents Field%>                                                                            
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call InitSpreadComboBox
    
    
    Call SetToolBar("1100110100101111")										<%'버튼 툴바 제어 %>
    Call CookiePage(0)
    
    frm1.txtMajor.focus 
    
End Sub

Sub vspdData_ButtonClicked(Col, Row, ButtonDown)

	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
    
	End With
End Sub

Sub vspdData_Change(ByVal Col , ByVal Row )
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

   	If Frm1.vspdData.CellType = Parent.SS_CELL_TYPE_FLOAT Then
      If CDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
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
    Else
    	frm1.vspdData.Row = Row
		frm1.vspdData.Col = C_MajorCd
	
		lsConcd = frm1.vspdData.Text		
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

<%  '-----------------------
    'Check previous data area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    
<%  '-----------------------
    'Erase contents area
    '----------------------- %>
    frm1.txtMajorNm.value = ""
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
    If DbQuery	= False Then														<%'Query db data%>
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

    FncCopy = False                                                               '☜: Processing is NG

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
 
    ggoSpread.Source = Frm1.vspdData

 
 
	With frm1.vspdData
		If .ActiveRow > 0 Then
			.focus
			.ReDraw = false
		
			ggoSpread.Source = frm1.vspdData	
			ggoSpread.CopyRow
			SetSpreadColor .ActiveRow, .ActiveRow
    
			'key field clear
			.Col=C_MajorCd
			.Text=""
			
			'Hidden Field clear
			.Col=C_YesNo
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
    Dim lDelRows

    With frm1.vspdData 
    	.focus
    	.Col = C_Type 
    	If .Text = "시스템 정의" Then
    		Call DisplayMsgBox("900031", "X", "X", "X")
    		Exit Function
    	End If
    	.Col = C_YesNo
		If .Text = "Y" Then
    		Call DisplayMsgBox("900032", "X", "X", "X")
    		Exit Function
    	End If
    	
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
	Call InitData()
End Sub

Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>

	Call LayerShowHide(1)
	
	Dim strVal
    
    With frm1
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
    <% '@Query_Hidden %>    
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001						<%'Hidden의 검색조건으로 Query%>
		strVal = strVal & "&txtMajor=" & .hMajorCd.value				
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    Else
    <% '@Query_Text %>    
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001						<%'현재 검색조건으로 Query%>
		strVal = strVal & "&txtMajor=" & .txtMajor.value				
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    End If
    
	Call RunMyBizASP(MyBizASP, strVal)										<%'☜: 비지니스 ASP 를 가동 %>
        
    End With
    
    DbQuery = True
    
End Function

Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    '-----------------------
    'Reset variables area
    '-----------------------
	lgIntFlgMode = Parent.OPMD_UMODE
    Call InitData()
	Call SetToolBar("110011010011111")										<%'버튼 툴바 제어 %>
	
End Function

Function DbSave() 
    Dim pP21011																<%'As New P21011ManageIndReqSvr%>
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
		.txtUpdtUserId.value = Parent.gUsrID
		.txtInsrtUserId.value = Parent.gUsrID
    
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

            Case ggoSpread.InsertFlag									    <%'☜: 신규 %>
				
				strVal = strVal & "C" & Parent.gColSep					  			<%'☜: C=Create, Row위치 정보 %>

                .vspdData.Col = C_MajorCd	'1
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                
                .vspdData.Col = C_MajorNm	'2
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                
                .vspdData.Col = C_MinorLen	'3
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                
                .vspdData.Col = C_Type		
		            
		            Select Case Trim(.vspdData.Text)
						Case "시스템 정의"
							strVal = strVal & "S" & Parent.gRowSep			'4
						Case "사용자 정의"								
							strVal = strVal & "U" & Parent.gRowSep			'4
		            End Select
                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.UpdateFlag
                
                strVal = strVal & "U" & Parent.gColSep								<%'☜: C=Create, Row위치 정보 %>

                .vspdData.Col = C_MajorCd	'1
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                
                .vspdData.Col = C_MajorNm	'2
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                
                .vspdData.Col = C_MinorLen	'3
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                
                .vspdData.Col = C_Type		
		            
		            Select Case Trim(.vspdData.Text)
						Case "시스템 정의"
							strVal = strVal & "S" & Parent.gRowSep			'4
						Case "사용자 정의"								
							strVal = strVal & "U" & Parent.gRowSep			'4
		            End Select
                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.DeleteFlag										<%'☜: 삭제 %>

				strDel = strDel & "D" & Parent.gColSep								<%'☜: D=Update, Row위치 정보 %>
				
                .vspdData.Col = C_MajorCd	'1
                strDel = strDel & Trim(.vspdData.Text) & Parent.gRowSep									
                
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>Major코드등록</font></td>
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
									<TD CLASS="TD656">
										<INPUT TYPE=TEXT NAME="txtMajor" SIZE=10 MAXLENGTH=5 tag="11XXXU" ALT="Major코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMajorCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenMajor()">
										<INPUT TYPE=TEXT NAME="txtMajorNm" tag="14X">
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
									<script language =javascript src='./js/b1a01ma1_I539911192_vspdData.js'></script>
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
    <tr HEIGHT="20">
      <td WIDTH="100%">
      <table <%=LR_SPACE_TYPE_30%>>
        <tr>
          <td WIDTH="*" ALIGN=RIGHT><a href = "VBSCRIPT:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage 1">Minor코드등록</a></td>
				<TD WIDTH=10>&nbsp;</TD>
        </tr>
      </table>
      </td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="B1a01mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24">
</TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hMajorCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

