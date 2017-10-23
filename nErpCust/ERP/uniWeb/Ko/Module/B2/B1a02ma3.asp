
<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(사용자정의 Minor코드등록)
'*  3. Program ID           : B1a02ma3.asp
'*  4. Program Name         : B1a02ma3.asp
'*  5. Program Desc         :
'*  6. Modified date(First) : 2000/09/19
'*  7. Modified date(Last)  : 2002/12/10
'*  8. Modifier (First)     : Hwang Jeong-won
'*  9. Modifier (Last)      : Sim Hae Young
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<% '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################%>
<% '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* %>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/lgvariables.inc"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'☜: indicates that All variables must be declared in advance


Const BIZ_PGM_MINOR = "B1a02mb4.asp"												'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_MAJOR = "B1a02mb5.asp"												'☆: 비지니스 로직 ASP명 

Dim C_Minor    
Dim C_MinorNm  
Dim C_MinorType

Dim C_Minor2
Dim C_MinorNm2  
Dim C_MinorType2

Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim lgIntGrpCount              ' Group View Size를 조사할 변수 
Dim lgIntFlgMode               ' Variable is for Operation Status

Dim lgStrPrevKey
Dim lgLngCurRows

Dim IsOpenPop
Dim lgSortKey

Sub InitSpreadPosVariables(ByVal pvSpdNo)  
    If pvSpdNo = "A" Then
        C_Minor        = 1
        C_MinorNm      = 2
        C_MinorType    = 3
    ElseIf pvSpdNo = "B" Then
        C_Minor2        = 1
        C_MinorNm2      = 2
        C_MinorType2    = 3
    End If
End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    
    lgStrPrevKey = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
    
End Sub

Sub InitSpreadSheet(ByVal pvSpdNo)
    If pvSpdNo = "" OR pvSpdNo = "A" Then
         Call initSpreadPosVariables("A")  

	     With frm1.vspdData
	         ggoSpread.Source = frm1.vspdData	
            'patch version
             ggoSpread.Spreadinit "V20021202",,parent.gAllowDragDropSpread    
             
	         .ReDraw = false

	         .MaxCols = C_MinorType + 1												'☜: 최대 Columns의 항상 1개 증가시킴 
	         .Col = .MaxCols														'☆: 사용자 별 Hidden Column
	         .ColHidden = True    
	                   
             .MaxRows = 0
             ggoSpread.ClearSpreadData

             Call GetSpreadColumnPos("A")  

	         ggoSpread.SSSetEdit C_Minor, "시스템정의 Minor코드", 26,,,10,2         '1
	         ggoSpread.SSSetEdit C_MinorNm, "Minor코드명", 50,,,50    '2
	         ggoSpread.SSSetCheck C_MinorType, "Minor코드 정의형식", 40, , "시스템 정의", True       '3
	
	         .ReDraw = true
	
             Call SetSpreadLock(1) 
         End With
    End If
    
    If pvSpdNo = "" OR pvSpdNo = "B" Then
         Call initSpreadPosVariables("B")  

	     With frm1.vspdData2
	         ggoSpread.Source = frm1.vspdData2	
            'patch version
             ggoSpread.Spreadinit "V20021202",,parent.gAllowDragDropSpread    
             
	         .ReDraw = false

	         .MaxCols = C_MinorType2 + 1												'☜: 최대 Columns의 항상 1개 증가시킴 
	         .Col = .MaxCols														'☆: 사용자 별 Hidden Column
	         .ColHidden = True    
	                   
             .MaxRows = 0
            ggoSpread.ClearSpreadData
	
             Call GetSpreadColumnPos("B")  
	
	         ggoSpread.SSSetEdit C_Minor2, "사용자정의 Minor코드", 26,,,10,2
	         ggoSpread.SSSetEdit C_MinorNm2, "Minor코드명", 50,,,50	'2
	         ggoSpread.SSSetCheck C_MinorType2, "Minor코드 정의형식", 40, , "사용자 정의", True	'3

	        .ReDraw = true
	
            Call SetSpreadLock(2)
        End With
    End If
    
End Sub

Sub SetSpreadLock(Byval intSheet)

	With frm1
	    Select Case Cint(intSheet)
	    	Case 1
	    		ggoSpread.Source = .vspdData
	    		.vspdData.ReDraw = False
	    		ggoSpread.SpreadLock C_Minor, -1, C_Minor
	    		ggoSpread.SpreadLock C_MinorNm, -1, C_MinorNm
	    		ggoSpread.SpreadLock C_MinorType, -1, C_MinorType
				ggoSpread.SSSetProtected .vspdData.MaxCols, -1, -1
	    		.vspdData.ReDraw = True
	    		
	    	Case 2
	    		ggoSpread.Source = .vspdData2
	    		.vspdData2.ReDraw = False
	    		ggoSpread.SpreadLock C_Minor2, -1, C_Minor2
	    		ggoSpread.SSSetRequired C_MinorNm2, -1, -1
	    		ggoSpread.SpreadLock C_MinorType2, -1, C_MinorType2 
				ggoSpread.SSSetProtected .vspdData2.MaxCols, -1, -1
	    		.vspdData2.ReDraw = True
	    End Select
	End With		
	
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
        ggoSpread.Source = .vspdData2
        .vspdData2.ReDraw = False
        ggoSpread.SSSetRequired  C_Minor2,     pvStartRow, pvEndRow
        ggoSpread.SSSetRequired  C_MinorNm2,   pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_MinorType2, pvStartRow, pvEndRow
        .vspdData2.ReDraw = True
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_Minor       = iCurColumnPos(1)
            C_MinorNm     = iCurColumnPos(2)
            C_MinorType   = iCurColumnPos(3)
       Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_Minor2      = iCurColumnPos(1)
            C_MinorNm2    = iCurColumnPos(2)
            C_MinorType2  = iCurColumnPos(3)
    End Select    
End Sub

Function OpenMajor()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	
	arrParam(0) = "Major코드 팝업"		' 팝업 명칭 
	arrParam(1) = "B_MAJOR"				 	' TABLE 명칭 
	arrParam(2) = frm1.txtMajor.value		' Code Condition
	arrParam(3) = ""						' Name Cindition
	arrParam(4) = "B_MAJOR.TYPE = " & FilterVar("U", "''", "S") & " "      ' Where Condition
	arrParam(5) = "Major코드"			
	
    arrField(0) = "major_cd"				' Field명(0)
    arrField(1) = "major_nm"				' Field명(1)
    
    arrHeader(0) = "Major코드"			' Header명(0)
    arrHeader(1) = "Major코드명"		' Header명(1)
    
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

Sub Form_Load()
    Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field%>                         
    
    Call InitSpreadSheet("")    
    Call InitVariables                                                      'Initializes local global variables
      
    Call SetToolbar("1100110100101111")										'버튼 툴바 제어 
    frm1.txtMajor.focus 
    
End Sub

Sub vspdData2_Change(ByVal Col, ByVal Row)

	ggoSpread.Source = frm1.vspdData2
	ggoSpread.UpdateRow Row
        
End Sub

Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("0000011111") 

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

Sub vspdData2_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SP1C"   

    Set gActiveSpdSheet = frm1.vspdData2
   
    If frm1.vspdData2.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData2
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If

	frm1.vspdData2.Row = Row
	
End Sub

Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
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

Sub vspdData2_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
        
    If Row <= 0 Then
        Exit Sub
    End If
        
    If frm1.vspdData2.MaxRows = 0 Then
        Exit Sub
    End If

End Sub
  
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

End Sub

Sub vspdData2_GotFocus()
    ggoSpread.Source = Frm1.vspdData2
    
End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)
    
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    

Sub vspdData2_MouseDown(Button , Shift , x , y)
    
    If Button = 2 And gMouseClickStatus = "SP1C" Then
       gMouseClickStatus = "SP1CR"
    End If
        
End Sub    

Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData2_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("B")
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

    Select Case gActiveSpdSheet.id
		Case "vaSpread"
			Call InitSpreadSheet("A")
		Case "vaSpread2"
			Call InitSpreadSheet("B")      		
	End Select 

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
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop)  Then	    '☜: 재쿼리 체크 
    	If lgStrPrevKey <> "" Then                  '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(parent.TBC_QUERY)					'☜ : Query 버튼을 disable 시킴.
			If DBQuery = False Then 
			   Call RestoreToolBar()
			   Exit Sub 
			End If 
    	End If

    End if
    
End Sub

Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then					'☜: 조회중이면 다음 조회 안하도록 체크 
        Exit Sub
	End If
	
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then	'hmkwon 2003.07.28 tracker 393
    'if frm1.vspdData2.MaxRows < NewTop + C_SHEETMAXROWS Then	 '   ☜: 재쿼리 체크 
    	If lgStrPrevKey <> "" Then                 '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
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
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    On Error Resume Next                                                    '☜: Protect system from crashing
        '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData2
    
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    '데이타가 변경되었습니다. 조회하시겠습니까?
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
    Call InitVariables
    															'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function										'☜: Query db data
       
    FncQuery = True																'⊙: Processing is OK
    
End Function

Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    On Error Resume Next                                                    '☜: Protect system from crashing
    
    
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

    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData

    Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
    Call InitVariables                                                      '⊙: Initializes local global variables
    
    FncNew = True                                                           '⊙: Processing is OK

End Function

Function FncDelete() 
    Dim IntRetCD 
    
    FncDelete = False                                                       '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    On Error Resume Next                                                    '☜: Protect system from crashing
    
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                                  '☆:
        Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    If DbDelete = False Then                                                '☜: Delete db data
       Exit Function                                                        '☜:
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                                         '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                         '⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData

    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData

    
    FncDelete = True                                                        '⊙: Processing is OK
    
End Function

Function FncSave() 
    
    FncSave = False                                                         '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    On Error Resume Next                                                    '☜: Protect system from crashing
    
    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData2
    If ggoSpread.SSCheckChange = False Then
        Call DisplayMsgBox("900001", "X", "X", "X")                          'No data changed!!
        Exit Function
    End If
    
    '-----------------------
    'Check content area
    '-----------------------
    ggoSpread.Source = frm1.vspdData2
    If Not ggoSpread.SSDefaultCheck Then     'Not chkField(Document, "2") OR    '⊙: Check contents area
       Exit Function
    End If
        
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          '⊙: Processing is OK
    
End Function

Function FncCopy() 
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

    If Frm1.vspdData2.MaxRows < 1 Then
       Exit Function
    End If
 
    ggoSpread.Source = Frm1.vspdData2

    With frm1.vspdData2
		If .ActiveRow > 0 Then
			.focus
			.ReDraw = false
		
			ggoSpread.CopyRow
			SetSpreadColor .ActiveRow, .ActiveRow
    
			'key field clear
			.Col=C_Minor2
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
	ggoSpread.Source = frm1.vspdData2 
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
	
	If Trim(frm1.txtMajor.value) = "" Then
		Call DisplayMsgBox("122204", "X", "X", "X")
		Exit Function
	End If
	
	If frm1.rdoChargeCd2.checked = True Then
		Call DisplayMsgBox("122504", "X", "X", "X")
		Exit Function
	End If

	With frm1
	    .vspdData2.focus
	    ggoSpread.Source = .vspdData2
	    .vspdData2.ReDraw = False
        ggoSpread.InsertRow ,imRow
        SetSpreadColor .vspdData2.ActiveRow, .vspdData2.ActiveRow + imRow - 1

        For iRow = .vspdData2.ActiveRow to .vspdData2.ActiveRow + imRow - 1
            .vspdData2.Row = iRow
	        .vspdData2.Col = C_MinorType2
	        .vspdData2.Text = "U"
        Next
	    .vspdData2.ReDraw = True
    End With
    lgBlnFlgChgValue = True

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function

Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
    
    With frm1.vspdData2 
    
    .focus
    ggoSpread.Source = frm1.vspdData2 
    
	lDelRows = ggoSpread.DeleteRow

    lgBlnFlgChgValue = True

    End With
    
End Function

Function FncPrev() 
    Dim strVal
    Dim IntRetCD
    
    If frm1.txtMajor2.value = "" Then                                      'Check if there is retrived data        
        Call DisplayMsgBox("900002", "X", "X", "X")  '☜ 바뀐부분        
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData2
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    '데이타가 변경되었습니다. 조회하시겠습니까?
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
            
    strVal = BIZ_PGM_MAJOR & "?txtMode=" & parent.UID_M0003							'☜: 
    strVal = strVal & "&txtMajor=" & frm1.txtMajor2.Value							'☆: 조회 조건 데이타 
    
	Call RunMyBizASP(MyBizASP, strVal)
    
End Function

Function FncNext() 
    Dim strVal
	Dim IntRetCD
	
    If frm1.txtMajor2.value = "" Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")  '☜ 바뀐부분        
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData2
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    '데이타가 변경되었습니다. 조회하시겠습니까?
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    
    strVal = BIZ_PGM_MAJOR & "?txtMode=" & parent.UID_M0004							'☜: 비지니스 처리 ASP의 상태값 
    strVal = strVal & "&txtMajor=" & frm1.txtMajor2.Value							'☆: 조회 조건 데이타 
    
	Call RunMyBizASP(MyBizASP, strVal)
    
End Function

Function FncPrint() 
    Call parent.FncPrint()                                                   '☜: Protect system from crashing
End Function

Function FncExit()
	Dim IntRetCD
	FncExit = False
    ggoSpread.Source = frm1.vspdData2
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    '데이타가 변경되었습니다. 조회하시겠습니까?
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    FncExit = True
End Function

Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)                                                   '☜: Protect system from crashing
End Function

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                         '☜:화면 유형, Tab 유무 
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
		strVal = BIZ_PGM_MINOR & "?txtMode=" & parent.UID_M0001							'☜: 
		strVal = strVal & "&txtMajor=" & .hMajor.value 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
    Else
		strVal = BIZ_PGM_MINOR & "?txtMode=" & parent.UID_M0001							'☜: 
		strVal = strVal & "&txtMajor=" & Trim(.txtMajor.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
    End If    

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True
    
End Function

Function DbQueryOk(ByVal lngMaxRow)														'☆: 조회 성공후 실행로직 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode    
    Call SetToolbar("1100111111111111")										'버튼 툴바 제어 
    
	frm1.vspdData2.col = C_minor2
    frm1.vspdData2.row = -1
    frm1.vspdData2.TypeMaxEditLen = cint(frm1.txtMinLen.value)
    
End Function

Function DbPrevNextOk()														'☆: 조회 성공후 실행로직 
	
	lgIntFlgMode = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
	frm1.vspdData2.MaxRows = 0
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
	
    Call MainQuery()

End Function

Function DbSave() 
    Dim pP21011     'As New P21011ManageIndReqSvr
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal, strDel
	
    DbSave = False                                                          '⊙: Processing is NG
    
    Call LayerShowHide(1)
    
    On Error Resume Next                                                   '☜: Protect system from crashing

	With frm1
		.txtMode.value = parent.UID_M0002
		.txtInsertUserId.value = parent.gUsrID
		
    
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
    For lRow = 1 To .vspdData2.MaxRows
    
        .vspdData2.Row = lRow
        .vspdData2.Col = 0
        
        Select Case .vspdData2.Text

            Case ggoSpread.InsertFlag											'☜: 신규 
				
				strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep 					'☜: C=Create, Row위치 정보 
                
                strVal = strVal & Trim(.txtMajor2.Value) & parent.gColSep
                
                .vspdData2.Col = C_Minor2	    '1
                strVal = strVal & Trim(.vspdData2.Text) & parent.gColSep
                
                .vspdData2.Col = C_MinorNm2  '2
                strVal = strVal & Trim(.vspdData2.Text) & parent.gColSep
                
                .vspdData2.Col = C_MinorType2	'3
				strVal = strVal & Trim(.vspdData2.Text) & parent.gRowSep
				
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.UpdateFlag											'☜: 수정 
				
				strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep 					'☜: U=Update, Row위치 정보 
				
				strVal = strVal & Trim(.txtMajor2.Value) & parent.gColSep
				
				.vspdData2.Col = C_Minor2	    '1
                strVal = strVal & Trim(.vspdData2.Text) & parent.gColSep
                
                .vspdData2.Col = C_MinorNm2  '2
                strVal = strVal & Trim(.vspdData2.Text) & parent.gColSep
                
                .vspdData2.Col = C_MinorType2	'3
				strVal = strVal & Trim(.vspdData2.Text) & parent.gRowSep 
				
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.DeleteFlag											'☜: 삭제 

				strDel = strDel & "D" & parent.gColSep & lRow & parent.gColSep									'☜: D=Delete, Row위치 정보 

                strDel = strDel & Trim(.txtMajor2.Value) & parent.gColSep
                
                .vspdData2.Col = C_Minor2	    '1
                strDel = strDel & Trim(.vspdData2.Text) & parent.gRowSep
                
                lGrpCnt = lGrpCnt + 1
                
        End Select
                
    Next
	
	.txtMaxRows.value = lGrpCnt-1
	.txtSpread.value = strDel & strVal
	
	Call ExecMyBizASP(frm1, BIZ_PGM_MINOR)										'☜: 비지니스 ASP 를 가동 
	
	End With
	
    DbSave = True                                                           '⊙: Processing is NG
    
End Function

Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
   
	Call InitVariables
	frm1.vspdData2.MaxRows = 0
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData

    Call MainQuery()

End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc"  -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>사용자정의 Minor코드등록</font></td>
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
										<INPUT TYPE=TEXT NAME="txtMajorNm" tag="14X"></TD>
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
						<TR>				
							<TD CLASS="TD5">Major코드</TD>
						    <TD CLASS="TD6">
										<INPUT TYPE=TEXT NAME="txtMajor2" SIZE=10 MAXLENGTH=5 tag="24" ALT="Major코드">
										<INPUT TYPE=TEXT NAME="txtMajorNm2" tag="24"></TD>

						</TR>
						<TR>
							
							<TD CLASS="TD5">Minor코드 길이</TD>
							<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtMinLen" SIZE=10 MAXLENGTH=5 tag="24" ALT="Minor길이"></TD>							
						</TR>
						<TR>
							
							<TD CLASS="TD5">사용자정의 Minor코드 추가가능여부</TD>
							<TD CLASS="TD6">
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoChargeCd" TAG="24" VALUE="Y" CHECKED ID="rdoChargeCd1">
							<LABEL FOR="rdoChargeCd1">가능</LABEL>
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoChargeCd" TAG="24" VALUE="N" ID="rdoChargeCd2">
							<LABEL FOR="rdoChargeCd2">불가능</LABEL>							
							</TD>							
						</TR>
						<TR>
							<TD HEIGHT="50%" WIDTH="100%" COLSPAN=2>
								<script language =javascript src='./js/b1a02ma3_vaSpread_vspdData.js'></script>
							</TD>
						</TR>
						<TR>	
							<TD HEIGHT="50%" WIDTH="100%" COLSPAN=2>
								<script language =javascript src='./js/b1a02ma3_vaSpread2_vspdData2.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="B1A02mb4.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsertUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hMajor" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

