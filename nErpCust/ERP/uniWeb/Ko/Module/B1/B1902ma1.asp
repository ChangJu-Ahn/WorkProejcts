
<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Count Format)
'*  3. Program ID           : B1902ma1.asp
'*  4. Program Name         : B1902ma1.asp
'*  5. Program Desc         :
'*  6. Modified date(First) : 2000/09/18
'*  7. Modified date(Last)  : 2002/12/04
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
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit		
															

Const BIZ_PGM_ID = "B1902mb1.asp"							'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_COMMON_FORMAT = "b1901ma1"
Const BIZ_PGM_NUMERIC_FORMAT = "b1903ma1"

Const TAB1 = 1										            <%'Tab의 위치 %>
Const TAB2 = 2

Dim C_ModuleNm
Dim C_ModuleCd
Dim C_Decimals
Dim C_RndUnit
Dim C_RndPolicy
Dim C_RndPolicyNm
Dim C_DataFormat 

Dim IsOpenPop
Dim gSelframeFlg                                            <%'Current Tab Page%>
Dim iCurrentCD    ''이전TAB2의 선택모듈CD값 셋팅 

<!-- #Include file="../../inc/lgvariables.inc" -->	

Sub InitSpreadPosVariables()
    C_ModuleNm = 1
    C_ModuleCd = 2
    C_Decimals = 3
    C_RndUnit = 4
    C_RndPolicy = 5
    C_RndPolicyNm = 6
    C_DataFormat = 7
End Sub

Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size

    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count    
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

    .MaxCols = C_DataFormat + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
	.Col = .MaxCols														'☆: 사용자 별 Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData
	
	Call AppendNumberPlace("6","1","0")
    Call GetSpreadColumnPos("A")  
    
    ggoSpread.SSSetCombo C_ModuleNm, "업무", 26
    ggoSpread.SSSetCombo C_ModuleCd, " ", 10
	ggoSpread.SSSetFloat C_Decimals,"소수점자리수" ,16,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0","6"
    ggoSpread.SSSetEdit C_RndUnit, "올림처리단위", 20, 1
	ggoSpread.SSSetCombo C_RndPolicy, " ", 20
	ggoSpread.SSSetCombo C_RndPolicyNm, "올림구분", 20
	ggoSpread.SSSetEdit C_DataFormat, "포맷", 33 , 1                        

    Call ggoSpread.SSSetColHidden(C_ModuleCd,C_ModuleCd,True)
    Call ggoSpread.SSSetColHidden(C_RndPolicy,C_RndPolicy,True)

	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub

Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SpreadLock C_ModuleNm, -1, C_ModuleNm
    ggoSpread.SSSetRequired	C_Decimals, -1, -1
    ggoSpread.SpreadLock C_RndUnit, -1, C_RndUnit
    ggoSpread.SSSetRequired	C_RndPolicyNm, -1, -1
    ggoSpread.SpreadLock C_DataFormat, -1, C_DataFormat
	ggoSpread.SSSetProtected .vspdData.MaxCols, -1, -1
    .vspdData.ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1    
        .vspdData.ReDraw = False
        
        ggoSpread.SSSetRequired C_ModuleNm,    pvStartRow, pvEndRow
        ggoSpread.SSSetRequired C_Decimals,    pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_RndUnit,    pvStartRow, pvEndRow
        ggoSpread.SSSetRequired C_RndPolicyNm, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_DataFormat, pvStartRow, pvEndRow
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
            C_Decimals = iCurColumnPos(3) 
            C_RndUnit = iCurColumnPos(4) 
            C_RndPolicy = iCurColumnPos(5) 
            C_RndPolicyNm = iCurColumnPos(6) 
            C_DataFormat = iCurColumnPos(7) 
            
    End Select    
End Sub

Function ClickTab1()
	Dim IntRetCD
	
	If gSelframeFlg = TAB1 Then Exit Function
	
	ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X") '☜ 바뀐부분 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	
	Call changeTabs(TAB1)                                               <%'첫번째 Tab%>
	gSelframeFlg = TAB1
	
	frm1.cboModuleCd.value = "*"
	frm1.cboModuleCd.disabled = True
	frm1.cboFormType.value = "I"
	Call SetToolbar("1100100000011111")
	Call MainQuery()
	
End Function

Function ClickTab2()	
	Dim IntRetCD
	
	If gSelframeFlg = TAB2 Then Exit Function
	
	If lgBlnFlgChgValue = True Then 
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X") '☜ 바뀐부분 
		If IntRetCD = vbNo Then
			Exit Function		
		Else
			lgBlnFlgChgValue = False
		End If
	End If
		
	Call changeTabs(TAB2)
	gSelframeFlg = TAB2
	
	''frm1.cboModuleCd.value = "*"
	frm1.cboModuleCd.value = iCurrentCD  '''8/28 : Srh
	frm1.cboModuleCd.disabled = False	
	frm1.cboFormType.value = "Q"
	Call SetToolbar("1100111100111111")										'⊙: 버튼 툴바 제어 
	Call MainQuery()   
	
End Function

Sub InitSpreadComboBox()
	Dim strCboData
	Dim strCboData2
	
	''MODULE
	Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("B0001", "''", "S") & "  ", _
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
	        
	strCboData = Replace(lgF0,chr(11),vbTab)
    strCboData2 = Replace(lgF1,chr(11),vbTab)
    strCboData = Left(strCboData,Len(strCboData) - 1)
    strCboData2 = Left(strCboData2,Len(strCboData2) - 1)
    
	ggoSpread.SetCombo strCboData, C_ModuleCd
	ggoSpread.SetCombo strCboData2, C_ModuleNm
	
	''FLAG
	Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("B0004", "''", "S") & "  ", _	
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
	strCboData = Replace(lgF0,chr(11),vbTab)
    strCboData2 = Replace(lgF1,chr(11),vbTab)
    strCboData = Left(strCboData,Len(strCboData) - 1)
    strCboData2 = Left(strCboData2,Len(strCboData2) - 1)
	ggoSpread.SetCombo strCboData, C_RndPolicy
	ggoSpread.SetCombo strCboData2, C_RndPolicyNm
	
End Sub

Sub InitComboBox()
	''MODULE
	Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("B0001", "''", "S") & "  ", _
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
	Call SetCombo2(frm1.cboModuleCd, lgF0, lgF1, Chr(11))	
	        
	''FORM TYPE
	Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("B0002", "''", "S") & "  ", _
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
    	               
    Call SetCombo2(frm1.cboFormType, lgF0, lgF1, Chr(11))
	''FLAG
	Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("B0004", "''", "S") & "  ", _	
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
	Call SetCombo2(frm1.cboFlag, lgF0, lgF1, Chr(11))
	
End Sub

Function LoadCommonFormat()
    
    PgmJump(BIZ_PGM_COMMON_FORMAT)

End Function

Function LoadNumericFormat()
    
    PgmJump(BIZ_PGM_NUMERIC_FORMAT)

End Function

Sub Form_Load()
	
	Call LoadInfTB19029                          <%'Load table , B_numeric_format%>

    Call ggoOper.LockField(Document, "N")        <%'Lock  Suitable  Field%>    					                                           <%'Format Numeric Contents Field%>                                                     
	Call AppendNumberPlace("6", "1", "0")    					                                                    
	Call AppendNumberRange("0", "0", "6")

  	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart, ggStrMaxPart)
  	
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    
    Call InitVariables                             '⊙: Initializes local global variables
    
    '----------  Coding part  -------------------------------------------------------------    
	Call InitSpreadComboBox
	Call InitComboBox		 
    
    gIsTab     = "Y"
    gTabMaxCnt = 2
    
    Call ClickTab1
    
End Sub

Function LoadCountFormat()
    
    PgmJump(BIZ_PGM_COUNT_FORMAT)

End Function

Function LoadNumericFormat()
    
    PgmJump(BIZ_PGM_NUMERIC_FORMAT)

End Function

Sub txtDec_Change()
Dim i, j
Dim lRound, lRoundP
	
With frm1
		j = .txtDec.value
		If Len(j) > 1 Or j > 6 Or j < 0 Then
			.txtDec.value = 0
			j = 0
		End If
		
		lRound = 0.1
		lRoundP = 1
	
		If j > 0 Then
			For i = 1 To j
				lRound = lRound * 0.1
			Next	
			.txtFlag.value = lRound
			
		ElseIf j = 0 Then
			.txtFlag.value = lRound
			
		Else
			For i = 1 To (j * -1)
				lRoundP = lRoundP * 10
		    Next
		        
		    lRoundP = lRoundP / 10
		    
		    .txtFlag.value = lRoundP
		End If	

		.txtFlag.value = replace (.txtFlag.value, ".", "@")
		.txtFlag.value = replace (.txtFlag.value, ",", "$")

		.txtFlag.value = replace (.txtFlag.value, "@", parent.gComNumDec)
		.txtFlag.value = replace (.txtFlag.value, "$", parent.gComNum1000)
	
	 	
		   Select Case CInt(j)
			  Case -1
				.txtFormat.value = "##,###,###,##0"
			  Case -2
				.txtFormat.value = "##,###,###,#00"
		      Case -3
				.txtFormat.value = "##,###,###,000"		
		  	  Case -4
				.txtFormat.value = "##,###,##0,000"
			  Case 0
				.txtFormat.value = "##,###,###,###"
			  Case 1
				.txtFormat.value = "##,###,###,###.0"
		 	  Case 2
				.txtFormat.value = "##,###,###,###.00"
			  Case 3
				.txtFormat.value = "##,###,###,###.000"
			  Case 4
				.txtFormat.value = "##,###,###,###.0000"
			  Case 5
				.txtFormat.value = "#,###,###,###.00000"
			  Case 6
				.txtFormat.value = "###,###,###.000000"			
			  Case Else
				.txtFormat.value = "##,###,###,###"
		  End Select
		  
			.txtFormat.value = replace (.txtFormat.value, ".", "@")
		    .txtFormat.value = replace (.txtFormat.value, ",", "$")
		    .txtFormat.value = replace (.txtFormat.value, "@", parent.gComNumDec)
		    .txtFormat.value = replace (.txtFormat.value, "$", parent.gComNum1000)
	End With
	
	lgBlnFlgChgValue = True

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

Sub vspdData_Change(ByVal Col, ByVal Row )
Dim i, j
Dim intIndex
Dim lRound, lRoundP

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

    'lgBlnFlgChgValue = True
    
    With frm1.vspdData
    
		If Col = C_Decimals Then
    
		  .Col = Col
		  .Row = Row
		  j = .value
    
		  lRound = 0.1
		  lRoundP = 1
    
		    If j > 0 Then
		        For i = 1 To j
		            lRound = lRound * 0.1
		        Next
		        
		        .Col = C_RndUnit
		        .Row = Row
		        .value = lRound
		        
		    ElseIf j = 0 Then
		        .Col = C_RndUnit
		        .Row = Row
		        .value = lRound
		        
		    Else
		        For i = 1 To (j * -1)
		            lRoundP = lRoundP * 10
		        Next
		        
		        lRoundP = lRoundP / 10
		        .Col = C_RndUnit
		        .Row = Row
		        .value = lRoundP
		        
		    End If    
		   
	
		.value = replace (.value, ".", "@")
		.value = replace (.value, ",", "$")
		.value = replace (.value, "@", parent.gComNumDec)
		.value = replace (.value, "$", parent.gComNum1000)

			 
		.Col = C_DataFormat
		.Row = Row
			    
		   Select Case CInt(j)
			  Case -1
				.Text = "##,###,###,##0"
			  Case -2
				.Text = "##,###,###,#00"
		      Case -3
				.Text = "##,###,###,000"		
		  	  Case -4
				.Text = "##,###,##0,000"
			  Case 0
				.Text = "##,###,###,###"
			  Case 1
				.Text = "##,###,###,###.0"
		 	  Case 2
				.Text = "##,###,###,###.00"
			  Case 3
				.Text = "##,###,###,###.000"
			  Case 4
				.Text = "##,###,###,###.0000"
			  Case 5
				.Text = "#,###,###,###.00000"
			  Case 6
				.Text = "###,###,###.000000"			
			  Case Else
				.Text = "##,###,###,###"
		  End Select
		  .text = replace (.text, ".", "@")
		  .text = replace (.text, ",", "$")
		  .text = replace (.text, "@", parent.gComNumDec)			'Parent.gComNumDec
		  .text = replace (.text, "$", parent.gComNum1000)			'Parent.gComNum1000
		End If													
		
    End with
            
End Sub


function FormatChanging(ByVal cnt )
	
	dim strFormat
			    
	 Select Case CInt(cnt)
	    Case -1
		  	strFormat = "##,###,###,##0"
	    Case -2
		  	strFormat = "##,###,###,#00"
	    Case -3
	  		strFormat = "##,###,###,000"		
	    Case -4
	  		strFormat = "##,###,##0,000"
	    Case 0
	  		strFormat = "##,###,###,###"
	    Case 1
	  		strFormat = "##,###,###,###.0"
	    Case 2
	  		strFormat = "##,###,###,###.00"
	    Case 3
	  		strFormat = "##,###,###,###.000"
	    Case 4
	  		strFormat = "##,###,###,###.0000"
	    Case 5
	  		strFormat = "#,###,###,###.00000"
	    Case 6
	  		strFormat = "###,###,###.000000"			
	    Case Else
	  		strFormat = "##,###,###,###"
	End Select
	
	strFormat = replace (strFormat, ".", "@")
	strFormat = replace (strFormat, ",", "$")
	strFormat = replace (strFormat, "@", parent.gComNumDec)			'Parent.gComNumDec
	strFormat = replace (strFormat, "$", parent.gComNum1000)			'Parent.gComNum1000
		
	FormatChanging = strFormat            
    
End function

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

End Sub

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then					'☜: 조회중이면 다음 조회 안하도록 체크 
        Exit Sub
	End If
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
  
End Sub

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim index
	
	With frm1.vspdData
		If Col = C_RndPolicyNm And Row > 0 Then
			.Row = Row
			.Col = Col
			index = .TypeComboBoxCurSel
			
			.Col = C_RndPolicy
			.TypeComboBoxCurSel = index
		ElseIf Col = C_ModuleNm And Row > 0 Then
			.Row = Row
			.Col = Col
			index = .TypeComboBoxCurSel
			
			.Col = C_ModuleCd
			.TypeComboBoxCurSel = index		
		End If
	End With
	
	ggoSpread.UpdateRow Row
End Sub

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
	
<%  '-----------------------
    'Check previous data area
    '----------------------- %>    
    If gSelframeFlg = TAB1 Then
    	If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
    		If IntRetCD = vbNo Then
      			Exit Function
    		End If
    	End If    	
    Else
    	ggoSpread.Source = frm1.vspdData
    	If ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
    		If IntRetCD = vbNo Then
      			Exit Function
    		End If
    	End If
	End If
	        
    '-----------------------
    'Erase contents area
    '-----------------------
	Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
	ggoSpread.ClearSpreadData

    Call InitVariables	

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = false Then														'☜: Query db data
		Exit Function
	End If	
	       
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
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X") '☜ 바뀐부분		
		If IntRetCD = vbNo Then
			Exit Function
		End If
       
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")
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
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                                  '☆:
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X")  '☜ 바뀐부분 
    If IntRetCD = vbNo Then
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
    Call ggoOper.ClearField(Document, "A")
    ggoSpread.ClearSpreadData
    
    FncDelete = True                                                        '⊙: Processing is OK
    
End Function

Function FncSave()

    FncSave = False                                                         '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    On Error Resume Next                                                    '☜: Protect system from crashing
    
    frm1.txtLogInCnt.value = "0"	'현재로그인 user를 초기화 시킴    
    '-----------------------
    'Precheck area
    '-----------------------    
    If gSelframeFlg = TAB1 Then    
		If lgBlnFlgChgValue = False Then
			Call DisplayMsgBox("900001", "X", "X", "X")                            '⊙: No data changed!!        
			Exit Function
		End If
	Else
		ggoSpread.Source = frm1.vspdData
		If ggoSpread.SSCheckChange = False Then
			Call DisplayMsgBox("900001", "X", "X", "X")                            '⊙: No data changed!!        
			Exit Function
		End If
	End If
    
    '-----------------------
    'Check content area
    '-----------------------
    If gSelframeFlg = TAB2 Then
	    ggoSpread.Source = frm1.vspdData
		If Not ggoSpread.SSDefaultCheck Then  			'Not chkField(Document, "2")
			Call changeTabs(TAB2)
			Exit Function
		End If
    End If
    '-----------------------
    'Save function call area
    '-----------------------    
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          '⊙: Processing is OK
    
End Function

Function FncCopy()
    FncCopy = False                                                               '☜: Processing is NG

    Err.Clear                                                               '☜: Protect system from crashing
    On Error Resume Next                                                    '☜: Protect system from crashing

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
 
    ggoSpread.Source = Frm1.vspdData
 
    With frm1.vspdData
		If .ActiveRow > 0 Then
			.focus
			.ReDraw = false
		
			ggoSpread.CopyRow
			SetSpreadColor .ActiveRow, .ActiveRow 
			
			'key field clear
			.Col = C_ModuleNm
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
    .vspdData.ReDraw = False
    .vspdData.focus
    ggoSpread.Source = .vspdData
    ggoSpread.InsertRow ,imRow
    SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
    
    For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1
        .vspdData.Row = iRow
        .vspdData.Col = C_RndUnit
        .vspdData.text = "0.1"
	    .vspdData.text = replace (.vspdData.text, ".", "@")		'***********************************020419수정 
	    .vspdData.text = replace (.vspdData.text, ",", "$")
	    .vspdData.text = replace (.vspdData.text, "@", parent.gComNumDec)
	    .vspdData.text = replace (.vspdData.text, "$", parent.gComNum1000)

        .vspdData.Col = C_DataFormat    
	    .vspdData.text = "##,###,###,###"				'******************020408
	    .vspdData.text = replace (.vspdData.text, ".", "@")		'***********************************020419수정 
	    .vspdData.text = replace (.vspdData.text, ",", "$")
	    .vspdData.text = replace (.vspdData.text, "@", parent.gComNumDec)			'Parent.gComNumDec
	    .vspdData.text = replace (.vspdData.text, "$", parent.gComNum1000)			'Parent.gComNum1000
	
	    .txtFormat.value = replace (.vspdData.text, ",", "$")			'******************020408
	    .vspdData.text = replace (.vspdData.text, "$", parent.gComNum1000)	'******************020408
    Next

    .vspdData.ReDraw = True
    
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
    ggoSpread.Source = frm1.vspdData 
    
	lDelRows = ggoSpread.DeleteRow
    
    End With
End Function

Function FncPrint() 
    Call parent.FncPrint()                                                  '☜: Protect system from crashing
End Function

Function FncPrev() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

Function FncNext() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)                                                   <%'☜: Protect system from crashing%>
End Function

Function FncFind()
	If gSelframeFlg = TAB1 Then	
		Call parent.FncFind(Parent.C_SINGLE, False)                                         <%'☜:화면 유형, Tab 유무 %>
	Else
		Call parent.FncFind(Parent.C_MULTI, False)                                         <%'☜:화면 유형, Tab 유무 %>
	End If
End Function

Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")			<%'⊙: "Will you destory previous data"%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

Function DbQuery() 
    
    DbQuery = False
    
    Call LayerShowHide(1)
    
    Err.Clear                                                               '☜: Protect system from crashing

	Dim strVal
    
    With frm1        
    
    If gSelframeFlg = Tab1 Then        
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001							'☜: 
			strVal = strVal & "&cboModuleCd=" & "*" 				'☆: 조회 조건 데이타 
			strVal = strVal & "&cboFormType=" & "I" 
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001							'☜: 
			strVal = strVal & "&cboModuleCd=" & "*"			     	'☆: 조회 조건 데이타 
			strVal = strVal & "&cboFormType=" & "I"
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If    
    Else    		
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001							'☜: 
			strVal = strVal & "&cboModuleCd=" & .hModuleCd.value 				'☆: 조회 조건 데이타 
			strVal = strVal & "&cboFormType=" & .hFormType.value 
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows			
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001							'☜: 
			strVal = strVal & "&cboModuleCd=" & Trim(.cboModuleCd.value)				'☆: 조회 조건 데이타 
			strVal = strVal & "&cboFormType=" & "Q"
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows			
		End If     
    End If   
    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    End With
    
    lgBlnFlgChgValue = False
        
    DbQuery = True    

End Function

Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    
    'Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
    
    If gSelframeFlg = TAB1 Then
		Call SetToolbar("1100100000011111")
    Else
		Call SetToolbar("1100111100111111")									'⊙: 버튼 툴바 제어 
    End If
    
    lgBlnFlgChgValue = False
	
	If gSelframeFlg = TAB1 Then
		frm1.cboModuleCd.value = "*"
		frm1.cboModuleCd.disabled = True
    Else
		If frm1.hModuleCd.value = "" Then
			frm1.cboModuleCd.value = ""
		End If		
		frm1.cboModuleCd.disabled = False
    End If	    
End Function

Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal, strDel
	
	Dim temp,strTempDataFormat
	
    DbSave = False                                                          '⊙: Processing is NG
    
    Call LayerShowHide(1)
    
    On Error Resume Next                                                   '☜: Protect system from crashing

	With frm1
		.txtMode.value = Parent.UID_M0002
		.txtInsrtUserId.value = Parent.gUsrID
		
    
    '-----------------------
    'Data manipulate area
    '-----------------------
    lGrpCnt = 1    
    strVal = ""
    strDel = ""
    
    '-----------------------
    'Data manipulate area
    '-----------------------
  If gSelframeFlg = TAB1 Then    '*************020408

		strVal = strVal & "U" & Parent.gColSep & "1" & Parent.gColSep				'☜: U=Update
		strVal = strVal & "*" & Parent.gColSep									'Module_cd
		strVal = strVal & "I" & Parent.gColSep									'Form Type
		strVal = strVal & Trim(.txtDec.value) & Parent.gColSep					'Decimals
		
		strVal = strVal & UNIConvNum(Trim(.txtFlag.value),0) & Parent.gColSep  'Round Unit     
		
		strVal = strVal & Trim(.cboFlag.value) & Parent.gColSep				'Round Policy
		
		strTempDataFormat = FormatChanging(Trim(.txtDec.value))         'Data Format
		strVal = strVal & strTempDataFormat & Parent.gRowSep      
		
		lGrpCnt = lGrpCnt + 1  

  Else  
    For lRow = 1 To .vspdData.MaxRows
    
        .vspdData.Row = lRow
        .vspdData.Col = 0
        
        Select Case .vspdData.Text

            Case ggoSpread.InsertFlag											'☜: 신규 
				
				strVal = strVal & "C" & Parent.gColSep & lRow & Parent.gColSep				'☜: C=Create

                .vspdData.Col = C_ModuleCd	'1
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                
                strVal = strVal & "Q" & Parent.gColSep
                                
                .vspdData.Col = C_Decimals	'5
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                
                .vspdData.Col = C_RndUnit	'6                
                strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep  'Round Unit     
                
                .vspdData.Col = C_RndPolicy	'7
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                                    
				.vspdData.Col = C_Decimals
				strTempDataFormat = FormatChanging(Trim(.vspdData.Text))
		        strVal = strVal & strTempDataFormat & Parent.gRowSep      
		
                                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.UpdateFlag											'☜: 신규 

				strVal = strVal & "U" & Parent.gColSep & lRow & Parent.gColSep				'☜: U=Update

                .vspdData.Col = C_ModuleCd	'1
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

				strVal = strVal & "Q" & Parent.gColSep
				
                .vspdData.Col = C_Decimals	'5
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                
                .vspdData.Col = C_RndUnit	'6
                strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep  'Round Unit     
                
                .vspdData.Col = C_RndPolicy	'7
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

				.vspdData.Col = C_Decimals
				strTempDataFormat = FormatChanging(Trim(.vspdData.Text))
		        strVal = strVal & strTempDataFormat & Parent.gRowSep      
                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.DeleteFlag											'☜: 삭제 

				strDel = strDel & "D" & Parent.gColSep & lRow & Parent.gColSep				'☜: D=Delete

                .vspdData.Col = C_ModuleCd	'1
                strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
                
                strDel = strDel & "Q" & Parent.gRowSep								
                
                lGrpCnt = lGrpCnt + 1
        End Select
                
    Next
  End If
  
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

Function CheckLogInUser() 
    Dim IntRetCD 
	Dim strLogInCnt
	Dim arrRet
	Dim arrParam(5)
    Dim tempMsg
    Dim iCalledAspName
    
	arrParam(0) = ""
	arrParam(1) = ""
    
    Err.Clear			
    strLogInCnt = Cint(frm1.txtLogInCnt.value)
    
    tempMsg = "접속중인 사용자가 존재하므로 저장할 수 없습니다 " & vbCrLf
    tempMsg = tempMsg & "이 자료는 시스템관리자 1명만 접속했을 때 저장할 수 있습니다" & vbCrLf
    tempMsg = tempMsg & "접속중인 사용자 정보를 보시겠습니까?"
      
    intRetCD = MsgBox(tempMsg,vbExclamation + vbYesNo, gLogoName & "-[Warning]")
    
    If IntRetCD = vbNo Then
		Exit Function
	End If

	iCalledAspName = AskPRAspName("LoginUserList")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "LoginUserList", "X")
		lgIsOpenPop = False
		Exit Function
	End If


	arrRet = window.showModalDialog(iCalledAspName,Array(window.parent,arrParam),, "dialogWidth=400px; dialogHeight=600px; center: Yes; help: No; resizable: No; status: No;")


End Function


Function cboFlag_onChange()
	lgBlnFlgChgValue = True
End Function

Function cboModuleCd_onChange()
    If gSelframeFlg = TAB2 Then
        iCurrentCD = frm1.cboModuleCd.value         
	End if	
End Function

Function DbDelete() 
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
					<TR>
						<TD WIDTH=10>&nbsp;</TD>
						<TD CLASS="CLSMTABP">
							<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
								<TR>
									<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/seltab_up_left.gif" width=9></td>
									<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>수량포맷(입력)</font></td>
									<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/seltab_up_right.gif" width=10></td>
							    </TR>
							</TABLE>
						</TD>
						<TD CLASS="CLSMTABP">
							<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
								<TR>
									<td background="../../../CShared/image/table/tab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/tab_up_left.gif" width=9></td>
									<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>수량포맷(조회)</font></td>
									<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/tab_up_right.gif" width=10></td>
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
					<FIELDSET CLASS="CLSFLD"><TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS="TD5">업무</TD>
						<TD CLASS="TD6" COLSPAN=3><SELECT NAME="cboModuleCd" tag="11X" STYLE="WIDTH: 160px;"><OPTION value=""></OPTION></SELECT></TD>
						<TD CLASS="TD5">화면종류</TD>
						<TD CLASS="TD6"><SELECT NAME="cboFormType"tag="14X" STYLE="WIDTH: 160px;"><OPTION value=""></OPTION></SELECT></TD>
					</TR>
						</TABLE></FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
	
				<!-- 첫번째 탭 내용 -->
					<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
					<TABLE <%=LR_SPACE_TYPE_60%>>
						<TR>
							<TD CLASS="TD5" NOWRAP>소수점자리수</TD>							
							<TD CLASS="TD656" NOWRAP>
							<script language =javascript src='./js/b1902ma1_txtDec_txtDec.js'></script></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>올림처리단위</TD>
							<TD CLASS=TD656 NOWRAP><INPUT NAME="txtFlag" ALT="올림처리단위" TYPE="Text" MAXLENGTH="18" SIZE=24 tag="24" STYLE="Text-Align:Right"></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>올림구분</TD>
							<TD CLASS=TD656 NOWRAP><SELECT NAME="cboFlag" tag="22X" STYLE="WIDTH: 180px;"></SELECT></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>포맷</TD>
							<TD CLASS=TD656 NOWRAP><INPUT NAME="txtFormat" ALT="포맷" TYPE="Text" MAXLENGTH="18" SIZE=24 tag="24" STYLE="Text-Align:Right"></TD>
						</TR>			
						<% Call SubFillRemBodyTD656(18)%>			
					</TABLE>
					</DIV>

				<!-- 두번째 탭 내용 -->
					<DIV ID="TabDiv" STYLE="DISPLAY: none" SCROLL=no>
					<TABLE <%=LR_SPACE_TYPE_20%>>
						<TR>
							<TD HEIGHT="100%">
								<script language =javascript src='./js/b1902ma1_vaSpread1_vspdData.js'></script>
							</TD>
						</TR>
					</TABLE>
					</DIV>	
		</TABLE></TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
			<TD WIDTH=100%>
				<TABLE <%=LR_SPACE_TYPE_30%>>
					<TR>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="vbscript:LoadCommonFormat">공통포맷</A>&nbsp;|&nbsp;<A HREF="vbscript:LoadNumericFormat">Numeric포맷</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
					</TR>
				</TABLE>
			</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="B1902mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hModuleCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hFormType" tag="24">
<INPUT TYPE=HIDDEN NAME="txtLogInCnt" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
