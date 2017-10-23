
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : Master Data
'*  2. Function Name        : Master Data
'*  3. Program ID           : B1201ma1.asp
'*  4. Program Name         : B1201ma1.asp
'*  5. Program Desc         : 자동채번등록 
'*  6. Comproxy list        : 
'*  7. Modified date(First) : 2000/09/04
'*  8. Modified date(Last)  : 2002/12/10
'*  9. Modifier (First)     : Hwang Jeong-won
'* 10. Modifier (Last)      : Sim Hae Young
'* 11. Comment              :
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

Const BIZ_PGM_ID = "B1201mb1.asp"                                                 '비지니스 로직 ASP명 

<!-- #Include file="../../inc/lgvariables.inc" -->
 
Dim C_NumberingCd        
Dim C_NumberingTypePopUp 
Dim C_AutoNumbering      
Dim C_NumberingNm        
Dim C_MaxLen             
Dim C_ValidDt            
Dim C_PrefixCd           
Dim C_DateType           
Dim C_Inc                
Dim C_NumMaxLen          
Dim C_Detail             
Dim C_DateInfo           
Dim C_Number             
Dim C_LastNum            


Dim IsOpenPop          
Dim lgStrPrevDt

Sub InitSpreadPosVariables()
    C_NumberingCd        = 1
    C_NumberingTypePopUp = 2
    C_AutoNumbering      = 3
    C_NumberingNm        = 4
    C_MaxLen             = 5
    C_ValidDt            = 6
    C_PrefixCd           = 7
    C_DateType           = 8
    C_Inc                = 9
    C_NumMaxLen          = 10
    C_Detail             = 11
    C_DateInfo           = 12
    C_Number             = 13
    C_LastNum            = 14
End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevDt = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
End Sub

Sub SetDefaultVal()
    frm1.txtValidDt.text = UniConvDateAToB(ReadAutoNumeringEffectDt(),parent.gServerDateFormat,parent.gDateFormat)
					'iDate = UniConvDateAToB("<%=GetSvrDate%>" ,parent.gServerDateFormat,parent.gDateFormat)    
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "B","NOCOOKIE","MA") %>
End Sub

Sub InitSpreadSheet()

    Call initSpreadPosVariables()  

	With frm1.vspdData
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20021203",,parent.gAllowDragDropSpread    
        
	.ReDraw = false

    .MaxCols = C_LastNum + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
	.Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
    .ColHidden = True
    
    .MaxRows = 0
    ggoSpread.ClearSpreadData

    Call GetSpreadColumnPos("A")

    ggoSpread.SSSetEdit    C_NumberingCd   , "채번유형코드"    , 15,,,2,2 '1
    ggoSpread.SSSetButton  C_NumberingTypePopUp '2
	ggoSpread.SSSetCheck   C_AutoNumbering , "자동채번여부"    , 14, 2, "자동", True '3
    ggoSpread.SSSetEdit    C_NumberingNm   , "채번유형"        , 16 '4
    ggoSpread.SSSetFloat   C_MaxLen        , "최대길이"        , 20,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"1","18"
    ggoSpread.SSSetDate    C_ValidDt       , "적용시작일"      , 10, 2, parent.gDateFormat '6    
    ggoSpread.SSSetEdit    C_PrefixCd      , "Prefix코드"      , 10,,,2,2 '7
    ggoSpread.SSSetCombo   C_DateType      , "날짜유형"        , 12 '8
   	ggoSpread.SSSetFloat   C_Inc           , "채번증가수"      , 10,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"1","99"
    ggoSpread.SSSetFloat   C_NumMaxLen     , "숫자최대길이"    , 18,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"1","15"
    ggoSpread.SSSetEdit    C_Detail        , "상세정보"        , 20,,,50 '11
    ggoSpread.SSSetEdit    C_DateInfo      , "날짜정보"        , 10, 2 '12
	ggoSpread.SSSetFloat   C_Number        , "숫자"            ,  8,"7",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetEdit    C_LastNum       , "최종자동채번번호", 20 '14
	
	call ggoSpread.MakePairsColumn(C_NumberingCd,C_NumberingTypePopUp)

    Call ggoSpread.SSSetColHidden(C_AutoNumbering,C_AutoNumbering,True)
	
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub

Sub SetSpreadLock()
    With frm1
	    .vspdData.ReDraw = False
		ggoSpread.SpreadLock C_NumberingCd, -1, C_NumberingCd 
		ggoSpread.SpreadLock C_NumberingTypePopUp, -1, C_NumberingTypePopUp 
		ggoSpread.SpreadLock C_NumberingNm, -1, C_NumberingNm 
		ggoSpread.SpreadLock C_MaxLen, -1, C_MaxLen 
		ggoSpread.SpreadLock C_ValidDt, -1, C_ValidDt 
		ggoSpread.SpreadLock C_PrefixCd, -1, C_PrefixCd 
		ggoSpread.SpreadLock C_DateType, -1, C_DateType 
		ggoSpread.SpreadLock C_Inc, -1, C_Inc 
		ggoSpread.SpreadLock C_NumMaxLen, -1, C_NumMaxLen 
		ggoSpread.SpreadLock C_Detail, -1, C_Detail 
		ggoSpread.SpreadLock C_DateInfo, -1, C_DateInfo 
		ggoSpread.SpreadLock C_Number, -1, C_Number 
		ggoSpread.SpreadLock C_LastNum, -1, C_LastNum 
		ggoSpread.SSSetProtected .vspdData.MaxCols, -1, -1
	    .vspdData.ReDraw = True
    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
		ggoSpread.SSSetRequired C_AutoNumbering, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_NumberingCd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_NumberingNm, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_ValidDt, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_PrefixCd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_MaxLen, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_Inc, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_NumMaxLen, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_DateInfo, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Number, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_LastNum, pvStartRow, pvEndRow
    .vspdData.ReDraw = True
    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
            C_NumberingCd        = iCurColumnPos(1)
            C_NumberingTypePopUp = iCurColumnPos(2)
            C_AutoNumbering      = iCurColumnPos(3)
            C_NumberingNm        = iCurColumnPos(4)
            C_MaxLen             = iCurColumnPos(5)
            C_ValidDt            = iCurColumnPos(6)
            C_PrefixCd           = iCurColumnPos(7)
            C_DateType           = iCurColumnPos(8)
            C_Inc                = iCurColumnPos(9)
            C_NumMaxLen          = iCurColumnPos(10)
            C_Detail             = iCurColumnPos(11)
            C_DateInfo           = iCurColumnPos(12)
            C_Number             = iCurColumnPos(13)
            C_LastNum            = iCurColumnPos(14)
            
    End Select    
End Sub

Sub InitSpreadComboBox()
ggoSpread.SetCombo "" & vbTab _
					 & "YYYYMMDD" & vbTab _
					 & "YYMMDD" & vbTab _
					 & "YYYYMM" & vbTab _
					 & "YYMM" & vbTab _
					 & "YMMDD" & vbTab _
					 & "YYYY" & vbTab _
					 & "YY", C_DateType					 					 
End Sub

Function OpenNumberingType(iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "채번유형 팝업"				  ' 팝업 명칭 
	arrParam(1) = "b_minor a, b_configuration c"	  ' TABLE 명칭 
	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(2) = frm1.txtMinor.value			  ' Code Condition
		''arrParam(2) = FilterVar(frm1.txtMinor.value,"''","S") 
	Else 'spread
		arrParam(2) = frm1.vspdData.Text			  ' Code Condition
	End If
	arrParam(3) = ""								  ' Name Cindition
	arrParam(4) = " a.major_cd = " & FilterVar("B0006", "''", "S") & " "_
	              & " And c.major_cd =* a.major_cd And c.minor_cd =* a.minor_cd"_
	              & " And c.seq_no = 1"				  ' Where Condition
	arrParam(5) = "채번유형"					  ' 조건필드의 라벨 명칭 
	
    arrField(0) = "a.minor_cd"						  ' Field명(0)
    arrField(1) = "a.minor_nm"						  ' Field명(1)
    arrField(2) = "Case c.reference When null Then " & FilterVar("18", "''", "S") & "  Else c.reference End"		' Field명(2)
    
    arrHeader(0) = "채번유형코드"				  ' Header명(0)
    arrHeader(1) = "채번유형"					  ' Header명(1)
    arrHeader(2) = "채번최대길이"				  ' Header명(2)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtMinor.focus
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetNumberingType(arrRet, iWhere)
	End If	
			
End Function

Function SetNumberingType(Byval arrRet, Byval iWhere)
		
	With frm1
		If iWhere = 0 Then 'TextBox(Condition)
			.txtMinor.value = arrRet(0)
			.txtMinorNm.value = arrRet(1)
		Else 'spread
			.vspdData.Col = C_NumberingCd
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_NumberingNm
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_MaxLen
			If IsEmpty(arrRet(2)) Or arrRet(2) = "" Then
				.vspdData.Text = "18"
			Else
				.vspdData.Text = arrRet(2)
			End If
			.vspdData.Col = C_PrefixCd
			.vspdData.Text = arrRet(0)
			
			lgBlnFlgChgValue = True
		End If
	End With
End Function

Sub Form_Load()
    Call LoadInfTB19029                                                     <%'Load table , B_numeric_format%>
    
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetDefaultVal
    Call InitSpreadComboBox
    Call SetToolbar("1100110100101111")										<%'버튼 툴바 제어 %>
    frm1.txtMinor.focus
    
End Sub

Sub txtValidDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtValidDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtValidDt.focus
    End If
End Sub

Sub vspdData_Change(ByVal Col , ByVal Row )
   Frm1.vspdData.Row = Row
   Frm1.vspdData.Col = Col

   If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If CDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
   End If
   
   ggoSpread.Source = frm1.vspdData
   ggoSpread.UpdateRow Row

End Sub

Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101111111") 
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

Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

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

	With frm1.vspdData 
	
    ggoSpread.Source = frm1.vspdData
   
    If Row > 0 And Col = C_NumberingTypePopUp Then
		    .Row = Row
		    .Col = C_NumberingCd

		    Call OpenNumberingType(1)        
    End If
    
    End With
End Sub

Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    
    If CheckRunningBizProcess = True Then					'☜: 조회중이면 다음 조회 안하도록 체크 
        Exit Sub
	End If
	
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

End Sub

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               <%'Protect system from crashing%>

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    
    Call ggoOper.ClearField(Document, "2")									<%'Clear Contents  Field%>
    ggoSpread.ClearSpreadData
    Call InitVariables                                                      <%'Initializes local global variables%>
    															
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If
    
    If DbQuery = False Then Exit Function
       
    FncQuery = True															
    
End Function

Function FncSave() 
    On Error Resume Next                                                       '☜: Protect system from crashing

    FncSave = False                                                         
    
    Err.Clear                                                                  '☜: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData                                           'Precheck area
    If ggoSpread.SSCheckChange = False Then
        Call DisplayMsgBox("900001", "X", "X", "X")                            'No data changed!!
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData                                           'Check content area
    If Not ggoSpread.SSDefaultCheck Then                                       '⊙: Check contents area
       Exit Function
    End If
    
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
    
			.Col = C_NumberingCd : .Text = ""      'Key field clear
			.Col = C_NumberingNm : .Text = ""
			.Col = C_maxLen      : .Text = ""
			.Col = C_PrefixCd    : .Text = ""
			.Col = C_DateInfo    : .Text = ""
			.Col = C_Number      : .Text = ""
			.Col = C_LastNum     : .Text = ""
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
    	    .vspdData.Col = C_AutoNumbering
		    .vspdData.Value = 1

		    .vspdData.Col = C_ValidDt
		    .vspdData.Text = UNIFormatDate(Date)
      
		    .vspdData.Col = C_Inc
		    .vspdData.Text = 1
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
	Dim strVal

	Call LayerShowHide(1)
	
    DbQuery = False
    
    Err.Clear

    frm1.txtMinorNm.value = ReadAutoNumeringTypeName(frm1.txtMinor.value)

    With frm1    
        If lgIntFlgMode = parent.OPMD_UMODE Then
		   strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 
		   strVal = strVal & "&txtMinor="     & .hMinor.value 			'☆: 조회 조건 데이타 
           strVal = strVal & "&txtValidDt="   & Trim(.hValidDt.value)        
           strVal = strVal & "&txtMaxRows="   & .vspdData.MaxRows
	    Else
		    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 
           strVal = strVal & "&txtMinor="     & Trim(.txtMinor.value)	'☆: 조회 조건 데이타 
           strVal = strVal & "&txtValidDt="   & Trim(.txtValidDt.Text)        
           strVal = strVal & "&txtMaxRows="   & .vspdData.MaxRows
        End If
    
	    Call RunMyBizASP(MyBizASP, strVal)								<%'☜: 비지니스 ASP 를 가동 %>        
    End With
    
    DbQuery = True
    
End Function

Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE

	Call SetToolbar("110011110011111")										<%'버튼 툴바 제어 %>
	
End Function

Function DbSave() 
    Dim lRow        
    Dim lGrpCnt  
	Dim strVal, strDel
	Dim a, b
	
    DbSave = False                                                          
    
    Call LayerShowHide(1)
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>

	With frm1
		.txtMode.value = parent.UID_M0002
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
    
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
		        Case ggoSpread.InsertFlag
					strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep '☜: C=Create, Row위치 정보 

				    'Validation Check
					a=0
					b=0
				
					.vspdData.Col = C_PrefixCd		'*************  창 020304
					a = Len(Trim(.vspdData.Text))
					
					If a < 1 Then    '*************  창 020304 Prefix자리수에서 Tracking Number을 한자리로 맞추는 과정에서 수정완료 
						Call DisplayMsgBox("120706", "X", "X", "X")

					    .vspdData.Row = lRow
					    .vspdData.Action = 0 'ActionActiveCell
					    .vspdData.EditMode = True
					    Call LayerShowHide(0)
					    Exit Function
					End If
					
					.vspdData.Col = C_DateType
					b = Len(Trim(.vspdData.Text))
					a = a + b

					.vspdData.Col = C_NumMaxLen
					If Trim(.vspdData.Text) <> "" Then
						b = CInt(.vspdData.Text)
						If b > 18 - a Or b = 0 Then
							.vspdData.Row = 0
						    Call DisplayMsgBox("970025", "X", .vspdData.Text, CStr(18 - a))

						    .vspdData.Row = lRow
						    .vspdData.Action = 0 'ActionActiveCell
						    .vspdData.EditMode = True
						    Call LayerShowHide(0)
						    Exit Function
						End If
					End If
					'End Validation Check

		        Case ggoSpread.UpdateFlag
					strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep '☜: U=Update, Row위치 정보 
			End Select			

		    .vspdData.Row = lRow
		    .vspdData.Col = 0
		    
		    Select Case .vspdData.Text
		        Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag		'☜: 신규, 수정 

		            .vspdData.Col = C_AutoNumbering	'1
	                If .vspdData.Value = 1 Then
			            strVal = strVal & "O" & parent.gColSep
			        Else
			            strVal = strVal & "X" & parent.gColSep
					End If
		            
		            .vspdData.Col = C_NumberingCD	'2
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ValidDt		'6
		            strVal = strVal & UNIConvDate(Trim(.vspdData.text)) & parent.gColSep

		            .vspdData.Col = C_PrefixCd		'7
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

		            .vspdData.Col = C_DateType		'8
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

		            .vspdData.Col = C_Inc			'9
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

		            .vspdData.Col = C_NumMaxLen		'10
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

		            .vspdData.Col = C_Detail		'11
		            strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep

		            lGrpCnt = lGrpCnt + 1

		        Case ggoSpread.DeleteFlag							'☜: 삭제 

					strDel = strDel & "D" & parent.gColSep	& lRow & parent.gColSep 
					
					.vspdData.Col = C_AutoNumbering	'1
	                If .vspdData.Value = 1 Then
			            strDel = strDel & "O" & parent.gColSep
			        Else
			            strDel = strDel & "X" & parent.gColSep
					End If
					
		            .vspdData.Col = C_NumberingCD	'2
		            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ValidDt	'6
		            strDel = strDel & UNIConvDate(Trim(.vspdData.text)) & parent.gRowSep
  
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

Sub txtValidDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Sub

Function ReadAutoNumeringEffectDt()
    Dim iRet
    Dim iEffectDt

    If CommonQueryRs(" max(CONVERT(CHAR(10),EFFECT_FROM_DT, 20)) "," B_AUTO_NUMBERING ","EFFECT_FROM_DT <= getdate() " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) =  True Then
       iEffectDt                = Split(lgF0,Chr(11))
       ReadAutoNumeringEffectDt = iEffectDt(0)
    Else
       ReadAutoNumeringEffectDt = ""
    End If
End Function

Function ReadAutoNumeringTypeName(pMinorCd)
    Dim iRet
    Dim strSQL
    Dim iMinorMn
    
    strSQL = "a.major_cd = b.major_cd "
    strSQL = strSQL  & "and  a.major_cd = " & FilterVar("B0006", "''", "S") & " "
    strSQL = strSQL  & "and  b.minor_cd =  " & FilterVar(pMinorCd , "''", "S") & ""

    If CommonQueryRs(" b.minor_nm "," b_major a, b_minor b ",strSQL ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) =  True Then
       iMinorMn =  Split(lgF0,Chr(11))
       ReadAutoNumeringTypeName =  iMinorMn(0)
    Else
       ReadAutoNumeringTypeName = ""
    End If
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
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>자동채번</font></td>
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
									<TD CLASS="TD5">채번유형</TD>
									<TD CLASS="TD6">
										<INPUT TYPE=TEXT NAME="txtMinor" SIZE=10 MAXLENGTH=2 tag="11XXXU" ALT="채번유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAutoNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenNumberingType(0)">
										<INPUT TYPE=TEXT NAME="txtMinorNm" tag="14X">
									</TD>
									<TD CLASS="TD5">적용시작일</TD>
									<TD CLASS="TD6">
									<script language =javascript src='./js/b1201ma1_I194613635_txtValidDt.js'></script>
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
									<script language =javascript src='./js/b1201ma1_vaSpread1_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="B1201mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hMinor" tag="24"><INPUT TYPE=HIDDEN NAME="hValidDt" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

