
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Daily Exchange Rate)(변동환율등록)
'*  3. Program ID           : B1702ma1.asp
'*  4. Program Name         : B1702ma1.asp
'*  5. Program Desc         :
'*  6. Modified date(First) : 2000/09/05
'*  7. Modified date(Last)  : 2002/12/11
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
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	<%'☜: indicates that All variables must be declared in advance%>

Const BIZ_PGM_ID = "B1702mb1.asp"												<%'비지니스 로직 ASP명 %>
 
 
 
Dim C_ValidDt
Dim C_Currency
Dim C_CurrencyPopup
Dim C_ToCurrency
Dim C_ToCurrencyPopup
Dim C_MultiDivide
Dim C_StdRate
Dim C_BuyRate
Dim C_SellRate
Dim C_CashBuyRate
Dim C_CashSellRate
Dim C_UsdRate
Dim C_Scope_Average

<% EndDate= GetSvrDate %>

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop          

Dim lgStrPrevToKey

Sub InitSpreadPosVariables()
    C_ValidDt           = 1  
    C_Currency          = 2
    C_CurrencyPopup     = 3
    C_ToCurrency        = 4
    C_ToCurrencyPopup   = 5
    C_MultiDivide       = 6
    C_StdRate           = 7
    C_BuyRate           = 8
    C_SellRate          = 9
    C_CashBuyRate       = 10
    C_CashSellRate      = 11
    C_UsdRate           = 12
    C_Scope_Average = 13
End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevToKey = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
End Sub

Sub SetDefaultVal()

    Dim strYear
    Dim strMonth
    Dim strDay
    

'	frm1.txtValidDt.Text = "<%=EndDate%>"
   
    Call ExtractDateFrom("<%= GetSvrDate %>",parent.gServerDateFormat , parent.gServerDateType      ,strYear,strMonth,strDay)

	frm1.txtValidDt.Year  = strYear
	frm1.txtValidDt.Month = strMonth
	frm1.txtValidDt.Day   = strDay
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

    .MaxCols = C_Scope_Average + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
	.Col = .MaxCols														'☆: 사용자 별 Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData
	
    Call GetSpreadColumnPos("A")  

    ggoSpread.SSSetDate C_ValidDt, "일자", 10, 2, parent.gDateFormat '1
    ggoSpread.SSSetEdit C_Currency, "기준통화", 10,,,3,2 '2
    ggoSpread.SSSetButton C_CurrencyPopup
    ggoSpread.SSSetEdit C_ToCurrency, "변환통화", 10,,,3,2 '4
    ggoSpread.SSSetButton C_ToCurrencyPopup
    ggoSpread.SSSetCombo C_MultiDivide, "계산구분자", 12 '6   
    ggoSpread.SSSetFloat C_StdRate,"기준환율",15,parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_BuyRate,"전신환매입율",15,parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_SellRate,"전신환매도율",15,parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_CashBuyRate,"현금매입율",15,parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_CashSellRate,"현금매도율",15,parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_UsdRate,"달러환산율",15,parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat C_Scope_Average,"환율오차범위",15,parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	
    call ggoSpread.MakePairsColumn(C_Currency,C_CurrencyPopup)
    call ggoSpread.MakePairsColumn(C_ToCurrency,C_ToCurrencyPopup)

	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub

Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False
    ggoSpread.SpreadLock C_ValidDt, -1, C_ValidDt
	ggoSpread.SpreadLock C_Currency, -1, C_Currency
	ggoSpread.SpreadLock C_CurrencyPopup, -1, C_CurrencyPopup
	ggoSpread.SpreadLock C_ToCurrency, -1, C_ToCurrency
	ggoSpread.SpreadLock C_ToCurrencyPopup, -1, C_ToCurrencyPopup
	ggoSpread.SSSetRequired C_MultiDivide, -1, -1
	ggoSpread.SSSetRequired C_StdRate, -1, -1
	ggoSpread.SSSetProtected .vspdData.MaxCols, -1, -1
    .vspdData.ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SSSetRequired C_ValidDt, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired C_Currency, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired C_ToCurrency, pvStartRow, pvEndRow	
	ggoSpread.SSSetRequired C_MultiDivide, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired C_StdRate, pvStartRow, pvEndRow
    .vspdData.ReDraw = True
    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_ValidDt          = iCurColumnPos(1) 
            C_Currency         = iCurColumnPos(2)
            C_CurrencyPopup    = iCurColumnPos(3)
            C_ToCurrency       = iCurColumnPos(4)
            C_ToCurrencyPopup  = iCurColumnPos(5)
            C_MultiDivide      = iCurColumnPos(6)
            C_StdRate          = iCurColumnPos(7)
            C_BuyRate          = iCurColumnPos(8)
            C_SellRate         = iCurColumnPos(9)
            C_CashBuyRate      = iCurColumnPos(10)
            C_CashSellRate     = iCurColumnPos(11)
            C_UsdRate          = iCurColumnPos(12)
            C_Scope_Average = iCurColumnPos(13)
    End Select    
End Sub

Sub InitSpreadComboBox()
	ggoSpread.SetCombo "*" & vbTab & "/", C_MultiDivide
End Sub

Function OpenCurrency(iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "통화 팝업"						<%' 팝업 명칭 %>
	arrParam(1) = "b_currency"						<%' TABLE 명칭 %>
	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(2) = Trim(frm1.txtCurrency.value)	<%' Code Condition%>
	Else 'spread
		arrParam(2) = Trim(frm1.vspdData.Text)		<%' Code Condition%>
	End If
	arrParam(3) = ""								<%' Name Cindition%>
	arrParam(4) = ""								<%' Where Condition%>
	arrParam(5) = "통화"							<%' 조건필드의 라벨 명칭 %>
	
    arrField(0) = "currency"						<%' Field명(0)%>
    arrField(1) = "currency_desc"					<%' Field명(1)%>
    
    arrHeader(0) = "통화"							<%' Header명(0)%>
    arrHeader(1) = "비고"							<%' Header명(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtCurrency.focus 
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCurrency(arrRet, iWhere)
	End If	
	
End Function

Function SetCurrency(Byval arrRet, Byval iWhere)
	With frm1
		If iWhere = 0 Then 'TextBox(Condition)
			.txtCurrency.value = arrRet(0)
		Else 'spread
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
    
    Call SetDefaultVal
    Call InitSpreadComboBox
    Call SetToolbar("1100110100101111")										<%'버튼 툴바 제어 %>
    frm1.txtValidDt.focus
    
End Sub

Sub txtValidDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtValidDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtValidDt.Focus
    End If
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
End Sub

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With frm1.vspdData 
		If Row > 0 And Col = C_CurrencyPopup Then
		    .Row = Row
		    .Col = C_Currency

		    Call OpenCurrency(1)
		ElseIf Row > 0 And Col = C_ToCurrencyPopup Then
		    .Row = Row
		    .Col = C_ToCurrency

		    Call OpenCurrency(1)
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
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) _
	And Not(lgStrPrevKey = "" And lgStrPrevToKey = "") Then
		Call DisableToolBar(parent.TBC_QUERY)					'☜ : Query 버튼을 disable 시킴.
		If DBQuery = False Then 
		   Call RestoreToolBar()
		   Exit Sub 
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
    If Not ggoSpread.SSDefaultCheck Then     'Not chkField(Document, "2") OR      '⊙: Check contents area
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
			
			ggoSpread.CopyRow
            SetSpreadColor .ActiveRow, .ActiveRow    
			'Key field clear
			.Col = C_Currency
			.Text = ""
			
			.Col = C_ToCurrency
			.Text = ""

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
		    
		    .vspdData.Col = C_ValidDt
		    .vspdData.Text = .txtValidDt.text

		    .vspdData.Col = C_MultiDivide
		    .vspdData.Text = "*"
		
		    .vspdData.ReDraw = True
		Next
		
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

    DbQuery = False
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
	
	Call LayerShowHide(1)
	
	Dim strVal
    
    With frm1
    
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode="    & parent.UID_M0001							'☜: 
		strVal = strVal & "&txtValidDt="     & Trim(.hValidDt.value)        '☆: 조회 조건 데이타 
		strVal = strVal & "&txtCurrency="    & .hCurrency.value 			'☆: 조회 조건 데이타		
		strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey
		strVal = strVal & "&lgStrPrevToKey=" & lgStrPrevToKey
    Else
		strVal = BIZ_PGM_ID & "?txtMode="    & parent.UID_M0001							'☜: 
		'strVal = strVal & "&txtValidDt="     & UniConvDateToYYYYMMDD(.txtValidDt.Text,parent.gDateFormat,parent.gServerDateType)
		strVal = strVal & "&txtValidDt="     & .txtValidDt.Text
		strVal = strVal & "&txtCurrency="    & Trim(.txtCurrency.value)			'☆: 조회 조건 데이타		
		strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey
		strVal = strVal & "&lgStrPrevToKey=" & lgStrPrevToKey
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
	DIm strCur, strToCur
	
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
		        Case ggoSpread.InsertFlag
					strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep		'☜: C=Create
		        Case ggoSpread.UpdateFlag
					strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep		'☜: U=Update
			End Select			

		    Select Case .vspdData.Text
		        Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag		'☜: 신규, 수정 
		        
		            .vspdData.Col = C_ValidDt		'1
		            strVal = strVal & UniConvDateAToB(Trim(.vspdData.text),parent.gDateFormat,parent.gServerDateFormat) & parent.gColSep
		            
		             
		            .vspdData.Col = C_Currency		'2
		            strCur = Trim(.vspdData.Text)
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ToCurrency	'4
		            strToCur = Trim(.vspdData.Text)
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep		            
		            IF strCur = StrToCur Then
						Call DisplayMsgBox("121404", "X", "X", "X")
						
					    .vspdData.Row = lRow
						.vspdData.Action = 0

					    Call LayerShowHide(0)
						Exit Function
					End If
					
					.vspdData.Col = C_MultiDivide	'6
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_StdRate		'7
		            strVal = strVal & .vspdData.Value & parent.gColSep
					If .vspdData.value <= 0 Then
						.vspdData.Row = 0
						Call DisplayMsgBox("970022", "X", .vspdData.Text, "0")
						
					    .vspdData.Row = lRow
						.vspdData.Action = 0

					    Call LayerShowHide(0)
						Exit Function
					End If
					
					.vspdData.Col = C_BuyRate		'8
		            
					If Trim(.vspdData.value) = "" Then 					
		            	strVal = strVal & "0" & parent.gColSep
		            Else
		            	If .vspdData.value < 0 Then					
							.vspdData.Row = 0
							Call DisplayMsgBox("970022", "X", .vspdData.Text, "0")
						
					    	.vspdData.Row = lRow
							.vspdData.Action = 0

					    	Call LayerShowHide(0)
							Exit Function
						Else
							strVal = strVal & .vspdData.Value & parent.gColSep
						End If				
		            End If
					
					.vspdData.Col = C_SellRate		'9
		            
					If Trim(.vspdData.value) = "" Then 					
		            	strVal = strVal & "0" & parent.gColSep
		            Else
		            	If .vspdData.value < 0 Then					
							.vspdData.Row = 0
							Call DisplayMsgBox("970022", "X", .vspdData.Text, "0")
						
					    	.vspdData.Row = lRow
							.vspdData.Action = 0

					    	Call LayerShowHide(0)
							Exit Function
						Else
							strVal = strVal & .vspdData.Value & parent.gColSep
						End If				
		            End If
					
					.vspdData.Col = C_CashBuyRate		'10
		            
					If Trim(.vspdData.value) = "" Then 					
		            	strVal = strVal & "0" & parent.gColSep
		            Else
		            	If .vspdData.value < 0 Then					
							.vspdData.Row = 0
							Call DisplayMsgBox("970022", "X", .vspdData.Text, "0")
						
					    	.vspdData.Row = lRow
							.vspdData.Action = 0

					    	Call LayerShowHide(0)
							Exit Function
						Else
							strVal = strVal & .vspdData.Value & parent.gColSep
						End If				
		            End If
					
					.vspdData.Col = C_CashSellRate		'11
		            
					If Trim(.vspdData.value) = "" Then 					
		            	strVal = strVal & "0" & parent.gColSep
		            Else
		            	If .vspdData.value < 0 Then					
							.vspdData.Row = 0
							Call DisplayMsgBox("970022", "X", .vspdData.Text, "0")
						
					    	.vspdData.Row = lRow
							.vspdData.Action = 0

					    	Call LayerShowHide(0)
							Exit Function
						Else
							strVal = strVal & .vspdData.Value & parent.gColSep
						End If				
		            End If
					
					.vspdData.Col = C_UsdRate		'12
		            
					If Trim(.vspdData.value) = "" Then 					
		            	strVal = strVal & "0" & parent.gRowSep
		            Else
		            	If .vspdData.value < 0 Then					
							.vspdData.Row = 0
							Call DisplayMsgBox("970022", "X", .vspdData.Text, "0")
						
					    	.vspdData.Row = lRow
							.vspdData.Action = 0

					    	Call LayerShowHide(0)
							Exit Function
						Else
							strVal = strVal & .vspdData.Value & parent.gColSep
						End If				
		            End If

					.vspdData.Col = C_Scope_Average		'13
		            
					If Trim(.vspdData.value) = "" Then 					
		            	strVal = strVal & "0" & parent.gRowSep
		            Else
		            	If .vspdData.value < 0 Then					
							.vspdData.Row = 0
							Call DisplayMsgBox("970022", "X", .vspdData.Text, "0")
						
					    	.vspdData.Row = lRow
							.vspdData.Action = 0

					    	Call LayerShowHide(0)
							Exit Function
						Else
							strVal = strVal & .vspdData.Value & parent.gRowSep
						End If				
		            End If
		            
		            lGrpCnt = lGrpCnt + 1

		        Case ggoSpread.DeleteFlag							'☜: 삭제 

					strDel = strDel & "D" & parent.gColSep & lRow & parent.gColSep					'☜: U=Update

		            .vspdData.Col = C_ValidDt		'1
		            strDel = strDel & UniConvDateAToB(Trim(.vspdData.text),parent.gDateFormat,parent.gServerDateFormat) & parent.gColSep

		            .vspdData.Col = C_Currency		'3
		            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ToCurrency	'5
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

Sub txtValidDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 13 Then
		Call MainQuery()
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>변동환율</font></td>
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
									<TD CLASS="TD5">일자</TD>
									<TD CLASS="TD6">
										<script language =javascript src='./js/b1702ma1_I941833829_txtValidDt.js'></script>
									</TD>
									<TD CLASS="TD5">기준통화</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCurrency" SIZE=16 MAXLENGTH=3 tag="11XXXU"  ALT="기준통화"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCurrency" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCurrency(0)">
									<div style="display:none"> <INPUT TYPE=TEXT NAME="dummy"></div>
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
									<script language =javascript src='./js/b1702ma1_vaSpread1_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="B1702mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hValidDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hCurrency" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

