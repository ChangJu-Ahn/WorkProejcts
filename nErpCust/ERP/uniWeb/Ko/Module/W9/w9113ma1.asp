
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 기타 서식 
'*  3. Program ID           : W1113MA1
'*  4. Program Name         : W1113MA1.asp
'*  5. Program Desc         : 제 55호 소득자료명세서 
'*  6. Modified date(First) : 2005/01/27
'*  7. Modified date(Last)  : 2006/02/02
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : HJO 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%
'=========================  로긴중인 유저의 법인코드를 출력하기 위해  ======================
    Call LoadBasisGlobalInf()
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliGrid.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../WCM/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  상수/변수 선언  ====================================

<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_MNU_ID = "w9113ma1"
Const BIZ_PGM_ID = "w9113mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const EBR_RPT_ID = "w9113oa1"
Dim C_SEQ_NO
Dim C_W1_CD
Dim C_W1
Dim C_W2
Dim C_W3
Dim C_W4
Dim C_W5
Dim C_W6
Dim C_W7
Dim C_W8

Dim IsOpenPop    
Dim gSelframeFlg , lgCurrGrid      
Dim lgStrPrevKey2
Dim	lgFISC_START_DT, lgFISC_END_DT

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()
	C_SEQ_NO			= 1
    C_W1_CD				= 2
    C_W1				= 3
    C_W2				= 4
    C_W3				= 5
    C_W4				= 6
    C_W5				= 7
    C_W6				= 8
    C_W7				= 9
    C_W8				= 10
End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevKey2 = ""                          'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1

End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub



'============================================  신고구분 콤보 박스 채우기  ====================================

Sub InitComboBox()
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
End Sub

Sub InitSpreadSheet()
	Dim ret
		
	Call initSpreadPosVariables()  

	With frm1.vspdData
		
		ggoSpread.Source = frm1.vspdData	
		'patch version
		ggoSpread.Spreadinit "V20041222",,parent.gForbidDragDropSpread    
			 
		.ReDraw = false

		'헤더를 2줄로    
		.ColHeaderRows = 2   
			 
		.MaxCols = C_W8 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols														'☆: 사용자 별 Hidden Column
		.ColHidden = True    
				       
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		ggoSpread.SSSetEdit		C_SEQ_NO, "순번", 10,,,100,1
		ggoSpread.SSSetCombo	C_W1_CD,	"코드"    , 15
		ggoSpread.SSSetCombo	C_W1,		"(1)소득구분"    , 15
		ggoSpread.SSSetDate		C_W2,		"(2)사업연도",	10,		2,		Parent.gDateFormat,	-1
		ggoSpread.SSSetMask		C_W3,		"(3)소득귀속"	, 10, 2, "9999년" 
		ggoSpread.SSSetFloat	C_W4,		"(4)배당.상여및" & vbCrLf & "기타소득금액",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
		ggoSpread.SSSetFloat	C_W5,		"(5)원천징수할" & vbCrLf & "소득세액" ,		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
		ggoSpread.SSSetEdit		C_W6,		"(6)주소", 25,,,100,1
		ggoSpread.SSSetEdit		C_W7,		"(7)성명", 10,,,10
		ggoSpread.SSSetMask		C_W8,		"(8)주민등록번호" & vbCrlf & "일수"	, 15, 2, "999999-9999999" 
				   
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_W1_CD,True)
							
		Call InitSpreadComboBox()

		' 그리드 헤더 합침 정의 
		ret = .AddCellSpan(C_SEQ_NO		, -1000, 1, 2)	' SEQ_NO 합침 
		ret = .AddCellSpan(C_W1			, -1000, 1, 2)	
		ret = .AddCellSpan(C_W2			, -1000, 1, 2)	
		ret = .AddCellSpan(C_W3			, -1000, 1, 2)	
		ret = .AddCellSpan(C_W4			, -1000, 1, 2)	
		ret = .AddCellSpan(C_W5			, -1000, 1, 2)
		ret = .AddCellSpan(C_W6			, -1000, 3, 1)
			
		' 첫번째 헤더 출력 글자 
		.Row = -1000
		.Col = C_W6
		.Text = "소 득 자"
					
		' 두번째 헤더 출력 글자 
		.Row = -999	
		.Col = C_W6
		.Text = "(6)주소"
		.Col = C_W7
		.Text = "(7)성명"
		.Col = C_W8	
		.Text = "(8)주민등록번호"

		.rowheight(-999) = 12	' 높이 재지정 

		.ReDraw = true
			
		Call SetSpreadLock 
	    
	End With

End Sub


'============================================  그리드 함수  ====================================

Sub InitSpreadComboBox()
    Dim IntRetCD1

	' 차/대변 
	IntRetCD1 = CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", " MAJOR_CD='W1005' ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  

	If IntRetCD1 <> False Then
		ggoSpread.Source = frm1.vspddata
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_W1_CD
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_W1
	End If

End Sub

Sub SetSpreadLock()
    With frm1.vspdData

		.ReDraw = False
		ggoSpread.SpreadLock C_SEQ_NO, 1, C_SEQ_NO, .MaxRows	' 전체 적용 
		ggoSpread.SpreadLock C_SEQ_NO, .MaxRows, C_W8, .MaxRows	' 전체 적용 
		ggoSpread.SSSetRequired C_W1, 1, .MaxRows-1
		.ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1

    .vspdData.ReDraw = False
 
  	ggoSpread.SSSetRequired C_W1, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_SEQ_NO, pvStartRow, pvEndRow
      
    .vspdData.ReDraw = True
    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_SEQ_NO			= iCurColumnPos(1)
            C_W1_CD				= iCurColumnPos(2)
            C_W1				= iCurColumnPos(3)
            C_W2				= iCurColumnPos(4)
            C_W3				= iCurColumnPos(5)
            C_W4				= iCurColumnPos(6)
            C_W5				= iCurColumnPos(7)
            C_W6				= iCurColumnPos(8)
            C_W7				= iCurColumnPos(9)
            C_W8				= iCurColumnPos(10)

    End Select    
End Sub

Sub InitData()
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

	Call GetFISC_DATE
End Sub

Sub SetSpreadTotalLine()
	Dim iRow
	ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
		If .MaxRows > 0 Then
			.Row = .MaxRows
			.Col = C_W1 : .CellType = 1	: .Text = "계"	: .TypeHAlign = 2
			ggoSpread.SSSetProtected -1, .MaxRows, .MaxRows
		End If
	End With
End Sub 

'============================================  조회조건 함수  ====================================
Sub GetFISC_DATE()	' 법인의 조회조건에 만족하는 당기시작,종료일을 가져온다.
	Dim sFiscYear, sRepType, sCoCd, ret, datFISC_START_DT, datFISC_END_DT
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	call CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	' 법인 기간은 필수입력 
	lgFISC_START_DT = CDate(lgF0)
	lgFISC_END_DT = CDate(lgF1)
	
End Sub

'============================================  폼 함수  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                                    <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100110100000111")										<%'버튼 툴바 제어 %>

	' 변경한곳 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

	Call InitData()

    Call FncQuery
    
End Sub


'============================================  이벤트 함수  ====================================
Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub

Sub cboREP_TYPE_onChange()	' 신고기준을 바꾸면..
	Call GetFISC_DATE
End Sub


Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	 ' 이 Template 화면에서는 없는 로직임, 콤보(Name)가 변경되면 콤보(Code, Hidden)도 변경시켜주는 로직 
	With frm1.vspdData
		.Row = Row

		Select Case Col
			Case  C_W1
				.Col = Col
				intIndex = .Value
				.Col = C_W1_CD
				.Value = intIndex	
			Case  C_W1_CD
				.Col = Col
				intIndex = .Value
				.Col = C_W1
				.Value = intIndex		
		End Select
		
		' 사업연도 종료일을 출력한다.
		.Col = C_W2
		.Text = lgFISC_END_DT
	End With
End Sub

'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

    If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
    End If
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

	With frm1.vspdData

		Select Case Col
			Case C_W4, C_W5
					
				Call FncSumSheet(frm1.vspdData, Col, 1, .MaxRows - 1, true, .MaxRows, Col, "V")	' 합계 
				
				ggoSpread.UpdateRow .MaxRows
		End Select 
			
	End With
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
	Call ggoSpread.ReOrderingSpreadData()
End Sub


'============================================  툴바지원 함수  ====================================

Function FncNew() 
    Dim IntRetCD 

    FncNew = False

  '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")           '⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")
    Call ggoOper.LockField(Document, "N")
    Call InitVariables
    Call InitData

    Call SetToolbar("1100110000000111")

	Call ClickTab1()
	frm1.txtCO_CD.focus

    FncNew = True

End Function

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
	
     
    CALL DBQuery()
    
End Function

Function FncSave() 
    Dim blnChange, dblSum
    
    FncSave = False                                                         
    blnChange = False
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		blnChange = True
    End If

	If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
	      Exit Function
	End If    

	If blnChange = False Then
        Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
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

	With frm1
		If .vspdData.ActiveRow > 0 Then
			.vspdData.focus
			.vspdData.ReDraw = False
		
			ggoSpread.CopyRow
			SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow

			.vspdData.Col = C_DOC_AMT
			.vspdData.Text = ""
    
			.vspdData.Col = C_COMPANY_NM
			.vspdData.Text = ""
			
			.vspdData.Col = C_STOCK_RATE
			.vspdData.Text = ""
			
			.vspdData.Col = C_ACQUIRE_AMT
			.vspdData.Text = ""
			
			.vspdData.ReDraw = True
		End If
	End With
	
    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 
  dim iRow
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData	
	    iRow = .Row 
	    If .ActiveRow = .MaxRows Or .MaxRows = 0 Then Exit Function
	
		
		ggoSpread.EditUndo                                                  '☜: Protect system from crashing
		
        if .Maxrows = 1  then 
           	ggoSpread.EditUndo  
        else 
			Call FncSumSheet(frm1.vspdData, C_W4, 1, .MaxRows - 1, true, .MaxRows, C_W4, "V")	' 합계 
			Call FncSumSheet(frm1.vspdData, C_W5, 1, .MaxRows - 1, true, .MaxRows, C_W5, "V")	' 합계 
					
		    ggoSpread.UpdateRow .MaxRows   
		end if    
		 .Row  = iRow
	
	End With 
End Function

Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim iRow, iSeqNo

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim uCountID

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
 
	With frm1.vspdData	' 포커스된 그리드 
			
		ggoSpread.Source = frm1.vspdData
			
		iRow = .ActiveRow
		.ReDraw = False
					
		If .MaxRows = 0 Then	' 첫 InsertRow는 1줄+합계줄 

			iRow = 1
			ggoSpread.InsertRow , 2
			Call SetSpreadColor( iRow, iRow+1) 
			.Row = iRow		
			

			.Col = C_SEQ_NO : .Text = iRow	
			
			iRow = 2		: .Row = iRow
			.Col = C_SEQ_NO : .Text = SUM_SEQ_NO	
			.Col = C_W1		: .CellType = 1	: .Text = "계"	: .TypeHAlign = 2
			ggoSpread.SpreadLock C_W1, iRow, .MaxCols-1, iRow

		Else
				
			If iRow = .MaxRows Then	' -- 마지막 합계줄에서 InsertRow를 하면 상위에 추가한다.
				ggoSpread.InsertRow iRow-1 , imRow 
				SetSpreadColor iRow, iRow + imRow - 1

				Call SetDefaultVal(iRow, imRow)
			Else
				ggoSpread.InsertRow ,imRow
				SetSpreadColor iRow+1, iRow + imRow

				Call SetDefaultVal(iRow+1, imRow)
			End If   
		End If
	End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function

' 그리드에 SEQ_NO, TYPE 넣는 로직 
Function SetDefaultVal( iRow, iAddRows)
	
	Dim i, iSeqNo
	
	With frm1.vspdData	' 포커스된 그리드 

	ggoSpread.Source = frm1.vspdData
	
	If iAddRows = 1 Then ' 1줄만 넣는경우 
		.Row = iRow
		MaxSpreadVal frm1.vspdData, C_SEQ_NO, iRow
	Else
		iSeqNo = MaxSpreadVal(frm1.vspdData, C_SEQ_NO, iRow)	' 현재의 최대SeqNo를 구한다 
		
		For i = iRow to iRow + iAddRows -1
			.Row = i
			.Col = C_SEQ_NO : .Value = iSeqNo : iSeqNo = iSeqNo + 1
		Next
	End If
	End With
End Function

Function FncDeleteRow() 
    Dim lDelRows

    With frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	If .ActiveRow = .MaxRows Or .MaxRows = 0 Then Exit Function
    	
    	lDelRows = ggoSpread.DeleteRow
    	
		Call FncSumSheet(frm1.vspdData, C_W4, 1, .MaxRows - 1, true, .MaxRows, C_W4, "V")	' 합계 
		Call FncSumSheet(frm1.vspdData, C_W5, 1, .MaxRows - 1, true, .MaxRows, C_W5, "V")	' 합계 

    	If .MaxRows = 2 Then 
    		.SetActiveCell 1, .MaxRows
    		lDelRows = ggoSpread.DeleteRow
    	Else
    		ggoSpread.UpdateRow .MaxRows
    	End If
		
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


'============================================  DB 억세스 함수  ====================================
Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
	
	Call LayerShowHide(1)
	
	Dim strVal
    
    With Frm1
    
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '☜: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '☜: Query Key   
        
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '☜:  Run biz logic

    DbQuery = True  
  
End Function

Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    If frm1.vspdData.MaxRows > 0 Then
    
		lgIntFlgMode = parent.OPMD_UMODE
		
		Call SetSpreadLock
		
		Call SetSpreadTotalLine
		
		Call SetToolbar("1100111100000111")										<%'버튼 툴바 제어 %>
	End If
	
	frm1.vspdData.focus			
End Function


'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim pP21011
    Dim lRow, lCol   
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow, lEndRow     
    Dim lRestGrpCnt 
    Dim strVal, strDel, lMaxRows, lMaxCols, sTmp
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if

    strVal = ""
    strDel = ""
    lGrpCnt = 1
    
	With frm1.vspdData
		' ----- 1번째 그리드 
		ggoSpread.Source = frm1.vspdData
		lMaxRows = .MaxRows : lMaxCols = .MaxCols
				
		For lRow = 1 To lMaxRows
		    
		   .Row = lRow : .Col = 0 : sTmp = ""
		   
		  ' 모든 그리드 데이타 보냄     
		  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
				For lCol = C_SEQ_NO To lMaxCols
					Select Case lCol
						Case C_W3
							.Col = lCol : sTmp = sTmp & Trim(.Value) &  Parent.gColSep
						Case Else
							.Col = lCol : sTmp = sTmp & Trim(.Text) &  Parent.gColSep
					End Select
				Next
				sTmp = sTmp & Trim(.Text) &  Parent.gRowSep
		  End If  

		.Row = lRow : .Col = 0  
		   ' I/U/D 플래그 처리 
		   Select Case .Text
		       Case  ggoSpread.InsertFlag                                      '☜: Insert
		                                          strVal = strVal & "C"  &  Parent.gColSep & sTmp
		            lGrpCnt = lGrpCnt + 1
		                    
		       Case  ggoSpread.UpdateFlag                                      '☜: Update                                                  
		                                           strVal = strVal & "U"  &  Parent.gColSep  & sTmp                                                
		            lGrpCnt = lGrpCnt + 1                                                 
		       Case  ggoSpread.DeleteFlag                                      '☜: Delete
		                                          strDel = strDel & "D"  &  Parent.gColSep & sTmp
		            lGrpCnt = lGrpCnt + 1  
		  End Select
		   
		Next
	End With

	frm1.txtMode.value        =  Parent.UID_M0002
    frm1.txtSpread.value      = strDel & strVal
    strVal = ""
    
	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
	
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
					<TD CLASS="CLSMTABP_BAK">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()" width=200>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB_BAK"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
									<TD CLASS="TD5">사업연도</TD>
									<TD CLASS="TD6"><script language =javascript src='./js/w9113ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
									<TD CLASS="TD5">법인명</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCO_CD" Size=10 tag="14">
										<INPUT TYPE=TEXT NAME="txtCO_NM" Size=20 tag="14">
									</TD>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">신고구분</TD>
									<TD CLASS="TD6"><SELECT NAME="cboREP_TYPE" ALT="신고구분" STYLE="WIDTH: 50%" tag="14X"></SELECT>
									</TD>
									<TD CLASS="TD5"></TD>
									<TD CLASS="TD6"></TD>
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
									<script language =javascript src='./js/w9113ma1_vaSpread1_vspdData.js'></script>
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
		<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE CLASS="TB3" CELLSPACING=0>
	
		
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint('VIEW')" Flag=1>미리보기</BUTTON>&nbsp;
						<BUTTON NAME="bttnPrint"	CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint('PRINT')"   Flag=1>인쇄</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

