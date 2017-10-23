<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 기타 서식 
'*  3. Program ID           : w9111ma1
'*  4. Program Name         : w9111ma1.asp
'*  5. Program Desc         : 제 54호 주식변동상황명세서(을)
'*  6. Modified date(First) : 2005/02/03
'*  7. Modified date(Last)  : 2006/02/03
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
<STYLE>
	.RADIO2 {
		BORDER: 0;
	}
</STYLE>
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

Const BIZ_MNU_ID = "w9111ma1"
Const BIZ_PGM_ID = "w9111mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_REF_PGM_ID = "w9111mb2.asp"
Const EBR_RPT_ID = "w9111oa1"

Const TYPE_1 = 0
Const TYPE_2 = 1

Dim C_W1
Dim C_W2
Dim C_W3
Dim C_W4
Dim C_W10

Dim C_SEQ_NO
Dim C_SEQ_NO_VIEW
Dim C_W5
Dim C_W6
Dim C_W7
Dim C_W8
Dim C_W9

Dim IsOpenPop    
Dim gSelframeFlg , lgCurrGrid      
Dim lgStrPrevKey2
Dim	IsRunEvents, lgvspdData(1)
dim lgfisc_start_dt,	lgfisc_end_dt
'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()
	
	C_W1		= 0	' 콘트롤배열 순서(HTML기준)
	C_W2		= 1
	C_W3		= 2
	C_W4		= 3
	C_W10		= 4
	
	C_SEQ_NO	= 1	' 그리드배열 
	C_SEQ_NO_VIEW = 2
	C_W5		= 3
	C_W6		= 4
	C_W7		= 5
	C_W8		= 6
	C_W9		= 7

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

	IsRunEvents = False
	lgCurrGrid = TYPE_2
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
	
	Set lgvspdData(TYPE_1) = frm1.txtData
	Set lgvspdData(TYPE_2) = frm1.vspdData
	
	Call initSpreadPosVariables()  

	With lgvspdData(TYPE_2)
		
		ggoSpread.Source = lgvspdData(TYPE_2)	
		'patch version
		ggoSpread.Spreadinit "V20041222" & TYPE_2 ,,parent.gForbidDragDropSpread    
			 
		.ReDraw = false
			 
		.MaxCols = C_W9 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols														'☆: 사용자 별 Hidden Column
		.ColHidden = True    

  		'헤더를 3줄로    
		.ColHeaderRows = 2  
						       
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		ggoSpread.SSSetEdit		C_SEQ_NO,		"일련번호", 10,,,10,1
		ggoSpread.SSSetEdit		C_SEQ_NO_VIEW,	"일련번호", 7,2,,10,1

		ggoSpread.SSSetEdit		C_W5,		"(5)성명" , 20,,,50,1
		ggoSpread.SSSetMask		C_W6,		"(6)주민등록번호"	, 20, 2, "999999-9999999" 
		ggoSpread.SSSetDate		C_W7,		"(7)양도일자",			15,		2,		Parent.gDateFormat,	-1
		ggoSpread.SSSetDate		C_W8,		"(8)취득일자",			15,		2,		Parent.gDateFormat,	-1
		ggoSpread.SSSetFloat	C_W9,		"(9)주식수"  , 20,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,"Z" 
		
		ret = .AddCellSpan(C_SEQ_NO			, -1000, 1, 2)	
		ret = .AddCellSpan(C_SEQ_NO_VIEW	, -1000, 1, 2)	
		ret = .AddCellSpan(C_W5		, -1000, 2, 1)	
		ret = .AddCellSpan(C_W7		, -1000, 3, 1)
		
		' 첫번째 헤더 출력 글자 
		.Row = -1000
		.Col = C_SEQ_NO_VIEW	: .Text = "일련" &  vbCrLf & "번호"
		.Col = C_W5		: .Text = "주 식 양 도 자"
		.Col = C_W7	: .Text = "주식.출자지분 양도내용"
		
		.Row = -999
		.Col = C_W5		: .Text = "(5)성명"
		.Col = C_W6		: .Text = "(6)주민등록번호"
		.Col = C_W7		: .Text = "(7)양도일자"
		.Col = C_W8		: .Text = "(8)취득일자"
		.Col = C_W9		: .Text = "(9)주식수" & vbCrLf & "(출자좌수)"
		
		.rowheight(-1000) = 15					
		.rowheight(-999) = 20					
					   
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO, C_SEQ_NO, True)
				
		.ReDraw = true
			
		'Call SetSpreadLock 
	    
	End With

End Sub


'============================================  그리드 함수  ====================================

Sub InitSpreadComboBox()

End Sub

Sub SetSpreadLock()
    With lgvspdData(TYPE_2)

		.ReDraw = False
		ggoSpread.SpreadLock C_SEQ_NO,  1, C_W9, 1
		ggoSpread.SpreadLock C_SEQ_NO_VIEW,  1, C_SEQ_NO_VIEW, .MaxRows
		ggoSpread.SSSetRequired  C_W5, 2, .MaxRows
  		ggoSpread.SSSetRequired  C_W6, 2, .MaxRows
  		ggoSpread.SSSetRequired  C_W7, 2, .MaxRows
  		ggoSpread.SSSetRequired  C_W8, 2, .MaxRows
  		ggoSpread.SSSetRequired  C_W9, 2, .MaxRows
  	
		.ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With lgvspdData(TYPE_2)

		.ReDraw = False

		If pvStartRow > 1 Then
			ggoSpread.SSSetProtected C_SEQ_NO, pvStartRow, pvEndRow
  			ggoSpread.SSSetProtected C_SEQ_NO_VIEW, pvStartRow, pvEndRow
  			ggoSpread.SSSetRequired  C_W5, pvStartRow, pvEndRow
  			ggoSpread.SSSetRequired  C_W6, pvStartRow, pvEndRow
  			ggoSpread.SSSetRequired  C_W7, pvStartRow, pvEndRow
  			ggoSpread.SSSetRequired  C_W8, pvStartRow, pvEndRow
  			ggoSpread.SSSetRequired  C_W9, pvStartRow, pvEndRow
  		End If
  		
		.ReDraw = True
    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
  
End Sub

Sub InitData()
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"
	frm1.txtW4_1.checked = true

	
	
	call CommonQueryRs("fisc_start_dt ,fisc_end_dt "," TB_COMPANY_HISTORY "," CO_CD = "&filterVar(frm1.txtCO_CD.value,"''","S")&" and FISC_YEAR="&filterVar(frm1.txtFISC_YEAR.text,"''","S")&" and REP_TYPE="&filterVar(frm1.cboREP_TYPE.value,"''","S")&" ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	lgfisc_start_dt=Replace(lgF0,chr(11),"")
	lgfisc_end_dt=Replace(lgF1,chr(11),"")


End Sub

'============================== 레퍼런스 함수  ========================================
Function GetRef()	' 그리드1의 금액가져오기 링크 클릭시 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD, arrW1, arrW2, iMaxRows, sTmp, iRow, arrADDR
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg

	' 온로드시 레퍼런스메시지 가져온다.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	sMesg = wgRefDoc & vbCrLf & vbCrLf

	IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '⊙: "Will you destory previous data"

	If IntRetCD = vbNo Then
		 Exit Function
	End If

	lgvspdData(TYPE_2).MaxRows = 0
    ggoSpread.ClearSpreadData
	
	Dim strVal
    
    With Frm1
    
		strVal = BIZ_REF_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal			& "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '☜: Query Key        
        strVal = strVal			& "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '☜: Query Key   
        
		Call RunMyBizASP(MyBizASP, strVal)   
    End With  

End Function



' 해당 그리드에서 데이타가져오기 
Function GetGrid(Byval pCol, Byval pRow)
	With lgvspdData(TYPE_2)
		.Col = pCol	: .Row = pRow : GetGrid = .Value
	End With
End Function

' 해당 그리드에서 데이타가져오기 
Function GetGridText(Byval pCol, Byval pRow)
	With lgvspdData(TYPE_2)
		.Col = pCol	: .Row = pRow : GetGridText = .Text
	End With
End Function

' 해당 그리드에서 데이타가져오기 
Function PutGrid(Byval pCol, Byval pRow, Byval pVal)
	With lgvspdData(TYPE_2)
		.Col = pCol	: .Row = pRow : .Value = pVal
	End With
End Function

' 해당 그리드에서 데이타가져오기 
Function PutGridText(Byval pCol, Byval pRow, Byval pVal)
	With lgvspdData(TYPE_2)
		.Col = pCol	: .Row = pRow : .Text = pVal
	End With
End Function
'============================================  그리드 팝업  ====================================


' 헤더 재계산 
Sub SetHeadReCalc()	
	Dim dblSum, dblData(40)
	
	If IsRunEvents Then Exit Sub	' 아래 .vlaue = 에서 이벤트가 발생해 재귀함수로 가는걸 막는다.
	
	IsRunEvents = True
	
	With frm1

		
	End With

	lgBlnFlgChgValue= True ' 변경여부 
	IsRunEvents = False	' 이벤트 발생금지를 해제함 
End Sub

'============================================  조회조건 함수  ====================================

'============================================  폼 함수  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                                    <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitVariables                                                      <%'Initializes local global variables%>
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    
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

Sub txtW10_Click(pIdx)
	With frm1
		.txtW10(0).checked = false
		.txtW10(1).checked = false
		.txtW10(2).checked = false
		.txtW10(pIdx).checked = true
		.txtData(C_W10).value = pIdx + 1
	End With
End Sub

Sub GetFISC_DATE()	' 법인의 조회조건에 만족하는 당기시작,종료일을 가져온다.

End Sub

Sub Make_W4()
	Dim i
	With frm1
		For i = 1 To 7
			If document.all("txtW4_"&i).checked Then
				lgvspdData(TYPE_1)(C_W4).value = i
				Exit Sub
			End If
		Next
	End With
End Sub



sub txtW4_1_Onclick()
    lgBlnFlgChgValue= True 
end sub

sub txtW4_2_Onclick()
    lgBlnFlgChgValue= True 
end sub

sub txtW4_3_Onclick()
    lgBlnFlgChgValue= True 
end sub

sub txtW4_4_Onclick()
    lgBlnFlgChgValue= True 
end sub

sub txtW4_5_Onclick()
    lgBlnFlgChgValue= True 
end sub

sub txtW4_6_Onclick()
    lgBlnFlgChgValue= True 
end sub

sub txtW4_7_Onclick()
    lgBlnFlgChgValue= True 
end sub

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)

End Sub

'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	With lgvspdData(TYPE_2)
		lgBlnFlgChgValue= True ' 변경여부 
		.Row = Row
		.Col = Col

		If .CellType = parent.SS_CELL_TYPE_FLOAT Then
		  If UNICDbl(.text) < UNICDbl(.TypeFloatMin) Then
		     .text = .TypeFloatMin
		  End If
		End If
	
		ggoSpread.Source = lgvspdData(TYPE_2)
		ggoSpread.UpdateRow Row

		Dim datW7, datW8

		.Row = Row
		.Col = Col
				
		Select Case Col
			Case C_W6
				If Len(Trim(.Text)) < 14 Then
					Call DisplayMsgBox("WC0036", parent.VB_INFORMATION, .Text, GetGrid(Col, -999))           '⊙: "Will you destory previous data"
					Call PutGridText(Col, Row, "")
				End If
				
			Case C_W7
				
				datW7 = GetGridText(C_W7, Row)
				datW8 = GetGridText(C_W8, Row)
				
				If datW7 = "" Then Exit Sub

				If Len(datW7) < 10 Then
					Call DisplayMsgBox("WC0036", parent.VB_INFORMATION, datW7, GetGrid(Col, -999))           '⊙: "Will you destory previous data"
					Call PutGridText(Col, Row, "")
				End If
				
				If datW8 = "" Then Exit Sub

				If datW7 < datW8 Then
					Call DisplayMsgBox("WC0010", parent.VB_INFORMATION, GetGrid(C_W8, -999), GetGrid(C_W7, -999))           '⊙: "Will you destory previous data"
					Call PutGridText(C_W7, Row, "")
					Exit Sub
				End If
				
			Case C_W8
				
				datW7 = GetGridText(C_W7, Row)
				datW8 = GetGridText(C_W8, Row)
				
				If datW8 = "" Then Exit Sub

				If Len(datW8) < 10 Then
					Call DisplayMsgBox("WC0036", parent.VB_INFORMATION, datW8, GetGrid(Col, -999))           '⊙: "Will you destory previous data"
					Call PutGridText(Col, Row, "")
				End If
				
				If datW7 = "" Then Exit Sub
					
				If datW7 < datW8 Then
					Call DisplayMsgBox("WC0010", parent.VB_INFORMATION, GetGrid(C_W8, -999), GetGrid(C_W7, -999))           '⊙: "Will you destory previous data"
					Call PutGridText(C_W8, Row, "")
					Exit Sub
				End If
			
			Case C_W9
				Call FncSumSheet(lgvspdData(TYPE_2), C_W9, 2, .MaxRows , true, 1, C_W9, "V")	' 합계 
				.Row = 1	: .Col = 1
				ggoSpread.UpdateRow .Row
		End Select 
			
	End With
End Sub

Sub ReCalcW36()
	Dim iMaxRows, iRow, dblSum, dblW35, dblW36
	
	dblSum = GetGrid(C_W35, 1)
	
	With lgvspdData(TYPE_2)
		ggoSpread.Source = lgvspdData(TYPE_2)
		
		iMaxRows = .MaxRows
		
		For iRow = 2 To iMaxRows
			.Row = iRow		: .Col = C_W35	: dblW35 = UNICDbl(.value)
			dblW36 = dblW35 / dblSum
			.Col = C_W36	: .Value = dblW36

			ggoSpread.UpdateRow iRow
		Next
		
		Call PutGrid(C_W36, 1, "1")
		ggoSpread.UpdateRow 1	' 합계쪽 플래그변경 
	End With
End Sub

Sub vspdData_Click(ByVal Col, ByVal Row)
    'Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = lgvspdData(TYPE_2)
   
    If lgvspdData(TYPE_2).MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = lgvspdData(TYPE_2)
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If

	lgvspdData(TYPE_2).Row = Row
End Sub

Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = lgvspdData(TYPE_2)
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If lgvspdData(TYPE_2).MaxRows = 0 Then
        Exit Sub
    End If

End Sub

Sub vspdData_GotFocus()
    ggoSpread.Source = lgvspdData(TYPE_2)

End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    

Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = lgvspdData(TYPE_2)
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
    
    if lgvspdData(TYPE_2).MaxRows < NewTop + VisibleRowCnt(lgvspdData(TYPE_2),NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
      	   Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreTooBar()
			    Exit Sub
			End If  				
    	End If
    End if
End Sub

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With lgvspdData(TYPE_2)
		Select Case Col
			Case C_W37_P
				.Col = Col - 1
				Call OpenW1034(.Value)
			Case C_W16_P
				.Col = Col - 1
				Call OpenW1039(.Value)
		End Select
    End With
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

    Call SetToolbar("1100110100000111")

	frm1.txtCO_CD.focus

    FncNew = True

End Function

Function FncDelete() 
    Dim IntRetCD

    FncDelete = False
    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003",parent.VB_YES_NO, "X", "X")
    IF IntRetCD = vbNO Then
		Exit Function
	End IF

    Call DbDelete

    FncDelete = True
End Function

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               <%'Protect system from crashing%>

<%  '-----------------------
    'Check previous data area
    '----------------------- %>
    If lgBlnFlgChgValue Then
		ggoSpread.Source = lgvspdData(TYPE_2)
		If ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
			If IntRetCD = vbNo Then
		  	Exit Function
			End If
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
    If lgBlnFlgChgValue = False Then
		Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
		Exit Function
	End If
	
	If lgvspdData(TYPE_2).MaxRows > 0 Then
	
		ggoSpread.Source = lgvspdData(TYPE_2)
		If ggoSpread.SSCheckChange = True Then
			blnChange = True
		End If

		If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
		      Exit Function
		End If    
	
		'If blnChange = False Then
		'    Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
		'    Exit Function
		'End If
	End If
	
	Call Make_W4
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          
    
End Function

Function FncCopy() 
  
	
End Function

Function FncCancel() 
	With lgvspdData(TYPE_2)
		ggoSpread.Source = lgvspdData(TYPE_2)	
		If .ActiveRow = 1 Or .MaxRows = 0 Then Exit Function
		.Row = .ActiveRow
		ggoSpread.EditUndo                                                   '☜: Protect system from crashing
		
		Call FncSumSheet(frm1.vspdData, C_W9, 2, .MaxRows, true, 1, C_W9, "V")	' 합계 
    End With
    ' 삭제후 결과를 다른행에 반영한다.
    'Call ReCalcGrid()
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
 
	With lgvspdData(TYPE_2)	' 포커스된 그리드 
			
		ggoSpread.Source = lgvspdData(TYPE_2)
			
		iRow = .ActiveRow
		.ReDraw = False
					
		If .MaxRows = 0 Then	' 첫 InsertRow는 1줄+합계줄 

			iRow = 1
			ggoSpread.InsertRow , 1
			ggoSpread.SpreadLock C_SEQ_NO,  1, C_W9, 1
			.Row = iRow		

			iRow = 1				: .Row = iRow
			.Col = C_SEQ_NO			: .Value = "1"
			.Col = C_SEQ_NO_VIEW	: .Value = "계"
					
			Call SetTotalLine
		End If
				
		If iRow = 1 Then	' -- 합계줄에서 InsertRow를 하면 하위에 추가한다.
			iRow = .MaxRows 
			ggoSpread.InsertRow iRow , imRow 

		Else
			ggoSpread.InsertRow ,imRow
		End If   
			
		SetSpreadColor iRow+1, iRow + imRow
		Call SetDefaultVal( iRow+1, imRow)

	End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
         
End Function

' 그리드에 SEQ_NO, TYPE 넣는 로직 
Function SetDefaultVal(iRow, iAddRows)
	
	Dim i, iSeqNo
	
	With lgvspdData(TYPE_2)	' 포커스된 그리드 

	ggoSpread.Source = lgvspdData(TYPE_2)
	
	If iAddRows = 1 Then ' 1줄만 넣는경우 
		.Row = iRow
		iSeqNo = MaxSpreadVal(lgvspdData(TYPE_2), C_SEQ_NO, iRow)	' 현재의 최대SeqNo를 구한다 
		.Col = C_SEQ_NO_VIEW	: .Value = iSeqNo - 1
	Else
		iSeqNo = MaxSpreadVal(lgvspdData(TYPE_2), C_SEQ_NO, iRow)	' 현재의 최대SeqNo를 구한다 
		
		For i = iRow to iRow + iAddRows -1
			.Row = i	
			.Col = C_SEQ_NO			: .Value = iSeqNo 
			.Col = C_SEQ_NO_VIEW	: .Value = iSeqNo - 1
			iSeqNo = iSeqNo + 1
		Next
	End If
	End With
End Function

Sub SetTotalLine()
	With lgvspdData(TYPE_2)
		.Row = 1
		.Col = C_SEQ_NO_VIEW	: .TypeHAlign = 2	: .value = "계"
			
	End With
End Sub

Function FncDeleteRow() 
    Dim lDelRows

    With lgvspdData(TYPE_2) 
    	.focus
    	ggoSpread.Source = lgvspdData(TYPE_2)
    	If .ActiveRow = 1 Or .MaxRows = 0 Then Exit Function
    	If (.ActiveRow = 1 Or .ActiveRow = 2) And .MaxRows > 2 Then
    		Call  DisplayMsgBox("WC0032", parent.VB_INFORMATION, "X", "X") 
    		Exit Function
    	End If
    	
    	lDelRows = ggoSpread.DeleteRow
    	
    	If .MaxRows = 2 Then 
    		.SetActiveCell 1, 1
    		lDelRows = ggoSpread.DeleteRow
    	Else
    		ggoSpread.UpdateRow 1
    	End If    	
    	
    	' 삭제후 결과를 다른행에 반영한다.
    	'Call ReCalcGrid()
    	Call FncSumSheet(frm1.vspdData, C_W9, 2, .MaxRows, true, 1, C_W9, "V")	' 합계 
    	
    	lgBlnFlgChgValue = True
    End With
End Function

Function ReCalcGrid()
	Dim iRow, iMaxRows, dblW(30), dblW35Sum, dblW50Sum, dblW35, dblW50
	
	With lgvspdData(TYPE_2)
		.ReDraw  = False
		iMaxRows = .MaxRows
		
		Call FncSumSheet(lgvspdData(TYPE_2), C_W9, 1, .MaxRows, true, 1, C_W9, "V")	' 합계 
		
		.ReDraw  = True
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
	
    ggoSpread.Source = lgvspdData(TYPE_2)	
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
    'If lgvspdData(TYPE_2).MaxRows > 0 Then
    
		lgIntFlgMode = parent.OPMD_UMODE
		
		Call SetSpreadLock
		
		Call SetToolbar("1101111100000111")										<%'버튼 툴바 제어 %>
	'End If
	
	lgvspdData(TYPE_2).focus			
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
    lGrpCnt = 0
	   
	' 그리드 부분 
	With lgvspdData(TYPE_2)
		ggoSpread.Source = lgvspdData(TYPE_2)
		lMaxRows = .MaxRows : lMaxCols = .MaxCols
				
		For lRow = 1 To lMaxRows
		    
		   .Row = lRow : .Col = 0 : sTmp = ""
		   
		  ' 모든 그리드 데이타 보냄     
		  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
				For lCol = 1 To lMaxCols
					.Col = lCol
					 if .col=C_W7 then
						
						if lRow >1 and (.text < lgfisc_start_dt or .text > lgfisc_end_dt)  then
							
							call DisplayMsgBox("971012","X", "양도일자","X")
							call LayerShowHide(0)				
							.col=C_W7		
							.Row=lRow		
							.action=0
							
							exit function
							
						end if 
					
					 end if
					 .Col = lCol:  sTmp = sTmp & Trim(.Text) &  Parent.gColSep
					 
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
		                                          strVal = strVal & "U"  &  Parent.gColSep & sTmp                                                 
		            lGrpCnt = lGrpCnt + 1                                                 
		       Case  ggoSpread.DeleteFlag                                      '☜: Delete
		                                          strDel = strDel & "D"  &  Parent.gColSep & sTmp
		            lGrpCnt = lGrpCnt + 1  
		  End Select
		 
 
		Next
	End With
	
	frm1.txtMode.value        =  Parent.UID_M0002
    frm1.txtSpread1.value      = strDel & strVal
    strVal = ""
    strDel = ""
    
    On Error Resume Next
    
	' 헤더(HTML)부분 
	With frm1
		.txtData(C_W10).value = lgvspdData(TYPE_2).MaxRows-1
		For lRow = C_W1 To C_W10
			Select Case lRow
				Case Else
					strVal = strVal & .txtData(lRow).Value & Parent.gColSep
			End Select
		Next 
		.txtSpread0.value   = strVal
		.txtHeadMode.value	= lgIntFlgMode
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
	
    DbSave = True                                                           
End Function


Function DbSaveOk()													        <%' 저장 성공후 실행 로직 %>
	Call InitVariables
	lgvspdData(TYPE_2).MaxRows = 0
    ggoSpread.Source = lgvspdData(TYPE_2)
    ggoSpread.ClearSpreadData
	
    Call MainQuery()
End Function


'========================================================================================
Function DbDelete() 
    Err.Clear
    DbDelete = False

    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '☜: Query Key        
    strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '☜: Query Key            
	
	Call RunMyBizASP(MyBizASP, strVal)

    DbDelete = True
End Function

Function DbDeleteOk()
	Call FncNew()
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
<SCRIPT LANGUAGE=javascript FOR=txtData EVENT=Change>
<!--
    SetHeadReCalc();
//-->
</SCRIPT>
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
					<TD CLASS="CLSMTABP_BAK">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP_BAK"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:GetRef()">주식양도내역 불러오기</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
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
									<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFISC_YEAR CLASS=FPDTYYYY title=FPDATETIME ALT="사업연도" tag="14X1" id=txtFISC_YEAR></OBJECT>');</SCRIPT>
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
						<TABLE <%=LR_SPACE_TYPE_60%> BORDER=1>
                            <TR HEIGHT=15%>
                                <TD WIDTH="100%" VALIGN=TOP >
								<TABLE <%=LR_SPACE_TYPE_20%> border="0" width="100%">
								   <TR>
										<TD>
											<TABLE <%=LR_SPACE_TYPE_20%> border="1" width="100%">
											<TR HEIGHT=10>
											       <TD CLASS="TD51" WIDTH="13%">(1)법인명</TD>
												   <TD CLASS="TD61" WIDTH="20%"><INPUT TYPE=TEXT tag="24" style="width: 100%" id="txtData" name=txtData></TD>
												   <TD CLASS="TD51" WIDTH="13%">(2)사업자등록번호</TD>
												   <TD CLASS="TD61" WIDTH="20%"><INPUT TYPE=TEXT tag="24" style="width: 100%" id="txtData" name=txtData style="text-align: center"></TD>
												   <TD CLASS="TD51" WIDTH="13%">(3)대표자</TD>
												   <TD CLASS="TD61" WIDTH="20%" COLSPAN=2><INPUT TYPE=TEXT tag="24" style="width: 100%" id="txtData" name=txtData></OBJECT></TD>
				
											</TR>
											<TR>
											       <TD CLASS="TD51" ROWSPAN=5>(4)주식구분</TD>
												   <TD CLASS="TD61" COLSPAN=5>소득세법 제94조 제1항 4호 나목(특정시설물 이용권부여)</TD>
												   <TD CLASS="TD61" ALIGN=CENTER><INPUT TYPE=RADIO CLASS="RADIO2" ID=txtW4_1 NAME=txtW4 tag="25" CHECKED><LABEL FOR=txtW4_1>1</LABEL></TD>
										   </TD>
											<TR>
												   <TD CLASS="TD61" COLSPAN=5>소득세법시행령 제158조 제1항 1호(부동산 등 50%이상 보유.양도)</TD>
												   <TD CLASS="TD61" ALIGN=CENTER><INPUT TYPE=RADIO CLASS="RADIO2" ID=txtW4_2 NAME=txtW4 tag="25"><LABEL FOR=txtW4_2>2</LABEL></TD>
										   </TD>
											<TR>
												   <TD CLASS="TD61" COLSPAN=5>소득세법시행령 제158조 제1항 5호(골프장 등 영위, 부동산등 80%이상)</TD>
												   <TD CLASS="TD61" ALIGN=CENTER><INPUT TYPE=RADIO CLASS="RADIO2" ID=txtW4_3 NAME=txtW4 tag="25"><LABEL FOR=txtW4_3>3</LABEL></TD>
										   </TD>
											<TR>
												   <TD CLASS="TD61" COLSPAN=4>소득세법 제94조 제1항 3호 가목 또는 나목(상장.등록법인)</TD>
												   <TD CLASS="TD61" ALIGN=CENTER><INPUT TYPE=RADIO CLASS="RADIO2" ID=txtW4_4 NAME=txtW4 tag="25"><LABEL FOR=txtW4_4>4.중소</LABEL></TD>
												   <TD CLASS="TD61" ALIGN=CENTER><INPUT TYPE=RADIO CLASS="RADIO2" ID=txtW4_5 NAME=txtW4 tag="25"><LABEL FOR=txtW4_5>5.일반</LABEL></TD>
										   </TD>
											<TR>
												   <TD CLASS="TD61" COLSPAN=4>소득세법 제94조 제1항 3호 다목(비상장법인)</TD>
												   <TD CLASS="TD61" ALIGN=CENTER><INPUT TYPE=RADIO CLASS="RADIO2" ID=txtW4_6 NAME=txtW4 tag="25"><LABEL FOR=txtW4_6>6.중소</LABEL></TD>
												   <TD CLASS="TD61" ALIGN=CENTER><INPUT TYPE=RADIO CLASS="RADIO2" ID=txtW4_7 NAME=txtW4 tag="25"><LABEL FOR=txtW4_7>7.일반</LABEL></TD>
										   </TR>
											</TABLE>
										<TD>
									</TR>
								</TD>
							</TR>
							<TR HEIGHT=85%>
								<TD WIDTH="100%" VALIGN=TOP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread0> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
<INPUT TYPE=HIDDEN NAME="txtData" tag="24"><% ' 라디오버튼 값 %>
<INPUT TYPE=HIDDEN NAME="txtData" tag="24"><% ' 그리드갯수 %>
<TEXTAREA CLASS="hidden" NAME="txtSpread0" tag="24"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread1" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHeadMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

