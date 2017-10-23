
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 각 과목별 조정 
'*  3. Program ID           : W3105MA1
'*  4. Program Name         : W3105MA1.asp
'*  5. Program Desc         : 제34호 대손충당금 및 대손금 조정명세서 
'*  6. Modified date(First) : 2005/01/07
'*  7. Modified date(Last)  : 2006/01/23
'*  8. Modifier (First)     : LSHSAT
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

Const BIZ_MNU_ID		= "W3105MA1"
Const BIZ_PGM_ID		= "W3105MB1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_REF_PGM_ID	= "W3105MB2.asp"
Const EBR_RPT_ID		= "W3105OA1"

Dim C_SEQ_NO1
Dim C_W16
Dim C_W16_BT
Dim C_W16_NM
Dim C_W17
Dim C_W18
Dim C_W19
Dim C_W20
Dim C_W21
Dim C_DESC1

Dim C_SEQ_NO2
Dim C_W22
Dim C_W23
Dim C_W23_BT
Dim C_W23_NM
Dim C_W24
Dim C_W25
Dim C_W26
Dim C_W27
Dim C_W28
Dim C_W29
Dim C_W30
Dim C_W31
Dim C_W32
Dim C_DESC2

Dim IsOpenPop          
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 
Dim gCurrGrid

Dim IsRunEvents	' ㅠㅠ 무한이벤트반복을 막음 

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()
	gCurrGrid	=1

	C_SEQ_NO1	= 1	' -- 1번 그리드 
    C_W16		= 2
    C_W16_BT	= 3
    C_W16_NM	= 4
    C_W17		= 5
    C_W18		= 6
    C_W19		= 7
    C_W20		= 8
    C_W21		= 9
    C_DESC1		= 10	

 	C_SEQ_NO2	= 1  ' -- 2번 그리드 
    C_W22		= 2 
    C_W23		= 3
    C_W23_BT	= 4
    C_W23_NM	= 5
    C_W24		= 6
    C_W25		= 7
    C_W26		= 8
    C_W27		= 9
    C_W28		= 10
    C_W29		= 11
    C_W30		= 12
    C_W31		= 13
    C_W32		= 14
    C_DESC2		= 15
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
    lgRefMode = False
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

Sub InitMinor()	'대손표준율 가져오기 
	Dim iArrCd, iArrRf
	
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	
	call CommonQueryRs("seq_no, reference"," b_configuration "," MAJOR_CD = 'W2004' AND minor_cd = (select comp_type2 from TB_COMPANY_HISTORY where co_cd = '" &sCoCd &"' and fisc_year = '" &sFiscYear &"'and Rep_Type = '" & sRepType &"') ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	if lgF0 = "" then Exit sub
	iArrCd = Split(lgF0, Chr(11))
	iArrRf = Split(lgF1, Chr(11))
	If iArrCd(0) = "1" Then
    	Frm1.txtW2_1_1.value = iArrRf(0)
    	Frm1.txtW2_1_2.value = iArrRf(1)
    Else
    	Frm1.txtW2_1_1.Value = iArrRf(1)
    	Frm1.txtW2_1_2.Value = iArrRf(0)
   End If
End Sub

Sub InitSpreadSheet()
	Dim ret
	
    Call initSpreadPosVariables()  

	' 1번 그리드 
	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData	
	   'patch version
	    ggoSpread.Spreadinit "V20041222",,parent.gForbidDragDropSpread    
	    
		.ReDraw = false
	
	    .MaxCols = C_DESC1 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols									'☆: 사용자 별 Hidden Column
		.ColHidden = True    
		       
	    .MaxRows = 0
	    ggoSpread.ClearSpreadData
	    
	    'Call AppendNumberPlace("6","3","2")
	
	    ggoSpread.SSSetEdit		C_SEQ_NO1,	"순번"		, 10,,,6,1	' 히든컬럼 
		ggoSpread.SSSetEdit		C_W16,		"(16)계정과목"	, 10,,,50,1	
	    ggoSpread.SSSetButton 	C_W16_BT     '4
		ggoSpread.SSSetEdit		C_W16_NM,	"(16)계정과목"	, 15,,,50,1	
	    ggoSpread.SSSetFloat	C_W17,		"(17)채권잔액의" & vbCrLf & "장부가액", 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetFloat	C_W18,		"(18)기말현재" & vbCrLf & "대손금부인누계", 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec , ,                        ,        ,         ,"" 
	    ggoSpread.SSSetFloat	C_W19,		"(19)합계" & vbCrLf & "{(17)+(18)}", 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetFloat	C_W20,		"(20)충당금 " & vbCrLf & "설정제외 채권", 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetFloat	C_W21,		"(21)채권잔액 " & vbCrLf & "{(19)-(20)}", 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetEdit		C_DESC1,	"비 고"	, 10,,,50,1	

	    ret = .AddCellSpan(C_W16, -1000, 3, 1)
	
		Call ggoSpread.MakePairsColumn(C_W16,C_W16_BT)
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO1,C_SEQ_NO1,True)
					
		.rowheight(-1000) = 20	' 높이 재지정 
		.ReDraw = true
		
	    'Call SetSpreadLock 
    
    End With

 	' -----  2번 그리드 
	With frm1.vspdData2
	
		ggoSpread.Source = frm1.vspdData2	
	   'patch version
	    ggoSpread.Spreadinit "V20041222_2",,parent.gForbidDragDropSpread    
	    
		.ReDraw = false
	    
	    .MaxCols = C_DESC2 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols									'☆: 사용자 별 Hidden Column
		.ColHidden = True    
		       
	    .MaxRows = 0
	    ggoSpread.ClearSpreadData
	
		'헤더를 2줄로    
	    .ColHeaderRows = 2
	    'Call AppendNumberPlace("6","3","2")
	
	    ggoSpread.SSSetEdit		C_SEQ_NO2,	"순번", 10,,,6,1	' 히든컬럼 
	    ggoSpread.SSSetDate		C_W22,		"(22)일자"      , 10, 2, parent.gDateFormat '6    
		ggoSpread.SSSetEdit		C_W23,		"(23)계정과목", 10,,,50,1
	    ggoSpread.SSSetButton 	C_W23_BT     '4
		ggoSpread.SSSetEdit		C_W23_NM,	"(23)계정과목", 10,,,50,1
		ggoSpread.SSSetEdit		C_W24,		"(24)채권내역", 10,,,50,1
		ggoSpread.SSSetEdit		C_W25,		"(25)대손사유", 10,,,50,1
	    ggoSpread.SSSetFloat	C_W26,		"(26)금액",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0",""
	    ggoSpread.SSSetFloat	C_W27,		"(27)계" ,		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0","" 
		ggoSpread.SSSetFloat	C_W28,		"(28)시인액",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetFloat	C_W29,		"(29)부인액",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0",""
	    ggoSpread.SSSetFloat	C_W30,		"(30)계",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0","" 
	    ggoSpread.SSSetFloat	C_W31,		"(31)시인액" ,15,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0",""
	    ggoSpread.SSSetFloat	C_W32,		"(32)부인액" , 15,Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0",""
		ggoSpread.SSSetEdit		C_DESC2,	"비 고", 20,,,50,1
	
	 
	    ret = .AddCellSpan(0, -1000, 1, 2)
	    ret = .AddCellSpan(1, -1000, 1, 2)
	    ret = .AddCellSpan(2, -1000, 1, 2)
	    ret = .AddCellSpan(3, -1000, 3, 2)
	    ret = .AddCellSpan(6, -1000, 1, 2)
	    ret = .AddCellSpan(7, -1000, 1, 2)
	    ret = .AddCellSpan(8, -1000, 1, 2)
	    ret = .AddCellSpan(9, -1000, 3, 1)
	    ret = .AddCellSpan(12, -1000, 3, 1)
	    ret = .AddCellSpan(15, -1000, 1, 2) 
	    
	    ' 첫번째 헤더 출력 글자 
		.Row = -1000
		.Col = 9
		.Text = "대손충당금상계액"
	
		.Col = 12
		.Text = "당기손금계상액"
	
		' 두번째 헤더 출력 글자 
		.Row = -999
		.Col = 9
		.Text = "(27)계"
		.Col = 10
		.Text = "(28)시인액"
		.Col = 11
		.Text = "(29)부인액"
		.Col = 12
		.Text = "(30)계"
		.Col = 13
		.Text = "(31)시인액"
		.Col = 14
		.Text = "(32)부인액"
		.rowheight(-999) = 20	' 높이 재지정 
		
		Call ggoSpread.MakePairsColumn(C_W23,C_W23_BT)
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO2,C_SEQ_NO2,True)
					
		
		.ReDraw = true
		
	    Call SetSpreadLock 
    
    End With
       
End Sub


'============================================  그리드 함수  ====================================
Function OpenAccount(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strWhere

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "계정 팝업"					<%' 팝업 명칭 %>
	arrParam(1) = "TB_ACCT_MATCH"					<%' TABLE 명칭 %>
	

	If iWhere = 1 then
		frm1.vspdData.Col = C_W16
	    arrParam(2) = frm1.vspdData.Text		<%' Code Condition%>
	ElseIf iWhere = 2 then
		frm1.vspdData.Col = C_W23
	    arrParam(2) = frm1.vspdData2.Text		<%' Code Condition%>
	End If
	arrParam(3) = ""							<%' Name Cindition%>

	strWhere = " MATCH_CD = '07'"
	strWhere = strWhere & " AND CO_CD = '" & frm1.txtCO_CD.value & "' "
	strWhere = strWhere & " AND FISC_YEAR = '" & frm1.txtFISC_YEAR.text & "' "
	strWhere = strWhere & " AND REP_TYPE = '" & frm1.cboREP_TYPE.value & "' "

	arrParam(4) = strWhere							<%' Where Condition%>
	arrParam(5) = "계정"						<%' 조건필드의 라벨 명칭 %>
	
    arrField(0) = "ED5" & Chr(11) & "ACCT_GP_CD" & Chr(11)					<%' Field명(0)%>
    arrField(1) = "ED10" & Chr(11) & "dbo.ufn_GetCodeName('W1085', ACCT_GP_CD)" & Chr(11)					<%' Field명(1)%>
    arrField(2) = "ED7" & Chr(11) & "ACCT_CD" & Chr(11)					<%' Field명(2)%>
    arrField(3) = "ED15" & Chr(11) & "ACCT_NM" & Chr(11)					<%' Field명(3)%>
    
    arrHeader(0) = "대표계정코드"					<%' Header명(0)%>
    arrHeader(1) = "대표계정명"						<%' Header명(1)%>
    arrHeader(2) = "계정코드"					<%' Header명(2)%>
    arrHeader(3) = "계정명"						<%' Header명(3)%>
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetAccount(arrRet,iWhere)
	End If	
	
End Function

Function SetAccount(byval arrRet,Byval iWhere)
    With frm1
		If iWhere = 1 Then 'Spread1(Condition)
			.vspdData.Col = C_W16
			.vspdData.Text = arrRet(2)
			.vspdData.Col = C_W16_NM
			.vspdData.Text = arrRet(3)
		    ggoSpread.Source = frm1.vspdData
		    ggoSpread.UpdateRow frm1.vspdData.ActiveRow
			lgBlnFlgChgValue = True
		ElseIf iWhere = 2 Then 'Spread2(Condition)
			.vspdData2.Col = C_W23
			.vspdData2.Text = arrRet(2)
			.vspdData2.Col = C_W23_NM
			.vspdData2.Text = arrRet(3)
		    ggoSpread.Source = frm1.vspdData
		    ggoSpread.UpdateRow frm1.vspdData.ActiveRow
			lgBlnFlgChgValue = True
		End If
	End With
End Function


Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False
    .vspdData2.ReDraw = False
        
    ggoSpread.Source = frm1.vspdData2	

	ggoSpread.SSSetProtected C_W27, -1, -1
	ggoSpread.SSSetProtected C_W30, -1, -1
'    ggoSpread.SpreadLock C_W19, -1, -1
'	ggoSpread.SSSetRequired C_W28, -1, -1
	
	
    ggoSpread.Source = frm1.vspdData
        
	ggoSpread.SSSetProtected C_W19, -1, -1
	ggoSpread.SSSetProtected C_W21, -1, -1
'    ggoSpread.SpreadLock C_W17, -1, C_W21, .vspdData.MaxRows
'	ggoSpread.SSSetRequired C_DESC1, -1, -1

    .vspdData.ReDraw = True
    .vspdData2.ReDraw = True

    End With
End Sub

' InsertRow/Copy 할때 호출됨 
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	Dim iRow

	If gCurrGrid = 1 Then
	    With frm1.vspdData
			.ReDraw = False
	 
			ggoSpread.Source = frm1.vspdData
			
			For iRow = pvStartRow To pvEndRow
				.Row = iRow	
				If iRow <>  .MaxRows Then
			  		ggoSpread.SSSetRequired C_W16, iRow, iRow
			  		ggoSpread.SSSetProtected C_W16_NM, iRow, iRow
			 		ggoSpread.SSSetRequired C_W17, iRow, iRow
'			 		ggoSpread.SSSetRequired C_W18, iRow, iRow 		
					ggoSpread.SSSetProtected C_W19 , iRow, iRow
					ggoSpread.SSSetProtected C_W21 , iRow, iRow
			    End If
			    .Col = C_SEQ_NO1
			    If .Text = "999999" and .MaxRows = iRow Then
					ggoSpread.SpreadLock C_W16, iRow, C_W21, iRow
				End If
			Next
		
			.ReDraw = True
		End With
    Else
	    With frm1.vspdData2
			.ReDraw = False
 
			ggoSpread.Source = frm1.vspdData2
	
			For iRow = pvStartRow To pvEndRow
				.Row = iRow	
				If iRow <>  .MaxRows Then
			  		ggoSpread.SSSetRequired C_W22, iRow, iRow
			 		ggoSpread.SSSetRequired C_W23, iRow, iRow
			 		ggoSpread.SSSetProtected C_W23_NM, iRow, iRow
			 		ggoSpread.SSSetRequired C_W26, iRow, iRow
'			 		ggoSpread.SSSetRequired C_W28, iRow, iRow
'			 		ggoSpread.SSSetRequired C_W29, iRow, iRow
'			 		ggoSpread.SSSetRequired C_W31, iRow, iRow
'			 		ggoSpread.SSSetRequired C_W32, iRow, iRow
					ggoSpread.SSSetProtected C_W27, iRow, iRow
					ggoSpread.SSSetProtected C_W30, iRow, iRow
			    End If	 
	
			    .Col = C_SEQ_NO2
			    If .Text = "999999" and .MaxRows = iRow Then
					ggoSpread.SpreadLock C_W22, iRow, C_W32, iRow
				End If
			Next
		
			.ReDraw = True
		End With
    End If
    
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_SEQ_NO	= iCurColumnPos(1)
            C_W7		= iCurColumnPos(2)
            C_W9		= iCurColumnPos(3)
            C_W8		= iCurColumnPos(4)
            C_W8_NM		= iCurColumnPos(5)
            C_W9		= iCurColumnPos(6)
            C_W10		= iCurColumnPos(7)
            C_W11		= iCurColumnPos(8)
            C_W12       = iCurColumnPos(9)
            C_W13		= iCurColumnPos(10)
            C_W15		= iCurColumnPos(11)
            C_W16		= iCurColumnPos(12)
            C_W17		= iCurColumnPos(13)
            C_W18		= iCurColumnPos(14)
            C_W19		= iCurColumnPos(15)
            C_W20		= iCurColumnPos(16)
    End Select    
End Sub

'============================== 레퍼런스 함수  ========================================
Function GetRef()	' 금액불러오기 링크 클릭시 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	Dim BackColor_w,BackColor_g

	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg

	sMesg = wgRefDoc & vbCrLf & vbCrLf
	BackColor_w = frm1.txtW6.BackColor
	
	frm1.txtW4.BackColor =&H009BF0A2&
	frm1.txtW6.BackColor =&H009BF0A2&
	frm1.txtW8.BackColor =&H009BF0A2&
	frm1.txtW10.BackColor =&H009BF0A2&
	frm1.txtW14.BackColor =&H009BF0A2&

    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '⊙: "Will you destory previous data"

	frm1.txtW4.BackColor = BackColor_w
	frm1.txtW6.BackColor = BackColor_w
	frm1.txtW8.BackColor = BackColor_w
	frm1.txtW10.BackColor = BackColor_w
	frm1.txtW14.BackColor = BackColor_w

	If IntRetCD = vbNo Then
		Exit Function
	End If
	
	Dim strVal
    
    With Frm1
    
		strVal = BIZ_REF_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '☜: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '☜: Query Key           
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With  
    	
End Function

Function GetRefOK(ByVal pStrData)
	Dim arrRowVal, arrColVal
	Dim lLngMaxRow, iDx, iRow
	Dim iW22, iW23

	If pStrData <> "" Then
		arrRowVal = Split(pStrData, Parent.gRowSep)                                 '☜: Split Row    data
		lLngMaxRow = UBound(arrRowVal)
		gCurrGrid = 2
		
		With Frm1.vspdData2
			For iDx = 1 To lLngMaxRow
			    arrColVal = Split(arrRowVal(iDx-1), Parent.gColSep)    
	
				If .MaxRows > 0 Then
					For iRow = 1 To .MaxRows - 1
						.Row	= iRow
						.Col 	= C_W22	:	iW22 = .Text
						.Col	= C_W23	:	iW23 = .Text
						.Col	= C_SEQ_NO2
						If iW23	= arrColVal(C_W23) And iW22	= arrColVal(C_W22) Then
							.Row	= iRow
							Exit For
						End If
					Next
					If iRow = .MaxRows Then
						If arrColVal(C_SEQ_NO2) <> "999999" Then Call FncInsertRow(1)
						iRow = iRow
					End If
					.Row	= iRow
				Else
					Call FncInsertRow(1)
					.Row	= 1
				End If
				.Col	= C_W22	:	.Text	= arrColVal(C_W22)
				.Col	= C_W23	:	.Text	= arrColVal(C_W23)
				.Col	= C_W23_NM	:	.Text	= arrColVal(C_W23_NM)
				.Col	= C_W24	:	.Text	= arrColVal(C_W24)
				.Col	= C_W25	:	.Text	= arrColVal(C_W25)
				.Col	= C_W26	:	.Text	= arrColVal(C_W26)
				.Col	= C_W28	:	.Text	= arrColVal(C_W28)
				.Col	= C_W29	:	.Text	= arrColVal(C_W29)
				.Col	= C_W31	:	.Text	= arrColVal(C_W31)
				.Col	= C_W32	:	.Text	= arrColVal(C_W32)
				Call CheckReCalc2(C_W28, .Row)
			Next
		End With
		
	End IF

End Function

Function OpenRefDebt()	'채권금액 조회 

    Dim arrRet
    Dim arrParam(4)
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD
	Dim arrRowVal
    Dim arrColVal, lLngMaxRow
    Dim iDx, iRow
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

'    iCalledAspName = AskPRAspName("W3105RA1")
    
 '   If Trim(iCalledAspName) = "" Then
  '      IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "W3105RA1", "x")
   '     IsOpenPop = False
    '    Exit Function
   ' End If
    
	arrParam(0) = frm1.txtCO_CD.Value
	arrParam(1) = frm1.txtCO_NM.Value		
	arrParam(2) = frm1.txtFISC_YEAR.Text		
	arrParam(3) = frm1.cboREP_TYPE.Value		

    arrRet = window.showModalDialog("W3105RA1.asp", Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	If arrRet(0,0) <> "" Then
		arrRowVal = Split(arrRet(0,0), Parent.gRowSep)                                 '☜: Split Row    data
		lLngMaxRow = UBound(arrRowVal)
		gCurrGrid = 1
		
		For iDx = 1 To lLngMaxRow
		    arrColVal = Split(arrRowVal(iDx-1), Parent.gColSep)    

			If Frm1.vspdData.MaxRows > 0 Then
				For iRow = 1 To Frm1.vspdData.MaxRows - 1
					Frm1.vspdData.Row	= iRow
					Frm1.vspdData.Col	= C_W16
					If Frm1.vspdData.Text	= arrColVal(C_W16) Then
						Frm1.vspdData.Row	= iRow
						Exit For
					End If
				Next
				If iRow = Frm1.vspdData.MaxRows Then
'					Frm1.vspdData.ActiveRow = iRow - 1
					Frm1.vspdData.Row = iRow
					Call FncInsertRow(1)
					iRow = Frm1.vspdData.ActiveRow
				End If
				Frm1.vspdData.Row	= iRow
			Else
				Call FncInsertRow(1)
				Frm1.vspdData.Row	= 1
			End If
			Frm1.vspdData.Col	= C_W16
			Frm1.vspdData.Text	= arrColVal(C_W16)
			Frm1.vspdData.Col	= C_W16_NM
			Frm1.vspdData.Text	= arrColVal(C_W16_NM)
			Frm1.vspdData.Col	= C_W17
			Frm1.vspdData.Text	= arrColVal(C_W17)
			Frm1.vspdData.Col	= C_W18
			Frm1.vspdData.Text	= arrColVal(C_W18)
		    ggoSpread.Source = frm1.vspdData
		    ggoSpread.UpdateRow frm1.vspdData.Row
			Call CheckReCalc(C_W17, Frm1.vspdData.Row)
		Next
		
	End IF
    
    IsOpenPop = False
    
    
End Function

Function OpenRefMenu()

    Dim arrRet
    Dim arrParam(2)
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

'    iCalledAspName = AskPRAspName("W5105RA1")
    
 '   If Trim(iCalledAspName) = "" Then
  '      IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "W5105RA1", "x")
   '     IsOpenPop = False
    '    Exit Function
   ' End If
    
    With frm1
        If .vspdData.ActiveRow > 0 then 
            arrParam(0) = GetSpreadText(.vspdData, 3, .vspdData.ActiveRow, "X", "X")
            arrParam(1) = GetSpreadText(.vspdData, 4, .vspdData.ActiveRow, "X", "X")
        End If            
    End With    

    arrRet = window.showModalDialog("../W5/W5105RA1.asp", Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    
End Function

'============================================  조회조건 함수  ====================================
Sub CheckFISC_DATE()	' 요청법인의 조회조건에 만족하는 당기시작,종료일을 가져온다.
	Dim sFiscYear, sRepType, sCoCd, sFISC_START_DT, sFISC_END_DT, datMonCnt, i, datNow
	
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	call CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If lgF0 <> "" Then 
		sFISC_START_DT = CDate(lgF0)
	Else
		sFISC_START_DT = ""
	End if

    If lgF1 <> "" Then 
		sFISC_END_DT = CDate(lgF1)
	Else
		sFISC_END_DT = ""
	End if
	
End Sub

'============================================  폼 함수  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                                    <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>

    Call SetToolbar("1100111100000111")										<%'버튼 툴바 제어 %>

	' 변경한곳 
    Call InitComboBox
      
    Call AppendNumberRange("0","","")
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
    

	Call InitData
	Call InitMinor
     
    Call FncQuery
End Sub

'============================================  사용자 함수  ====================================
Function Fn_TxtCalc()

	If IsRunEvents Then Exit FUnction	' 아래 .vlaue = 에서 이벤트가 발생해 재귀함수로 가는걸 막는다.
	
	IsRunEvents = True
	

    If unicdbl(Frm1.txtW2_2.value) /100< unicdbl(Frm1.txtW2_1_1.value) Then
    	Frm1.txtW3.value = unicdbl(Frm1.txtw1.value) * unicdbl(Frm1.txtw2_1_1.value)
    Else
    	Frm1.txtW3.value = unicdbl(Frm1.txtw1.value) * unicdbl(Frm1.txtw2_2.value) /100
    End If

    Frm1.txtW5.value = unicdbl(Frm1.txtw6.value) - unicdbl(Frm1.txtw4.value)
    Frm1.txtW12.value = unicdbl(Frm1.txtw5.value)

    Frm1.txtW7.value = unicdbl(Frm1.txtw6.value) - unicdbl(Frm1.txtw3.value)
    If unicdbl(Frm1.txtw7.value) < 0 then Frm1.txtW7.value = 0

    'Frm1.txtW10.value = unicdbl(Frm1.txtw10_BEFORE.value) - unicdbl(Frm1.txtw9.value)

    Frm1.txtW13.value = unicdbl(Frm1.txtw8.value) - unicdbl(Frm1.txtw9.value) - unicdbl(Frm1.txtw10.value) - unicdbl(Frm1.txtw11.value) - unicdbl(Frm1.txtw12.value)

    Frm1.txtW15.value = unicdbl(Frm1.txtw13.value) - unicdbl(Frm1.txtw14.value)
    
	lgBlnFlgChgValue= True ' 변경여부 
	IsRunEvents = False	' 이벤트 발생금지를 해제함 
	
End Function

'============================================  이벤트 함수  ====================================
Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub

' DBQuery 에서 DATA 가져온뒤 호출됨 
Function QueryTotalLine()
	Dim ret, iTmpGrid
	
	iTmpGrid = gCurrGrid
	
	With frm1.vspdData
		If .maxrows>0 then 
	
			ggoSpread.Source = frm1.vspdData
			
			If .MaxRows > 0 Then	' 한줄 추가 
				.Row = .MaxRows
			    ret = .AddCellSpan(C_W16, .MaxRows, 3, 1) 
				.Col = C_W16	: .CellType = 1	: .Text = "계"	: .TypeHAlign = 2
				
			End If
			gCurrGrid = 1
			SetSpreadColor 1, .MaxRows
		End If
	End With
	
	With frm1.vspdData2
		If .maxrows>0 then 
			ggoSpread.Source = frm1.vspdData2
			
			If .MaxRows > 0 Then	' 한줄 추가 
				.Row = .MaxRows
			    ret = .AddCellSpan(C_W22, .MaxRows, 6, 1) 
				.Col = C_W22	: .CellType = 1	: .Text = "계"	: .TypeHAlign = 2
				
			End If
			gCurrGrid = 2
			SetSpreadColor 1, .MaxRows
		End If
	End With
	gCurrGrid = iTmpGrid
End Function


' 레퍼런스에서 넣었으므로 입력으로 변환해주면서 계산도 해 준다.
Function ChangeRowFlg()
	Dim iRow
	
	With frm1.vspdData2
		ggoSpread.Source = frm1.vspdData2
		
		For iRow = 1 To .MaxRows
			.Col = 0 : .Row = iRow : .Value = ggoSpread.InsertFlag
			Call CheckReCalc2(C_W28, iRow)
		Next
	End With
End Function


Sub txtW1_Change( )
    Call Fn_TxtCalc()
End Sub

Sub txtW2_2_Change( )
    Call Fn_TxtCalc()
End Sub

Sub txtw3_Change( )
    Call Fn_TxtCalc()
End Sub

Sub txtw4_Change( )
    Call Fn_TxtCalc()
End Sub

Sub txtw5_Change( )
    Call Fn_TxtCalc()
End Sub

Sub txtw6_Change( )
    Call Fn_TxtCalc()
End Sub

'⑦ 한도초과액				Max ( ( ⑥ 계- ③ 한도액금액) , 0 ) 을 계산하여 입력함.
'				동 금액이  "0" 이 아닌경우 15-1호 서식에 (1)과목명 "대손충당금" (2) 금액은 동 금액 
'				(3)소득처분에는 "유보(증가)"를 입력하고,
'				조정내용은 "대손충당금 한도초과액을 손금불산입하고 유보처분함."을 입력하고 경고하여 줌.
'Sub txtw7_Change( )
'    lgBlnFlgChgValue = True
'    Frm1.txtW13.value = unicdbl(Frm1.txtw8.value) - unicdbl(Frm1.txtw9.value) - unicdbl(Frm1.txtw10.value) - unicdbl(Frm1.txtw11.value) - unicdbl(Frm1.txtw12.value)
'End Sub

Sub txtw8_Change( )
    Call Fn_TxtCalc()
End Sub

Sub txtw9_Change( )
	Dim tmpVal
	
	tmpVal = unicdbl(frm1.txtW10.Text)
	If tmpVal -unicdbl(frm1.txtw9.Text)<0 then 
		frm1.txtW10.Text=unicdbl(0)
	Else
		frm1.txtW10.Text=tmpVal -unicdbl(frm1.txtw9.Text)
	End IF
    Call Fn_TxtCalc()
End Sub

Sub txtw10_Change( )	
    Call Fn_TxtCalc()
'    lgBlnFlgChgValue = True
End Sub

Sub txtw11_Change( )
    Call Fn_TxtCalc()
End Sub

Sub txtw12_Change( )
    Call Fn_TxtCalc()
End Sub

Sub txtw13_Change( )
    Call Fn_TxtCalc()
End Sub

Sub txtw14_Change( )
    Call Fn_TxtCalc()
End Sub

Sub SetW2_2()
	Dim iDblAmt

	With Frm1.vspdData2
		If .MaxRows > 0 Then
		    .Row = .MaxRows
		    .Col = C_W28 :	iDblAmt = unicdbl(.text)
		    .Col = C_W31 :	iDblAmt = iDblAmt + unicdbl(.text)
		    
		    ' ( (28)시인액 합계 + (31)시인액 합계 ) ÷ 직전사업연도의 동 서식의 ①채권잔액 금액 
		    If unicdbl(Frm1.txtW1_BEFORE.Value) > 0 Then
		    	Frm1.txtW2_2.text =round(unicdbl( iDblAmt / unicdbl(Frm1.txtW1_BEFORE.Value))*100,2)
		    End If
		End If
	End With
End Sub

Sub SetW4()

	Dim iDblAmt

	With Frm1.vspdData2
		If .MaxRows > 0 Then
		    .Row = .MaxRows
		    .Col = C_W30 :	iDblAmt = unicdbl(.text)

		    ' ( (4)당기계상액 = 손익계산서상 대손상각비 - (30)당기손금계상액 합계 )
		    iDblAmt = unicdbl(Frm1.txtW4.Value) - iDblAmt

	    else
			iDblAmt=0
	    	iDblAmt = unicdbl(Frm1.txtW4.Value) - iDblAmt
		End If
		
		 If iDblAmt < 0 Then iDblAmt = 0
	    	Frm1.txtW4.Value = iDblAmt
	    	
	End With
End Sub

Sub SetAllTxtRecalc()
	
'	If unicdbl(Frm1.txtW2_2.Value) > 0 Then
		Call SetW2_2()
'	End If
	Call SetW4()
	If unicdbl(Frm1.txtw10_BEFORE.value) - unicdbl(Frm1.txtw9.value)<0 then 
		Frm1.txtW10.value = unicdbl(0)
	Else
		Frm1.txtW10.value = unicdbl(Frm1.txtw10_BEFORE.value) - unicdbl(Frm1.txtw9.value)
	End If
	Call Fn_TxtCalc()
End Sub

'==========================================================================================
Sub InitData()

	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"

	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

	' 온로드시 레퍼런스메시지 가져온다.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	Call CheckFISC_DATE
End Sub

Sub vspdData_Change(ByVal Col , ByVal Row )
	Dim iDblAmt
	Dim iIntRow
	Dim iDblW17, iDblW18, iDblW19, iDblW20, iDblW21
	Dim arrVal
	
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

    If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
    End If
	
	If Col = C_W16 Then
	    frm1.vspdData.Col = 0
'		If  frm1.vspdData.Text = ggoSpread.InsertFlag Then
			frm1.vspdData.Col = C_W16

			If Len(frm1.vspdData.Text) > 0 Then
				frm1.vspdData.Row = Row
				frm1.vspdData.Col = C_W16

'				If CommonQueryRs("ACCT_NM", " TB_WORK_6 (NOLOCK)" , "ACCT_CD = '" & Frm1.vspdData.Text &"' AND ACCT_CD IN (SELECT ACCT_CD FROM TB_ACCT_MATCH (NOLOCK) WHERE MATCH_CD = '07')", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
				If CommonQueryRs("MINOR_NM", " B_MINOR " , "MAJOR_CD = 'W1085' AND MINOR_CD = '" & Frm1.vspdData.Text &"' ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
			    	frm1.vspdData.Col	= C_W16_NM
			    	arrVal				= Split(lgF0, Chr(11))
					frm1.vspdData.Text	= arrVal(0)
				Else
			'		frm1.vspdData.Text	= ""
					frm1.vspdData.Col	= C_W16_NM
					frm1.vspdData.Text	= ""
				End If
			Else
				frm1.vspdData.Col = C_W16_NM
				frm1.vspdData.Text = ""
			End If
'		End If
	ElseIf Col = C_W17 Or Col = C_W18 Or Col = C_W19 Or Col = C_W20 Then

		Call CheckReCalc(Col, Row)
	End If

	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

Sub CheckReCalc(ByVal Col , ByVal Row)
	Dim iDblAmt, iDblW20
	Dim dblSum

	With Frm1.vspdData
		If Row > 0 Then
			.Row = Row
			'(19) 계는 (17) + (18) 을 계산하여 출력한다.
			'(21) 계는 (19) - (20) 을 계산하여 출력한다.
			If Col = C_W17 Or Col = C_W18 Or Col = C_W19 Or Col = C_W20 Then
			    .Col = C_W17 :	iDblAmt = unicdbl(.text)				
			    .Col = C_W18 :	iDblAmt = iDblAmt + unicdbl(.text)	
				.Col = C_W20 :	iDblW20 = unicdbl(.text)
			
			    '(19) < (20) 이면 오류 (메세지 WC0010)
			    If iDblAmt < iDblW20 Then
'			        Call DisplayMsgBox("WC0010", "X", "(19)합계", "(20)충당금설정제외채권")
				    .Col = Col
				    .text = 0
			    End If
			    
			    .Col = C_W17 :	iDblAmt = unicdbl(.text)				
			    .Col = C_W18 :	iDblAmt = iDblAmt + unicdbl(.text)	
			    .Col = C_W19 :	.text = iDblAmt	
			    .Col = C_W20 :	iDblAmt = iDblAmt - unicdbl(.text)				
			    .Col = C_W21 :	.text = iDblAmt			
			End If
		End If
	End With

	With Frm1.vspdData
		If .MaxRows > 0 Then
			dblSum = FncSumSheet(Frm1.vspdData, C_W17, 1, .MaxRows - 1, true, .MaxRows, C_W17, "V")	' 합계 
			dblSum = FncSumSheet(Frm1.vspdData, C_W18, 1, .MaxRows - 1, true, .MaxRows, C_W18, "V")	' 합계 
			dblSum = FncSumSheet(Frm1.vspdData, C_W19, 1, .MaxRows - 1, true, .MaxRows, C_W19, "V")	' 합계 
			dblSum = FncSumSheet(Frm1.vspdData, C_W20, 1, .MaxRows - 1, true, .MaxRows, C_W20, "V")	' 합계 
			dblSum = FncSumSheet(Frm1.vspdData, C_W21, 1, .MaxRows - 1, true, .MaxRows, C_W21, "V")	' 합계 
		    .Row = .MaxRows :	.Col = C_W21 :	Frm1.txtW1.value = .text
		    ggoSpread.Source = frm1.vspdData
		    ggoSpread.UpdateRow .MaxRows
		End If
	End With
End Sub

Sub vspdData2_Change(ByVal Col , ByVal Row )
	Dim iDblAmt
	Dim iIntRow
	Dim iDblW26, iDblW27, iDblW28, iDblW29, iDblW30, iDblW31, iDblW32
	Dim arrVal
	
    Frm1.vspdData2.Row = Row
    Frm1.vspdData2.Col = Col

    If Frm1.vspdData2.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData2.text) < UNICDbl(Frm1.vspdData2.TypeFloatMin) Then
         Frm1.vspdData2.text = Frm1.vspdData2.TypeFloatMin
      End If
    End If
	
    If Col = C_W23 Then
	    frm1.vspdData2.Col = 0
'		If  frm1.vspdData2.Text = ggoSpread.InsertFlag Then
			frm1.vspdData2.Col = C_W23

			If Len(frm1.vspdData2.Text) > 0 Then
				frm1.vspdData2.Row = Row
				frm1.vspdData2.Col = C_W23

'				If CommonQueryRs("ACCT_NM", " TB_WORK_6 (NOLOCK)" , "ACCT_CD = '" & Frm1.vspdData2.Text &"' AND ACCT_CD IN (SELECT ACCT_CD FROM TB_ACCT_MATCH (NOLOCK) WHERE MATCH_CD = '7')", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
				If CommonQueryRs("MINOR_NM", " B_MINOR " , "MAJOR_CD = 'W1085' AND MINOR_CD = '" & Frm1.vspdData2.Text &"' ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
			    	frm1.vspdData2.Col	= C_W23_NM
			    	arrVal				= Split(lgF0, Chr(11))
					frm1.vspdData2.Text	= arrVal(0)
				Else
			'		frm1.vspdData2.Text	= ""
					frm1.vspdData2.Col	= C_W23_NM
					frm1.vspdData2.Text	= ""
				End If
			Else
				frm1.vspdData2.Col = C_W23_NM
				frm1.vspdData2.Text = ""
			End If
'		End If
	ElseIf Col = C_W26 Or Col = C_W27 Or Col = C_W28 Or Col = C_W29 Or Col = C_W30 Or Col = C_W31 Or Col = C_W32 Then
		Frm1.vspdData2.Col = C_W28 :	iDblAmt = unicdbl(Frm1.vspdData2.text)			
		Frm1.vspdData2.Col = C_W29 :	iDblAmt = iDblAmt + unicdbl(Frm1.vspdData2.text)		    
		Frm1.vspdData2.Col = C_W27 :	Frm1.vspdData2.text = iDblAmt
		Frm1.vspdData2.Col = C_W31 :	iDblAmt = unicdbl(Frm1.vspdData2.text)			
		Frm1.vspdData2.Col = C_W32 :	iDblAmt = iDblAmt + unicdbl(Frm1.vspdData2.text)		    
		Frm1.vspdData2.Col = C_W30 :	Frm1.vspdData2.text = iDblAmt
	    Frm1.vspdData2.Col = C_W26
	    iDblW26 = unicdbl(Frm1.vspdData2.text)		
	    Frm1.vspdData2.Col = C_W27
	    iDblW27 = unicdbl(Frm1.vspdData2.text)		
	    Frm1.vspdData2.Col = C_W30
	    iDblW30 = unicdbl(Frm1.vspdData2.text)

	    '(26) < (27) 이면 오류 (메세지 WC0010)
	    If iDblW26 < iDblW27 Then
'	        Call DisplayMsgBox("WC0010", "X", "(27)대손충당금상계액계", "(26)금액")
'		    Frm1.vspdData2.Col = Col
'		    Frm1.vspdData2.text = 0
	    End If

	    '(26) < (30) 이면 오류 (메세지 WC0010)
	    If iDblW26 < iDblW30 Then
'	        Call DisplayMsgBox("WC0010", "X", "(30)당기손금계상액계", "(26)금액")
'		    Frm1.vspdData2.Col = Col
'		    Frm1.vspdData2.text = 0
	    End If

	    '(26) < (27) + (30) 이면 오류 (메세지 WC0010)
	    If iDblW26 < iDblW27 + iDblW30 Then
'	        Call DisplayMsgBox("WC0012", "X", "(27)대손충당금상계액계", "(30)당기손금계상액계")', "(16)금액")
'		    Frm1.vspdData2.Col = Col
'		    Frm1.vspdData2.text = 0
	    End If
		Call CheckReCalc2(Col, Row)
	End If
	
	
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.UpdateRow Row

End Sub

Sub CheckReCalc2(ByVal Col , ByVal Row)
	Dim iDblAmt
	Dim dblSum

	With Frm1.vspdData2
		'(27) 계는 (28) + (29) 을 계산하여 출력한다.
		'(30) 계는 (31) + (32) 을 계산하여 출력한다.
		If Col = C_W27 Or Col = C_W28 Or Col = C_W29 Or Col = C_W30 Or Col = C_W31 Or Col = C_W32 Then
		    .Col = C_W28 :	iDblAmt = unicdbl(.text)			
		    .Col = C_W29 :	iDblAmt = iDblAmt + unicdbl(.text)		    
		    .Col = C_W27 :	.text = iDblAmt
		    .Col = C_W31 :	iDblAmt = unicdbl(.text)			
		    .Col = C_W32 :	iDblAmt = iDblAmt + unicdbl(.text)		    
		    .Col = C_W30 :	.text = iDblAmt		
		End If
	End With

	With Frm1.vspdData2
		If .MaxRows > 0 Then
			dblSum = FncSumSheet(Frm1.vspdData2, C_W26, 1, .MaxRows - 1, true, .MaxRows, C_W26, "V")	' 합계 
			dblSum = FncSumSheet(Frm1.vspdData2, C_W27, 1, .MaxRows - 1, true, .MaxRows, C_W27, "V")	' 합계 
			dblSum = FncSumSheet(Frm1.vspdData2, C_W28, 1, .MaxRows - 1, true, .MaxRows, C_W28, "V")	' 합계 
			dblSum = FncSumSheet(Frm1.vspdData2, C_W29, 1, .MaxRows - 1, true, .MaxRows, C_W29, "V")	' 합계 
			dblSum = FncSumSheet(Frm1.vspdData2, C_W30, 1, .MaxRows - 1, true, .MaxRows, C_W30, "V")	' 합계 
			dblSum = FncSumSheet(Frm1.vspdData2, C_W31, 1, .MaxRows - 1, true, .MaxRows, C_W31, "V")	' 합계 
			dblSum = FncSumSheet(Frm1.vspdData2, C_W32, 1, .MaxRows - 1, true, .MaxRows, C_W32, "V")	' 합계 
		    .Row = .MaxRows :	.Col = C_W27 :	Frm1.txtW11.value = .text

		    ggoSpread.Source = frm1.vspdData2
		    ggoSpread.UpdateRow .MaxRows

			Call SetW2_2()

		End If
	End With
	
	call SetW4
End Sub

Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101011111") 
    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
'    If Row <= 0 Then
'       ggoSpread.Source = frm1.vspdData
       
'       If lgSortKey = 1 Then
'           ggoSpread.SSSort Col               'Sort in ascending
'           lgSortKey = 2
'       Else
'           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
'           lgSortKey = 1
'       End If
       
'       Exit Sub
'    End If

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

	gCurrGrid = 1
	ggoSpread.Source = Frm1.vspdData
End Sub    

Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
'    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
'    Call GetSpreadColumnPos("A")
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


Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then					'☜: 조회중이면 다음 조회 안하도록 체크 
        Exit Sub
	End If
	
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    <%'☜: 재쿼리 체크 %>
      
    	If lgStrPrevKey <> "" And lgStrPrevKey2 <> "" Then                  <%'다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 %>
      		Call DisableToolBar(parent.TBC_QUERY)					'☜ : Query 버튼을 disable 시킴.
			If DBQuery = False Then 
			   Call RestoreToolBar()
			   Exit Sub 
			End If 
    	End If

    End If
End Sub

'============================================  2번 그리드 함수  ====================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData2
   
    If frm1.vspdData2.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
'    If Row <= 0 Then
'       ggoSpread.Source = frm1.vspdData2
       
'       If lgSortKey = 1 Then
'           ggoSpread.SSSort Col               'Sort in ascending
'           lgSortKey = 2
'       Else
'           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
'           lgSortKey = 1
'       End If
       
'       Exit Sub
'    End If

	frm1.vspdData2.Row = Row
End Sub

Sub vspdData2_GotFocus()
    ggoSpread.Source = Frm1.vspdData2
End Sub

Sub vspdData2_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

	gCurrGrid = 2
	ggoSpread.Source = Frm1.vspdData2
End Sub  

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
Dim strTemp
Dim intPos1
   
	With frm1.vspdData 
	
    ggoSpread.Source = frm1.vspdData
   
    If Row > 0 And Col = C_W16_BT Then
        .Col = Col - 1
        .Row = Row        
        Call OpenAccount(1)
        
    End If
    
    End With
      
End Sub

Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
Dim strTemp
Dim intPos1
   
	With frm1.vspdData2 
	
    ggoSpread.Source = frm1.vspdData2
   
    If Row > 0 And Col = C_W23_BT Then
        .Col = Col - 1
        .Row = Row
        
        Call OpenAccount(2)
        
    End If
    
    End With
      
End Sub

'============================================  툴바지원 함수  ====================================

Function FncQuery() 
    Dim IntRetCD, blnChange
    
    FncQuery = False      
	blnChange = False
	IsRunEvents = True 
    
    Err.Clear                                                               <%'Protect system from crashing%>

<%  '-----------------------
    'Check previous data area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
    	blnChange = True
    End If
    
    ggoSpread.Source = frm1.vspdData2
    If ggoSpread.SSCheckChange = True Then
    	blnChange = True
    End If
    If blnChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
    	If IntRetCD = vbNo Then
    		IsRunEvents = False
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
    	IsRunEvents = False
		Exit Function
	End If    
	
     
    CALL DBQuery()
    
End Function

Function FncSave() 
	Dim blnChange
        
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

    ggoSpread.Source = frm1.vspdData2
    If ggoSpread.SSCheckChange = True Then
    	blnChange = True
    End If

    If lgBlnFlgChgValue = False and blnChange = False Then
		Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
		Exit Function
	End If

    ggoSpread.Source = frm1.vspdData
	If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
	      Exit Function
	End If    

    ggoSpread.Source = frm1.vspdData2
	If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
	      Exit Function
	End If    

	If Verification = False Then Exit Function

<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          
    
End Function


' ---------------------- 서식내 검증 -------------------------
Function  Verification()
	Dim iDblW19, iDblW20, iDblW26, iDblW27, iDblW30, iRow
	
	Verification = False

    If unicdbl(Frm1.txtw6.value) < unicdbl(Frm1.txtw4.value) Then
        Call DisplayMsgBox("WC0010", "X", "(6)계", "(4)당기계상액")
	    Frm1.txtw4.value = 0
	    Exit Function
    End If

	With Frm1.vspdData
		For iRow = 1 To .MaxRows
			.Row = iRow
		    .Col = C_W19 :	iDblW19 = unicdbl(.text)
			.Col = C_W20 :	iDblW20 = unicdbl(.text)
			
		    '(19) < (20) 이면 오류 (메세지 WC0010)
		    If iDblW19 < iDblW20 Then
		        Call DisplayMsgBox("WC0010", "X", "(19)합계", "(20)충당금설정제외채권")
			    .Col = C_W20
			    .text = 0
			    Exit Function
		    End If
		Next
	End With
	With Frm1.vspdData2
		For iRow = 1 To .MaxRows
			.Row = iRow
		    .Col = C_W26 :	iDblW26 = unicdbl(.text)
			.Col = C_W27 :	iDblW27 = unicdbl(.text)
			.Col = C_W30 :	iDblW30 = unicdbl(.text)

			'(26) < (27) 이면 오류 (메세지 WC0010)
			If iDblW26 < iDblW27 Then
			    Call DisplayMsgBox("WC0010", "X", "(27)대손충당금상계액계", "(26)금액")
			    .Col = C_W27
			    .text = 0
			    Exit Function
			End If
			
			'(26) < (30) 이면 오류 (메세지 WC0010)
			If iDblW26 < iDblW30 Then
			    Call DisplayMsgBox("WC0010", "X", "(30)당기손금계상액계", "(26)금액")
			    .Col = C_W30
			    .text = 0
			    Exit Function
			End If
			
			'(26) < (27) + (30) 이면 오류 (메세지 WC0010)
			If iDblW26 < iDblW27 + iDblW30 Then
			    Call DisplayMsgBox("WC0012", "X", "(27)대손충당금상계액계", "(30)당기손금계상액계")', "(16)금액")
			    .Col = Col
			    .text = 0
			    Exit Function
			End If
		Next
	End With

	Verification = True	
End Function

'========================================================================================
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
    Call InitMinor
    Call SetToolbar("1100111100000111")										<%'버튼 툴바 제어 %>

	frm1.txtCO_CD.focus

    FncNew = True

End Function


Function FncCopy() 
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

'    If Frm1.vspdData.MaxRows < 1 Or Frm1.vspdData.ActiveRow = Frm1.vspdData.MaxRows Then
       Exit Function
'    End If
 
    ggoSpread.Source = Frm1.vspdData

	With frm1
		If .vspdData.ActiveRow > 0 Then
			.vspdData.focus
			.vspdData.ReDraw = False
		
			ggoSpread.CopyRow
			SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow

			.vspdData.ReDraw = True
		End If
	End With
	
    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 
    Dim lDelRows
    
	If gCurrGrid = 1 Then
		With Frm1.vspdData
			.focus

			ggoSpread.Source = Frm1.vspdData
			If .MaxRows <= 0 Then
				Exit Function
			ElseIf CheckTotRow(Frm1.vspdData, .ActiveRow) = True Then
				MsgBox "합계는 삭제할 수 없습니다.", vbCritical
				Exit Function
			Else
				lDelRows = ggoSpread.EditUndo
				lgBlnFlgChgValue = True
				lDelRows = CheckLastRow(Frm1.vspdData, lDelRows)
				If lDelRows > 0 Then ggoSpread.EditUndo lDelRows
			End If
			
		End With
	    Call CheckReCalc(0,0)
	Else
		With Frm1.vspdData2
			.focus

		    ggoSpread.Source = frm1.vspdData2
			If .MaxRows <= 0 Then
				Exit Function
			ElseIf CheckTotRow(Frm1.vspdData2, .ActiveRow) = True Then
				MsgBox "합계는 삭제할 수 없습니다.", vbCritical
				Exit Function
			Else
				lDelRows = ggoSpread.EditUndo
				lgBlnFlgChgValue = True
				lDelRows = CheckLastRow(Frm1.vspdData2, lDelRows)
				If lDelRows > 0 Then ggoSpread.EditUndo lDelRows
			End If
			
		End With
	    Call CheckReCalc2(0,0)
	End If
End Function

' -- 합계 행인지 체크(Header Grid)
Function CheckTotRow(Byref pObj, Byval pRow) 
	CheckTotRow = False
	pObj.Col = C_SEQ_NO1 : pObj.Row = pRow
	If pObj.Text = "" Then Exit Function
	If pObj.Text = 999999 Then	 ' 합계 행 
		CheckTotRow = True
	End If
End Function

' -- Detail Data가 존재하는지 체크 
Function CheckLastRow(Byref pObj, Byval pRow) 
	Dim iCnt, iRow, iMaxRow
	CheckLastRow = 0
	iCnt = 0
	
	With pObj

		For iRow = 1 To .MaxRows
			.Row = iRow : .Col = 0
			If .Text <> ggoSpread.DeleteFlag Then
				iCnt = iCnt + 1
				iMaxRow = iRow
			End If
		Next
		.Col = C_SEQ_NO1	:	.Row = iMaxRow
		If .Text = 999999 and iCnt = 1 Then
			CheckLastRow = iMaxRow
		End If
	End With
	
End Function

Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim iRow, iSeqNo
    Dim ret

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
   
	With frm1	
	
		If gCurrGrid = 1 Then
		
			.vspdData.focus
			ggoSpread.Source = .vspdData
		
			'.vspdData.ReDraw = False
			iSeqNo = .vspdData.MaxRows+1
		
			if 	.vspdData.MaxRows = 0 then
			
			     ggoSpread.InsertRow  imRow 
			     .vspdData.Col	= C_SEQ_NO1
				.vspdData.Text	= 1
			     ggoSpread.InsertRow  imRow 
			     .row = .vspdData.MaxRows
			    .vspdData.Col	= C_SEQ_NO1
				.vspdData.Text	= "999999"

			    ret = .vspdData.AddCellSpan(C_W16, .vspdData.MaxRows, 3, 1) 
			    
			    ' 첫번째 헤더 출력 글자 
				.vspdData.Col = C_W16
				.vspdData.CellType = 1
				.vspdData.Text = "계"
				.vspdData.TypeHAlign = 2

				 SetSpreadColor 1, .vspdData.MaxRows

			else
					'.vspdData.ReDraw = False	' 이 행이 ActiveRow 값을 사라지게 함, 특별히 긴 로직이 아니라 ReDraw를 허용함. - 최영태 
			     
					iRow = .vspdData.ActiveRow
			
					If iRow = .vspdData.MaxRows Then
					    .vspdData.ActiveRow  = .vspdData.MaxRows -1
						ggoSpread.InsertRow iRow-1 , imRow 
						Call SetSpreadColor(iRow, iRow + imRow - 1)
						Call SetDefaultVal( iRow, imRow)
					Else
			            ggoSpread.InsertRow ,imRow
						Call SetSpreadColor(iRow + 1, iRow +  imRow)   
						Call SetDefaultVal( iRow + 1, imRow)
					End If


	        End if 	

   
		Else
			.vspdData.focus
			ggoSpread.Source = .vspdData2
		
			'.vspdData2.ReDraw = False
			iSeqNo = .vspdData2.MaxRows+1
		
			if 	.vspdData2.MaxRows = 0 then
			
			     ggoSpread.InsertRow  imRow 
			     .vspdData2.Col	= C_SEQ_NO2
				.vspdData2.Text	= 1
			     ggoSpread.InsertRow  imRow 
			     .row = .vspdData2.MaxRows
			    .vspdData2.Col	= C_SEQ_NO2
				.vspdData2.Text	= "999999"
				 
			    ret = .vspdData2.AddCellSpan(C_W22, .vspdData2.MaxRows, 6, 1) 
			    
			    ' 첫번째 헤더 출력 글자 
				.vspdData2.Col = C_W22
				.vspdData2.CellType = 1
				.vspdData2.Text = "계"
				.vspdData2.TypeHAlign = 2

				 SetSpreadColor 1, .vspdData2.MaxRows
				 
			else
					'.vspdData2.ReDraw = False	' 이 행이 ActiveRow 값을 사라지게 함, 특별히 긴 로직이 아니라 ReDraw를 허용함. - 최영태 
			     
					iRow = .vspdData2.ActiveRow
			
					If iRow = .vspdData2.MaxRows Then
					    .vspdData2.ActiveRow  = .vspdData2.MaxRows -1
						ggoSpread.InsertRow iRow-1 , imRow 
						Call SetSpreadColor(iRow, iRow + imRow - 1)
						Call SetDefaultVal( iRow, imRow)
					
					Else
			            ggoSpread.InsertRow ,imRow
						Call SetSpreadColor(iRow + 1, iRow +  imRow)   
						Call SetDefaultVal(iRow + 1, imRow)
					End If
	        End if 	
			'.vspdData2.ReDraw = True
		End If
		
    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function


' 그리드에 SEQ_NO, TYPE 넣는 로직 
Function SetDefaultVal(iRow, iAddRows)
	
	Dim i, iSeqNo
	With frm1	
	
		If gCurrGrid = 1 Then
		
			ggoSpread.Source = .vspdData
		
			If iAddRows = 1 Then ' 1줄만 넣는경우 
				.vspdData.Row = iRow
				.vspdData.Value = MaxSpreadVal(.vspdData, C_SEQ_NO1, iRow)
			Else
				iSeqNo = MaxSpreadVal(.vspdData, C_SEQ_NO1, iRow)	' 현재의 최대SeqNo를 구한다 
				
				For i = iRow to iRow + iAddRows -1
					.vspdData.Row = i
					.vspdData.Col = C_SEQ_NO1 : .vspdData.Value = iSeqNo : iSeqNo = iSeqNo + 1
				Next
			End If
		Else
			ggoSpread.Source = .vspdData2
		
			If iAddRows = 1 Then ' 1줄만 넣는경우 
				.vspdData2.Row = iRow
				.vspdData2.Value = MaxSpreadVal(.vspdData2, C_SEQ_NO2, iRow)
			Else
				iSeqNo = MaxSpreadVal(.vspdData2, C_SEQ_NO2, iRow)	' 현재의 최대SeqNo를 구한다 
				
				For i = iRow to iRow + iAddRows -1
					.vspdData2.Row = i
					.vspdData2.Col = C_SEQ_NO2 : .vspdData2.Value = iSeqNo : iSeqNo = iSeqNo + 1
				Next
			End If
		End If
	End With
End Function


Function FncDeleteRow() 
    Dim lDelRows

	If gCurrGrid = 1	Then
		With frm1.vspdData 
			.focus
			ggoSpread.Source = frm1.vspdData 

			If .MaxRows <= 0 Then
				Exit Function
			ElseIf CheckTotRow(Frm1.vspdData, .ActiveRow) = True Then
				MsgBox "합계는 삭제할 수 없습니다.", vbCritical
				Exit Function
			Else
				lDelRows = ggoSpread.DeleteRow
				lgBlnFlgChgValue = True
				lDelRows = CheckLastRow(Frm1.vspdData, lDelRows)
				If lDelRows > 0 Then ggoSpread.DeleteRow lDelRows
			End If
			Call CheckReCalc(0,0)
		End With
    Else
		With frm1.vspdData2 
			.focus
			ggoSpread.Source = frm1.vspdData2

			If .MaxRows <= 0 Then
				Exit Function
			ElseIf CheckTotRow(Frm1.vspdData2, .ActiveRow) = True Then
				MsgBox "합계는 삭제할 수 없습니다.", vbCritical
				Exit Function
			Else
				lDelRows = ggoSpread.DeleteRow
				lgBlnFlgChgValue = True
				lDelRows = CheckLastRow(Frm1.vspdData2, lDelRows)
				If lDelRows > 0 Then ggoSpread.DeleteRow lDelRows
			End If
			Call CheckReCalc2(0,0)
		End With    
    End If

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
	Dim bRtn1, bRtn2
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData
    bRtn1 = ggoSpread.SSCheckChange

    ggoSpread.Source = frm1.vspdData2
    bRtn2 = ggoSpread.SSCheckChange
    If bRtn1 = True Or bRtn2 = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 종료 하시겠습니까?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================
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
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows            
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '☜:  Run biz logic

    DbQuery = True  
    IsRunEvents = False
  
End Function

Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'If UNICDbl(frm1.txtW1.value) > 0 Or UNICDbl(frm1.vspdData.MaxRows) > 0 Or UNICDbl(frm1.vspdData2.MaxRows) > 0Then
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE
		    
		Call SetToolbar("1101111100000111")										<%'버튼 툴바 제어 %>
'	Else
'		Call SetToolbar("1100111100000111")										<%'버튼 툴바 제어 %>
'	End If
	If frm1.vspdData.MaxRows > 0 or frm1.vspdData2.MaxRows>0 Then
		Call QueryTotalLine
	End If
End Function


'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim pP21011
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow, lEndRow     
    Dim lRestGrpCnt 
    Dim strVal, strDel
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if
    strVal = ""
    strDel = ""
    lGrpCnt = 1
    
	With Frm1
		' ----- 1번째 그리드 
		For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
        
               Case  ggoSpread.InsertFlag                                      '☜: Insert
                                                  strVal = strVal & "C"  &  Parent.gColSep
                                                'strVal = strVal & lRow &  Parent.gColSep
                    .vspdData.Col = C_SEQ_NO1   : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W16		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W16_NM	: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W17       : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W18		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W19       : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W20		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W21		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_DESC1		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gRowSep

 
                    lGrpCnt = lGrpCnt + 1
                    
               Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                  strVal = strVal & "U"  &  Parent.gColSep
                                                  'strVal = strVal & lRow &  Parent.gColSep
                    .vspdData.Col = C_SEQ_NO1   : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W16		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W16_NM	: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W17       : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W18		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W19       : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W20		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W21		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_DESC1		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gRowSep
                    
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.DeleteFlag                                      '☜: Delete
                                                  strVal = strVal & "D"  &  Parent.gColSep
                                                  'strVal = strVal & lRow &  Parent.gColSep
                    .vspdData.Col = C_SEQ_NO1   : strVal = strVal & Trim(.vspdData.Text) &  Parent.gRowSep
   
                    lGrpCnt = lGrpCnt + 1
           End Select
		Next
		
       .txtSpread.value      = strVal
       strVal = ""
       
		' ----- 2번째 그리드 
 		For lRow = 1 To .vspdData2.MaxRows
    
           .vspdData2.Row = lRow
           .vspdData2.Col = 0
        
           Select Case .vspdData2.Text
        
               Case  ggoSpread.InsertFlag                                      '☜: Insert
													strVal = strVal & "C"  &  Parent.gColSep
													'strVal = strVal & lRow &  Parent.gColSep
                    .vspdData2.Col = C_SEQ_NO2		: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W22			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W23			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W23_NM		: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W24			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W25			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W26			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W27			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W28			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W29			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W30			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W31			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W32			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_DESC2		: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gRowSep
 
                    lGrpCnt = lGrpCnt + 1
                    
               Case  ggoSpread.UpdateFlag                                      '☜: Update
													strVal = strVal & "U"  &  Parent.gColSep
													'strVal = strVal & lRow &  Parent.gColSep
                    .vspdData2.Col = C_SEQ_NO2		: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W22			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W23			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W23_NM		: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W24			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W25			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W26			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W27			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W28			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W29			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W30			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W31			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W32			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_DESC2		: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gRowSep
  
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.DeleteFlag                                      '☜: Delete
													strVal = strVal & "D"  &  Parent.gColSep
													'strVal = strVal & lRow &  Parent.gColSep
                    .vspdData2.Col = C_SEQ_NO2      : strVal = strVal & Trim(.vspdData2.Text) &  Parent.gRowSep
   
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
      
		.txtMode.value        =  Parent.UID_M0002
		'.txtUpdtUserId.value  =  Parent.gUsrID
		'.txtInsrtUserId.value =  Parent.gUsrID
		.txtMaxRows.value     = lGrpCnt-1 
		.txtSpread2.value      = strVal
		.txtFlgMode.value     = lgIntFlgMode
		
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
    DbSave = True                                                           
End Function


Function DbSaveOk()													        <%' 저장 성공후 실행 로직 %>
	Call InitVariables
	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
	frm1.vspdData2.MaxRows = 0
    ggoSpread.Source = frm1.vspdData2
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


'========================================================================================
Function DbDeleteOk()
	lgBlnFlgChgValue = False
	Call FncNew()
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
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
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT>&nbsp;
						<a href="vbscript:OpenRefDebt">채권금액조회</A>|<a href="vbscript:GetRef">금액불러오기</A>|<a href="vbscript:OpenRefMenu">소득금액합계표조회</A>
					</TD>
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
						<DIV ID="ViewDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%; overflow=auto"><% ' -- overflow=auto : 컨텐츠 구역을 브라우저 크기에 따라 스크롤바가 생성되게 한다 %>
						<TABLE <%=LR_SPACE_TYPE_60%> CLASS="TD61" BORDER=0>
                            <TR>
                                <TD WIDTH="100%" VALIGN=TOP>
                                   <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN="LEFT"> 1. 대손충당금 조정</LEGEND><BR>
                                   ㄱ.손금산입액조정<BR>
                                   <TABLE width="100%" bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table1">
									   <TR>
									       <TD CLASS="TD51" width="20%" ALIGN=CENTER rowspan="2">(1)채권잔액<br>(21)의 금액</TD>
									       <TD CLASS="TD51" width="20%" ALIGN=CENTER colspan="3">(2)설정률</TD>
									       <TD CLASS="TD51" width="10%" ALIGN=CENTER rowspan="2">(3)한도액<BR>(1) X (2)</TD>
								           <TD CLASS="TD51" width="40%" ALIGN=CENTER colspan="3">회사계상액</TD>
								           <TD CLASS="TD51" width="10%" ALIGN=CENTER rowspan="2">(7)한도초과액<BR>{(6)-(3)}</TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" width="7%" ALIGN=CENTER>(ㄱ)<BR>1(2)/100</TD>
									       <TD CLASS="TD51" width="7%" ALIGN=CENTER>(ㄴ)<BR>실적률</TD>
									       <TD CLASS="TD51" width="6%" ALIGN=CENTER>(ㄷ)<BR>표준<BR>비율</TD>
									       <TD CLASS="TD51" width="12%" ALIGN=CENTER>(4)당기계상액</TD>
									       <TD CLASS="TD51" width="12%" ALIGN=CENTER>(5)보충액</TD>
									       <TD CLASS="TD51" width="16%" ALIGN=CENTER>(6)계</TD>
									  </TR>
									  <TR>
											<TD CLASS="TD61" width="20%" align=center><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW1" name=txtW1 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD61" width="7%" align=right><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW2_1_2" name=txtW2_1_2 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="14X2Z" width= 100%></OBJECT>');</SCRIPT>
											<input name="txtW2_1_1" tag="14XZ0" type="hidden">
										    </TD>
											<TD CLASS="TD61" width="7%" align=right><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW2_2" name=txtW2_2 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21XZ0" width = 85%></OBJECT>');</SCRIPT>%</TD>
											<TD CLASS="TD61" width="6%" align=center><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW2_3" name=txtW2_3 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21XZ0" width = 85%></OBJECT>');</SCRIPT>%</TD>
											<TD CLASS="TD61" width="10%" align=center><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW3" name=txtW3 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD61" width="12%" align=center><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW4" name=txtW4 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD61" width="12%" align=center><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW5" name=txtW5 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD61" width="16%" align=center><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW6" name=txtW6 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD61" width="10%" align=center><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW7" name=txtW7 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
								  </table><BR>
                                   ㄴ.익금산입액조정<BR>
                                   <TABLE width="100%" bgcolor=#696969  border=0 cellpadding=1 cellspacing=1 ID="Table2">
									   <TR>
									       <TD CLASS="TD51" width="15%" ALIGN=CENTER>(8)장부상<br>충당금<br>기초잔액</TD>
									       <TD CLASS="TD51" width="10%" ALIGN=CENTER>(9)기중<br>충당금<br>환입액</TD>
									       <TD CLASS="TD51" width="10%" ALIGN=CENTER>(10)충당금<br>부인<br>누계액</TD>
								           <TD CLASS="TD51" width="15%" ALIGN=CENTER>(11)당기대손금<br>상계액<br>{(27)의금액}</TD>
								           <TD CLASS="TD51" width="10%" ALIGN=CENTER>(12)당기설정<br>충당금<br>보충액</TD>
								           <TD CLASS="TD51" width="12%" ALIGN=CENTER>(13)환입할금액<br>{(8)-(9)-(10)<br>-(11)-(12)}</TD>
								           <TD CLASS="TD51" width="13%" ALIGN=CENTER>(14)회사<br>환입액</TD>
								           <TD CLASS="TD51" width="15%" ALIGN=CENTER>(15)과소환입·<br>과다환입(△)<br>{(13)－(14)}</TD>
									  </TR>
									  <TR>
											<TD CLASS="TD61" width="15%" align=center><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8" name=txtW8 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD61" width="10%" align=center><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW9" name=txtW9 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD61" width="10%" align=center><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW10" name=txtW10 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD61" width="15%" align=center><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW11" name=txtW11 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD61" width="10%" align=center><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW12" name=txtW12 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD61" width="12%" align=center><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW13" name=txtW13 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X2X" width = 100%></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD61" width="13%" align=center><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW14" name=txtW14 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD61" width="15%" align=center><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW15" name=txtW15 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X2X" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
								  </table><BR>
								  채  권  잔  액<BR>
                                   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=200 tag="23" TITLE="SPREAD" id=vaSpread> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT><BR>
								  </FIELDSET>
								</TD>
							</TR>
							<TR>
								<TD WIDTH=100% valign=top>
                                   <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN="LEFT">  2.대손금조정</LEGEND><BR>
                                   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=200 tag="23" TITLE="SPREAD" id=vaSpread2> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								  </FIELDSET>
								</TD>
							</TR>
						</TABLE>
						</DIV>
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
				        <TD><INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check1" ><LABEL FOR="prt_check1"><별지>채권잔액</LABEL>&nbsp;
				                                 <INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check2" ><LABEL FOR="prt_check2"><별지>대손금조정</LABEL>&nbsp;
	
                </TR>
			
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint('VIEW')" Flag=1>미리보기</BUTTON>&nbsp;
						<BUTTON NAME="bttnPrint"	CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint('PRINT')"   Flag=1>인쇄</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" style="display:none"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread2" tag="24" style="display:none"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows2" tag="24">
<INPUT TYPE=HIDDEN NAME="txtW1_BEFORE" tag="24" value="0">
<INPUT TYPE=HIDDEN NAME="txtW4_BEFORE" tag="24" value="0">
<INPUT TYPE=HIDDEN NAME="txtW10_BEFORE" tag="24" value="0">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
