<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 소득금액조정 
'*  3. Program ID           : w5107mA1
'*  4. Program Name         : w5107mA1.asp
'*  5. Program Desc         : 제21호 기부금 조정명세서 
'*  6. Modified date(First) : 2005/02/17
'*  7. Modified date(Last)  : 2005/02/17
'*  8. Modifier (First)     : LSHSAT
'*  9. Modifier (Last)      : 
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

Const BIZ_MNU_ID		= "w5107mA1"
Const BIZ_PGM_ID		= "w5107mB1.asp"											 '☆: 비지니스 로직 ASP명 
Const EBR_RPT_ID	    = "w5107OA1"

' -- 그리드 컬럼 정의 
Dim C_W_YEAR
Dim C_W_TYPE
Dim C_W_NAME
Dim C_W26
Dim C_W27
Dim C_W28
Dim C_W29


Dim IsOpenPop  
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()
	C_W_YEAR	= 1
	C_W_TYPE	= 2
	C_W_NAME	= 3
	C_W26		= 4
	C_W27		= 5
	C_W28		= 6
	C_W29		= 7
End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
    lgRefMode = False

End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub



'============================================  콤보 박스 채우기  ====================================

Sub InitComboBox()
	' 조회조건(구분)
	Dim IntRetCD1
	
	Call CommonQueryRs("MINOR_CD,MINOR_NM"," dbo.ufn_TB_MINOR('W1018', '" & C_REVISION_YM & "') ","  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
	
End Sub


Sub InitSpreadComboBox()
    Dim IntRetCD1

'	IntRetCD1 = CommonQueryRs("MINOR_CD, MINOR_NM", "dbo.ufn_TB_MINOR('W1008', '" & C_REVISION_YM & "')", "  ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  

'	If IntRetCD1 <> False Then
'		ggoSpread.Source = Frm1.vspdData
'		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_W6
'		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_W6_NM
'	End If

End Sub

' 여기선 사용안함.
Function OpenAccount()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strWhere

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "계정 팝업"					<%' 팝업 명칭 %>
	arrParam(1) = "TB_WORK_6"					<%' TABLE 명칭 %>
	

	frm1.vspdData.Col = C_W1
    arrParam(2) = frm1.vspdData.Text		<%' Code Condition%>

	arrParam(3) = ""							<%' Name Cindition%>

	strWhere = " ACCT_CD IN (SELECT ACCT_CD FROM TB_ACCT_MATCH (NOLOCK) WHERE MATCH_CD = '18')"
	strWhere = strWhere & " AND CO_CD = '" & frm1.txtCO_CD.value & "' "
	strWhere = strWhere & " AND FISC_YEAR = '" & frm1.txtFISC_YEAR.text & "' "
	strWhere = strWhere & " AND REP_TYPE = '" & frm1.cboREP_TYPE.value & "' "
	
	arrParam(4) = strWhere							<%' Where Condition%>
	arrParam(5) = "계정"						<%' 조건필드의 라벨 명칭 %>
	
    arrField(0) = "ACCT_CD"					<%' Field명(0)%>
    arrField(1) = "ACCT_NM"					<%' Field명(1)%>
    
    arrHeader(0) = "계정코드"					<%' Header명(0)%>
    arrHeader(1) = "계정명"						<%' Header명(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetAccount(arrRet)
	End If	
	
End Function

Function SetAccount(byval arrRet)
    With frm1
		.vspdData.Col = C_W1
		.vspdData.Text = arrRet(0)
		.vspdData.Col = C_W1_NM
		.vspdData.Text = arrRet(1)
	    ggoSpread.Source = frm1.vspdData
	    ggoSpread.UpdateRow frm1.vspdData.ActiveRow
		lgBlnFlgChgValue = True
	End With
End Function


Sub InitSpreadSheet()
	Dim ret, iRow
	
    Call initSpreadPosVariables()  

	Call AppendNumberPlace("6","3","2")
	
	' 1번 그리드 

	With Frm1.vspdData
				
		ggoSpread.Source = Frm1.vspdData
		'patch version
		ggoSpread.Spreadinit "V20041222_1",,parent.gForbidDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W29 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols									'☆: 사용자 별 Hidden Column
		.ColHidden = True    
 
		.MaxRows = 0
		ggoSpread.ClearSpreadData

'	    ggoSpread.SSSetEdit		C_SEQ_NO,	"순번",				5,,,6,1	' 히든컬럼 
		ggoSpread.SSSetMask		C_W_YEAR,	"사업연도",			15, 2, "9999" 
		ggoSpread.SSSetEdit		C_W_TYPE,	"기부금코드",			7,,,10,1
		ggoSpread.SSSetEdit		C_W_NAME,	"기부금 종류",			25,,,50,1

	    ggoSpread.SSSetFloat	C_W26,		"(26)한도초과손금" & VbCrlf & "불산입액",		17,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetFloat	C_W27,		"(27)당해사업연도이전" & VbCrlf & "손금추인액누계액",		17,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetFloat	C_W28,		"(28)이월액잔액" & VbCrlf & "{(26)-(27)}",		17,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,,"" 
	    ggoSpread.SSSetFloat	C_W29,		"(29)당해사업연도" & VbCrlf & "손금추인액",		17,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 

		.rowheight(-1000) = 30

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_W_TYPE,C_W_TYPE,True)

		Call InitSpreadComboBox

		.ReDraw = true	

		Call SetSpreadLock()

	End With 
	
	
					
End Sub


'============================================  그리드 함수  ====================================

Sub InitData()
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

	Call GetFISC_DATE
	
	'Exit Sub
		
End Sub


Sub SetSpreadLock()

	With Frm1.vspdData
	
		ggoSpread.Source = Frm1.vspdData

		ggoSpread.SpreadUnLock C_W_YEAR, -1, C_W29	' 전체 적용 
		ggoSpread.SpreadLock C_W_NAME, -1, C_W_NAME	' 전체 적용 
		ggoSpread.SpreadLock C_W28, -1, C_W28	' 전체 적용 

'		ggoSpread.SSSetRequired C_W1, -1, -1
'		ggoSpread.SSSetRequired C_W1_NM, -1, -1
'		ggoSpread.SSSetRequired C_W2, -1, -1
'		ggoSpread.SSSetRequired C_W3, -1, -1
'		ggoSpread.SSSetRequired C_W6, -1, -1
'		ggoSpread.SSSetRequired C_W6_NM, -1, -1
		
	End With	
End Sub

' InsertRow/Copy 할때 호출됨 
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	Dim iRow, sITEM_CD

	With Frm1.vspdData

		ggoSpread.Source = Frm1.vspdData
		
		ggoSpread.SSSetRequired C_W_YEAR, pvStartRow, pvEndRow
		ggoSpread.SpreadLock C_W_NAME, -1, C_W_NAME
		ggoSpread.SpreadLock C_W28, -1, C_W28

		If pvEndRow > .MaxRows - 2 Then
			ggoSpread.SpreadLock C_W_YEAR, .MaxRows - 1, C_W29
			ggoSpread.SpreadLock C_W_YEAR, .MaxRows, C_W29
		End If
			
	End With	
End Sub

' -- 헤더쪽 그리드 재조정 
Sub RedrawSumRow()
	Dim iRow
	
	iRow = 1
	
	ggoSpread.Source = Frm1.vspdData
	With Frm1.vspdData
		ggoSpread.SpreadUnLock C_W1, iRow, C_W7, .MaxRows - 1	' 전체 적용 
		ggoSpread.SSSetRequired C_W1, iRow, .MaxRows - 1
		ggoSpread.SSSetRequired C_W1_NM, iRow, .MaxRows - 1
		ggoSpread.SSSetRequired C_W2, iRow, .MaxRows - 1
		ggoSpread.SSSetRequired C_W3, iRow, .MaxRows - 1
		ggoSpread.SSSetRequired C_W6, iRow, .MaxRows - 1
		ggoSpread.SSSetRequired C_W6_NM, iRow, .MaxRows - 1

		ggoSpread.SpreadLock C_W1,   .MaxRows, C_W7, .MaxRows

		.Row = .MaxRows
		Call .AddCellSpan(C_W1, .MaxRows, 3, 1) 
		.Col = C_W1	:	.CellType = 1	:	.Text = "합계"	:	.TypeHAlign = 2
		.Col = C_W6_NM	:	.CellType = 1
	End With	
End Sub


'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr, parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <>  parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub



Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_SEQ_NO		= iCurColumnPos(1)
            C_W1		= iCurColumnPos(2)
            C_W1_BT		= iCurColumnPos(3)
            C_W1_NM	= iCurColumnPos(4)
            C_W2		= iCurColumnPos(5)
            C_W3		= iCurColumnPos(6)
            C_W4		= iCurColumnPos(7)
            C_W5		= iCurColumnPos(8)
            C_W6		= iCurColumnPos(9)
            C_W6_NM		= iCurColumnPos(10)
            C_W7		= iCurColumnPos(11)
    End Select    
End Sub

'============================== 레퍼런스 함수  ========================================
Function GetRef()	' 금액가져오기 링크 클릭시 

	Dim IntRetCD , i
	Dim sMesg
	Dim sFiscYear, sRepType, sCoCd
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6,BackColor_w,BackColor_g
    
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	if wgConfirmFlg = "Y" then
		Exit function
	end if
	
	wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
	
	sMesg = wgRefDoc & vbCrLf & vbCrLf
	BackColor_g = frm1.txtW1.BackColor
	BackColor_w = frm1.txtW2.BackColor

	frm1.txtW1.BackColor =&H009BF0A2&
	frm1.txtW2.BackColor =&H009BF0A2&
	frm1.txtW3.BackColor =&H009BF0A2&
	frm1.txtW4.BackColor =&H009BF0A2&
	frm1.txtW8.BackColor =&H009BF0A2&
	frm1.txtW14.BackColor =&H009BF0A2&
	frm1.txtW19.BackColor =&H009BF0A2&
	
	IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '⊙: "Will you destory previous data"
	
	frm1.txtW1.BackColor = BackColor_g
	frm1.txtW2.BackColor = BackColor_w
	frm1.txtW3.BackColor = BackColor_g
	frm1.txtW4.BackColor = BackColor_g
	frm1.txtW8.BackColor = BackColor_g
	frm1.txtW14.BackColor = BackColor_g
	frm1.txtW19.BackColor = BackColor_g

	If IntRetCD = vbNo Then
		Exit Function
	End If

	call CommonQueryRs("W_R1, W2, W3, W4, W8, W14, W19 ","dbo.ufn_TB_21_GetRef('" & sCoCd & "', '" & sFiscYear & "', '" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	If lgF0 = "" Then	 Exit Function
	
	frm1.txtW_R1.Value = REPLACE(lgF0, chr(11),"")
	frm1.txtW2.Value = REPLACE(lgF1, chr(11),"")
	frm1.txtW3.Value = REPLACE(lgF2, chr(11),"")
	frm1.txtW4.Value = REPLACE(lgF3, chr(11),"")
	frm1.txtW8.Value = REPLACE(lgF4, chr(11),"")
	frm1.txtW14.Value = REPLACE(lgF5, chr(11),"")
	frm1.txtW19.Value = REPLACE(lgF6, chr(11),"")

	Call txtW2_change

'	frm1.txtW7_C.value =  cdbl(arrW2)
'	Call txtW7_C_change

End Function

Function GetRefOk()													<%'조회 성공후 실행로직 %>
	
    Dim iNameArr, iSeqNo, iLastRow, iRow
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    lgIntFlgMode = parent.OPMD_CMODE
    lgBlnFlgChgValue = True
    Call Fn_txtCalc
	Call SetToolbar("1100111100001111")										<%'버튼 툴바 제어 %>

End Function

Function ChangeRowFlg(iObj)
	Dim iRow
	
	With iObj
		ggoSpread.Source = iObj
		
		For iRow = 1 To .MaxRows
			.Col = 0 : .Row = iRow : .Value = ggoSpread.InsertFlag
		Next
	End With
End Function


Sub GetFISC_DATE()	' 법인의 조회조건에 만족하는 당기시작,종료일을 가져온다.

		
End Sub

Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub


Sub txtw2_Change( )
    lgBlnFlgChgValue = True
    Call Fn_txtCalc
End Sub

Sub txtw3_Change( )
    lgBlnFlgChgValue = True
    Call Fn_txtCalc
End Sub

Sub txtw4_Change( )
    lgBlnFlgChgValue = True
    Call Fn_txtCalc
End Sub

Sub txtw8_Change( )
    lgBlnFlgChgValue = True
    Call Fn_txtCalc
End Sub

Sub txtw14_Change( )
    lgBlnFlgChgValue = True
    Call Fn_txtCalc
End Sub

'============================================  조회조건 함수  ====================================


'============================================  폼 함수  ====================================
Sub Form_Load()
   
    Call LoadInfTB19029                                                       <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
		
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100111100001111")										<%'버튼 툴바 제어 %>
	  
	' 변경한곳 

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
 
	Call InitComboBox	' 먼저해야 한다. 기업의 회계기준일을 읽어오기 위해 
	Call ggoOper.ClearField(Document, "1")	
	Call InitData

	Call FncQuery
    
    
End Sub

'============================================  사용자 함수  ====================================
Function  Fn_txtCalc()
	Dim dblW1, dblW3, dblW1Max, dblW6, dblW7, dblW9, dblW12, dblW15, dblW18, dblW28
	
	' (1) 계산 : 표준손익계산서 당기순이익 + 제15호 익금산입 합계 - 제15호 손금산입 합계 + (3) + (4) + (8) + (14) 을 계산하여 입력함.
	Frm1.txtW1.Value = UNICdbl(Frm1.txtW_R1.Value) + UNICdbl(Frm1.txtW3.Value) + UNICdbl(Frm1.txtW4.Value) + UNICdbl(Frm1.txtW8.Value) + UNICdbl(Frm1.txtW14.Value)
	
	' (5) 계산 : MIN(max((①-②),0),(③+④))를 계산하여 입력함.
	dblW1 = UNICdbl(Frm1.txtW1.Value) - UNICdbl(Frm1.txtW2.Value)
	dblW3 = UNICdbl(Frm1.txtW3.Value) + UNICdbl(Frm1.txtW4.Value)
	If dblW1 < 0 Then
		dblW1Max = 0
	Else
		dblW1Max = dblW1
	End If
	If dblW1Max > dblW3 Then
		Frm1.txtW5.Value = dblW3
	Else
		Frm1.txtW5.Value = dblW1Max
	End If
	
	' (6) 계산 : Max ( ①－② - ⑤ , 0 )을 계산하여 입력함.
	dblW6 = UNICdbl(Frm1.txtW1.Value) - UNICdbl(Frm1.txtW2.Value) - UNICdbl(Frm1.txtW5.Value)
	If dblW6 > 0 Then
		Frm1.txtW6.Value = dblW6
	Else
		Frm1.txtW6.Value = 0
	End If
	
	' (7) 계산 : max [(③+④)-max(0,(①-②)), 0 ]을 계산하여 입력함.
	dblW7 = dblW3 - dblW1Max
	If dblW7 > 0 Then
		Frm1.txtW7.Value = dblW7
	Else
		Frm1.txtW7.Value = 0
	End If

	' (9) 계산 : ⑥×50% 계산하여 입력함.
	dblW6 = UNICdbl(Frm1.txtW6.Value)
	Frm1.txtW9.Value = dblW6 * 0.5
	
	' (10) : ⑧과⑨중 적은금액을 입력함.
	Frm1.txtW11.Value = 0
	If UNICdbl(Frm1.txtW9.Value) > UNICdbl(Frm1.txtW8.Value) Then
		Frm1.txtW10.Value = Frm1.txtW8.Value

		' (11)⑨>⑧ 인 경우로서(⑨-⑧) 과 "(28)이월액" 잔액 합계액의 조특법73조 금액중 적은 금액을 입력함.
		dblW9 = UNICdbl(Frm1.txtW9.Value) - UNICdbl(Frm1.txtW8.Value)
		If Frm1.vspdData.MaxRows > 0 Then
			Frm1.vspdData.Row = Frm1.vspdData.MaxRows - 1
			Frm1.vspdData.Col = C_W28
			dblW28 = UNICdbl(Frm1.vspdData.Text)
		Else
			dblW28 = 0
		End If
		If dblW9 > dblW28 Then
			Frm1.txtW11.Value = dblW28
		Else
			Frm1.txtW11.Value = dblW9
		End If
	Else
		Frm1.txtW10.Value = Frm1.txtW9.Value
		Frm1.txtW11.Value = 0
	End If

	' (12) : max [ (⑥ - ⑩ - ⑪) , 0 ] 를 계산하여 입력함.
	dblW6 = UNICdbl(Frm1.txtW6.Value) - UNICdbl(Frm1.txtW10.Value) - UNICdbl(Frm1.txtW11.Value)
	If dblW6 > 0 Then
		Frm1.txtW12.Value = dblW6
	Else
		Frm1.txtW12.Value = 0
	End If
	
	' (13) : max [ (⑧ - ⑨ ) , 0 ] 를 계산하여 입력함.
	dblW9 = UNICdbl(Frm1.txtW8.Value) - UNICdbl(Frm1.txtW9.Value)
	If dblW9 > 0 Then
		Frm1.txtW13.Value = dblW9
	Else
		Frm1.txtW13.Value = 0
	End If

	' (15) : ⑫×5% 를 계산하여 입력함.
	dblW12 = UNICdbl(Frm1.txtW12.Value)
	Frm1.txtW15.Value = dblW12 * 0.05

	' (16) : ⑭와⑮중 적은금액을 입력함.
	dblW15 = 0
	If UNICdbl(Frm1.txtW15.Value) > UNICdbl(Frm1.txtW14.Value) Then
		Frm1.txtW16.Value = Frm1.txtW14.Value

		' (17) : ⑭<⑮인 경우로서(⑮-⑭)과 (28)이월액 잔액합계액의 지정기부금 중 적은금액을 입력함.
		dblW15 = UNICdbl(Frm1.txtW15.Value) - UNICdbl(Frm1.txtW14.Value)
		If Frm1.vspdData.MaxRows > 0 Then
			Frm1.vspdData.Row = Frm1.vspdData.MaxRows
			Frm1.vspdData.Col = C_W28
			dblW28 = UNICdbl(Frm1.vspdData.Text)
		Else
			dblW28 = 0
		End If
		If dblW15 > dblW28 Then
			Frm1.txtW17.Value = dblW28
		Else
			Frm1.txtW17.Value = dblW15
		End If
	Else
		Frm1.txtW16.Value = Frm1.txtW15.Value
	End If

	' (18) : max [ (⑭ - ⑮) , 0 ] 를 계산하여 입력함.
	dblW15 = UNICdbl(Frm1.txtW14.Value) - UNICdbl(Frm1.txtW15.Value)
	If dblW15 > 0 Then
		Frm1.txtW18.Value = dblW15
	Else
		Frm1.txtW18.Value = 0
	End If

	' (20)		(12) × 3% 를 계산하여 입력합니다.
	dblW12 = UNICdbl(Frm1.txtW12.Value)
	Frm1.txtW20.Value = dblW12 * 0.03

	' (21)		(18),(19),(20) 중 최소금액을 입력합니다.
	If UNICdbl(Frm1.txtW18.Value) > UNICdbl(Frm1.txtW19.Value) Then
		If UNICdbl(Frm1.txtW19.Value) > UNICdbl(Frm1.txtW20.Value) Then
			Frm1.txtW21.Value = Frm1.txtW20.Value
		Else
			Frm1.txtW21.Value = Frm1.txtW19.Value
		End If
	Else
		If UNICdbl(Frm1.txtW18.Value) > UNICdbl(Frm1.txtW20.Value) Then
			Frm1.txtW21.Value = Frm1.txtW20.Value
		Else
			Frm1.txtW21.Value = Frm1.txtW18.Value
		End If
	End If

	' (22)		max ((18)-(21) , 0) 를 계산하여 입력합니다.
	dblW18 = UNICdbl(Frm1.txtW18.Value) - UNICdbl(Frm1.txtW21.Value)
	If dblW18 > 0 Then
		Frm1.txtW22.Value = dblW18
	Else
		Frm1.txtW22.Value = 0
	End If


	' (23) : ③+④+⑧+⑭ 를 계산하여 입력함.
	Frm1.txtW23.Value = UNICdbl(Frm1.txtW3.Value) + UNICdbl(Frm1.txtW4.Value) + UNICdbl(Frm1.txtW8.Value) + UNICdbl(Frm1.txtW14.Value)

	' (20) : (5) + (10) + (16) + (21) 를 계산하여 입력합니다.
	Frm1.txtW24.Value = UNICdbl(Frm1.txtW5.Value) + UNICdbl(Frm1.txtW10.Value) + UNICdbl(Frm1.txtW16.Value) + UNICdbl(Frm1.txtW21.Value)

	' (25) : Max ( (23) - (24) , 0 ) 를 계산하여 입력합니다.
	dblW18 = UNICdbl(Frm1.txtW23.Value) - UNICdbl(Frm1.txtW24.Value)
	If dblW18 > 0 Then
		Frm1.txtW25.Value = dblW18
	Else
		Frm1.txtW25.Value = 0
	End If
	'Frm1.txtW21.Value = UNICdbl(Frm1.txtW7.Value) + UNICdbl(Frm1.txtW13.Value) + UNICdbl(Frm1.txtW17.Value)

End Function

Function Fn_gridCalc(ByVal pCol, ByVal pRow)
	Dim iRow
	Dim dblC30W26, dblC30W27, dblC30W28, dblC30W29
	Dim dblC40W26, dblC40W27, dblC40W28, dblC40W29
	
	If Frm1.vspdData.MaxRows <= 0 Then Exit Function
	
	Call Fn_gridChck(pCol, pRow)
	

	dblC30W26 = 0	:	dblC30W27 = 0	:	dblC30W28 = 0	:	dblC30W29 = 0
	dblC40W26 = 0	:	dblC40W27 = 0	:	dblC40W28 = 0	:	dblC40W29 = 0
	
    ggoSpread.Source = Frm1.vspdData

	With Frm1.vspdData
		For iRow = 1 To .MaxRows - 2
			.Row = iRow	:	.Col = 0
			If .Text <> ggoSpread.DeleteFlag Then
				.Row = iRow	:	.Col = C_W_TYPE
				
				If .Text = "30" Then
					.Col = C_W26	:	dblC30W26 = dblC30W26 + UNICdbl(.Text)
					.Col = C_W27	:	dblC30W27 = dblC30W27 + UNICdbl(.Text)
					.Col = C_W28	:	dblC30W28 = dblC30W28 + UNICdbl(.Text)
					.Col = C_W29	:	dblC30W29 = dblC30W29 + UNICdbl(.Text)
				Else
					.Col = C_W26	:	dblC40W26 = dblC40W26 + UNICdbl(.Text)
					.Col = C_W27	:	dblC40W27 = dblC40W27 + UNICdbl(.Text)
					.Col = C_W28	:	dblC40W28 = dblC40W28 + UNICdbl(.Text)
					.Col = C_W29	:	dblC40W29 = dblC40W29 + UNICdbl(.Text)
				End If
			End If
			
		Next

		.Row = .MaxRows - 1
		.Col = C_W26	:	.Text = dblC30W26
		.Col = C_W27	:	.Text = dblC30W27
		.Col = C_W28	:	.Text = dblC30W28
		.Col = C_W29	:	.Text = dblC30W29
	    ggoSpread.UpdateRow .MaxRows - 1

		
		.Row = .MaxRows
		.Col = C_W26	:	.Text = dblC40W26
		.Col = C_W27	:	.Text = dblC40W27
		.Col = C_W28	:	.Text = dblC40W28
		.Col = C_W29	:	.Text = dblC40W29
	    ggoSpread.UpdateRow .MaxRows

	End With
	Call Fn_txtCalc
End Function

Function Fn_gridChck(ByVal pCol, ByVal pRow)
	Dim dblW26, dblW27, dblW28, dblW29
	
	With Frm1.vspdData
		If pRow > 0 Then
			.Row = pRow	:	.Col = C_W26	:	dblW26 = UNICdbl(.Text)
			.Row = pRow	:	.Col = C_W27	:	dblW27 = UNICdbl(.Text)
			.Row = pRow	:	.Col = C_W29	:	dblW29 = UNICdbl(.Text)
			
			If pCol = C_W26 Or pCol = C_W27 Then
				If dblW26 < dblW27 Then
					Call DisplayMsgBox("WC0010", "X", "(27)", "(26)")                          <%'WC0010  (26) < (27) Error!!%>
					.Col = pCol	:	.Text = 0
					Exit Function
				End If
			End If
			
			dblW28 = dblW26 - dblW27
			If dblW28 < dblW29 Then
				Call DisplayMsgBox("WC0010", "X", "(29)", "(28)")                          <%'WC0010  (24) < (25) Error!!%>
				.Col = pCol	:	.Text = 0
				Exit Function
			End If
			.Col = C_W28	:	.Text = dblW28
		End If
	End With
End Function

'============================================  이벤트 함수  ====================================
'============================================  이벤트 호출 함수  ====================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim iIdx

	With Frm1.vspdData
		Select Case Col
			Case C_W6
				.Col = Col	: .Row = Row
				iIdx = UNICDbl(.Value)
				.Col = Col +1
				.Value = iIdx
			Case C_W6_NM
				.Col = Col	: .Row = Row
				iIdx = UNICDbl(.Value)
				.Col = Col -1
				.Value = iIdx
		End Select
		
	End With
End Sub

Sub vspdData_Change(ByVal Col , ByVal Row )
	Dim dblSum, sTemp, iRow
	
	lgBlnFlgChgValue= True ' 변경여부 
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

    If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
    End If
	
    ggoSpread.Source = Frm1.vspdData
    ggoSpread.UpdateRow Row

	Select Case Col
		Case C_W26, C_W27, C_W28, C_W29
			Call Fn_gridCalc(Col, Row)	
		Case C_W_YEAR
			Frm1.vspdData.Col = Col	: Frm1.vspdData.Row = Row
			sTemp = Frm1.vspdData.Text
			If (Row mod 2) = 1 Then
				iRow = Row + 1
			Else
				iRow = Row - 1
			End If
			Frm1.vspdData.Row = iRow	:	Frm1.vspdData.Text = sTemp

    		ggoSpread.UpdateRow iRow

	End Select    
    
End Sub

Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = Frm1.vspdData
   
    If Frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
    	Exit Sub
       ggoSpread.Source = Frm1.vspdData
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If

	Frm1.vspdData.Row = Row
End Sub

Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = Frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If Frm1.vspdData.MaxRows = 0 Then
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

	ggoSpread.Source = Frm1.vspdData
End Sub    

Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = Frm1.vspdData
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
    
    if Frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(Frm1.vspdData,NewTop) Then	           
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
	With Frm1.vspdData
		If Row > 0 And Col = C_W1_BT Then
		    .Row = Row
		    .Col = C_W1_BT

		    Call OpenAccount()
		End If
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

Function FncQuery() 
    Dim IntRetCD , i, blnChange
    
    FncQuery = False                                                        
    blnChange = False
    
    Err.Clear                                                               <%'Protect system from crashing%>

	
<%  '-----------------------
    'Check previous data area
    '----------------------- %>
	ggoSpread.Source = Frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		blnChange = True
	End If
    
    If lgBlnFlgChgValue Or blnChange Then
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
    Call InitVariables													<%'Initializes local global variables%>
    Call InitData                              
    															
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If    

    'Call SetToolbar("1100111100001111")

     
    CALL DBQuery()
    
End Function

Function FncSave() 
    Dim blnChange, i, sMsg
    
    blnChange = False
    
    FncSave = False                                                         
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange <> False Then
		blnChange = True
'	    Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
'	    Exit Function
	End If

    If lgBlnFlgChgValue = False and blnChange = True Then
		Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
		Exit Function
	End If

	
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

' -- 2006.03.02 추가(cyt)
Function Verification()
	Dim IntRetCD, dblSum
	Verification = False
	
	Dim sWhere
	sWhere = "CO_CD=" & FilterVar(Trim(frm1.txtCO_CD.value),"''","S")
	sWhere = sWhere & " AND FISC_YEAR=" & FilterVar(Trim(frm1.txtFISC_YEAR.text),"''","S")
	sWhere = sWhere & " AND REP_TYPE=" & FilterVar(Trim(frm1.cboREP_TYPE.value),"''","S")

	IntRetCD = CommonQueryRs("W05, W54, W04", "TB_3", sWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  

	If IntRetCD <> False Then
		lgF0 = Replace(lgF0, Chr(11), "")
		lgF1 = Replace(lgF1, Chr(11), "")
		lgF2 = Replace(lgF2, Chr(11), "")
		
		If UNICDbl(frm1.txtW25.value) <> UNICDbl(lgF0) Then
			Call DisplayMsgBox("WC0004", parent.VB_INFORMATION, "(25)한도초과액합계", "제3호 법인세과세표준 및 세액조정계산서의 (105)기부금한도초과액")           '⊙: "Will you destory previous data"
			Exit Function
		End If

		If frm1.vspdData.MaxRows > 0 Then 
			dblSum = UNICDbl(GetGrid(C_W29, frm1.vspdData.MaxRows -1)) + UNICDbl(GetGrid(C_W29, frm1.vspdData.MaxRows))

			If dblSum <> UNICDbl(lgF1) Then
				Call DisplayMsgBox("WC0004", parent.VB_INFORMATION, "(29)당해사업연도 손금추인액", "제3호 법인세과세표준 및 세액조정계산서의 (106)기부금한도초과액 이월액손금산입")           '⊙: "Will you destory previous data"
				Exit Function
			End If
		ElseIf UNICDbl(lgF1) > 0 Then
			Call DisplayMsgBox("X", parent.VB_INFORMATION, "제3호 법인세과세표준 및 세액조정계산서의 (106)기부금한도초과액 이월액손금산입이 0보다 큰 경우 (29)당해사업연도 손금추인액의 합계가 존재해야 합니다.", "")           '⊙: "Will you destory previous data"
			Exit Function
		End If		
		
		dblSum = UNICDbl(frm1.txtW1.value) - (UNICDbl(frm1.txtW3.value) + UNICDbl(frm1.txtW4.value) + UNICDbl(frm1.txtW8.value) + UNICDbl(frm1.txtW14.value))
		If dblSum <> UNICDbl(lgF2) Then
			Call DisplayMsgBox("WC0004", parent.VB_INFORMATION, "기부금조정명세서의 소득금액계산에 적용된 차가감소득금액{ (1)- [(3)+(4)+(8)+(14)] }", "제3호 법인세과세표준 및 세액조정계산서의 (104)차가감소득금액")           '⊙: "Will you destory previous data"
			Exit Function
		End If

	End If
	
	
	Verification = True
End Function

Function GetGrid(Byval pCol, Byval pRow)
	With frm1.vspdData
		.Col = pCol : pRow = pRow : GetGrid = .Value
	End With
End Function

'========================================================================================
Function FncNew() 
    Dim IntRetCD , blnErr

    FncNew = False : blnErr = False

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

    Call SetToolbar("1100111100001111")
    lgIntFlgMode = parent.OPMD_CMODE

	frm1.txtCO_CD.focus

	' -- 3호 서식 체크후 데이타 없을시 에러
	Dim sWhere
	sWhere = "CO_CD=" & FilterVar(Trim(frm1.txtCO_CD.value),"''","S")
	sWhere = sWhere & " AND FISC_YEAR=" & FilterVar(Trim(frm1.txtFISC_YEAR.text),"''","S")
	sWhere = sWhere & " AND REP_TYPE=" & FilterVar(Trim(frm1.cboREP_TYPE.value),"''","S")

	IntRetCD = CommonQueryRs("W04", "TB_3", sWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  

	If IntRetCD = False Then
		blnErr = True
	Else
		lgF0 = Replace(lgF0, Chr(11), "")
		If UNICDbl(lgF0) = 0 Then	blnErr = True
	End If
	
	If blnErr Then
		Call DisplayMsgBox("W60006", parent.VB_INFORMATION, "(104) 차가감소득금액", "X")           '⊙: "Will you destory previous data"
		Call SetToolbar("1100000000001111")
	End If


	frm1.txtW1.value = 0
	frm1.txtW2.value = 0
	frm1.txtW3.value = 0
	frm1.txtW4.value = 0
	frm1.txtW8.value = 0
	frm1.txtW14.value = 0
	
    FncNew = True

End Function


Function FncCopy() 
	Dim iActiveRow
	
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If

	With Frm1.vspdData
	    ggoSpread.Source = Frm1.vspdData
	    iActiveRow = .ActiveRow

		If .ActiveRow > 0 Then
			.focus
			.ReDraw = False
		
			ggoSpread.CopyRow

			.Row = .ActiveRow
			.Col = C_W1
			.Text = ""
			
			.Col = C_W1_NM
			.Text = ""
    
			.Col = C_W2
			.Text = ""

			.Col = C_W3
			.Text = ""
    
			.ReDraw = True

			SetSpreadColor .ActiveRow, .ActiveRow
			Call SetDefaultVal(ActiveRow, ActiveRow)
		End If
	End With
	
    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 
    Dim lDelRows, dblSum, iRow

	With Frm1.vspdData
		.focus

		ggoSpread.Source = Frm1.vspdData
		If .MaxRows <= 0 Then
			Exit Function
		ElseIf CheckTotRow(Frm1.vspdData, .ActiveRow) = True Then
			MsgBox "합계는 삭제할 수 없습니다.", vbCritical
			Exit Function
		Else
			iRow = .ActiveRow
			lDelRows = CheckLastRow(Frm1.vspdData, iRow)
			If lDelRows > 0 Then
				ggoSpread.EditUndo lDelRows
				ggoSpread.EditUndo lDelRows - 1
			End If
			
			If (iRow Mod 2) = 1 Then
				iRow = iRow + 1
			End If
			lDelRows = ggoSpread.EditUndo(iRow)
			lDelRows = ggoSpread.EditUndo(iRow - 1)
			
			lgBlnFlgChgValue = True
		End If
		
	End With

	Call Fn_gridCalc(0, 0)
	Call SetDefaultVal(1, Frm1.vspdData.MaxRows)


End Function

' -- 합계 행인지 체크(Header Grid)
Function CheckTotRow(Byref pObj, Byval pRow) 
	CheckTotRow = False
	pObj.Col = C_W_YEAR : pObj.Row = pRow
	If pObj.Text = "" Then Exit Function
	If pObj.Text = "합계" Then	 ' 합계 행 
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
		If iCnt = 4 Then
			CheckLastRow = iMaxRow
		ElseIf iCnt = 0 Then
			CheckLastRow = .MaxRows
		End If
	End With
	
End Function


Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim iRow

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

	With Frm1.vspdData
	
		.focus
		ggoSpread.Source = Frm1.vspdData

		if .MaxRows = 0 then

			' 한번에 네줄을 만들고 두줄을 묶고 기부금종류명을 넣는다.
			ggoSpread.InsertRow  ,4

'			ret = .AddCellSpan(C_W_YEAR	, 1, 1, 2)	' 사업연도 2열 합침 
'			ret = .AddCellSpan(C_W_YEAR	, 3, 1, 2)	' 사업연도 2열 합침 
			
			.Row = 1	:	.Col	= C_W_YEAR	:	.TypeHAlign = 2	:	.TypeVAlign = 2
			.Row = 3	:	.Col	= C_W_YEAR	:	.CellType = 1	:	.Text	= "합계"	:	.TypeHAlign = 2	:	.TypeVAlign = 2
			.Row = 4	:	.Col	= C_W_YEAR	:	.CellType = 1	:	.Text	= "합계"

			SetSpreadColor 1, 4
			Call SetDefaultVal(1, 4)
			
			.Row  = 1
			.ActiveRow = 1

		else
			iRow = .ActiveRow

			If iRow > .MaxRows - 2 Then	' -- 마지막 합계줄에서 InsertRow를 하면 상위에 추가한다.
				iRow = .MaxRows - 2
				ggoSpread.InsertRow iRow, 2 

'				ret = .AddCellSpan(C_W_YEAR	, iRow + 1, 1, 2)	' 사업연도 2열 합침 
				.Row = iRow + 1	:	.Col	= C_W_YEAR	:	.TypeHAlign = 2	:	.TypeVAlign = 2
				SetSpreadColor iRow + 1, iRow + 1
				Call SetDefaultVal(iRow + 1, 2)
			Else
				If (iRow Mod 2) = 1 Then iRow = iRow + 1
				ggoSpread.InsertRow iRow, 2 

'				ret = .AddCellSpan(C_W_YEAR	, iRow + 1, 1, 2)	' 사업연도 2열 합침 
				.Row = iRow + 1	:	.Col	= C_W_YEAR	:	.TypeHAlign = 2	:	.TypeVAlign = 2
				SetSpreadColor iRow + 1, iRow + 1
				Call SetDefaultVal(iRow + 1, 2)
			End If   
			.vspdData.Row  = iRow + 1
			.vspdData.ActiveRow = iRow +1
			
        End if 	
		
    End With

    'Call SetToolbar("1101111100001111")

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function

' 그리드에 SEQ_NO, TYPE 넣는 로직 
Function SetDefaultVal(iRow, iAddRows)
	
	Dim i
	
	With Frm1.vspdData
	
		For i = iRow to iRow + iAddRows -1
			.Row = i
			If (i Mod 2) = 1 Then
				Call .AddCellSpan(C_W_YEAR, i, 1, 2) 
				
				.Col = C_W_TYPE	:	.Text = "30"
				.Col = C_W_NAME	:	.Text = "조세특례제한법 제73조"
				.Col = C_W_YEAR	:	.TypeHAlign = 2	:	.TypeVAlign = 2
			Else
				.Col = C_W_TYPE	:	.Text = "40"
				.Col = C_W_NAME	:	.Text = "지정기부금"
				.Col = C_W_YEAR	:	.TypeHAlign = 2	:	.TypeVAlign = 2
			End If
		Next
	End With
End Function



Function FncDeleteRow() 
    Dim lDelRows, iRow, dblSum 

	With Frm1.vspdData
		.focus

		ggoSpread.Source = Frm1.vspdData
		If .MaxRows <= 0 Then
			Exit Function
		ElseIf CheckTotRow(Frm1.vspdData, .ActiveRow) = True Then
			MsgBox "합계는 삭제할 수 없습니다.", vbCritical
			Exit Function
		Else
			iRow = .ActiveRow
			lDelRows = CheckLastRow(Frm1.vspdData, iRow)
			If lDelRows > 0 Then
				ggoSpread.DeleteRow lDelRows
				ggoSpread.DeleteRow lDelRows - 1
			End If
			
			If (iRow Mod 2) = 1 Then
				iRow = iRow + 1
			End If
			lDelRows = ggoSpread.DeleteRow(iRow)
			lDelRows = ggoSpread.DeleteRow(iRow - 1)
			
			lgBlnFlgChgValue = True
		End If
		
	End With

	Call Fn_gridCalc(0, 0)
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
	
    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
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


	    strVal = strVal     & "&lgStrPrevKey="		& lgStrPrevKey             '☜: Next key tag
	    strVal = strVal     & "&txtMaxRows="		& Frm1.vspdData.MaxRows         '☜: Max fetched data

		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '☜:  Run biz logic

    DbQuery = True  
  
End Function

		
Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    Dim iNameArr
    Dim iDx, iRow, iMaxRows
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	lgBlnFlgChgValue = False
	
	With Frm1.vspdData
		If .MaxRows > 0 Or frm1.txtW1.value <> ""  Then
			'-----------------------
			'Reset variables area
			'-----------------------
			lgIntFlgMode = parent.OPMD_UMODE
			
	
			Call SetToolbar("1101111100001111")										<%'버튼 툴바 제어 %>
			
			If .MaxRows > 0 Then
				Call SetSpreadColor(1, Frm1.vspdData.MaxRows)
				Call SetDefaultVal(1, Frm1.vspdData.MaxRows)
	
	
				.Row = .MaxRows -1	:	.Col	= C_W_YEAR	:	.CellType = 1	:	.Text	= "합계"	:	.TypeHAlign = 2	:	.TypeVAlign = 2
				.Row = .MaxRows	:	.Col	= C_W_YEAR	:	.CellType = 1	:	.Text	= "합계"
			End If
		End If
	End With
	
	Frm1.vspdData.focus			
End Function


'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim lRow, lCol, lMaxRows, lMaxCols , i    
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
    
	With Frm1.vspdData

		ggoSpread.Source = Frm1.vspdData
		lMaxRows = .MaxRows : lMaxCols = .MaxCols
		
		' ----- 1번째 그리드 
		For lRow = 1 To .MaxRows

	       .Row = lRow
	       .Col = 0
	    
	       Select Case .Text
	           Case  ggoSpread.InsertFlag                                      '☜: Insert
	                                              strVal = strVal & "C"  &  Parent.gColSep
	           Case  ggoSpread.UpdateFlag                                      '☜: Update
	                                              strVal = strVal & "U"  &  Parent.gColSep
		       Case  ggoSpread.DeleteFlag                                      '☜: Delete
		                                          strVal = strVal & "D"  &  Parent.gColSep
	       End Select
	       
		  ' 모든 그리드 데이타 보냄     
		  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
				For lCol = 1 To lMaxCols
					.Col = lCol : strVal = strVal & Trim(.Text) &  Parent.gColSep
				Next
				strVal = strVal & Trim(.Text) &  Parent.gRowSep
		  End If  
		Next
	
	End With

	Frm1.txtSpread.value		=	strVal
	strVal = ""

	Frm1.txtMode.value			=	Parent.UID_M0002
	Frm1.txtFlgMode.Value 		=	lgIntFlgMode
	'.txtUpdtUserId.value  =  Parent.gUsrID
	'.txtInsrtUserId.value =  Parent.gUsrID
				
	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
	
    DbSave = True                                                           
End Function


Function DbSaveOk()		
	Dim iRow											        <%' 저장 성공후 실행 로직 %>
	Call InitVariables
	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
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
	Call FncNew()
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
					<TD WIDTH=* align=right><a href="vbscript:GetRef">금액불러오기</a></TD>
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
									<TD CLASS="TD6"><script language =javascript src='./js/w5107ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script></TD>
									<TD CLASS="TD5">법인명</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCO_CD" Size=10 tag="14">
										<INPUT TYPE=TEXT NAME="txtCO_NM" Size=20 tag="14">
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
					<TD WIDTH="100%" HEIGHT=* valign=top>
						<DIV ID="ViewDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%; overflow=auto"><% ' -- overflow=auto : 컨텐츠 구역을 브라우저 크기에 따라 스크롤바가 생성되게 한다 %>
						<TABLE <%=LR_SPACE_TYPE_60%> CLASS="TD61" BORDER=0>
							<TR>
							     <TD WIDTH="100%" VALIGN=TOP HEIGHT="100%"><BR>
							     1.법정기부금 등 손금산입한도액 계산<BR>
                                   <TABLE width="100%" bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table1">
									   <TR>
									       <TD CLASS="TD51" width="18%" ALIGN=CENTER>(1)소득금액계</TD>
									       <TD CLASS="TD51" width="15%" ALIGN=CENTER>(2)법인세법 제13조<BR> 제1호에 의한 이월<BR> 결손금합계액</TD>
									       <TD CLASS="TD51" width="12%" ALIGN=CENTER>(3)법인세법 제24조<BR> 제2항 기부금<BR> 해당금액</TD>
								           <TD CLASS="TD51" width="12%" ALIGN=CENTER>(4)조세특례제한법<BR>제76조 및 동법 제73<BR>조제1항제1호 기부금<BR>해당금액</TD>
								           <TD CLASS="TD51" width="12%" ALIGN=CENTER TITLE='금액이  음수(-)인 경우에는 “0”을 기입합니다.'>(5)손금산입액[{(1)-(2)}와<BR> {(3)+(4)}중 적은금액]</TD>
									       <TD CLASS="TD51" width="15%" ALIGN=CENTER TITLE='금액이  음수(-)인 경우에는 “0”을 기입합니다.'>(6)소득금액잔액<BR>{(1)-(2)-(5)}</TD>
									       <TD CLASS="TD51" width="15%" ALIGN=CENTER TITLE='금액이  음수(-)인 경우에는 “0”을 기입합니다.'>(7)법정기부금등<BR> 한도초과액<BR>[{(3)+(4)}-{(1)-(2)}]</TD>
									  </TR>
									  <TR>
											<TD CLASS="TD61" align=center><script language =javascript src='./js/w5107ma1_txtW1_txtW1.js'></script></TD>
											<TD CLASS="TD61" align=center><script language =javascript src='./js/w5107ma1_txtW2_txtW2.js'></script></TD>
											<TD CLASS="TD61" align=center><script language =javascript src='./js/w5107ma1_txtW3_txtW3.js'></script></TD>
											<TD CLASS="TD61" align=center><script language =javascript src='./js/w5107ma1_txtW4_txtW4.js'></script></TD>
											<TD CLASS="TD61" align=center><script language =javascript src='./js/w5107ma1_txtW5_txtW5.js'></script></TD>
											<TD CLASS="TD61" align=center><script language =javascript src='./js/w5107ma1_txtW6_txtW6.js'></script></TD>
											<TD CLASS="TD61" align=center><script language =javascript src='./js/w5107ma1_txtW7_txtW7.js'></script></TD>
									  </TR>
								  </table><BR>
                                   2.조세특례제한법 제 73조 기부금 손금산입 한도액 계산<BR>
                                   <TABLE width="100%" bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table2">
									   <TR>
									       <TD CLASS="TD51" width="18%" ALIGN=CENTER>(8)조세특례제한법 제73조<BR> 기부금 해당금액(제1호 제외)</TD>
									       <TD CLASS="TD51" width="12%" ALIGN=CENTER>(9)한도액<BR>{(6)X50%}</TD>
									       <TD CLASS="TD51" width="16%" ALIGN=CENTER>(10)손금산입액<BR>{(8)과(9)중 적은금액}</TD>
								           <TD CLASS="TD51" width="28%" ALIGN=CENTER>(11)이월액 잔액중 손금산입액<BR>[{(9)>(8)}인 경우로서{(9)-(8)}과 <BR>(28)이월액잔액 합계중 적은 금액]</TD>
								           <TD CLASS="TD51" width="12%" ALIGN=CENTER>(12)소득금액<BR> 차감잔액<BR>{(6)-(10)-(11)}</TD>
								           <TD CLASS="TD51" width="14%" ALIGN=CENTER TITLE='금액이  음수(-)인 경우에는 “0”을 기입합니다.'>(13)한도초과액<BR>{(8)-(9)}</TD>
									  </TR>
									  <TR>
											<TD CLASS="TD61" align=center><script language =javascript src='./js/w5107ma1_txtW8_txtW8.js'></script></TD>
											<TD CLASS="TD61" align=center><script language =javascript src='./js/w5107ma1_txtW9_txtW9.js'></script></TD>
											<TD CLASS="TD61" align=center><script language =javascript src='./js/w5107ma1_txtW10_txtW10.js'></script></TD>
											<TD CLASS="TD61" align=center><script language =javascript src='./js/w5107ma1_txtW11_txtW11.js'></script></TD>
											<TD CLASS="TD61" align=center><script language =javascript src='./js/w5107ma1_txtW12_txtW12.js'></script></TD>
											<TD CLASS="TD61" align=center><script language =javascript src='./js/w5107ma1_txtW13_txtW13.js'></script></TD>
									  </TR>
								  </table><BR>
                                   3.지정기부금 손금산입 한도액 계산<BR>
                                   <TABLE width="100%" bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table3">
									   <TR>
									       <TD CLASS="TD51" width="18%" ALIGN=CENTER>(14)지정기부금 해당금액((19)기부금 포함)</TD>
									       <TD CLASS="TD51" width="23%" ALIGN=CENTER>(15)지정기부금<BR>한도액{(12)X5%}</TD>
									       <TD CLASS="TD51" width="18%" ALIGN=CENTER>(16)손금산입액{(14)와<BR>(15)중 적은금액}</TD>
								           <TD CLASS="TD51" width="26%" ALIGN=CENTER>(17)이월액잔액중 손금산입액<BR> [(14)<(15)인 경우로서 {(15)-(14)}와<BR> (28)이월액 잔액 합계액중 적은 금액]</TD>
								           <TD CLASS="TD51" width="16%" ALIGN=CENTER TITLE='금액이  음수(-)인 경우에는 “0”을 기입합니다.'>(18)지정기부금 한도초과액(ㄱ) {(14)-(15)}</TD>
									  </TR>
									  <TR>
											<TD CLASS="TD61" align=center><script language =javascript src='./js/w5107ma1_txtW14_txtW14.js'></script></TD>
											<TD CLASS="TD61" align=center><script language =javascript src='./js/w5107ma1_txtW15_txtW15.js'></script></TD>
											<TD CLASS="TD61" align=center><script language =javascript src='./js/w5107ma1_txtW16_txtW16.js'></script></TD>
											<TD CLASS="TD61" align=center><script language =javascript src='./js/w5107ma1_txtW17_txtW17.js'></script></TD>
											<TD CLASS="TD61" align=center><script language =javascript src='./js/w5107ma1_txtW18_txtW18.js'></script></TD>
									  </TR>
								  </table><BR>
                                   4. 지정기부금 중 조세특례제한법 제73조 제2항 기부금 추가 손금산입액 계산<BR>
                                   <TABLE width="100%" bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table4">
									   <TR>
									       <TD CLASS="TD51" width="23%" ALIGN=CENTER>(19)지정기부금 중 조세특례제한법 제73조 제2항 해당금액</TD>
									       <TD CLASS="TD51" width="25%" ALIGN=CENTER>(20)법인세법§24(1) (제1호-제2호)x3% {(12)x3%}</TD>
									       <TD CLASS="TD51" width="25%" ALIGN=CENTER>(21)추가 손금산입액 {(18),(19),(20)중 최소금액}</TD>
									       <TD CLASS="TD51" width="27%" ALIGN=CENTER>(22)지정기부금 한도초과액(ㄴ) {(18)-(21)}</TD>
									  </TR>
									  <TR>
											<TD CLASS="TD61" align=center><script language =javascript src='./js/w5107ma1_txtW19_txtW19.js'></script></TD>
											<TD CLASS="TD61" align=center><script language =javascript src='./js/w5107ma1_txtW20_txtW20.js'></script></TD>
											<TD CLASS="TD61" align=center><script language =javascript src='./js/w5107ma1_txtW21_txtW21.js'></script></TD>
											<TD CLASS="TD61" align=center><script language =javascript src='./js/w5107ma1_txtW22_txtW22.js'></script></TD>
									  </TR>
								  </table><BR>
                                   5.기부금 한도액 초과액 총계<BR>
                                   <TABLE width="100%" bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table4">
									   <TR>
									       <TD CLASS="TD51" width="45%" ALIGN=CENTER>(23)기부금합계{(3)+(4)+(8)+(14)}</TD>
									       <TD CLASS="TD51" width="28%" ALIGN=CENTER>(24)손금산입합계{(5)+(10)+(16)+(21)}</TD>
									       <TD CLASS="TD51" width="27%" ALIGN=CENTER>(25)한도초과액합계{(23)-(24)}=(7)+(13)+(22)</TD>
									  </TR>
									  <TR>
											<TD CLASS="TD61" align=center><script language =javascript src='./js/w5107ma1_txtW23_txtW23.js'></script></TD>
											<TD CLASS="TD61" align=center><script language =javascript src='./js/w5107ma1_txtW24_txtW24.js'></script></TD>
											<TD CLASS="TD61" align=center><script language =javascript src='./js/w5107ma1_txtW25_txtW25.js'></script></TD>
									  </TR>
								  </table><BR>
								  6.조세특례제한법 제73조제1항제2호 내지 제15호 및 법인세법상 지정기부금(조특법§73②포함) 이월액 명세<BR>
								     <script language =javascript src='./js/w5107ma1_vspdData_vspdData.js'></script>
							    </TD>
							</TR>
						</TABLE>
						</DIV>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE CLASS="TB3" CELLSPACING=0>
			    <TR>
				        <TD WIDTH=10>&nbsp;</TD>
				        <TD><INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check1" ><LABEL FOR="prt_check1"><별지>이월액명세</LABEL>&nbsp;
				           
				        </TD>
				            
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" style="display:'none'"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream" tag="24">
<INPUT TYPE=HIDDEN NAME="txtW_R1" tag="24" Value="0">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
