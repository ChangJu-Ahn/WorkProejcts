<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 기준정보 
'*  3. Program ID           : wb111mA1
'*  4. Program Name         : wb111mA1.asp
'*  5. Program Desc         : 계정 Mapping
'*  6. Modified date(First) : 2005/03/04
'*  7. Modified date(Last)  : 2005/03/04
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
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  상수/변수 선언  ====================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_MNU_ID		= "wb111mA1"
Const BIZ_PGM_ID		= "wb111mB1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_REF_ID		= "WB111RA1.asp"

' -- 그리드 컬럼 정의 
Dim C_ACCT_CD
Dim C_ACCT_BT
Dim C_ACCT_NM
Dim C_BS_PL_FG
Dim C_BS_PL_NM
Dim C_GP_CD
Dim C_GP_BT
Dim C_GP_NM


Dim IsOpenPop  
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()

	C_ACCT_CD	= 1
	C_ACCT_BT	= 2
	C_ACCT_NM	= 3
	C_BS_PL_FG	= 4
	C_BS_PL_NM	= 5
	C_GP_CD		= 6
	C_GP_BT		= 7
	C_GP_NM		= 8
	
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
	
	Call CommonQueryRs("MINOR_CD,MINOR_NM"," dbo.ufn_TB_MINOR('W1018', '" & C_REVISION_YM & "') "," ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
    Call CommonQueryRs("MINOR_CD,MINOR_NM"," dbo.ufn_TB_MINOR('W1081', '" & C_REVISION_YM & "') "," ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2( frm1.cboBS_PL_FG ,"" & Chr(11) & lgF0  ,"전체" & Chr(11) & lgF1  ,Chr(11))
	
End Sub


Sub InitSpreadComboBox()
    Dim IntRetCD1
    Dim iArr, iCnt, i

	IntRetCD1 = CommonQueryRs("MINOR_CD,MINOR_NM", "dbo.ufn_TB_MINOR('W1081', '" & C_REVISION_YM & "')", " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  

	If IntRetCD1 <> False Then
		ggoSpread.Source = Frm1.vspdData
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_BS_PL_FG
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_BS_PL_NM
	End If



End Sub

Function OpenAccount()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	DIm strWhere

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "계정 팝업"					<%' 팝업 명칭 %>
	arrParam(1) = "TB_WORK_6"					<%' TABLE 명칭 %>
	

	frm1.vspdData.Col = C_ACCT_CD
    arrParam(2) = frm1.vspdData.Text		<%' Code Condition%>

	arrParam(3) = ""							<%' Name Cindition%>

	strWhere = " CO_CD = '" & frm1.txtCO_CD.value & "' "
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
		.vspdData.Col = C_ACCT_CD
		.vspdData.Text = arrRet(0)
		.vspdData.Col = C_ACCT_NM
		.vspdData.Text = arrRet(1)
	    ggoSpread.Source = frm1.vspdData
	    ggoSpread.UpdateRow frm1.vspdData.ActiveRow
		lgBlnFlgChgValue = True
	End With
End Function

Function OpenAccountGP()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	DIm strWhere, sBS_PL_FG

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "계정그룹 팝업"					<%' 팝업 명칭 %>
	arrParam(1) = "TB_ACCT_GP"					<%' TABLE 명칭 %>
	
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_GP_CD
    arrParam(2) = frm1.vspdData.Text		<%' Code Condition%>

	arrParam(3) = ""							<%' Name Cindition%>

	frm1.vspdData.Col = C_BS_PL_FG

	strWhere = " REVISION_YM = '" & C_REVISION_YM & "'" & vbCrLf
	strWhere = strWhere & " AND BS_PL_FG = '" & frm1.vspdData.text & "' " & vbCrLf
	strWhere = strWhere & " AND COMP_TYPE2 = ( SELECT COMP_TYPE2 FROM TB_COMPANY_HISTORY WHERE " 
	strWhere = strWhere & " CO_CD = '" & frm1.txtCO_CD.value & "' "
	strWhere = strWhere & " AND FISC_YEAR = '" & frm1.txtFISC_YEAR.text & "' "
	strWhere = strWhere & " AND REP_TYPE = '" & frm1.cboREP_TYPE.value & "' )"
	strWhere = strWhere & " AND sum_fg = 'O' "
	
	arrParam(4) = strWhere							<%' Where Condition%>
	arrParam(5) = "계정그룹"						<%' 조건필드의 라벨 명칭 %>
	
    arrField(0) = "GP_CD"					<%' Field명(0)%>
    arrField(1) = "GP_NM"					<%' Field명(1)%>
    
    arrHeader(0) = "계정그룹코드"					<%' Header명(0)%>
    arrHeader(1) = "계정그룹명"						<%' Header명(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetAccountGP(arrRet)
	End If	
	
End Function

Function SetAccountGP(byval arrRet)
    With frm1
		.vspdData.Col = C_GP_CD
		.vspdData.Text = arrRet(0)
		.vspdData.Col = C_GP_NM
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
		ggoSpread.Spreadinit "V20041222_1",,parent.gAllowDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_GP_NM + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols									'☆: 사용자 별 Hidden Column
		.ColHidden = True    
 
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		ggoSpread.SSSetEdit		C_ACCT_CD,		"계정코드",			10,,,20,1
	    ggoSpread.SSSetButton 	C_ACCT_BT
		ggoSpread.SSSetEdit		C_ACCT_NM,		"계정명",			20,,,50,1
	    ggoSpread.SSSetCombo	C_BS_PL_FG,		"재무제표 구분",	15
	    ggoSpread.SSSetCombo	C_BS_PL_NM,		"재무제표 구분",	15
		ggoSpread.SSSetEdit		C_GP_CD,		"계정그룹코드",		10,,,20,1
	    ggoSpread.SSSetButton 	C_GP_BT
		ggoSpread.SSSetEdit		C_GP_NM,		"계정그룹명",		20,,,50,1

	    ret = .AddCellSpan(C_ACCT_CD, -1000, 2, 1)
	    ret = .AddCellSpan(C_GP_CD, -1000, 2, 1)

	    
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_BS_PL_FG,C_BS_PL_FG,True)
		
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

'		ggoSpread.SSSetRequired C_ACCT_CD, -1, -1
		ggoSpread.SpreadLock C_ACCT_CD, -1, C_ACCT_CD	' 전체 적용 
		ggoSpread.SpreadLock C_ACCT_NM, -1, C_ACCT_NM	' 전체 적용 
		ggoSpread.SSSetRequired C_BS_PL_FG, -1, -1
		ggoSpread.SSSetRequired C_BS_PL_NM, -1, -1
		ggoSpread.SSSetRequired C_GP_CD, -1, -1
		ggoSpread.SpreadLock C_GP_NM, -1, C_GP_NM	' 전체 적용 

'		ggoSpread.SpreadUnLock C_W1, -1, C_W9	' 전체 적용 
'		ggoSpread.SpreadLock C_PGM_ID, -1, C_PGM_ID	' 전체 적용 
'		ggoSpread.SSSetRequired C_W8, -1, -1
	End With	
End Sub

' InsertRow/Copy 할때 호출됨 
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	Dim iRow, sITEM_CD

	With Frm1.vspdData

		ggoSpread.Source = Frm1.vspdData
		ggoSpread.SSSetRequired C_ACCT_CD, pvStartRow, pvEndRow
		ggoSpread.SpreadLock C_ACCT_NM, pvStartRow, C_ACCT_NM, pvEndRow
		ggoSpread.SSSetRequired C_BS_PL_FG, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_BS_PL_NM, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_GP_CD, pvStartRow, pvEndRow
		ggoSpread.SpreadLock C_GP_NM, pvStartRow, C_GP_NM, pvEndRow
			
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
            C_W7		= iCurColumnPos(10)
            C_W8		= iCurColumnPos(11)
            C_W9		= iCurColumnPos(12)
    End Select    
End Sub

'============================== 레퍼런스 함수  ========================================
Function GetRef()	' ERP계정불러오기 조회 

    Dim arrRet
    Dim arrParam(5)
    Dim arrField, arrHeader
    Dim IntRetCD
    Dim strData
	Dim arrRowVal
    Dim arrColVal, lLngMaxRow
    Dim iDx, iRow
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

'    iCalledAspName = AskPRAspName("WB111RA1")
    
'    If Trim(iCalledAspName) = "" Then
'        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "WB111RA1", "x")
'        IsOpenPop = False
'        Exit Function
'    End If
	strData = ""
	With Frm1.vspdData
			.Col = C_ACCT_CD
		For iRow = 1 To .MaxRows
			.Row = iRow
			strData = strData & Chr(11) & .Text
		Next
	End With
    
	arrParam(0) = frm1.txtCO_CD.Value
	arrParam(1) = frm1.txtCO_NM.Value		
	arrParam(2) = frm1.txtFISC_YEAR.Text		
	arrParam(3) = frm1.cboREP_TYPE.Value		
	arrParam(4) = strData

    arrRet = window.showModalDialog(BIZ_REF_ID, Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	If arrRet(0,0) <> "" Then
		arrRowVal = Split(arrRet(0,0), Parent.gRowSep)                                 '☜: Split Row    data
		lLngMaxRow = UBound(arrRowVal)

		For iDx = 1 To lLngMaxRow
		    arrColVal = Split(arrRowVal(iDx-1), Parent.gColSep)   

			If Frm1.vspdData.MaxRows > 0 Then
				For iRow = 1 To Frm1.vspdData.MaxRows 
					Frm1.vspdData.Row	= iRow
					Frm1.vspdData.Col	= C_ACCT_CD
					If Frm1.vspdData.Text	= arrColVal(1) Then
						Frm1.vspdData.Row	= iRow
						Exit For
					End If
				Next
				If iRow > Frm1.vspdData.MaxRows Then
					Call Fn_InsertRow(iRow)
'					iRow = iRow + 1
				End If
				Frm1.vspdData.Row	= iRow
			Else
				Call Fn_InsertRow(1)
				Frm1.vspdData.Row	= 1
			End If
			Frm1.vspdData.Col	= C_ACCT_CD
			Frm1.vspdData.Text	= arrColVal(1)
			Frm1.vspdData.Col	= C_ACCT_NM
			Frm1.vspdData.Text	= arrColVal(2)
		Next
		
	End IF
    
    IsOpenPop = False
    
End Function

Function GetRef2()	
	Call window.open("WB111MA1.txt", BIZ_MNU_ID, _
	"Width=700px,Height=450px,center= Yes,status=yes,toolbar=no,menubar=no,location=no,scrollbars=yes")

End Function

Sub Fn_InsertRow(ByVal pRow)
	With Frm1.vspdData
		ggoSpread.Source = Frm1.vspdData

		ggoSpread.InsertRow pRow,1
		If pRow = 1 Then
			SetSpreadColor 1, 1
		Else
			SetSpreadColor pRow, pRow + 1
		End If
	End With
End Sub


Sub GetFISC_DATE()	' 법인의 조회조건에 만족하는 당기시작,종료일을 가져온다.

		
End Sub

'============================================  조회조건 함수  ====================================


'============================================  폼 함수  ====================================
Sub Form_Load()
   
    Call LoadInfTB19029                                                       <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
		
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1110111100001111")										<%'버튼 툴바 제어 %>
	  
	' 변경한곳 

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
 
	Call InitComboBox	' 먼저해야 한다. 기업의 회계기준일을 읽어오기 위해 
	Call ggoOper.ClearField(Document, "1")	
	Call InitData

	
    
    
End Sub



'============================================  사용자 함수  ====================================

'============================================  이벤트 함수  ====================================

'============================================  이벤트 호출 함수  ====================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	With Frm1.vspdData
		.Row = Row
		Select Case Col
			Case C_BS_PL_FG
				.Col = Col
				intIndex = .Value
				.Col = C_BS_PL_FG + 1
				.Value = intIndex	
			Case  C_BS_PL_NM
				.Col = Col
				intIndex = .Value
				.Col = C_BS_PL_FG
				.Value = intIndex		

		End Select
	End With
End Sub

Sub vspdData_Change(ByVal Col , ByVal Row )
	Dim dblSum, strWhere, arrVal
	
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
		Case	C_ACCT_CD
	
			With frm1.vspdData
	
				If Len(.Text) > 0 Then
					.Row = Row

					strWhere = 			" AND CO_CD = '" & frm1.txtCO_CD.value & "' "
					strWhere = strWhere & " AND FISC_YEAR = '" & frm1.txtFISC_YEAR.text & "' "
					strWhere = strWhere & " AND REP_TYPE = '" & frm1.cboREP_TYPE.value & "' "

					.Col = Col
					If CommonQueryRs("ACCT_NM", " TB_WORK_6 (NOLOCK)" , "ACCT_CD = '" & .Text &"' " & strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
				    	.Col	= C_ACCT_NM
				    	arrVal				= Split(lgF0, Chr(11))
						.Text	= arrVal(0)
					Else
						.Text	= ""
						.Col	= C_ACCT_NM
						.Text	= ""
					End If
				Else
					.Col = C_ACCT_NM
					.Text = ""
				End If
			End With

		Case	C_GP_CD
			With frm1.vspdData
	
				If Len(.Text) > 0 Then
					.Row = Row

					.Col = C_BS_PL_FG
					strWhere = 			" AND REVISION_YM = '" & C_REVISION_YM & "' "
					strWhere = strWhere & " AND BS_PL_FG = '" & .text & "' "
					strWhere = strWhere & " AND sum_fg = 'O' "

					.Col = Col
					If CommonQueryRs("GP_NM", " TB_ACCT_GP (NOLOCK)" , "GP_CD = '" & .Text &"' " & strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
				    	.Col	= C_GP_NM
				    	arrVal				= Split(lgF0, Chr(11))
						.Text	= arrVal(0)
					Else
						.Text	= ""
						.Col	= C_GP_NM
						.Text	= ""
					End If
				Else
					.Col = C_GP_NM
					.Text = ""
				End If
			End With

		Case	C_BS_PL_NM
			Call vspdData_ComboSelChange(C_BS_PL_NM, Row)

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
    
  '  if Frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(Frm1.vspdData,NewTop) Then	           
  '  	If lgStrPrevKey <> "" Then                         
  '    	   Call DisableToolBar(Parent.TBC_QUERY)
'			If DbQuery = False Then
'				Call RestoreTooBar()
'			    Exit Sub
'			End If  				
 '   	End If
  '  End if
End Sub

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With Frm1.vspdData
		If Row > 0 And Col = C_ACCT_BT Then
		    .Row = Row
		    .Col = C_ACCT_BT

		    Call OpenAccount()
		ElseIf Row > 0 And Col = C_GP_BT Then
		    .Row = Row
		    .Col = C_GP_BT

		    Call OpenAccountGP()
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
'    Call InitData                              
    															
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If    

    Call SetToolbar("1110111100001111")

     
    CALL DBQuery()
    
End Function

Function FncSave() 
    Dim i, sMsg
    
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

	
	If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
	      Exit Function
	End If    

    'If Verification = False Then Exit Function
    
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          
    
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

    Call SetToolbar("1110111100001111")
    lgIntFlgMode = parent.OPMD_CMODE

	frm1.txtCO_CD.focus

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
			.Col = C_ACCT_CD	:	.Text = ""
			.Col = C_ACCT_NM	:	.Text = ""
			
			.ReDraw = True

			SetSpreadColor .ActiveRow, .ActiveRow
			Call SetDefaultVal(iActiveRow + 1, 1)
		End If
	End With
	
    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 
    Dim lDelRows, dblSum 

	With Frm1.vspdData
		.focus

		ggoSpread.Source = Frm1.vspdData
		If .MaxRows <= 0 Then
			Exit Function
		Else
			lDelRows = ggoSpread.EditUndo
			lgBlnFlgChgValue = True
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

		iRow = .ActiveRow

		ggoSpread.InsertRow ,imRow
		SetSpreadColor iRow, iRow + imRow

		.vspdData.Row  = iRow + 1
		.vspdData.ActiveRow = iRow +1
			
    End With

    Call SetToolbar("1110111100101111")

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function

Function FncDeleteRow() 
    Dim lDelRows, iActiveRow, dblSum 

	With Frm1.vspdData
		.focus

		ggoSpread.Source = Frm1.vspdData
		If .MaxRows <= 0 Then
			Exit Function
		Else
			lDelRows = ggoSpread.DeleteRow
			lgBlnFlgChgValue = True
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
		strVal = strVal     & "&cboBS_PL_FG="        & Frm1.cboBS_PL_FG.Value      '☜: Query Key  

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
	
	If Frm1.vspdData.MaxRows > 0  Then
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE
		

		Call SetToolbar("1111111100101111")										<%'버튼 툴바 제어 %>
'		Call SetSpreadColor(1, Frm1.vspdData.MaxRows)
		Call SetSpreadLock()
		
	End If
	
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

	Frm1.txtSpread.value      = strVal
	strVal = ""

	Frm1.txtMode.value        =  Parent.UID_M0002
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

Function BtnExcelLoad()
    Dim Y, z, X
    Dim istrFileName
    Dim IntRetCD
    Dim listcount, handle, iRow
    Dim List(10)

	If Frm1.vspdData.MaxRows > 0 Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")           '⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Call LayerShowHide(1)
				
	Frm1.vspdData1.ScriptEnhanced = True
	istrFileName = Trim(Frm1.txtFILE_NAME.Value)
	
	X = Frm1.vspdData1.IsExcelFile(Trim(istrFileName))
	
    If istrFileName <> "" And X = 1 Then
    	
		ggoSpread.Source = Frm1.vspdData1
		ggoSpread.ClearSpreadData
        Y = Frm1.vspdData1.ScriptGetExcelSheetList(Trim(istrFileName), List, listcount, "", handle, True)
        ' If received sheet list, tell user, import file,
        ' and set result to z
        If Y = True Then
            z = Frm1.vspdData1.ImportExcelSheet(handle, 0)
            ' Tell user result based on true/false value of z
            If z = True Then
				ggoSpread.Source = Frm1.vspdData
				ggoSpread.ClearSpreadData

				Frm1.vspdData.MaxRows = Frm1.vspdData1.MaxRows

				Call SetSpreadColor(1, Frm1.vspdData.MaxRows)
				
				For iRow = 1 To Frm1.vspdData1.MaxRows
					Frm1.vspdData.Row = iRow	:	Frm1.vspdData1.Row = iRow
					
					Frm1.vspdData.Col = C_ACCT_CD	:	Frm1.vspdData1.Col = 1

					If Trim(CStr(Frm1.vspdData1.Text)) = "" Then
						Exit For
					End If
					Frm1.vspdData.Text = CStr(CLng(Frm1.vspdData1.Text))
'					Call vspdData_Change(C_ACCT_CD, iRow)
					
					Frm1.vspdData.Col = C_ACCT_NM	:	Frm1.vspdData1.Col = 2
					Frm1.vspdData.Text = CStr(Frm1.vspdData1.Text)

					Frm1.vspdData.Col = C_BS_PL_FG	:	Frm1.vspdData1.Col = 3
					Frm1.vspdData.Text = CStr(CLng(Frm1.vspdData1.Text))
'					Call vspdData_ComboSelChange(C_BS_PL_FG, iRow)
					Frm1.vspdData.Col = C_BS_PL_NM	:	Frm1.vspdData1.Col = 4
					Frm1.vspdData.Text = CStr(Frm1.vspdData1.Text)

		
					Frm1.vspdData.Col = C_GP_CD	:	Frm1.vspdData1.Col = 5
					Frm1.vspdData.Text = CStr(CLng(Frm1.vspdData1.Text))
'					Call vspdData_Change(C_GP_CD, iRow)

					Frm1.vspdData.Col = C_GP_NM	:	Frm1.vspdData1.Col =6
					Frm1.vspdData.Text = CStr(Frm1.vspdData1.Text)

					Frm1.vspdData.Col = 0
					Frm1.vspdData.Text = ggoSpread.InsertFlag
				Next
				Frm1.vspdData.MaxRows = iRow - 1

				Call SetSpreadColor(1, Frm1.vspdData.MaxRows)

            Else
                MsgBox "엑셀 파일 Import에 실패하였습니다."
            End If
        Else
            ' Tell user cannot obtain sheet list
            MsgBox "Cannot return information for Excel file."
        End If
    Else
        MsgBox "엑셀 파일을 선택하여 주세요."
    End If
	Call LayerShowHide(0)
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
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP_BAK">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB_BAK"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:GetRef">ERP계정불러오기</A><!--|<a href="vbscript:GetRef2">엑셀로 로딩시 참고</A>--> </TD>
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
									<TD CLASS="TD6"><script language =javascript src='./js/wb111ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script></TD>
									<TD CLASS="TD5">법인명</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCO_CD" Size=10 tag="14">
										<INPUT TYPE=TEXT NAME="txtCO_NM" Size=20 tag="14">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">신고구분</TD>
									<TD CLASS="TD6"><SELECT NAME="cboREP_TYPE" ALT="신고구분" STYLE="WIDTH: 50%" tag="14X"></SELECT>
									</TD>
									<TD CLASS="TD5">제무제표구분</TD>
									<TD CLASS="TD6"><SELECT NAME="cboBS_PL_FG" ALT="제무제표구분" STYLE="WIDTH: 50%" tag="11X"></SELECT></TD>
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
						<TABLE <%=LR_SPACE_TYPE_60%> CLASS="TD61" BORDER=0>
							<TR>
							     <TD WIDTH="100%" VALIGN=TOP HEIGHT=*>
							     <script language =javascript src='./js/wb111ma1_vspdData_vspdData.js'></script>
							     <script language =javascript src='./js/wb111ma1_vspdData1_vspdData1.js'></script>
							    </TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
   <!-- <TR HEIGHT=20>   
        <TD WIDTH=100%>   
            <TABLE <%=LR_SPACE_TYPE_30%>>   
                <TR>   
                <TD WIDTH=50%>&nbsp;<INPUT TYPE="FILE" NAME="txtFILE_NAME" SIZE="20" tag="21">
                	&nbsp;<BUTTON NAME="bttnExcelLoad"  CLASS="CLSSBTN" ONCLICK="vbscript:BtnExcelLoad()" Flag=1>읽어오기</BUTTON></TD>   
                <TD WIDTH=50%>
                </TD>   
                </TR>   
            </TABLE>   
        </TD>   
    </TR>--> 
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
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
