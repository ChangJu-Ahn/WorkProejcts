<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : AM
'*  2. Function Name        :
'*  3. Program ID           : A5965MA1
'*  4. Program Name         : 월차전표 내역 조회 
'*  5. Program Desc         : Single-Multi Sample
'*  6. Component List       :
'*  7. Modified date(First) : 2002/01/15
'*  8. Modified date(Last)  : 2002/01/16
'*  9. Modifier (First)     : song sang min
'* 10. Modifier (Last)      : song sang min
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>



<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">


<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js">   </SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance


'========================================================================================================

Const BIZ_PGM_ID      = "a5965mb1.asp"						           '☆: Biz Logic ASP Name


'--------------------------------------------------------------------------------------------------------

Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2

Dim C_DEPT_CD	        '부서코드 
Dim C_DEPT_CD_NM		'부서명 
Dim C_DR_ACCT_CD		'차변 계정명 
Dim C_DR_ACCT   	    '차변 계정명 
Dim C_DR_AMOUNT 		'차변 금액 
Dim C_CR_AMOUNT 		'대변 금액 
Dim C_TEMP_GL_NO		'결의전표번호 
Dim C_GL_NO				'전표번호 

Const COOKIE_SPLIT       = 4877	                                      'Cookie Split String

'========================================================================================================
'=                       4.3 Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop
Dim IsOpenPop
Dim StartDate
<%
StartDate	= GetSvrDate                                               'Get Server DB Date
%>

'========================================================================================================
Sub initSpreadPosVariables()
	 C_DEPT_CD 			= 1		'부서코드 
	 C_DEPT_CD_NM   	= 2		'부서명 
	 C_DR_ACCT_CD		= 3		'차변 계정명 
	 C_DR_ACCT			= 4		'차변 계정명 
	 C_DR_AMOUNT		= 5		'차변 금액 
	 C_CR_AMOUNT   		= 6		'대변 금액 
	 C_TEMP_GL_NO		= 7		'결의전표번호 
	 C_GL_NO			= 8		'전표번호 
End Sub
'========================================================================================================
' Name : InitVariables()
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
   	lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
   	lgStrPrevKeyIndex = ""                                      '⊙: initializes Previous Key Index
   	lgSortKey         = 1                                       '⊙: initializes sort direction

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	lgStrPrevKey = ""                                           'initializes Previous Key
   	lgLngCurRows = 0                                            'initializes Deleted Rows Count
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
Sub SetDefaultVal()


	Dim StartDate
	Dim strYear, strMonth, strDay

	StartDate	= "<%=StartDate%>"                           'Get Server DB Date

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	Call ExtractDateFrom(StartDate,Parent.gServerDateFormat,Parent.gServerDateType,strYear,strMonth,strDay)
	frm1.fpdtWk_yymm.Text	= UniConvDateAToB(StartDate ,parent.gServerDateFormat,parent.gDateFormat) 
	Call ggoOper.FormatDate(frm1.fpdtWk_yymm, Parent.gDateFormat,2)
	frm1.fpdtWk_yymm.focus
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub


'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "A", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("Q", "A", "NOCOOKIE", "MA") %>
End Sub


'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value
'========================================================================================================
Sub CookiePage(Kubun)
   '------ Developer Coding part (Start ) --------------------------------------------------------------
   '------ Developer Coding part (End   ) --------------------------------------------------------------
End Sub

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
    Dim strYYYYMM
    Dim strYear,strMonth,strDay
    '------ Developer Coding part (Start ) --------------------------------------------------------------
    Call ExtractDateFrom(frm1.fpdtWk_yymm.Text,frm1.fpdtWk_yymm.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
    strYYYYMM =strYear&strMonth
  
    lgKeyStream = strYYYYMM   & parent.gColSep       '날짜 
    lgKeyStream = lgKeyStream & Trim(Frm1.txtCurrencyCode.Value) & parent.gColSep '계정그룹코드 
    lgkeyStream = lgkeyStream & Trim(frm1.txtReg.value) & parent.gColSep '날짜 
  
   '------ Developer Coding part (End) --------------------------------------------------------------
End Sub

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr
    Dim iNameArr
    Dim iDx
 
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 컬럼을 클릭할 경우 발생하는 콤보 박스 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021103",,parent.gAllowDragDropSpread  
		.ReDraw = false
		
       .MaxCols   = C_GL_NO + 1                                                 ' ☜:☜: Add 1 to Maxcols
	   .Col       = .MaxCols                                                        ' ☜:☜: Hide maxcols
       .ColHidden = True                                                            ' ☜:☜:
       '.MaxRows = 0
		ggoSpread.Source = frm1.vspdData
		ggospread.ClearSpreadData		'Buffer Clear

		Call AppendNumberPlace("6","15","2")
		
		Call GetSpreadColumnPos("A")
		
        ggoSpread.SSSetEdit     C_DEPT_CD        ,     "부서 코드"  ,13,,,5,2
        ggoSpread.SSSetEdit     C_DEPT_CD_NM     ,     "부서명"   	 ,21,,,50,2
        ggoSpread.SSSetEdit   	C_DR_ACCT_CD		 , "계정코드"   ,13,,,20,2
        ggoSpread.SSSetEdit   	C_DR_ACCT		 ,     "계정명"   	 ,20,,,50,2
        ggoSpread.SSSetFloat    C_DR_AMOUNT      ,     "차변 금액"   ,20,2,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
        ggoSpread.SSSetFloat    C_CR_AMOUNT      ,     "대변 금액"   ,20,2,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit   	C_TEMP_GL_NO	 ,		"결의전표번호"   ,16,,,20,2
        ggoSpread.SSSetEdit   	C_GL_NO		 	 ,     "회계전표번호"  	 ,16,,,20,2

		'Call ggoSpread.MakePairsColumn(C_AcctCd,C_AcctCdPopup,"1")      
		'Call ggoSpread.SSSetColHidden(C_DeprMthd,C_DeprMthd,True)
		
	   .ReDraw = true

       Call SetSpreadLock
    End With

End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
    .vspdData.ReDraw = False
	ggoSpread.SpreadLockWithOddEvenRowColor()
	ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    .vspdData.ReDraw = True
    End With
End Sub
'======================================================================================================
'	Name : OpenCode()
'	Description : 사업장 
'=======================================================================================================%>
Function OpenCode()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "사업장 팝업"		            <%' 팝업 명칭 %>
	arrParam(1) = "B_BIZ_AREA "	                <%' TABLE 명칭 %>
	arrParam(2) = frm1.txtCurrencyCode.value            <%' Code Condition%>
	arrParam(3) = ""   		                    <%' Name Cindition%>
	arrParam(4) = ""        <%' Where Condition%>
	arrParam(5) = "사업장"

   	arrField(0) = "BIZ_AREA_CD"	     			    <%' Field명(1)%>
    arrField(1) = "BIZ_AREA_NM"					    <%' Field명(0)%>


    arrHeader(0) = "사업장"			    <%' Header명(0)%>
    arrHeader(1) = "사업장명"				<%' Header명(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtCurrencyCode.focus
		Exit Function
	Else
		Call SetCode(arrRet)
	End If

End Function
'======================================================================================================
'	Name : SetCode()
'	Description : 사업장코드 Popup에서 Return되는 값 setting
'=======================================================================================================%>
Function SetCode(Byval arrRet)
	With frm1
		.txtCurrencyCode.focus
		.txtCurrencyCode.value = arrRet(0)
		.txtCurrency.value = arrRet(1)
	End With
End Function
'======================================================================================================
'	Name : OpenCodeCon()
'	Description : 월차 구분 코드 
'=======================================================================================================%>
Function OpenCodeCon()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "월차구분"		           <%' 팝업 명칭 %>
	arrParam(1) = "B_MINOR A, B_MAJOR B, A_MONTHLY_BASE C "	   <%' TABLE 명칭 %>
	arrParam(2) = frm1.txtReg.value   <%' Code Condition%>
	arrParam(3) = ""   		                   <%' Name Cindition%>
	arrParam(4) = "A.MINOR_TYPE = " & FilterVar("S", "''", "S") & "  AND A.MAJOR_CD = B.MAJOR_CD AND B.MAJOR_CD = " & FilterVar("A1029", "''", "S") & "  AND a.minor_cd = c.reg_cd and a.minor_cd <> " & FilterVar("10", "''", "S") & "  AND c.USE_YN = " & FilterVar("Y", "''", "S") & "  "        <%' Where Condition%>
	arrParam(5) = "월차구분"

   	arrField(0) = "A.MINOR_CD"	     			<%' Field명(1)%>
    arrField(1) = "A.MINOR_NM"					<%' Field명(0)%>


    arrHeader(0) = "월차구분"			    <%' Header명(0)%>
    arrHeader(1) = "월차구분명"				<%' Header명(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtReg.focus
		Exit Function
	Else
		Call SetReg(arrRet)
	End If

End Function
'======================================================================================================
'	Name : SetCode()
'	Description : 월차구분코드 Popup에서 Return되는 값 setting
'=======================================================================================================%>
Function SetReg(Byval arrRet)
	With frm1
		.txtReg.focus
		.txtReg.value = arrRet(0)
		.txtRegnm.value = arrRet(1)
	End With
End Function

'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStarRow, ByVal pvEndRow)
       
    With frm1
    .vspdData.ReDraw = False
    
    .vspdData.ReDraw = True
    End With
End Sub


Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = frm1.vspdData
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_DEPT_CD 			= iCurColumnPos(1)
			C_DEPT_CD_NM   		= iCurColumnPos(2)
			C_DR_ACCT_CD		= iCurColumnPos(3)
			C_DR_ACCT			= iCurColumnPos(4)
			C_DR_AMOUNT			= iCurColumnPos(5)
			C_CR_AMOUNT   		= iCurColumnPos(6)
			C_TEMP_GL_NO		= iCurColumnPos(7)
			C_GL_NO		   		= iCurColumnPos(8)
	End Select
End Sub

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0
              Exit For
           End If

       Next

    End If
End Sub


Sub PopSaveSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet()
	Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub


'========================================================================================================
Sub Form_Load()
    Err.Clear 
	Call LoadInfTB19029
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")
    Call InitSpreadSheet
    Call InitVariables
    Call SetDefaultVal
    Call SetToolbar("1100000000001111")
    Call CookiePage(0)
    '------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD

    FncQuery = False															  '☜: Processing is NG
    Err.Clear                                                                     '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")			          '☜: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "2")										  '⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
  ' Call SetDefaultVal
    Call InitVariables															  '⊙: Initializes local global variables

    If Not chkField(Document, "1") Then									          '⊙: This function check indispensable field
       Exit Function
    End If
  
	'------ Developer Coding part (Start ) --------------------------------------------------------------
 
    Call MakeKeyStream("X")
 
   '------ Developer Coding part (End )   --------------------------------------------------------------
    If DbQuery = False Then                                                       '☜: Query db data
       Exit Function
    End If
    
	
    Set gActiveElement = document.ActiveElement
    FncQuery = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD

    FncNew = False																  '☜: Processing is NG
    Err.Clear                                                                     '☜: Clear err status

    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x")					  '☜: Data is changed.  Do you want to make it new?
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If

    Call ggoOper.ClearField(Document, "1")                                        '☜: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                        '☜: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
    Call ggoOper.LockField(Document , "N")                                        '☜: Lock  Field
	'------ Developer Coding part (Start ) --------------------------------------------------------------

	
    Call SetToolbar("1100000000001111")									<%'버튼 툴바 제어 %>
    
    Call SetDefaultVal
    Call InitVariables

	'------ Developer Coding part (End )   --------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
    FncNew = True															      '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD

    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                           'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                 '☜: Please do Display first.
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		                 '☜: Do you want to delete?
	If IntRetCD = vbNo Then
		Exit Function
	End If

	'------ Developer Coding part (Start ) --------------------------------------------------------------

	'------ Developer Coding part (End )   --------------------------------------------------------------
    If DbDelete = False Then                                                     '☜: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement
    FncDelete = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD
    
    FncSave = False                                                              '☜: Processing is NG

    Err.Clear                                                                    '☜: Clear err status

    
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")                           '⊙: No data changed!!
        Exit Function
    End If

    If Not chkField(Document, "2") Then
       Exit Function
    End If

	ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
       Exit Function
    End If

	'------ Developer Coding part (Start ) --------------------------------------------------------------

   '  Call MakeKeyStream("X")
    
	'------ Developer Coding part (End )   --------------------------------------------------------------
    If DbSave = False Then                                                       '☜: Query db data
       Call LayerShowHide(0)
       Exit Function
    End If
    Set gActiveElement = document.ActiveElement
    FncSave = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If

	With Frm1

		If .vspdData.ActiveRow > 0 Then
			.vspdData.ReDraw = False

			ggoSpread.Source = frm1.vspdData
			ggoSpread.CopyRow
			SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	' Clear key field
	'----------------------------------------------------------------------------------------------------

	'------ Developer Coding part (End )   --------------------------------------------------------------

			.vspdData.ReDraw = True
			.vspdData.focus
		End If
	End With
    Set gActiveElement = document.ActiveElement
    FncCopy = True                                                                '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel()
    FncCancel = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
    FncCancel = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow()
    Err.Clear                                                                    '☜: Clear err status

	Dim imRow
	FncInsertRow = False

	if IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	else
		imRow = AskSpdSheetAddRowcount()

	If ImRow="" then
		Exit Function
	End If
	End If

	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, frm1.vspdData.ActiveRow
       .vspdData.ReDraw = True
    End With
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
    FncInsertRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows

    FncDeleteRow = False														 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If Frm1.vspdData.MaxRows < 1 then
       Exit function
	End if
    With Frm1.vspdData
    	.focus
    	ggoSpread.Source = frm1.vspdData
    	lDelRows = ggoSpread.DeleteRow
    End With
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
    FncDeleteRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()
    FncPrint = False	                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
	Call Parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True	                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev()
    Dim strVal
    Dim IntRetCD
    FncPrev = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                           '☜: Please do Display first.
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    Call ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear

    Call SetDefaultVal
    Call InitVariables													         '⊙: Initializes local global variables

    if LayerShowHide(1) = false then
	    Exit Function
	end if

    Call MakeKeyStream("X")

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
    strVal = strVal     & "&txtPrevNext="        & "P"	                         '☆: Direction


	'------ Developer Coding part (Start)  --------------------------------------------------------------

	'------ Developer Coding part (End )   --------------------------------------------------------------

	Call RunMyBizASP(MyBizASP, strVal)										     '☜: Run Biz
    Set gActiveElement = document.ActiveElement
    FncPrev = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext()
    Dim strVal
    Dim IntRetCD

    FncNext = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                           '☜: Please do Display first.
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    Call ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear

    Call SetDefaultVal
    Call InitVariables														     '⊙: Initializes local global variables

    if LayerShowHide(1) = false then
	    Exit Function
	end if

    Call MakeKeyStream("X")

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
    strVal = strVal     & "&txtPrevNext="        & "N"	                         '☆: Direction


	'------ Developer Coding part (Start )   --------------------------------------------------------------
	'------ Developer Coding part (End   )   --------------------------------------------------------------

	Call RunMyBizASP(MyBizASP, strVal)										     '☜: Run Biz
    Set gActiveElement = document.ActiveElement
    FncNext = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel()
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncExport(parent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind()
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncFind(parent.C_MULTI, True)

    FncFind = True                                                               '☜: Processing is OK
End Function


'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub


'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")			         '⊙: Data is changed.  Do you want to exit?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True                                                               '☜: Processing is OK
End Function


'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery()
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

    if LayerShowHide(1) = false then
	    Exit Function
	end if                                                       '☜: Show Processing Message

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
	'------ Developer Coding part (Start)  --------------------------------------------------------------

	'------ Developer Coding part (End )   --------------------------------------------------------------
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is OK

    Set gActiveElement = document.ActiveElement
End Function
'========================================================================================================
' Name : DbSave
' Desc : This sub is called by FncSave
'========================================================================================================
Function DbSave()
    Dim pP21011
    Dim lRow
    Dim lGrpCnt
    Dim retVal
    Dim boolCheck
    Dim lStartRow
    Dim lEndRow
    Dim lRestGrpCnt
    Dim strVal, strDel
    DIm IntRetCD
 
    Err.Clear                                                                    '☜: Clear err status

 
    DbSave = False                                                               '☜: Processing is NG
    if LayerShowHide(1) = false then
	    Exit Function
	end if                                                     '☜: Show Processing Message

	'------ Developer Coding part (Start)  --------------------------------------------------------------
    With frm1
        .txtMode.value        = parent.UID_M0002                                        '☜: Delete
        .txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
    End With

   '------ Developer Coding part (add)  --------------------------------------------------------------

   '------ Developer Coding part (add)  --------------------------------------------------------------
    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1

       For lRow = 1 To .vspdData.MaxRows

           .vspdData.Row = lRow
           .vspdData.Col = 0
           
    
           Select Case .vspdData.Text

               Case ggoSpread.InsertFlag                                      '☜: Update
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.UpdateFlag                                      '☜: Update
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete
       End Select
       Next

	   .txtMaxRows.value     = lGrpCnt-1
	   .txtSpread.value      = strDel & strVal

	End With

	'------ Developer Coding part (End )   --------------------------------------------------------------

	'------ Developer Coding part (End )   --------------------------------------------------------------
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    DbSave = True


    Set gActiveElement = document.ActiveElement
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This sub is called by FncDelete
'========================================================================================================
Function DbDelete()

    Err.Clear                                                                    '☜: Clear err status
    DbDelete = False                                                             '☜: Processing is NG
    'Call LayerShowHide(1)                                                        '☜: Show Processing Message

    DbDelete = True                                                             '☜: Processing is OK
	'Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Sub DbQueryOk()

	lgIntFlgMode      = parent.OPMD_UMODE                                                   '⊙: Indicates that current mode is Create mode
    Call InitData()
	Call ggoOper.LockField(Document, "Q")
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement
End Sub

'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Sub DbSaveOk()
	Call InitVariables
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
    
    Call SetToolbar("1100000000001111")
    FncQuery()
    Set gActiveElement = document.ActiveElement
End Sub


'-----------------------------------------  OpenPopuptempGL()  --------------------------------------------------
'	Name : OpenPopuptempGL()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopuptempGL()
	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName
	Dim IntRetCD
	
	iCalledAspName = AskPRAspName("a5130ra1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function

	With frm1.vspdData
		.Row = .ActiveRow
		.Col = C_TEMP_GL_NO
		arrParam(0) = Trim(.Text)							        '결의전표번호 
	    arrParam(1) = ""											'Reference번호	
	End With


	If arrParam(0) = "" Then
		IntRetCD = DisplayMsgBox("970000","X" , "결의전표", "X") 	
		IsOpenPop = False
		Exit Function
	End If	
	
	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

'-----------------------------------------  OpenPopupGL()  --------------------------------------------------
'	Name : OpenPopupGL()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopupGL()
	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName
	Dim IntRetCD

	iCalledAspName = AskPRAspName("a5120ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	With frm1.vspdData
		.Row = .ActiveRow
		.Col = C_GL_NO
		arrParam(0) = Trim(.Text)							        '회계전표번호 
	    arrParam(1) = ""											'Reference번호	
	End With
	
	If arrParam(0) = "" Then
		IntRetCD = DisplayMsgBox("970000","X" , "회계전표", "X") 	
		IsOpenPop = False
		Exit Function
	End If

	
	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function


'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(Col , Row)

    Dim iDx
    Dim IntRetCD,EFlag
    
    EFlag = False

   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col


	'------ Developer Coding part (End   ) --------------------------------------------------------------

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)


	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

    '------ Developer Coding part (Start ) --------------------------------------------------------------
    '데이터 확인시 틀린데이터에 대해 undo 해준다.
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = 0

    If EFlag And Frm1.vspdData.Text <> ggoSpread.InsertFlag Then
		Call FncCancel()
	End If
	'------ Developer Coding part (End   ) --------------------------------------------------------------
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(Col, Row)
    Call SetPopupMenuItemInf("0000111111")
    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData

    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal Col1, ByVal Col2)		' 7) 컬럼 width 변경 이벤트 핸들러 
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(Col1,Col2)

End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				'8) 컬럼 title 변경 
    Dim iColumnName
 	If Row<=0 then
		Exit Sub
	End If
	If Frm1.vspdData.MaxRows =0 then
		Exit Sub
	End if
End Sub

'======================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 특정 column를 click할때 
'======================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    if Frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
       If lgStrPrevKeyIndex <> "" Then
          lgCurrentSpd = "M"
          Call MakeKeyStream("X")
          Call DisableToolBar(parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End if
       End If
    End if

End Sub


'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================

Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData 
  If Row >= NewRow Then
      Exit Sub
  End If
    End With
End Sub


Sub vspdData_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    Call GetSpreadColumnPos("A")
End Sub


'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
 Dim index
 
 With frm1.vspdData
  .Row = Row

  .Col = Col
  index = .Value
   
  .Col = Col - 1
  .Value = index
 End With
 
End Sub

'========================================================================================================
'   Event Name : cboYesNo_OnChange
'   Event Desc :
'========================================================================================================
Sub cboYesNo_OnChange()
    lgBlnFlgChgValue = True
End Sub

'======================================================================================================
' Name : fpdtWk_yymm_DblClick
' Desc :
'=======================================================================================================

Sub fpdtWk_yymm_DblClick(Button)
	If Button = 1 Then
		frm1.fpdtWk_yymm.Action = 7
		Call SetFocusToDocument("M")
		Frm1.fpdtWk_yymm.Focus
	End If
End Sub

'======================================================================================================
' Name : fpdtWk_yymm_KeyPress
' Desc : Call Mainquery
'=======================================================================================================
Sub fpdtWk_yymm_KeyPress(Key)
    If key = 13 Then
        Call FncQuery
		End If
End Sub
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'########################################################################################################
-->
<BODY SCROLL="no" TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>월차전표내역조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A></TD>
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
									<TD CLASS=TD5 NOWRAP WIDTH=14%>년월</TD>
									<TD CLASS=TD6 NOWRAP WIDTH=86%><script language =javascript src='./js/a5965ma1_fpDateTime3_fpdtWk_yymm.js'></script></TD>
									<TD CLASS=TD5 NOWRAP>사업장</TD>
									<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE=TEXT NAME="txtCurrencyCode" SIZE=10 MAXLENGTH=10 tag="12XXXU"  ALT="사업장" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCode1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCode()">
									<INPUT TYPE=TEXT NAME="txtCurrency" SIZE=22 MAXLENGTH=50 tag="14XXXU"  ALT="사업장명">
									</TD>
									</TR>
							    	<TR>
                                    <TD CLASS=TD5 NOWRAP>월차 구분</TD>
									<TD CLASS=TD6 NOWRAP>
									    <INPUT TYPE=TEXT NAME="txtReg" SIZE=10 MAXLENGTH=30 tag="12XXXU"  ALT="월차 구분"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCode1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCodeCon()">
     								    <INPUT TYPE=TEXT NAME="txtRegnm" SIZE=22 MAXLENGTH=50 tag=14XXXU  ALT="월차 구분명">
									</TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP>
									</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_30%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/a5965ma1_vaSpread_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_30%>>
							<TR>
						   	<TD CLASS=TD5 NOWRAP>차변금액 합계</TD>
						   	<TD CLASS=TD6 NOWRAP> <script language =javascript src='./js/a5965ma1_fpDoubleSingle1_txtdrAmt.js'></script></TD>
						   	<TD CLASS=TD5 NOWRAP>대변금액 합계</TD>
						   	<TD CLASS=TD6 NOWRAP> <script language =javascript src='./js/a5965ma1_fpDoubleSingle2_txtcrAmt.js'></script></TD>
							</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BizSize%>> <IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO  noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows"      TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMode"         TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"    TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"   TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId"  TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"      TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"     TAG="24" TABINDEX="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm" TABINDEX="-1"></IFRAME>
</DIV>
</BODY>
</HTML>

