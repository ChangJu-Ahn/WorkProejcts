<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Accounting
*  2. Function Name        : 
*  3. Program ID           : a5446ma1.asp
*  4. Program Name         : 
*  5. Program Desc         : 
*  6. Comproxy List        :
*  7. Modified date(First) : 2003/05/28
*  8. Modified date(Last)  : 2003/05/28
*  9. Modifier (First)     :
* 10. Modifier (Last)      :
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<Script Language="VBScript">

Option Explicit                                                        '☜: Turn on the Option Explicit option.


'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID 		= "A5446MB1.asp"                              '☆: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Dim	C_ROW_NO
Dim	C_ASST_ACCT_CD
Dim	C_ASST_ACCT_NM
Dim	C_ASST_NO
Dim	C_ASST_NM
Dim	C_DUR_YRS_FG
Dim	C_HIS_SEQ
Dim	C_HIS_FG
Dim	C_HIS_FG_NM
Dim	C_HIS_DT
Dim	C_GL_NO
Dim	C_GL_DT
Dim	C_TEMP_GL_NO
Dim	C_TEMP_GL_DT
Dim	C_DEPT_CD
Dim	C_DEPT_NM
Dim	C_ORG_CHANGE_ID
Dim	C_HIS_INV_QTY_INC
Dim	C_HIS_INV_QTY_DEC
Dim	C_HIS_COST_INC
Dim	C_HIS_COST_DEC
Dim	C_ACCU_DEPR_ACCT_CD
Dim	C_ACCU_DEPR_ACCT_NM
Dim	C_HIS_ACCU_DEPR_INC
Dim	C_HIS_ACCU_DEPR_DEC
Dim	C_REF_NO
Dim	C_HIS_DUR_YRS
Dim	C_HIS_DUR_MNTH
Dim	C_HIS_RES_AMT
Dim	C_DEPR_EXP_ACCT_CD
Dim	C_DEPR_EXP_ACCT_NM
Dim	C_HIS_DESC

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop                                          
Dim IsOpenPop  

Dim lgCookValue

Dim lgStrColorFlag

Dim lgSaveRow 

<% 
   BaseDate     = GetSvrDate                                                                  'Get DB Server Date
%>

Dim EndDate
Dim StartDate

EndDate = UNIConvDateAToB("<%=BaseDate%>",parent.gServerDateFormat,parent.gDateFormat)
StartDate = UNIDateAdd("M", -1, EndDate, Parent.gDateFormat)

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================

Sub initSpreadPosVariables()
	C_ROW_NO			= 1
	C_ASST_ACCT_CD		= 2
	C_ASST_ACCT_NM		= 3
	C_ASST_NO			= 4
	C_ASST_NM			= 5
	C_DUR_YRS_FG		= 6
	C_HIS_FG			= 7
	C_HIS_FG_NM			= 8
	C_HIS_COST_INC		= 9
	C_HIS_COST_DEC		= 10
	C_ACCU_DEPR_ACCT_CD	= 11
	C_ACCU_DEPR_ACCT_NM	= 12
	C_HIS_ACCU_DEPR_INC	= 13
	C_HIS_ACCU_DEPR_DEC	= 14
	C_REF_NO			= 15
	C_HIS_DT			= 16
	C_GL_NO				= 17
	C_GL_DT				= 18
	C_TEMP_GL_NO		= 19
	C_TEMP_GL_DT		= 20
	C_DEPR_EXP_ACCT_CD	= 21
	C_DEPR_EXP_ACCT_NM	= 22
	C_DEPT_CD			= 23
	C_DEPT_NM			= 24
	C_ORG_CHANGE_ID		= 25
	C_HIS_INV_QTY_INC	= 26
	C_HIS_INV_QTY_DEC	= 27
	C_HIS_DUR_YRS		= 28
	C_HIS_DUR_MNTH		= 29
	C_HIS_RES_AMT		= 30
	C_HIS_DESC			= 31
	C_HIS_SEQ			= 32
End Sub

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================	
Sub InitVariables()
    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = Parent.OPMD_CMODE                          'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	frm1.txtFrDt.text = StartDate
	frm1.txtToDt.text = EndDate
    frm1.txtRadio.value = "01"
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
' Para : 1. Currency
'        2. I(Input),Q(Query),P(Print),B(Batch)
'        3. "*" is for Common module
'           "A" is for Accounting
'           "I" is for Inventory
'           ...
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029A("Q", "*","NOCOOKIE","QA") %>                                '☆: 
	<% Call LoadBNumericFormatA("Q", "*", "NOCOOKIE", "QA") %>

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
'	Name : CookiePage()
'	Description : JUMP시 Load화면의 조건부로 Value
'========================================================================================================
Function CookiePage(ByVal Kubun)

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Dim strTemp, arrVal

	Const CookieSplit = 4877

	If Kubun = 0 Then                                              ' Called Area
       strTemp = ReadCookie(CookieSplit)

       If strTemp = "" then Exit Function

       arrVal = Split(strTemp, parent.gRowSep)

       Frm1.txtSchoolCd.Value = ReadCookie ("SchoolCd")
       Frm1.txtGrade.Value   = arrVal(0)
				
       Call MainQuery()

       WriteCookie CookieSplit , ""
	
	ElseIf Kubun = 1 then                                         ' If you want to call
		Call vspdData_Click(Frm1.vspdData.ActiveCol,Frm1.vspdData.ActiveRow)
		WriteCookie CookieSplit , lgCookValue		
		Call PgmJump(BIZ_PGM_JUMP_ID2)
	End IF
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
		
End Function


'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()
	With frm1.vspdData
		 ggoSpread.Source = frm1.vspdData
         ggoSpread.Spreadinit "V20030626",, parent.gAllowDragDropSpread
		.ReDraw = false
		.MaxCols   = C_HIS_SEQ + 1                                                  ' ☜:☜: Add 1 to Maxcols		

		.Col = .MaxCols
		.ColHidden = True
		
		.MaxRows = 0                                                                  ' ☜: Clear spreadsheet data 
		
		Call GetSpreadColumnPos("A")
								'Col			Header			Width		Grp			IntegeralPart			DeciPointpart												Align	Sep		PZ		Min		Max 
'		ggoSpread.SSSetFloat	C_ROW_NO,			"",				3,		"2",	,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,		1,		True
		ggoSpread.SSSetEdit		C_ROW_NO,			"",				8,		,		,		2
		ggoSpread.SSSetEdit		C_ASST_ACCT_CD,		"자산계정",		9,		,		,		18
		ggoSpread.SSSetEdit		C_ASST_ACCT_NM,		"계정명",		12,		,		,		40
		ggoSpread.SSSetEdit		C_ASST_NO,			"자산번호",		9,		,		,		18
		ggoSpread.SSSetEdit		C_ASST_NM,			"자산명",		12,		,		,		40
		ggoSpread.SSSetEdit		C_DUR_YRS_FG,		"상각구분",		8,		,		,		2

		ggoSpread.SSSetEdit		C_HIS_FG,			"",				5,		,		,		18
		ggoSpread.SSSetEdit		C_HIS_FG_NM,		"구분",			5,		,		,		18

		ggoSpread.SSSetFloat	C_HIS_COST_INC,		"취득액+",		10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_HIS_COST_DEC,		"취득액-",		10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

		ggoSpread.SSSetEdit		C_ACCU_DEPR_ACCT_CD,"상각누계",		9,		,		,		18
		ggoSpread.SSSetEdit		C_ACCU_DEPR_ACCT_NM,"계정명",		12,		,		,		40

		ggoSpread.SSSetFloat	C_HIS_ACCU_DEPR_INC,"상각누계+",	10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_HIS_ACCU_DEPR_DEC,"상각누계-",	10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

		ggoSpread.SSSetEdit		C_REF_NO,			"참조번호",		10,		,		,		100

		ggoSpread.SSSetDate		C_HIS_DT,			"거래일",		10,		2,		Parent.gDateFormat,	-1
		
		ggoSpread.SSSetEdit		C_GL_NO,			"전표번호",		10,		,		,		18
		ggoSpread.SSSetDate		C_GL_DT,			"일자",			10,		2,		Parent.gDateFormat,	-1
		ggoSpread.SSSetEdit		C_TEMP_GL_NO,		"결의전표",		10,		,		,		18
		ggoSpread.SSSetDate		C_TEMP_GL_DT,		"일자",			10,		2,		Parent.gDateFormat,	-1
		
		ggoSpread.SSSetEdit		C_DEPR_EXP_ACCT_CD,	"상각계정",		9,		,		,		18
		ggoSpread.SSSetEdit		C_DEPR_EXP_ACCT_NM,	"계정명",		12,		,		,		40

		ggoSpread.SSSetEdit		C_DEPT_CD,			"부서",			8,		,		,		10
		ggoSpread.SSSetEdit		C_DEPT_NM,			"부서명",		12,		,		,		40
		ggoSpread.SSSetEdit		C_ORG_CHANGE_ID,	"조직변경ID",	8,		,		,		5
		
		ggoSpread.SSSetFloat	C_HIS_INV_QTY_INC,	"수량+",		6,		Parent.ggqtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_HIS_INV_QTY_DEC,	"수량-",		6,		Parent.ggqtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

'		ggoSpread.SSSetFloat	C_HIS_DUR_YRS,		"내용연수",		4,		"6",	ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,		1,		True
		ggoSpread.SSSetEdit		C_HIS_DUR_YRS,		"내용연수",		8,		,		,		5
		
'		ggoSpread.SSSetFloat	C_HIS_DUR_MNTH,		"내용월수",		4,		"7",	ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,		1,		True
		ggoSpread.SSSetEdit		C_HIS_DUR_MNTH,		"내용월수",		8,		,		,		5
		
		ggoSpread.SSSetFloat	C_HIS_RES_AMT,		"잔존가액",		10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

		ggoSpread.SSSetEdit		C_HIS_DESC,			"비고",			20,		,		,		100

		ggoSpread.SSSetEdit		C_HIS_SEQ,			"순서",			5,		,		,		4
'		ggoSpread.SSSetFloat	C_HIS_SEQ,			"순서",			4,		"6",	ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,		1,		True

		Call ggoSpread.SSSetColHidden(C_ROW_NO, C_ROW_NO, True)
		Call ggoSpread.SSSetColHidden(C_DUR_YRS_FG, C_DUR_YRS_FG, True)
		Call ggoSpread.SSSetColHidden(C_HIS_FG, C_HIS_FG, True)
		Call ggoSpread.SSSetColHidden(C_ORG_CHANGE_ID, C_ORG_CHANGE_ID, True)
'--------------------------       
		.ReDraw = true
				
		Call SetSpreadLock
			    
	End With
    
End Sub


'========================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock()
    With frm1
		.vspdData.ReDraw = False
		'------ Developer Coding part (Start ) -------------------------------------------------------------- 
		ggoSpread.SpreadLockWithOddEvenRowColor()
		'------ Developer Coding part (End   ) -------------------------------------------------------------- 
		.vspdData.ReDraw = True
	End With
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_ROW_NO			= iCurColumnPos(1)
			C_ASST_ACCT_CD		= iCurColumnPos(2)
			C_ASST_ACCT_NM		= iCurColumnPos(3)
			C_ASST_NO			= iCurColumnPos(4)
			C_ASST_NM			= iCurColumnPos(5)
			C_DUR_YRS_FG		= iCurColumnPos(6)
			C_HIS_FG			= iCurColumnPos(7)
			C_HIS_FG_NM			= iCurColumnPos(8)
			C_HIS_COST_INC		= iCurColumnPos(9)
			C_HIS_COST_DEC		= iCurColumnPos(10)
			C_ACCU_DEPR_ACCT_CD	= iCurColumnPos(11)
			C_ACCU_DEPR_ACCT_NM	= iCurColumnPos(12)
			C_HIS_ACCU_DEPR_INC	= iCurColumnPos(13)
			C_HIS_ACCU_DEPR_DEC	= iCurColumnPos(14)
			C_REF_NO			= iCurColumnPos(15)
			C_HIS_DT			= iCurColumnPos(16)
			C_GL_NO				= iCurColumnPos(17)
			C_GL_DT				= iCurColumnPos(18)
			C_TEMP_GL_NO		= iCurColumnPos(19)
			C_TEMP_GL_DT		= iCurColumnPos(20)
			C_DEPR_EXP_ACCT_CD	= iCurColumnPos(21)
			C_DEPR_EXP_ACCT_NM	= iCurColumnPos(22)
			C_DEPT_CD			= iCurColumnPos(23)    
			C_DEPT_NM			= iCurColumnPos(24)
			C_ORG_CHANGE_ID		= iCurColumnPos(25)
			C_HIS_INV_QTY_INC	= iCurColumnPos(26)
			C_HIS_INV_QTY_DEC	= iCurColumnPos(27)
			C_HIS_DUR_YRS		= iCurColumnPos(28)
			C_HIS_DUR_MNTH		= iCurColumnPos(29)
			C_HIS_RES_AMT		= iCurColumnPos(30)
			C_HIS_DESC			= iCurColumnPos(31)
			C_HIS_SEQ			= iCurColumnPos(32)
    End Select    
End Sub

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()

   
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 


	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset ComboBox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
End Sub

'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029				                                           '⊙: Load table , B_numeric_format

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")

	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolBar("1100000000011111")										

	Call InitComboBox
    Call CookiePage(0)

    Frm1.fpDateTime1.focus
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub
'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
 
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncQuery = False                                                              '⊙: Processing is NG
    
    Call ggoOper.ClearField(Document, "2")									      '⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.ClearSpreadData()
    
    Call InitVariables 														      '⊙: Initializes local global variables
    
    If Not chkField(Document, "1") Then								              '⊙: This function check indispensable field
       Exit Function
    End If

	If CompareDateByFormat(frm1.txtFrDt.text,frm1.txtToDt.text,frm1.txtFrDt.Alt,frm1.txtToDt.Alt, _
        	               "970023",frm1.txtFrDt.UserDefinedFormat,parent.gComDateType, true) = False Then
	   frm1.txtToDt.focus
	   Exit Function
	End If

    If DbQuery = False Then 
       Exit Function
    End If   

    If Err.number = 0 Then
       FncQuery = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
End Function


'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow()
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrint = False                                                              '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                        '☜: Protect system from crashing

    If Err.number = 0 Then
       FncPrint = True                                                            '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 
End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExcel = False                                                              '☜: Processing is NG

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
	Call Parent.FncExport(parent.C_MULTI)
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    If Err.number = 0 Then
       FncExcel = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncFind = False                                                               '☜: Processing is NG

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
	Call Parent.FncFind(parent.C_MULTI, True)
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    If Err.number = 0 Then
       FncFind = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

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

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExit = False                                                               '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    If Err.number = 0 Then
       FncExit = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery() 
	Dim strVal

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    DbQuery = False
    
	Call LayerShowHide(1)
    
    With frm1

		strVal = BIZ_PGM_ID	& "?txtMode="        & Parent.UID_M0001                      '☜: Query
		strVal = strVal		& "&txtMaxRows=" 	 & Frm1.vspdData.MaxRows				'☜: Max fetched data

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
        If lgIntFlgMode  <> parent.OPMD_UMODE Then   ' This means that it is first search
           strVal = strVal & "&txtRadio="			& Trim(.txtRadio.value)
           strVal = strVal & "&txtFrDt="			& UNIConvDateAToB(.txtFrDt.Text, parent.gDateFormat,parent.gServerDateFormat)
           strVal = strVal & "&txtToDt="			& UNIConvDateAToB(.txtToDt.Text, parent.gDateFormat,parent.gServerDateFormat)
           strVal = strVal & "&txtBizAreaCd="		& Trim(.txtBizAreaCd.value)    	
           strVal = strVal & "&txtFrAcctCd="		& Trim(.txtFrAcctCd.value)    	
           strVal = strVal & "&txtToAcctCd="		& Trim(.txtToAcctCd.value)    	
           strVal = strVal & "&txtFrAsstCd="		& Trim(.txtFrAsstCd.value)    	
           strVal = strVal & "&txtToAsstCd="		& Trim(.txtToAsstCd.value)    	
        Else
           strVal = strVal & "?txtRadio="			& Trim(.htxtRadio.value)
           strVal = strVal & "&txtFrDt="			& Trim(.htxtFrDt.value)
           strVal = strVal & "&txtToDt="			& Trim(.htxtToDt.value)
           strVal = strVal & "&txtBizAreaCd="		& Trim(.htxtBizAreaCd.value)    	
           strVal = strVal & "&txtFrAcctCd="		& Trim(.htxtFrAcctCd.value)    	
           strVal = strVal & "&txtToAcctCd="		& Trim(.htxtToAcctCd.value)    	
           strVal = strVal & "&txtFrAsstCd="		& Trim(.htxtFrAsstCd.value)    	
           strVal = strVal & "&txtToAsstCd="		& Trim(.htxtToAsstCd.value)    	
        End If   
           
    '--------- Developer Coding Part (End) ------------------------------------------------------------

        Call RunMyBizASP(MyBizASP, strVal)	                                         '☜: 비지니스 ASP 를 가동 

    End With

    If Err.number = 0 Then
       DbQuery = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()												

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	lgBlnFlgChgValue = False
    lgIntFlgMode     = parent.OPMD_UMODE												 '⊙: Indicates that current mode is Update mode
    lgSaveRow        = 1

    Set gActiveElement = document.ActiveElement   

	Call SetQuerySpreadColor

End Function

'========================================================================================================
'	Description : 스프레트시트의 특정 컬럼의 배경색상을 변경 
'========================================================================================================
Sub SetQuerySpreadColor()

	Dim iArrColor1, iArrColor2
	Dim iLoopCnt
	Dim Spread
	
	Set Spread = frm1.vspdData
	
	iArrColor1 = Split(lgStrColorFlag,Parent.gRowSep)
	
	For iLoopCnt=0 to ubound(iArrColor1,1) - 1
		iArrColor2 = Split(iArrColor1(iLoopCnt),Parent.gColSep)

		Spread.Col = -1
		Spread.Row =  iArrColor2(0)
		
		Select Case iArrColor2(1)
			Case "1"
				Spread.BackColor = RGB(204,255,153) '연두 
			Case "2"
				Spread.BackColor = RGB(176,234,244) '하늘색 
			Case "3"
				Spread.BackColor = RGB(224,206,244) '연보라 
			Case "4"  
				Spread.BackColor = RGB(251,226,153) '연주황 
			Case "5" 
				Spread.BackColor = RGB(255,255,153) '연노랑 
		End Select
	Next

End Sub
'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'========================================================================================================
'	Name : OpenBizAreaPopUp()
'	Description : OpenBizAreaPopUp PopUp
'========================================================================================================
Function OpenBizAreaPopUp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "사업장팝업"                             ' Popup Name
	arrParam(1) = "B_BIZ_AREA"                                        ' Table Name
	arrParam(2) = Frm1.txtBizAreaCd.value                         ' Code Condition
	arrParam(3) = ""                                              ' Name Cindition
	arrParam(4) = ""                                              ' Where Condition
	arrParam(5) = "사업장코드"
	
    arrField(0) = "BIZ_AREA_CD"                                      ' Field명(0)
    arrField(1) = "BIZ_AREA_NM"                                      ' Field명(1)
    
    arrHeader(0) = "사업장코드"	                              ' Header명(0)
    arrHeader(1) = "사업장명"                               ' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBizAreaCd.focus
		Exit Function
	Else
		Call SetBizArea(arrRet)
	End If	


End Function

Sub SetBizArea(Byval arrRet)
   
	With Frm1
	  .txtBizAreaCd.value = arrRet(0)
	  .txtBizAreaNm.value = arrRet(1)
	End With
   
End Sub

'========================================================================================================
'	Name : OpenAcctCd()
'	Description : Plant PopUp
'========================================================================================================
Function OpenAcctCd(byval strText, byval iwhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "자산계정팝업"                             ' Popup Name
	arrParam(1) = "A_ASSET_ACCT A, A_ACCT B"                                        ' Table Name
	arrParam(2) = Trim(strText)                         ' Code Condition
	arrParam(3) = ""                                              ' Name Cindition
	arrParam(4) = "A.ACCT_CD = B.ACCT_CD"                                              ' Where Condition
	arrParam(5) = "자산계정"
	
    arrField(0) = "A.ACCT_CD"                                      ' Field명(0)
    arrField(1) = "B.ACCT_NM"                                      ' Field명(1)
    
    arrHeader(0) = "자산계정"	                              ' Header명(0)
    arrHeader(1) = "자산계정명"                               ' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If Cint(iwhere) = 1 then
			frm1.txtFrAcctCd.focus
		else
			frm1.txtToAcctCd.focus
		end if
		Exit Function
	Else
		Call SetAcctCd(arrRet, iwhere)
	End If	


End Function

Sub SetAcctCd(ByVal arrRet, ByVal iwhere)
	With Frm1
		If Cint(iwhere) = 1 then
			.txtFrAcctCd.value = arrRet(0)
			.txtFrAcctNm.value = arrRet(1)
		else
			.txtToAcctCd.value = arrRet(0)
			.txtToAcctNm.value = arrRet(1)
			call txtToAcctCd_onChange()
		end if   
	End With
   
End Sub
'========================================================================================================
'	Name : OpenAsstCd()
'	Description : Plant PopUp
'========================================================================================================
Function OpenAsstCd(byval strText, byval iwhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "자산번호팝업"                             ' Popup Name
	arrParam(1) = "A_ASSET_MASTER"                                        ' Table Name
	arrParam(2) = Trim(strText)                         ' Code Condition
	arrParam(3) = ""                                              ' Name Cindition
	arrParam(4) = ""                                              ' Where Condition
	arrParam(5) = "자산번호"
	
    arrField(0) = "ASST_NO"                                      ' Field명(0)
    arrField(1) = "ASST_NM"                                      ' Field명(1)
    
    arrHeader(0) = "자산번호"	                              ' Header명(0)
    arrHeader(1) = "자산명"                               ' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If Cint(iwhere) = 1 then
			frm1.txtFrAsstCd.focus
		else
			frm1.txtToAsstCd.focus
		end if
		Exit Function
	Else
		Call SetAsstCd(arrRet, iwhere)
	End If	


End Function

Sub SetAsstCd(ByVal arrRet, ByVal iwhere)
	With Frm1
		If Cint(iwhere) = 1 then
			.txtFrAsstCd.value = arrRet(0)
			.txtFrAsstNm.value = arrRet(1)
		else
			.txtToAsstCd.value = arrRet(0)
			.txtToAsstNm.value = arrRet(1)
			call txtToAsstCd_onChange()
		end if   
	End With
   
End Sub

'==================================================================================
' Name : PopZAdoConfigGrid
' Desc :
'==================================================================================
Sub PopZAdoConfigGrid()

  If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
     Exit Sub
  End If

  Call OpenOrderBy("A")
  
End Sub



'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub


'========================================================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'========================================================================================================
Sub OpenOrderBy(ByVal pvPsdNo)
	Dim arrRet
	
	If lgIsOpenPop = True Then Exit Sub
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData(pvPsdNo),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then         ' Means that nothing is happened!!!
	   Exit Sub
	Else
	   Call ggoSpread.SaveXMLData(pvPsdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Sub


'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub


'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : Occurs when the user clicks the left mouse button while the pointer is in a cell
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub    

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)

End Sub


'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
	If Frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(Frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then								
           Call DisableToolBar(parent.TBC_QUERY)
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If
    
End Sub

'========================================================================================================
'   Event Name : fpdtFromEnterDt
'   Event Desc : Date OCX Double Click
'========================================================================================================
Sub txtFrDt_DblClick(Button)
	If Button = 1 Then
       Frm1.txtFrDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtFrDt.Focus
	End If
End Sub
'========================================================================================================
'   Event Name : fpdtToEnterDt
'   Event Desc : Date OCX Double Click
'========================================================================================================
Sub txtToDt_DblClick(Button)
	If Button = 1 Then
       Frm1.txtToDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtToDt.Focus
	End If
End Sub

'========================================================================================================
'   Event Name : fpdtFromEnterDt_KeyPress()
'   Event Desc : 
'========================================================================================================
Sub txtFrDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

'========================================================================================================
'   Event Name : fpdtToEnterDt_KeyPress()
'   Event Desc : 
'========================================================================================================
Sub txtToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub


Function Radio1_onChange()
	if frm1.txtRadio.value <> "01" then
		TitleDate.innerHTML = "전표일자"
		frm1.txtRadio.value = "01"
	end if
End Function

Function Radio2_onChange()
	if frm1.txtRadio.value <> "02" then         
		TitleDate.innerHTML = "거래일자"
		frm1.txtRadio.value = "02"
	end if
End Function

'========================================================================================================
'========================================================================================================
'   Event Name : txtBizAreaCd_onChange
'   Event Desc : 
'========================================================================================================
Sub txtBizAreaCd_onChange()
	Dim IntRetCD
	Dim arrVal

	If frm1.txtBizAreaCd.value = "" Then Exit Sub

	If CommonQueryRs("BIZ_AREA_NM", "B_BIZ_AREA ", " BIZ_AREA_CD=  " & FilterVar(frm1.txtBizAreaCd.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		frm1.txtBizAreaNm.value= Trim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("124200","X","X","X")  	
		frm1.txtBizAreaCd.focus
	End If
End Sub

'========================================================================================================
'   Event Name : txtAcctCd_onChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub txtFrAcctCd_onChange()
	Dim IntRetCD
	Dim arrVal
	
	If CommonQueryRs("ACCT_NM", "A_ACCT ", " ACCT_CD=  " & FilterVar(frm1.txtFrAcctCd.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		frm1.txtFrAcctNm.value= Trim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("117100","X","X","X")
		frm1.txtFrAcctCd.value = ""
		frm1.txtFrAcctNm.value = ""
		frm1.txtFrAcctCd.focus
	End If
End Sub
'========================================================================================================
'   Event Name : txtAcctCd_onChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub txtToAcctCd_onChange()
	Dim IntRetCD
	Dim arrVal
	
	If frm1.txtToAcctCd.value < frm1.txtFrAcctCd.value then
		IntRetCD = DisplayMsgBox("970023","X",frm1.txtToAcctCd.alt, frm1.txtFrAcctCd.alt)
		frm1.txtToAcctCd.value = ""
		frm1.txtToAcctNm.value = ""
		exit sub
	end if 
	
	If CommonQueryRs("ACCT_NM", "A_ACCT ", " ACCT_CD=  " & FilterVar(frm1.txtToAcctCd.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		frm1.txtToAcctNm.value= Trim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("117100","X","X","X")
		frm1.txtToAcctCd.value = ""
		frm1.txtToAcctNm.value = ""
		frm1.txtToAcctCd.focus
	End If
End Sub

'========================================================================================================
'   Event Name : txtAcctCd_onChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub txtFrAsstCd_onChange()
	Dim IntRetCD
	Dim arrVal
	
	If CommonQueryRs("ASST_NM", "A_ASSET_MASTER ", " ASST_NO=  " & FilterVar(frm1.txtFrAsstCd.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		frm1.txtFrAsstNm.value= Trim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("117400","X","X","X")
		frm1.txtFrAsstCd.value = ""
		frm1.txtFrAsstNm.value = ""
		frm1.txtFrAsstCd.focus
	End If
End Sub
'========================================================================================================
'   Event Name : txtAsstCd_onChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub txtToAsstCd_onChange()
	Dim IntRetCD
	Dim arrVal
	
	If frm1.txtToAsstCd.value < frm1.txtFrAsstCd.value then
		IntRetCD = DisplayMsgBox("970023","X",frm1.txtToAsstCd.alt, frm1.txtFrAsstCd.alt)
		frm1.txtToAsstCd.value = ""
		frm1.txtToAsstNm.value = ""
		exit sub
	end if 
	
	If CommonQueryRs("ASST_NM", "A_ASSET_MASTER ", " ASST_NO=  " & FilterVar(frm1.txtToAsstCd.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		frm1.txtToAsstNm.value= Trim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("117400","X","X","X")
		frm1.txtToAsstCd.value = ""
		frm1.txtToAsstNm.value = ""
		frm1.txtToAsstCd.focus
	End If
End Sub



</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>고정자산확인작업</font></td>
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
						<FIELDSET>
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>조회구분</TD>
									<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio1 ID=Rb_WK1 Checked onclick=radio1_onchange()><LABEL FOR=Rb_WK1>전표일자</LABEL>&nbsp;<INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO1 ID=Rb_WK2 onclick=radio2_onchange()><LABEL FOR=Rb_WK2>거래일자</LABEL></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" ID = "TitleDate">전표일자</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/a5446ma1_fpDateTime1_txtFrDt.js'></script>&nbsp;~&nbsp;
												           <script language =javascript src='./js/a5446ma1_fpDateTime2_txtToDt.js'></script></TD>
									<TD CLASS="TD5" NOWRAP>자산관리사업장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="사업장코드"><IMG SRC="../../image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBizAreaPopUp()"> <INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=18 tag="14X" ALT="사업장명"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>고정자산계정</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtFrAcctCd" SIZE=12 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="자산계정코드(From)"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFrAcctCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenAcctCd(frm1.txtFrAcctCd.value, 1)"> <INPUT TYPE=TEXT NAME="txtFrAcctNm" SIZE=25 tag="14">&nbsp;~</TD>
									<TD CLASS="TD5" NOWRAP>자산번호</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtFrAsstCd" SIZE=12 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="자산번호(From)"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAsstCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenAsstCd(frm1.txtFrAsstCd.value,1)"> <INPUT TYPE=TEXT NAME="txtFrAsstNm" SIZE=25 tag="14">&nbsp;~</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtToAcctCd" SIZE=12 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="자산계정코드(To)"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnToAcctCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenAcctCd(frm1.txtToAcctCd.value, 2)"> <INPUT TYPE=TEXT NAME="txtToAcctNm" SIZE=25 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtToAsstCd" SIZE=12 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="자산번호(To)"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnToAsstCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenAsstCd(frm1.txtToAsstCd.value,2)"> <INPUT TYPE=TEXT NAME="txtToAsstNm" SIZE=25 tag="14"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/a5446ma1_vspdData_vspdData.js'></script>
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
    <TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH="*" ALIGN="RIGHT"></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtRadio"		tag="34" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="htxtRadio"		tag="34" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="htxtFrDt"		tag="34" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="htxtToDt"		tag="34" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="htxtBizAreaCd"	tag="34" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="htxtFrAcctCd"	tag="34" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="htxtToAcctCd"	tag="34" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="htxtFrAsstNo"	tag="34" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="htxtToAsstNo"	tag="34" TABINDEX = "-1">
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
