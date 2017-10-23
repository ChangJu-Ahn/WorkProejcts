
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : 경영손익 
*  2. Function Name        :
*  3. Program ID           : GC006MA1
*  4. Program Name         : GC006MA1
*  5. Program Desc         : 경영손익 Profit Center별 손익현황 조회 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/12/05
*  8. Modified date(Last)  : 2001/12/05
*  9. Modifier (First)     : 권기수 
* 10. Modifier (Last)      : 권기수 
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
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID      = "GC006MB1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_JUMP_ID = "GC008MA1"
'========================================================================================================
'=                       4.2 Constant variables
'========================================================================================================
'--------------------------------------------------------------------------------------------------------
'  Constants for SpreadSheet #1
'--------------------------------------------------------------------------------------------------------

Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2

Dim C_PROFIT_NM       
Dim C_PROFIT_AMT      
Dim C_PROFIT_PCNT     
Dim C_OLD_PROFIT_AMT  
Dim C_OLD_PROFIT_PCNT 
Dim C_AMT_DIFF        
Dim C_INC_PCNT        

'Const C_SHEETMAXROWS     = 100                                          '☜: Fetch count at a time
'Const C_SHEETMAXROWS_D   = 100                                          '☜: Fetch count at a time

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


'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : InitVariables()
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
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
' Name : SetDefaultVal()
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) --------------------------------------------------------------

	Dim StartDate
	Dim EndDate
	
	StartDate	= "<%=GetSvrDate%>"                                               'Get Server DB Date
	EndDate		= StartDate
	
	frm1.fpdtWk_yymm.focus
	frm1.fpdtWk_yymm.text	= UniConvDateAToB(StartDate ,parent.gServerDateFormat,parent.gDateFormat)
	frm1.fpdtWk_yymm1.text	= UniConvDateAToB(EndDate ,parent.gServerDateFormat,parent.gDateFormat) 
	 
	Call ggoOper.FormatDate(frm1.fpdtWk_yymm, Parent.gDateFormat, 2)
	Call ggoOper.FormatDate(frm1.fpdtWk_yymm1, Parent.gDateFormat, 2)
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub


'========================================================================================================
' Name : LoadInfTB19029()
' Desc : Set System Number format
' Para : 1. Currency
'        2. I(Input),Q(Query),P(Print),B(Bacth)
'        3. "*" is for Common module
'           "A" is for Accounting
'           "I" is for Inventory
'           ...
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) --------------------------------------------------------------

  <% Call loadInfTB19029A("G", "*", "COOKIE", "QA") %>

	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value
'========================================================================================================
Sub CookiePage(Kubun)

Dim iRow,iCol
Dim strYear,strMonth,strDay
Dim TempFrDt,TempToDt
   '------ Developer Coding part (Start ) --------------------------------------------------------------    
	With Frm1                      	
		Select Case Kubun
		Case 0
		
			If ReadCookie("JumpFlag")	<>"" Then .txtJumpFlag.Value	= ReadCookie("JumpFlag")
			
			If UCase(Trim(.txtJumpFlag.Value)) = "GC008MA1" Then
				If ReadCookie("FrYYYYMM")	<>"" Then 
					TempFrDt			= ReadCookie("FrYYYYMM")
					Call ExtractDateFrom(TempFrDt, Parent.gDateFormat, Parent.gComDateType, strYear, strMonth, strDay)
					.fpdtWk_yymm.Year	= strYear
					.fpdtWk_yymm.Month	= strMonth
					.fpdtWk_yymm.Day	= strDay
				End If
				
				If ReadCookie("ToYYYYMM")	<>"" Then 
					TempToDt			= ReadCookie("ToYYYYMM")
					Call ExtractDateFrom(TempToDt, Parent.gDateFormat, Parent.gComDateType, strYear, strMonth, strDay)
					.fpdtWk_yymm1.Year	= strYear
					.fpdtWk_yymm1.Month = strMonth
					.fpdtWk_yymm1.Day	= strDay
				End If

				If ReadCookie("ItemGrp")	<>"" Then .txtDeptCd.Value			= ReadCookie("ItemGrp")'				
				
				WriteCookie "FrYYYYMM"		, ""
				WriteCookie "ToYYYYMM"		, ""
				WriteCookie "ItemGrp"		, ""						
			
				If Trim(.fpdtWk_yymm.Text) <> "" and Trim(.fpdtWk_yymm1.Text) <> "" and Trim(.txtDeptCd.Value) <> "" Then
					Call MainQuery()
      			End If
      		End If
      		
      		WriteCookie "JumpFlag"		, ""
      		
		Case 1								
		    WriteCookie "FrYYYYMM"	, UniConvYYYYMMDDToDate(Parent.gDateFormat,Trim(.fpdtWk_yymm.Year), Right("0" & Trim(.fpdtWk_yymm.Month), 2),"01")
		    WriteCookie "ToYYYYMM"	, UniConvYYYYMMDDToDate(Parent.gDateFormat,Trim(.fpdtWk_yymm1.Year),Right("0" & Trim(.fpdtWk_yymm1.Month),2),"01")
		    
		    WriteCookie "JumpFlag"	, "GC006MA1"
		End Select
	End With
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : PgmJumpCheck
' Desc : This method set focus to pos of err
'========================================================================================================
Function PgmJumpCheck()         
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
	
	
	Call PgmJump(BIZ_PGM_JUMP_ID)
	    
End Function	  
'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================

Sub MakeKeyStream(pOpt)
    Dim strYYYYMM
    Dim strYear,strMonth,strDay

   '------ Developer Coding part (Start ) --------------------------------------------------------------
    Call ExtractDateFrom(frm1.fpdtWk_yymm.Text,frm1.fpdtWk_yymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
    strYYYYMM = strYear & Parent.gServerDateType  & strMonth

    lgKeyStream       = strYYYYMM & Parent.gColSep       'You Must append one character(Parent.gColSep)

    Call ExtractDateFrom(frm1.fpdtWk_yymm1.Text,frm1.fpdtWk_yymm1.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
    strYYYYMM = strYear & Parent.gServerDateType  & strMonth

    lgKeyStream       = lgKeyStream + strYYYYMM & Parent.gColSep
    lgKeyStream       = lgKeyStream + frm1.txtDeptCd.Value & Parent.gColSep              'You Must append one character(Parent.gColSep)
   '------ Developer Coding part (End   ) --------------------------------------------------------------
End Sub

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr
    Dim iNameArr
    Dim iDx

	'------ Developer Coding part (Start ) --------------------------------------------------------------


	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	 C_PROFIT_NM       = 1
	 C_PROFIT_AMT      = 2
	 C_PROFIT_PCNT     = 3
	 C_OLD_PROFIT_AMT  = 4
	 C_OLD_PROFIT_PCNT = 5
	 C_AMT_DIFF        = 6
	 C_INC_PCNT        = 7
End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	
	Call InitSpreadPosVariables()
	With frm1.vspdData

       .MaxCols   = C_INC_PCNT + 1                                                 ' ☜:☜: Add 1 to Maxcols
	   .Col       = .MaxCols                                                        ' ☜:☜: Hide maxcols
       .ColHidden = True                                                            ' ☜:☜:

        ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021216", ,parent.gAllowDragDropSpread
		
		ggoSpread.ClearSpreadData

	   .ReDraw = false

		Call GetSpreadColumnPos("A")
		
       ggoSpread.SSSetEdit   C_PROFIT_NM    , "손익항목"  ,23
       ggoSpread.SSSetFloat  C_PROFIT_AMT    , "금액",19,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	   ggoSpread.SSSetFloat  C_PROFIT_PCNT    , "%"  ,11,Parent.ggExchRateNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
       ggoSpread.SSSetFloat  C_OLD_PROFIT_AMT    , "전년동기",19,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	   ggoSpread.SSSetFloat  C_OLD_PROFIT_PCNT    , "%"  ,11,Parent.ggExchRateNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	   ggoSpread.SSSetFloat  C_AMT_DIFF    , "차이",19,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	   ggoSpread.SSSetFloat  C_INC_PCNT    , "증감율"  ,11,Parent.ggExchRateNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec

	   .ReDraw = true

       Call SetSpreadLock

    End With

End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
      ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1

    .vspdData.ReDraw = False
    .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> Parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to
              Exit For
           End If

       Next

    End If
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

			C_PROFIT_NM				= iCurColumnPos(1)
			C_PROFIT_AMT			= iCurColumnPos(2)
			C_PROFIT_PCNT			= iCurColumnPos(3)
			C_OLD_PROFIT_AMT		= iCurColumnPos(4)
			C_OLD_PROFIT_PCNT		= iCurColumnPos(5)
			C_AMT_DIFF				= iCurColumnPos(6)
			C_INC_PCNT				= iCurColumnPos(7)
    End Select    
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
'	Call InitData()
'	Call initMinor()
'	Call txtCurrencyCode_OnChange()
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

    Err.Clear                                                                        '☜: Clear err status

	Call LoadInfTB19029                                                              '☜: Load table , B_numeric_format

	'------ Developer Coding part (Start ) --------------------------------------------------------------

	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                            '⊙: Lock  Suitable  Field

    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                               '⊙: Setup the Spread sheet
	Call SetDefaultVal

	Call SetToolbar("1100000000001111")                                              '☆: Developer must customize
	Call CookiePage (0) 
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
    Dim fromYYYYMM
    Dim toYYYYMM
    Dim strYear,strMonth,strDay

    FncQuery = False															  '☜: Processing is NG
    Err.Clear                                                                     '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")			          '☜: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "2")										  '⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    If Not chkField(Document, "1") Then									          '⊙: This function check indispensable field
       Exit Function
    End If

'    Call SetDefaultVal
    Call InitVariables															  '⊙: Initializes local global variables

	'------ Developer Coding part (Start ) --------------------------------------------------------------
    If CompareDateByFormat(frm1.fpdtWk_yymm.Text,frm1.fpdtWk_yymm1.Text,frm1.fpdtWk_yymm.Alt,frm1.fpdtWk_yymm1.Alt,"970023",frm1.fpdtWk_yymm.UserDefinedFormat,Parent.gComDateType,True) = False Then
        frm1.fpdtWk_yymm1.focus
        Set gActiveElement = document.activeElement
        Exit Function
    End If

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
       IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")					  '☜: Data is changed.  Do you want to make it new?
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If

    Call ggoOper.ClearField(Document, "1")                                        '☜: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                        '☜: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    
    Call ggoOper.LockField(Document , "N")                                        '☜: Lock  Field
	'------ Developer Coding part (Start ) --------------------------------------------------------------
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

    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                           'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                 '☜: Please do Display first.
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")		                 '☜: Do you want to delete?
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
	'------ Developer Coding part (End )   --------------------------------------------------------------
    If DbSave = False Then                                                       '☜: Query db data
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
    Dim IntRetCD
    Dim imRow
    
    On Error Resume Next
    
    FncInsertRow = False                                                         '☜: Processing is NG
    Err.Clear  
    
    imRow = AskSpdSheetAddRowcount()                                                                  '☜: Clear err status

	If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

	If imRow = "" Then
		Exit function
	End If
		
	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow -1
       .vspdData.ReDraw = True
    End With
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
    
    IF Err.number = 0 Then
	    FncInsertRow = True                                                          '☜: Processing is OK
	End If
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

    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                           '☜: Please do Display first.
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    Call ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData


    Call SetDefaultVal
    Call InitVariables													         '⊙: Initializes local global variables

    if LayerShowHide(1) = false then
	    Exit Function
	end if

    Call MakeKeyStream("X")

    strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex             '☜: Next key tag
 '   strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D)        '☜: Max fetched data at a time
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

    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                           '☜: Please do Display first.
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    Call ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData


    Call SetDefaultVal
    Call InitVariables														     '⊙: Initializes local global variables

    if LayerShowHide(1) = false then
	    Exit Function
	end if

    Call MakeKeyStream("X")

    strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex             '☜: Next key tag
'    strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D)        '☜: Max fetched data at a time
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

	Call Parent.FncExport(Parent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind()
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncFind(Parent.C_MULTI, True)

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
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")			         '⊙: Data is changed.  Do you want to exit?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True                                                               '☜: Processing is OK
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
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

    if LayerShowHide(1) = false then
	    Exit Function
	end if

    strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex             '☜: Next key tag
'    strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D)        '☜: Max fetched data at a time
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

    Err.Clear                                                                    '☜: Clear err status
    DbSave = False                                                               '☜: Processing is NG
    if LayerShowHide(1) = false then
	    Exit Function
	end if

	'------ Developer Coding part (Start)  --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
    DbSave = True                                                                '☜: Processing is OK
    Set gActiveElement = document.ActiveElement
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This sub is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
    DbDelete = False                                                             '☜: Processing is NG
    if LayerShowHide(1) = false then
	    Exit Function
	end if

	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003                                '☜: Delete
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------

    DbDelete = True                                                              '☜: Processing is OK
	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Sub DbQueryOk()
	lgIntFlgMode      = Parent.OPMD_UMODE                                                   '⊙: Indicates that current mode is Create mode
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	Call SetToolbar("11000000000111111")                                              '☆: Developer must customize
	'------ Developer Coding part (End )   --------------------------------------------------------------
'    Call InitData()
'	Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement
End Sub

'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Sub DbSaveOk()
	Call InitVariables
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
    MainQuery()
    Set gActiveElement = document.ActiveElement
End Sub

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Sub DbDeleteOk()
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
	Call MainNew()
End Sub


'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================


'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(Col , Row)

    Dim iDx

   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

	'------ Developer Coding part (Start ) --------------------------------------------------------------
    Select Case Col
         Case  C_LcKindNm
                iDx = Frm1.vspdData.value
   	            Frm1.vspdData.Col = C_LcKindCd
                Frm1.vspdData.value = iDx
         Case Else
    End Select
	'------ Developer Coding part (End   ) --------------------------------------------------------------

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
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
'   Event Name : vspdData_ColWidthChange
'   Event Desc : This function is called where spread sheet column width change
'========================================================================================================
sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub  


'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
        If lgStrPrevKeyIndex <> "" Then
      	   Call DisableToolBar(Parent.TBC_QUERY)
      	   If DBQuery = false Then
      	    Call RestoreToolBar()
      	    Exit Sub
      	   End If
        End If
    End if
End Sub

<%'======================================================================================================
'	Name : OpenCode()
'	Description : Major PopUp
'=======================================================================================================%>
Function OpenCode()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "Profit Center"		    	<%' 팝업 명칭 %>
	arrParam(1) = "B_COST_CENTER"           <%' TABLE 명칭 %>
	arrParam(2) = frm1.txtDeptCd.value                        <%' Code Condition%>
	arrParam(3) = "" 		            	<%' Name Cindition%>
	arrParam(4) = ""                        <%' Where Condition%>
	arrParam(5) = "Profit Center"

    arrField(0) = "COST_CD"	     			<%' Field명(1)%>
    arrField(1) = "COST_NM"					<%' Field명(0)%>


    arrHeader(0) = "Profit Center코드"			    	<%' Header명(0)%>
    arrHeader(1) = "Profit Center명"				<%' Header명(1)%>

	arrRet = window.showModalDialog("../../comasp/AdoCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=470px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtDeptCd.focus
		Exit Function
	Else
		Call SetCode(arrRet)
	End If

End Function

<%'======================================================================================================
'	Name : SetCode()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================%>
Function SetCode(Byval arrRet)
	With frm1
		.txtDeptCd.focus
		.txtDeptCd.value = arrRet(0)
		.txtDeptNm.value = arrRet(1)
	End With
End Function

'========================================================================================================
'   Event Name : fpdtWk_yymm_DblClick
'   Event Desc :
'========================================================================================================
Sub fpdtWk_yymm_DblClick(Button)
	If Button = 1 Then
		frm1.fpdtWk_yymm.Action = 7
		Call SetFocusToDocument("M")
		frm1.fpdtWk_yymm.focus
	End If
End Sub

'========================================================================================================
'   Event Name : fpdtWk_yymm_KeyPress
'   Event Desc :
'========================================================================================================
Sub fpdtWk_yymm_KeyPress(Key)
    If key = 13 Then
        Call MainQuery
	End If
End Sub

'========================================================================================================
'   Event Name : fpdtWk_yymm1_DblClick
'   Event Desc :
'========================================================================================================
Sub fpdtWk_yymm1_DblClick(Button)
	If Button = 1 Then
		frm1.fpdtWk_yymm1.Action = 7
		Call SetFocusToDocument("M")
		frm1.fpdtWk_yymm1.focus
	End If
End Sub

'========================================================================================================
'   Event Name : fpdtWk_yymm_KeyPress
'   Event Desc :
'========================================================================================================
Sub fpdtWk_yymm1_KeyPress(Key)
    If key = 13 Then
        Call MainQuery
	End If
End Sub
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'########################################################################################################
-->
<BODY SCROLL="No" TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD NOWRAP  <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD NOWRAP >
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD NOWRAP  WIDTH=10>&nbsp;</TD>
					<TD NOWRAP  CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD NOWRAP  background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<TD NOWRAP  background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>P/C별 손익현황</font></td>
								<TD NOWRAP  background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD NOWRAP  WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD NOWRAP  WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD NOWRAP  <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD NOWRAP  HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD NOWRAP  CLASS=TD5 NOWRAP>작업년월</TD>
									<TD NOWRAP  CLASS=TD6 NOWRAP><script language =javascript src='./js/gc006ma1_fpdtWk_yymm_fpdtWk_yymm.js'></script>~<script language =javascript src='./js/gc006ma1_fpdtWk_yymm1_fpdtWk_yymm1.js'></script></TD>
                                    <TD NOWRAP  CLASS=TD5 NOWRAP>Profit Center</TD>
									<TD NOWRAP  CLASS=TD6 NOWRAP>
									 <INPUT TYPE=TEXT NAME="txtDeptCd" SIZE=15 MAXLENGTH=10 tag="12XXXU" ALT="Profit Center코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCode" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCode()">
									 <INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=20 MAXLENGTH=20 tag="14" ALT="Profit Center명" >
									 </TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD NOWRAP  <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD NOWRAP  WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD NOWRAP  HEIGHT="100%">
									<script language =javascript src='./js/gc006ma1_vspdData_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD NOWRAP  <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			</TABLE>
		</TD>

	</TR>
	<TR >
		<TD NOWRAP  WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=0 noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows"      TAG="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtMode"         TAG="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"    TAG="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"      TAG="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"     TAG="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtJumpFlag"	  TAG="24" TABINDEX = "-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

