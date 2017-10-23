
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : a8103ma1
'*  4. Program Name         : 본지점전표승인 
'*  5. Program Desc         : 입력된 결의전표중 본지점전표인 것을 모전표/자전표를 동시 승인한다.
'*  6. Component List       : PADG035.dll
'*  7. Modified date(First) : 2001/01/18
'*  8. Modified date(Last)  : 2003/06/05
'*  9. Modifier (First)     : hersheys
'* 10. Modifier (Last)      : Lim YOung Woon
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="vbscript">
Option Explicit    
'########################################################################################################
'#                       4.  Data Declaration Part
'========================================================================================================
'=                       4.1 External ASP File
Const BIZ_PGM_ID = "a8107mb1_KO441.asp"                                              '☆: 비지니스 로직 ASP명 
'========================================================================================================
'=                       4.2 Constant variables For spreadsheet
'========================================================================================================
'=                       4.3 Common variables 
<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================================
'=                       4.4 User-defind Variables
Dim C_Confirm
Dim C_Conf_Nm
Dim C_Conf_fg
Dim C_bConf_fg
Dim C_TempGlDt
Dim C_GlDt
Dim C_TempGlNo
Dim C_DeptNm
Dim C_Currency
Dim C_TempGlAmt
Dim C_TempGlLocAmt
Dim C_GlNo
Dim C_InputType
Dim C_InputTypeNm
Dim C_HqBrchNo
Dim C_BConfirm
Dim C_BConf_Nm
Dim C_BTempGlDt
Dim C_BTempGlNo
Dim C_BDeptNm
Dim C_BCurrency
Dim C_BTempGlAmt
Dim C_BTempGlLocAmt
Dim C_BGlDt
Dim C_BGlNo
Dim C_BInputType
Dim C_BInputTypeNm
Dim C_BHqBrchNo
Dim C_USER
Dim C_BUSER

Dim lgStrPrevKeyTempGlDt
Dim lgStrPrevKeyTempGlNo
Dim lgStrPrevKeyTempGlNo2
Dim IsOpenPop
Dim lgCurrRow1
Dim lgQueryfg

<%
Dim lsSvrDate
lsSvrDate = GetSvrDate
%>

'########################################################################################################
'#                       5.Method Declaration Part
'========================================================================================================
'                        5.1 Common Group-1
'========================================================================================================
Sub initSpreadPosVariables(ByVal pOpt)
	Select Case  pOpt
		Case "A"
		
			C_Confirm      = 1
			C_Conf_Nm      = 2
			C_Conf_Fg      = 3
			C_USER         = 4 
			C_TempGlDt     = 5 
			C_GlDt         = 6 
			C_TempGlNo     = 7 
			C_DeptNm       = 8 
			C_Currency     = 9 
			C_TempGlAmt    = 10
			C_TempGlLocAmt = 11
			C_GlNo         = 12
			C_InputType    = 13
			C_InputTypeNm  = 14
			C_HqBrchNo     = 15
		 
		 
		Case "B"
			 C_BConfirm      = 1
			 C_BConf_Nm      = 2
			 C_BConf_Fg      = 3
			 C_BTempGlDt     = 4
			 C_BGlDt         = 5
			 C_BTempGlNo     = 6
			 C_BDeptNm       = 7
			 C_BCurrency     = 8
			 C_BTempGlAmt    = 9
			 C_BTempGlLocAmt = 10
			 C_BGlNo         = 11
			 C_BInputType    = 12
			 C_BInputTypeNm  = 13
			 C_BHqBrchNo     = 14
			 C_BUSER         = 15
	End Select
End Sub

'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE   
    lgBlnFlgChgValue = False  
    lgIntGrpCount = 0        
    lgStrPrevKeyTempGlDt = ""    
    lgStrPrevKeyTempGlNo = ""    
    lgStrPrevKeyTempGlNo2 = ""   
    lgLngCurRows = 0       
    lgPageNo = 0       


End Sub
'========================================================================================================
Sub SetDefaultVal()

	frm1.txtFromTempGlDt.text = UNIDateAdd("m", -1, UNIDateClientFormat("<%=lsSvrDate%>"), parent.gDateFormat)
	frm1.txtToTempGlDt.text   = UNIDateClientFormat("<%=lsSvrDate%>")
	frm1.cboConfFg.value	=	"U"
	frm1.txtBizAreaCd.focus	

End Sub

'========================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub
'========================================================================================================
Sub InitComboBox()

	Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("A1007", "''", "S") & "  ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
    Call SetCombo2(frm1.cboConfFg, lgF0, lgF1, Chr(11))

End Sub

'=========================================================================================================
Sub InitSpreadComboBox(Byval pOpt)
	
	Select Case pOpt
		Case "A"	
			ggoSpread.Source = frm1.vspdData

			Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", "MAJOR_CD=" & FilterVar("A1001", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
			ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_InputType
			ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_InputTypeNm

		Case "B"
			ggoSpread.Source = frm1.vspdData2

			Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", "MAJOR_CD=" & FilterVar("A1001", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
			ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_InputType
			ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_InputTypeNm

	End Select

End Sub

'========================================================================================================
Sub InitData(ByVal pOpt)

	Dim intRow
	Dim intIndex

	Select Case pOpt
		Case "A"
			For intRow = 1 to frm1.vspdData.MaxRows
				frm1.vspdData.Row = intRow
'				frm1.vspdData.Col = C_Confirm
'				intIndex = frm1.vspdData.value
'				frm1.vspdData.col = C_Conf_Nm
'				frm1.vspdData.value = intindex
				frm1.vspdData.Col = C_InputType
				intIndex = frm1.vspdData.value
				frm1.vspdData.col = C_InputTypeNm
				frm1.vspdData.value = intindex
			Next

		Case "B"
			For intRow = 1 to frm1.vspdData2.MaxRows
				frm1.vspdData2.Row = intRow
'				frm1.vspdData2.Col = C_BConfirm
'				intIndex = frm1.vspdData2.value
'				frm1.vspdData2.col = C_BConf_Nm
'				frm1.vspdData2.value = intindex
				frm1.vspdData2.Col = C_BInputType
				intIndex = frm1.vspdData2.value
				frm1.vspdData2.col = C_BInputTypeNm
				frm1.vspdData2.value = intindex
			Next

	End Select

End Sub

'========================================================================================================
Sub InitSpreadSheet(ByVal pOpt)

    Dim sList

	Call initSpreadPosVariables(pOpt)

	Select Case pOpt
		Case "A"
			ggoSpread.Source = frm1.vspdData
			ggoSpread.SpreadInit "V20021225",,parent.gAllowDragDropSpread

			With frm1.vspdData

				.ReDraw = False

				.MaxCols	= C_HqBrchNo + 1
				Call ggoSpread.ClearSpreadData()
				Call GetSpreadColumnPos(pOpt)
'				ggoSpread.SSSetCombo C_Confirm,      "",					8
'				ggoSpread.SSSetCombo C_Conf_Nm,      "승인",			10

				ggoSpread.SSSetCheck C_Confirm,     "",     8,  -10, "", True, -1         
				ggoSpread.SSSetEdit C_Conf_Nm,      "", 8, 2,,3                  
			        ggoSpread.SSSetEdit  C_Conf_Fg,      "결재여부",   15, ,,15
			        ggoSpread.SSSetEdit	 C_USER,		 "결재자",   10,,,10
				ggoSpread.SSSetDate  C_TempGlDt,     "결의전표일",		12, 2, parent.gDateFormat
				ggoSpread.SSSetDate  C_GlDt,         "전표일",			12, 2, parent.gDateFormat
				ggoSpread.SSSetEdit  C_TempGlNo,     "결의번호",		12, 2, , 20
				ggoSpread.SSSetEdit  C_DeptNm,       "회계부서명",		19,  , , 30
				ggoSpread.SSSetEdit  C_Currency,     "통화",			8,   , , 3
				ggoSpread.SSSetFloat C_TempGlAmt,    "금액",			13, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
				ggoSpread.SSSetFloat C_TempGlLocAmt, "금액(자국)",		13, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
				ggoSpread.SSSetEdit  C_GlNo,         "전표번호",		12, 2, , 20
				ggoSpread.SSSetCombo C_InputType,    "",					8
				ggoSpread.SSSetCombo C_InputTypeNm,  "입력경로",		14
				ggoSpread.SSSetEdit  C_HqBrchNo,     "참조번호",		14,  , , 20
				
	
				Call ggoSpread.SSSetColHidden(C_Conf_Nm,C_Conf_Nm,True)
				Call ggoSpread.SSSetColHidden(C_Currency,C_Currency,True)
				Call ggoSpread.SSSetColHidden(C_InputType,C_InputType,True)
				Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)

				.ReDraw = True

			End With

		Case "B"
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.SpreadInit "V20021224",,parent.gAllowDragDropSpread

			With frm1.vspdData2

				.ReDraw = False

				.MaxCols	 = C_BUSER + 1
				Call ggoSpread.ClearSpreadData()
				Call GetSpreadColumnPos(pOpt)

				ggoSpread.SSSetCheck C_BConfirm,     "",     8,  -10, "", True, -1         
				ggoSpread.SSSetEdit C_BConf_Nm,      "", 10, 2,,3                  
                                ggoSpread.SSSetEdit  C_BConf_Fg,      "최종승인여부",   15, ,,15
'				ggoSpread.SSSetCombo C_BConfirm,      "",				8
'				ggoSpread.SSSetCombo C_BConf_Nm,      "승인",		10
				ggoSpread.SSSetDate  C_BTempGlDt,     "결의전표일",	12, 2, parent.gDateFormat
				ggoSpread.SSSetDate  C_BGlDt,         "전표일",		12, 2, parent.gDateFormat
				ggoSpread.SSSetEdit  C_BTempGlNo,     "결의번호",	12, 2, , 20
				ggoSpread.SSSetEdit  C_BDeptNm,       "회계부서명",	19,  , , 30
				ggoSpread.SSSetEdit  C_BCurrency,     "통화",		8,   , , 3
				ggoSpread.SSSetFloat C_BTempGlAmt,    "금액",		13, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
				ggoSpread.SSSetFloat C_BTempGlLocAmt, "금액(자국)",	13, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
				ggoSpread.SSSetEdit  C_BGlNo,         "전표번호",	12, 2, , 20
				ggoSpread.SSSetCombo C_BInputType,    "",				8
				ggoSpread.SSSetCombo C_BInputTypeNm,  "입력경로",	14
				ggoSpread.SSSetEdit  C_BHqBrchNo,     "참조번호",	14, , , 20
                                ggoSpread.SSSetEdit	 C_BUSER,		 "최종승인자",   10,,,10

				Call ggoSpread.SSSetColHidden(C_BConf_Nm,C_BConf_Nm,True)
				Call ggoSpread.SSSetColHidden(C_BConf_Fg,C_BConf_Fg,True)
				Call ggoSpread.SSSetColHidden(C_BUSER,C_BUSER,True)
				
				Call ggoSpread.SSSetColHidden(C_BCurrency,C_BCurrency,True)
				Call ggoSpread.SSSetColHidden(C_BInputType,C_BInputType,True)
				Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)


				.ReDraw = True

			End With

	End Select

	Call SetSpreadLock(pOpt, -1, -1, "I")

End Sub

'=========================================================================================================
Sub SetSpreadLock(Byval pOpt, ByVal lRow , ByVal lRow2 , Byval stsFg)

	Select Case pOpt
		Case "A"
			ggoSpread.Source = frm1.vspdData
			frm1.vspdData.ReDraw = False

			ggoSpread.SpreadLock C_Conf_Nm,      lRow, C_Conf_Nm,      lRow2
                        ggoSpread.SpreadLock C_Conf_Fg,      lRow, C_Conf_Fg,      lRow2
			ggoSpread.SpreadLock C_TempGlDt,     lRow, C_TempGlDt,     lRow2
			ggoSpread.SpreadLock C_TempGlNo,     lRow, C_TempGlNo,     lRow2
			ggoSpread.SpreadLock C_DeptNm,       lRow, C_DeptNm,       lRow2
			ggoSpread.SpreadLock C_TempGlAmt,    lRow, C_TempGlAmt,    lRow2
			ggoSpread.SpreadLock C_TempGlLocAmt, lRow, C_TempGlLocAmt, lRow2
			ggoSpread.SpreadLock C_GlNo,         lRow, C_GlNo,         lRow2
			ggoSpread.SpreadLock C_USER,         lRow, C_USER,         lRow2

			
			If stsFg = "Q" Then
				ggoSpread.SpreadLock C_GlDt,         lRow, C_GlDt,         lRow2
			End If
			
			ggoSpread.SpreadLock C_InputTypeNm,  lRow, C_InputTypeNm,  lRow2
			ggoSpread.SpreadLock C_HqBrchNo,     lRow, C_HqBrchNo,     lRow2
			ggoSpread.SSSetProtected	frm1.vspdData.MaxCols,-1,-1	
			frm1.vspdData.ReDraw = True

		Case "B"
			ggoSpread.Source = frm1.vspdData2
			frm1.vspdData2.ReDraw = False
			ggoSpread.SpreadLock C_BConfirm,      lRow, C_BConfirm,      lRow2
			ggoSpread.SpreadLock C_BConf_Nm,      lRow, C_BConf_Nm,      lRow2
                        ggoSpread.SpreadLock C_BConf_Fg,      lRow, C_BConf_Fg,      lRow2
			ggoSpread.SpreadLock C_BTempGlDt,     lRow, C_BTempGlDt,     lRow2
			ggoSpread.SpreadLock C_BTempGlNo,     lRow, C_BTempGlNo,     lRow2
			ggoSpread.SpreadLock C_BDeptNm,       lRow, C_BDeptNm,       lRow2
			ggoSpread.SpreadLock C_BTempGlAmt,    lRow, C_BTempGlAmt,    lRow2
			ggoSpread.SpreadLock C_BTempGlLocAmt, lRow, C_BTempGlLocAmt, lRow2
			ggoSpread.SpreadLock C_BGlDt,         lRow, C_BGlDt,         lRow2
			ggoSpread.SpreadLock C_BGlNo,         lRow, C_BGlNo,         lRow2
			ggoSpread.SpreadLock C_BInputTypeNm,  lRow, C_BInputTypeNm,  lRow2
			ggoSpread.SpreadLock C_BHqBrchNo,     lRow, C_BHqBrchNo,     lRow2
			ggoSpread.SpreadLock C_BUSER,         lRow, C_BUSER,         lRow2
			ggoSpread.SSSetProtected	frm1.vspdData2.MaxCols,-1,-1	
			frm1.vspdData2.ReDraw = True
	End Select

End Sub

'=========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

	
	if  pvEndRow = "" THEN	pvEndRow = pvStartRow
    With frm1
	    .vspdData.ReDraw = False
		IF frm1.cboConfFg.value = "C" Then
			ggoSpread.SSSetProtected		C_GlDt,         pvStartRow, pvEndRow
		ELSE
			ggoSpread.SSSetRequired		C_GlDt,         pvStartRow, pvEndRow
		END IF	
		ggoSpread.SSSetProtected	C_GlNo,			pvStartRow, pvEndRow 
		.vspdData.ReDraw = True
    End With

End Sub
'=========================================================================================================

Sub SetSpreadColor2(ByVal lRow)
    
    With frm1.vspdData2
    
	    .Redraw = False
    
	    ggoSpread.Source = frm1.vspdData2
		
		' SSSetRequired(ByVal Col, ByVal Row, Optional ByVal Row2 = -10)
		'ggoSpread.SSSetRequired C_CtrlCtrlNm, lRow, lRow		' 관리항목 

	    .Col = 1
		.Row = .ActiveRow
	    .Action = 0                         'SS_ACTION_ACTIVE_CELL
		.EditMode = True
    
		.Redraw = True
    
    End With

End Sub

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
              Frm1.vspdData.Action = 0
              Exit For
           End If
       Next
    End If

End Sub

'======================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)

	Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
			C_Confirm      = iCurColumnPos(1  )
			C_Conf_Nm      = iCurColumnPos(2  )
			C_Conf_Fg      = iCurColumnPos(3  )
			C_USER         = iCurColumnPos(4  )
			C_TempGlDt     = iCurColumnPos(5  )
			C_GlDt         = iCurColumnPos(6  )
			C_TempGlNo     = iCurColumnPos(7  )
			C_DeptNm       = iCurColumnPos(8  )
			C_Currency     = iCurColumnPos(9  )
			C_TempGlAmt    = iCurColumnPos(10 )
			C_TempGlLocAmt = iCurColumnPos(11 )
			C_GlNo         = iCurColumnPos(12 )
			C_InputType    = iCurColumnPos(13 )
			C_InputTypeNm  = iCurColumnPos(14 )
			C_HqBrchNo     = iCurColumnPos(15 )

		Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_BConfirm       = iCurColumnPos(1)
			C_BConf_Nm       = iCurColumnPos(2)
			C_BConf_Fg       = iCurColumnPos(3)
			C_BTempGlDt      = iCurColumnPos(4)
			C_BGlDt          = iCurColumnPos(5)
			C_BTempGlNo      = iCurColumnPos(6)
			C_BDeptNm        = iCurColumnPos(7)
			C_BCurrency      = iCurColumnPos(8)
			C_BTempGlAmt     = iCurColumnPos(9)
			C_BTempGlLocAmt  = iCurColumnPos(10)
			C_BGlNo          = iCurColumnPos(11)
			C_BInputType     = iCurColumnPos(12)
			C_BInputTypeNm   = iCurColumnPos(13)
			C_BHqBrchNo      = iCurColumnPos(14)
			C_BUSER      = iCurColumnPos(15)

    End Select

End Sub

'========================================================================================================
'                        5.2 Common Group-2
'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029()                                                      'Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                      'Lock  Suitable  Field

    Call InitSpreadSheet("A")                                                  'Setup the Spread shee

    Call InitSpreadSheet("B")
	Call InitComboBox()
	Call InitSpreadComboBox("A")
	Call InitSpreadComboBox("B")
    Call SetDefaultVal()

    Call SetToolbar("1100000000001111")                                        '버튼 툴바 제어 

End Sub

'========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=========================================================================================================
Function FncQuery()

	Dim IntRetCD
	
	on Error Resume Next
	Err.Clear

    FncQuery = False

	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
    
	Call ggoSpread.ClearSpreadData()
	
    If Not chkField(Document, "1") Then                                        'This function check indispensable field
       Exit Function
    End If

    If CompareDateByFormat(frm1.txtFromTempGlDt.Text, frm1.txtToTempGlDt.Text, frm1.txtFromTempGlDt.Alt, frm1.txtToTempGlDt.Alt, _
                        "970025", frm1.txtFromTempGlDt.UserDefinedFormat, parent.gComDateType, True) = False Then
          Exit Function
    End If

    Call InitVariables()                                                       'Initializes local global variables
    
    If DbQuery = False Then
		Exit Function
    End If                                                                     '☜: Query db data

    If Err.number = 0 Then
       FncQuery = True                                                         '☜: Processing is OK
    End If

    lgQueryfg = false

	Set gActiveElement = document.ActiveElement

End Function

'==========================================================================================================
Function FncNew()

	Dim IntRetCD

	on Error Resume Next
	Err.Clear

    FncNew = False
    If Err.number = 0 Then
		FncNew = True                                                          '⊙: Processing is OK
    End If

	Set gActiveElement = document.ActiveElement

End Function

'==========================================================================================================
Function FncDelete()

	Dim IntRetCD

	on Error Resume Next
	Err.Clear

    FncDelete = False
    If Err.number = 0 Then
		FncDelete = True
	End If
    
    Set gActiveElement = document.ActiveElement

End Function

'==========================================================================================================
Function FncSave() 

	Dim IntRetCD

    On Error Resume Next
    Err.Clear

    FncSave = False

	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSDefaultCheck = False Then
	    IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                      '☜ 바뀐부분 
	    Exit Function
	End If

    If Not chkField(Document, "2") Then                                        'Check contents area
       Exit Function
    End If

    If DbSave = False Then                                                     '☜: Save db data
		Exit Function
	End If

    If Err.number = 0 Then
		FncSave = True
	End If

    Set gActiveElement = document.ActiveElement    

End Function

'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    On Error Resume Next
    Err.Clear

    FncCopy = False                                                            '☜: Processing is NG

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
    
    ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
            ggoSpread.CopyRow
			SetSpreadColor .ActiveRow, .ActiveRow
            .ReDraw = True
		    .Focus
		 End If
	End With
    If Err.number = 0 Then	
       FncCopy = True                                                          '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function

'==========================================================================================================
Function FncCancel()

	Dim FormCnt

    On Error Resume Next
    Err.Clear

	FncCancel = False

	If frm1.vspdData.MaxRows < 1 Then Exit Function

    frm1.vspdData.Row = frm1.vspdData.ActiveRow
    frm1.vspdData.Col = 0

    ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo
	
	Call InitData("A")
'	Call vspdData_ComboSelChange(C_Conf_Nm, frm1.vspdData.ActiveRow)

    If Err.number = 0 Then	
		FncCancel = True
	End If
	
	Set gActiveElement = document.ActiveElement

End Function

'==========================================================================================================
Function FncInsertRow()

	Dim lngNum

    On Error Resume Next                                                       '☜: If process fails
    Err.Clear                                                                  '☜: Clear error status
    
    FncInsertRow = False                                                       '☜: Processing is NG

    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else
        imRow = AskSpdSheetAddRowCount()
        If imRow = "" Then
            Exit Function
        End If
    End If

	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow ,imRow
        Call SetSpreadColor(.vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1)
        .vspdData.ReDraw = True
    End With
    If Err.number = 0 Then
       FncInsertRow = True                                                     '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement

End Function

'==========================================================================================================
Function FncDeleteRow()
    Dim lDelRows

    On Error Resume Next                                                       '☜: If process fails
    Err.Clear                                                                  '☜: Clear error status

    FncDeleteRow = False                                                       '☜: Processing is NG

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
    If Err.number = 0 Then	
       FncDeleteRow = True                                                     '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function FncPrint()

    On Error Resume Next                                                       '☜: If process fails
    Err.Clear                                                                  '☜: Clear error status

    FncPrint = False                                                           '☜: Processing is NG
	Call Parent.FncPrint()                                                     '☜: Protect system from crashing

    If Err.number = 0 Then	 
       FncPrint = True                                                         '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function FncPrev() 

    On Error Resume Next                                                       '☜: If process fails
    Err.Clear                                                                  '☜: Clear error status

    FncPrev = False                                                            '☜: Processing is NG
    If Err.number = 0 Then	 
       FncPrev = True                                                          '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function FncNext() 

    On Error Resume Next                                                       '☜: If process fails
    Err.Clear                                                                  '☜: Clear error status

    FncNext = False                                                            '☜: Processing is NG

    If Err.number = 0 Then	 
       FncNext = True                                                          '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function FncExcel() 

    On Error Resume Next                                                       '☜: If process fails
    Err.Clear                                                                  '☜: Clear error status

    FncExcel = False                                                           '☜: Processing is NG

	Call Parent.FncExport(Parent.C_MULTI)

    If Err.number = 0 Then	 
       FncExcel = True                                                         '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
Function FncFind() 

    On Error Resume Next                                                       '☜: If process fails
    Err.Clear                                                                  '☜: Clear error status

    FncFind = False                                                            '☜: Processing is NG

	Call Parent.FncFind(Parent.C_MULTI, True)

    If Err.number = 0 Then	 
       FncFind = True                                                          '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'==========================================================================================================
Function FncExit()

	Dim IntRetCD

    On Error Resume Next                                                       '☜: If process fails
    Err.Clear                                                                  '☜: Clear error status

	FncExit = False

	ggoSpread.Source = Frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")         '⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    If Err.number = 0 Then	 
       FncExit = True                                                          '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function

'==========================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'==========================================================================================================
Sub PopRestoreSpreadColumnInf()

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Select Case Trim(UCase(gActiveSpdSheet.Name))
		Case "VSPDDATA"
			Call InitSpreadSheet("A")
			Call InitSpreadComboBox("A")
			Call ggoSpread.ReOrderingSpreadData()
			Call InitData("A")
'			Call vspdData_ComboSelChange(C_Conf_Nm, frm1.vspdData.ActiveRow)
			
		Case "VSPDDATA2"
			Call InitSpreadSheet("B")
			Call InitSpreadComboBox("B")
			Call ggoSpread.ReOrderingSpreadData()
			Call InitData("B")
'			Call vspdData_ComboSelChange(C_Conf_Nm, frm1.vspdData.ActiveRow)

	End Select

End Sub

'==========================================================================================================
Function DbQuery()

    Dim strVal

    On Error Resume Next                                                       '☜: If process fails
    Err.Clear                                                                  '☜: Clear error status

    DbQuery = False
	
	ggoSpread.Source = frm1.vspdData2
	Call ggoSpread.ClearSpreadData()
	
	Call DisableToolBar(Parent.TBC_QUERY)
    Call LayerShowHide(1)

	frm1.hOrgChangeId.value = parent.gChangeOrgId

    With frm1


		strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0001						'☜:조회표시 
		strVal = strVal & "&lgStrPrevKeyTempGlDt=" & lgStrPrevKeyTempGlDt
		strVal = strVal & "&lgStrPrevKeyTempGlNo=" & lgStrPrevKeyTempGlNo
		strVal = strVal & "&lgStrPrevKeyTempGlNo2=" & lgStrPrevKeyTempGlNo2
		strVal = strVal & "&txtBizAreaCd="         & Trim(.txtBizAreaCd.value)
		strVal = strVal & "&cboConfFg="            & Trim(.cboConfFg.value)
		strVal = strVal & "&hOrgChangeId="         & Trim(.hOrgChangeId.value)
		strVal = strVal & "&txtFromTempGlDt="      & UNICONVDATE(Trim(.txtFromTempGlDt.TEXT))
		strVal = strVal & "&txtToTempGlDt="        & UNICONVDATE(Trim(.txtToTempGlDt.Text))
		strVal = strVal & "&txtMaxRows="           & .vspdData.MaxRows

    End With

	Call RunMyBizASP(MyBizASP, strVal)                                         '☜: 비지니스 ASP 를 가동 
	
    If Err.number = 0 Then
       DbQuery = True                                                          '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function

'=========================================================================================================
Function DbQuery2(ByVal Row)

	Const iUID_M9999 = 9999

	Dim strVal
	Dim lngRows

	on Error Resume Next
	Err.Clear

	DbQuery2 = False

	Call DisableToolBar(Parent.TBC_QUERY)
	Call LayerShowHide(1)

	frm1.hOrgChangeID.value = parent.gChangeOrgId

	With frm1

		.vspdData.Row = Row
		.vspdData.Col = C_Confirm
        .txtConfirm.value = .vspdData.Text

        .vspdData.Col = C_HqBrchNo
		.txtHqBrchNo.value = .vspdData.Text

		.vspdData.Col = C_TempGlNo
		.txtTempGLNo.value = .vspdData.Text

		strVal = BIZ_PGM_ID & "?txtMode="		& iUID_M9999
		strVal = strVal		 & "&txtBizAreaCd=" & Trim(.txtBizAreaCd.value)
		strVal = strVal		 & "&txtHqBrchNo=" & Trim(.txtHqBrchNo.value)
		strVal = strVal		 & "&cboConfFg="    & Trim(.txtConfirm.value)
		strVal = strVal		 & "&txtTempGLNo=" & Trim(.txtTempGLNo.value)
		strVal = strVal		 & "&txtMaxRows="   & .vspdData2.MaxRows
		strVal = strVal		 & "&hOrgChangeID=" & Trim(.hOrgChangeID.value)
		strVal = strVal		 & "&vspdDataRow=" & Row


	End With

	Call RunMyBizASP(MyBizASP, strVal)                                         '☜: 비지니스 ASP 를 가동 

    If Err.number = 0 Then
       DbQuery = True                                                          '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function

'==========================================================================================================
Function DbSave() 

    Dim lRow
    Dim lGrpCnt
	Dim strVal

	on Error Resume Next
	Err.Clear

    DbSave = False

    Call DisableToolBar(Parent.TBC_SAVE)                                       '☜: Disable Save Button Of ToolBar
    Call LayerShowHide(1)

	With frm1

		.txtMode.value        = parent.UID_M0002
		.txtInsrtUserId.value = parent.gUsrID
		.txtUpdtUserId.value  = parent.gUsrID
		lGrpCnt = 1
		strVal = ""


		For lRow = 1 To .vspdData.MaxRows
    		.vspdData.Row	=	lRow
			.vspdData.col	=	C_Confirm	
			
			IF 	.vspdData.text = "1" Then
					strVal = strVal & "U" & parent.gColSep				'☜: U=Update
					
					.vspdData.Col = C_GlDt
					strVal = strVal & UNIConvDate(Trim(.vspdData.Text)) & parent.gColSep
					
					.vspdData.Col = C_TempGlNo
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					
					IF Trim(.cboConfFg.value) = "U" Then 
						strVal = strVal & "C" & parent.gColSep
					Else
						strVal = strVal & "U" & parent.gColSep
					END IF	
					strVal = strVal & Trim(.txtBizAreaCd.value) &parent.gColSep & lRow & parent.gRowSep
					lGrpCnt = lGrpCnt + 1
			End If
		Next
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value  = strVal
		
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)                                        '☜: 비지니스 ASP 를 가동 

    If Err.number = 0 Then	 
       DbSave = True                                                           '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'==========================================================================================================
Function DbDelete()

    On Error Resume Next                                                       '☜: If process fails
    Err.Clear                                                                  '☜: Clear error status

    DbDelete = False                                                           '☜: Processing is NG
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	 
       DbDelete = True                                                         '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function
    
'=========================================================================================================
Sub DbQueryOk()
Dim i
Dim IntRetCD 
	lgIntFlgMode = parent.OPMD_UMODE                                       'Indicates that current mode is Update mode
	Call InitData("A")                                                     'Combo의 Name을 Code를 기준으로 맞춤 
    Call SetToolbar("1100100100011111")                                    '버튼 툴바 제어 
	Call SetSpreadColor (1, frm1.vspddata.maxrows )


              
	
    If frm1.vspdData.MaxRows > 0 Then
        Call DbQuery2(1)                                                   '조회된 상단그리드의 첫번째 레코드를 기준으로 하단그리드의 내용을 가져온다.
    End If

    Call ggoOper.LockField(Document, "Q")                                  'This function lock the suitable field

	lgQueryfg = true
End Sub

'=========================================================================================================
Sub DbQueryOk2()
	Dim Cnt
	Dim intItemCnt	
Dim IntRetCD 
    frm1.vspdData2.Redraw = True

	With frm1
		.vspdData.Col = 1:    intItemCnt = .vspddata.MaxRows
        '-----------------------
        'Reset variables area
        '-----------------------
        ggoSpread.Source = .vspdData2
	    Call SetSpreadLock( 1, 1, "", "Q")
	    Call InitData("2")							' Combo의 Name을 Code를 기준으로 맞춤 
		For Cnt = 1 to .vspdData2.MaxRows
	        Call SetSpreadColor2(Cnt)
			' 상단의 전표일자를 해당하는 하단의 전표일자에 값을 넣는다.
	        .vspdData.Col = C_GlDt
	        .vspdData.row = .vspdData.activerow
	        .vspdData2.Col = C_GlDt
	        .vspdData2.row = Cnt
	        .vspdData2.text = .vspdData.text

			

		Next
		ggoSpread.Source = .vspdData
    End With

'	Call vspdData_ComboSelChange(C_Conf_Nm, frm1.vspdData.ActiveRow)
	
End Sub

'==========================================================================================================
Sub DbSaveOk()
    
    Call InitVariables()                                                       '⊙: Initializes local global variables
    ggoSpread.source = frm1.vspdData                                             '⊙: Clear Contents  Field
    Call ggoSpread.ClearSpreadData()
    ggoSpread.source = frm1.vspdData2
    Call ggoSpread.ClearSpreadData()
    ggoSpread.SSDeleteFlag 1

    If DbQuery = False Then
		Call RestoreToolBar()
		Exit Sub
    End If                                                                     '☜: Query db data

	Set gActiveElement = document.ActiveElement
	
End Sub

'==========================================================================================================
Sub DbDeleteOk()

    On Error Resume Next                                                       '☜: If process fails
    Err.Clear                                                                  '☜: Clear error status

    Set gActiveElement = document.ActiveElement

End Sub

'=========================================================================================================
Function OpenPopUp(Byval strCode, Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
		Case 0
			arrParam(0) = "사업장팝업"				' 팝업 명칭 
			arrParam(1) = "B_BIZ_AREA"		 			' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "사업장코드"				' 조건필드의 라벨 명칭 

			arrField(0) = "BIZ_AREA_CD"	     			' Field명(0)
			arrField(1) = "BIZ_AREA_NM"			    	' Field명(1)

			arrHeader(0) = "사업장코드"				' Header명(0)
			arrHeader(1) = "사업장명"				' Header명(1)

	End Select

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetPopUp(arrRet, iWhere)
	End If
	frm1.txtBizAreaCd.focus
	
End Function

'=========================================================================================================
Function SetPopUp(Byval arrRet, Byval iWhere)

	With frm1
		Select Case iWhere
			Case 0
				.txtBizAreaCd.value = arrRet(0)
				.txtBizAreaNm.value = arrRet(1)

		End Select

	End With

End Function

'=========================================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row )

    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col
    Call CheckMinNumSpread(frm1.vspdData, Col, Row)		

	ggoSpread.Source = frm1.vspdData
	'ggoSpread.UpdateRow Row

End Sub

'=========================================================================================================
Sub vspdData2_Change(ByVal Col, ByVal Row)

    Frm1.vspdData2.Row = Row
    Frm1.vspdData2.Col = Col
    Call CheckMinNumSpread(frm1.vspdData2, Col, Row)		
	
    ggoSpread.Source = frm1.vspdData2
	ggoSpread.UpdateRow Row

End Sub	

'=========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SPC"

	Set gActiveSpdSheet = frm1.vspdData

	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If

	If Row = 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort
			lgSortKey = 2
		Else
			ggoSpread.SSSort ,lgSortKey
			lgSortKey = 1
		End If
		Exit Sub
	End If

	If frm1.vspdData.MaxRows <= 0 Then
		Exit Sub
	End If
	If Row = frm1.vspdData.ActiveRow Then
		Exit Sub
	End If
	If Len(frm1.vspdData.Text) > 0 Then
		frm1.vspddata2.MaxRows = 0			 		
	 	Call DbQuery2(Row)		
	End If	
End Sub

'==========================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0010111111")
	gMouseClickStatus = "SP1C"

	Set gActiveSpdSheet = frm1.vspdData2

	If frm1.vspdData2.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If

	If Row = 0 Then
		ggoSpread.Source = frm1.vspdData2
		If lgSortKey = 1 Then
			ggoSpread.SSSort
			lgSortKey = 2
		Else
			ggoSpread.SSSort ,lgSortKey
			lgSortKey = 1
		End If
		Exit Sub
	End If

End Sub

'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)

	If Col = C_Confirm and lgQueryfg = true Then
		If BUTTONDOWN = 0 or BUTTONDOWN = 1 then
			Call SetVspdData2Checked(Row)
		End If
	End If

End Sub

'=======================================================================================================
Sub SetVspdData2Checked(Byval Row)

	Dim i
	Dim StrConf1
	Dim iCol, iRow
	Dim strChgFlag	
	frm1.vspdData.Row	=	Row
	frm1.vspdData.col	=	C_Confirm
	
	iF Row = 0 Then
		Exit sub
	End If
		
	IF 	frm1.vspdData.text = "1" Then
			If frm1.vspdData2.MaxRows > 0 Then 
				For iRow = 1 To frm1.vspdData2.MaxRows
					frm1.vspdData2.Row = iRow
					frm1.vspdData2.col	= C_BConfirm		
					frm1.vspdData2.text = "1" 
				Next
			End If	
	Else
		If frm1.vspdData2.MaxRows > 0 Then 
			For iRow = 1 To frm1.vspdData2.MaxRows
				frm1.vspdData2.Row = iRow
				frm1.vspdData2.col	= C_BConfirm		
				frm1.vspdData2.text = "0" 
			Next
		End If	
	End If			
		
	'//체크한목록이 있는지 확인 
	
	strChgFlag = False
	For iRow=1 To frm1.vspdData.MaxRows 
		frm1.vspdData.row = iRow
		frm1.vspdData.Col = C_Confirm
		If frm1.vspdData.text = "1" Then
			strChgFlag = True
			Exit for
		End If	
	Next
	
	If strChgFlag = True Then
		lgBlnFlgChgValue = True
	Else
		lgBlnFlgChgValue = False
	End If
	
	
End Sub

'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)

    If Row <= 0 Then
       Exit Sub
	End If

	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If

End Sub

'========================================================================================================
Sub vspdData2_DblClick(ByVal Col, ByVal Row)

    If Row <= 0 Then
       Exit Sub
	End If

	If frm1.vspdData2.MaxRows = 0 Then
		Exit Sub
	End If
End Sub

'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

End Sub

'========================================================================================================
Sub vspdData2_GotFocus()
    ggoSpread.Source = Frm1.vspdData2

End Sub

'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    
End Sub    

'==========================================================================================
Sub vspdData2_MouseDown(Button, Shift, X, Y)

    If Button = 2 And gMouseClickStatus = "SP1C" Then
       gMouseClickStatus = "SP1CR"
    End If
    
End Sub

'========================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)

	Dim indx1, indx2

	frm1.vspdData.Row = Row

	Select Case Col
		Case  C_Conf_Nm
			frm1.vspdData.Col = Col
			indx1 = frm1.vspdData.Value
			frm1.vspdData.Col = C_Confirm
			frm1.vspdData.Value = indx1

			For indx2 = 1 to frm1.vspdData2.MaxRows
				frm1.vspdData2.Row	 = indx2
				frm1.vspdData2.Col   = C_BConfirm
				frm1.vspdData2.Value = indx1
				frm1.vspdData2.Col   = C_BConf_Nm
				frm1.vspdData2.Value = indx1
			Next

	End Select

End Sub

'=========================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

    If Row <> NewRow And NewRow > 0 Then
        ggoSpread.source = frm1.vspdData2
		Call ggoSpread.ClearSpreadData()
        lgCurrRow1 = NewRow

		If Len(frm1.vspdData.Text) > 0 Then
		  frm1.vspddata.Col = 0
		  Call DbQuery2(NewRow)
		End If
    End If

End Sub

'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If

    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
    	If lgPageNo <> "" Then
			Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
    	End If
    End If

End Sub

Sub ChkControl(Byval pVal)
	Dim iLoop, i, iLoopCnt, pObj

	Set pObj = frm1.vspdData
	
	With pObj
	iLoopCnt = .MaxRows

	For iLoop = 1 To iLoopCnt
		.Row = iLoop 
		.Col = C_Confirm
		.Value = pVal
	Next
	End With
End Sub


'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")

End Sub

'========================================================================================================
Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")

End Sub

'========================================================================================================
Sub txtFromTempGlDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery
End Sub

'========================================================================================================
Sub txtToTempGlDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery
End Sub

'========================================================================================================
Sub txtFromTempGlDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromTempGlDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtFromTempGlDt.focus

    End If
End Sub

'========================================================================================================
Sub txtFromTempGlDt_Change()
    lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Sub txttoTempGlDt_DblClick(Button)
    If Button = 1 Then
        frm1.txttoTempGlDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txttoTempGlDt.focus
        
    End If
End Sub

'========================================================================================================
Sub txttoTempGlDt_Change(Button)
    lgBlnFlgChgValue = True
End Sub



'팝업추가>>air
'======================================================================================================
'   Function Name : OpenPopupTempGL()
'   Function Desc : 
'=======================================================================================================
Function OpenPopupTempGL()

	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	
	With frm1.vspdData
	   if .maxrows > 0 Then	
		.Row = .ActiveRow
		.Col = C_TempGlNo

	
		arrParam(0) = Trim(.Text)	'결의전표번호 
		arrParam(1) = ""			'Reference번호 
	   End if	
	End With

'	arrParam(0) = Trim(GetKeyPosVal("A", 1))	'전표번호 
'	arrParam(1) = ""			      
	IsOpenPop = True
    
    iCalledAspName = AskPRAspName("a5130ra1")    
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
End Function

'======================================================================================================
'   Function Name : OpenPopUpgl()
'   Function Desc : 
'=======================================================================================================
Function OpenPopupGL()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName
	
	iCalledAspName = AskPRAspName("a5120ra1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If	

	If IsOpenPop = True Then Exit Function
	
	With frm1.vspdData
	   if .maxrows > 0 Then	
		.Row = .ActiveRow
		.Col = C_BGlNo	
		
		arrParam(0) = Trim(.Text)				'회계전표번호 
		arrParam(1) = ""						'Reference번호 
	   End if	
	End With
	
	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function




</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc"  -->
</HEAD>

<BODY TABINDEX="-1" SCROLL=NO>
<FORM NAME=frm1 TARGET=MyBizASP METHOD=POST>
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>><!-- ' 상위 여백 --></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS=CLSMTABP>
						<TABLE ID=MyTab CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif"><IMG SRC="../../../CShared/image/table/seltab_up_left.gif" WIDTH=9 HEIGHT=23></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN=CENTER CLASS=CLSMTABP><FONT COLOR=WHITE>본지점전표승인</FONT></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN=RIGHT><IMG SRC="../../../CShared/image/table/seltab_up_right.gif" WIDTH=10 HEIGHT=23></TD>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH="60" align=right><A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</A>&nbsp;</td>
					<TD WIDTH="10">&nbsp;</TD>
					<TD WIDTH="60" align=right><A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A>&nbsp;</td>
					<TD WIDTH="10">&nbsp;</TD>					
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS=Tab11>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH="100%"></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS=CLSFLD>
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>결의일자</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> ID=fpDateTime1 NAME=txtFromTempGlDt CLASS=FPDTYYYYMMDD TITLE=FPDATETIME tag="12" ALT="시작일자"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
 										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> ID=fpDateTime2 NAME=txtToTempGlDt CLASS=FPDTYYYYMMDD TITLE=FPDATETIME tag="12" ALT="종료일자"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 NOWRAP>사업장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS=clstxt TYPE=TEXT    NAME=txtBizAreaCd SIZE=10 MAXLENGTH=10 tag="12XXXU" ALT="사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME=btnBizAreaCd ALIGN=TOP TYPE=BUTTON ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCd.Value, 0)">
													     <INPUT TYPE=TEXT ID=txtBizAreaNm NAME=txtBizAreaNm SIZE=25 tag="14X">
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>승인상태</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME=cboConfFg tag="12" STYLE="WIDTH:100px:"></OPTION></SELECT></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH="100%"></TD>
				</TR>
				<TR>
					<TD WIDTH="100%" HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="70%">
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData tag="2" WIDTH="100%" HEIGHT="100%" TITLE=SPREAD> <PARAM NAME=MaxCols VALUE=0><PARAM NAME=MaxRows VALUE=0> </OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD HEIGHT="30%">
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 tag="2" WIDTH="100%" HEIGHT="100%" TITLE=SPREAD> <PARAM NAME=MaxCols VALUE=0><PARAM NAME=MaxRows VALUE=0> </OBJECT>');</SCRIPT></TD>
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
  		<TD WIDTH="100%">
  			<TABLE <%=LR_SPACE_TYPE_30%>>
   				<TR>
   					<TD WIDTH=10>&nbsp;</TD>
   					<TD><BUTTON NAME="button1" CLASS="CLSMBTN" ONCLICK="vbscript:ChkControl(1)" Flag=1>일괄선택</BUTTON>&nbsp;
						<BUTTON NAME="button1" CLASS="CLSMBTN" ONCLICK="vbscript:ChkControl(0)" Flag=1>일괄취소</BUTTON>&nbsp;
   					</TD>
   				</TR>
   			</TABLE> 
  		</TD>
    </TR>
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>>
			<IFRAME NAME=MyBizASP WIDTH="100%" HEIGHT=<%=BizSize%> SRC="../../blank.htm" FRAMEBORDER=0 SCROLLING=NO NORESIZE FRAMESPACING=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS=HIDDEN NAME=txtSpread	tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME=txtMode			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=txtFlgMode		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=txtConfirm		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=txtFormCnt		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=txtHqBrchNo		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=txtTempGLNo		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=txtUpdtUserId	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=txtInsrtUserId	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=txtMaxRows		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=hOrgChangeId	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=hBizAreaCd		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=hFromTempGlDt	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=hToTempGlDt		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=hcboConfFg		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=htxtWorkFg		tag="24" TABINDEX="-1">
</FORM>
<DIV ID=MousePT NAME=MousePT>
	<IFRAME NAME=MouseWindow FRAMEBORDER=0 SCROLLING=NO NORESIZE FRAMESPACING=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
