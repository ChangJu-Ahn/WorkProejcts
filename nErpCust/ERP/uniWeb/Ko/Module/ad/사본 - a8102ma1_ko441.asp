<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : a5101ma1
'*  4. Program Name         : 결의전표미달거래연결 
'*  5. Program Desc         : 결의전표미달거래연결 
'*  6. Component List       : PADG010.dll
'*  7. Modified date(First) : 2000/09/22,2000/10/07
'*  8. Modified date(Last)  : 2003/06/05
'*  9. Modifier (First)     : 안혜진 
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
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js">			</SCRIPT>

<SCRIPT LANGUAGE="vbscript">
Option Explicit
'########################################################################################################
'#                       4.  Data Declaration Part
'========================================================================================================
'=                       4.1 External ASP File
Const BIZ_PGM_ID      = "a8102mb1_ko441.asp"					'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_QRY_ID2 = "a8102mb7_ko441.asp"
'========================================================================================================
'=                       4.2 Constant variables For spreadsheet
Const TAB1 = 1											'☜: Tab의 위치 
Const TAB2 = 2

Const C_SHEETMAXROWS	= 30							' : 한 화면에 보여지는 최대갯수*1.5
'========================================================================================================
'=                       4.3 Common variables 
<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================================
'=                       4.4 User-defind Variables
Dim C_WKCHK											'☆: ITEM SPREAD SHEET 의 COLUMNS 인덱스 
Dim C_TEMPGLDT
Dim C_BIZAREACD
Dim C_BIZAREANM
Dim C_TEMPGLNO
Dim C_DRLOCAMT

Dim C_ITEMSEQ										'☆: DTL SPREAD SHEET 의 COLUMNS 인덱스 
Dim C_ACCTNM
Dim C_DRCRFGNM
Dim C_ITEMLOCAMT

Dim lgintItemCnt
Dim lgSelframeFlg
Dim lsClickRow
Dim lsClickRow2
Dim lsClickRow6
Dim lsClickRow7
Dim IsOpenPop
Dim lgBlnFlgChgValue2

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

	Select Case pOpt
		Case "A", "B", "E", "F"
			C_WKCHK		= 1								'☆: ITEM SPREAD SHEET 의 COLUMNS 인덱스 
			C_TEMPGLDT	= 2
			C_BIZAREACD	= 3
			C_BIZAREANM	= 4
			C_TEMPGLNO	= 5
			C_DRLOCAMT	= 6
		
		Case "C", "D", "G", "H"
			C_ITEMSEQ	= 1								'☆: DTL SPREAD SHEET 의 COLUMNS 인덱스 
			C_ACCTNM	= 2
			C_DRCRFGNM	= 3
			C_ITEMLOCAMT= 4

	End Select

End Sub

'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode		= parent.OPMD_CMODE	
    lgBlnFlgChgValue	= False			
    lgStrPrevKey		= ""			
    lsClickRow			= ""
	lsClickRow2			= ""
	lsClickRow6			= ""
	lsClickRow7			= ""

End Sub
'=========================================================================================================
Sub InitUserVariables()

	lgIntGrpCount	= 0                           'initializes Group View Size
	lgLngCurRows	= 0                           'initializes Deleted Rows Count
	lgSelframeFlg	= TAB1

End Sub
'=========================================================================================================
Sub SetDefaultVal()

	frm1.txtfromdt.text = UniConvDateAToB("<%=lsSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
    frm1.txttodt.text   = UniConvDateAToB("<%=lsSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

End Sub
'========================================================================================================
Sub LoadInfTB19029()

<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "A", "NOCOOKIE", "MA") %>
<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>

End Sub

'========================================================================================================
Sub InitSpreadSheet(ByVal pOpt)

	Call initSpreadPosVariables(pOpt)

	Select Case pOpt
		Case "A"
			ggoSpread.Source = frm1.vspdData
			ggoSpread.SpreadInit "V20021224",,parent.gAllowDragDropSpread
	
			With frm1.vspdData

			    .ReDraw = False
			    
			    .MaxCols	= C_DRLOCAMT + 1
			    Call ggoSpread.ClearSpreadData()
				
				Call GetSpreadColumnPos(pOpt)
				ggoSpread.SSSetCheck C_WKCHK,     "작업",     8,  -10, "", True,	-1
			    ggoSpread.SSSetDate  C_TEMPGLDT,  "결의일",   10,   2, parent.gDateFormat
			    ggoSpread.SSSetEdit  C_BIZAREACD, "",			  10
				ggoSpread.SSSetEdit  C_BIZAREANM, "사업장",   10,    ,   , 20
				ggoSpread.SSSetEdit  C_TEMPGLNO,  "결의번호", 12,   2,   , 30
				ggoSpread.SSSetFloat C_DRLOCAMT,  "금액",     15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
				
				Call ggoSpread.SSSetColHidden(C_BIZAREACD,C_BIZAREACD,True)
				Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)

				.ReDraw = True

			End with
		
		Case "B"
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.SpreadInit "V20021224",,parent.gAllowDragDropSpread

			With frm1.vspdData2

			    .ReDraw = False
			    
			    .MaxCols	= C_DRLOCAMT + 1
			    Call ggoSpread.ClearSpreadData()

				Call GetSpreadColumnPos(pOpt)
				ggoSpread.SSSetCheck C_WKCHK,     "작업",      8, -10, "", True,	-1
			    ggoSpread.SSSetDate  C_TEMPGLDT,  "결의일",   10,   2, parent.gDateFormat
			    ggoSpread.SSSetEdit  C_BIZAREACD, "",			  10
				ggoSpread.SSSetEdit  C_BIZAREANM, "사업장",   10,    ,   ,	20
				ggoSpread.SSSetEdit  C_TEMPGLNO,  "결의번호", 12,   2,   ,	30
				ggoSpread.SSSetFloat C_DRLOCAMT,  "금액",     15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec

				Call ggoSpread.SSSetColHidden(C_BIZAREACD,C_BIZAREACD,True)
				Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)

				.ReDraw = True

			End With
		
		Case "C"
			ggoSpread.Source = frm1.vspdData3
			ggoSpread.SpreadInit "V20021224",,parent.gAllowDragDropSpread

			With frm1.vspdData3

			    .ReDraw = False

			    .MaxCols	= C_ITEMLOCAMT + 1
			    Call ggoSpread.ClearSpreadData()

				Call AppendNumberPlace("6","3","0")
				Call GetSpreadColumnPos(pOpt)
			    ggoSpread.SSSetFloat C_ITEMSEQ,    "순번",	10, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, 2
				ggoSpread.SSSetEdit  C_ACCTNM,     "계정명",	20,    , , 20
				ggoSpread.SSSetEdit  C_DRCRFGNM,   "차/대",	10,   2, , 10
				ggoSpread.SSSetFloat C_ITEMLOCAMT, "금액",	15,	parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec

				Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)

				.ReDraw = True

			End with

		Case "D"
			ggoSpread.Source = frm1.vspdData4
			ggoSpread.SpreadInit "V20021224",,parent.gAllowDragDropSpread

			With frm1.vspdData4

			    .ReDraw = False

			    .MaxCols	= C_ITEMLOCAMT + 1
			    Call ggoSpread.ClearSpreadData()

				Call AppendNumberPlace("6","3","0")
				Call GetSpreadColumnPos(pOpt)
			    ggoSpread.SSSetFloat C_ITEMSEQ,    "순번",	10, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, 2
				ggoSpread.SSSetEdit  C_ACCTNM,     "계정명",	20,    , , 20
				ggoSpread.SSSetEdit  C_DRCRFGNM,   "차/대",	10,   2, , 10
				ggoSpread.SSSetFloat C_ITEMLOCAMT, "금액",	15,	parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec

				Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)

				.ReDraw = True

			End With

		Case "E"
			ggoSpread.Source = frm1.vspdData6
			ggoSpread.SpreadInit "V20021224",,parent.gAllowDragDropSpread

			With frm1.vspdData6

			    .ReDraw = False

			    .MaxCols	= C_DRLOCAMT + 1
			    Call ggoSpread.ClearSpreadData()
				
				Call GetSpreadColumnPos(pOpt)
				ggoSpread.SSSetCheck C_WKCHK,     "작업",      8, -10, "", True, -1
			    ggoSpread.SSSetDate  C_TEMPGLDT,  "결의일",   10,   2, parent.gDateFormat
			    ggoSpread.SSSetEdit  C_BIZAREACD, "",			  10
				ggoSpread.SSSetEdit  C_BIZAREANM, "사업장",   10,    , , 20
				ggoSpread.SSSetEdit  C_TEMPGLNO,  "결의번호", 12,   2, , 30
				ggoSpread.SSSetFloat C_DRLOCAMT,  "금액",     15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec

				Call ggoSpread.SSSetColHidden(C_BIZAREACD,C_BIZAREACD,True)
				Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)

				.ReDraw = True
			            
			End With

		Case "F"
			ggoSpread.Source = frm1.vspdData7
			ggoSpread.SpreadInit "V20021224",,parent.gAllowDragDropSpread

			With frm1.vspdData7

			    .ReDraw = False

			    .MaxCols	= C_DRLOCAMT + 1
			    Call ggoSpread.ClearSpreadData()
				
				Call GetSpreadColumnPos(pOpt)
				ggoSpread.SSSetCheck C_WKCHK,     "작업",      8, -10, "", True, -1
			    ggoSpread.SSSetDate  C_TEMPGLDT,  "결의일",   10,   2, parent.gDateFormat
			    ggoSpread.SSSetEdit  C_BIZAREACD, "",			  10
				ggoSpread.SSSetEdit  C_BIZAREANM, "사업장",   10,    , , 20
				ggoSpread.SSSetEdit  C_TEMPGLNO,  "결의번호", 12,   2, , 30
				ggoSpread.SSSetFloat C_DRLOCAMT,  "금액",     15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec

				Call ggoSpread.SSSetColHidden(C_WKCHK,C_WKCHK,True)
				Call ggoSpread.SSSetColHidden(C_BIZAREACD,C_BIZAREACD,True)
				Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)

				.ReDraw = True

			End With

		Case "G"
			ggoSpread.Source = frm1.vspdData8
			ggoSpread.SpreadInit "V20021224",,parent.gAllowDragDropSpread
	
			With frm1.vspdData8

				.ReDraw = False
				
			    .MaxCols	= C_ITEMLOCAMT + 1
			    Call ggoSpread.ClearSpreadData()

				Call AppendNumberPlace("6","3","0")
				Call GetSpreadColumnPos(pOpt)
			    ggoSpread.SSSetFloat C_ITEMSEQ,    "순번",   10, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, 2
				ggoSpread.SSSetEdit  C_ACCTNM,     "계정명", 20, , , 20
				ggoSpread.SSSetEdit  C_DRCRFGNM,   "차/대",  10, 2, , 10
				ggoSpread.SSSetFloat C_ITEMLOCAMT, "금액",   15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec

				Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)

				.ReDraw = True

			End With

		Case "H"
			ggoSpread.Source = frm1.vspdData9
			ggoSpread.SpreadInit "V20021224",,parent.gAllowDragDropSpread

			With frm1.vspdData9

				.ReDraw = False

			    .MaxCols	= C_ITEMLOCAMT + 1
			    Call ggoSpread.ClearSpreadData()

				Call AppendNumberPlace("6","3","0")
				Call GetSpreadColumnPos(pOpt)
			    ggoSpread.SSSetFloat C_ITEMSEQ,    "순번",   10, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,2
				ggoSpread.SSSetEdit  C_ACCTNM,     "계정명", 20, ,  , 20
				ggoSpread.SSSetEdit  C_DRCRFGNM,   "차/대",  10, 2, , 10
				ggoSpread.SSSetFloat C_ITEMLOCAMT, "금액",   15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec

				Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)

				.ReDraw = True

			End With

	End Select
	
	Call SetSpreadLock("Q", pOpt, -1, -1)

End Sub
'=======================================================================================================
Sub SetSpreadLock(Byval stsFg, Byval pOpt, ByVal lRow  , ByVal lRow2 )

	Select Case pOpt
		Case "A"
			ggoSpread.Source = frm1.vspdData
			With frm1.vspdData
				.ReDraw = False
				ggoSpread.SpreadLockWithOddEvenRowColor()
				'ggoSpread.SpreadLock C_TEMPGLDT, lRow, C_DRLOCAMT, lRow2
'				ggoSpread.SpreadLock C_TEMPGLDT,	-1  , C_TEMPGLDT
'				ggoSpread.SpreadLock C_BIZAREACD,	-1  , C_BIZAREACD
'				ggoSpread.SpreadLock C_BIZAREANM,	-1  , C_BIZAREANM
'				ggoSpread.SpreadLock C_TEMPGLNO,	-1  , C_TEMPGLNO
'				ggoSpread.SpreadLock C_DRLOCAMT,	-1  , C_DRLOCAMT
				.ReDraw = True
			End With
			
		Case "B"
			ggoSpread.Source = frm1.vspdData2
			With frm1.vspdData2
				.ReDraw = False
				ggoSpread.SpreadLockWithOddEvenRowColor()				
				'ggoSpread.SpreadLock C_TEMPGLDT, lRow, C_DRLOCAMT, lRow2
'				ggoSpread.SpreadLock C_TEMPGLDT,	-1  , C_TEMPGLDT
'				ggoSpread.SpreadLock C_BIZAREACD,	-1  , C_BIZAREACD
'				ggoSpread.SpreadLock C_BIZAREANM,	-1  , C_BIZAREANM
'				ggoSpread.SpreadLock C_TEMPGLNO,	-1  , C_TEMPGLNO
'				ggoSpread.SpreadLock C_DRLOCAMT,	-1  , C_DRLOCAMT
				.ReDraw = True
			End With
		
		Case "C"
			ggoSpread.Source = frm1.vspdData3
			With frm1.vspdData3
				.ReDraw = False
				ggoSpread.SpreadLockWithOddEvenRowColor()								
				'ggoSpread.SpreadLock C_ITEMSEQ, lRow, C_ITEMLOCAMT, lRow2
'				ggoSpread.SpreadLock C_ITEMSEQ,		-1  , C_ITEMSEQ
'				ggoSpread.SpreadLock C_ACCTNM,		-1  , C_ACCTNM
'				ggoSpread.SpreadLock C_DRCRFGNM,	-1  , C_DRCRFGNM
'				ggoSpread.SpreadLock C_ITEMLOCAMT,	-1  , C_ITEMLOCAMT				
				.ReDraw = True
			End With

		Case "D"
			ggoSpread.Source = frm1.vspdData4
			With frm1.vspdData4
				.ReDraw = False
				ggoSpread.SpreadLockWithOddEvenRowColor()								
				'ggoSpread.SpreadLock C_ITEMSEQ, lRow, C_ITEMLOCAMT, lRow2
'				ggoSpread.SpreadLock C_ITEMSEQ,		-1  , C_ITEMSEQ
'				ggoSpread.SpreadLock C_ACCTNM,		-1  , C_ACCTNM
'				ggoSpread.SpreadLock C_DRCRFGNM,	-1  , C_DRCRFGNM
'				ggoSpread.SpreadLock C_ITEMLOCAMT,	-1  , C_ITEMLOCAMT
				.ReDraw = True
			End With

		Case "E"
			ggoSpread.Source = frm1.vspdData6
			With frm1.vspdData6
				.ReDraw = False
				ggoSpread.SpreadLockWithOddEvenRowColor()								
				'ggoSpread.SpreadLock C_TEMPGLDT, lRow, C_DRLOCAMT, lRow2
'				ggoSpread.SpreadLock C_TEMPGLDT,	-1  , C_TEMPGLDT
'				ggoSpread.SpreadLock C_BIZAREACD,	-1  , C_BIZAREACD
'				ggoSpread.SpreadLock C_BIZAREANM,	-1  , C_BIZAREANM
'				ggoSpread.SpreadLock C_TEMPGLNO,	-1  , C_TEMPGLNO
'				ggoSpread.SpreadLock C_DRLOCAMT,	-1  , C_DRLOCAMT
				.ReDraw = True
			End With

		Case "F"
			ggoSpread.Source = frm1.vspdData7
			With frm1.vspdData7
				.ReDraw = False
				ggoSpread.SpreadLockWithOddEvenRowColor()								
				'ggoSpread.SpreadLock C_TEMPGLDT, lRow, C_DRLOCAMT, lRow2
'				ggoSpread.SpreadLock C_TEMPGLDT,	-1  , C_TEMPGLDT
'				ggoSpread.SpreadLock C_BIZAREACD,	-1  , C_BIZAREACD
'				ggoSpread.SpreadLock C_BIZAREANM,	-1  , C_BIZAREANM
'				ggoSpread.SpreadLock C_TEMPGLNO,	-1  , C_TEMPGLNO
'				ggoSpread.SpreadLock C_DRLOCAMT,	-1  , C_DRLOCAMT
				.ReDraw = True
			End With

		Case "G"
			ggoSpread.Source = frm1.vspdData8
			With frm1.vspdData8
				.ReDraw = False
				ggoSpread.SpreadLockWithOddEvenRowColor()								
				'ggoSpread.SpreadLock C_ITEMSEQ, lRow, C_ITEMLOCAMT, lRow2
'				ggoSpread.SpreadLock C_ITEMSEQ,		-1  , C_ITEMSEQ
'				ggoSpread.SpreadLock C_ACCTNM,		-1  , C_ACCTNM
'				ggoSpread.SpreadLock C_DRCRFGNM,	-1  , C_DRCRFGNM
'				ggoSpread.SpreadLock C_ITEMLOCAMT,	-1  , C_ITEMLOCAMT
				.ReDraw = True
			End With

		Case "H"
			ggoSpread.Source = frm1.vspdData9
			With frm1.vspdData9
				.ReDraw = False
				ggoSpread.SpreadLockWithOddEvenRowColor()								
				'ggoSpread.SpreadLock C_ITEMSEQ, lRow, C_ITEMLOCAMT, lRow2
'				ggoSpread.SpreadLock C_ITEMSEQ,		-1  , C_ITEMSEQ
'				ggoSpread.SpreadLock C_ACCTNM,		-1  , C_ACCTNM
'				ggoSpread.SpreadLock C_DRCRFGNM,	-1  , C_DRCRFGNM
'				ggoSpread.SpreadLock C_ITEMLOCAMT,	-1  , C_ITEMLOCAMT
				.ReDraw = True
			End With
	End Select
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
              Frm1.vspdData.Action = 0 ' go to
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
			C_WKCHK     = iCurColumnPos(1)
			C_TEMPGLDT  = iCurColumnPos(2)
			C_BIZAREACD = iCurColumnPos(3)
			C_BIZAREANM = iCurColumnPos(4)
			C_TEMPGLNO  = iCurColumnPos(5)
			C_DRLOCAMT  = iCurColumnPos(6)

		Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_WKCHK     = iCurColumnPos(1)
			C_TEMPGLDT  = iCurColumnPos(2)
			C_BIZAREACD = iCurColumnPos(3)
			C_BIZAREANM = iCurColumnPos(4)
			C_TEMPGLNO  = iCurColumnPos(5)
			C_DRLOCAMT  = iCurColumnPos(6)

		Case "C"
            ggoSpread.Source = frm1.vspdData3
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_ITEMSEQ    = iCurColumnPos(1)
			C_ACCTNM     = iCurColumnPos(2)
			C_DRCRFGNM   = iCurColumnPos(3)
			C_ITEMLOCAMT = iCurColumnPos(4)

		Case "D"
            ggoSpread.Source = frm1.vspdData4
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_ITEMSEQ    = iCurColumnPos(1)
			C_ACCTNM     = iCurColumnPos(2)
			C_DRCRFGNM   = iCurColumnPos(3)
			C_ITEMLOCAMT = iCurColumnPos(4)

		Case "E"
            ggoSpread.Source = frm1.vspdData6
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_WKCHK     = iCurColumnPos(1)
			C_TEMPGLDT  = iCurColumnPos(2)
			C_BIZAREACD = iCurColumnPos(3)
			C_BIZAREANM = iCurColumnPos(4)
			C_TEMPGLNO  = iCurColumnPos(5)
			C_DRLOCAMT  = iCurColumnPos(6)

		Case "F"
            ggoSpread.Source = frm1.vspdData7
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_WKCHK     = iCurColumnPos(1)
			C_TEMPGLDT  = iCurColumnPos(2)
			C_BIZAREACD = iCurColumnPos(3)
			C_BIZAREANM = iCurColumnPos(4)
			C_TEMPGLNO  = iCurColumnPos(5)
			C_DRLOCAMT  = iCurColumnPos(6)

		Case "G"
            ggoSpread.Source = frm1.vspdData8
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_ITEMSEQ    = iCurColumnPos(1)
			C_ACCTNM     = iCurColumnPos(2)
			C_DRCRFGNM   = iCurColumnPos(3)
			C_ITEMLOCAMT = iCurColumnPos(4)

		Case "H"
            ggoSpread.Source = frm1.vspdData9
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_ITEMSEQ    = iCurColumnPos(1)
			C_ACCTNM     = iCurColumnPos(2)
			C_DRCRFGNM   = iCurColumnPos(3)
			C_ITEMLOCAMT = iCurColumnPos(4)

    End Select

End Sub

'=======================================================================================================
Function ClickTab1()

	If lgSelframeFlg = TAB1 Then Exit Function

	Call changeTabs(TAB1)										'~~~ 첫번째 Tab 
	lgSelframeFlg = TAB1

	If lgIntFlgMode <> parent.OPMD_UMODE Then
	    Call SetToolbar("1100100000011111")			'⊙: 버튼 툴바 제어 
	Else
	    Call SetToolbar("1100100000011111")			'⊙: 버튼 툴바 제어 
	End If

End Function

'=======================================================================================================
Function ClickTab2()

	If lgSelframeFlg = TAB2 Then Exit Function

	Call changeTabs(TAB2)										'~~~ 첫번째 Tab 
	lgSelframeFlg = TAB2

End Function

'========================================================================================================
'                        5.2 Common Group-2
'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                             '⊙: Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)

    Call ggoOper.ClearField(Document, "1")							'⊙: Condition field clear
	Call ggoOper.LockField(Document, "N")							'⊙: Lock  Suitable  Field
    Call InitSpreadSheet("A")                                       '⊙: Setup the Spread sheet
    Call InitSpreadSheet("B")                                       '⊙: Setup the Spread sheet
    Call InitSpreadSheet("C")                                       '⊙: Setup the Spread sheet
    Call InitSpreadSheet("D")                                       '⊙: Setup the Spread sheet
    Call InitSpreadSheet("E")                                       '⊙: Setup the Spread sheet
    Call InitSpreadSheet("F")                                       '⊙: Setup the Spread sheet
    Call InitSpreadSheet("G")                                       '⊙: Setup the Spread sheet
    Call InitSpreadSheet("H")                                       '⊙: Setup the Spread sheet
    Call InitVariables()                                            '⊙: Initializes local global variables
	Call InitUserVariables()
    
    Call SetDefaultVal()
    Call SetToolbar("1100000000001111")							    '⊙: 버튼 툴바 제어 

    frm1.txtCommandMode.value = "CREATE"

	gIsTab     = "Y"												'Tab 있는 화면 
	gTabMaxCnt = 2													'Tab 갯수 

	frm1.txtBizArea.Focus

End Sub
'========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'========================================================================================================
Function FncQuery()

    Dim IntRetCD
    Dim RetFlag

	on Error Resume Next
	Err.Clear																		'☜: Protect system from crashing

    FncQuery = False																'⊙: Processing is NG

    ggoSpread.Source = frm1.vspdData
     
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")		         '데이타가 변경되었습니다. 조회하시겠습니까?
    	If IntRetCD = vbNo Then
      		Exit Function
     	End If
    End If

    If Not chkField(Document, "1") Then												'⊙: This function check indispensable field
       Exit Function
    End If

	If CompareDateByFormat(frm1.txtFromDt.text, frm1.txtToDt.text, frm1.txtFromDt.Alt, frm1.txtToDt.Alt, _
                        "970025", frm1.txtFromDt.UserDefinedFormat, parent.gComDateType, True) = False Then
    	Exit Function
    End If

    If UNICDbl(frm1.txtFromAmt.Text) > UNICDbl(frm1.txtToAmt.Text) Then
		IntRetCD = DisplayMsgBox("113123", "X", "X", "X")					        '⊙: "Will you destory previous data"
		Exit Function
    End If

    If frm1.txtBizArea.value = "" Then
		frm1.txtBizAreaNm.value = ""
    End If

    Call InitSpreadSheet("A")
    Call InitSpreadSheet("B")	
    Call InitSpreadSheet("C")
    Call InitSpreadSheet("D")
    Call InitSpreadSheet("E")
    Call InitSpreadSheet("F")
    Call InitSpreadSheet("G")
    Call InitSpreadSheet("H")
    Call InitVariables()	

    If  DbQuery	= False Then														'☜: Query db data
		Exit Function
    End If

    If Err.number = 0 Then
       FncQuery = True	
    End If

	Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function FncNew()

	Dim IntRetCD

	On Error Resume Next
    Err.Clear			
    
    FncNew = False		

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")		         '데이타가 변경되었습니다. 조회하시겠습니까?
    	If IntRetCD = vbNo Then
      		Exit Function
    	End If
    End If

    Call ggoOper.ClearField(Document, "1")									        '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
	Call ggoSpread.ClearSpreadData()
	ggoSpread.Source = frm1.vspdData2
	Call ggoSpread.ClearSpreadData()
	ggoSpread.Source = frm1.vspdData3
	Call ggoSpread.ClearSpreadData()
	ggoSpread.Source = frm1.vspdData4
	Call ggoSpread.ClearSpreadData()
	ggoSpread.Source = frm1.vspdData6
	Call ggoSpread.ClearSpreadData()
	ggoSpread.Source = frm1.vspdData7
	Call ggoSpread.ClearSpreadData()
	ggoSpread.Source = frm1.vspdData8
	Call ggoSpread.ClearSpreadData()
	ggoSpread.Source = frm1.vspdData9
	Call ggoSpread.ClearSpreadData()
	
    Call ggoOper.LockField(Document, "N")									        '⊙: Lock  Suitable  Field
    Call InitVariables()															'⊙: Initializes local global variables
    Call SetDefaultVal()
    Call InitSpreadSheet("A")														'⊙: Setup the Spread sheet
    Call InitSpreadSheet("B")														'⊙: Setup the Spread sheet
    Call InitSpreadSheet("C")														'⊙: Setup the Spread sheet
    Call InitSpreadSheet("D")														'⊙: Setup the Spread sheet
    Call InitSpreadSheet("E")														'⊙: Setup the Spread sheet
    Call InitSpreadSheet("F")														'⊙: Setup the Spread sheet
    Call InitSpreadSheet("G")														'⊙: Setup the Spread sheet
    Call InitSpreadSheet("H")														'⊙: Setup the Spread sheet

    frm1.txtCommandMode.value = "CREATE"
    Call SetToolbar("1100100000011111")	

    If Err.number = 0 Then
		FncNew = True							            						'⊙: Processing is OK
    End If

	Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function FncDelete()

	Dim IntRetCD 

    On Error Resume Next	
    Err.Clear		

    FncDelete = False		

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900002", parent.VB_YES_NO, "X", "X")		        '데이타가 변경되었습니다. 조회하시겠습니까?
    	If IntRetCD = vbNo Then
      		Exit Function
    	End If
    End If

    If DbDelete = False Then
		Exit Function
	End If																            '☜: Delete db data

	If Err.number = 0 Then
		FncDelete = True
	End If

	Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function FncSave()

    Dim IntRetCD

    On Error Resume Next
    Err.Clear																		'☜: Protect system from crashing

    FncSave = False
    If 	lgBlnFlgChgValue =False And lgBlnFlgChgValue2 = False Then
		IntRetCD = DisplayMsgBox("900001", "X", "X", "X")			       'No data changed!!
		Exit Function    
    End if

    If DbSave = False Then												           '☜: Save db data
		Exit Function
	End If

	If Err.number  = 0 Then
		FncSave = True
    End If

	Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function FncCopy()

	Dim  IntRetCD

    On Error Resume Next												            '☜: Protect system from crashing
    Err.Clear

	FncCopy = False

	frm1.vspdData.ReDraw = False

    ggoSpread.Source = frm1.vspdData
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow
    MaxSpreadVal frm1.vspdData.ActiveRow

	If Err.number = 0 Then
		FncCopy = True
	End If

	Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function FncCancel()

    Dim iItemSeq

    On Error Resume Next	
    Err.Clear

	FncCancel = False

    With frm1.vspdData

        .Row = .ActiveRow
        .Col = 0
        If .Text = ggoSpread.InsertFlag Then
            .Col = C_ITEMSEQ
        End if

        .Col = C_WKCHK
		iItemSeq = .Text

        ggoSpread.Source = frm1.vspdData
        ggoSpread.EditUndo

        .Row = .ActiveRow
        .Col = 0
        If .Text = ggoSpread.InsertFlag Then
            .Col = C_ITEMSEQ
            frm1.htempglno.value = .Text
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.ClearSpreadData()
        Else
            .Col = C_ITEMSEQ
            frm1.htempglno.value = .Text
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.ClearSpreadData()
        End if

    End With
    
    If Err.number = 0 Then
		FncCancel = True
	End If

	Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function FncInsertRow()

    On Error Resume Next	
    Err.Clear

    FncInsertRow = False

	With frm1

		If .vspdData.MaxRows = 50 Then  Exit Function

		.vspdData.focus
		ggoSpread.Source = .vspdData
		.vspdData.ReDraw = False
		ggoSpread.InsertRow
		.vspdData.ReDraw = True
		SetSpreadColor .vspdData.ActiveRow
		MaxSpreadVal .vspdData.ActiveRow

    End With

	If Err.number = 0 Then
		FncInsertRow = True
	End If

	Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function FncPrint()

    On Error Resume Next  
    Err.Clear                                                                     '☜: Clear error status

    FncPrint = False                                                              '☜: Processing is NG
	Call Parent.FncPrint()                                                        '☜: Protect system from crashing

    If Err.number = 0 Then
       FncPrint = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function FncPrev() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrev = False                                                               '☜: Processing is NG
    If Err.number = 0 Then
       FncPrev = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function FncNext() 

    On Error Resume Next      
    Err.Clear       

    FncNext = False         
    If Err.number = 0 Then
       FncNext = True     
    End If

    Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function FncExcel() 

    On Error Resume Next
    Err.Clear           

    FncExcel = False    

	Call Parent.FncExport(Parent.C_MULTI)

    If Err.number = 0 Then
       FncExcel = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function


'========================================================================================================
Function FncFind() 

    On Error Resume Next 
    Err.Clear    

    FncFind = False         

	Call Parent.FncFind(Parent.C_MULTI, True)

    If Err.number = 0 Then
       FncFind = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'========================================================================================================
Function FncExit()

	Dim IntRetCD

    On Error Resume Next  
    Err.Clear     

	FncExit = False

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
	IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")	
    	If IntRetCD = vbNo Then
      		Exit Function
    	End If
    End If

    ggoSpread.Source = frm1.vspdData2
    If ggoSpread.SSCheckChange = True Then
	IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")	
    	If IntRetCD = vbNo Then
      		Exit Function
    	End If
    End If

    ggoSpread.Source = frm1.vspdData6
    If ggoSpread.SSCheckChange = True Then
	IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")	
    	If IntRetCD = vbNo Then
      		Exit Function
    	End If
    End If

    If Err.number = 0 Then
       FncExit = True           
    End If

    Set gActiveElement = document.ActiveElement

End Function

'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
Sub PopRestoreSpreadColumnInf()

	Dim strSpread

    ggoSpread.Source = gActiveSpdSheet
    Select Case UCase(ggoSpread.Source.Name)
		Case "VSPDDATA"
			strSpread = "A"
		Case "VSPDDATA2"
			strSpread = "B"
		Case "VSPDDATA3"
			strSpread = "C"
		Case "VSPDDATA4"
			strSpread = "D"
		Case "VSPDDATA6"
			strSpread = "E"
		Case "VSPDDATA7"
			strSpread = "F"
		Case "VSPDDATA8"
			strSpread = "G"
		Case "VSPDDATA9"
			strSpread = "H"
	End Select

    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet(strSpread)
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()

End Sub

'========================================================================================================
'                        5.3 Common Group-3
'========================================================================================================
Function DbQuery()

	Dim strVal
	Dim RetFlag

	On Error Resume Next
	Err.Clear

    DbQuery = False

	Call DisableToolBar(parent.TBC_QUERY)
    Call LayerShowHide(1)
    
	ggoSpread.Source = frm1.vspdData3
	Call ggoSpread.ClearSpreadData()
	
    lsClickRow  = 1
    lsClickRow2 = 1

    With frm1

		strVal = BIZ_PGM_ID & "?txtMode="		& parent.UID_M0001 
		strVal = strVal		& "&txtBizArea="	& Trim(.txtBizArea.value)
		strVal = strVal		& "&txtFromdt="		& Trim(.txtFromdt.Text)	 
		strVal = strVal		& "&txtTodt="		& Trim(.txtTodt.Text)	
		strVal = strVal		& "&txtfromamt="	& Trim(.txtfromamt.Text)	
		strVal = strVal		& "&txttoamt="		&  Trim(.txttoamt.Text)	
		strVal = strVal		& "&lgStrPrevKey="	& lgStrPrevKey
		strVal = strVal		& "&txtMaxRows1="	& .vspdData.MaxRows
		strVal = strVal		& "&txtMaxRows2="	& .vspdData2.MaxRows
		strval = strVal		& "&txttab="		& lgSelframeFlg

    End With
    
    Call RunMyBizASP(MyBizASP, strVal)		                                  '☜: 비지니스 ASP 를 가동 
    
	If Err.number = 0 Then
       DbQuery = True														  '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function DbQuery2(Byval strtxttab,ByVal Row, Byval iWhere )

	Dim strVal
	Dim boolExist
	Dim lngRows

	On Error Resume Next
	Err.Clear

	DbQuery2 = False

	Call DisableToolBar(parent.TBC_QUERY)
	Call LayerShowHide(1)
	boolExist = False

	With frm1
	
		If strtxttab = "1" Then
			Select Case iWhere
				Case 1	
					ggoSpread.Source = frm1.vspdData
					Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
					Call GetSpreadColumnPos("A")
							
					.vspdData.Row	 = Row
					.vspdData.Col	 = C_TEMPGLNO
					.htempglno.Value = .vspdData.Text
					'msgbox "1=" & .vspdData.Text
					If Trim(.htempglno.Value) = ""           Then Exit Function

				    If lgIntFlgMode = parent.OPMD_UMODE Then
				       '@Query_Hidden
				       strVal = BIZ_PGM_QRY_ID2 & "?txtMode="		& parent.UID_M0001		'☜: 
				       strVal = strVal			& "&txttempglno="	& .htempglno.Value 		'☆: 조회 조건 데이타 
				       strVal = strVal			& "&txtgubun="		& cstr(iWhere)
				       strVal = strVal			& "&txtstrtab="		& strtxttab
		 		       strVal = strVal			& "&txtMaxRows="	& .vspdData3.MaxRows
				    Else
				       '@Query_Text
				       strVal = BIZ_PGM_QRY_ID2 & "?txtMode="		& parent.UID_M0001		'☜: 
				       strVal = strVal			& "&txttempglno="	& .vspdData.Text		'☆: 조회 조건 
				       strVal = strVal			& "&txtgubun="		& cstr(iWhere)
				       strVal = strVal			& "&txtstrtab="		& strtxttab
				       strVal = strVal			& "&txtMaxRows="	& .vspdData3.MaxRows
				    End If

				    If lsClickRow <> Row Then
						lsClickRow		 = Row
						.vspdData.Col	 = C_WKCHK
						.vspdData.Action = 0
					End If
					
			   Case 2
					ggoSpread.Source = frm1.vspdData3
					Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
					Call GetSpreadColumnPos("B")
			       .vspdData2.Row	= Row
				   .vspdData2.Col	= C_TEMPGLNO
				   .htempglno.Value = .vspdData2.Text
					If Trim(.htempglno.Value) = ""           Then Exit Function

				    If lgIntFlgMode = parent.OPMD_UMODE Then
				       '@Query_Hidden
				       strVal = BIZ_PGM_QRY_ID2 & "?txtMode="		& parent.UID_M0001		'☜: 
				       strVal = strVal			& "&txttempglno="	& .htempglno.Value		'☆: 조회 조건 데이타 
				       strVal = strVal			& "&txtgubun="		& cstr(iWhere)
				       strVal = strVal			& "&txtstrtab="		& strtxttab
		 		       strVal = strVal			& "&txtMaxRows="	& .vspdData4.MaxRows
				    Else
				       '@Query_Text
				       strVal = BIZ_PGM_QRY_ID2 & "?txtMode="		& parent.UID_M0001		'☜: 
				       strVal = strVal			& "&txttempglno="	& .htempglno.Value		'☆: 조회 조건 
				       strVal = strVal			& "&txtgubun="		& cstr(iWhere)
				       strVal = strVal			& "&txtstrtab="		& strtxttab
				       strVal = strVal			& "&txtMaxRows="	& .vspdData4.MaxRows
				    End If
				
				    If lsClickRow2 <> Row Then
						lsClickRow2		  = Row
						.vspdData2.Col	  = C_WKCHK
						.vspdData2.Action = 0
					End If

			End Select
		Else
			Select Case iWhere
				Case 1
					ggoSpread.Source = frm1.vspdData6
					Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
					Call GetSpreadColumnPos("E")
					.vspdData6.row	 = Row
					.vspdData6.col	 = C_TEMPGLNO
					.htempglno.Value = .vspdData6.Text

					If Trim(.htempglno.Value) = ""           Then Exit Function

				    If lgIntFlgMode = parent.OPMD_UMODE Then
						'@Query_Hidden
						strVal = BIZ_PGM_QRY_ID2 & "?txtMode="		& parent.UID_M0001			'☜:
						strVal = strVal			 & "&txttempglno="	& .htempglno.Value			'☆: 조회 조건 데이타 
					   	strVal = strVal			 & "&txtgubun="		& cstr(iWhere)
 				        strVal = strVal			 & "&txtstrtab="	& strtxttab
				        strVal = strVal			 & "&txtMaxRows="	& .vspdData8.MaxRows
					Else
					    '@Query_Text
					    strVal = BIZ_PGM_QRY_ID2 & "?txtMode="		& parent.UID_M0001			'☜:
					    strVal = strVal			 & "&txttempglno="	& .htempglno.Value			'☆: 조회 조건 
					    strVal = strVal			 & "&txtgubun="		& cstr(iWhere)
 				        strVal = strVal			 & "&txtstrtab="	& strtxttab
					    strVal = strVal			 & "&txtMaxRows="	& .vspdData8.MaxRows
					End If

					If lsClickRow6 <> Row Then
						lsClickRow6		  = Row
						.vspdData6.Col	  = C_WKCHK
						.vspdData6.Action = 0
					End If

				Case 2
					ggoSpread.Source = frm1.vspdData6
					Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
					Call GetSpreadColumnPos("F")
				    .vspdData7.row	 = Row
					.vspdData7.col	 = C_TEMPGLNO
					.htempglno.Value = .vspdData7.Text

					If Trim(.htempglno.Value) = ""           Then Exit Function

				    If lgIntFlgMode = parent.OPMD_UMODE Then
						'@Query_Hidden
					    strVal = BIZ_PGM_QRY_ID2 & "?txtMode="		& parent.UID_M0001			'☜:
					    strVal = strVal			 & "&txttempglno=" & .htempglno.Value			'☆: 조회 조건 데이타 
					    strVal = strVal			 & "&txtgubun="		& cstr(iWhere)
 				        strVal = strVal			 & "&txtstrtab="	& strtxttab
		 			    strVal = strVal			 & "&txtMaxRows="	& .vspdData9.MaxRows
					Else
					    '@Query_Text
					    strVal = BIZ_PGM_QRY_ID2 & "?txtMode="		& parent.UID_M0001			'☜:
					    strVal = strVal			 & "&txttempglno="	& .htempglno.Value			'☆: 조회 조건 
					    strVal = strVal			 & "&txtgubun="		& cstr(iWhere)
 				        strVal = strVal			 & "&txtstrtab="	& strtxttab
 				        strVal = strVal			 & "&txtMaxRows="	& .vspdData9.MaxRows
					End If

					If lsClickRow7 <> Row Then
						lsClickRow7		 = Row
						.vspdData7.Col	 = C_WKCHK
						.vspdData7.Action = 0
					End If

			End Select
		End if

	End With

	strVal = strVal & "&txtquerychk=" & "N"

	Call RunMyBizASP(MyBizASP, strVal)															'☜: 비지니스 ASP 를 가동 

	If Err.number = 0 Then
       DbQuery2 = True														                     '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function DbQuery6(Byval strtxttab, Byval strgubun, byval strtempglno )

	Dim strval
	
	On Error Resume Next
	Err.Clear

	DbQuery6 = False

	Call DisableToolBar(parent.TBC_QUERY)
	Call LayerShowHide(1)
	
	strVal = BIZ_PGM_QRY_ID2 & "?txtMode="			& parent.UID_M0001	'☜: 
	strVal = strVal			 & "&txttempglno="		& strtempglno		'☆: 조회 조건 
	strVal = strVal			 & "&txtgubun="			& strgubun
	strVal = strVal			 & "&txtstrtab="		& strtxttab
	strVal = strVal			 & "&txtquerychk="		& "Y"

	Call RunMyBizASP(MyBizASP, strVal)								     '☜: 비지니스 ASP 를 가동 

	If Err.number = 0 Then
       DbQuery6 = True
    End If

    Set gActiveElement = document.ActiveElement

End function

'========================================================================================================
Function DbSave()

    Dim pAP010M
    Dim lngRows, itemRows, lngLChkRows, lngRChkRows
    Dim lGrpcnt
    DIM strVal
    Dim strDel
    Dim tempItemSeq
    Dim IntRetCD

    On Error Resume Next
	Err.Clear 

    DbSave = False

	Call DisableToolBar(parent.TBC_SAVE)
    Call LayerShowHide(1)

    Call SetSumItem()

    With frm1

		.txtUpdtUserId.value = parent.gUsrID
		.txtMode.value = parent.UID_M0002
		.txtstrtab.value = cstr(lgSelframeFlg)

	End With

    strVal = ""
    lGrpCnt = 1

    ggoSpread.Source = frm1.vspdData

	If lgSelframeFlg = TAB1 Then

		lngLChkRows = 0
		lngRChkRows = 0

        With frm1.vspdData

			For lngRows = 1 to .MaxRows
				.Row = lngRows
				.Col = C_WKCHK
				If Trim(.Text) = "1"  Then
					lnglChkRows = lnglChkRows + 1
					.Col = C_TEMPGLNO	'1
					strVal = strVal & Trim(.Text) & parent.gColSep & parent.gRowSep
					lGrpCnt = lGrpCnt + 1
				End If
			Next

		End With

		ggoSpread.Source = frm1.vspdData2

		With frm1.vspdData2

			For lngRows = 1 to .MaxRows
				.Row = lngRows
				.Col = C_WKCHK
				If Trim(.Text) = "1"  Then
					lngRChkRows = lngRChkRows + 1
					.Col = C_TEMPGLNO	'1
					strVal = strVal & Trim(.Text) & parent.gColSep & parent.gRowSep
					lGrpCnt = lGrpCnt + 1
				End If
			 Next

        End With

        If lnglChkRows = 0 Or lngRchkRows = 0 Then
			IntRetCD = DisplayMsgBox("113124", "X", "X", "X")				'⊙: "Will you destory previous data"
			Call LayerShowHide(0)
			Exit Function
        End If
	Else
		lngLChkRows = 0

		With frm1.vspdData6

			For lngRows = 1 to .MaxRows
				.Row = lngRows
				.Col = C_WKCHK
				
				If Trim(.Text) = "1"  Then
					lngLChkRows = lngLChkRows + 1
				   .Col = C_TEMPGLNO	'1
				   strVal = strVal & Trim(.Text) & parent.gColSep & parent.gRowSep
				   lGrpCnt = lGrpCnt + 1
				End If
			Next
		End With

        If lngLChkRows = 0 Then
			IntRetCD = DisplayMsgBox("113118", "X", "X", "X")				'⊙: "Will you destory previous data"
			Call LayerShowHide(0)
			Exit Function
        End If
	End if

    frm1.txtMaxRows.value = lGrpCnt-1										'Spread Sheet의 변경된 최대갯수 
    frm1.txtSpread.value  = strVal											'Spread Sheet 내용을 저장 

    Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'저장 비지니스 ASP 를 가동 

	If Err.number = 0 Then
       DbSave = True														'☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function DbDelete()

	Dim strVal

	On Error Resume Next
    Err.Clear

    DbDelete = False														   '⊙: Processing is NG

    Call DisableToolBar(parent.TBC_DELETE)
    Call LayerShowHide(1)

	frm1.txtOrgChangeId.value = parent.gChangeOrgId

	strVal = BIZ_PGM_ID & "?txtMode="		 & parent.UID_M0003					'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal		& "&txtDeptCd="		 & Trim(frm1.txtDeptCd.value)
	strVal = strVal		& "&txtOrgChangeId=" & Trim(frm1.txtOrgChangeId.value)
    strVal = strVal		& "&txtGlinputType=" & Trim(frm1.txtGlinputType.value)
	Call RunMyBizASP(MyBizASP, strVal)

    If Err.number = 0 Then
       DbDelete = True                                                           '☜: Processing is OK
    End If

	Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Sub DbQueryOk(byval strtxttab)

	With frm1
			
	   .vspdData.Col = C_WKCHK
	   lgintItemCnt = .vspddata.MaxRows

        lgIntFlgMode = parent.OPMD_UMODE	
        Call ggoOper.LockField(Document, "1")	
        Call SetToolbar("1100100000011111")			
        frm1.txtCommandMode.value = "UPDATE"
        Call SetSumItem()

        If strtxttab = "1" Then
			If .vspdData.MaxRows > 0 Then
				.vspdData.Row	 = 1
				.vspdData.Col	 = C_TEMPGLNO
				.htempglno.Value = .vspdData.Text
				Call DbQuery2(strtxttab,1,1)
            Else
				If .vspdData2.MaxRows > 0 Then
					.vspdData2.Row	 = 1
					.vspdData2.Col	 = C_TEMPGLNO
					.htempglno.Value = .vspdData2.Text
					Call DbQuery2(strtxttab,1,2)
				End If   
			End If
		Else
			If strtxttab = "2" Then
				If .vspdData6.MaxRows > 0 Then
				   .vspdData6.Row = 1
				   .vspdData6.Col = C_TEMPGLNO
				   .htempglno.Value = .vspdData6.Text
					Call DbQuery2(strtxttab,1,1)
				Else
					If .vspdData7.MaxRows > 0 Then
					    .vspdData7.Row = 1
				        .vspdData7.Col = C_TEMPGLNO
						.htempglno.Value = .vspdData7.Text
						Call DbQuery2("1",1,2)
					End If
				End If
		    End If
		End if

    End With

End Sub

'========================================================================================================
Sub DbQueryOk2(byval strtxttab)

	With frm1

        If strtxttab = "1" Then
			If .vspdData2.MaxRows > 0 Then
			   .vspdData2.Row	= 1
			   .vspdData2.Col	= C_TEMPGLNO
			   .htempglno.Value = .vspdData2.Text
			    Call DbQuery2(strtxttab,1,2)
			End If
        Else
           If strtxttab = "2" Then
        		If .vspdData7.MaxRows > 0 Then
				   .vspdData7.Row	= 1
				   .vspdData7.Col	= C_TEMPGLNO
				   .htempglno.Value = .vspdData7.Text
				    Call DbQuery2(strtxttab,1,2)
				End If
		   End If
        End if

    End With

End Sub

'========================================================================================================
Sub DbQueryOk3()

	If frm1.vspddata7.MaxRows > 0 Then
		lsClickRow7 = frm1.vspddata7.ActiveRow
	End If

End Sub

'========================================================================================================
Sub DbSaveOk()

	Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
	ggoSpread.Source = frm1.vspdData
	Call ggoSpread.ClearSpreadData()
	ggoSpread.Source = frm1.vspdData2
	Call ggoSpread.ClearSpreadData()
	ggoSpread.Source = frm1.vspdData3
	Call ggoSpread.ClearSpreadData()
	ggoSpread.Source = frm1.vspdData4
	Call ggoSpread.ClearSpreadData()
	ggoSpread.Source = frm1.vspdData6
	Call ggoSpread.ClearSpreadData()
	ggoSpread.Source = frm1.vspdData7
	Call ggoSpread.ClearSpreadData()
	ggoSpread.Source = frm1.vspdData8
	Call ggoSpread.ClearSpreadData()
	ggoSpread.Source = frm1.vspdData9
	Call ggoSpread.ClearSpreadData()

	If DbQuery = False Then
		Call RestoreToolBar()
		Exit Sub
    End If                                                                     '☜: Query db data

End Sub

'========================================================================================================
Sub DbDeleteOk()
	Call FncNew()
End Sub

'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================
' Function Name : OpenCtrlPB()
' Function Desc : PopUp 관리항목 
'========================================================================================================
Function OpenPopUp(Byval strCode, Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strBizAreaCd

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	frm1.txtOrgChangeId.value = parent.gChangeOrgId

	Select Case iWhere
		Case 0
			arrParam(0) = "사업장 팝업"				' 팝업 명칭 
			arrParam(1) = "B_BIZ_AREA"					' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""
			arrParam(5) = "사업장코드"				' 조건필드의 라벨 명칭 
			arrField(0) = "BIZ_AREA_CD"					' Field명(0)
			arrField(1) = "BIZ_AREA_NM"					' Field명(1)

			arrHeader(0) = "사업장코드"				' Header명(0)
			arrHeader(1) = "사업장명"				' Header명(1)

	End Select

   	If iWhere = 0 Then
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=400px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetPopUp(arrRet, iWhere)
	End If
	frm1.txtBizArea.focus
	
End Function

'========================================================================================================
Function SetPopUp(Byval arrRet, Byval iWhere)

	With frm1

		Select Case iWhere
			Case 0
				.txtBizArea.value	= arrRet(0)
				.txtBizAreanm.value = arrRet(1)
		End Select

	End With

End Function

'========================================================================================================
Function MaxSpreadVal(byval Row)

  Dim iRows
  Dim MaxValue
  Dim tmpVal

	MAxValue = 0

	With frm1

		For iRows = 1 to  .vspdData.MaxRows
			.vspddata.row = iRows
	        .vspddata.col = C_ITEMSEQ
			If .vspdData.Text = "" Then
			   tmpVal = 0
			Else
  			   tmpVal = UNICDbl(.vspdData.Text)
			End If

			If tmpval > MaxValue   Then
			   MaxValue = UNICDbl(tmpVal)
			End If
		Next

		MaxValue = MaxValue + 1

		.vspdData.Row	= Row
		.vspdData.Col	= C_ITEMSEQ
		.vspdData.Text	= MaxValue

	End With

End Function

'=======================================================================================================
Function SetSumItem()

    Dim DblTotDrAmt
    Dim DblTotCrAmt
    Dim lngRows

	SetSumItem = False
	ggoSpread.Source = frm1.vspdData

    With frm1.vspdData

		If .MaxRows > 0 Then
			For lngRows = 1 to .MaxRows
		        .Row = lngRows
		        .Col = C_WKCHK
		        If Trim(.Text) = "1"  Then
			       .Col = C_DRLOCAMT
				   If .Text = "" Then
				        DblTotDrAmt = UNICDbl(DblTotDrAmt) + 0
				   Else
				        DblTotDrAmt = UNICDbl(DblTotDrAmt) + UNICDbl(.Text)
				   End If
			    End If
			Next
		End If

    End With

    With frm1.vspdData2

		If .MaxRows > 0 Then
			For lngRows = 1 to .MaxRows
		        .Row = lngRows
		        .Col = C_WKCHK
		        If Trim(.Text) = "1"  Then
			       .Col = C_DRLOCAMT
				   If .Text = "" Then
				        DblTotCrAmt = UNICDbl(DblTotCrAmt) + 0
				   Else
				        DblTotCrAmt = UNICDbl(DblTotCrAmt) + UNICDbl(.Text)
				   End If
				End If
			Next
		 End If

    End With

    frm1.txtDrAmt.Text = UNIConvNumPCToCompanyByCurrency(DblTotDrAmt, parent.gCurrency, parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
    frm1.txtCrAmt.Text = UNIConvNumPCToCompanyByCurrency(DblTotCrAmt, parent.gCurrency, parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")

    If DblTotDrAmt <> DblTotCrAmt Then
       SetSumItem = true
    End If

End Function

'========================================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row)
	frm1.vspdData.Row = Row
   	frm1.vspdData.Col = Col
	Call CheckMinNumSpread(frm1.vspdData, Col, Row)
	ggoSpread.Source = frm1.vspdData
End Sub

'========================================================================================================
Sub vspdData2_Change(ByVal Col, ByVal Row)
	frm1.vspdData2.Row = Row
   	frm1.vspdData2.Col = Col
	Call CheckMinNumSpread(frm1.vspdData2, Col, Row)
	ggoSpread.Source = frm1.vspdData2
End Sub

'=======================================================================================================
Sub vspdData3_Change(ByVal Col, ByVal Row)
	frm1.vspdData3.Row = Row
   	frm1.vspdData3.Col = Col
	Call CheckMinNumSpread(frm1.vspdData3, Col, Row)
	ggoSpread.Source = frm1.vspdData3
End Sub

'=======================================================================================================
Sub vspdData4_Change(ByVal Col, ByVal Row)
	frm1.vspdData4.Row = Row
   	frm1.vspdData4.Col = Col
	Call CheckMinNumSpread(frm1.vspdData4, Col, Row)
	ggoSpread.Source = frm1.vspdData4
End Sub

'=======================================================================================================
Sub vspdData6_Change(ByVal Col, ByVal Row)
	frm1.vspdData6.Row = Row
   	frm1.vspdData6.Col = Col
	Call CheckMinNumSpread(frm1.vspdData6, Col, Row)
	ggoSpread.Source = frm1.vspdData6
End Sub

'=======================================================================================================
Sub vspdData7_Change(ByVal Col, ByVal Row)
	frm1.vspdData7.Row = Row
   	frm1.vspdData7.Col = Col
	Call CheckMinNumSpread(frm1.vspdData7, Col, Row)
	ggoSpread.Source = frm1.vspdData7
End Sub

'=======================================================================================================
Sub vspdData8_Change(ByVal Col, ByVal Row)
	frm1.vspdData8.Row = Row
   	frm1.vspdData8.Col = Col
	Call CheckMinNumSpread(frm1.vspdData8, Col, Row)
	ggoSpread.Source = frm1.vspdData8
End Sub

'=======================================================================================================
Sub vspdData9_Change(ByVal Col, ByVal Row)
	frm1.vspdData9.Row = Row
   	frm1.vspdData9.Col = Col
	Call CheckMinNumSpread(frm1.vspdData9, Col, Row)
	ggoSpread.Source = frm1.vspdData9
End Sub

'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SPC"				'Split 상태코드 
	Set gActiveSpdSheet = frm1.vspdData

	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
	
	ggoSpread.Source = frm1.vspdData

    If Row <= 0 Then
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey
			lgSortKey = 1
		End If
		Exit Sub 
    End If
	
    If Row > 0 And Row <> lsClickRow Then
		ggoSpread.Source  = frm1.vspdData
		frm1.vspddata.Row = frm1.vspddata.ActiveRow
		lsClickRow		  = Row
		
'		ggoSpread.Source  = frm1.vspdData2
'		Call ggoSpread.ClearSpreadData()
		ggoSpread.Source  = frm1.vspdData3
		Call ggoSpread.ClearSpreadData()
		ggoSpread.Source  = frm1.vspdData4
		Call ggoSpread.ClearSpreadData()

		Call DbQuery2("1",Row,1)
	End If
	
	If Col = C_WKCHK Then 
		frm1.vspddata.COL =  C_WKCHK
		if frm1.vspddata.value = "1" then		
			frm1.vspddata.value = "0"
		else
			frm1.vspddata.value = "1"
		end if				
		lgBlnFlgChgValue = True
	End If

End Sub

'========================================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SP1C"				'Split 상태코드 
	Set gActiveSpdSheet = frm1.vspdData2

	If frm1.vspdData2.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
	
	ggoSpread.Source = frm1.vspdData2

    If Row <= 0 Then
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey
			lgSortKey = 1
		End If
		Exit Sub 
    End If
    
   	If Row > 0 And Row <> lsClickRow2 Then
		gMouseClickStatus = "SP2C"	'Split 상태코드 
		ggoSpread.Source  = frm1.vspdData2
		frm1.vspddata.Row = frm1.vspddata2.ActiveRow
		lsClickRow2		  = Row
		
		ggoSpread.Source  = frm1.vspdData4
		Call ggoSpread.ClearSpreadData()
		
		Call DbQuery2("1",Row,2)
	End If

	If Col = C_WKCHK Then 
		frm1.vspdData2.COL =  C_WKCHK
		if frm1.vspdData2.value = "1" then		
			frm1.vspdData2.value = "0"
		else
			frm1.vspdData2.value = "1"
		end if				

		lgBlnFlgChgValue2 = True
	End If

End Sub

'========================================================================================================
Sub vspdData3_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SP2C"				'Split 상태코드 
	Set gActiveSpdSheet = frm1.vspdData3

	If frm1.vspdData3.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
	
	ggoSpread.Source = frm1.vspdData3

    If Row <= 0 Then
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey
			lgSortKey = 1
		End If
		Exit Sub 
    End If

End Sub

'========================================================================================================
Sub vspdData4_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SP3C"				'Split 상태코드 
	Set gActiveSpdSheet = frm1.vspdData4

	If frm1.vspdData4.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
	
	ggoSpread.Source = frm1.vspdData4

    If Row <= 0 Then
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey
			lgSortKey = 1
		End If
		Exit Sub 
    End If

End Sub

'========================================================================================================
Sub vspdData6_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SP4C"				'Split 상태코드 
	Set gActiveSpdSheet = frm1.vspdData6

	If frm1.vspdData6.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
	
	ggoSpread.Source = frm1.vspdData6

    If Row <= 0 Then
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey
			lgSortKey = 1
		End If
		Exit Sub 
    End If
    
    If Row > 0 And Row <> lsClickRow6 Then
		ggoSpread.Source  = frm1.vspdData6
		frm1.vspddata.Row = frm1.vspddata6.ActiveRow
		lsClickRow6		  = Row
		
		ggoSpread.Source  = frm1.vspdData7
		Call ggoSpread.ClearSpreadData()
		ggoSpread.Source  = frm1.vspdData8
		Call ggoSpread.ClearSpreadData()
		ggoSpread.Source  = frm1.vspdData9
		Call ggoSpread.ClearSpreadData()
		
		Call DbQuery2("2",Row,1)
	End If
	
	If Col = C_WKCHK Then 
		frm1.vspdData6.COL =  C_WKCHK
		if frm1.vspdData6.value = "1" then		
			frm1.vspdData6.value = "0"
		else
			frm1.vspdData6.value = "1"
		end if				

		lgBlnFlgChgValue2 = True
	End If

End Sub

'========================================================================================================
Sub vspdData7_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SP5C"				'Split 상태코드 
	Set gActiveSpdSheet = frm1.vspdData7

	If frm1.vspdData7.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
	
	ggoSpread.Source = frm1.vspdData7

    If Row <= 0 Then
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey
			lgSortKey = 1
		End If
		Exit Sub 
    End If

    If Row > 0 And Row <> lsClickRow7 Then
		ggoSpread.Source  = frm1.vspdData7
		frm1.vspddata.Row = frm1.vspddata7.ActiveRow
		lsClickRow7		  = Row
		
		ggoSpread.Source  = frm1.vspdData9
		Call ggoSpread.ClearSpreadData()

		Call DbQuery2("2",Row,2)
	End If

End Sub

'========================================================================================================
Sub vspdData8_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SP6C"				'Split 상태코드 
	Set gActiveSpdSheet = frm1.vspdData8

	If frm1.vspdData8.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
	
	ggoSpread.Source = frm1.vspdData8

    If Row <= 0 Then
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey
			lgSortKey = 1
		End If
		Exit Sub 
    End If

End Sub

'========================================================================================================
Sub vspdData9_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SP7C"				'Split 상태코드 
	Set gActiveSpdSheet = frm1.vspdData9

	If frm1.vspdData9.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
	
	ggoSpread.Source = frm1.vspdData9

    If Row <= 0 Then
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey
			lgSortKey = 1
		End If
		Exit Sub 
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
Sub vspdData3_ColWidthChange(ByVal pvCol1, ByVal pvCol2)

    ggoSpread.Source = frm1.vspdData3
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
Sub vspdData4_ColWidthChange(ByVal pvCol1, ByVal pvCol2)

    ggoSpread.Source = frm1.vspdData4
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
Sub vspdData6_ColWidthChange(ByVal pvCol1, ByVal pvCol2)

    ggoSpread.Source = frm1.vspdData6
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
Sub vspdData7_ColWidthChange(ByVal pvCol1, ByVal pvCol2)

    ggoSpread.Source = frm1.vspdData7
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
Sub vspdData8_ColWidthChange(ByVal pvCol1, ByVal pvCol2)

    ggoSpread.Source = frm1.vspdData8
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
Sub vspdData9_ColWidthChange(ByVal pvCol1, ByVal pvCol2)

    ggoSpread.Source = frm1.vspdData9
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
Sub vspdData_onfocus()
    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call SetToolbar("1100100000011111")                                     '버튼 툴바 제어 
    Else
        Call SetToolbar("1100100000011111")                                     '버튼 툴바 제어 
    End If
End Sub

'========================================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'========================================================================================================
Sub vspdData2_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP1C" Then
		gMouseClickStatus = "SP1CR"
	End If
End Sub

'========================================================================================================
Sub vspdData3_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub

'========================================================================================================
Sub vspdData4_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP3C" Then
		gMouseClickStatus = "SP3CR"
	End If
End Sub

'========================================================================================================
Sub vspdData6_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP4C" Then
		gMouseClickStatus = "SP4CR"
	End If
End Sub

'========================================================================================================
Sub vspdData7_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP5C" Then
		gMouseClickStatus = "SP5CR"
	End If
End Sub

'========================================================================================================
Sub vspdData8_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP6C" Then
		gMouseClickStatus = "SP6CR"
	End If
End Sub

'========================================================================================================
Sub vspdData9_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP7C" Then
		gMouseClickStatus = "SP7CR"
	End If
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
Sub vspdData3_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

    ggoSpread.Source = frm1.vspdData3
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("C")

End Sub

'========================================================================================================
Sub vspdData4_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

    ggoSpread.Source = frm1.vspdData4
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("D")

End Sub

'========================================================================================================
Sub vspdData6_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

    ggoSpread.Source = frm1.vspdData6
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("E")

End Sub

'========================================================================================================
Sub vspdData7_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

    ggoSpread.Source = frm1.vspdData7
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("F")

End Sub

'========================================================================================================
Sub vspdData8_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

    ggoSpread.Source = frm1.vspdData8
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("G")

End Sub

'========================================================================================================
Sub vspdData9_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

    ggoSpread.Source = frm1.vspdData9
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("H")

End Sub

'========================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

    With frm1.vspdData

		If Row >= NewRow Then
		    Exit Sub
		End If

    End With

    call SetSumItem()

End Sub

'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

End Sub

'========================================================================================================
Sub txtFromdt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery
End Sub

'========================================================================================================
Sub txtTodt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery
End Sub

'========================================================================================================
Sub txtfromamt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery
End Sub

'========================================================================================================
Sub txttoamt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery
End Sub

'========================================================================================================
Sub txtfromdt_DblClick(Button)
    If Button = 1 Then
        frm1.txtfromdt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtfromdt.focus
    End If
End Sub
'========================================================================================================
Sub txttodt_DblClick(Button)
    If Button = 1 Then
        frm1.txttodt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txttodt.focus
    End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2KCM.inc"  -->
</HEAD>

<BODY TABINDEX="-1" SCROLL=NO>
<FORM NAME=frm1 TARGET=MyBizASP METHOD=POST>
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
				     <TD WIDTH=10>&nbsp;</TD>
				 	<TD CLASS=CLSMTABP>
				 		<TABLE ID=MyTab CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab1()">
				 			<TR>
				 	           <TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif"><IMG SRC="../../../CShared/image/table/seltab_up_left.gif" WIDTH=9 HEIGHT=23></TD>
				 	           <TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN=CENTER CLASS=CLSMTABP><FONT COLOR=WHITE>미달거래연결</FONT></TD>
				 	           <TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN=RIGHT><IMG SRC="../../../CShared/image/table/seltab_up_right.gif" WIDTH=10 HEIGHT=23></TD>
				            </TR>
				        </TABLE>
				     </TD>
				     <TD CLASS=CLSMTABP>
				 		<TABLE ID=MyTab CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
				 			<TR>
				 			   <TD BACKGROUND="../../../CShared/image/table/tab_up_bg.gif"><IMG HEIGHT=23 SRC="../../../CShared/image/table/tab_up_left.gif" WIDTH=9></TD>
				 			   <TD BACKGROUND="../../../CShared/image/table/tab_up_bg.gif" ALIGN=CENTER CLASS=CLSMTABP><FONT COLOR=WHITE>미달거래연결취소</FONT></TD>
				 			   <TD BACKGROUND="../../../CShared/image/table/tab_up_bg.gif" ALIGN=RIGHT><IMG HEIGHT=23 SRC="../../../CShared/image/table/tab_up_right.gif" WIDTH=10></TD>
				 		    </TR>
				 		</TABLE>
				 	</TD>
				     <TD WIDTH=*>&nbsp;</TD>
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
							<TABLE CLASS=BasicTB CELLSPACING=0>
								<TR>
									<TD CLASS=TD5 NOWRAP>사업장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME=txtBizArea   ALT="사업장"   MAXLENGTH=10 SIZE=10 STYLE="TEXT-ALIGN: LEFT" tag ="11"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME=btnBizArea ALIGN=TOP TYPE=BUTTON ONCLICK="vbscript:Call OpenPopup(frm1.txtBizArea.Value, 0)">&nbsp;
														 <INPUT TYPE=TEXT NAME=txtBizAreaNm ALT="사업장명" MAXLENGTH=20 SIZE=20 STYLE="TEXT-ALIGN: LEFT" tag ="14X"></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>결의일</TD>
									<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a8102ma1_fpDateTime1_txtFromdt.js'></script>&nbsp;~&nbsp;
							                             <script language =javascript src='./js/a8102ma1_fpDateTime2_txtTodt.js'></script></TD>
							        <TD CLASS=TD5 NOWRAP>금액</TD>
									<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a8102ma1_fpDoubleSingle1_txtfromamt.js'></script>&nbsp;~&nbsp;
														 <script language =javascript src='./js/a8102ma1_fpDoubleSingle1_txttoamt.js'></script></TD>
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
					<!--첫번째 TAB  -->
					<DIV ID=TabDiv STYLE="FLOAT: left; HEIGHT: 100%; OVERFLOW: AUTO; WIDTH: 100%" SCROLL=NO>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
							  <TD COLSPAN=3>
							  <script language =javascript src='./js/a8102ma1_I609974461_vspdData.js'></script>
							  </TD>
							  <TD COLSPAN=3>
							  <script language =javascript src='./js/a8102ma1_I534250283_vspdData2.js'></script>
							  </TD>
							</TR>
				            <TR HEIGHT=20>
								<TD CLASS=TD5 NOWRAP>차대합계(거래)</TD>
								<TD ><script language =javascript src='./js/a8102ma1_I157405334_txtDrAmt.js'></script></TD>
								<TD ><script language =javascript src='./js/a8102ma1_I579530906_txtCrAmt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>차대합계(자국)</TD>
								<TD ><script language =javascript src='./js/a8102ma1_I310657852_txtDrLocAmt.js'></script></TD>
								<TD ><script language =javascript src='./js/a8102ma1_I768516027_txtCrLocAmt.js'></script></TD>
	                        </TR>
                            <TR HEIGHT="30%">
								<TD COLSPAN=3><script language =javascript src='./js/a8102ma1_I334219449_vspdData3.js'></script>
								</TD>
								<TD COLSPAN=3><script language =javascript src='./js/a8102ma1_I359672057_vspdData4.js'></script>
								</TD>
							</TR>
						</TABLE>
					</DIV>
					<!--두번째 TAB  -->
					<DIV ID=TabDiv STYLE="FLOAT: left; HEIGHT: 100%; OVERFLOW: AUTO; WIDTH: 100%" SCROLL=NO>
						<TABLE <%=LR_SPACE_TYPE_20%>>
						    <TR>
								<TD COLSPAN=3>
								<script language =javascript src='./js/a8102ma1_I375983911_vspdData6.js'></script>
								</TD>
								<TD COLSPAN=3>
								<script language =javascript src='./js/a8102ma1_I442789316_vspdData7.js'></script>
								</TD>
						    </TR>
						    <TR HEIGHT=20>
								<TD CLASS=TD5 NOWRAP>차대합계(거래)</TD>
								<TD ><script language =javascript src='./js/a8102ma1_fpDoubleSingle2_txtDrAmt2.js'></script></TD>
								<TD ><script language =javascript src='./js/a8102ma1_fpDoubleSingle3_txtCrAmt2.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>차대합계(자국)</TD>
								<TD ><script language =javascript src='./js/a8102ma1_fpDoubleSingle4_txtDrLocAmt2.js'></script></TD>
								<TD ><script language =javascript src='./js/a8102ma1_fpDoubleSingle5_txtCrLocAmt2.js'></script></TD>
						    </TR>
						    <TR HEIGHT="30%">
								<TD COLSPAN=3><script language =javascript src='./js/a8102ma1_I513900202_vspdData8.js'></script>
								</TD>
								<TD COLSPAN=3><script language =javascript src='./js/a8102ma1_I329736244_vspdData9.js'></script>
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
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO NORESIZE FRAMESPACING=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
	
</TABLE>
<TEXTAREA CLASS=HIDDEN NAME=txtSpread			tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA CLASS=HIDDEN NAME=txtSpread3			tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT    TYPE=HIDDEN  NAME=txtMode             tag="24" TABINDEX="-1">
<INPUT    TYPE=HIDDEN  NAME=txtCommandMode		tag="24" TABINDEX="-1">
<INPUT    TYPE=HIDDEN  NAME=txtOrgChangeId		tag="24" TABINDEX="-1">
<INPUT    TYPE=HIDDEN  NAME=txtUpdtUserId		tag="24" TABINDEX="-1">
<INPUT    TYPE=HIDDEN  NAME=txtstrtab			tag="24" TABINDEX="-1">
<INPUT    TYPE=HIDDEN  NAME=txtMaxRows			tag="24" TABINDEX="-1">
<INPUT    TYPE=HIDDEN  NAME=txtFlgMode			tag="24" TABINDEX="-1">
<INPUT    TYPE=HIDDEN  NAME=txtGlinputType		tag="24" TABINDEX="-1">
<INPUT    TYPE=HIDDEN  NAME=txtNextKey			tag="24" TABINDEX="-1">
<INPUT    TYPE=HIDDEN  NAME=htempglno			tag="24" TABINDEX="-1">
<INPUT    TYPE=HIDDEN  NAME=hAcctCd				tag="24" TABINDEX="-1">
<INPUT    TYPE=HIDDEN  NAME=txtMaxRows3			tag="24" TABINDEX="-1">


</FORM>
<DIV ID=MousePT NAME=MousePT>
<IFRAME NAME=MouseWindow FRAMEBORDER=0 SCROLLING=NO NORESIZE FRAMESPACING=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
