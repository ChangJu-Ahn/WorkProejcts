<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
'*  1. Module Name          : Human Resource
'*  2. Function Name        : 연말정산관리 
'*  3. Program ID           : h9111ma1.asp
'*  4. Program Name         : h9111ma1.asp
'*  5. Program Desc         : 연말정산급/상여내역조회 
'*  6. Modified date(First) : 2001/06/08
'*  7. Modified date(Last)  : 2003/06/13
'*  8. Modifier (First)     : Song Bong-kyu
'*  9. Modifier (Last)      : Lee SiNa
'* 10. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/IncHRQuery.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/IncCliRdsQuery.vbs">   </SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const CookieSplit = 1233
Const BIZ_PGM_ID = "h9111mb1.asp"                                      'Biz Logic ASP 
Const C_SHEETMAXROWS    = 21	                                      '한 화면에 보여지는 최대갯수*1.5%>

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd
Dim IsOpenPop          
Dim lsInternal_cd
Dim lgStrPrevKey1
Dim topleftOK

Dim C_PayDt
Dim C_PayCd
Dim C_PayTotAmt
Dim C_BonusTotAmt
Dim C_NotaxAmt
Dim C_TaxAmt
Dim C_IncomTaxAmt
Dim C_ResTaxAmt
Dim C_SaveFundAmt
															
Dim C_PayDt2  '      = 1
Dim C_PayText '    = 2
Dim C_IncomTaxAmt2'= 4															
Dim C_ResTaxAmt2  '= 5															

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgStrPrevKey1     = ""                                      '⊙: initializes Previous Key    
    lgCurrentSpd      = 1 
    lgSortKey         = 1                                       '⊙: initializes sort direction
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	Dim strYear,strMonth,strDay
	lgBlnFlgChgValue = False
	
   	frm1.txtYear.focus
    Call ExtractDateFrom("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gServerDateType,strYear,strMonth,strDay)	
    frm1.txtYear.Year	= strYear
End Sub	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H","NOCOOKIE","MA") %>
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)
   
    lgKeyStream       = frm1.txtYear.Year & parent.gColSep                                           'You Must append one character(parent.gColSep)
	lgKeyStream       = lgKeyStream & Frm1.txtEmp_no.Value & parent.gColSep
End Sub        

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	frm1.vspdData1S.Col = C_PayTotAmt
	frm1.vspdData1S.Row = 1
	frm1.vspdData1S.Text = FncSumSheet(frm1.vspdData1,C_PayTotAmt,1,frm1.vspdData1.MaxRows,False,7,9,"V")

	frm1.vspdData1S.Col = C_BonusTotAmt
	frm1.vspdData1S.Row = 1
	frm1.vspdData1S.Text = FncSumSheet(frm1.vspdData1,C_BonusTotAmt,1,frm1.vspdData1.MaxRows,False,7,9,"V")

	frm1.vspdData1S.Col = C_NotaxAmt
	frm1.vspdData1S.Row = 1
	frm1.vspdData1S.Text = FncSumSheet(frm1.vspdData1,C_NotaxAmt,1,frm1.vspdData1.MaxRows,False,7,9,"V")
	
	frm1.vspdData1S.Col = C_TaxAmt
	frm1.vspdData1S.Row = 1
	frm1.vspdData1S.Text = FncSumSheet(frm1.vspdData1,C_TaxAmt,1,frm1.vspdData1.MaxRows,False,7,9,"V")

	frm1.vspdData1S.Col = C_IncomTaxAmt
	frm1.vspdData1S.Row = 1
	frm1.vspdData1S.Text = FncSumSheet(frm1.vspdData1,C_IncomTaxAmt,1,frm1.vspdData1.MaxRows,False,7,9,"V")

	frm1.vspdData1S.Col = C_ResTaxAmt
	frm1.vspdData1S.Row = 1
	frm1.vspdData1S.Text = FncSumSheet(frm1.vspdData1,C_ResTaxAmt,1,frm1.vspdData1.MaxRows,False,7,9,"V")
	
	frm1.vspdData1S.Col = C_SaveFundAmt
	frm1.vspdData1S.Row = 1	
	frm1.vspdData1S.Text = FncSumSheet(frm1.vspdData1,C_SaveFundAmt,1,frm1.vspdData1.MaxRows,False,7,9,"V")

	frm1.vspdData2S.Col = C_PayTotAmt
	frm1.vspdData2S.Row = 1
	frm1.vspdData2S.Text = FncSumSheet(frm1.vspdData2,C_PayTotAmt,1,frm1.vspdData2.MaxRows,False,7,9,"V")

	frm1.vspdData2S.Col = C_IncomTaxAmt2
	frm1.vspdData2S.Row = 1
	frm1.vspdData2S.Text = FncSumSheet(frm1.vspdData2,C_IncomTaxAmt2,1,frm1.vspdData2.MaxRows,False,7,9,"V")

	frm1.vspdData2S.Col = C_ResTaxAmt2
	frm1.vspdData2S.Row = 1
	frm1.vspdData2S.Text = FncSumSheet(frm1.vspdData2,C_ResTaxAmt2,1,frm1.vspdData2.MaxRows,False,7,9,"V")

End Sub


sub InitSpreadPosVariables(spd)
	if (spd = "A" or spd="B")  then
		C_PayDt     = 1
		C_PayCd       = 2
		C_PayTotAmt   = 3															
		C_BonusTotAmt = 4															
		C_NotaxAmt    = 5															
		C_TaxAmt      = 6															
		C_IncomTaxAmt = 7															
		C_ResTaxAmt   = 8															
		C_SaveFundAmt = 9
	end if
	if (spd = "C" or spd="D")  then																
		C_PayDt2      = 1
		C_PayText     = 2
		C_IncomTaxAmt2= 4															
		C_ResTaxAmt2  = 5
	end if 
end sub


'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet(spd)
	call InitSpreadPosVariables(spd)
	if (spd = "A")  then
		With frm1.vspdData1
			ggoSpread.Source = Frm1.vspdData1
			ggoSpread.Spreadinit "V20021128",, parent.gAllowDragDropSpread
			.MaxCols = C_SaveFundAmt + 1										<%'☜: 최대 Columns의 항상 1개 증가시킴 %>

			.Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
			.ColHidden = True                                                            ' ☜:☜:

			.MaxRows = 0

			.ReDraw = false
				
			Call GetSpreadColumnPos(spd)

			ggoSpread.SSSetEdit  C_PayDt       , "지급일자", 10
			ggoSpread.SSSetEdit  C_PayCd       , "급여구분", 15
			ggoSpread.SSSetFloat C_PayTotAmt   , "급여총액", 15,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat C_BonusTotAmt , "상여총액", 15,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat C_NotaxAmt    , "비과세금액", 15,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat C_TaxAmt      , "과세금액", 15,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat C_IncomTaxAmt , "소득세", 15,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat C_ResTaxAmt   , "주민세", 15,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat C_SaveFundAmt , "재형기금", 15,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

			.ReDraw = true	
		End With
	End if
    if (spd = "B")  then
		With frm1.vspdData1S
			ggoSpread.Source = Frm1.vspdData1S
			ggoSpread.Spreadinit "V20021128",, parent.gAllowDragDropSpread
			.MaxCols = C_SaveFundAmt
		    
			.MaxRows = 1

			.ReDraw = false
			.DisplayColHeaders = False
				
			Call GetSpreadColumnPos(spd)
			.Col = C_PayCd 
			.Row = 1
			.Text = "합계"

			ggoSpread.SSSetEdit  C_PayDt       , "", 10
			ggoSpread.SSSetEdit  C_PayCd       , "", 15
			ggoSpread.SSSetFloat C_PayTotAmt   , "", 15,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat C_BonusTotAmt , "", 15,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat C_NotaxAmt    , "", 15,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat C_TaxAmt      , "", 15,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat C_IncomTaxAmt , "", 15,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat C_ResTaxAmt   , "", 15,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat C_SaveFundAmt , "", 15,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

			.ReDraw = true	
		End With
	end if
	
	if (spd = "C")  then
		With frm1.vspdData2
			ggoSpread.Source = Frm1.vspdData2
			ggoSpread.Spreadinit "V20021128",, parent.gAllowDragDropSpread
			.MaxCols = C_ResTaxAmt2 + 1										<%'☜: 최대 Columns의 항상 1개 증가시킴 %>

			.Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
			.ColHidden = True                                                            ' ☜:☜:

			.MaxRows = 0

			.ReDraw = false

			Call GetSpreadColumnPos(spd)

			ggoSpread.SSSetEdit  C_PayDt2      , "현물지급일", 15
			ggoSpread.SSSetEdit  C_PayText     , "급여구분", 30
			ggoSpread.SSSetFloat C_PayTotAmt   , "지급액", 20,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat C_IncomTaxAmt2, "소득세", 20,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat C_ResTaxAmt2  , "주민세", 20,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

			.ReDraw = true	
		End With
	end if
	
	if (spd = "D")  then    
		With frm1.vspdData2S
			ggoSpread.Source = Frm1.vspdData2S
			ggoSpread.Spreadinit "V20021128",, parent.gAllowDragDropSpread
			.MaxCols = C_ResTaxAmt2
		    
			.MaxRows = 1

			.ReDraw = false

			.DisplayColHeaders = False
			.OperationMode = 1

				
				Call GetSpreadColumnPos(spd)
			.Col = C_PayCd 
			.Row = 1
			.Text = "합계"

			ggoSpread.SSSetEdit  C_PayDt2      , "현물지급일", 15
			ggoSpread.SSSetEdit  C_PayText     , "급여구분", 30
			ggoSpread.SSSetFloat C_PayTotAmt   , "지급액", 20,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,,"10"
			ggoSpread.SSSetFloat C_IncomTaxAmt2, "소득세", 20,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,,"10"
			ggoSpread.SSSetFloat C_ResTaxAmt2  , "주민세", 20,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,,"10"

		.ReDraw = true	
		End With
	end if
	
    Call SetSpreadLock(spd)

End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock(spd)

	Select Case spd
		case "A"
			ggoSpread.Source = Frm1.vspdData1   
			ggoSpread.SpreadLockWithOddEvenRowColor()
      	case "B"
			ggoSpread.Source = Frm1.vspdData1S
			With frm1.vspdData1S
				.ReDraw = False
				ggoSpread.SpreadLock      -1,-1,-1
				ggoSpread.SSSetProtected  .MaxCols   , -1, -1
				.ReDraw = True
			End With
		case "C"
			ggoSpread.Source = Frm1.vspdData2
			ggoSpread.SpreadLockWithOddEvenRowColor()
		case "D"
			ggoSpread.Source = Frm1.vspdData2S
			With frm1.vspdData2S
				.ReDraw = False
				ggoSpread.SpreadLock      -1,-1,-1
				ggoSpread.SSSetProtected  .MaxCols   , -1, -1
				.ReDraw = True
			End With
	End Select
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData1.ReDraw = False
     ggoSpread.SSSetProtected    C_PayCd , pvStartRow, pvEndRow
     ggoSpread.SSSetProtected    C_ResTaxAmt  , pvStartRow, pvEndRow
    .vspdData1.ReDraw = True
    
    End With
End Sub

'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData1.MaxCols - 1
           Frm1.vspdData1.Col = iDx
           Frm1.vspdData1.Row = iRow
           If Frm1.vspdData1.ColHidden <> True And Frm1.vspdData1.BackColor <> parent.UC_PROTECTED Then
              Frm1.vspdData1.Col = iDx
              Frm1.vspdData1.Row = iRow
              Frm1.vspdData1.Action = 0 ' go to 
              Exit For
           End If          
       Next
    End If   
End Sub


Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_PayDt     = iCurColumnPos(1)
			C_PayCd       = iCurColumnPos(2)
			C_PayTotAmt   = iCurColumnPos(3)															
			C_BonusTotAmt = iCurColumnPos(4)															
			C_NotaxAmt    = iCurColumnPos(5)															
			C_TaxAmt      = iCurColumnPos(6)															
			C_IncomTaxAmt = iCurColumnPos(7)															
			C_ResTaxAmt   = iCurColumnPos(8)															
			C_SaveFundAmt = iCurColumnPos(9)
		Case "B"
            ggoSpread.Source = frm1.vspdData1s
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_PayDt     = iCurColumnPos(1)
			C_PayCd       = iCurColumnPos(2)
			C_PayTotAmt   = iCurColumnPos(3)															
			C_BonusTotAmt = iCurColumnPos(4)															
			C_NotaxAmt    = iCurColumnPos(5)															
			C_TaxAmt      = iCurColumnPos(6)															
			C_IncomTaxAmt = iCurColumnPos(7)															
			C_ResTaxAmt   = iCurColumnPos(8)															
			C_SaveFundAmt = iCurColumnPos(9)
		Case "C"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_PayDt2      = iCurColumnPos(1)
			C_PayText     = iCurColumnPos(2)
			C_IncomTaxAmt2= iCurColumnPos(4)														
			C_ResTaxAmt2  = iCurColumnPos(5)
		Case "D"
            ggoSpread.Source = frm1.vspdData2s
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_PayDt2      = iCurColumnPos(1)
			C_PayText     = iCurColumnPos(2)
			C_IncomTaxAmt2= iCurColumnPos(4)														
			C_ResTaxAmt2  = iCurColumnPos(5)
    End Select    
End Sub

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
	if	UCase(gActiveSpdSheet.id) = "VASPREAD1" then
	    ggoSpread.Source = frm1.vspdData1s 
	    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
	elseif	UCase(gActiveSpdSheet.id) = "VASPREAD2" then
	    ggoSpread.Source = frm1.vspdData2s 
	    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
	end if

End Sub


'========================================================================================
' Function Name : PopSaveSpreadColumnInf()
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()

    if isEmpty(TypeName(gActiveSpdSheet)) then
		exit sub
	elseif	UCase(gActiveSpdSheet.id) = "VASPREAD1" then
		ggoSpread.Source = frm1.vspdData1s 
		Call ggoSpread.SaveSpreadColumnInf()
	elseif	UCase(gActiveSpdSheet.id) = "VASPREAD2" then
		ggoSpread.Source = frm1.vspdData2s 
		Call ggoSpread.SaveSpreadColumnInf()
	end if
    
    
End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf()
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()

	dim sprd
	
	Select case UCase(gActiveSpdSheet.id)
		Case "VASPREAD1"
			sprd = "A"
		Case "VASPREAD1s"
			sprd = "B"
		Case "VASPREAD2"
			sprd = "C"
		Case "VASPREAD2s"
			sprd = "D"
	end select
	
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet(sprd)      
	Call ggoSpread.ReOrderingSpreadData()
	
	
	if sprd = "A" then
		ggoSpread.Source = frm1.vspdData1s
		Call ggoSpread.RestoreSpreadInf()
		Call InitSpreadSheet("B")      
		Call ggoSpread.ReOrderingSpreadData()
	elseif sprd = "C" then
		ggoSpread.Source = frm1.vspdData2s
		Call ggoSpread.RestoreSpreadInf()
		Call InitSpreadSheet("D")      

		Call ggoSpread.ReOrderingSpreadData()
	end if
	Call InitData()
	
End Sub

Sub vspdData1_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
    call vspdData1s_link_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 

Sub vspdData1s_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    cancel = true
End Sub 

Sub vspdData1s_link_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData1s
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub 

Sub vspdData2_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
    call vspdData2s_link_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 

Sub vspdData2s_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    cancel = true
End Sub 

Sub vspdData2s_link_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData2s
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub 

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
    ggoSpread.Source = Frm1.vspdData1

	Call AppendNumberPlace("6", "18", "2")

    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtYear, parent.gDateFormat, 3)

	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
           
    Call InitSpreadSheet("A")                                                           'Setup the Spread sheet
    Call InitSpreadSheet("B")                                                           'Setup the Spread sheet
    Call InitSpreadSheet("C")                                                           'Setup the Spread sheet
    Call InitSpreadSheet("D")                                                           'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    Call SetDefaultVal
    
    Call FuncGetAuth(gStrRequestMenuID, parent.gUsrID, lgUsrIntCd)                                ' 자료권한:lgUsrIntCd ("%", "1%")
    
	Call SetToolbar("1100000000001111")												'⊙: Set ToolBar

	Call CookiePage (0)                                                             '☜: Check Cookie
End Sub

'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 

    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData1
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgbox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    If txtEmp_no_Onchange() Then         'ENTER KEY 로 조회시 사원과 사번을 CHECK 한다 
        Exit Function
    End if

    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("X")
	
	topleftOK = false
    If DbQuery = False Then  
		Exit Function
	End If

    FncQuery = True                                                              '☜: Processing is OK

End Function

'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncNew = True																 '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncDelete = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
       
    FncSave = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()

End Function

'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
    ggoSpread.Source = Frm1.vspdData1	
    ggoSpread.EditUndo  
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow() 
End Function

'========================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow() 
End Function

'========================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function FncPrint()
    Call parent.FncPrint()
End Function
'========================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)                                         '☜: 화면 유형 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                    '☜:화면 유형, Tab 유무 
End Function

'========================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()
    Dim IntRetCD
	
    FncExit = True
End Function

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery() 
    Dim strVal
    Err.Clear                                                                        '☜: Clear err status

    DbQuery = False                                                                  '☜: Processing is NG
    
    Call LayerShowHide(1)

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                         '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
    strVal = strVal     & "&lgCurrentSpd="       & lgCurrentSpd                   '☜: Next key tag
        
	if lgCurrentSpd = "1" then
		strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
	else
		strVal = strVal     & "&lgStrPrevKey1=" & lgStrPrevKey1             '☜: Next key tag
	end if	

    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData1.MaxRows          '☜: Max fetched data
    
    Call RunMyBizASP(MyBizASP, strVal)                                               '☜:  Run biz logic

    DbQuery = True                                                                   '☜: Processing is NG

End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
	
    DbSave = True                                                           
    
End Function

'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
    Dim IntRetCd
    
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
	
    lgIntFlgMode = parent.OPMD_UMODE    
    Call ggoOper.LockField(Document, "Q")										'⊙: Lock field
    Call InitData()

	Call SetToolbar("1100000000001111")											'⊙: Set ToolBar
	
	Call DbQueryTotal                                                           ' 총액(single)에 뿌려주는 Sub
	frm1.vspdData1.focus	
End Function

'========================================================================================================
' Function Name : DbQueryNo
' Function Desc : Called by MB Area when query operation is not successful
'========================================================================================================
Function DbQueryNo()
    Call InitData()
	
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    
    Call InitVariables															'⊙: Initializes local global variables
	call DBQuery()
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()

End Function

'======================================================================================================
'	Name : DbQueryTotal()
'	Description : 총액 필드 계산 
'=======================================================================================================
Sub DbQueryTotal()

    Dim iTaxAmt, iIncomTaxAmt, iResTaxAmt
    
	With frm1.vspdData1S
        ggoSpread.Source = Frm1.vspdData1S

        .Col = C_PayTotAmt 
        .Row = 1
        Frm1.txtPayTotAmt.text = .text
        

        .Col = C_BonusTotAmt 
        .Row = 1
        Frm1.txtBonusTotAmt.text = .text

        .Col = C_SaveFundAmt 
        .Row = 1
        Frm1.txtSaveFundAmt.text = .text

        .Col = C_TaxAmt 
        .Row = 1
		iTaxAmt = UNICDbl(.text)
		
        .Col = C_IncomTaxAmt 
        .Row = 1
		iIncomTaxAmt = UNICDbl(.text)
		
        .Col = C_ResTaxAmt 
        .Row = 1
		iResTaxAmt = UNICDbl(.text)

	End With
		
	With frm1.vspdData2S
        ggoSpread.Source = Frm1.vspdData2S
        
        .Col = C_PayTotAmt 
        .Row = 1
        Frm1.txtTaxAmt.text = UNIFormatNumber(UNICDbl(.text) + Cdbl(iTaxAmt), ggAmtOfMoney.DecPoint, -2,  0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.Rndunit)

        .Col = C_IncomTaxAmt2 
        .Row = 1
        Frm1.txtIncomeTaxAmt.text = UNIFormatNumber(UNICDbl(.text) + Cdbl(iIncomTaxAmt), ggAmtOfMoney.DecPoint, -2,  0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.Rndunit)

        .Col = C_ResTaxAmt2 
        .Row = 1
        Frm1.txtResTaxAmt.text = UNIFormatNumber(UNICDbl(.text) + Cdbl(iResTaxAmt), ggAmtOfMoney.DecPoint, -2,  0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.Rndunit)

        .Col = C_PayTotAmt 
        .Row = 1
        Frm1.txtEtcTotAmt.text = .text

	End With

End Sub

'========================================================================================================
' Name : OpenEmpName()
' Desc : developer describe this line 
'========================================================================================================

Function OpenEmpName(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

    If  iWhere = 0 Then
	    arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""'frm1.txtName.value			' Name Cindition
	    arrParam(2) = lgUsrIntCd                    ' 자료권한 Condition  
    Else 'spread
        frm1.vspdData.Col = C_EMP_NO
		arrParam(0) = frm1.vspdData.Text			' Code Condition
        frm1.vspdData.Col = C_NAME
	    arrParam(1) = ""'frm1.vspdData.Text			' Name Cindition
	    arrParam(2) = lgUsrIntCd                    ' 자료권한 Condition  
	End If

	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtEmp_no.focus
		Exit Function
	Else
		Call SubSetCondEmp(arrRet, iWhere)
	End If	
			
End Function

'======================================================================================================
'	Name : SetCondArea()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetCondEmp(Byval arrRet, Byval iWhere)
	With frm1
		If iWhere = 0 Then 'TextBox(Condition)
			.txtEmp_no.value = arrRet(0)
			.txtName.value = arrRet(1)
			.txtEmp_no.focus
		end if
	End With
End Sub
'========================================================================================================
'   Event Name : txtEmp_no_Onchange             
'   Event Desc :
'========================================================================================================
Function txtEmp_no_Onchange()
    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd
    Dim strVal

    If frm1.txtEmp_no.value = "" Then
		frm1.txtName.value = ""

	    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
	    
        Call initData()		
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                              strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    if  IntRetCd < 0 then
	        if  IntRetCd = -1 then
    			Call DisplayMsgbox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call DisplayMsgbox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            end if
			frm1.txtName.value = ""
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true

		    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
		    Call initData()		
        Else
            frm1.txtName.value = strName
        End if 
    End if  
    
End Function 


'=======================================================================================================
'   Event Name : txtYear_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtYear_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtYear.Action = 7
        frm1.txtYear.focus
    End If
End Sub

'==========================================================================================
'   Event Name : txtpay_yymm_dt_KeyDown()
'   Event Desc : 조회조건부의 txtpay_yymm_dt_KeyDown시 EnterKey일 경우는 Query
'==========================================================================================
Sub txtYear_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call FncQuery()
End Sub
'========================================================================================================
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
        frm1.vspdData1S.LeftCol=NewLeft   	
		Exit Sub
	End If
	
	If frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
	
		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			Call DisableToolBar(Parent.TBC_QUERY)
			topleftOK = true			
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If  
End Sub
'========================================================================================================
'   Event Name : vspdData1S_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1S_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        frm1.vspdData1.LeftCol=NewLeft
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        frm1.vspdData2S.LeftCol=NewLeft
         exit sub
    End If
	If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then
    	If lgStrPrevKey1 <> "" Then                         
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
    	End If
    End if
End Sub
'========================================================================================================
'   Event Name : vspdData1S_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData2S_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        frm1.vspdData2.LeftCol=NewLeft
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData1_ColWidthChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1 , ByVal pvCol2 )
    ggoSpread.Source = frm1.vspdData1
    call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
    frm1.vspdData1S.ColWidth(pvCol1) = frm1.vspdData1.ColWidth(pvCol1)
    ggoSpread.Source = frm1.vspdData1s
    call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData1S_ColWidthChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1S_ColWidthChange(ByVal pvCol1 , ByVal pvCol2 )
    ggoSpread.Source = frm1.vspdData1s
    call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
    frm1.vspdData1.ColWidth(pvCol1) = frm1.vspdData1S.ColWidth(pvCol1)
    ggoSpread.Source = frm1.vspdData1
    call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData1_ColWidthChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1 , ByVal pvCol2 )
    ggoSpread.Source = frm1.vspdData2
    call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
    frm1.vspdData2S.ColWidth(pvCol1) = frm1.vspdData2.ColWidth(pvCol1)
    ggoSpread.Source = frm1.vspdData2s
    call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData1S_ColWidthChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData2S_ColWidthChange(ByVal pvCol1 , ByVal pvCol2 )
    ggoSpread.Source = frm1.vspdData2s
    call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
    frm1.vspdData2.ColWidth(pvCol1) = frm1.vspdData2S.ColWidth(pvCol1)
    ggoSpread.Source = frm1.vspdData2
    call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)
 	Call SetPopupMenuItemInf("0000101111")
    gMouseClickStatus = "SPC" 
    Set gActiveSpdSheet = frm1.vspdData1

	if frm1.vspddata1.MaxRows <= 0 then
		exit sub
	end if
	lgCurrentSpd = "1"	
	if Row <=0 then
		ggoSpread.Source = frm1.vspdData1
		if lgSortkey = 1 then
			ggoSpread.SSSort Col
			lgSortKey = 2
		else
			ggoSpread.SSSort Col, lgSortkey
			lgSortKey = 1
		end if
		Exit sub
	end if
	frm1.vspdData1.Row = Row

End Sub



Sub vspdData2_Click(ByVal Col, ByVal Row)
 	Call SetPopupMenuItemInf("0000101111")
    gMouseClickStatus = "SPC2" 
    Set gActiveSpdSheet = frm1.vspdData2

	if frm1.vspddata2.MaxRows <= 0 then
		exit sub
	end if
	lgCurrentSpd = "2"
	if Row <=0 then
		ggoSpread.Source = frm1.vspdData2
		if lgSortkey = 1 then
			ggoSpread.SSSort Col
			lgSortKey = 2
		else
			ggoSpread.SSSort Col, lgSortkey
			lgSortKey = 1
		end if
		Exit sub
	end if
	frm1.vspdData2.Row = Row

End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================
Sub vspdData1_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
     End If
End Sub  

Sub vspdData1s_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC1" Then
       gMouseClickStatus = "SP1CR"
     End If
End Sub  

Sub vspdData2_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC2" Then
       gMouseClickStatus = "SP2CR"
     End If
End Sub  


Sub vspdData2s_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC3" Then
       gMouseClickStatus = "SP3CR"
     End If
End Sub  

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
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
								<td background="../../../Cshared/Image/table/seltab_up_bg.gif"><img src="../../../Cshared/Image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../Cshared/Image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>연말정산급/상여내역조회</font></td>
								<td background="../../../Cshared/Image/table/seltab_up_bg.gif" align="right"><img src="../../../Cshared/Image/table/seltab_up_right.gif" width="10" height="23"></td>
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
									<TD CLASS=TD5 NOWRAP>정산년도</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/h9111ma1_fpDateTime1_txtYear.js'></script>
									</TD>	
									<TD CLASS=TD5 NOWRAP>사번</TD>
			     					<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no" ALT="사번" TYPE="Text" SiZE=15 MAXLENGTH=13 tag="12XXXU"><IMG SRC="../../../Cshared/Image/btnPopup.gif" NAME="btnEmpNo" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenEmpName('0')">
									                     <INPUT NAME="txtName" MAXLENGTH="30" SIZE="20"  ALT ="성명" tag="14XXXU"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				
				<TR HEIGHT=120>
					<TD HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_40%>>
						<TR>
							<TD WIDTH="100%" HEIGHT=56%>
								<script language =javascript src='./js/h9111ma1_vaSpread1_vspdData1.js'></script>
							</TD>
						</TR>
						<TR HEIGHT=38>
							<TD WIDTH="100%" >
								<script language =javascript src='./js/h9111ma1_vaSpread1s_vspdData1S.js'></script>
							</TD>
						</TR>
						<TR>
							<TD WIDTH="100%" HEIGHT=18%>
								<script language =javascript src='./js/h9111ma1_vaSpread2_vspdData2.js'></script>
							</TD>
						</TR>
						<TR HEIGHT=38>
							<TD WIDTH="100%" >
								<script language =javascript src='./js/h9111ma1_vaSpread2s_vspdData2S.js'></script>
							</TD>
						</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_50%>>
						<TR>  
							<TD CLASS=TD5>급여총액</TD>
							<TD CLASS=TD6><script language =javascript src='./js/h9111ma1_fpDoubleSingle2_txtPayTotAmt.js'></script></TD>
							<TD CLASS=TD5>소득세총액</TD>
							<TD CLASS=TD6><script language =javascript src='./js/h9111ma1_fpDoubleSingle2_txtIncomeTaxAmt.js'></script></TD>
						</TR>
						<TR>  
							<TD CLASS=TD5>상여총액</TD>
							<TD CLASS=TD6><script language =javascript src='./js/h9111ma1_fpDoubleSingle2_txtBonusTotAmt.js'></script></TD>
							<TD CLASS=TD5>주민세총액</TD>
							<TD CLASS=TD6><script language =javascript src='./js/h9111ma1_fpDoubleSingle2_txtResTaxAmt.js'></script></TD>
						</TR>
						<TR>  
							<TD CLASS=TD5>기타소득총액</TD>
							<TD CLASS=TD6><script language =javascript src='./js/h9111ma1_fpDoubleSingle2_txtEtcTotAmt.js'></script></TD>
							<TD CLASS=TD5>재형기금총액</TD>
							<TD CLASS=TD6><script language =javascript src='./js/h9111ma1_fpDoubleSingle2_txtSaveFundAmt.js'></script></TD>
						</TR>
						<TR>  
							<TD CLASS=TD5>과세총액</TD>
							<TD CLASS=TD6><script language =javascript src='./js/h9111ma1_fpDoubleSingle2_txtTaxAmt.js'></script></TD>
							<TD CLASS=TD5></TD>
							<TD CLASS=TD6></TD>
						</TR>
						</TABLE>	
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>  
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="h9111mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
<!--		<TD HEIGHT=80><IFRAME NAME="MyBizASP" SRC="h9111mb1.asp" WIDTH=100% HEIGHT=100% FRAMEBORDER=1 SCROLLING=YES noresize framespacing=0></IFRAME> -->
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="lgCurrentSpd"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

