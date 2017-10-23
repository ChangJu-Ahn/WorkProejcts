<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
'*  1. Module Name          : Human Resource
'*  2. Function Name        : 인사급여 
'*  3. Program ID           : h6025ma1
'*  4. Program Name         : h6025ma1.asp
'*  5. Program Desc         : 급여구분별지급현황조회 
'*  6. Modified date(First) : 2003/06/25
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Lee Si Na
'*  9. Modifier (Last)      : 
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
Const BIZ_PGM_ID = "h6025mb1.asp"                                      'Biz Logic ASP 
Const CookieSplit = 1233
Const C_SHEETMAXROWS    = 21	                                      '☜: Visble row
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

Dim C_DEPT_CD
Dim C_ANN_SAL_NUM
Dim C_ANN_SALARY
Dim C_MON_PAY_NUM
Dim C_MON_PAY
Dim C_DLY_WAGES_NUM
Dim C_DLY_WAGES
Dim C_HOUR_WAGES_NUM
Dim C_HOUR_WAGES
Dim C_PAY_TOTAL
Dim THIS_MON_PERSON
Dim THIS_MON_PAY
Dim THIS_MON_RETIRE 
Dim THIS_MON_RETIRE_PAY

Dim C_DEPT_CD1
Dim C_ANN_SAL_NUM1
Dim C_ANN_SALARY1
Dim C_MON_PAY_NUM1
Dim C_MON_PAY1
Dim C_DLY_WAGES_NUM1
Dim C_DLY_WAGES1
Dim C_HOUR_WAGES_NUM1
Dim C_HOUR_WAGES1
Dim C_PAY_TOTAL1
Dim THIS_MON_PERSON1
Dim THIS_MON_PAY1
Dim THIS_MON_RETIRE1
Dim THIS_MON_RETIRE_PAY1

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize the position
'========================================================================================================

sub InitSpreadPosVariables()
	C_DEPT_CD			=1
	C_ANN_SAL_NUM		=2
	C_ANN_SALARY		=3
	C_MON_PAY_NUM		=4
	C_MON_PAY			=5
	C_DLY_WAGES_NUM		=6
	C_DLY_WAGES			=7	
	C_HOUR_WAGES_NUM	=8
	C_HOUR_WAGES		=9
	C_PAY_TOTAL			=10
	THIS_MON_PERSON		=11
	THIS_MON_PAY		=12
	THIS_MON_RETIRE		=13
	THIS_MON_RETIRE_PAY	=14

	C_DEPT_CD1			=1
	C_ANN_SAL_NUM1		=2
	C_ANN_SALARY1		=3
	C_MON_PAY_NUM1		=4
	C_MON_PAY1			=5
	C_DLY_WAGES_NUM1	=6
	C_DLY_WAGES1		=7	
	C_HOUR_WAGES_NUM1	=8
	C_HOUR_WAGES1		=9
	C_PAY_TOTAL1		=10
	THIS_MON_PERSON1	=11
	THIS_MON_PAY1		=12
	THIS_MON_RETIRE1	=13
	THIS_MON_RETIRE_PAY1=14

end sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
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
Sub MakeKeyStream(pOpt)
    Dim strPayYYYYMM
    Dim strPayYYYY
    DIm strPayMM
    Dim StrDt
    StrDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtpay_yymm_dt.Year, Right("0" & frm1.txtpay_yymm_dt.month , 2), "01")
    strPayYYYY = Frm1.txtpay_yymm_dt.Year
    strPayMM  = Frm1.txtpay_yymm_dt.Month
    
    If len(strPayMM) = 1 Then
		strPayMM = "0" & strPayMM
	End if
    strPayYYYYMM = strPayYYYY & strPayMM
    
    lgKeyStream  = strPayYYYYMM & Parent.gColSep
    lgKeyStream  = lgKeyStream & Trim(Frm1.txtemp_no.value) & Parent.gColSep
    lgKeyStream  = lgKeyStream & Trim(Frm1.txtPay_grd1.value) & Parent.gColSep
    lgKeyStream  = lgKeyStream & Trim(Frm1.txtprov_cd.value) & Parent.gColSep    
    lgKeyStream  = lgKeyStream & Trim(Frm1.txtfr_internal_cd.value) & Parent.gColSep
    lgKeyStream  = lgKeyStream & Trim(Frm1.txtto_internal_cd.value) & Parent.gColSep
    lgKeyStream  = lgKeyStream & lgUsrIntcd & Parent.gColSep
    lgKeyStream  = lgKeyStream & StrDt & Parent.gColSep    
End Sub        
'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
    Dim dblSum

	With frm1.vspdData1
        ggoSpread.Source = frm1.vspdData2
		ggoSpread.UpdateRow 1
        frm1.vspdData2.Col = 0
        
        frm1.vspdData2.Text = "합계"
        frm1.vspdData2.Col = C_ANN_SAL_NUM1
        frm1.vspdData2.text = FncSumSheet(frm1.vspddata1,C_ANN_SAL_NUM, 1, .MaxRows, FALSE ,-1, -1, "V")
        frm1.vspdData2.Col = C_ANN_SALARY1
        frm1.vspdData2.text = FncSumSheet(frm1.vspddata1,C_ANN_SALARY, 1, .MaxRows,FALSE , -1, -1, "V")
        frm1.vspdData2.Col = C_MON_PAY_NUM1
        frm1.vspdData2.text = FncSumSheet(frm1.vspddata1,C_MON_PAY_NUM, 1, .MaxRows,FALSE , -1, -1, "V")
        frm1.vspdData2.Col = C_MON_PAY1
        frm1.vspdData2.text = FncSumSheet(frm1.vspddata1,C_MON_PAY, 1, .MaxRows,FALSE , -1, -1, "V")
        frm1.vspdData2.Col = C_DLY_WAGES_NUM1
        frm1.vspdData2.text = FncSumSheet(frm1.vspddata1,C_DLY_WAGES_NUM, 1, .MaxRows,FALSE , -1, -1, "V")
        frm1.vspdData2.Col = C_DLY_WAGES1
        frm1.vspdData2.text = FncSumSheet(frm1.vspddata1,C_DLY_WAGES, 1, .MaxRows,FALSE , -1, -1, "V")
        frm1.vspdData2.Col = C_HOUR_WAGES_NUM1
        frm1.vspdData2.text = FncSumSheet(frm1.vspddata1,C_HOUR_WAGES_NUM, 1, .MaxRows,FALSE , -1, -1, "V")
        frm1.vspdData2.Col = C_HOUR_WAGES1
        frm1.vspdData2.text = FncSumSheet(frm1.vspddata1,C_HOUR_WAGES, 1, .MaxRows,FALSE , -1, -1, "V")
        frm1.vspdData2.Col = C_PAY_TOTAL1
        frm1.vspdData2.text = FncSumSheet(frm1.vspddata1,C_PAY_TOTAL, 1, .MaxRows,FALSE , -1, -1, "V")
		frm1.vspdData2.Col = THIS_MON_PERSON1
		frm1.vspdData2.text = FncSumSheet(frm1.vspddata1,THIS_MON_PERSON, 1, .MaxRows,FALSE , -1, -1, "V")
		frm1.vspdData2.Col = THIS_MON_PAY1
		frm1.vspdData2.text = FncSumSheet(frm1.vspddata1,THIS_MON_PAY, 1, .MaxRows,FALSE , -1, -1, "V")
		frm1.vspdData2.Col = THIS_MON_RETIRE1
		frm1.vspdData2.text = FncSumSheet(frm1.vspddata1,THIS_MON_RETIRE, 1, .MaxRows,FALSE , -1, -1, "V")
		frm1.vspdData2.Col = THIS_MON_RETIRE_PAY1
        frm1.vspdData2.text = FncSumSheet(frm1.vspddata1,THIS_MON_RETIRE_PAY, 1, .MaxRows,FALSE , -1, -1, "V")
        
		call SetSpreadLock2

    End With
    
End Sub

Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay
	
	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	frm1.txtpay_yymm_dt.focus()	
	frm1.txtpay_yymm_dt.Year = strYear 		 '년월일 default value setting
	frm1.txtpay_yymm_dt.Month = strMonth 
End Sub
	
'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)
	call InitSpreadPosVariables()

    If pvSpdNo = "" OR pvSpdNo = "A" Then

		With frm1.vspdData1
			ggoSpread.Source = Frm1.vspdData1
			ggoSpread.Spreadinit "V20021128",, parent.gAllowDragDropSpread
			
			.ReDraw = false
			.MaxCols = THIS_MON_RETIRE_PAY + 1										<%'☜: 최대 Columns의 항상 1개 증가시킴 %>

			.Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
			.ColHidden = True                                                            ' ☜:☜:

			.MaxRows = 0
				
			Call GetSpreadColumnPos("A")
		    Call AppendNumberPlace("6","5","0")

			ggoSpread.SSSetEdit  C_DEPT_CD      , "부서명", 15
		    ggoSpread.SSSetFloat C_ANN_SAL_NUM	, "연봉제인원"	, 9,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec			
			ggoSpread.SSSetFloat C_ANN_SALARY   , "연봉제임금"	, 14,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		    ggoSpread.SSSetFloat C_MON_PAY_NUM	, "월급직인원"	, 9,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec			
			ggoSpread.SSSetFloat C_MON_PAY		, "월급직임금"	, 14,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		    ggoSpread.SSSetFloat C_DLY_WAGES_NUM, "일급직인원"	, 9,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec			
			ggoSpread.SSSetFloat C_DLY_WAGES	, "일급직임금"	, 14,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		    ggoSpread.SSSetFloat C_HOUR_WAGES_NUM	, "시급직인원"	, 9,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec			
			ggoSpread.SSSetFloat C_HOUR_WAGES		, "시급직임금"	, 14,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat C_PAY_TOTAL		, "임금합계"	, 15,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		    ggoSpread.SSSetFloat THIS_MON_PERSON	, "당월입사인원"	, 10,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec			
			ggoSpread.SSSetFloat THIS_MON_PAY		, "당월입사임금"	, 14,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		    ggoSpread.SSSetFloat THIS_MON_RETIRE	, "당월퇴사인원"	, 10,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec			
			ggoSpread.SSSetFloat THIS_MON_RETIRE_PAY	, "당월퇴사임금"	, 14,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

						
			.ReDraw = true	
		End With
    End If

    If pvSpdNo = "" OR pvSpdNo = "B" Then
	    With frm1.vspdData2

            ggoSpread.Source = frm1.vspdData2
            ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread

	       .ReDraw = false
           .MaxCols   = THIS_MON_RETIRE_PAY1 + 1                                                      ' ☜:☜: Add 1 to Maxcols
	       .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
           .ColHidden = True                                                            ' ☜:☜:
 
           .MaxRows = 0
            ggoSpread.ClearSpreadData

           .DisplayColHeaders = False

            Call GetSpreadColumnPos("B") 

		    Call AppendNumberPlace("6","5","0")
            
			ggoSpread.SSSetEdit  C_DEPT_CD1      , "부서명", 15
		    ggoSpread.SSSetFloat C_ANN_SAL_NUM1	, "연봉제인원"	, 9,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec			
			ggoSpread.SSSetFloat C_ANN_SALARY1   , "연봉제임금"	, 14,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		    ggoSpread.SSSetFloat C_MON_PAY_NUM1	, "월급직인원"	, 9,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec			
			ggoSpread.SSSetFloat C_MON_PAY1		, "월급직임금"	, 14,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		    ggoSpread.SSSetFloat C_DLY_WAGES_NUM1, "일급직인원"	, 9,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec			
			ggoSpread.SSSetFloat C_DLY_WAGES1	, "일급직임금"	, 14,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		    ggoSpread.SSSetFloat C_HOUR_WAGES_NUM1	, "시급직인원"	, 9,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec			
			ggoSpread.SSSetFloat C_HOUR_WAGES1		, "시급직임금"	, 14,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat C_PAY_TOTAL1		, "임금합계"	, 15,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		    ggoSpread.SSSetFloat THIS_MON_PERSON1	, "당월입사인원"	, 10,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec			
			ggoSpread.SSSetFloat THIS_MON_PAY1		, "당월입사임금"	, 14,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		    ggoSpread.SSSetFloat THIS_MON_RETIRE1	, "당월퇴사인원"	, 10,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec			
			ggoSpread.SSSetFloat THIS_MON_RETIRE_PAY1	, "당월퇴사임금"	, 14,"2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
                     
	       .ReDraw = true
    
        End With
    End If

    Call SetSpreadLock()

End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData1
	ggoSpread.SpreadLockWithOddEvenRowColor()
	call SetSpreadLock2
    
End Sub
Sub SetSpreadLock2()
	ggoSpread.Source = frm1.vspdData2 
    With frm1
		.vspdData2.ReDraw = False
		ggoSpread.SpreadLock    C_DEPT_CD1, -1, C_DEPT_CD1, -1
		ggoSpread.SpreadLock    C_ANN_SAL_NUM1, -1, C_ANN_SAL_NUM1, -1 
		ggoSpread.SpreadLock    C_ANN_SALARY1, -1, C_ANN_SALARY1, -1 
		ggoSpread.SpreadLock    C_MON_PAY_NUM1, -1, C_MON_PAY_NUM1, -1 
		ggoSpread.SpreadLock    C_MON_PAY1, -1, C_MON_PAY1, -1 
		ggoSpread.SpreadLock    C_DLY_WAGES_NUM1, -1, C_DLY_WAGES_NUM1, -1 
		ggoSpread.SpreadLock    C_DLY_WAGES1, -1, C_DLY_WAGES1, -1 
		ggoSpread.SpreadLock    C_HOUR_WAGES_NUM1, -1, C_HOUR_WAGES_NUM1, -1 
		ggoSpread.SpreadLock    C_HOUR_WAGES1, -1, C_HOUR_WAGES1, -1 
		ggoSpread.SpreadLock    C_PAY_TOTAL1, -1, C_PAY_TOTAL1, -1 
		ggoSpread.SpreadLock    THIS_MON_PERSON1, -1, THIS_MON_PERSON1, -1 
		ggoSpread.SpreadLock    THIS_MON_PAY1, -1, THIS_MON_PAY1, -1 
		ggoSpread.SpreadLock    THIS_MON_RETIRE1, -1, THIS_MON_RETIRE1, -1 
		ggoSpread.SpreadLock    THIS_MON_RETIRE_PAY1, -1, THIS_MON_RETIRE_PAY1, -1 
		ggoSpread.SSSetProtected   .vspdData2.MaxCols   , -1, -1
		.vspdData2.ReDraw = True

    End With
End Sub
'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData1.ReDraw = False
'     ggoSpread.SSSetProtected    C_PayCd , pvStartRow, pvEndRow
 '    ggoSpread.SSSetProtected    C_ResTaxAmt  , pvStartRow, pvEndRow
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
			C_DEPT_CD		= iCurColumnPos(1)
			C_ANN_SAL_NUM	= iCurColumnPos(2)
			C_ANN_SALARY	= iCurColumnPos(3)															
			C_MON_PAY_NUM	= iCurColumnPos(4)															
			C_MON_PAY		= iCurColumnPos(5)															
			C_DLY_WAGES_NUM = iCurColumnPos(6)															
			C_DLY_WAGES		= iCurColumnPos(7)															
			C_HOUR_WAGES_NUM= iCurColumnPos(8)															
			C_HOUR_WAGES	= iCurColumnPos(9)
			C_PAY_TOTAL		= iCurColumnPos(10)
			THIS_MON_PERSON     = iCurColumnPos(11)
			THIS_MON_PAY		= iCurColumnPos(12)
			THIS_MON_RETIRE		= iCurColumnPos(13)														
			THIS_MON_RETIRE_PAY = iCurColumnPos(14)
       Case "B"
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_DEPT_CD1			= iCurColumnPos(1)
			C_ANN_SAL_NUM1		= iCurColumnPos(2)
			C_ANN_SALARY1		= iCurColumnPos(3)															
			C_MON_PAY_NUM1		= iCurColumnPos(4)															
			C_MON_PAY1			= iCurColumnPos(5)															
			C_DLY_WAGES_NUM1	= iCurColumnPos(6)															
			C_DLY_WAGES1		= iCurColumnPos(7)															
			C_HOUR_WAGES_NUM1	= iCurColumnPos(8)															
			C_HOUR_WAGES1		= iCurColumnPos(9)
			C_PAY_TOTAL1		= iCurColumnPos(10)
			THIS_MON_PERSON1    = iCurColumnPos(11)
			THIS_MON_PAY1		= iCurColumnPos(12)
			THIS_MON_RETIRE1	= iCurColumnPos(13)														
			THIS_MON_RETIRE_PAY1 = iCurColumnPos(14)
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
End Sub


'========================================================================================
' Function Name : PopSaveSpreadColumnInf()
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
	dim temp
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet("A")      
    ggoSpread.Source = frm1.vspdData1
	Call ggoSpread.ReOrderingSpreadData()

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet("B")      
    ggoSpread.Source = frm1.vspdData2
	Call ggoSpread.ReOrderingSpreadData()

	temp = GetSpreadText(frm1.vspdData1,1,1,"X","X")

	if temp <>"" then
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.InsertRow
		Call InitData()
	end if  
End Sub

'========================================================================================================
'   Event Name : vspdData1_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
    
    ggoSpread.Source = frm1.vspdData2    
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )        
    Call GetSpreadColumnPos("B")    
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")

    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
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
    Call ggoOper.FormatDate(frm1.txtpay_yymm_dt, parent.gDateFormat, 2)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
    Call InitSpreadSheet("")                                                           'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables

    Call FuncGetAuth(gStrRequestMenuID, parent.gUsrID, lgUsrIntCd)                                ' 자료권한:lgUsrIntCd ("%", "1%")
    Call SetDefaultVal    
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
    Dim StrDt
    StrDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtpay_yymm_dt.Year, Right("0" & frm1.txtpay_yymm_dt.month , 2), "01")	
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
    If  txtEmp_no_Onchange() then
        Exit Function
    End If

    If  txtFr_Dept_cd_Onchange()  then
        Exit Function
    End If

    If  txtTo_Dept_cd_Onchange() then
        Exit Function
    End If
    If  txtprov_cd_Onchange() then
        Exit Function
    End If  

    If  txtPay_grd1_OnChange() then
        Exit Function
    End If      
    
    Dim Fr_dept_cd , To_dept_cd, rFrDept ,rToDept
    
    Fr_dept_cd = frm1.txtFr_internal_cd.value
    To_dept_cd = frm1.txtTo_internal_cd.value

    If fr_dept_cd = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,StrDt, rFrDept ,rToDept)
		
		frm1.txtFr_internal_cd.value = rFrDept
		frm1.txtFr_dept_nm.value = ""
	End If	
	
	If to_dept_cd = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,StrDt, rFrDept ,rToDept)
		frm1.txtTo_internal_cd.value = rToDept
		frm1.txtTo_dept_nm.value = ""
	End If  
    Fr_dept_cd = frm1.txtFr_internal_cd.value
    To_dept_cd = frm1.txtTo_internal_cd.value

    If (Fr_dept_cd<> "") AND (To_dept_cd<>"") Then       
  
        If Fr_dept_cd > To_dept_cd then
            Call DisplayMsgBox("800153","X","X","X")	'시작부서코드는 종료부서코드보다 작아야합니다.
            frm1.txtFr_internal_cd.value = ""
            frm1.txtTo_internal_cd.value = ""
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.activeElement
            Exit Function
        End IF 
        
    END IF   

    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("X")

	ggoSpread.Source = frm1.vspdData2
    ggoSpread.InsertRow
    call SetSpreadLock2
	
    If DbQuery = False Then  
		Exit Function
	End If

    FncQuery = True                                                              '☜: Processing is OK

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
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
    ggoSpread.Source = Frm1.vspdData1	
    ggoSpread.EditUndo  
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
	strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData1.MaxRows          '☜: Max fetched data
    
    Call RunMyBizASP(MyBizASP, strVal)                                               '☜:  Run biz logic

    DbQuery = True                                                                   '☜: Processing is NG

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
	
	frm1.vspdData1.focus	
End Function

'========================================================================================================
' Function Name : DbQueryNo
' Function Desc : Called by MB Area when query operation is not successful
'========================================================================================================
Function DbQueryNo()
	
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
End Function
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
	End If

	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtEmp_no.focus
		Exit Function
	Else
		frm1.txtEmp_no.value = arrRet(0)
		frm1.txtName.value = arrRet(1)
		frm1.txtEmp_no.focus
	End If	
			
End Function
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
        Else
            frm1.txtName.value = strName
        End if 
    End if  
    
End Function 

'========================================================================================================
'   Event Name : txtprov_cd_Onchange()            
'   Event Desc :
'========================================================================================================
function txtprov_cd_Onchange()
    Dim iDx
    Dim IntRetCd
    
    IF frm1.txtprov_cd.value = "" THEN
        frm1.txtprov_nm.value = ""
    ELSE
        IntRetCd = CommonQueryRs(" minor_nm "," b_minor "," major_cd = " & FilterVar("H0040", "''", "S") & " and minor_cd =  " & FilterVar(frm1.txtprov_cd.value , "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false then
			Call DisplayMsgBox("800140","X","X","X")	'지급내역코드에 등록되지 않은 코드입니다.
            frm1.txtprov_nm.value = ""
            frm1.txtprov_cd.focus
            txtprov_cd_Onchange = true
        ELSE    
            frm1.txtprov_nm.value = Trim(Replace(lgF0,Chr(11),""))   '수당코드 
        END IF
    END IF 

End Function


'======================================================================================================
'	Name : OpenCode()
'	Description : Code PopUp at vspdData
'=======================================================================================================
Function OpenCondAreaPopup(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then  
	   Exit Function
	End If   

	IsOpenPop = True
	Select Case iWhere
	   
        Case "2"
            arrParam(0) = "지급구분 팝업"			' 팝업 명칭 
	        arrParam(1) = "B_MINOR"				 		' TABLE 명칭 
	        arrParam(2) = frm1.txtprov_cd.value		    ' Code Condition
	        arrParam(3) = ""'frm1.txtprov_nm.value			' Name Cindition
	        arrParam(4) = " MAJOR_CD = " & FilterVar("H0040", "''", "S") & " AND MINOR_CD NOT IN (" & FilterVar("B", "''", "S") & " ," & FilterVar("C", "''", "S") & " ," & FilterVar("Z", "''", "S") & " )"    ' Where Condition							' Where Condition
	        arrParam(5) = "지급구분"			    ' TextBox 명칭 
	
            arrField(0) = "minor_cd"					' Field명(0)
            arrField(1) = "minor_nm"				    ' Field명(1)
    
            arrHeader(0) = "지급구분코드"				' Header명(0)
            arrHeader(1) = "지급구분명"
	
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
		
	
	If arrRet(0) = "" Then
		frm1.txtprov_cd.focus
		Exit Function
	Else
		Frm1.txtProv_Cd.value = arrRet(0)
		Frm1.txtProv_Nm.value = arrRet(1)		
		Frm1.txtProv_Cd.focus
	End If	
	
End Function
'========================================================================================================
' Name : OpenDept
' Desc : 부서 POPUP
'========================================================================================================
Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(2)
    Dim strBasDt 
    Dim StrDt
    StrDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtpay_yymm_dt.Year, Right("0" & frm1.txtpay_yymm_dt.month , 2), "01")	
    
	strBasDt = UNIGetLastDay(frm1.txtpay_yymm_dt.Text,Parent.gDateFormatYYYYMM)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then
		arrParam(0) = frm1.txtFr_dept_cd.value			            '  Code Condition
	ElseIf iWhere = 1 Then
		arrParam(0) = frm1.txtTo_dept_cd.value			            ' Code Condition
	End If
	
    arrParam(1) = strDt
	arrParam(2) = lgUsrIntCd                              ' 자료권한 Condition  

	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent, arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
		     Case "0"
               frm1.txtFr_dept_cd.focus
             Case "1"  
               frm1.txtTo_dept_cd.focus
        End Select	
		Exit Function
	Else
		Call SetDept(arrRet, iWhere)
	End If	
			
End Function
'------------------------------------------  SetDept()  ------------------------------------------------
'	Name : SetDept()
'	Description : Dept Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetDept(Byval arrRet, Byval iWhere)
		
	With frm1
		Select Case iWhere
		     Case "0"
               .txtFr_dept_cd.value = arrRet(0)
               .txtFr_dept_nm.value = arrRet(1)
               .txtFr_Internal_cd.value = arrRet(2)
               .txtFr_dept_cd.focus
             Case "1"  
               .txtTo_dept_cd.value = arrRet(0)
               .txtTo_dept_nm.value = arrRet(1) 
               .txtTo_Internal_cd.value = arrRet(2)
               .txtTo_dept_cd.focus
        End Select
	End With
End Function        		
'========================================================================================================
'   Event Name : txtFr_dept_cd_change
'   Event Desc :
'========================================================================================================
Function txtFr_dept_cd_Onchange()
    Dim IntRetCd
    Dim strDept_nm
    Dim StrDt
    StrDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtpay_yymm_dt.Year, Right("0" & frm1.txtpay_yymm_dt.month , 2), "01")	
    If frm1.txtFr_dept_cd.value = "" Then
		frm1.txtFr_dept_nm.value = ""
		frm1.txtFr_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtFr_dept_cd.value,StrDt,lgUsrIntCd, strDept_nm, lsInternal_cd)

        if  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call DisplayMsgBox("800012", "x","x","x")   '부서코드정보에 등록되지 않은 코드입니다.
            else
                Call DisplayMsgBox("800455", "x","x","x")   ' 자료권한이 없습니다.
            end if
		    frm1.txtFr_dept_nm.value = ""
		    frm1.txtFr_internal_cd.value = ""
            lsInternal_cd = ""
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.ActiveElement
       	    txtFr_dept_cd_Onchange = True
            Exit Function      
        else
            frm1.txtFr_dept_nm.value = strDept_nm
            frm1.txtFr_internal_cd.value = lsInternal_cd
        end if        
    End if  
End Function

'========================================================================================================
'   Event Name : txtTo_dept_cd_Onchange
'   Event Desc :
'========================================================================================================
Function txtTo_dept_cd_Onchange()
    Dim IntRetCd
    Dim strDept_nm
    Dim StrDt
    StrDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtpay_yymm_dt.Year, Right("0" & frm1.txtpay_yymm_dt.month , 2), "01")	    
    If frm1.txtTo_dept_cd.value = "" Then
		frm1.txtTo_dept_nm.value = ""
		frm1.txtTo_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtTo_dept_cd.value,StrDt,lgUsrIntCd,strDept_nm,lsInternal_cd)
        
        if  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call DisplayMsgBox("800012", "x","x","x")   '부서코드정보에 등록되지 않은 코드입니다.
            else
                Call DisplayMsgBox("800455", "x","x","x")   ' 자료권한이 없습니다.
            end if
		    frm1.txtTo_dept_nm.value = ""
		    frm1.txtTo_internal_cd.value = ""
            lsInternal_cd = ""
            frm1.txtTo_dept_cd.focus
            Set gActiveElement = document.ActiveElement
            txtTo_dept_cd_Onchange = true
            Exit Function      
        else
            frm1.txtTo_dept_nm.value = strDept_nm
            frm1.txtTo_internal_cd.value = lsInternal_cd
        end if
    End if  
End Function


'===========================================================================
' Function Name : OpenSItemDC
' Function Desc : OpenSItemDC Reference Popup
'===========================================================================
Function OpenSItemDC(Byval strCode, Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	    Case 1  ' 급호 
	    	arrParam(1) = "B_minor"				            	' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtPay_grd1.Value)	        ' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0001", "''", "S") & ""		    		' Where Condition
	    	arrParam(5) = "급호"		    				    ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)%>
    
	    	arrHeader(0) = "급호코드"			        		' Header명(0)%>
	    	arrHeader(1) = "급호명"	        					' Header명(1)%>
	End Select

    arrParam(3) = ""	
	arrParam(0) = arrParam(5)								    ' 팝업 명칭 

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPay_grd1.focus
		Exit Function
	Else
		frm1.txtPay_grd1.value = arrRet(0)
		frm1.txtPay_grd1_nm.value = arrRet(1)  
		frm1.txtPay_grd1.focus
	End If	
	
End Function
'======================================================================================================
'   Event Name : txtPay_grd_OnChange
'   Event Desc : 직급코드가 변경될 경우 
'=======================================================================================================
Function txtPay_grd1_OnChange()
    Dim IntRetCd

    If  frm1.txtPay_grd1.value = "" Then
        frm1.txtPay_grd1.Value=""
        frm1.txtPay_grd1_nm.Value=""
    Else
        IntRetCD =  CommonQueryRs(" minor_cd,minor_nm "," b_minor "," major_cd=" & FilterVar("H0001", "''", "S") & " And minor_cd =  " & FilterVar(frm1.txtPay_grd1.value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCD=False And Trim(frm1.txtPay_grd1.Value)<>""  Then
            Call  DisplayMsgBox("800054","X","X","X")                         '☜ : 등록되지 않은 코드입니다.
            frm1.txtPay_grd1_nm.Value=""
            txtPay_grd1_OnChange = True
        Else
            frm1.txtPay_grd1_nm.Value=Trim(Replace(lgF1,Chr(11),""))
        End If
    End If
End Function

'=======================================================================================================
'   Event Name : txtpay_yymm_dt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtpay_yymm_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtpay_yymm_dt.Action = 7
        frm1.txtpay_yymm_dt.focus
    End If
End Sub

Sub txtpay_yymm_dt_Keypress(KeyAscii)
	If KeyAscii = 13	Then Call MainQuery   'Call FncQuery()
End Sub
'========================================================================================================
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
        frm1.vspdData2.LeftCol=NewLeft   	   	
		Exit Sub
	End If

	If frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then

'		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
'			Call DisableToolBar(Parent.TBC_QUERY)
'			If DBQuery = False Then
'				Call RestoreToolBar()
'				Exit Sub
'			End If
'		End If
	End If  
End Sub

Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        frm1.vspdData1.LeftCol=NewLeft
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData1_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

    frm1.vspdData1.Col = pvCol1
    frm1.vspdData2.ColWidth(pvCol1) = frm1.vspdData1.ColWidth(pvCol1)

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

    frm1.vspdData2.Col = pvCol1
    frm1.vspdData1.ColWidth(pvCol1) = frm1.vspdData2.ColWidth(pvCol1)

    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub
'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_ScriptLeaveCell(ByVal Col , ByVal Row, ByVal newCol , ByVal newRow ,Cancel )
    frm1.vspdData2.Col = newCol
    frm1.vspdData2.Action = 0

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
'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc :
'==========================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0000000000")

    gMouseClickStatus = "SP1C" 

    Set gActiveSpdSheet = frm1.vspdData2
   
End Sub
'========================================================================================================
'   Event Name : vspdData1_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData1_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData1.MaxRows = 0 Then
        Exit Sub
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData2_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData2_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData2.MaxRows = 0 Then
        Exit Sub
    End If
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

Sub vspdData2_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SP1C" Then
       gMouseClickStatus = "SP1CR"
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
								<td background="../../../Cshared/Image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>급여구분별지급현황조회</font></td>
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
								<TD CLASS="TD6" NOWRAP><OBJECT classid=<%=gCLSIDFPDT%> id=txtpay_yymm_dt NAME="txtpay_yymm_dt"  CLASS=FPDTYYYYMM  title=FPDATETIME  ALT="급여년월" tag="12X1" VIEWASTEXT></OBJECT></TD>		
									<TD CLASS=TD5 NOWRAP>부서코드</TD>
									<TD CLASS=TD6 NOWRAP><INPUT ID="txtFr_dept_cd" NAME="txtFr_dept_cd" ALT="부서코드" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBP_CD" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenDept(0)">&nbsp;
									                     <INPUT ID="txtFr_dept_nm" NAME="txtFr_dept_nm" TYPE="Text" MAXLENGTH="50" SIZE=30 tag="14XXXU">&nbsp;~</TD>								
									                     <INPUT NAME="txtFr_Internal_cd" ALT="내부부서코드" TYPE="HIDDEN" SiZE=7 MAXLENGTH=7 tag="14XXXU">
								</TR>
							    <TR>								
	                            <TD CLASS="TD5" NOWRAP>지급구분</TD>
	                        	<TD CLASS="TD6" NOWRAP><INPUT NAME="txtProv_cd" MAXLENGTH="1" SIZE="10" ALT ="지급구분" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCondAreaPopup(2)">
	                        	                       <INPUT NAME="txtProv_nm" MAXLENGTH="20" SIZE="20" ALT ="지급구분명" tag="14XXXU"></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP><INPUT ID="txtTo_dept_cd" NAME="txtTo_dept_cd" ALT="" TYPE="Text" MAXLENGTH="18" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnITEM_CD" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenDept(1)">&nbsp;
									                     <INPUT ID="txtTo_dept_nm" NAME="txtTo_dept_nm" TYPE="Text" MAXLENGTH="40" SIZE=30 tag="14XXXU"></TD>	
									                     <INPUT NAME="txtTo_Internal_cd" ALT="내부부서코드" TYPE="HIDDEN" SiZE=7 MAXLENGTH=7 tag="14XXXU">

							    </TR>								
							    <TR>
									<TD CLASS=TD5 NOWRAP>사번</TD>
			     					<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no" ALT="사번" TYPE="Text" SiZE=15 MAXLENGTH=13 tag="11XXXU"><IMG SRC="../../../Cshared/Image/btnPopup.gif" NAME="btnEmpNo" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenEmpName('0')">
									                     <INPUT NAME="txtName" MAXLENGTH="30" SIZE="20"  ALT ="성명" tag="14XXXU"></TD>
					    			<TD CLASS="TD5" NOWRAP>급호</TD>
					    			<TD CLASS="TD6"><INPUT NAME="txtPay_grd1" ALT="급호" TYPE="Text" MAXLENGTH=2 SiZE=5 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSItemDC frm1.txtName.value,1">&nbsp;<INPUT NAME="txtPay_grd1_nm" TYPE="Text" MAXLENGTH=20 SIZE=20 tag="24"></TD>								</TR>
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
							<TD WIDTH="100%" HEIGHT=66%>
								<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData1 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"  id=vaSpread1>
									<PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0">
								</OBJECT>
							</TD>
						</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=44 VALIGN=TOP>  
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT="100%">
									<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"  id=vaSpread2>
										<PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0">
									</OBJECT>
								</TD>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
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
