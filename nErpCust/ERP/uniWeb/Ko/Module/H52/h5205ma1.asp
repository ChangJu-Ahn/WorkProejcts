<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        :
*  3. Program ID           : h5205ma1
*  4. Program Name         : 월대부상환현황 조회및 조정 
*  5. Program Desc         : 월대부상환현황 조회,수정 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/30
*  8. Modified date(Last)  : 2003/06/11
*  9. Modifier (First)     : TGS(CHUN HYUNG WON)
* 10. Modifier (Last)      : Lee SiNa
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables
'========================================================================================================
Const BIZ_PGM_ID      = "h5205mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_ID1     = "h5205mb2.asp"						           '☆: Biz Logic ASP Name
Const C_SHEETMAXROWS    = 21	                                      '☜: Visble row

'========================================================================================================
'=                       4.3 Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================

Dim gSelframeFlg                                                       '현재 TAB의 위치를 나타내는 Flag %>
Dim gblnWinEvent                                                       'ShowModal Dialog(PopUp) Window가 여러 개 뜨는 것을 방지하기 위해 
Dim lgBlnFlawChgFlg
Dim gtxtChargeType
Dim IsOpenPop
Dim lgOldRow
Dim lsInternal_cd

Dim C_EMP_NO
Dim C_EMP_NO_POP
Dim C_NAME
Dim C_HDD020T_BORW_CD
Dim C_HDD020T_BORW_NM_POP
Dim C_HDD020T_BORW_NM
Dim C_HDD020T_BORW_DT
Dim C_HDD020T_INTREST_TYPE
Dim C_HDD020T_INTREST_TYPE_NM
Dim C_HDD020T_PAY_INTCHNG_AMT
Dim C_HDD020T_INTREST_AMT
Dim C_HDD020T_BONUS_INTCHNG_AMT
Dim C_HDD020T_TOT_INTCHNG_AMT
Dim C_HDD020T_BORW_BALN_AMT
Dim C_HDD020T_BORW_TOT_AMT
Dim C_HDD020T_PAY_INTCHNG_CNT
Dim C_HDD020T_BONUS_INTCHNG_CNT

Dim C_EMP_NO2
Dim C_EMP_NO_POP2
Dim C_NAME2
Dim C_HDD020T_BORW_CD2
Dim C_HDD020T_BORW_NM_POP2
Dim C_HDD020T_BORW_NM2
Dim C_HDD020T_BORW_DT2
Dim C_HDD020T_INTREST_TYPE2
Dim C_HDD020T_INTREST_TYPE_NM2
Dim C_HDD020T_PAY_INTCHNG_AMT2
Dim C_HDD020T_INTREST_AMT2
Dim C_HDD020T_BONUS_INTCHNG_AMT2
Dim C_HDD020T_TOT_INTCHNG_AMT2
Dim C_HDD020T_BORW_BALN_AMT2
Dim C_HDD020T_BORW_TOT_AMT2
Dim C_HDD020T_PAY_INTCHNG_CNT2
Dim C_HDD020T_BONUS_INTCHNG_CNT2

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  

    C_EMP_NO = 1
    C_EMP_NO_POP = 2
    C_NAME = 3                                                        'Column constant for Spread Sheet
    C_HDD020T_BORW_CD = 4
    C_HDD020T_BORW_NM_POP = 5
    C_HDD020T_BORW_NM = 6
    C_HDD020T_BORW_DT = 7
    C_HDD020T_INTREST_TYPE = 8
    C_HDD020T_INTREST_TYPE_NM = 9
    C_HDD020T_PAY_INTCHNG_AMT = 10
    C_HDD020T_INTREST_AMT = 11
    C_HDD020T_BONUS_INTCHNG_AMT = 12
    C_HDD020T_TOT_INTCHNG_AMT = 13
    C_HDD020T_BORW_BALN_AMT = 14
    C_HDD020T_BORW_TOT_AMT = 15
    C_HDD020T_PAY_INTCHNG_CNT = 16
    C_HDD020T_BONUS_INTCHNG_CNT = 17
    
    C_EMP_NO2 = 1
    C_EMP_NO_POP2 = 2
    C_NAME2 = 3                                                        'Column constant for Spread Sheet
    C_HDD020T_BORW_CD2 = 4
    C_HDD020T_BORW_NM_POP2 = 5
    C_HDD020T_BORW_NM2 = 6
    C_HDD020T_BORW_DT2  = 7
    C_HDD020T_INTREST_TYPE2  = 8
    C_HDD020T_INTREST_TYPE_NM2 = 9
    C_HDD020T_PAY_INTCHNG_AMT2 = 10
    C_HDD020T_INTREST_AMT2 = 11
    C_HDD020T_BONUS_INTCHNG_AMT2 = 12
    C_HDD020T_TOT_INTCHNG_AMT2 = 13
    C_HDD020T_BORW_BALN_AMT2 = 14
    C_HDD020T_BORW_TOT_AMT2 = 15
    C_HDD020T_PAY_INTCHNG_CNT2 = 16
    C_HDD020T_BONUS_INTCHNG_CNT2 = 17

End Sub
'========================================================================================================
' Name : InitVariables()
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      =  parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
	lgOldRow = 0
	lsInternal_cd     = ""

	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False
End Sub

'========================================================================================================
' Name : SetDefaultVal()
' Desc : Set default value
'========================================================================================================

Sub SetDefaultVal()
    Dim strYear,strMonth,strDay
	frm1.txtIntchng_yymm_dt.focus
	Call  ExtractDateFrom("<%=GetSvrDate%>", parent.gServerDateFormat ,  parent.gServerDateType ,strYear,strMonth,strDay)	
	frm1.txtIntchng_yymm_dt.Year	=  strYear
	frm1.txtIntchng_yymm_dt.Month	=  strMonth
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

    lgKeyStream       = Trim(frm1.txtIntchng_yymm_dt.Year & Right("0" & frm1.txtIntchng_yymm_dt.Month,2)) & parent.gColSep       'You Must append one character( parent.gColSep)
    lgKeyStream       = lgKeyStream & Trim(frm1.txtBorw_cd.value) & parent.gColSep
    lgKeyStream       = lgKeyStream & Trim(frm1.txtEmp_no.value) & parent.gColSep
    lgKeyStream       = lgKeyStream & Trim(frm1.txtIntrest_type.value) & parent.gColSep
    If  lsInternal_cd = "" then
        lgKeyStream = lgKeyStream & lgUsrIntCd & parent.gColSep
    Else
        lgKeyStream = lgKeyStream & lsInternal_cd & parent.gColSep
    End If
End Sub

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr
    Dim iNameArr

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("h0044", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) '이자구분 
    iCodeArr = lgF0
    iNameArr = lgF1
     ggoSpread.Source = frm1.vspdData
     ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_HDD020T_INTREST_TYPE
     ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_HDD020T_INTREST_TYPE_NM
     
    Call  SetCombo2(frm1.txtintrest_type,iCodeArr, iNameArr,Chr(11))
End Sub

Sub InitSpreadComboBox()
    Dim iCodeArr
    Dim iNameArr

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("h0044", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) '이자구분 
    iCodeArr = lgF0
    iNameArr = lgF1
     ggoSpread.Source = frm1.vspdData
     ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_HDD020T_INTREST_TYPE
     ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_HDD020T_INTREST_TYPE_NM
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)

	Call initSpreadPosVariables()   'sbk 

    If pvSpdNo = "" OR pvSpdNo = "A" Then

    	With frm1.vspdData

            ggoSpread.Source = frm1.vspdData
            ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    'sbk

    	   .ReDraw = false

           .MaxCols   = C_HDD020T_BONUS_INTCHNG_CNT + 1                                       ' ☜:☜: Add 1 to Maxcols
    	   .Col       = .MaxCols                                                              ' ☜:☜: Hide maxcols
           .ColHidden = True                                                                  ' ☜:☜:

           .MaxRows = 0
            ggoSpread.ClearSpreadData

           Call GetSpreadColumnPos("A") 'sbk

           Call  AppendNumberPlace("6","6","0")
    	  
                 ggoSpread.SSSetEdit     C_EMP_NO,                       "사번", 12,,,13,2
                 ggoSpread.SSSetButton   C_EMP_NO_POP
                 ggoSpread.SSSetEdit     C_NAME,                         "성명", 12,,,,2
                 ggoSpread.SSSetEdit     C_HDD020T_BORW_CD,              "대부코드",8,,,10,2
                 ggoSpread.SSSetButton   C_HDD020T_BORW_NM_POP
                 ggoSpread.SSSetEdit     C_HDD020T_BORW_NM,              "대부코드명",10,,,50,2
                 ggoSpread.SSSetDate     C_HDD020T_BORW_DT,              "대부일", 10,2,  parent.gDateFormat
                 ggoSpread.SSSetCombo    C_HDD020T_INTREST_TYPE,         "",10
                 ggoSpread.SSSetCombo    C_HDD020T_INTREST_TYPE_NM,      "이자구분",10,,false
                 ggoSpread.SSSetFloat    C_HDD020T_PAY_INTCHNG_AMT,      "급여상환액", 12, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
                 ggoSpread.SSSetFloat    C_HDD020T_INTREST_AMT,          "이자" ,12, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
                 ggoSpread.SSSetFloat    C_HDD020T_BONUS_INTCHNG_AMT,    "상여상환액" ,12, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
                 ggoSpread.SSSetFloat    C_HDD020T_TOT_INTCHNG_AMT,      "상환총액" ,15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
                 ggoSpread.SSSetFloat    C_HDD020T_BORW_BALN_AMT,        "잔액" ,12, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
                 ggoSpread.SSSetFloat    C_HDD020T_BORW_TOT_AMT,         "대부총액" ,15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
                 ggoSpread.SSSetFloat    C_HDD020T_PAY_INTCHNG_CNT,      "급여상환횟수" ,12,"6", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
                 ggoSpread.SSSetFloat    C_HDD020T_BONUS_INTCHNG_CNT,    "상여상환횟수" ,12,"6", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"

            Call ggoSpread.MakePairsColumn(C_EMP_NO,C_EMP_NO_POP)    'sbk
            Call ggoSpread.MakePairsColumn(C_HDD020T_BORW_CD,C_HDD020T_BORW_NM_POP)    'sbk
            Call ggoSpread.SSSetColHidden(C_HDD020T_INTREST_TYPE,C_HDD020T_INTREST_TYPE,True)

    	   .ReDraw = true

        End With
    End If

    If pvSpdNo = "" OR pvSpdNo = "B" Then
    	With frm1.vspdData2

            ggoSpread.Source = frm1.vspdData2
            ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    'sbk

    	   .ReDraw = false

           .MaxCols   = C_HDD020T_BONUS_INTCHNG_CNT2 + 1                                      ' ☜:☜: Add 1 to Maxcols
    	   .Col       = .MaxCols                                                              ' ☜:☜: Hide maxcols
           .ColHidden = True                                                                  ' ☜:☜:

           .MaxRows = 0
            ggoSpread.ClearSpreadData

           .DisplayColHeaders = False

           Call GetSpreadColumnPos("B") 'sbk

           Call  AppendNumberPlace("6","6","0")

                 ggoSpread.SSSetEdit     C_EMP_NO2,                       "", 12,,,13,2
                 ggoSpread.SSSetButton   C_EMP_NO_POP2
                 ggoSpread.SSSetEdit     C_NAME2,                         "", 12,,,,2
                 ggoSpread.SSSetEdit     C_HDD020T_BORW_CD2,              "",8,,,50,2
                 ggoSpread.SSSetButton   C_HDD020T_BORW_NM_POP2
                 ggoSpread.SSSetEdit     C_HDD020T_BORW_NM2,              "",10
                 ggoSpread.SSSetDate     C_HDD020T_BORW_DT2,              "", 10,2,  parent.gDateFormat
                 ggoSpread.SSSetEdit     C_HDD020T_INTREST_TYPE2,         "",10
                 ggoSpread.SSSetEdit     C_HDD020T_INTREST_TYPE_NM2,      "",10,,false
                 ggoSpread.SSSetFloat    C_HDD020T_PAY_INTCHNG_AMT2,      "급여상환액", 12, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
                 ggoSpread.SSSetFloat    C_HDD020T_INTREST_AMT2,          "이자" ,12, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
                 ggoSpread.SSSetFloat    C_HDD020T_BONUS_INTCHNG_AMT2,    "상여상환액" ,12, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
                 ggoSpread.SSSetFloat    C_HDD020T_TOT_INTCHNG_AMT2,      "" ,15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
                 ggoSpread.SSSetFloat    C_HDD020T_BORW_BALN_AMT2,        "" ,12, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
                 ggoSpread.SSSetFloat    C_HDD020T_BORW_TOT_AMT2,         "" ,15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
                 ggoSpread.SSSetFloat    C_HDD020T_PAY_INTCHNG_CNT2,      "" ,12,"6", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
                 ggoSpread.SSSetFloat    C_HDD020T_BONUS_INTCHNG_CNT2,    "" ,12,"6", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
    	   .ReDraw = true

            Call ggoSpread.MakePairsColumn(C_EMP_NO2,C_EMP_NO_POP2)    'sbk
            Call ggoSpread.MakePairsColumn(C_HDD020T_BORW_CD2,C_HDD020T_BORW_NM_POP2)    'sbk
            Call ggoSpread.SSSetColHidden(C_HDD020T_INTREST_TYPE2,C_HDD020T_INTREST_TYPE2,True)

        End With
    End If

    Call SetSpreadLock

End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
     ggoSpread.Source = frm1.vspdData
    With frm1
    .vspdData.ReDraw = False
         ggoSpread.SpreadLock    C_EMP_NO, -1, C_EMP_NO, -1
         ggoSpread.SpreadLock    C_EMP_NO_POP, -1, C_EMP_NO_POP, -1
         ggoSpread.SpreadLock    C_NAME, -1, C_NAME, -1
         ggoSpread.SpreadLock    C_HDD020T_BORW_CD, -1, C_HDD020T_BORW_CD, -1
         ggoSpread.SpreadLock    C_HDD020T_BORW_NM, -1, C_HDD020T_BORW_NM, -1
         ggoSpread.SpreadLock    C_HDD020T_BORW_NM_POP, -1, C_HDD020T_BORW_NM_POP, -1
         ggoSpread.SpreadLock    C_HDD020T_BORW_DT, -1, C_HDD020T_BORW_DT, -1
         ggoSpread.SpreadLock    C_HDD020T_INTREST_TYPE, -1, C_HDD020T_INTREST_TYPE, -1
         ggoSpread.SpreadLock    C_HDD020T_INTREST_TYPE_NM, -1, C_HDD020T_INTREST_TYPE_NM, -1
         ggoSpread.SSSetRequired	C_HDD020T_PAY_INTCHNG_AMT, -1, -1
         ggoSpread.SSSetRequired	C_HDD020T_INTREST_AMT, -1, -1
         ggoSpread.SSSetRequired	C_HDD020T_BONUS_INTCHNG_AMT, -1, -1
         ggoSpread.SSSetRequired	C_HDD020T_TOT_INTCHNG_AMT, -1, -1
         ggoSpread.SSSetRequired	C_HDD020T_BORW_BALN_AMT, -1, -1
         ggoSpread.SSSetRequired	C_HDD020T_BORW_TOT_AMT, -1, -1
         ggoSpread.SSSetProtected   .vspdData.MaxCols   , -1, -1
    .vspdData.ReDraw = True

    End With

     ggoSpread.Source = frm1.vspdData2
    With frm1.vspdData2
    .ReDraw = False
         ggoSpread.SpreadLock    C_EMP_NO2, -1, C_EMP_NO2, -1
         ggoSpread.SpreadLock    C_EMP_NO_POP2, -1, C_EMP_NO_POP2, -1
         ggoSpread.SpreadLock    C_NAME2, -1, C_NAME2, -1
         ggoSpread.SpreadLock    C_HDD020T_BORW_CD2, -1, C_HDD020T_BORW_CD2, -1
         ggoSpread.SpreadLock    C_HDD020T_BORW_NM2, -1, C_HDD020T_BORW_NM2, -1
         ggoSpread.SpreadLock    C_HDD020T_BORW_NM_POP2, -1, C_HDD020T_BORW_NM_POP2, -1
         ggoSpread.SpreadLock    C_HDD020T_BORW_DT2, -1, C_HDD020T_BORW_DT2, -1
         ggoSpread.SpreadLock    C_HDD020T_INTREST_TYPE2, -1, C_HDD020T_INTREST_TYPE2, -1
         ggoSpread.SpreadLock    C_HDD020T_INTREST_TYPE_NM2, -1, C_HDD020T_INTREST_TYPE_NM2, -1
         ggoSpread.SpreadLock    C_HDD020T_PAY_INTCHNG_AMT2, -1, C_HDD020T_PAY_INTCHNG_AMT2, -1
         ggoSpread.SpreadLock    C_HDD020T_INTREST_AMT2, -1, C_HDD020T_INTREST_AMT2, -1
         ggoSpread.SpreadLock    C_HDD020T_BONUS_INTCHNG_AMT2, -1, C_HDD020T_BONUS_INTCHNG_AMT2, -1
         ggoSpread.SpreadLock    C_HDD020T_TOT_INTCHNG_AMT2, -1, C_HDD020T_TOT_INTCHNG_AMT2, -1
         ggoSpread.SpreadLock    C_HDD020T_BORW_BALN_AMT2, -1, C_HDD020T_BORW_BALN_AMT2, -1
         ggoSpread.SpreadLock    C_HDD020T_BORW_TOT_AMT2, -1, C_HDD020T_BORW_TOT_AMT2, -1
         ggoSpread.SpreadLock    C_HDD020T_PAY_INTCHNG_CNT2, -1, C_HDD020T_PAY_INTCHNG_CNT2, -1
         ggoSpread.SpreadLock    C_HDD020T_BONUS_INTCHNG_CNT2, -1, C_HDD020T_BONUS_INTCHNG_CNT2, -1
         ggoSpread.SSSetProtected   .MaxCols   , -1, -1
    .ReDraw = True

    End With

End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1.vspdData
    .ReDraw = False
         ggoSpread.SSSetRequired		C_EMP_NO, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected       C_NAME , pvStartRow, pvEndRow
         ggoSpread.SSSetRequired        C_HDD020T_BORW_CD , pvStartRow, pvEndRow
         ggoSpread.SSSetProtected       C_HDD020T_BORW_NM , pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_HDD020T_BORW_DT, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_HDD020T_INTREST_TYPE_NM, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_HDD020T_PAY_INTCHNG_AMT, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_HDD020T_INTREST_AMT, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_HDD020T_BONUS_INTCHNG_AMT, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_HDD020T_TOT_INTCHNG_AMT, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_HDD020T_BORW_BALN_AMT, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_HDD020T_BORW_TOT_AMT, pvStartRow, pvEndRow
    .ReDraw = True
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

            C_EMP_NO = iCurColumnPos(1)
            C_EMP_NO_POP = iCurColumnPos(2)
            C_NAME = iCurColumnPos(3)
            C_HDD020T_BORW_CD = iCurColumnPos(4)
            C_HDD020T_BORW_NM_POP = iCurColumnPos(5)
            C_HDD020T_BORW_NM = iCurColumnPos(6)
            C_HDD020T_BORW_DT  = iCurColumnPos(7)
            C_HDD020T_INTREST_TYPE  = iCurColumnPos(8)
            C_HDD020T_INTREST_TYPE_NM = iCurColumnPos(9)
            C_HDD020T_PAY_INTCHNG_AMT = iCurColumnPos(10)
            C_HDD020T_INTREST_AMT = iCurColumnPos(11)
            C_HDD020T_BONUS_INTCHNG_AMT = iCurColumnPos(12)
            C_HDD020T_TOT_INTCHNG_AMT = iCurColumnPos(13)
            C_HDD020T_BORW_BALN_AMT = iCurColumnPos(14)
            C_HDD020T_BORW_TOT_AMT = iCurColumnPos(15)
            C_HDD020T_PAY_INTCHNG_CNT = iCurColumnPos(16)
            C_HDD020T_BONUS_INTCHNG_CNT = iCurColumnPos(17)
            
        Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
                                                                   
            C_EMP_NO2 = iCurColumnPos(1)
            C_EMP_NO_POP2 = iCurColumnPos(2)
            C_NAME2 = iCurColumnPos(3)
            C_HDD020T_BORW_CD2 = iCurColumnPos(4)
            C_HDD020T_BORW_NM_POP2 = iCurColumnPos(5)
            C_HDD020T_BORW_NM2 = iCurColumnPos(6)
            C_HDD020T_BORW_DT2  = iCurColumnPos(7)
            C_HDD020T_INTREST_TYPE2  = iCurColumnPos(8)
            C_HDD020T_INTREST_TYPE_NM2 = iCurColumnPos(9)
            C_HDD020T_PAY_INTCHNG_AMT2 = iCurColumnPos(10)
            C_HDD020T_INTREST_AMT2 = iCurColumnPos(11)
            C_HDD020T_BONUS_INTCHNG_AMT2 = iCurColumnPos(12)
            C_HDD020T_TOT_INTCHNG_AMT2 = iCurColumnPos(13)
            C_HDD020T_BORW_BALN_AMT2 = iCurColumnPos(14)
            C_HDD020T_BORW_TOT_AMT2 = iCurColumnPos(15)
            C_HDD020T_PAY_INTCHNG_CNT2 = iCurColumnPos(16)
            C_HDD020T_BONUS_INTCHNG_CNT2 = iCurColumnPos(17)
            
    End Select    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtIntchng_yymm_dt,  parent.gDateFormat, 2)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field

    Call InitSpreadSheet("")                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables

    Call FuncGetAuth(gStrRequestMenuID,  parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
    Call SetDefaultVal
	Call SetToolbar("1100110100101111")												'⊙: Set ToolBar

    Call InitComboBox
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
    Dim RetStatus
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt

    FncQuery = False                                                            '☜: Processing is NG

    Err.Clear                                                                   '☜: Protect system from crashing

     ggoSpread.Source = Frm1.vspdData
    If lgBlnFlgChgValue = True Or  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call  ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call InitVariables															'⊙: Initializes local global variables

    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If

    If txtEmp_no_Onchange() then
        Exit Function
    End If

    If txtBorw_cd_OnChange() Then
        Exit Function
    End If

    Call MakeKeyStream("X")

    Call  DisableToolBar( parent.TBC_QUERY)
	If DbQuery = False Then
		Call  RestoreToolBar()
		Exit Function
	End if														'☜: Query db data

    FncQuery = True																'☜: Processing is OK


End Function

'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD

    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If lgBlnFlgChgValue = True Then
       IntRetCD =  DisplayMsgBox("900015",  parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to make it new?
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If

    Call  ggoOper.ClearField(Document, "A")                                       '☜: Clear Condition Field
    Call  ggoOper.LockField(Document , "N")                                       '☜: Lock  Field
	Call SetToolbar("1100111100111111")							                 '⊙: Set ToolBar
    Call InitVariables                                                           '⊙: Initializes local global variables

    Set gActiveElement = document.ActiveElement

    FncNew = True																 '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim IntRetCd

    FncDelete = False                                                             '☜: Processing is NG

    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                            'Check if there is retrived data
        Call  DisplayMsgBox("900002","X","X","X")                                  '☜: Please do Display first.
        Exit Function
    End If

    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"X","X")		                  '☜: Do you want to delete?
	If IntRetCD = vbNo Then
		Exit Function
	End If

    If DbDelete= False Then
       Exit Function
    End If
												                  '☜: Delete db data

    FncDelete=  True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD
    Dim iRow
    Dim dblintchng_amt,dblBorw_baln,dblBorw_tot_amt
    Dim dblPay_intchng, dblBonus_intchng
    Dim intPay_int_cnt,intBonus_int_cnt,strIntrest_type,dblintrest_amt
    Dim strIntchng_yymm_dt, strBrow_dt
    
    FncSave = False                                                              '☜: Processing is NG

    Err.Clear                                                                    '☜: Clear err status

     ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And  ggoSpread.SSCheckChange = False Then
        IntRetCD =  DisplayMsgBox("900001","X","X","X")                           '⊙: No data changed!!
        Exit Function
    End If

    If Not chkField(Document, "2") Then
       Exit Function
    End If

	 ggoSpread.Source = frm1.vspdData
    If Not  ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
       Exit Function
    End If

    If frm1.txtIntchng_yymm_dt.Text="" Then
        Call  DisplayMsgBox("970021","X",frm1.txtIntchng_yymm_dt.alt,"X")        '상환년월은 필수 입력사항입니다.
        frm1.txtIntchng_yymm_dt.focus ' go to
        Exit Function
    End If
    
    With Frm1.vspdData
        For iRow = 1 To  .MaxRows
            .Row = iRow
            .Col = 0
           Select Case .Text
               Case  ggoSpread.InsertFlag, ggoSpread.UpdateFlag
					.Col = C_NAME

					If IsNull(Trim(.Text)) OR Trim(.Text) = "" Then
					    Call DisplayMsgBox("800048","X","X","X")
						Exit Function
					end if
					
					.Col = C_HDD020T_BORW_NM

					If IsNull(Trim(.Text)) OR Trim(.Text) = "" Then
						Call DisplayMsgBox("970000", "X","대부코드","x")
						Exit Function
					end if
					               

                    strIntchng_yymm_dt =  UNIGetFirstDay(frm1.txtIntchng_yymm_dt.Text, parent.gDateFormatYYYYMM)
                    strIntchng_yymm_dt = Mid( UniConvDateToYYYYMMDD(strIntchng_yymm_dt,  parent.gDateFormat, ""),1,6)
                    
                    .Col = C_HDD020T_BORW_DT
                    strBrow_dt = Mid( UniConvDateToYYYYMMDD(.Text,  parent.gDateFormat, ""),1,6)
                    
                    If strIntchng_yymm_dt < strBrow_dt Then
	                    Call  DisplayMsgBox("970025","X","대부일",frm1.txtIntchng_yymm_dt.alt)	'
  	                    .Action=0
                        Set gActiveElement = document.activeElement
                        Exit Function
   	                End If
                    
   	                .Col = C_HDD020T_PAY_INTCHNG_AMT    '급여상환액 
                    dblPay_intchng  =  UNICDbl(.Text)
   	                .Col = C_HDD020T_BONUS_INTCHNG_AMT  '상여상환액 
                    dblBonus_intchng =  UNICDbl(.Text)

   	                .Col = C_HDD020T_TOT_INTCHNG_AMT    '상환총액 
                    dblintchng_amt  =  UNICDbl(.Text)
   	                .Col = C_HDD020T_BORW_BALN_AMT      '잔액 
                    dblBorw_baln  =  UNICDbl(.Text)
   	                .Col = C_HDD020T_BORW_TOT_AMT       '대부총액 
                    dblBorw_tot_amt  =  UNICDbl(.Text)

   	                If ((dblintchng_amt = 0) And (dblBorw_baln = 0) And (dblBorw_tot_amt =0)) Then
	                    Call  DisplayMsgBox("800132","X","X","X")	'상환총액, 잔액, 대부총액을 확인하십시오.
	                    .Row = iRow
  	                    .Col = C_HDD020T_BORW_TOT_AMT
  	                    .Action=0
                        Set gActiveElement = document.activeElement
                        Exit Function
   	                End If

   	                If (dblPay_intchng + dblBonus_intchng) > dblintchng_amt Then
	                    Call  DisplayMsgBox("800129","X","X","X")	'급여/상여 상환액을 조정하시오.
	                    .Row = iRow
  	                    .Col = C_HDD020T_PAY_INTCHNG_AMT
  	                    .Action=0
                        Set gActiveElement = document.activeElement
                        Exit Function
   	                End If
   	                
   	                If (dblintchng_amt + dblBorw_baln) <> dblBorw_tot_amt Then
	                    Call  DisplayMsgBox("800132","X","X","X")	'상환총액, 잔액, 대부총액을 확인하십시오.
	                    .Row = iRow
  	                    .Col = C_HDD020T_BORW_TOT_AMT
  	                    .Action=0
                        Set gActiveElement = document.activeElement
                        Exit Function
   	                End If

   	                .Col = C_HDD020T_PAY_INTCHNG_CNT
                    intPay_int_cnt  =  UNICDbl(.Text)
   	                .Col = C_HDD020T_BONUS_INTCHNG_CNT
                    intBonus_int_cnt  =  UNICDbl(.Text)

   	                If dblPay_intchng > 0 And intPay_int_cnt = 0 Then
	                    Call  DisplayMsgBox("800483","X","급여","X")	'
	                    .Row = iRow
  	                    .Col = C_HDD020T_PAY_INTCHNG_CNT
  	                    .Action=0
                        Set gActiveElement = document.activeElement
                        Exit Function
   	                End If

   	                If dblBonus_intchng > 0 And intBonus_int_cnt = 0 Then
	                    Call  DisplayMsgBox("800483","X","상여","X")	'상여상환횟수가 0이므로 상여상환액을 입력할 수 없습니다.
	                    .Row = iRow
  	                    .Col = C_HDD020T_BONUS_INTCHNG_CNT
  	                    .Action=0
                        Set gActiveElement = document.activeElement
                        Exit Function
   	                End If

	            	If intPay_int_cnt < 0 Or intBonus_int_cnt < 0 Then
	                    Call  DisplayMsgBox("800138","X","X","X")	'상환횟수는 음수일수 없습니다.
	                    .Row = iRow
  	                    .Col = C_HDD020T_PAY_INTCHNG_CNT
  	                    .Action=0
                        Set gActiveElement = document.activeElement
                        Exit Function
       	            End If

   	                .Col = C_HDD020T_INTREST_TYPE
                    strIntrest_type  = .Text
   	                .Col = C_HDD020T_INTREST_AMT
                    dblintrest_amt   =  UNICDbl(.Text)

		            If IsNull(dblintrest_amt) Then
		                dblIntrest_rate = 0
		            End If
		                                                                    '이자구분이 Y인 경우 이자율은 zero면 에러 
		            If strIntrest_type = "Y" And dblintrest_amt = 0 then   '이자구분이 N인 경우 이자율은 zero보다 커야함 
                        Call  DisplayMsgBox("141157","X","X","X")   '이자율을 입력하십시오.
                        .Action = 0 ' go to
                        Set gActiveElement = document.activeElement
          	            Exit Function
		            Else
		                If strIntrest_type = "N" And dblintrest_amt > 0 then
                            Call  DisplayMsgBox("800235","X","X","X")	'이자율을 입력할 수 없습니다.
                            .Action = 0 ' go to
                            Set gActiveElement = document.activeElement
                            Exit Function
                        End If
 		            End If
           End Select
        Next
    End With

    Call MakeKeyStream("X")

    If DbSave = False Then
       Exit Function
    End If
				                                                    '☜: Save db data

    FncSave = True                                                              '☜: Processing is OK

End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()

    If gMouseClickStatus <> "SP1C" Then
        If Frm1.vspdData.MaxRows < 1 Then
           Exit Function
        End If

	    With frm1.vspdData

	    	If .ActiveRow > 0 Then
	    		.ReDraw = False

	    		 ggoSpread.Source = frm1.vspdData
	    		 ggoSpread.CopyRow
                 SetSpreadColor .ActiveRow, .ActiveRow

               .Row  = .ActiveRow
                .Col  = C_NAME
                .Text = ""
                .Col  = C_EMP_NO
                .Text = ""
                .Col  = C_HDD020T_BORW_CD
                .Text = ""
                .Col  = C_HDD020T_BORW_NM
                .Text = ""
                .Col  = C_HDD020T_BORW_DT
                .Text = ""
                .Col  = C_HDD020T_INTREST_TYPE
                .Text = ""
                .Col  = C_HDD020T_INTREST_TYPE_NM
                .Text = ""

	    		.ReDraw = True
	    		.focus
	    	End If
	    End With
	End If
	
    Set gActiveElement = document.ActiveElement

End Function


'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel()
     ggoSpread.Source = frm1.vspdData
     ggoSpread.EditUndo
    Call  initData()

End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)

    Dim imRow, iCnt

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    If gMouseClickStatus <> "SP1C" Then
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
            .vspdData.ReDraw = False
            .vspdData.focus
             ggoSpread.Source = .vspdData
             ggoSpread.InsertRow .vspdData.ActiveRow, imRow
             SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1

            For iCnt = 1 To imRow 
                .vspdData.Row = .vspdData.ActiveRow + iCnt - 1
                
                .vspdData.col=C_HDD020T_BORW_DT                             'On the envent Insert Row, set defalut value
                .vspdData.text= UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
                .vspdData.col=C_HDD020T_PAY_INTCHNG_AMT
                .vspdData.value=0
                .vspdData.col=C_HDD020T_INTREST_AMT
                .vspdData.value=0
                .vspdData.col=C_HDD020T_BONUS_INTCHNG_AMT
                .vspdData.value=0
                .vspdData.col=C_HDD020T_TOT_INTCHNG_AMT
                .vspdData.value=0
                .vspdData.col=C_HDD020T_BORW_BALN_AMT
                .vspdData.value=0
                .vspdData.col=C_HDD020T_BORW_TOT_AMT
                .vspdData.value=0
                .vspdData.col=C_HDD020T_PAY_INTCHNG_CNT
                .vspdData.value=0
                .vspdData.col=C_HDD020T_BONUS_INTCHNG_CNT
                .vspdData.value=0
            Next

           .vspdData.ReDraw = True
        End With
    End If
    
    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows
    If Frm1.vspdData.MaxRows < 1 then
       Exit function
	End if
    With Frm1.vspdData
    	.focus
    	 ggoSpread.Source = frm1.vspdData
    	lDelRows =  ggoSpread.DeleteRow
    End With
    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                    '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncPrev
' Desc : developer describe this line Called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev()

    Dim strVal
    Dim IntRetCD

    FncPrev = False                                                              '☜: Processing is OK
    Err.Clear                                                                    '☜: Clear err status

    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                           '☜: Please do Display first.
        Call  DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If

	If lgBlnFlgChgValue = True Then
		IntRetCD =  DisplayMsgBox("900017",  parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    Call  ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area

    Call InitVariables														 '⊙: Initializes local global variables

    If LayerShowHide(1) = false Then
        Exit Function
    End If

    Call MakeKeyStream("X")

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
    strVal = strVal     & "&txtPrevNext="        & "P"	                         '☆: Direction

	Call RunMyBizASP(MyBizASP, strVal)										     '☜: Run Biz

    FncPrev = True                                                               '☜: Processing is OK

End Function
'========================================================================================================
' Name : FncNext
' Desc : developer describe this line Called by MainNext in Common.vbs
'========================================================================================================
Function FncNext()
    Dim strVal
    Dim IntRetCD

    FncNext = False                                                              '☜: Processing is OK
    Err.Clear                                                                    '☜: Clear err status

    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                           '☜: Please do Display first.
        Call  DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If

	If lgBlnFlgChgValue = True Then
		IntRetCD =  DisplayMsgBox("900017",  parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    Call  ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area

    Call InitVariables														 '⊙: Initializes local global variables

    If LayerShowHide(1) = false Then
        Exit Function
    End If

    Call MakeKeyStream("X")

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
    strVal = strVal     & "&txtPrevNext="        & "N"	                         '☆: Direction

	Call RunMyBizASP(MyBizASP, strVal)										     '☜: Run Biz

    FncNext = True                                                               '☜: Processing is OK

End Function
'========================================================================================================
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel()
	Call Parent.FncExport( parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind()
	Call Parent.FncFind( parent.C_SINGLE, True)
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

    ggoSpread.Source = frm1.vspdData2 
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub


'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
    
    If isEmpty(TypeName(gActiveSpdSheet)) Then
		Exit Sub
	Elseif	UCase(gActiveSpdSheet.id) = "VASPREAD" Then
		ggoSpread.Source = frm1.vspdData2 
		Call ggoSpread.SaveSpreadColumnInf()
	End if

End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet("A")      
    Call InitSpreadComboBox
    ggoSpread.Source = frm1.vspdData
	Call ggoSpread.ReOrderingSpreadData()

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet("B")      
    ggoSpread.Source = frm1.vspdData2
	Call ggoSpread.ReOrderingSpreadData()

    Frm1.vspdData2.Col = 0
    Frm1.vspdData2.Text = "합계"

	Call InitData()
End Sub

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()

	Dim IntRetCD
	FncExit = False
	 ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900016",  parent.VB_YES_NO,"X","X")			 '⊙: Data is changed.  Do you want to exit?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True

End Function

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

    If LayerShowHide(1) = false Then
        Exit Function
    End If

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
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

	DbSave = False														         '☜: Processing is NG

    If LayerShowHide(1) = false Then
        Exit Function
    End If

	With frm1
		.txtMode.value        =  parent.UID_M0002                                        '☜: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1

       For lRow = 1 To .vspdData.MaxRows

           .vspdData.Row = lRow
           .vspdData.Col = 0

           Select Case .vspdData.Text
              Case  ggoSpread.InsertFlag                                       '☜: Insert
                                                                   strVal = strVal & "C" & parent.gColSep
                                                                   strVal = strVal & lRow & parent.gColSep
                                                                   strVal = strVal & Trim(.txtIntchng_yymm_dt.Year & Right("0" & .txtIntchng_yymm_dt.Month,2)) & parent.gColSep
                    .vspdData.Col = C_EMP_NO                     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD020T_BORW_CD	         : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD020T_BORW_DT            : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD020T_INTREST_TYPE	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD020T_PAY_INTCHNG_AMT	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD020T_INTREST_AMT	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD020T_BONUS_INTCHNG_AMT	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD020T_TOT_INTCHNG_AMT	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD020T_BORW_BALN_AMT	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD020T_BORW_TOT_AMT	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD020T_PAY_INTCHNG_CNT	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD020T_BONUS_INTCHNG_CNT  : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                                   strVal = strVal & "U" & parent.gColSep
                                                                   strVal = strVal & lRow & parent.gColSep
                                                                   strVal = strVal & Trim(.txtIntchng_yymm_dt.Year & Right("0" & .txtIntchng_yymm_dt.Month,2)) & parent.gColSep
                    .vspdData.Col = C_EMP_NO                     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD020T_BORW_CD	         : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD020T_BORW_DT            : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD020T_INTREST_TYPE	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD020T_PAY_INTCHNG_AMT	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD020T_INTREST_AMT	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD020T_BONUS_INTCHNG_AMT	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD020T_TOT_INTCHNG_AMT	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD020T_BORW_BALN_AMT	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD020T_BORW_TOT_AMT	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD020T_PAY_INTCHNG_CNT	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD020T_BONUS_INTCHNG_CNT  : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.DeleteFlag                                      '☜: Delete

                                                                   strDel = strDel & "D" & parent.gColSep
                                                                   strDel = strDel & lRow & parent.gColSep
                                                                   strDel = strDel & Trim(.txtIntchng_yymm_dt.Year & Right("0" & .txtIntchng_yymm_dt.Month,2)) & parent.gColSep
                    .vspdData.Col = C_EMP_NO                     : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD020T_BORW_CD	         : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD020T_BORW_DT            : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD020T_INTREST_TYPE	     : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next

	   .txtMaxRows.value     = lGrpCnt-1
	   .txtSpread.value      = strDel & strVal

	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    DbSave  = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status

	DbDelete = False			                                                 '☜: Processing is NG

    If LayerShowHide(1) = false Then
        Exit Function
    End If

	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003                                '☜: Delete
	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
	DbDelete = True                                                              '⊙: Processing is NG

End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
    Dim strVal
	lgIntFlgMode      =  parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode

     ggoSpread.Source       = Frm1.vspdData2
    Frm1.vspdData2.MaxRows = 0

    Call MakeKeyStream("X")
    If LayerShowHide(1) = false Then
        Exit Function
    End If

    strVal = BIZ_PGM_ID1 & "?txtMode="            & parent.UID_M0001                    '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & 1                             '☜: Max fetched data
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    Call  ggoOper.LockField(Document, "Q")
	frm1.vspdData.focus
End Function
'========================================================================================================
' Function Name : DbQueryOk1
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk1()

	lgIntFlgMode      =  parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode

    Frm1.vspdData2.Col = 0
    Frm1.vspdData2.Text = "합계"
    Frm1.txtIntchng_yymm_dt.focus
	Call SetToolbar("1100111100111111")												'⊙: Set ToolBar

    Call InitData()
	Frm1.vspdData.focus
    Call  ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
	Call InitVariables
     ggoSpread.Source = Frm1.vspdData
    Frm1.vspdData.MaxRows = 0
    Call MainQuery()
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
	Call MainNew()
End Function

'========================================================================================================
' Name : FncOpenPopup
' Desc : developer describe this line
'========================================================================================================
Function FncOpenPopup(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then
	   Exit Function
	End If

	IsOpenPop = True
	Select Case iWhere
	    Case "1"

	        arrParam(0) = "대부코드조회 팝업"			' 팝업 명칭 
	        arrParam(1) = "B_MINOR"				 		' TABLE 명칭 
	        arrParam(2) = frm1.txtBorw_cd.value         ' Code Condition
	        arrParam(3) = ""'frm1.txtBorw_nm.value		' Name Cindition
	        arrParam(4) = " MAJOR_CD=" & FilterVar("h0043", "''", "S") & ""	    	' Where Condition
	        arrParam(5) = "대부코드"			    ' TextBox 명칭 

            arrField(0) = "MINOR_CD"					' Field명(0)
            arrField(1) = "MINOR_NM"				    ' Field명(1)

            arrHeader(0) = "대부코드"				' Header명(0)
            arrHeader(1) = "대부코드명"		    	    ' Header명(1)
	End Select

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBorw_cd.focus
		Exit Function
	Else
		Call SubSetOpenPop(arrRet,iWhere)
	End If

End Function

'======================================================================================================
'	Name : SubSetOpenPop()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetOpenPop(Byval arrRet, Byval iWhere)
	With Frm1
		Select Case iWhere
		    Case "1"
		        .txtBorw_cd.value = arrRet(0)
		        .txtBorw_nm.value = arrRet(1)
		        .txtBorw_cd.focus
        End Select
	End With
End Sub
'======================================================================================================
'	Name : OpenCode()
'	Description : Code PopUp at vspdData
'=======================================================================================================
Function OpenCode(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	    Case C_HDD020T_BORW_NM_POP
	        arrParam(0) = "대부코드조회 팝업"			        ' 팝업 명칭 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = ""                          			' Code Condition
	    	arrParam(3) = strCode								' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("h0043", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "대부코드" 			            ' TextBox 명칭 

	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)

	    	arrHeader(0) = "대부코드"	   		    	    ' Header명(0)
	    	arrHeader(1) = "대부코드명"	    		        ' Header명(1)
	End Select

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.vspdData.Col = C_HDD020T_BORW_CD
		frm1.vspdData.action =0	
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
       	 ggoSpread.Source = frm1.vspdData
         ggoSpread.UpdateRow Row
	End If

End Function
'======================================================================================================
'	Name : SetCode()
'	Description : Code PopUp에서 Return되는 값 setting
'=======================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)

	With frm1

		Select Case iWhere
		    Case C_HDD020T_BORW_NM_POP
		    	.vspdData.Col = C_HDD020T_BORW_NM
		    	.vspdData.text = arrRet(1)
		        .vspdData.Col = C_HDD020T_BORW_CD
		    	.vspdData.text = arrRet(0)
				.vspdData.action =0
        End Select

	End With
End Function
'======================================================================================================
'	Name : OpenEmp()
'	Description : Employee PopUp
'======================================================================================================
Function OpenEmp(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = frm1.txtEmp_no.value			' Code Condition
		If frm1.txtEmp_no.value="" Then
    	    arrParam(1) = ""'frm1.txtName.value		' Name Cindition
    	Else
    	    arrParam(1) = ""                		' Name Cindition
        End If
	Else 'spread
	    frm1.vspdData.Col = C_EMP_NO
		arrParam(0) = frm1.vspdData.Text			' Code Condition
	    frm1.vspdData.Col = C_NAME
	    arrParam(1) = ""'frm1.vspdData.Text			' Name Cindition
	End If
	arrParam(2) = lgUsrIntCd             			' Internal_cd

	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		If iWhere = 0 Then
			frm1.txtEmp_no.focus
		Else
			frm1.vspdData.Col = C_EMP_NO
			frm1.vspdData.action =0
		End If
	
		Exit Function
	Else
		Call SetEmp(arrRet, iWhere)
	End If

End Function

'======================================================================================================
'	Name : SetEmp()
'	Description : Employee Popup에서 Return되는 값 setting
'======================================================================================================
Function SetEmp(Byval arrRet, Byval iWhere)

	With frm1
		If iWhere = 0 Then
			.txtEmp_no.value = arrRet(0)
			.txtName.value = arrRet(1)
			.txtEmp_no.focus
		Else
			.vspdData.Col = C_NAME
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_EMP_NO
			.vspdData.Text = arrRet(0)
			.vspdData.action =0
		End If
	End With
End Function

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col-1
	Select Case Col
	    Case C_EMP_NO_POP
                    Call OpenEmp(1)
	    Case C_HDD020T_BORW_NM_POP
                    Call OpenCode(frm1.vspdData.Text, C_HDD020T_BORW_NM_POP, Row)
    End Select
End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Function vspdData_Change(ByVal Col , ByVal Row)

    Dim iDx
    Dim IntRetCD
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd
    Dim strVal

   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    With frm1
        Select Case Col
             Case  C_EMP_NO
                If Trim(.vspdData.Text) = "" Then
               	    .vspdData.Text = ""
    	            .vspdData.Col = C_NAME
                    .vspdData.Text = ""
               	 Else
	                    IntRetCd =  FuncGetEmpInf2(Trim(.vspdData.Text),lgUsrIntCd,strName,strDept_nm,_
	                                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	                    If  IntRetCd < 0 then
	                        If  IntRetCd = -1 then
    	                		Call  DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
                            Else
                                Call  DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
                            End if
    	                        .vspdData.Col = C_NAME
                                .vspdData.Text = ""
    	                        .vspdData.Col = C_EMP_NO
                                .vspdData.Action = 0 ' go to
                                Set gActiveElement = document.ActiveElement
                                vspdData_Change =true
                                Exit Function
                        Else
    	                    .vspdData.Col = C_NAME
                            .vspdData.Text=strName
                        End If
                  End If
	         Case C_HDD020T_BORW_CD
                    IntRetCD =  CommonQueryRs(" minor_nm "," b_minor "," major_cd=" & FilterVar("H0043", "''", "S") & " And minor_cd =  " & FilterVar(frm1.vspdData.Text, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
                    If IntRetCD=False And Trim(frm1.vspdData.Text)<>""  Then
						Call DisplayMsgBox("970000", "X","대부코드","x")
    	                frm1.vspdData.Col = C_HDD020T_BORW_NM
	                    frm1.vspdData.Text=""
                        vspdData_Change =true
                        
                    ElseIf  CountStrings(lgF0, Chr(11) ) > 1 Then    ' 같은명일 경우 pop up
                        Call OpenCode(frm1.vspdData.Text, C_HDD020T_BORW_NM_POP, Row)
                    Else
    	                frm1.vspdData.Col = C_HDD020T_BORW_NM
                        frm1.vspdData.Text=Trim(Replace(lgF0,Chr(11),""))
                    End If
        End Select
    End With

   	If Frm1.vspdData.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData.text) <  UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If

	 ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row

End Function


'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	With frm1.vspdData

		.Row = Row
        Select Case Col
            Case C_HDD020T_INTREST_TYPE_NM
                .Col = Col
                intIndex = .Value
				.Col = C_HDD020T_INTREST_TYPE
				.Value = intIndex

		End Select
	End With

   	 ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row

End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    
    Call SetPopupMenuItemInf("1101101111")

    gMouseClickStatus = "SPC" 

    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If
End Sub
'-----------------------------------------
Sub vspdData_MouseDown(Button , Shift , x , y)

       If Button = 2 And  gMouseClickStatus = "SPC" Then
           gMouseClickStatus = "SPCR"
        End If
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)

    gMouseClickStatus = "SP1C" 

    Set gActiveSpdSheet = frm1.vspdData2
End Sub
'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
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

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

    frm1.vspdData.Col = pvCol1
    frm1.vspdData2.ColWidth(pvCol1) = frm1.vspdData.ColWidth(pvCol1)

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
    frm1.vspdData.ColWidth(pvCol1) = frm1.vspdData2.ColWidth(pvCol1)

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
        frm1.vspdData2.LeftCol=NewLeft   	
		Exit Sub
	End If
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If  
End Sub
'========================================================================================================
'   Event Name : vspdData2_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
        frm1.vspdData.LeftCol=NewLeft   	
		Exit Sub
	End If
	If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then
		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If  
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

	frm1.txtName.value = ""

    If  frm1.txtEmp_no.value = "" Then
        frm1.txtName.value = ""
    Else
	    IntRetCd =  FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    if  IntRetCd < 0 then
	        if  IntRetCd = -1 then
    			Call  DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call  DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            end if
            Call  ggoOper.ClearField(Document, "2")
            call InitVariables()
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true
        Else
            frm1.txtName.value = strName
        End if
    End if

End Function
'=======================================================================================================
'   Event Name : txtYear_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtIntchng_yymm_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtIntchng_yymm_dt.Action = 7
        frm1.txtIntchng_yymm_dt.focus
    End If
End Sub
'=======================================================================================================
'   Event Name : txtIntchng_yymm_dt_Keypress(Key)
'   Event Desc : 3rd party control에서 Enter 키를 누르면 조회 실행 
'=======================================================================================================
Sub txtIntchng_yymm_dt_Keypress(Key)
    If Key = 13 Then
        Call MainQuery()
    End If
End Sub
'======================================================================================================
'   Event Name : txtBorw_cd_OnChange
'   Event Desc : 대부코드가 변경될 경우 
'=======================================================================================================
Function txtBorw_cd_OnChange()
    Dim IntRetCd

    If Trim(frm1.txtBorw_cd.value) = "" Then
        frm1.txtBorw_nm.Value=""
    Else
        IntRetCD =  CommonQueryRs(" minor_nm "," b_minor "," major_cd=" & FilterVar("H0043", "''", "S") & " And minor_cd =  " & FilterVar(frm1.txtBorw_cd.value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCD=False And Trim(frm1.txtBorw_cd.Value)<>""  Then
            Call  DisplayMsgBox("800054","X","X","X")                         '☜ : 등록되지 않은 코드입니다.
            frm1.txtBorw_nm.Value=""
            txtBorw_cd_OnChange = true
        Else
            frm1.txtBorw_nm.Value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%> ></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>월대부상환현황</font></td>
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
						<FIELDSET CLASS="CLSFLD">
                           <TABLE <%=LR_SPACE_TYPE_40%>>
						    <TR>
							    <TD CLASS=TD5 NOWRAP>상환년월</TD>
			                    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h5205ma1_fpDateTime12_txtIntchng_yymm_dt.js'></script></TD>
			    	    		<TD CLASS=TD5 NOWRAP>사원</TD>
			    	    		<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no"  SIZE=13 MAXLENGTH=13 ALT="사번" TYPE="Text"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmp(0)">
			    	        	                     <INPUT NAME="txtName"  SIZE=20  MAXLENGTH=30 ALT="성명" TYPE="Text"  tag="14XXXU"></TD>
							</TR>
						    <TR>
					            <TD CLASS=TD5 NOWRAP>이자구분</TD>
					            <TD CLASS=TD6 NOWRAP><SELECT Name="txtIntrest_type" ALT="이자구분" STYLE="WIDTH: 100px" TAG="11"><OPTION VALUE=""></OPTION></SELECT></TD>
			    	    		<TD CLASS=TD5 NOWRAP>대부코드</TD>
			    	    		<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBorw_cd"  SIZE=10 MAXLENGTH=10 ALT="대부코드"   TYPE="Text"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:FncOpenPopup(1)">
			    	        	                     <INPUT NAME="txtBorw_nm"  SIZE=20 MAXLENGTH=50 ALT="대부코드명" TYPE="Text"  tag="14XXXU"></TD>
							</TR>
						   </TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%>></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT=100%>
									<script language =javascript src='./js/h5205ma1_vaSpread_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=43 VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT=100%>
									<script language =javascript src='./js/h5205ma1_vaSpread1_vspdData2.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=100% FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
