<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          	: Human Resources
*  2. Function Name        	: Multi Sample
*  3. Program ID           	: H6011ma1
*  4. Program Name         	: H6011ma1
*  5. Program Desc         	: 월별임금대비표조회 
*  6. Comproxy List        	:
*  7. Modified date(First) 	: 2001/04/18
*  8. Modified date(Last)  	: 2003/06/13
*  9. Modifier (First)     	: TGS 최용철 
* 10. Modifier (Last)      	: Lee SiNa
* 11. Comment              	:
=======================================================================================================-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID      = "h6011mb1.asp"						           '☆: Biz Logic ASP Name
Const CookieSplit = 1233
Const C_SHEETMAXROWS    = 21	                                      '☜: Visble row

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop
Dim lgOldRow

Dim C_ALLOW_CD
Dim C_ALLOW_NM
Dim C_BAS_GG_AMT
Dim C_BAS_GG_TOTAL_AMT
Dim C_MM01_AMT
Dim C_MM02_AMT
Dim C_MM03_AMT
Dim C_MM04_AMT
Dim C_MM05_AMT
Dim C_MM06_AMT
Dim C_MM07_AMT
Dim C_MM08_AMT
Dim C_MM09_AMT
Dim C_MM10_AMT
Dim C_MM11_AMT
Dim C_MM12_AMT

Dim C_ALLOW_CD2
Dim C_ALLOW_NM2
Dim C_BAS_GG_AMT2 
Dim C_BAS_GG_TOTAL_AMT2 
Dim C_MM01_AMT2
Dim C_MM02_AMT2
Dim C_MM03_AMT2
Dim C_MM04_AMT2
Dim C_MM05_AMT2
Dim C_MM06_AMT2
Dim C_MM07_AMT2
Dim C_MM08_AMT2
Dim C_MM09_AMT2
Dim C_MM10_AMT2
Dim C_MM11_AMT2
Dim C_MM12_AMT2

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
 
    C_ALLOW_CD = 1															<%'Spread Sheet의 Column별 상수 %>
    C_ALLOW_NM = 2
    C_BAS_GG_AMT = 3
    C_BAS_GG_TOTAL_AMT = 4
    C_MM01_AMT = 5
    C_MM02_AMT = 6
    C_MM03_AMT = 7
    C_MM04_AMT = 8
    C_MM05_AMT = 9
    C_MM06_AMT = 10
    C_MM07_AMT = 11
    C_MM08_AMT = 12
    C_MM09_AMT = 13
    C_MM10_AMT = 14
    C_MM11_AMT = 15
    C_MM12_AMT = 16
    
    C_ALLOW_CD2 = 1															<%'Spread Sheet의 Column별 상수 %>
    C_ALLOW_NM2 = 2
    C_BAS_GG_AMT2 = 3
    C_BAS_GG_TOTAL_AMT2 = 4
    C_MM01_AMT2 = 5
    C_MM02_AMT2 = 6
    C_MM03_AMT2 = 7
    C_MM04_AMT2 = 8
    C_MM05_AMT2 = 9
    C_MM06_AMT2 = 10
    C_MM07_AMT2 = 11
    C_MM08_AMT2 = 12
    C_MM09_AMT2 = 13
    C_MM10_AMT2 = 14
    C_MM11_AMT2 = 15
    C_MM12_AMT2 = 16

End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
	lgOldRow = 0
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay
	
	Call ExtractDateFrom("<%=GetSvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	frm1.txtBas_yy.focus 		
	frm1.txtBas_yy.Year = strYear 		 '년월일 default value setting
	
	frm1.txtDiff_yymm_dt.Year = strYear
	frm1.txtDiff_yymm_dt.Month = strMonth 
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H", "NOCOOKIE", "MA") %>
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
    Dim StrDt
    StrDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtDiff_yymm_dt.Year, Right("0" & frm1.txtDiff_yymm_dt.month , 2), "01")
    lgKeyStream  = Frm1.txtDiff_yymm_dt.Year & Right("0" & Frm1.txtDiff_yymm_dt.Month, 2) & Parent.gColSep
    lgKeyStream  = lgKeyStream & mid(Frm1.txtBas_yy.text,1,4) & Parent.gColSep   '기준년 
    lgKeyStream  = lgKeyStream & Trim(Frm1.txtfr_internal_cd.Value) & Parent.gColSep
    lgKeyStream  = lgKeyStream & Trim(Frm1.txtto_internal_cd.Value) & Parent.gColSep
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
    	
	With frm1.vspdData
        ggoSpread.Source = frm1.vspdData2
		ggoSpread.UpdateRow 1

        frm1.vspdData2.Col = 0
        frm1.vspdData2.Text = "합계"

        frm1.vspdData2.Col = C_BAS_GG_AMT2
        frm1.vspdData2.text = FncSumSheet(frm1.vspdData,C_BAS_GG_AMT, 1, .MaxRows, FALSE ,-1, -1, "V")
        frm1.vspdData2.Col = C_BAS_GG_TOTAL_AMT2
        frm1.vspdData2.text = FncSumSheet(frm1.vspdData,C_BAS_GG_TOTAL_AMT, 1, .MaxRows,FALSE , -1, -1, "V")
        frm1.vspdData2.Col = C_MM01_AMT2
        frm1.vspdData2.text = FncSumSheet(frm1.vspdData,C_MM01_AMT, 1, .MaxRows,FALSE , -1, -1, "V")
        frm1.vspdData2.Col = C_MM02_AMT2
        frm1.vspdData2.text = FncSumSheet(frm1.vspdData,C_MM02_AMT, 1, .MaxRows,FALSE , -1, -1, "V")
        frm1.vspdData2.Col = C_MM03_AMT2
        frm1.vspdData2.text = FncSumSheet(frm1.vspdData,C_MM03_AMT, 1, .MaxRows,FALSE , -1, -1, "V")
        frm1.vspdData2.Col = C_MM04_AMT2
        frm1.vspdData2.text = FncSumSheet(frm1.vspdData,C_MM04_AMT, 1, .MaxRows,FALSE , -1, -1, "V")
        frm1.vspdData2.Col = C_MM05_AMT2
        frm1.vspdData2.text = FncSumSheet(frm1.vspdData,C_MM05_AMT, 1, .MaxRows,FALSE , -1, -1, "V")
        frm1.vspdData2.Col = C_MM06_AMT2
        frm1.vspdData2.text = FncSumSheet(frm1.vspdData,C_MM06_AMT, 1, .MaxRows,FALSE , -1, -1, "V")
        frm1.vspdData2.Col = C_MM07_AMT2
        frm1.vspdData2.text = FncSumSheet(frm1.vspdData,C_MM07_AMT, 1, .MaxRows,FALSE , -1, -1, "V")
        frm1.vspdData2.Col = C_MM08_AMT2
        frm1.vspdData2.text = FncSumSheet(frm1.vspdData,C_MM08_AMT, 1, .MaxRows,FALSE , -1, -1, "V")
        frm1.vspdData2.Col = C_MM09_AMT2
        frm1.vspdData2.text = FncSumSheet(frm1.vspdData,C_MM09_AMT, 1, .MaxRows,FALSE , -1, -1, "V")
        frm1.vspdData2.Col = C_MM10_AMT2
        frm1.vspdData2.text = FncSumSheet(frm1.vspdData,C_MM10_AMT, 1, .MaxRows,FALSE , -1, -1, "V")
        frm1.vspdData2.Col = C_MM11_AMT2
        frm1.vspdData2.text = FncSumSheet(frm1.vspdData,C_MM11_AMT, 1, .MaxRows,FALSE , -1, -1, "V")
        frm1.vspdData2.Col = C_MM12_AMT2
        frm1.vspdData2.text = FncSumSheet(frm1.vspdData,C_MM12_AMT, 1, .MaxRows,FALSE , -1, -1, "V")
        
		call SetSpreadLock2

    End With
    
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
	
           .MaxCols   = C_MM12_AMT + 1                                                      ' ☜:☜: Add 1 to Maxcols
	                                               ' ☜:☜: Add 1 to Maxcols
	       .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
           .ColHidden = True                                                            ' ☜:☜:

           .MaxRows = 0
            ggoSpread.ClearSpreadData

            Call GetSpreadColumnPos("A") 'sbk
           
            Call AppendNumberPlace("6","15","0")
            
            ggoSpread.SSSetEdit  C_ALLOW_CD     , "수당코드", 8 ,,,3,2		'Lock/ Edit
            ggoSpread.SSSetEdit  C_ALLOW_NM     , "수당명", 14 ,,,50,2		'Lock/ Edit
            ggoSpread.SSSetFloat C_BAS_GG_AMT   , "비교연월수당금액" ,13,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_BAS_GG_TOTAL_AMT , "누계" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
    
            ggoSpread.SSSetFloat C_MM01_AMT     , "01월" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_MM02_AMT     , "02월" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_MM03_AMT     , "03월" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_MM04_AMT     , "04월" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_MM05_AMT     , "05월" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_MM06_AMT     , "06월" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_MM07_AMT     , "07월" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_MM08_AMT     , "08월" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_MM09_AMT     , "09월" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_MM10_AMT     , "10월" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_MM11_AMT     , "11월" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_MM12_AMT     , "12월" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
         
	       .ReDraw = true
    
        End With
    End If

    If pvSpdNo = "" OR pvSpdNo = "B" Then
    
	    With frm1.vspdData2

            ggoSpread.Source = frm1.vspdData2
            ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    'sbk

	       .ReDraw = false
	
           .MaxCols   = C_MM12_AMT2 + 1                                                      ' ☜:☜: Add 1 to Maxcols
	                                               ' ☜:☜: Add 1 to Maxcols
	       .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
           .ColHidden = True                                                            ' ☜:☜:
 
           .MaxRows = 0
            ggoSpread.ClearSpreadData

           .DisplayColHeaders = False

            Call GetSpreadColumnPos("B") 'sbk

            Call AppendNumberPlace("6","15","0")

            ggoSpread.SSSetEdit C_ALLOW_CD2      , "", 8,,,15,2		'Lock/ Edit
            ggoSpread.SSSetEdit C_ALLOW_NM2      , "", 14,,,12,2		'Lock/ Edit
            ggoSpread.SSSetFloat C_BAS_GG_AMT2   , "비교연월수당금액" ,13,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_BAS_GG_TOTAL_AMT2 , "누계" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_MM01_AMT2     , "01월" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_MM02_AMT2     , "02월" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_MM03_AMT2     , "03월" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_MM04_AMT2     , "04월" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_MM05_AMT2     , "05월" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_MM06_AMT2     , "06월" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_MM07_AMT2     , "07월" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_MM08_AMT2     , "08월" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_MM09_AMT2     , "09월" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_MM10_AMT2     , "10월" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_MM11_AMT2     , "11월" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_MM12_AMT2     , "12월" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
                     
	       .ReDraw = true
    
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
	ggoSpread.SpreadLockWithOddEvenRowColor()

	call SetSpreadLock2
    
End Sub

Sub SetSpreadLock2()
	ggoSpread.Source = frm1.vspdData2 
    With frm1
		.vspdData2.ReDraw = False
		ggoSpread.SpreadLock    C_ALLOW_CD2, -1, C_ALLOW_CD2, -1
		ggoSpread.SpreadLock    C_ALLOW_NM2, -1, C_ALLOW_NM2, -1 
		ggoSpread.SpreadLock    C_BAS_GG_AMT2, -1, C_BAS_GG_AMT2, -1 
		ggoSpread.SpreadLock    C_BAS_GG_TOTAL_AMT2, -1, C_BAS_GG_TOTAL_AMT2, -1 
		ggoSpread.SpreadLock    C_MM01_AMT2, -1, C_MM01_AMT2, -1 
		ggoSpread.SpreadLock    C_MM02_AMT2, -1, C_MM02_AMT2, -1 
		ggoSpread.SpreadLock    C_MM03_AMT2, -1, C_MM03_AMT2, -1 
		ggoSpread.SpreadLock    C_MM04_AMT2, -1, C_MM04_AMT2, -1 
		ggoSpread.SpreadLock    C_MM05_AMT2, -1, C_MM05_AMT2, -1 
		ggoSpread.SpreadLock    C_MM06_AMT2, -1, C_MM06_AMT2, -1 
		ggoSpread.SpreadLock    C_MM07_AMT2, -1, C_MM07_AMT2, -1 
		ggoSpread.SpreadLock    C_MM08_AMT2, -1, C_MM08_AMT2, -1 
		ggoSpread.SpreadLock    C_MM09_AMT2, -1, C_MM09_AMT2, -1 
		ggoSpread.SpreadLock    C_MM10_AMT2, -1, C_MM10_AMT2, -1
		ggoSpread.SpreadLock    C_MM11_AMT2, -1, C_MM11_AMT2, -1 
		ggoSpread.SpreadLock    C_MM12_AMT2, -1, C_MM12_AMT2, -1 				 
		ggoSpread.SSSetProtected   .vspdData2.MaxCols   , -1, -1
		.vspdData2.ReDraw = True

    End With
End Sub
'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
End Sub

'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
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
            
            C_ALLOW_CD = iCurColumnPos(1)
            C_ALLOW_NM = iCurColumnPos(2)
            C_BAS_GG_AMT = iCurColumnPos(3)
            C_BAS_GG_TOTAL_AMT = iCurColumnPos(4)
            C_MM01_AMT = iCurColumnPos(5)
            C_MM02_AMT = iCurColumnPos(6)
            C_MM03_AMT = iCurColumnPos(7)
            C_MM04_AMT = iCurColumnPos(8)
            C_MM05_AMT = iCurColumnPos(9)
            C_MM06_AMT = iCurColumnPos(10)
            C_MM07_AMT = iCurColumnPos(11)
            C_MM08_AMT = iCurColumnPos(12)
            C_MM09_AMT = iCurColumnPos(13)
            C_MM10_AMT = iCurColumnPos(14)
            C_MM11_AMT = iCurColumnPos(15)
            C_MM12_AMT = iCurColumnPos(16)
            
        Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

            C_ALLOW_CD2 = iCurColumnPos(1)
            C_ALLOW_NM2 = iCurColumnPos(2)                
            C_BAS_GG_AMT2 = iCurColumnPos(3)              
            C_BAS_GG_TOTAL_AMT2 = iCurColumnPos(4)        
            C_MM01_AMT2 = iCurColumnPos(5)                
            C_MM02_AMT2 = iCurColumnPos(6)                
            C_MM03_AMT2 = iCurColumnPos(7)                
            C_MM04_AMT2 = iCurColumnPos(8)                
            C_MM05_AMT2 = iCurColumnPos(9)                
            C_MM06_AMT2 = iCurColumnPos(10)               
            C_MM07_AMT2 = iCurColumnPos(11)               
            C_MM08_AMT2 = iCurColumnPos(12)               
            C_MM09_AMT2 = iCurColumnPos(13)               
            C_MM10_AMT2 = iCurColumnPos(14)               
            C_MM11_AMT2 = iCurColumnPos(15)               
            C_MM12_AMT2 = iCurColumnPos(16)               
            
    End Select    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
            
    Call InitSpreadSheet("")                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    
    Call ggoOper.FormatDate(frm1.txtBas_yy, Parent.gDateFormat, 3)                         '싱글에서 년월말 입력하고 싶은경우 다음 함수를 콜한다.
    Call ggoOper.FormatDate(frm1.txtDiff_yymm_dt, Parent.gDateFormat, 2) 
    
    Call FuncGetAuth("H6011MA1", Parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
    
    Call SetDefaultVal
	Call SetToolbar("1100000000001111")												'⊙: Set ToolBar

    frm1.txtBas_yy.focus()

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
    Dim strwhere 
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    If  txtFr_Dept_cd_Onchange()  then
        Exit Function
    End If
    If  txtTo_Dept_cd_Onchange()  then
        Exit Function
    End If
 
    Dim Fr_dept_cd , To_dept_cd, rFrDept ,rToDept,StrDt
    
    StrDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtDiff_yymm_dt.Year, Right("0" & frm1.txtDiff_yymm_dt.month , 2), "01")

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
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.activeElement
            Exit Function
        End IF 
        
    End If   

    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("X")
    
    Call DisableToolBar(Parent.TBC_QUERY)
    
	ggoSpread.Source = frm1.vspdData2
    ggoSpread.InsertRow
    call SetSpreadLock2
    
	IF DBQUERY =  False Then
		Call RestoreToolBar()
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
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If
    
    Call DisableToolBar(Parent.TBC_SAVE)
	IF DBSAVE =  False Then
		Call RestoreToolBar()
		Exit Function
	End If
    
    FncSave = True                                                              '☜: Processing is OK
    
End Function


'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
    
    ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
            ggoSpread.CopyRow
			SetSpreadColor .ActiveRow
    
            .ReDraw = True
		    .Focus
		 End If
	End With
	
    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo  
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow() 
	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow
        SetSpreadColor .vspdData.ActiveRow
       .vspdData.ReDraw = True
    End With
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
    	lDelRows = ggoSpread.DeleteRow
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
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
	Call Parent.FncExport(Parent.C_MULTI)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                    '☜:화면 유형, Tab 유무 
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
	dim temp
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet("A")      
    ggoSpread.Source = frm1.vspdData
	Call ggoSpread.ReOrderingSpreadData()

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet("B")      
    ggoSpread.Source = frm1.vspdData2
	Call ggoSpread.ReOrderingSpreadData()

	temp = GetSpreadText(frm1.vspdData,1,1,"X","X")

	if temp <>"" then
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.InsertRow
		Call InitData()
	end if  
End Sub

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
    Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
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

    DbQuery = False
    
    Err.Clear                                                                        '☜: Clear err status

	If LayerShowHide(1) = False then
    		Exit Function 
    End if
	
	Dim strVal
    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
    End With
		
    If lgIntFlgMode = Parent.OPMD_UMODE Then
    Else
    End If
	
	Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic
    
    DbQuery = True
    
End Function
'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
	Dim strVal, strDel
	
    DbSave = False                                                          
    
    If LayerShowHide(1) = False then
    	Exit Function 
    End if

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
 
           Case ggoSpread.InsertFlag                                      '☜: Update
                                            strVal = strVal & "C" & Parent.gColSep
                                            strVal = strVal & lRow & Parent.gColSep
                                         
             .vspdData.Col = C_NAME	      : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
             .vspdData.Col = C_EMP_NO	  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
             .vspdData.Col = C_DEPT_CD	  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
             .vspdData.Col = C_ALLOW_CD   : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
             .vspdData.Col = C_ALLOW_AMT  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
             .vspdData.Col = C_PAY_CD     : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
             .vspdData.Col = C_PROV_TYPE  : strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
             
             lGrpCnt = lGrpCnt + 1
      
           Case ggoSpread.UpdateFlag                                      '☜: Update
                                           strVal = strVal & "U" & Parent.gColSep
                                           strVal = strVal & lRow & Parent.gColSep
             
             .vspdData.Col = C_NAME	      : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
             .vspdData.Col = C_EMP_NO	  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
             .vspdData.Col = C_DEPT_CD	  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
             .vspdData.Col = C_ALLOW_CD   : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
             .vspdData.Col = C_ALLOW_AMT  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
             .vspdData.Col = C_PAY_CD     : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
             .vspdData.Col = C_PROV_TYPE   : strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep   
             
             lGrpCnt = lGrpCnt + 1
             
           Case ggoSpread.DeleteFlag                                      '☜: Delete

                                           strDel = strDel & "D" & Parent.gColSep
                                           strDel = strDel & lRow & Parent.gColSep
             .vspdData.Col = C_NAME	     : strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
             .vspdData.Col = C_EMP_NO	 : strDel = strDel & Trim(.vspdData.Text) & Parent.gRowSep								
             lGrpCnt = lGrpCnt + 1
           End Select
      Next
	
       .txtMode.value        = Parent.UID_M0002
       .txtUpdtUserId.value  = Parent.gUsrID
       .txtInsrtUserId.value = Parent.gUsrID
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
    Dim IntRetCd
    
    FncDelete = False                                                      '⊙: Processing is NG
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                '☆:
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If
    
    Call DisableToolBar(Parent.TBC_DELETE)
	IF DBDELETE =  False Then
		Call RestoreToolBar()
		Exit Function
	End If
    
    FncDelete = True                                                        '⊙: Processing is OK


End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
	
    lgIntFlgMode = Parent.OPMD_UMODE    
    Call ggoOper.LockField(Document, "Q")										'⊙: Lock field
    Call InitData()
	Call SetToolbar("1100000000011111")									
	Frm1.vspdData.focus

    Set gActiveElement = document.ActiveElement   
	
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

'========================================================================================================
'	Name : OpenMajor()
'	Description : Major PopUp
'========================================================================================================
Function OpenMajor()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "Major코드 팝업"			' 팝업 명칭 
	arrParam(1) = "B_MAJOR"				 		' TABLE 명칭 
	arrParam(2) = frm1.txtMajorCd.value			' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "Major코드"			
	
    arrField(0) = "major_cd"					' Field명(0)
    arrField(1) = "major_nm"				    ' Field명(1)
    
    arrHeader(0) = "Major코드"		        ' Header명(0)
    arrHeader(1) = "Major코드명"			' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtMajorCd.focus	
		Exit Function
	Else
		Call SetMajor(arrRet)
	End If	

End Function

'========================================================================================================
'	Name : SetMajor()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function SetMajor(Byval arrRet)
	With frm1
		.txtMajorCd.value = arrRet(0)
		.txtMajorNm.value = arrRet(1)		
		.txtMajorCd.focus
	End With
End Function

'========================================================================================================
' Name : OpenDept
' Desc : 부서 POPUP
'========================================================================================================
Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(2)
	Dim strBasDt

	strBasDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtDiff_yymm_dt.Year, Right("0" & frm1.txtDiff_yymm_dt.month , 2), "01")

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then
		arrParam(0) = frm1.txtFr_dept_cd.value			            '  Code Condition
	ElseIf iWhere = 1 Then
		arrParam(0) = frm1.txtTo_dept_cd.value			            ' Code Condition
	End If

    	arrParam(1) = strBasDt
	arrParam(2) = lgUsrIntcd
	
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
               .txtFr_internal_cd.value = arrRet(2)
               .txtFr_dept_cd.focus
             Case "1"  
               .txtTo_dept_cd.value = arrRet(0)
               .txtTo_dept_nm.value = arrRet(1) 
               .txtTo_internal_cd.value = arrRet(2) 
               .txtTo_dept_cd.focus
        End Select
	End With
End Function       		


'========================================================================================================
'   Event Name : txtFr_dept_cd_change
'   Event Desc :
'========================================================================================================
Function txtFr_dept_cd_Onchange()
    
	Dim strBasDt
    Dim IntRetCd,Dept_Nm,Internal_cd
    Dim rDay,rDate
    
	strBasDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtDiff_yymm_dt.Year, Right("0" & frm1.txtDiff_yymm_dt.month , 2), "01")
    
    If frm1.txtFr_dept_cd.value = "" Then
		frm1.txtFr_dept_nm.value = ""
		frm1.txtFr_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtFr_dept_cd.value , strBasDt , lgUsrIntCd,Dept_Nm , Internal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800098","X","X","X")	'부서정보코드에 등록되지 않은 코드입니다.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
		    frm1.txtFr_dept_nm.value = ""
		    frm1.txtFr_internal_cd.value = ""
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.ActiveElement 
            txtFr_dept_cd_Onchange = true
            Exit Function      
        Else
			frm1.txtFr_dept_nm.value = Dept_Nm
		    frm1.txtFr_internal_cd.value = Internal_cd
        End if 
    End if  
    
End Function

'========================================================================================================
'   Event Name : txtTo_dept_cd_Onchange
'   Event Desc :
'========================================================================================================
Function txtTo_dept_cd_Onchange()
   
	Dim strBasDt
    Dim IntRetCd,Dept_Nm,Internal_cd
    
	strBasDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtDiff_yymm_dt.Year, Right("0" & frm1.txtDiff_yymm_dt.month , 2), "01")
 
    If frm1.txtTo_dept_cd.value = "" Then
		frm1.txtTo_dept_nm.value = ""
		frm1.txtTo_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtTo_dept_cd.value , strBasDt , lgUsrIntCd,Dept_Nm , Internal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800098","X","X","X")	'부서정보코드에 등록되지 않은 코드입니다.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
		    frm1.txtTo_dept_nm.value = ""
		    frm1.txtTo_internal_cd.value = ""
            frm1.txtTo_dept_cd.focus
            Set gActiveElement = document.ActiveElement 
            txtTo_dept_cd_Onchange = true
            Exit Function      
        Else
			frm1.txtTo_dept_nm.value = Dept_Nm
		    frm1.txtTo_internal_cd.value = Internal_cd
        End if 
    End if  
    
End Function   
'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col
             
   	If Frm1.vspdData.CellType = Parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    
    Call SetPopupMenuItemInf("0000101111")

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

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc :
'==========================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0000000000")

    gMouseClickStatus = "SP1C" 

    Set gActiveSpdSheet = frm1.vspdData2
   
End Sub
'-----------------------------------------

Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
        gMouseClickStatus = "SPCR"
    End If
End Sub    

Sub vspdData2_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SP1C" Then
       gMouseClickStatus = "SP1CR"
    End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
        Exit Sub
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
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
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col , ByVal Row, ByVal newCol , ByVal newRow ,Cancel )
    frm1.vspdData2.Col = newCol
    frm1.vspdData2.Action = 0
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

Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        frm1.vspdData.LeftCol=NewLeft
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
End Sub

'=======================================================================================================
'   Event Name : txtBas_yy_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtBas_yy_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtBas_yy.Action = 7
        frm1.txtBas_yy.focus
    End If
End Sub
'=======================================================================================================
'   Event Name : txtDiff_yymm_dt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtDiff_yymm_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtDiff_yymm_dt.Action = 7
        frm1.txtDiff_yymm_dt.focus
    End If
End Sub

Sub txtBas_yy_Keypress(KeyAscii)
    If KeyAscii = 13 Then
        Call MainQuery()
    End If
End Sub

Sub txtDiff_yymm_dt_Keypress(KeyAscii)
    If KeyAscii = 13 Then
        Call MainQuery()
    End If
End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
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
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>월별임금대비표조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=LIGHT STYLE="TEXT-ALIGN:right">단위:천원&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
		
    <TR HEIGHT=*>
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR><TD <%=HEIGHT_TYPE_02%>></TD></TR>
				<TR>
					<TD HEIGHT=20>
					  <FIELDSET CLASS="CLSFLD">
					   <TABLE <%=LR_SPACE_TYPE_40%>>
						    <TR>
								<TD CLASS=TD5 NOWRAP>기준년</TD>
								<TD CLASS=TD6 NOWRAP><OBJECT classid=<%=gCLSIDFPDT%> id=txtBas_yy NAME="txtBas_yy" CLASS=FPDTYYYY  title=FPDATETIME ALT="기준년" tag="12X1" VIEWASTEXT> </OBJECT></TD>
	                        	<TD CLASS=TD5 NOWRAP>부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFr_dept_cd" ALT="부서코드" TYPE="Text" SiZE=10 MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(0)">
			                                        <INPUT NAME="txtFr_dept_nm" ALT="부서코드명" TYPE="Text" SiZE=20 MAXLENGTH=40 tag="14XXXU">&nbsp;~
		                                           <INPUT NAME="txtFr_Internal_cd" ALT="내부부서코드" TYPE="hidden" SiZE=7 MAXLENGTH=7 tag="14XXXU">  
                            </TR>
	                        <TR>
								<TD CLASS=TD5 NOWRAP>비교년월</TD>
								<TD CLASS=TD6 NOWRAP><OBJECT classid=<%=gCLSIDFPDT%> id=txtDiff_yymm_dt NAME="txtDiff_yymm_dt" CLASS=FPDTYYYYMM title=FPDATETIME ALT="비교년월" tag="12X1" VIEWASTEXT> </OBJECT></TD>				                                           
		                        <TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtto_dept_cd" MAXLENGTH="10" SIZE=10 ALT ="Order ID" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(1)">
							                        <INPUT NAME="txtto_dept_nm" MAXLENGTH="40" SIZE=20 ALT ="Order ID" tag="14XXXU">
    			                                   <INPUT  NAME="txtTo_Internal_cd" ALT="내부부서코드" TYPE="hidden" SiZE=7 MAXLENGTH=7 tag="14XXXU"></TD>
                           </TR>
	                       
	                   </TABLE>
				     </FIELDSET>
				   </TD>
				</TR>
				<TR><TD <%=HEIGHT_TYPE_03%>></TD></TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT="100%">
									<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"  id=vaSpread>
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
									<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"  id=vaSpread1>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>

<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

