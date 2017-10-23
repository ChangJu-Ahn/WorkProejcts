<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        :
*  3. Program ID           : h5204ma1
*  4. Program Name         : 대부금 등록 
*  5. Program Desc         : 대부금 조회,등록,변경,삭제 
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables
'========================================================================================================
Const BIZ_PGM_ID      = "h5204mb1.asp"						           '☆: Biz Logic ASP Name
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

Dim C_HDD010T_BORW_CD
Dim C_HDD010T_BORW_NM
Dim C_HDD010T_BORW_NM_POP
Dim C_HDD010T_INTREST_TYPE
Dim C_HDD010T_INTREST_TYPE_NM
Dim C_HDD010T_BORW_DT
Dim C_HDD010T_EXPIR_DT
Dim C_HDD010T_RESRV_DUR
Dim C_HDD010T_BORW_TOT_AMT
Dim C_HDD010T_PAY_INTCHNG_AVG
Dim C_HDD010T_TOT_INTCHNG_CNT
Dim C_HDD010T_PAY_INTCHNG_AMT
Dim C_HDD010T_BONUS_INTCHNG_CNT
Dim C_HDD010T_BONUS_INTCHNG_AMT
Dim C_HDD010T_INTREST_RATE

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  

    C_HDD010T_BORW_CD = 1										        'Column Dimant for Spread Sheet
    C_HDD010T_BORW_NM = 2
    C_HDD010T_BORW_NM_POP = 3
    C_HDD010T_INTREST_TYPE = 4
    C_HDD010T_INTREST_TYPE_NM = 5
    C_HDD010T_BORW_DT = 6
    C_HDD010T_EXPIR_DT = 7
    C_HDD010T_RESRV_DUR = 8
    C_HDD010T_BORW_TOT_AMT = 9
    C_HDD010T_PAY_INTCHNG_AVG = 10  
    C_HDD010T_TOT_INTCHNG_CNT = 11
    C_HDD010T_PAY_INTCHNG_AMT = 12
    C_HDD010T_BONUS_INTCHNG_CNT = 13  
    C_HDD010T_BONUS_INTCHNG_AMT = 14
    C_HDD010T_INTREST_RATE = 15

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
    lgKeyStream       = frm1.txtEmp_no.value & parent.gColSep       'You Must append one character( parent.gColSep)
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
    Dim iDx

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("h0044", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) '이자구분 

    iCodeArr = lgF0
    iNameArr = lgF1

     ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_HDD010T_INTREST_TYPE
     ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_HDD010T_INTREST_TYPE_NM
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex

	With frm1.vspdData
		For intRow = 1 To .MaxRows
			.Row = intRow
			.Col = C_HDD010T_INTREST_TYPE
			intIndex = .Value
			.col = C_HDD010T_INTREST_TYPE_NM
			.Value = intindex
		Next
	End With
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()   'sbk 

	With frm1.vspdData

        ggoSpread.Source = frm1.vspdData
        ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    'sbk

	   .ReDraw = false

       .MaxCols   = C_HDD010T_INTREST_RATE + 1                                            ' ☜:☜: Add 1 to Maxcols
	   .Col       = .MaxCols                                                              ' ☜:☜: Hide maxcols
       .ColHidden = True                                                                  ' ☜:☜:

       .Col = C_HDD010T_BORW_CD
       .ColHidden = True
       .Col = C_HDD010T_INTREST_TYPE
       .ColHidden = True

       .MaxRows = 0
       ggoSpread.ClearSpreadData

       Call GetSpreadColumnPos("A") 'sbk

       Call  AppendNumberPlace("6","3","0")
       Call  AppendNumberPlace("7","6","0")
       Call  AppendNumberPlace("8","2","2")

             ggoSpread.SSSetEdit     C_HDD010T_BORW_CD,              "",5
             ggoSpread.SSSetEdit     C_HDD010T_BORW_NM,              "대부코드명",12,,,50,2
             ggoSpread.SSSetButton   C_HDD010T_BORW_NM_POP
             ggoSpread.SSSetCombo    C_HDD010T_INTREST_TYPE,         "",5
             ggoSpread.SSSetCombo    C_HDD010T_INTREST_TYPE_NM,      "이자구분",10,,False
             ggoSpread.SSSetDate     C_HDD010T_BORW_DT,              "대부일", 11,2,  parent.gDateFormat
             ggoSpread.SSSetDate     C_HDD010T_EXPIR_DT,             "만기일", 11,2,  parent.gDateFormat
             ggoSpread.SSSetFloat    C_HDD010T_RESRV_DUR,            "거치기간" ,8,"6", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
             ggoSpread.SSSetFloat    C_HDD010T_BORW_TOT_AMT,         "대부총액" ,15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z","1"
             ggoSpread.SSSetFloat    C_HDD010T_PAY_INTCHNG_AVG,      "평균상환액" ,12, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
             ggoSpread.SSSetFloat    C_HDD010T_TOT_INTCHNG_CNT,      "급여상환횟수" ,12,"7", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
             ggoSpread.SSSetFloat    C_HDD010T_PAY_INTCHNG_AMT,      "급여상환액" ,12, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
             ggoSpread.SSSetFloat    C_HDD010T_BONUS_INTCHNG_CNT,    "상여상환횟수" ,12,"7", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
             ggoSpread.SSSetFloat    C_HDD010T_BONUS_INTCHNG_AMT,    "상여상환액" ,12, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
             ggoSpread.SSSetFloat    C_HDD010T_INTREST_RATE,         "이자율" ,8,"8", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"

        Call ggoSpread.MakePairsColumn(C_HDD010T_BORW_NM,C_HDD010T_BORW_NM_POP)    'sbk

        Call ggoSpread.SSSetColHidden(C_HDD010T_BORW_CD,C_HDD010T_BORW_CD,True)
        Call ggoSpread.SSSetColHidden(C_HDD010T_INTREST_TYPE,C_HDD010T_INTREST_TYPE,True)

	   .ReDraw = true

       Call SetSpreadLock

    End With

End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
    .vspdData.ReDraw = False
         ggoSpread.SpreadLock    C_HDD010T_BORW_CD, -1, C_HDD010T_BORW_CD, -1
         ggoSpread.SpreadLock    C_HDD010T_BORW_NM, -1, C_HDD010T_BORW_NM, -1
         ggoSpread.SpreadLock    C_HDD010T_BORW_NM_POP, -1, C_HDD010T_BORW_NM_POP, -1
         ggoSpread.SpreadLock    C_HDD010T_INTREST_TYPE, -1, C_HDD010T_INTREST_TYPE, -1
         ggoSpread.SpreadLock    C_HDD010T_INTREST_TYPE_NM, -1, C_HDD010T_INTREST_TYPE_NM, -1
         ggoSpread.SpreadLock    C_HDD010T_BORW_DT, -1, C_HDD010T_BORW_DT, -1
         ggoSpread.SpreadLock    C_HDD010T_PAY_INTCHNG_AVG, -1, C_HDD010T_PAY_INTCHNG_AVG, -1
         ggoSpread.SSSetRequired C_HDD010T_BORW_TOT_AMT , -1, -1
         ggoSpread.SSSetRequired C_HDD010T_TOT_INTCHNG_CNT , -1, -1
         ggoSpread.SSSetRequired C_HDD010T_PAY_INTCHNG_AMT , -1, -1
         ggoSpread.SSSetRequired C_HDD010T_BONUS_INTCHNG_CNT , -1, -1
         ggoSpread.SSSetRequired C_HDD010T_BONUS_INTCHNG_AMT , -1, -1
         ggoSpread.SSSetRequired C_HDD010T_INTREST_RATE , -1, -1
         ggoSpread.SSSetProtected   .vspdData.MaxCols   , -1, -1
    .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    .vspdData.ReDraw = False
         ggoSpread.SSSetRequired		C_HDD010T_BORW_NM, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_HDD010T_INTREST_TYPE_NM, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_HDD010T_BORW_DT, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_HDD010T_BORW_TOT_AMT, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected	    C_HDD010T_PAY_INTCHNG_AVG, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_HDD010T_TOT_INTCHNG_CNT, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_HDD010T_PAY_INTCHNG_AMT, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_HDD010T_BONUS_INTCHNG_CNT, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_HDD010T_BONUS_INTCHNG_AMT, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_HDD010T_INTREST_RATE, pvStartRow, pvEndRow
    .vspdData.ReDraw = True

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

            C_HDD010T_BORW_CD = iCurColumnPos(1)
            C_HDD010T_BORW_NM = iCurColumnPos(2)
            C_HDD010T_BORW_NM_POP = iCurColumnPos(3)
            C_HDD010T_INTREST_TYPE = iCurColumnPos(4)
            C_HDD010T_INTREST_TYPE_NM = iCurColumnPos(5)
            C_HDD010T_BORW_DT = iCurColumnPos(6)
            C_HDD010T_EXPIR_DT = iCurColumnPos(7)
            C_HDD010T_RESRV_DUR = iCurColumnPos(8)
            C_HDD010T_BORW_TOT_AMT = iCurColumnPos(9)
            C_HDD010T_PAY_INTCHNG_AVG = iCurColumnPos(10)
            C_HDD010T_TOT_INTCHNG_CNT = iCurColumnPos(11)
            C_HDD010T_PAY_INTCHNG_AMT = iCurColumnPos(12)
            C_HDD010T_BONUS_INTCHNG_CNT = iCurColumnPos(13)
            C_HDD010T_BONUS_INTCHNG_AMT = iCurColumnPos(14)
            C_HDD010T_INTREST_RATE = iCurColumnPos(15)
            
    End Select    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format

    Call  ggoOper.FormatField(Document, "A", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	Call  ggoOper.LockField(Document, "N")											'⊙: Lock Field

    Call InitSpreadSheet                                                            'Setup the Spread sheet

    Call InitVariables                                                              'Initializes local global variables

    Call  FuncGetAuth(gStrRequestMenuID,  parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
	Call SetToolbar("1100110100101111")												'⊙: Set ToolBar
    frm1.txtEmp_no.focus

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

    ggoSpread.ClearSpreadData

    Call InitVariables															'⊙: Initializes local global variables

    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If
    If  txtEmp_no_Onchange()  then
        Exit Function
    End If

    Call MakeKeyStream("X")

    Call  DisableToolBar( parent.TBC_QUERY)
	If DbQuery = False Then
		Call  RestoreToolBar()
		Exit Function
	End if													'☜: Query db data

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

	Call SetToolbar("1110111100111111")							                 '⊙: Set ToolBar
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
    End If												                  '☜: Delete db data

    FncDelete=  True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD
    Dim iRow
    Dim strBorw_dt
    Dim strExpir_dt
    Dim intPay_intchng
    Dim intBonus_intchng
    Dim intBorw_tot_amt
    Dim strIntrest_type
    Dim dblIntrest_rate
    Dim intPay_intchng_cnt, intBonus_intchng_cnt
	dim adddate
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

    If frm1.txtEmp_no.value="" Then
        Call  DisplayMsgBox("970021","X",frm1.txtEmp_no.alt,"X") '사번은 필수 입력사항입니다.
        frm1.txtEmp_no.focus ' go to
        Exit Function
    End If

    With Frm1.vspdData
        For iRow = 1 To  .MaxRows
            .Row = iRow
            .Col = 0
           Select Case .Text
               Case  ggoSpread.InsertFlag,  ggoSpread.UpdateFlag
					.Col = C_HDD010T_BORW_CD

					If IsNull(Trim(.Text)) OR Trim(.Text) = "" Then
						Call DisplayMsgBox("970000", "X","대부코드","x")
						Exit Function
					end if
					               
   	                .Col = C_HDD010T_BORW_DT
                    strBorw_dt  = .Text
   	                .Col = C_HDD010T_EXPIR_DT
                    strExpir_dt = .Text

                    If (strExpir_dt <> "" ) And  CompareDateByFormat(strBorw_dt,strExpir_dt,"만기일자","대부일자","800130", parent.gDateFormat, parent.gComDateType,True) = False then	                    
	                    .Row = iRow
  	                    .Col = C_HDD010T_EXPIR_DT
  	                    .Action=0
                        Set gActiveElement = document.activeElement
                        Exit Function
                    End if

                    .Col = C_HDD010T_RESRV_DUR                      '거치기간이 null이면 zero 처리 
 
					adddate= UNIDateAdd("m",.Text,strBorw_dt,  parent.gDateFormat)
					if  CompareDateByFormat(adddate,strExpir_dt,"만기일","거치기간","972001", parent.gDateFormat, parent.gComDateType,True)= False   then
						.Row = iRow
  	                    .Action=0
                        Set gActiveElement = document.activeElement
                        Exit Function
                    End if
                    
                    If .Text = ""  Then
                       .Text = 0
                    End If

                    .Col = C_HDD010T_PAY_INTCHNG_AMT   '급여상환액 
                    intPay_intchng   =  UNICDbl(.Text)
                    .Col = C_HDD010T_BONUS_INTCHNG_AMT '상여상환액 
                    intBonus_intchng =  UNICDbl(.Text)
                    .Col = C_HDD010T_TOT_INTCHNG_CNT   '급여상환횟수 
                    intPay_intchng_cnt = .Text
                    .Col = C_HDD010T_BONUS_INTCHNG_CNT '상여상환횟수 
                    intBonus_intchng_cnt = .Text
                    .Col = C_HDD010T_BORW_TOT_AMT      '대부총액  
		            intBorw_tot_amt =  UNICDbl(.Text)

		            If  IsNull(intBorw_tot_amt) or intBorw_tot_amt = 0 Then
                        Call  DisplayMsgBox("800443","X","대부총액","0")
                        .Action = 0 ' go to
                        Exit Function
                    End If

		            If (intPay_intchng_cnt <> 0 and intPay_intchng = 0) Then  '
                        Call  DisplayMsgBox("800493","X","급여","X") '급여상환횟수가 > 0이므로 급여상환액을 입력하세요.
					    .Col = C_HDD010T_PAY_INTCHNG_AMT  
					    .Action=0
                        Exit Function
                    End If    

		            If (intBonus_intchng_cnt <> 0 and intBonus_intchng = 0) Then  '
                        Call  DisplayMsgBox("800493","X","상여","X") '상여여상환횟수가 > 0이므로 상여상환액을 입력하세요.
					    .Col = C_HDD010T_BONUS_INTCHNG_AMT  
					    .Action=0
                        Exit Function
                    End If    

                    If	((intPay_intchng*intPay_intchng_cnt + intBonus_intchng*intBonus_intchng_cnt) > intBorw_tot_amt) Then
                        Call  DisplayMsgBox("800129","X","X","X")	'급여/상여 상환액을 조정하시오.
                        .Col = C_HDD010T_PAY_INTCHNG_AMT
                        .Action = 0 ' go to
                        Exit Function
                    End If

                    If	((intPay_intchng*intPay_intchng_cnt + intBonus_intchng*intBonus_intchng_cnt) < intBorw_tot_amt) Then
                        IntRetCD =  DisplayMsgBox("800498", parent.VB_YES_NO,"X","X")	'미수금이 발생했습니다.그래도 입력?
                        If IntRetCD = vbNo Then     ' No 이면 
                           .Col = C_HDD010T_PAY_INTCHNG_AMT
                           .Action = 0 ' go to
                           Exit Function
                        End If
                    End If                           

                    .Col = C_HDD010T_INTREST_TYPE
                    strIntrest_type   = .Text
                    .Col = C_HDD010T_INTREST_RATE
                    dblIntrest_rate   =  UNICDbl(.Text)

		            If IsNull(dblIntrest_rate) Then
		                dblIntrest_rate = 0
		            End If
		                                                                    '이자구분이 Y인 경우 이자율은 zero면 에러 
		            If strIntrest_type = "Y" And dblIntrest_rate = 0 then   '이자구분이 N인 경우 이자율은 zero보다 커야함 
                        Call  DisplayMsgBox("141157","X","이자율","X")   '이자율을 입력하십시오.
                        .Action = 0 ' go to
          	            Exit Function
		            Else
		                If strIntrest_type = "N" And dblIntrest_rate > 0 then
                            Call  DisplayMsgBox("800235","X","X","X")	'이자율을 입력할 수 없습니다.
                            .Action = 0 ' go to
                            Exit Function
                        End If
 		            End If
           End Select
        Next
    End With

    Call MakeKeyStream("X")

    If DbSave = False Then
       Exit Function
    End If				                                                    '☜: Save db data

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

	With Frm1

		If .vspdData.ActiveRow > 0 Then
			.vspdData.ReDraw = False

			 ggoSpread.Source = frm1.vspdData
			 ggoSpread.CopyRow
			SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow
            .vspdData.Row  = .vspdData.ActiveRow

            .vspdData.Col  = C_HDD010T_BORW_CD
            .vspdData.Text = ""
            .vspdData.Col  = C_HDD010T_BORW_NM
            .vspdData.Text = ""
            .vspdData.Col  = C_HDD010T_INTREST_TYPE
            .vspdData.Text = ""
            .vspdData.Col  = C_HDD010T_INTREST_TYPE_NM
            .vspdData.Text = ""
            .vspdData.Col  = C_HDD010T_BORW_DT
            .vspdData.Text = ""

			.vspdData.ReDraw = True
			.vspdData.focus
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
            
            .vspdData.col=C_HDD010T_BORW_DT
            .vspdData.text=  UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
            .vspdData.col=C_HDD010T_EXPIR_DT
            .vspdData.text=""
            .vspdData.col=C_HDD010T_RESRV_DUR
            .vspdData.text=0
            .vspdData.col=C_HDD010T_BORW_TOT_AMT
            .vspdData.text=0
            .vspdData.col=C_HDD010T_PAY_INTCHNG_AVG 
            .vspdData.text=0
            .vspdData.col=C_HDD010T_TOT_INTCHNG_CNT
            .vspdData.text=0
            .vspdData.col=C_HDD010T_PAY_INTCHNG_AMT
            .vspdData.text=0
            .vspdData.col=C_HDD010T_BONUS_INTCHNG_CNT
            .vspdData.text=0     
            .vspdData.col=C_HDD010T_BONUS_INTCHNG_AMT
            .vspdData.text=0
            .vspdData.col=C_HDD010T_INTREST_RATE
            .vspdData.text=0
       Next
       
       .vspdData.ReDraw = True
    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
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

    ggoSpread.Source = Frm1.vspdData
    ggoSpread.ClearSpreadData

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

    ggoSpread.Source = Frm1.vspdData
    ggoSpread.ClearSpreadData

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
              Case  ggoSpread.InsertFlag                                      '☜: Insert
                                                                   strVal = strVal & "C" & parent.gColSep
                                                                   strVal = strVal & lRow & parent.gColSep
                                                                   strVal = strVal & Trim(.txtEmp_no.value) & parent.gColSep
                    .vspdData.Col = C_HDD010T_BORW_CD            : strVal = strVal & Trim(.vspdData.Text)  & parent.gColSep
                    .vspdData.Col = C_HDD010T_INTREST_TYPE	     : strVal = strVal & Trim(.vspdData.Text)  & parent.gColSep
                    .vspdData.Col = C_HDD010T_BORW_DT            : strVal = strVal & Trim(.vspdData.Text)  & parent.gColSep
                    .vspdData.Col = C_HDD010T_EXPIR_DT	         : strVal = strVal & Trim(.vspdData.Text)  & parent.gColSep
                    .vspdData.Col = C_HDD010T_RESRV_DUR	         : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD010T_BORW_TOT_AMT	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD010T_PAY_INTCHNG_AVG	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD010T_TOT_INTCHNG_CNT	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD010T_PAY_INTCHNG_AMT	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD010T_BONUS_INTCHNG_CNT	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD010T_BONUS_INTCHNG_AMT	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD010T_INTREST_RATE       : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                                   strVal = strVal & "U" & parent.gColSep
                                                                   strVal = strVal & lRow & parent.gColSep
                                                                   strVal = strVal & Trim(.txtEmp_no.value) & parent.gColSep
                    .vspdData.Col = C_HDD010T_BORW_CD            : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD010T_INTREST_TYPE	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD010T_BORW_DT            : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD010T_EXPIR_DT	         : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD010T_RESRV_DUR	         : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD010T_BORW_TOT_AMT	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD010T_PAY_INTCHNG_AVG	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD010T_TOT_INTCHNG_CNT	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD010T_PAY_INTCHNG_AMT	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD010T_BONUS_INTCHNG_CNT	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD010T_BONUS_INTCHNG_AMT	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD010T_INTREST_RATE       : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.DeleteFlag                                      '☜: Delete

                                                                  strDel = strDel & "D" & parent.gColSep
                                                                  strDel = strDel & lRow & parent.gColSep
                                                                  strDel = strDel & Trim(.txtEmp_no.value) & parent.gColSep
                    .vspdData.Col = C_HDD010T_BORW_CD           : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD010T_INTREST_TYPE	    : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDD010T_BORW_DT           : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
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
	strVal = strVal & "&txtGlNo=" & Trim(frm1.txtLcNo.value)             '☜:

	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic

	DbDelete = True                                                              '⊙: Processing is NG

End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()

	lgIntFlgMode      =  parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode

    Frm1.txtName.focus

	Call SetToolbar("1100111100111111")												'⊙: Set ToolBar

    Call  ggoOper.LockField(Document, "Q")
	frm1.vspdData.focus
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
	Call InitVariables
    ggoSpread.Source = Frm1.vspdData
    Frm1.vspdData.MaxRows = 0
    ggoSpread.ClearSpreadData
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
' Name : OpenEmptName()
' Desc : developer describe this line
'========================================================================================================

Function OpenEmptName(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""'frm1.txtName.value			' Name Cindition
	Else 'spread
        frm1.vspdData.Col = C_EMP_NO
		arrParam(0) = frm1.vspdData.Text			' Code Condition
        frm1.vspdData.Col = C_NAME
	    arrParam(1) = ""'frm1.vspdData.Text			' Name Cindition
	End If
	arrParam(2) = lgUsrIntCd                    ' 자료권한 Condition

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
		Call SubSetCondEmp(arrRet, iWhere)
	End If

End Function

'======================================================================================================
'	Name : SubSetCondEmp()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetCondEmp(Byval arrRet, Byval iWhere)
    Dim strVal

	With frm1
		If iWhere = 0 Then 'TextBox(Condition)
			.txtEmp_no.value = arrRet(0)
			.txtName.value = arrRet(1)
			.txtDept_nm.value = arrRet(2)
			.txtRoll_pstn.value = arrRet(3)
			.txtEntr_dt.Text = arrRet(5)
			.txtPay_grd.value = arrRet(4)

			Call txtEmp_no_Onchange()
			.txtEmp_no.focus
		Else 'spread
			.vspdData.Col = C_NAME
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_DEPT_CD
			.vspdData.Text = arrRet(2)
			.vspdData.Col = C_EMP_NO
			.vspdData.Text = arrRet(0)
			.vspdData.action =0
		End If
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
	    Case C_HDD010T_BORW_NM_POP
	        arrParam(0) = "대부코드조회 팝업"			    ' 팝업 명칭 
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
		frm1.vspdData.Col = C_HDD010T_BORW_NM
		frm1.vspdData.action = 0	
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
		    Case C_HDD010T_BORW_NM_POP
		        .vspdData.Col = C_HDD010T_BORW_CD
		    	.vspdData.text = arrRet(0)
		    	.vspdData.Col = C_HDD010T_BORW_NM
		    	.vspdData.text = arrRet(1)
				.vspdData.action = 0
        End Select

	End With
End Function

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col
	Select Case Col
	    Case C_HDD010T_BORW_NM_POP
                    Call OpenCode(frm1.vspdData.Text, C_HDD010T_BORW_NM_POP, Row)
    End Select
End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Function vspdData_Change(ByVal Col , ByVal Row)

    Dim iDx
    Dim IntRetCD
    Dim intPay_intchng_avg   ,intBorw_Tot_Amt      , intTot_intchng_cnt  ,intPay_intchng_amt 
    Dim intBonus_intchng_cnt ,intBonus_intchng_amt , outstanding_amt

   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    With frm1
        Select Case Col
	         Case C_HDD010T_BORW_NM
                    IntRetCD =  CommonQueryRs(" minor_cd,minor_nm "," b_minor "," major_cd=" & FilterVar("H0043", "''", "S") & " And minor_nm =  " & FilterVar(.vspdData.Text, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
                    If IntRetCD=False And Trim(.vspdData.Text)<>""  Then
						Call DisplayMsgBox("970000", "X","대부코드","x")
    	                .vspdData.Col = C_HDD010T_BORW_CD                        
                        .vspdData.Text=""
                        exit function
                    ElseIf  CountStrings(lgF0, Chr(11) ) > 1 Then    ' 같은명일 경우 pop up
                        Call OpenCode(.vspdData.Text, C_HDD010T_BORW_NM_POP, Row)
                    Else
    	                .vspdData.Col = C_HDD010T_BORW_CD
                        .vspdData.Text=Trim(Replace(lgF0,Chr(11),""))
                    End If

             Case C_HDD010T_TOT_INTCHNG_CNT     '급여상환횟수에 따른 평균상환금액 자동계산 

                    .vspdData.Col = C_HDD010T_BORW_TOT_AMT      '대부총액  
		            intBorw_Tot_Amt =  UNICDbl(.vspdData.Text)

		            If  IsNull(intBorw_Tot_Amt) or intBorw_Tot_Amt = 0 Then
                        Call  DisplayMsgBox("800489","X","대부총액","X")
                        .Action = 0 ' go to
                        Exit Function
                    End If

                    .vspdData.Col = C_HDD010T_TOT_INTCHNG_CNT      '급여상환횟수  
                    intTot_intchng_cnt = .vspdData.Text
  		            If intTot_intchng_cnt = 0 Then       '급여상환횟수가 0 일경우...
						.vspdData.Col = C_HDD010T_PAY_INTCHNG_AVG   
						.vspdData.Text= 0
                    Else
                       intPay_intchng_avg = intBorw_Tot_Amt / intTot_intchng_cnt

		               .vspdData.Col = C_HDD010T_PAY_INTCHNG_AVG  
                       .vspdData.Text=  UNIFormatNumber(intPay_intchng_avg,  ggAmtOfMoney.DecPoint,-2,0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
					End If

             Case C_HDD010T_BONUS_INTCHNG_CNT     '상여상환횟수에 따른 상여상환금액 설정 

                    .vspdData.Col = C_HDD010T_BORW_TOT_AMT      '대부총액  
		            intBorw_Tot_Amt =  UNICDbl(.vspdData.Text)

		            If  IsNull(intBorw_Tot_Amt) or intBorw_Tot_Amt = 0 Then
                        Call  DisplayMsgBox("800489","X","대부총액","X")
                        .Action = 0 ' go to
                        Exit Function
                    End If

                    .vspdData.Col = C_HDD010T_PAY_INTCHNG_AMT
                    intPay_intchng_amt  =  UNICDbl(frm1.vspdData.Text)

                    .vspdData.Col = C_HDD010T_TOT_INTCHNG_CNT
                    intTot_intchng_cnt  =  UNICDbl(frm1.vspdData.Text)

                    .vspdData.Col = C_HDD010T_BONUS_INTCHNG_CNT
                    intBonus_intchng_cnt  =  UNICDbl(frm1.vspdData.Text)

   					outstanding_amt = intBorw_Tot_Amt - (intPay_intchng_amt * intTot_intchng_cnt)

		            If intBonus_intchng_cnt > 0 Then       '상여상환횟수가 0 보다 클떄 
		                If outstanding_amt = 0 Then
                            Call  DisplayMsgBox("800482","X","X","X")  '☜ : 급여상환액이 이미 대부총액과 일치합니다. 급여상환액을 조정하세요.
		                    .vspdData.Col = C_HDD010T_BONUS_INTCHNG_CNT  
                            .vspdData.Text= 0
		                    .vspdData.Col = C_HDD010T_BONUS_INTCHNG_AMT  
                            .vspdData.Text= 0                            
 						    .vspdData.Action=0
 						Else
		                    intBonus_intchng_amt = outstanding_amt / intBonus_intchng_cnt 

		                    .vspdData.Col = C_HDD010T_BONUS_INTCHNG_AMT  
                            .vspdData.Text=  UNIFormatNumber(intBonus_intchng_amt,  ggAmtOfMoney.DecPoint,-2,0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
                        End If
                    End If

             Case C_HDD010T_BORW_TOT_AMT   '대부총액이 수정되면 급여/상여금액은 초기화...
                    .vspdData.Col = C_HDD010T_PAY_INTCHNG_AMT  
                    .vspdData.Text= 0
                   	.vspdData.Col = C_HDD010T_TOT_INTCHNG_CNT  
                   	.vspdData.Text= 0
                   	.vspdData.Col = C_HDD010T_PAY_INTCHNG_AVG  
                   	.vspdData.Text= 0
                    .vspdData.Col = C_HDD010T_BONUS_INTCHNG_CNT
                    .vspdData.Text= 0
                   	.vspdData.Col = C_HDD010T_BONUS_INTCHNG_AMT  
                   	.vspdData.Text= 0

             Case C_HDD010T_PAY_INTCHNG_AMT    '급여상환액이 수정되면 상여금액초기화 
                    intPay_intchng_amt  =  UNICDbl(.vspdData.Text)

                    .vspdData.Col = C_HDD010T_BORW_TOT_AMT      '대부총액  
		            intBorw_Tot_Amt =  UNICDbl(.vspdData.Text)

		            If  IsNull(intBorw_Tot_Amt) or intBorw_Tot_Amt = 0 Then
                        Call  DisplayMsgBox("800489","X","대부총액","X")
                        .Action = 0 ' go to
                        Exit Function
                    End If

                    .vspdData.Col = C_HDD010T_TOT_INTCHNG_CNT
                    intTot_intchng_cnt  =  UNICDbl(.vspdData.Text)
		            .vspdData.Col = C_HDD010T_PAY_INTCHNG_AVG  
                    intPay_intchng_avg  =  UNICDbl(.vspdData.Text)

                    IF intPay_intchng_amt > intPay_intchng_avg Then
                        Call  DisplayMsgBox("800129","X","X","X")	'급여/상여 상환액을 조정하시오.
                   	    .vspdData.Action= 0
                        Exit Function
                    End If    

                    .vspdData.Col = C_HDD010T_BONUS_INTCHNG_CNT
                    .vspdData.Text= 0
                    .vspdData.Col = C_HDD010T_BONUS_INTCHNG_AMT  
                   	.vspdData.Text= 0

             Case C_HDD010T_BONUS_INTCHNG_AMT    '상여상환액이 수정되면 대부총액과 체크 
                    intBonus_intchng_amt =  UNICDbl(.vspdData.Text)

                    .vspdData.Col = C_HDD010T_BORW_TOT_AMT      '대부총액  
		            intBorw_Tot_Amt =  UNICDbl(.vspdData.Text)

		            If  IsNull(intBorw_Tot_Amt) or intBorw_Tot_Amt = 0 Then
                        Call  DisplayMsgBox("800489","X","대부총액","X")
                        .Action = 0 ' go to
                        Exit Function
                    End If

                    .vspdData.Col = C_HDD010T_PAY_INTCHNG_AMT
                    intPay_intchng_amt  =  UNICDbl(frm1.vspdData.Text)

                    .vspdData.Col = C_HDD010T_TOT_INTCHNG_CNT
                    intTot_intchng_cnt  =  UNICDbl(frm1.vspdData.Text)

                    .vspdData.Col = C_HDD010T_BONUS_INTCHNG_CNT
                    intBonus_intchng_cnt  =  UNICDbl(frm1.vspdData.Text)

    					outstanding_amt = (intPay_intchng_amt * intTot_intchng_cnt) + (intBonus_intchng_amt * intBonus_intchng_cnt)
		            
		                If intBorw_Tot_Amt < outstanding_amt Then
                            Call  DisplayMsgBox("800131","X","X","X")  '☜ : 상여상환액을 조정하세요.
		                    .vspdData.Col = C_HDD010T_BONUS_INTCHNG_AMT  
 						    .vspdData.Action=0
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
            Case C_HDD010T_INTREST_TYPE_NM
                .Col = Col
                intIndex = .Value
				.Col = C_HDD010T_INTREST_TYPE
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

    Call SetPopupMenuItemInf("1101111111")
    
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
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
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
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
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
		frm1.txtEmp_no.value = ""
        frm1.txtName.value = ""
        frm1.txtDept_nm.value = ""
        frm1.txtRoll_pstn.value = ""
        frm1.txtPay_grd.value = ""
        frm1.txtEntr_dt.Text = ""
		Frm1.imgPhoto.src = ""
    Else
	    IntRetCd =  FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    If  IntRetCd < 0 then
			strVal = "../../../CShared/image/default_picture.jpg"
			Frm1.imgPhoto.src = strVal
				    
	        If  IntRetCd = -1 then
    			Call  DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call  DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            end if
            frm1.txtName.value = ""
            frm1.txtDept_nm.value = ""
            frm1.txtRoll_pstn.value = ""
            frm1.txtPay_grd.value = ""
            frm1.txtEntr_dt.text = ""
            ggoSpread.Source = Frm1.vspdData
            ggoSpread.ClearSpreadData
            call InitVariables()
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true
        Else
            frm1.txtName.value = strName
            frm1.txtDept_nm.value = strDept_nm
            frm1.txtRoll_pstn.value = strRoll_pstn
            frm1.txtPay_grd.value = strPay_grd1 & "-" & strPay_grd2
            frm1.txtEntr_dt.Text =  UNIDateClientFormat(strEntr_dt)
            
			Call CommonQueryRs(" COUNT(*) "," HAA070T "," emp_no= " & FilterVar( Frm1.txtEmp_no.value, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    		if   Replace(lgF0, Chr(11), "") > 0  then
				strVal = "../../ComASP/CPictRead.asp" & "?txtKeyValue=" & Frm1.txtEmp_no.value '☜: query key
				strVal = strVal     & "&txtDKeyValue=" & "default"                            '☜: default value
				strVal = strVal     & "&txtTable="     & "HAA070T"                            '☜: Table Name
				strVal = strVal     & "&txtField="     & "Photo"	                          '☜: Field
				strVal = strVal     & "&txtKey="       & "Emp_no"	                          '☜: Key
			else
				strVal = "../../../CShared/image/default_picture.jpg"
			end if
			
            Frm1.imgPhoto.src = strVal
        End if
    End if

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00 %>></TD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>대부금등록</font></td>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100% COLSPAN=2></TD>
				</TR>
				<TR>
    	            <TD HEIGHT=20 WIDTH=7%>
			            <TABLE <%=LR_SPACE_TYPE_40%>>
			                <TR HEIGHT=69>
			                    <TD>
                                    <img src="../../../CShared/image/default_picture.jpg" name="imgPhoto" WIDTH=60 HEIGHT=69 HSPACE=10 VSPACE=0 BORDER=1>
			                    </TD>
			                </TR>
			            </TABLE>
    	            </TD>
    	            <TD HEIGHT=20 WIDTH=90%>
						<FIELDSET CLASS="CLSFLD">
                           <TABLE <%=LR_SPACE_TYPE_40%>>
			    	        <TR>
			    	    		<TD CLASS="TD5" NOWRAP>사번</TD>
			    	    		<TD CLASS="TD6"><INPUT NAME="txtEmp_no"  SIZE=13 MAXLENGTH=13 ALT="사번" TYPE="Text"  tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmptName('0')"></TD>
			    	        	<TD CLASS="TD5" NOWRAP>성명</TD>
			    	    		<TD CLASS="TD6"><INPUT NAME="txtName"  SIZE=20  MAXLENGTH=30 ALT="성명" TYPE="Text"  tag="14XXXU"></TD>
			            	</TR>
			            	<TR>
			            		<TD CLASS="TD5" NOWRAP>부서명</TD>
			            		<TD CLASS="TD6"><INPUT NAME="txtDept_nm" SIZE=20  MAXLENGTH=40 ALT="부서명" TYPE="Text"  tag="14XXXU"></TD>
			            		<TD CLASS="TD5" NOWRAP>직위</TD>
			            		<TD CLASS="TD6"><INPUT NAME="txtRoll_pstn" SIZE=20  MAXLENGTH=50  ALT="직위" TYPE="Text"  tag="14XXXU"></TD>
			            	</TR>
			            	<TR>
			            		<TD CLASS="TD5" NOWRAP>입사일</TD>
			            		<TD CLASS="TD6"><script language =javascript src='./js/h5204ma1_fpDateTime2_txtEntr_dt.js'></script></TD>
			            		<TD CLASS="TD5" NOWRAP>급호</TD>
			            		<TD CLASS="TD6"><INPUT NAME="txtPay_grd" SIZE=20  MAXLENGTH=50 ALT="급호" TYPE="Text"  tag="14XXXU"></TD>
			            	</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> COLSPAN=2></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP COLSPAN=2>
                        <TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT=100%>
									<script language =javascript src='./js/h5204ma1_vaSpread1_vspdData.js'></script>
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
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
