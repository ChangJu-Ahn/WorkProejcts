<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        :
*  3. Program ID           : h5203ma1
*  4. Program Name         : 월저축불입현황 조회 
*  5. Program Desc         : 월저축불입현황 조회 
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
Const BIZ_PGM_ID      = "h5203mb1.asp"						           '☆: Biz Logic ASP Name
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

Dim C_NAME
Dim C_EMP_NO
Dim C_HDC020T_SAVE_CD
Dim C_HDC020T_BANK_ACCNT
Dim C_HDC020T_SAVE_TYPE
Dim C_HDC020T_SCRIPT_AMT
Dim C_HDC020T_SCRIPT_CNT
Dim C_HDC020T_BALN_CNT
Dim C_HDC010T_TOT_SCRIPT_CNT
Dim C_HDC010T_BANK_CD
Dim C_FAA090T_BANK_NAME
Dim C_HDC010T_EXPIR_AMT
Dim C_HDC010T_NEW_DT
Dim C_HDC010T_EXPIR_DT
Dim C_HDC010T_REVOKE_DT

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  

    C_NAME = 1
    C_EMP_NO = 2
    C_HDC020T_SAVE_CD = 3
    C_HDC020T_BANK_ACCNT = 4
    C_HDC020T_SAVE_TYPE = 5
    C_HDC020T_SCRIPT_AMT = 6
    C_HDC020T_SCRIPT_CNT = 7
    C_HDC020T_BALN_CNT = 8
    C_HDC010T_TOT_SCRIPT_CNT = 9    
    C_HDC010T_BANK_CD = 10
    C_FAA090T_BANK_NAME = 11
    C_HDC010T_EXPIR_AMT = 12
    C_HDC010T_NEW_DT = 13
    C_HDC010T_EXPIR_DT = 14
    C_HDC010T_REVOKE_DT = 15

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
    
	frm1.txtScript_yymm_dt.focus
	Call ggoOper.FormatDate(frm1.txtScript_yymm_dt,  parent.gDateFormat, 2)
	Call ExtractDateFrom("<%=GetSvrDate%>", parent.gServerDateFormat ,  parent.gServerDateType ,strYear,strMonth,strDay)	
	frm1.txtScript_yymm_dt.Year		=  strYear
	frm1.txtScript_yymm_dt.Month	=  strMonth
	frm1.txtScript_yymm_dt.Day		=  strDay	
End Sub

'========================================================================================================
' Name : LoadInfTB19029()
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()

<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "H","NOCOOKIE","MA") %>

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

    lgKeyStream       = Trim(Frm1.txtScript_yymm_dt.Year & Right("0" & Frm1.txtScript_yymm_dt.Month,2)) & parent.gColSep       'You Must append one character( parent.gColSep)
    lgKeyStream       = lgKeyStream & Trim(Frm1.txtSave_cd.value) & parent.gColSep
    lgKeyStream       = lgKeyStream & Trim(Frm1.txtSave_type.value) & parent.gColSep
    lgKeyStream       = lgKeyStream & Trim(Frm1.txtEmp_no.value) & parent.gColSep
    
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

    Call  CommonQueryRs(" MAJOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0041", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) '저축코드 
    iCodeArr = lgF0
    iNameArr = lgF1
    Call  SetCombo2(frm1.txtSave_cd,iCodeArr, iNameArr,Chr(11))

    Call  CommonQueryRs(" MAJOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0042", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) '저축구분 
    iCodeArr = lgF0
    iNameArr = lgF1
    Call  SetCombo2(frm1.txtSave_type,iCodeArr, iNameArr,Chr(11))

End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
    If Frm1.vspdData.MaxRows > 0 Then
        Call vspdData_Click(1, 1)
		Frm1.vspdData.focus
		
        Set gActiveElement = document.ActiveElement
	End If
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

       .MaxCols   = C_HDC010T_REVOKE_DT + 1                                               ' ☜:☜: Add 1 to Maxcols
	   .Col       = .MaxCols                                                              ' ☜:☜: Hide maxcols
       .ColHidden = True                                                                  ' ☜:☜:

       .MaxRows = 0
        ggoSpread.ClearSpreadData

       Call GetSpreadColumnPos("A") 'sbk

       Call  AppendNumberPlace("6","14","0")

             ggoSpread.SSSetEdit     C_NAME,                    "성명", 15
             ggoSpread.SSSetEdit     C_EMP_NO,                  "사번", 15
             ggoSpread.SSSetEdit     C_HDC020T_SAVE_CD,         "저축코드", 15
             ggoSpread.SSSetEdit     C_HDC020T_BANK_ACCNT,      "계좌번호", 15
             ggoSpread.SSSetEdit     C_HDC020T_SAVE_TYPE,       "저축구분", 12
             ggoSpread.SSSetFloat    C_HDC020T_SCRIPT_AMT,      "월불입액" ,15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
             ggoSpread.SSSetFloat    C_HDC020T_SCRIPT_CNT,      "불입횟수" ,10,"6", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
             ggoSpread.SSSetFloat    C_HDC020T_BALN_CNT,        "잔여불입횟수" ,10,"6", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
             ggoSpread.SSSetFloat    C_HDC010T_TOT_SCRIPT_CNT,  "총불입횟수" ,10,"6", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec

             ggoSpread.SSSetEdit     C_HDC010T_BANK_CD,         "은행코드", 10
             ggoSpread.SSSetEdit     C_FAA090T_BANK_NAME,       "은행명", 10
             ggoSpread.SSSetFloat    C_HDC010T_EXPIR_AMT,       "만기금액" ,15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
             ggoSpread.SSSetDate     C_HDC010T_NEW_DT,          "신규일자", 10,2,  parent.gDateFormat
             ggoSpread.SSSetDate     C_HDC010T_EXPIR_DT,        "만기일자", 10,2,  parent.gDateFormat
             ggoSpread.SSSetDate     C_HDC010T_REVOKE_DT,       "해약일자", 10,2,  parent.gDateFormat

        Call ggoSpread.SSSetColHidden(C_HDC010T_BANK_CD,C_HDC010T_BANK_CD,True)
        Call ggoSpread.SSSetColHidden(C_FAA090T_BANK_NAME,C_FAA090T_BANK_NAME,True)
        Call ggoSpread.SSSetColHidden(C_HDC010T_EXPIR_AMT,C_HDC010T_EXPIR_AMT,True)
        Call ggoSpread.SSSetColHidden(C_HDC010T_NEW_DT,C_HDC010T_NEW_DT,True)
        Call ggoSpread.SSSetColHidden(C_HDC010T_EXPIR_DT,C_HDC010T_EXPIR_DT,True)
        Call ggoSpread.SSSetColHidden(C_HDC010T_REVOKE_DT,C_HDC010T_REVOKE_DT,True)

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
             ggoSpread.SpreadLock       C_NAME, -1, C_NAME, -1
             ggoSpread.SpreadLock       C_EMP_NO, -1, C_EMP_NO, -1
             ggoSpread.SpreadLock       C_HDC020T_SAVE_CD, -1, C_HDC020T_SAVE_CD, -1
             ggoSpread.SpreadLock       C_HDC020T_BANK_ACCNT, -1, C_HDC020T_BANK_ACCNT, -1
             ggoSpread.SpreadLock       C_HDC020T_SAVE_TYPE, -1, C_HDC020T_SAVE_TYPE, -1
             ggoSpread.SpreadLock       C_HDC020T_SCRIPT_AMT, -1, C_HDC020T_SCRIPT_AMT, -1
             ggoSpread.SpreadLock       C_HDC020T_SCRIPT_CNT, -1, C_HDC020T_SCRIPT_CNT, -1
             ggoSpread.SpreadLock       C_HDC020T_BALN_CNT, -1, C_HDC020T_BALN_CNT, -1
             ggoSpread.SpreadLock       C_HDC010T_TOT_SCRIPT_CNT, -1, C_HDC010T_TOT_SCRIPT_CNT, -1
             ggoSpread.SpreadLock       C_HDC010T_BANK_CD, -1, C_HDC010T_BANK_CD, -1
             ggoSpread.SpreadLock       C_FAA090T_BANK_NAME, -1, C_FAA090T_BANK_NAME, -1
             ggoSpread.SpreadLock       C_HDC010T_EXPIR_AMT, -1, C_HDC010T_EXPIR_AMT, -1
             ggoSpread.SpreadLock       C_HDC010T_NEW_DT, -1, C_HDC010T_NEW_DT, -1
             ggoSpread.SpreadLock       C_HDC010T_EXPIR_DT, -1, C_HDC010T_EXPIR_DT, -1
             ggoSpread.SpreadLock       C_HDC010T_REVOKE_DT, -1, C_HDC010T_REVOKE_DT, -1
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

            C_NAME = iCurColumnPos(1)
            C_EMP_NO = iCurColumnPos(2)
            C_HDC020T_SAVE_CD = iCurColumnPos(3)
            C_HDC020T_BANK_ACCNT = iCurColumnPos(4)
            C_HDC020T_SAVE_TYPE = iCurColumnPos(5)
            C_HDC020T_SCRIPT_AMT = iCurColumnPos(6)
            C_HDC020T_SCRIPT_CNT = iCurColumnPos(7)
            C_HDC020T_BALN_CNT = iCurColumnPos(8)
            C_HDC010T_TOT_SCRIPT_CNT = iCurColumnPos(9)
            C_HDC010T_BANK_CD = iCurColumnPos(10)
            C_FAA090T_BANK_NAME = iCurColumnPos(11)
            C_HDC010T_EXPIR_AMT = iCurColumnPos(12)
            C_HDC010T_NEW_DT = iCurColumnPos(13)
            C_HDC010T_EXPIR_DT = iCurColumnPos(14)
            C_HDC010T_REVOKE_DT = iCurColumnPos(15)
            
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

    Call  FuncGetAuth(gStrRequestMenuID,  parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables

    Call SetDefaultVal
	Call SetToolbar("1100000000001111")												'⊙: Set ToolBar

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

    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field

    Call InitVariables															'⊙: Initializes local global variables

    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If

    If  txtEmp_no_Onchange then
        Exit Function
    End If

    Call MakeKeyStream("X")

    Call DisableToolBar( parent.TBC_QUERY)
	If DbQuery = False Then
		Call  RestoreToolBar()
		Exit Function
	End if

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

	Call SetToolbar("1100000000001111")												'⊙: Set ToolBar
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
			SetSpreadColor frm1.vspdData.ActiveRow
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

    Call InitVariables						    								 '⊙: Initializes local global variables

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
    Frm1.txtScript_yymm_dt.focus

	Call SetToolbar("1100000000011111")												'⊙: Set ToolBar
    Call InitData()
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
' Name : OpenEmp()
' Desc : developer describe this line
'========================================================================================================

Function OpenEmp()
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	arrParam(1) = ""'frm1.txtName.value			' Name Cindition
	arrParam(2) = lgUsrIntCd        			' Internal_cd

	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtEmp_no.focus	
		Exit Function
	Else
		Call SetEmp(arrRet)
	End If

End Function

'======================================================================================================
'	Name : SetEmp()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SetEmp(arrRet)
	With frm1
		.txtEmp_no.value = arrRet(0)
		.txtName.value = arrRet(1)

		Call  ggoOper.ClearField(Document, "2")					 '☜: Clear Contents  Field
        ggoSpread.ClearSpreadData
		.txtEmp_no.focus
		lgBlnFlgChgValue = False

	End With
End Sub

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)

    Dim iDx

   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

   	If Frm1.vspdData.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If CDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If

	 ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
 Sub vspdData_Click(ByVal Col , ByVal Row )
	Dim IntRetCD

    Call SetPopupMenuItemInf("0000111111")
    
    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If

	If lgOldRow <> Row Then

		frm1.vspdData.Col = 1
		frm1.vspdData.Row = row

		lgOldRow = Row

		With frm1

		    .vspdData.Col = C_HDC010T_BANK_CD
		    .txtHdc010t_bank_cd.value = .vspdData.Text

		    .vspdData.Col = C_FAA090T_BANK_NAME
		    .txtFaa090t_bank_name.value = .vspdData.Text

		    .vspdData.Col = C_HDC010T_EXPIR_AMT
		    .txtHdc010t_expir_amt.value = .vspdData.Text

		    .vspdData.Col = C_HDC010T_NEW_DT
		    .txtHdc010t_new_dt.Text = .vspdData.Text

		    .vspdData.Col = C_HDC010T_EXPIR_DT
		    .txtHdc010t_expir_dt.Text = .vspdData.Text

		    .vspdData.Col = C_HDC010T_REVOKE_DT
		    .txtHdc010t_revoke_dt.Text = .vspdData.Text

		End With
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
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : This function is called when cursor leave cell
'========================================================================================================
Sub vspdData_ScriptLeaveCell(Col,Row,NewCol,NewRow,Cancel)
	If NewRow <= 0 Or NewCol < 0 Then
		Exit Sub
	End If

	call vspdData_click(NewCol, NewRow)
		
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

Sub cboYesNo_OnChange()
    lgBlnFlgChgValue = True
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
Sub txtScript_yymm_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")
        frm1.txtScript_yymm_dt.Action = 7
        frm1.txtScript_yymm_dt.focus
    End If
End Sub
'=======================================================================================================
'   Event Name : txtScript_yymm_dt_Keypress(Key)
'   Event Desc : 3rd party control에서 Enter 키를 누르면 조회 실행 
'=======================================================================================================
Sub txtScript_yymm_dt_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00 %> ></TD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>월저축불입현황조회</font></td>
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
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
                           <TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
							    <TD CLASS=TD5 NOWRAP>불입년월</TD>
			                    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h5203ma1_fpDateTime12_txtScript_yymm_dt.js'></script></TD>
			    	    		<TD CLASS=TD5 NOWRAP>저축코드</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="txtSave_cd" ALT="저축코드" STYLE="WIDTH: 100px" TAG="11"><OPTION Value=""></OPTION></SELECT>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>사원</TD>
			    	    		<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no"  SIZE=13 MAXLENGTH=13 ALT="사번" TYPE="Text"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmp()">
			    	        	                     <INPUT NAME="txtName"  SIZE=20  MAXLENGTH=30 ALT="성명" TYPE="Text"  tag="14XXXU"></TD>
								<TD CLASS=TD5 NOWRAP>저축구분</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="txtSave_type" ALT="저축구분" STYLE="WIDTH: 100px" TAG="11"><OPTION Value=""></OPTION></SELECT></TD>
							</TR>
						  </TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
                        <TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT=100%>
									<script language =javascript src='./js/h5203ma1_vaSpread1_vspdData.js'></script>
								</TD>
							</TR>
	                        <TR>
                            	<TD>
                            		<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
	                        				<TR>
	                        					<TD CLASS="TD5" NOWRAP>신규일자</TD>
	                        					<TD CLASS="TD6" NOWRAP>
	                        					    <script language =javascript src='./js/h5203ma1_fpDateTime12_txtHdc010t_new_dt.js'></script>
	                        					</TD>
	                        					<TD CLASS="TD5" NOWRAP>만기일자</TD>
	                        					<TD CLASS="TD6" NOWRAP>
	                        					    <script language =javascript src='./js/h5203ma1_fpDateTime12_txtHdc010t_expir_dt.js'></script>
	                        					</TD>
	                        					<TD CLASS="TD5" NOWRAP>해약일자</TD>
	                        					<TD CLASS="TD6" NOWRAP>
	                        					    <script language =javascript src='./js/h5203ma1_fpDateTime12_txtHdc010t_revoke_dt.js'></script>
	                        					</TD>
	                        				</TR>
	                        				<TR>
	                        					<TD CLASS="TD5" NOWRAP>은행코드</TD>
	                        					<TD CLASS="TD6" NOWRAP>
	                        					    <INPUT TYPE=TEXT Name="txtHdc010t_bank_cd" Size="20" ALT="은행코드" Tag="24">
	                        					</TD>
	                        					<TD CLASS="TD5" NOWRAP>은행명</TD>
	                        					<TD CLASS="TD6" NOWRAP>
	                        					    <INPUT TYPE=TEXT Name="txtFaa090t_bank_name" Size="20" ALT="은행명" Tag="24">
	                        					</TD>
	                        					<TD CLASS="TD5" NOWRAP>만기금액</TD>
	                        					<TD CLASS="TD6" NOWRAP>
	                        					    <script language =javascript src='./js/h5203ma1_OBJECT_txtHdc010t_expir_amt.js'></script>
	                        					</TD>
	                        				</TR>
                            		</TABLE>
                            	</TD>
                            </TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=100% FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
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


