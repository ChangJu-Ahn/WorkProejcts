<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 
*  3. Program ID           : h5401ma1
*  4. Program Name         : 월국민연금 조회및 조정 
*  5. Program Desc         : 월국민연금 조회,수정 
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit  

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID      = "h5401mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_ID1     = "h5401mb2.asp"						           '☆: Biz Logic ASP Name
Const C_SHEETMAXROWS    = 22	                                      '☜: Visble row

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

Dim C_HDB020T_EMP_NO
Dim C_HAA010T_NAME
Dim C_HAA010T_DEPT_NM
Dim C_HDB020T_GRADE
Dim C_HDB020T_PRSN_INSUR_AMT
Dim C_HDB020T_COMP_INSUR_AMT
Dim C_HDB020T_ANUT_ACCUM_AMT
Dim C_HDF020T_ANUT_NO

Dim C_HDB020T_EMP_NO2
Dim C_HAA010T_NAME2
Dim C_HAA010T_DEPT_NM2
Dim C_HDB020T_GRADE2
Dim C_HDB020T_PRSN_INSUR_AMT2
Dim C_HDB020T_COMP_INSUR_AMT2
Dim C_HDB020T_ANUT_ACCUM_AMT2
Dim C_HDF020T_ANUT_NO2

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  

    C_HDB020T_EMP_NO = 1                                            'Column Dimant for Spread Sheet
    C_HAA010T_NAME = 2                                                        
    C_HAA010T_DEPT_NM = 3
    C_HDB020T_GRADE = 4
    C_HDB020T_PRSN_INSUR_AMT = 5
    C_HDB020T_COMP_INSUR_AMT  = 6   
    C_HDB020T_ANUT_ACCUM_AMT  = 7
    C_HDF020T_ANUT_NO = 8
    
    C_HDB020T_EMP_NO2 = 1                                            'Column Dimant for Spread Sheet
    C_HAA010T_NAME2 = 2                                                        
    C_HAA010T_DEPT_NM2 = 3
    C_HDB020T_GRADE2 = 4
    C_HDB020T_PRSN_INSUR_AMT2 = 5
    C_HDB020T_COMP_INSUR_AMT2  = 6   
    C_HDB020T_ANUT_ACCUM_AMT2  = 7
    C_HDF020T_ANUT_NO2 = 8

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
	lsInternal_cd     = ""
		
	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay

	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtPay_yymm.Focus
		
	frm1.txtPay_yymm.Year = strYear 		'년월 default value setting
	frm1.txtPay_yymm.Month = strMonth 
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H", "NOCOOKIE", "MA") %>
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
   Dim strYear
   Dim strMonth
   Dim strPayDt

    strYear = frm1.txtPay_yymm.year
    strMonth = frm1.txtPay_yymm.month
    
    If len(strMonth) = 1 then
		strMonth = "0" & strMonth
	End if

	strPayDt = strYear & strMonth

    lgKeyStream       = strPayDt & Parent.gColSep									'You Must append one character(Parent.gColSep)
    lgKeyStream       = lgKeyStream & Trim(frm1.txtGrade.value) & Parent.gColSep
    lgKeyStream       = lgKeyStream & Trim(frm1.txtEmp_no.value) & Parent.gColSep
    lgKeyStream       = lgKeyStream & Trim(frm1.txtInternal_cd.value) & Parent.gColSep
    lgKeyStream       = lgKeyStream & lgUsrIntCd & Parent.gColSep
End Sub        
	
'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
    ggoSpread.Source = frm1.vspdData2
    intIndex = ggoSpread.InsertRow
    
    With frm1.vspdData2        
        .Col = 0
        .Text = "합계"

        .Col   = C_HDB020T_PRSN_INSUR_AMT2
        .value = FncSumSheet(frm1.vspdData,C_HDB020T_PRSN_INSUR_AMT, 1, frm1.vspdData.MaxRows, false, -1, -1, "V")
        
        .Col   = C_HDB020T_COMP_INSUR_AMT2
        .value = FncSumSheet(frm1.vspdData,C_HDB020T_COMP_INSUR_AMT, 1, frm1.vspdData.MaxRows, false, -1, -1, "V")

        .Col   = C_HDB020T_ANUT_ACCUM_AMT2
        .value = FncSumSheet(frm1.vspdData,C_HDB020T_ANUT_ACCUM_AMT, 1, frm1.vspdData.MaxRows, false, -1, -1, "V")

        ggoSpread.Source = frm1.vspdData2

        .ReDraw = False
            ggoSpread.SpreadLock    C_HDB020T_EMP_NO2,  -1, C_HDB020T_EMP_NO2,  -1
            ggoSpread.SpreadLock    C_HAA010T_NAME2,  -1, C_HAA010T_NAME2,  -1
            ggoSpread.SpreadLock    C_HAA010T_DEPT_NM2,  -1, C_HAA010T_DEPT_NM2,  -1
            ggoSpread.SpreadLock    C_HDB020T_GRADE2,  -1, C_HDB020T_GRADE2,  -1
            ggoSpread.SpreadLock    C_HDB020T_PRSN_INSUR_AMT2,  -1, C_HDB020T_PRSN_INSUR_AMT2,  -1
            ggoSpread.SpreadLock    C_HDB020T_COMP_INSUR_AMT2,  -1, C_HDB020T_COMP_INSUR_AMT2,  -1
            ggoSpread.SpreadLock    C_HDB020T_ANUT_ACCUM_AMT2,  -1, C_HDB020T_ANUT_ACCUM_AMT2,  -1
            ggoSpread.SpreadLock    C_HDF020T_ANUT_NO2, -1, C_HDF020T_ANUT_NO2,  -1
        .ReDraw = True
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
            ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread    'sbk

    	   .ReDraw = false
    	
           .MaxCols   = C_HDF020T_ANUT_NO + 1                                                 ' ☜:☜: Add 1 to Maxcols
    	   .Col       = .MaxCols                                                              ' ☜:☜: Hide maxcols
           .ColHidden = True                                                                  ' ☜:☜:
           
           .MaxRows = 0
            ggoSpread.ClearSpreadData

           Call GetSpreadColumnPos("A") 'sbk
    	
           ggoSpread.SSSetEdit     C_HDB020T_EMP_NO,              "사번", 15,,,,2
           ggoSpread.SSSetEdit     C_HAA010T_NAME,                "성명", 15,,,,2
           ggoSpread.SSSetEdit     C_HAA010T_DEPT_NM,             "부서", 20,,,,2
           ggoSpread.SSSetEdit     C_HDB020T_GRADE,               "등급", 9,,,4,2
           ggoSpread.SSSetFloat    C_HDB020T_PRSN_INSUR_AMT,      "개인부담료", 15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
           ggoSpread.SSSetFloat    C_HDB020T_COMP_INSUR_AMT,      "회사부담료", 15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
           ggoSpread.SSSetFloat    C_HDB020T_ANUT_ACCUM_AMT,      "퇴직전환금누계", 16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
    	   ggoSpread.SSSetEdit	  C_HDF020T_ANUT_NO,	         "국민연금번호", 20,,,,2

    	   .ReDraw = true
        
        End With
    End If

    If pvSpdNo = "" OR pvSpdNo = "B" Then
    	With frm1.vspdData2

            ggoSpread.Source = frm1.vspdData2
            ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread    'sbk

    	   .ReDraw = false
    	
           .MaxCols   = C_HDF020T_ANUT_NO2 + 1                                                ' ☜:☜: Add 1 to Maxcols
    	   .Col       = .MaxCols                                                              ' ☜:☜: Hide maxcols
           .ColHidden = True                                                                  ' ☜:☜:

           .MaxRows = 0	
            ggoSpread.ClearSpreadData

           .DisplayColHeaders = False

            Call GetSpreadColumnPos("B") 'sbk

            ggoSpread.SSSetEdit     C_HDB020T_EMP_NO2,              "", 15,,,,2
            ggoSpread.SSSetEdit     C_HAA010T_NAME2,                "", 15,,,,2
            ggoSpread.SSSetEdit     C_HAA010T_DEPT_NM2,             "", 20,,,,2
            ggoSpread.SSSetEdit     C_HDB020T_GRADE2,               "", 9,,,,2
            ggoSpread.SSSetFloat    C_HDB020T_PRSN_INSUR_AMT2,      "개인부담료", 15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
            ggoSpread.SSSetFloat    C_HDB020T_COMP_INSUR_AMT2,      "회사부담료", 15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
            ggoSpread.SSSetFloat    C_HDB020T_ANUT_ACCUM_AMT2,      "퇴직전환금누계", 16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
    		ggoSpread.SSSetEdit	   C_HDF020T_ANUT_NO2,	           "", 20,,,,2

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
    With frm1
    .vspdData.ReDraw = False
        ggoSpread.SpreadLock    C_HDB020T_EMP_NO,  -1, C_HDB020T_EMP_NO,  -1
        ggoSpread.SpreadLock    C_HAA010T_NAME,  -1, C_HAA010T_NAME,  -1
        ggoSpread.SpreadLock    C_HAA010T_DEPT_NM,  -1, C_HAA010T_DEPT_NM,  -1
        ggoSpread.SpreadLock    C_HDB020T_GRADE,  -1, C_HDB020T_GRADE,  -1
        ggoSpread.SpreadLock    C_HDF020T_ANUT_NO, -1, C_HDF020T_ANUT_NO,  -1
        ggoSpread.SSSetProtected   .vspdData.MaxCols   , -1, -1
    .vspdData.ReDraw = True
    End With

    ggoSpread.Source = frm1.vspdData2
    With frm1.vspdData2
    .ReDraw = False
        ggoSpread.SpreadLock    C_HDB020T_EMP_NO2,  -1, C_HDB020T_EMP_NO2,  -1
        ggoSpread.SpreadLock    C_HAA010T_NAME2,  -1, C_HAA010T_NAME2,  -1
        ggoSpread.SpreadLock    C_HAA010T_DEPT_NM2,  -1, C_HAA010T_DEPT_NM2,  -1
        ggoSpread.SpreadLock    C_HDB020T_GRADE2,  -1, C_HDB020T_GRADE2,  -1
        ggoSpread.SpreadLock    C_HDB020T_PRSN_INSUR_AMT2,  -1, C_HDB020T_PRSN_INSUR_AMT2,  -1
        ggoSpread.SpreadLock    C_HDB020T_COMP_INSUR_AMT2,  -1, C_HDB020T_COMP_INSUR_AMT2,  -1
        ggoSpread.SpreadLock    C_HDB020T_ANUT_ACCUM_AMT2,  -1, C_HDB020T_ANUT_ACCUM_AMT2,  -1
        ggoSpread.SpreadLock    C_HDF020T_ANUT_NO2, -1, C_HDF020T_ANUT_NO2,  -1
        ggoSpread.SSSetProtected   .MaxCols   , -1, -1
    .ReDraw = True

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

            C_HDB020T_EMP_NO  = iCurColumnPos(1)                   
            C_HAA010T_NAME    = iCurColumnPos(2)                   
            C_HAA010T_DEPT_NM = iCurColumnPos(3)                   
            C_HDB020T_GRADE   = iCurColumnPos(4)                   
            C_HDB020T_PRSN_INSUR_AMT = iCurColumnPos(5)            
            C_HDB020T_COMP_INSUR_AMT = iCurColumnPos(6)            
            C_HDB020T_ANUT_ACCUM_AMT = iCurColumnPos(7)            
            C_HDF020T_ANUT_NO        = iCurColumnPos(8)            
            
        Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
                                                                   
            C_HDB020T_EMP_NO2  = iCurColumnPos(1)
            C_HAA010T_NAME2    = iCurColumnPos(2)
            C_HAA010T_DEPT_NM2 = iCurColumnPos(3)
            C_HDB020T_GRADE2   = iCurColumnPos(4)
            C_HDB020T_PRSN_INSUR_AMT2 = iCurColumnPos(5)
            C_HDB020T_COMP_INSUR_AMT2 = iCurColumnPos(6)
            C_HDB020T_ANUT_ACCUM_AMT2 = iCurColumnPos(7)
            C_HDF020T_ANUT_NO2        = iCurColumnPos(8)
            
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
    Call ggoOper.FormatDate(frm1.txtPay_yymm, Parent.gDateFormat, 2)	'년월 
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
            
    Call InitSpreadSheet("")                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    
    Call FuncGetAuth(gStrRequestMenuID, Parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
    Call SetDefaultVal
	Call SetToolbar("1100100000001111")												'⊙: Set ToolBar
    frm1.txtPay_yymm.focus() 
    
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
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    
	
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field

    Call InitVariables															'⊙: Initializes local global variables
    
    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If

    If  txtDept_cd_Onchange()  then
        Exit Function
    End If

    If  txtEmp_no_Onchange()  then
        Exit Function
    End If

    Call MakeKeyStream("X")

    Call DisableToolBar(Parent.TBC_QUERY)
	If DBQuery=False Then
	   Call RestoreToolBar()

	   Exit Function
	End If

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
       IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call ggoOper.ClearField(Document, "A")                                       '☜: Clear Condition Field
    Call ggoOper.LockField(Document , "N")                                       '☜: Lock  Field
    
	Call SetToolbar("1100100000001111")												'⊙: Set ToolBar
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
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                            'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                  '☜: Please do Display first. 
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")		                  '☜: Do you want to delete? 
	If IntRetCD = vbNo Then											        
		Exit Function	
	End If
    
    Call DisableToolBar(Parent.TBC_DELETE)
	If DBDelete=False Then
	   Call RestoreToolBar()
	   Exit Function
	End If
    
    FncDelete=  True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    Dim iRow
    Dim dblid_accum
    
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
	
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
	
	ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
       Exit Function
    End If

    If frm1.txtPay_yymm.value="" Then
        Call DisplayMsgBox("970021","X",frm1.txtPay_yymm.alt,"X")        '조회년월은 필수 입력사항입니다.
        frm1.txtPay_yymm.focus    ' go to 
        Exit Function
    End If                    

    With Frm1.vspdData
        For iRow = 1 To  .MaxRows
            .Row = iRow
            .Col = 0
           Select Case .Text
               Case ggoSpread.UpdateFlag
   	                .Col = C_HDB020T_ANUT_ACCUM_AMT
                    dblid_accum  = .Value
		            IF IsNull(dblid_accum) OR dblid_accum < 0 THEN
	                    Call DisplayMsgBox("800280","X","X","X")	'퇴직전환금누계는 0보다 크거나 같아야합니다.
	                    .Row = iRow
  	                    .Col = C_HDB020T_ANUT_ACCUM_AMT
  	                    .Action=0
                        Set gActiveElement = document.activeElement
                        Exit Function
                   End If
           End Select
        Next
    End With
    
    Call MakeKeyStream("X")
    Call DisableToolBar(Parent.TBC_SAVE)
	If DBSave=False Then
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
    
	With Frm1
	    
		If .vspdData.ActiveRow > 0 Then
			.vspdData.ReDraw = False
		
			ggoSpread.Source = frm1.vspdData	
			ggoSpread.CopyRow
			SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow
   
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
' Name : FncPrev
' Desc : developer describe this line Called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 

    Dim strVal
    Dim IntRetCD

    FncPrev = False                                                              '☜: Processing is OK
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
    Call InitVariables														 '⊙: Initializes local global variables

	If LayerShowHide(1) = False Then
		Exit Function
	End If

    Call MakeKeyStream("X")

    strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001                     '☜: Query
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
    Call InitVariables														 '⊙: Initializes local global variables

	If LayerShowHide(1) = False Then
		Exit Function
	End If

    Call MakeKeyStream("X")

    strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001                     '☜: Query
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
	Call Parent.FncExport(Parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind(Parent.C_SINGLE, True)
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
    ggoSpread.Source = frm1.vspdData
	Call ggoSpread.ReOrderingSpreadData()

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet("B")      
    ggoSpread.Source = frm1.vspdData2
	Call ggoSpread.ReOrderingSpreadData()

    Frm1.vspdData2.Col = 0
    Frm1.vspdData2.Text = "합계"

End Sub

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()

	Dim IntRetCD
	FncExit = False
	ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")			 '⊙: Data is changed.  Do you want to exit? 
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

	If LayerShowHide(1) = False Then
		Exit Function
	End If

    Call LayerShowHide(1)

    strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001                     '☜: Query
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
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
	Dim strVal, strDel
    Err.Clear                                                                    '☜: Clear err status
		
	DbSave = False														         '☜: Processing is NG
		
	If LayerShowHide(1) = False Then
		Exit Function
	End If

	With frm1
		.txtMode.value        = Parent.UID_M0002                                        '☜: Delete
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
               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                                           strVal = strVal & "U" & Parent.gColSep
                                                                           strVal = strVal & lRow & Parent.gColSep
                                                                           strVal = strVal & Trim(.txtPay_yymm.text) & Parent.gColSep
                    .vspdData.Col = C_HDB020T_EMP_NO                     : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_HDB020T_GRADE                      : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_HDB020T_PRSN_INSUR_AMT             : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_HDB020T_COMP_INSUR_AMT             : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_HDB020T_ANUT_ACCUM_AMT             : strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
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
		
	If LayerShowHide(1) = False Then
		Exit Function
	End If

		
	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003                                '☜: Delete
	strVal = strVal & "&txtGlNo=" & Trim(frm1.txtLcNo.value)             '☜: 
		
	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
	
	DbDelete = True                                                              '⊙: Processing is NG
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
    Dim strVal
	lgIntFlgMode      = Parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
	
    ggoSpread.Source       = Frm1.vspdData2
    Frm1.vspdData2.MaxRows = 0
    ggoSpread.Source       = Frm1.vspdData2
    ggoSpread.ClearSpreadData

    Call MakeKeyStream("X")
	If LayerShowHide(1) = False Then
		Exit Function
	End If


    strVal = BIZ_PGM_ID1 & "?txtMode="            & Parent.UID_M0001                    '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & 1                             '☜: Max fetched data
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    Call ggoOper.LockField(Document, "Q")
    frm1.vspdData.focus
End Function
'========================================================================================================
' Function Name : DbQueryOk1
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk1()

	lgIntFlgMode      = Parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
	
    Frm1.vspdData2.Col = 0
    Frm1.vspdData2.Text = "합계"
    Frm1.txtPay_yymm.focus 
	Call SetToolbar("1100100100011111")												'⊙: Set ToolBar

	Frm1.vspdData.focus
    Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   
End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
	Call InitVariables
    Frm1.vspdData.MaxRows = 0
    ggoSpread.Source = Frm1.vspdData2
    ggoSpread.ClearSpreadData
    ggoSpread.Source = Frm1.vspdData
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
' Name : OpenDept
' Desc : 부서 POPUP
'========================================================================================================
Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = frm1.txtDept_cd.value			' 조건부에서 누른 경우 Code Condition
	Else 'spread
		arrParam(0) = frm1.vspdData.Text			' Grid에서 누른 경우 Code Condition
	End If
	arrParam(1) = ""								' 날짜기준(""현재날짜)
	arrParam(2) = lgUsrIntCd        			    ' Internal_cd

	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent, arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = 0 Then
			frm1.txtDept_cd.focus
		Else 
			frm1.vspdData.Col = C_Dept
			frm1.vspdData.action =0 
		End If	
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
		If iWhere = 0 Then
			.txtDept_cd.value = arrRet(0)
			.txtDept_Nm.value = arrRet(1)
			.txtDept_cd.focus
		Else
			.vspdData.Col = C_DeptNm
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_Dept
			.vspdData.Text = arrRet(0)
			.vspdData.action =0 
		End If
	End With
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
	arrParam(1) = ""  'frm1.txtName.value			' Name Cindition
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
		Call ggoOper.ClearField(Document, "2")					 '☜: Clear Contents  Field
		Set gActiveElement = document.ActiveElement
		.txtEmp_no.focus
		lgBlnFlgChgValue = False
	End With
End Sub

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col
End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)

    Dim iDx
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

   	If Frm1.vspdData.CellType = Parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

   	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    
    Call SetPopupMenuItemInf("0001101111")

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

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
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
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    if  IntRetCd < 0 then
	        if  IntRetCd = -1 then
    			Call DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            end if
            Call ggoOper.ClearField(Document, "2")
            call InitVariables()
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true
        Else
            frm1.txtName.value = strName
        End if 
    End if
    
End Function
'======================================================================================================
'   Event Name : txtDept_cd_OnChange
'   Event Desc : 부서가 변경될 경우 
'=======================================================================================================
Function txtDept_cd_OnChange()

    Dim IntRetCd
    Dim strDept_nm

    frm1.txtDept_nm.value = ""

    If  frm1.txtDept_cd.value = "" Then
        frm1.txtDept_nm.value = ""
        frm1.txtInternal_cd.value = ""
        lsInternal_cd = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtDept_cd.value,"",lgUsrIntCd,strDept_nm,lsInternal_cd)
        If  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call DisplayMsgBox("800012", "x","x","x")   ' 등록되지 않은 부서코드입니다.
            else
                Call DisplayMsgBox("800455", "x","x","x")   ' 자료권한이 없습니다.
            end if
            frm1.txtInternal_cd.value = ""
            lsInternal_cd = ""
            frm1.txtDept_cd.focus
            Set gActiveElement = document.ActiveElement
            txtDept_cd_OnChange = true
        Else
            frm1.txtDept_nm.value = strDept_nm
            frm1.txtInternal_cd.value = lsInternal_cd
        End if
    End if
			
End Function
'=======================================================================================================
'   Event Name : txtYear_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtPay_yymm_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtPay_yymm.Action = 7
        frm1.txtPay_yymm.focus
    End If
End Sub
'=======================================================================================================
'   Event Name : txtPay_yymm_Keypress(Key)
'   Event Desc : 3rd party control에서 Enter 키를 누르면 조회 실행 
'=======================================================================================================
Sub txtPay_yymm_Keypress(Key)
    If Key = 13 Then
        Call MainQuery()
    End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>월국민연금조회및조정</font></td>
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
    	            <TD HEIGHT=20 WIDTH=90%>
						<FIELDSET CLASS="CLSFLD">
                           <TABLE <%=LR_SPACE_TYPE_40%>>
						    <TR>
							    <TD CLASS=TD5 NOWRAP>조회년월</TD>
			                    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h5401ma1_txtPay_yymm_txtPay_yymm.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>등 급</TD>       
					            <TD CLASS=TD6 NOWRAP><INPUT ID=txtGrade NAME="txtGrade" MAXLENGTH="4" SIZE=10 ALT ="등급" tag="11XXXU"></TD>
							</TR>
							<TR>
			    	    		<TD CLASS="TD5" NOWRAP>사원</TD>
			    	    		<TD CLASS="TD6" NOWRAP><INPUT ID=txtEmp_no NAME="txtEmp_no"  SIZE=13 MAXLENGTH=13 ALT="사번" TYPE="Text" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmp()">
			    	        			    	       <INPUT ID=txtName NAME="txtName"  SIZE=20  MAXLENGTH=30 ALT="성명" TYPE="Text"  tag="14XXXU"></TD>
							    <TD CLASS=TD5 NOWRAP>부서코드</TD>
								<TD CLASS=TD6 NOWRAP><INPUT ID=txtDept_cd NAME="txtDept_cd" SIZE=10  MAXLENGTH=10  ALT ="부서코드" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenDept(0)">
								                     <INPUT ID=txtDept_nm NAME="txtDept_nm" SIZE=20  MAXLENGTH=40  ALT ="부서코드명"   tag="14XXXU">
								                     <INPUT ID=txtInternal_cd NAME="txtInternal_cd" TYPE="HIDDEN"  SIZE=20 ALT ="부서코드명"   tag="14XXXU"></TD>
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
									<script language =javascript src='./js/h5401ma1_vaSpread_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=42 VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT=100%>
									<script language =javascript src='./js/h5401ma1_vaSpread1_vspdData2.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
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

