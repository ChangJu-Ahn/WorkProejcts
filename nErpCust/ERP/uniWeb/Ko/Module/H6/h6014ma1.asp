<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          	: Human Resources
*  2. Function Name        	: Multi Sample
*  3. Program ID           	: H6014ma1
*  4. Program Name         	: H6014ma1
*  5. Program Desc         	: 급여변동사원조회 
*  6. Comproxy List        	:
*  7. Modified date(First) 	: 2001/05/30
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
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"> </SCRIPT>

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID      = "h6014mb1.asp"						           '☆: Biz Logic ASP Name
Const C_SHEETMAXROWS    = 21	                                      '☜: Visble row

Const TAB1 = 1
Const TAB2 = 2

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

Dim C_EMP_NO
Dim C_NAME
Dim C_DEPT_CD
Dim C_PAY_GRD1
Dim C_ROLL_PSTN
Dim C_MINOR_NM
Dim C_BEFORE_AMT
Dim C_CURRENT_AMT
Dim C_DIFF_AMT

Dim C_EMP_NO2
Dim C_NAME2
Dim C_DEPT_CD2
Dim C_PAY_GRD12
Dim C_ROLL_PSTN2
Dim C_MINOR_NM2
Dim C_BEFORE_AMT2
Dim C_CURRENT_AMT2
Dim C_DIFF_AMT2

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables(ByVal pvSpdNo)  

    If pvSpdNo = "A" Then

        C_EMP_NO = 1
        C_NAME = 2
        C_DEPT_CD = 3
        C_PAY_GRD1 = 4
        C_ROLL_PSTN = 5
        C_MINOR_NM = 6
        C_BEFORE_AMT = 7
        C_CURRENT_AMT = 8
        C_DIFF_AMT = 9
    
    ElseIf pvSpdNo = "B" Then
        C_EMP_NO2 = 1
        C_NAME2 = 2
        C_DEPT_CD2 = 3
        C_PAY_GRD12 = 4
        C_ROLL_PSTN2 = 5
        C_MINOR_NM2 = 6
        C_BEFORE_AMT2 = 7
        C_CURRENT_AMT2 = 8
        C_DIFF_AMT2 = 9
    End If

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
	frm1.txtFrom_dt.focus 		
	frm1.txtFrom_dt.Year = strYear 		 '년월일 default value setting
	frm1.txtFrom_dt.Month = strMonth 
	
	frm1.txtTo_dt.Year = strYear 		 '년월일 default value setting
	frm1.txtTo_dt.Month = strMonth 
	
	frm1.txtFrom_dt2.Year = strYear 		 '년월일 default value setting
	frm1.txtFrom_dt2.Month = strMonth 
	
	frm1.txtTo_dt2.Year = strYear 		 '년월일 default value setting
	frm1.txtTo_dt2.Month = strMonth 
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
    Dim rbo_sort  
    Dim rbo_sort2
    Dim StrDt,StrDt2
    StrDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtTo_dt.Year, Right("0" & frm1.txtTo_dt.month , 2), "01")
    StrDt2 = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtTo_dt2.Year, Right("0" & frm1.txtTo_dt2.month , 2), "01")

    If gSelframeFlg = TAB1 Then
        If frm1.rbo_sort(0).checked Then
            rbo_sort="1"
        ELSEIF frm1.rbo_sort(1).checked THEN
                rbo_sort="2"
        End If
   
        lgKeyStream  = gSelframeFlg & Parent.gColSep
        lgKeyStream  = lgKeyStream & Frm1.txtFrom_dt.Year & Right("0" & Frm1.txtFrom_dt.Month,2) & Parent.gColSep   '1
        lgKeyStream  = lgKeyStream & Frm1.txtTo_dt.Year & Right("0" & Frm1.txtTo_dt.Month,2) & Parent.gColSep     '2
        lgKeyStream  = lgKeyStream & Trim(frm1.txtOcpt_type.value) & Parent.gColSep   '직종             '3    
        lgKeyStream  = lgKeyStream & Trim(frm1.txtSect_cd.value) & Parent.gColSep     '근무구역         '4
        lgKeyStream  = lgKeyStream & Trim(Frm1.txtfr_internal_cd.Value) & Parent.gColSep                '5
        lgKeyStream  = lgKeyStream & Trim(Frm1.txtto_internal_cd.Value) & Parent.gColSep                '6
        lgKeyStream  = lgKeyStream & rbo_sort & Parent.gColSep                        '조회구분         '7     
        lgKeyStream  = lgKeyStream & Trim(frm1.txtPayCd.value) & Parent.gColSep       '지급구분         '8
        lgKeyStream  = lgKeyStream & Trim(frm1.txtAllow_cd.value) & Parent.gColSep    '수당코드         '9
        lgKeyStream  = lgKeyStream & lgUsrIntcd & Parent.gColSep                                        '10
        lgKeyStream  = lgKeyStream & StrDt & Parent.gColSep
    Else    
        If frm1.rbo_sort2(0).checked Then
            rbo_sort2="1"
        ELSEIF frm1.rbo_sort2(1).checked THEN
                rbo_sort2="2"
        End If

        lgKeyStream  = gSelframeFlg & Parent.gColSep
        lgKeyStream  = lgKeyStream & Frm1.txtFrom_dt.Year & Right("0" & Frm1.txtFrom_dt2.Month,2) & Parent.gColSep 
        lgKeyStream  = lgKeyStream & Frm1.txtTo_dt.Year & Right("0" & Frm1.txtTo_dt2.Month,2) & Parent.gColSep 
        lgKeyStream  = lgKeyStream & Trim(frm1.txtOcpt_type2.value) & Parent.gColSep  '직종   
        lgKeyStream  = lgKeyStream & Trim(frm1.txtSect_cd2.value) & Parent.gColSep    '근무구역  
        lgKeyStream  = lgKeyStream & Trim(Frm1.txtfr_internal_cd2.Value) & Parent.gColSep  
        lgKeyStream  = lgKeyStream & Trim(Frm1.txtto_internal_cd2.Value) & Parent.gColSep  
        lgKeyStream  = lgKeyStream & rbo_sort2 & Parent.gColSep                       '조회구분  
        lgKeyStream  = lgKeyStream & Trim(frm1.txtsub_type2.value) & Parent.gColSep    '공제구분   
        lgKeyStream  = lgKeyStream & Trim(frm1.txtsub_cd2.value) & Parent.gColSep     '공제코드   
        lgKeyStream  = lgKeyStream & lgUsrIntcd & Parent.gColSep
        lgKeyStream  = lgKeyStream & StrDt2 & Parent.gColSep    
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
    
    ' 직종 
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0003", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr =  lgF0
    iNameArr =  lgF1
    Call SetCombo2(frm1.txtOcpt_type, iCodeArr, iNameArr, Chr(11))
    Call SetCombo2(frm1.txtOcpt_type2, iCodeArr, iNameArr, Chr(11))
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
End Sub

'========================================================================================================
'	name: Tab Click
'	desc: Tab Click시 필요한 기능을 수행한다.
'========================================================================================================
Function ClickTab1()
	If gSelframeFlg = TAB1 Then Exit Function
	
	Call changeTabs(TAB1)
	gSelframeFlg = TAB1
End Function
'-------------------------------------------------------------------------------------------
Function ClickTab2()
	If gSelframeFlg = TAB2 Then Exit Function
	
	Call changeTabs(TAB2)
	gSelframeFlg = TAB2
End Function

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)

    If pvSpdNo = "" OR pvSpdNo = "A" Then

    	Call initSpreadPosVariables("A")   'sbk 

	    With frm1.vspdData
	
            ggoSpread.Source = frm1.vspdData

            ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    'sbk
                                                            
	       .ReDraw = false

           .MaxCols   = C_DIFF_AMT + 1                                                     ' ☜:☜: Add 1 to Maxcols
	       .Col       = .MaxCols                                                              ' ☜:☜: Hide maxcols
           .ColHidden = True                                                                  ' ☜:☜:
           
           .MaxRows = 0
            ggoSpread.ClearSpreadData

            Call GetSpreadColumnPos("A") 'sbk
	
            Call AppendNumberPlace("6","2","0")
                
            ggoSpread.SSSetEdit   C_EMP_NO,       "사번"       , 10,,, 13,2
            ggoSpread.SSSetEdit   C_NAME,         "성명"       , 14,,, 30,2
            ggoSpread.SSSetEdit   C_DEPT_CD,      "부서"       , 17,,, 20,2
            ggoSpread.SSSetEdit   C_PAY_GRD1,     "직급"       , 12,,, 50,2
            ggoSpread.SSSetEdit   C_ROLL_PSTN,    "직위"       , 10,,, 50,2
            ggoSpread.SSSetEdit   C_MINOR_NM,     "직책"       , 10,,, 50,2
            ggoSpread.SSSetFloat  C_BEFORE_AMT,   "기준년월금액"   , 14,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat  C_CURRENT_AMT,  "비교년월금액"   , 14,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat  C_DIFF_AMT,     "차액"       , 14,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec

	       .ReDraw = true
	
            Call SetSpreadLock("A")
    
        End With
    End If
    
    If pvSpdNo = "" OR pvSpdNo = "B" Then

    	Call initSpreadPosVariables("B")   'sbk 

        With frm1.vspdData2
	
            ggoSpread.Source = frm1.vspdData2

            ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    'sbk
                                                            
	       .ReDraw = false

           .MaxCols   = C_DIFF_AMT2 + 1                                                     ' ☜:☜: Add 1 to Maxcols
	       .Col       = .MaxCols                                                              ' ☜:☜: Hide maxcols
           .ColHidden = True                                                                  ' ☜:☜:
           
           .MaxRows = 0
            ggoSpread.ClearSpreadData

            Call GetSpreadColumnPos("B") 'sbk
	
            Call AppendNumberPlace("6","2","0")

            ggoSpread.SSSetEdit   C_EMP_NO2,       "사번"       , 10,,, 13,2
            ggoSpread.SSSetEdit   C_NAME2,         "성명"       , 14,,, 30,2
            ggoSpread.SSSetEdit   C_DEPT_CD2,      "부서"       , 17,,, 20,2
            ggoSpread.SSSetEdit   C_PAY_GRD12,     "직급"       , 12,,, 50,2
            ggoSpread.SSSetEdit   C_ROLL_PSTN2,    "직위"       , 10,,, 50,2
            ggoSpread.SSSetEdit   C_MINOR_NM2,     "직책"       , 10,,, 50,2
            ggoSpread.SSSetFloat  C_BEFORE_AMT2,   "기준년월금액"   , 14,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat  C_CURRENT_AMT2,  "비교년월금액"   , 14,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat  C_DIFF_AMT2,     "차액"       , 14,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec

	       .ReDraw = true
	
            Call SetSpreadLock("B")
    
        End With
    End If
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)

    If pvSpdNo = "A" Then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()
    ElseIf pvSpdNo = "B" Then
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.SpreadLockWithOddEvenRowColor()
    End If
    
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
    With frm1
    .vspdData2.ReDraw = False
    .vspdData2.ReDraw = True
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
            
            C_EMP_NO = iCurColumnPos(1) 
            C_NAME = iCurColumnPos(2) 
            C_DEPT_CD = iCurColumnPos(3) 
            C_PAY_GRD1 = iCurColumnPos(4) 
            C_ROLL_PSTN = iCurColumnPos(5) 
            C_MINOR_NM = iCurColumnPos(6) 
            C_BEFORE_AMT = iCurColumnPos(7) 
            C_CURRENT_AMT = iCurColumnPos(8) 
            C_DIFF_AMT = iCurColumnPos(9) 
        
       Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

            C_EMP_NO2 = iCurColumnPos(1) 
            C_NAME2 = iCurColumnPos(2) 
            C_DEPT_CD2 = iCurColumnPos(3) 
            C_PAY_GRD12 = iCurColumnPos(4) 
            C_ROLL_PSTN2 = iCurColumnPos(5) 
            C_MINOR_NM2 = iCurColumnPos(6) 
            C_BEFORE_AMT2 = iCurColumnPos(7) 
            C_CURRENT_AMT2 = iCurColumnPos(8) 
            C_DIFF_AMT2 = iCurColumnPos(9) 
            
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
    Call InitVariables 
    
    Call ggoOper.FormatDate(frm1.txtFrom_dt , Parent.gDateFormat, 2)            '싱글에서 년월말 입력하고 싶은경우 다음 함수를 콜한다.
    Call ggoOper.FormatDate(frm1.txtTo_dt   , Parent.gDateFormat, 2)            
    Call ggoOper.FormatDate(frm1.txtFrom_dt2, Parent.gDateFormat, 2)         
    Call ggoOper.FormatDate(frm1.txtTo_dt2  , Parent.gDateFormat, 2)           
    
    Call FuncGetAuth("H6014MA1", Parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
    
    Call SetDefaultVal
    
	gSelframeFlg = TAB1
	Call changeTabs(TAB1)
	Call SetToolbar("1100000000001111")												'⊙: Set ToolBar
    gIsTab     = "Y" ' <- "Yes"의 약자 Y(와이) 입니다.[V(브이)아닙니다]
    gTabMaxCnt = 2   ' Tab의 갯수를 적어 주세요    
    
   	Call InitComboBox() 
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
    Dim iDx
    Dim strWhere
    Dim StrDt,StrDt2
    StrDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtTo_dt.Year, Right("0" & frm1.txtTo_dt.month , 2), "01")
    StrDt2 = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtTo_dt2.Year, Right("0" & frm1.txtTo_dt2.month , 2), "01")

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
   
   SELECT CASE gSelframeFlg                                                       '☜:TAB 에따른 필수입력사항체크 
   
   CASE TAB1
         IF Trim(frm1.txtallow_cd.value) = "" THEN
             CALL DisplayMsgBox("970021", "X","수당코드","X")
             frm1.txtallow_cd.focus()
             EXIT FUNCTION
         END IF    
         IF Trim(frm1.txtPayCd.value) = "" THEN
             CALL DisplayMsgBox("970021", "X","지급구분코드","X")
             frm1.txtPayCd.focus()
             EXIT FUNCTION
         END IF
                
        If  txtallow_cd_Onchange() = false then
            Exit Function
        End If

        If  txtPayCd_Onchange() = false then
            Exit Function
        End If

        If  txtSect_cd_Onchange() = false then
            Exit Function
        End If
    
        If  txtFr_Dept_cd_Onchange() = false then
            Exit Function
        End If

        If  txtTo_Dept_cd_Onchange() = false then
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
                frm1.txtFr_dept_cd.focus()
                Set gActiveElement = document.activeElement
                Exit Function
            End IF 
            
        END IF   
    
        
  CASE TAB2
         IF frm1.txtsub_cd2.value = "" THEN
             CALL DisplayMsgBox("970021", "X","공제코드","X")
             frm1.txtsub_cd2.focus
             EXIT FUNCTION
         END IF    
         IF frm1.txtsub_type2.value = "" THEN
             CALL DisplayMsgBox("970021", "X","공제구분","X")
             frm1.txtsub_type2.focus
             EXIT FUNCTION
         END IF

        If  txtsub_cd2_Onchange() = false then
            Exit Function
        End If

        If  txtsub_type2_Onchange() = false then
            Exit Function
        End If

        If  txtSect_cd2_Onchange() = false then
            Exit Function
        End If
    
        If  txtFr_Dept_cd2_Onchange() = false then
            Exit Function
        End If

        If  txtTo_Dept_cd2_Onchange() = false then
            Exit Function
        End If
 
        Dim Fr_dept_cd2 , To_dept_cd2
    
        Fr_dept_cd2 = frm1.txtFr_internal_cd2.value
        To_dept_cd2 = frm1.txtTo_internal_cd2.value
    
        If fr_dept_cd2 = "" then
            IntRetCd = FuncGetTermDept(lgUsrIntCd ,StrDt2, rFrDept ,rToDept)
	    	frm1.txtFr_internal_cd2.value = rFrDept
	    	frm1.txtFr_dept_nm2.value = ""
	    End If	
	
	    If to_dept_cd2 = "" then
            IntRetCd = FuncGetTermDept(lgUsrIntCd ,StrDt2, rFrDept ,rToDept)
	    	frm1.txtTo_internal_cd2.value = rToDept
	    	frm1.txtTo_dept_nm2.value = ""
	    End If  
        Fr_dept_cd2 = frm1.txtFr_internal_cd2.value
        To_dept_cd2 = frm1.txtTo_internal_cd2.value

        If (Fr_dept_cd2<> "") AND (To_dept_cd2<>"") Then       
    
            If Fr_dept_cd2 > To_dept_cd2 then
	            Call DisplayMsgBox("800153","X","X","X")	'시작부서코드는 종료부서코드보다 작아야합니다.
                frm1.txtFr_dept_cd2.value = ""
                frm1.txtFr_dept_nm2.value = ""
                frm1.txtFr_internal_cd2.value = ""
                
                frm1.txtTo_dept_cd2.value = ""
                frm1.txtTo_dept_nm2.value = ""
                frm1.txtTo_internal_cd2.value = ""
                frm1.txtFr_dept_cd2.focus
                Set gActiveElement = document.activeElement
                Exit Function
            End IF 
        END IF   
  
	END SELECT       
    Call InitVariables															'⊙: Initializes local global variables
    If (frm1.txtFrom_dt.Text = "") Then
        frm1.txtFrom_dt.Text = "" 
    End If
    If (frm1.txtTo_dt.Text = "") Then
        frm1.txtTo_dt.Text = "" 
    End If
    
    Call MakeKeyStream("X")
    
    Call DisableToolBar(Parent.TBC_QUERY)
	IF DBQUERY =  False Then
		Call RestoreToolBar()
		Exit Function
	End If  
    FncQuery = True																'☜: Processing is OK
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
    Err.Clear   
                                                                     '☜: Clear err status
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
	
	
	If gSelframeFlg = TAB2 Then
    	frm1.vspdData.ReDraw = False
    	
        ggoSpread.Source = frm1.vspdData	
        ggoSpread.CopyRow
        SetSpreadColor frm1.vspdData.ActiveRow
        
    	frm1.vspdData.ReDraw = True
    ElseIf gSelframeFlg = TAB1 Then 
        Call ggoOper.ClearField(Document, "1")                                  <%'Clear Condition Field%>
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
End Function
'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow() 
	If gSelframeFlg <> TAB2 Then
		Call ClickTab2		'sstData.Tab = 1
	End If
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
    Call parent.FncFind(Parent.C_MULTI, False)                                          '☜:화면 유형, Tab 유무 
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
    
    Select Case gActiveSpdSheet.id
		Case "vaSpread"
			Call InitSpreadSheet("A")
		Case "vaSpread1"
			Call InitSpreadSheet("B")      		
	End Select 
	
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

    DbQuery = False
    
    Err.Clear                                                                        '☜: Clear err status

	Call LayerShowHide(1)
	
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
		
	Call LayerShowHide(1)
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

	lgIntFlgMode      = Parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
	Call SetToolbar("1100000000011111")												'⊙: Set ToolBar
    Call InitData()
    Call ggoOper.LockField(Document, "Q")
    If gSelframeFlg = TAB1 Then    
		frm1.vspdData.focus
	else
		frm1.vspdData2.focus
	end if
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

    Call MAINQuery()
	Call ClickTab1()
End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
	Call FncNew()	
End Function
'========================================================================================================
' Name : OpenCondAreaPopup
' Desc : developer describe this line 
'========================================================================================================
Function OpenCondAreaPopup(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
    Dim strFrom_dt
    Dim strTo_dt

	If IsOpenPop = True  Then  
	   Exit Function
	End If   

	IsOpenPop = True
	Select Case iWhere
	    
	   Case "5" 	
	    	arrParam(0) = "수당코드 팝업"		    ' 팝업 명칭 
	        arrParam(1) = "HDA010T"				 	    ' TABLE 명칭 
	        arrParam(2) = frm1.txtAllow_cd.value	    ' Code Condition
	        arrParam(3) = ""'frm1.txtAllow_nm.value	    ' Name Cindition
	        arrParam(4) = " PAY_CD = " & FilterVar("*", "''", "S") & "  AND CODE_TYPE = " & FilterVar("1", "''", "S") & "  "  ' Where Condition
	        arrParam(5) = "수당코드"			
	    
            arrField(0) = "allow_cd"				    ' Field명(0)
            arrField(1) = "allow_nm"				    ' Field명(1)
    
            arrHeader(0) = "수당코드"	            ' Header명(0)
            arrHeader(1) = "수당코드명"			    ' Header명(1)
            
	   Case "6"
            arrParam(0) = "공제코드 팝업"			' 팝업 명칭 
	        arrParam(1) = "HDA010T"				 		' TABLE 명칭 
	        arrParam(2) = frm1.txtsub_cd2.value		    ' Code Condition
	        arrParam(3) = ""'frm1.txtsub_cd_nm2.value			' Name Cindition
	        arrParam(4) =  " PAY_CD = " & FilterVar("*", "''", "S") & "  AND CODE_TYPE = " & FilterVar("2", "''", "S") & " "   ' Where Condition
	        arrParam(5) = "공제코드"			    ' TextBox 명칭 
	
            arrField(0) = "ALLOW_CD"					' Field명(0)
            arrField(1) = "ALLOW_NM"				    ' Field명(1)
    
            arrHeader(0) = "공제코드"				' Header명(0)
            arrHeader(1) = "공제코드명"
            
       Case "7" 
            arrParam(0) = "근무구역 팝업"	       	' 팝업 명칭 
	        arrParam(1) = "B_minor"					    ' TABLE 명칭 
	    	arrParam(2) = frm1.txtSect_cd.Value	        ' Code Condition
	    	arrParam(3) = ""'frm1.txtSect_cd_nm.Value 		' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0035", "''", "S") & ""	    	' Where Condition
	    	arrParam(5) = "근무구역"    		    ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"					' Field명(0)
	    	arrField(1) = "minor_nm"    				' Field명(1)
    
	    	arrHeader(0) = "근무구역코드"			' Header명(0)
	    	arrHeader(1) = "근무구역명"     		' Header명(1) 
	  Case "8" 
            arrParam(0) = "근무구역 팝업"	       	' 팝업 명칭 
	        arrParam(1) = "B_minor"					    ' TABLE 명칭 
	    	arrParam(2) = frm1.txtSect_cd2.Value	    ' Code Condition
	    	arrParam(3) = ""'frm1.txtSect_cd_nm2.Value		' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0035", "''", "S") & ""	    	' Where Condition
	    	arrParam(5) = "근무구역"    		    ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"					' Field명(0)
	    	arrField(1) = "minor_nm"    				' Field명(1)
    
	    	arrHeader(0) = "근무구역코드"			' Header명(0)
	    	arrHeader(1) = "근무구역명"     		' Header명(1)  
	    	
	  Case "9"  	
            arrParam(0) = "지급구분팝업"			' 팝업 명칭 
            arrParam(1) = "B_MINOR"				 		' TABLE 명칭 
            arrParam(2) = frm1.txtPayCd.value		    ' Code Condition
            arrParam(3) = ""'frm1.txtPayNM.value			' Name Cindition
            arrParam(4) = "MAJOR_CD = " & FilterVar("H0040", "''", "S") & ""			' Where Condition
            arrParam(5) = "지급구분"			    ' TextBox 명칭 
	
            arrField(0) = "MINOR_CD"					' Field명(0)
            arrField(1) = "MINOR_NM"				    ' Field명(1)
    
            arrHeader(0) = "지급구분"				' Header명(0)
            arrHeader(1) = "지급명"			        ' Header명(1)
      
      Case "10"       
            arrParam(0) = "공제구분 팝업"			' 팝업 명칭 
	        arrParam(1) = "B_MINOR"				 		' TABLE 명칭 
	        arrParam(2) = frm1.txtsub_type2.value		    ' Code Condition
	        arrParam(3) = ""'frm1.txtsub_type_nm2.value		' Name Cindition
	        arrParam(4) = " MAJOR_CD = " & FilterVar("H0040", "''", "S") & " "    ' Where Condition							
	        arrParam(5) = "공제구분"			    ' TextBox 명칭 
	
            arrField(0) = "minor_cd"					' Field명(0)
            arrField(1) = "minor_nm"				    ' Field명(1)
    
            arrHeader(0) = "공제구분코드"				' Header명(0)
            arrHeader(1) = "공제구분명"  		                                    ' Header명(1)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Select Case iWhere
		    Case "5"    
		        frm1.txtAllow_cd.focus
		    Case "6"
		        frm1.txtsub_cd2.focus
		    Case "7" 
			    frm1.txtSect_cd.focus
		    Case "8" 
			    frm1.txtSect_cd2.focus
			Case "9"
                frm1.txtPayCd.focus
			Case "10"
		        frm1.txtsub_type2.focus
        End Select	
		Exit Function
	Else
		Call SubSetCondArea(arrRet,iWhere)
	End If	
	
End Function

'======================================================================================================
'	Name : SetCondArea()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetCondArea(Byval arrRet, Byval iWhere)
	With Frm1
		Select Case iWhere
		   
		    Case "5"    
		        .txtAllow_cd.value    = arrRet(0)
		        .txtAllow_nm.value    = arrRet(1)		
		        .txtAllow_cd.focus
		    Case "6"
		        .txtsub_cd2.value     = arrRet(0)
		        .txtsub_cd_nm2.value  = arrRet(1)
		        .txtsub_cd2.focus
		    Case "7" 
			    .txtSect_cd.value     = arrRet(0) 
			    .txtSect_cd_nm.value  = arrRet(1)   
			    .txtSect_cd.focus
		    Case "8" 
			    .txtSect_cd2.value    = arrRet(0) 
			    .txtSect_cd_nm2.value = arrRet(1)
			    .txtSect_cd2.focus
			Case "9"
		    	.txtPayCd.value = arrRet(0)
                .txtPayNm.value = arrRet(1)		
                .txtPayCd.focus
			Case "10"
		        .txtsub_type2.value = arrRet(0)
		        .txtsub_type_nm2.value = arrRet(1)   	    
		        .txtsub_type2.focus
        End Select
	End With
End Sub
'========================================================================================================
' Name : OpenDept
' Desc : 부서 POPUP
'========================================================================================================
Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(2)
	Dim StrDt,StrDt2
    	StrDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtTo_dt.Year, Right("0" & frm1.txtTo_dt.month , 2), "01")
        StrDt2 = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtTo_dt2.Year, Right("0" & frm1.txtTo_dt2.month , 2), "01")
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then
		arrParam(0) = frm1.txtFr_dept_cd.value			            '  Code Condition
	ElseIf iWhere = 1 Then
		arrParam(0) = frm1.txtTo_dept_cd.value			            ' Code Condition
    	ElseIf iWhere = 2 Then
		arrParam(0) = frm1.txtFr_dept_cd2.value			            ' Code Condition
	ElseIf iWhere = 3 Then
		arrParam(0) = frm1.txtTo_dept_cd2.value			            ' Code Condition	    		
	End If
	
        If iWhere = 0 Then
		arrParam(1) = StrDt 
	ElseIf iWhere = 1 Then
		arrParam(1) = StrDt 
    	ElseIf iWhere = 2 Then
		arrParam(1) = StrDt2
	ElseIf iWhere = 3 Then
		arrParam(1) = StrDt2 		
	End If

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
             Case "2"
               frm1.txtFr_dept_cd2.focus
             Case "3"  
               frm1.txtTo_dept_cd2.focus
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
             Case "2"
               .txtFr_dept_cd2.value = arrRet(0)
               .txtFr_dept_nm2.value = arrRet(1)
               .txtFr_internal_cd2.value = arrRet(2)
               .txtFr_dept_cd2.focus
             Case "3"  
               .txtTo_dept_cd2.value = arrRet(0)
               .txtTo_dept_nm2.value = arrRet(1) 
               .txtTo_internal_cd2.value = arrRet(2) 
               .txtTo_dept_cd2.focus
        End Select
	End With
End Function    

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

End Sub

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)

End Sub
'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    
    Call SetPopupMenuItemInf("0000111111")

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

Sub vspdData2_Click(ByVal Col, ByVal Row)
    
    Call SetPopupMenuItemInf("0000111111")

    gMouseClickStatus = "SP1C" 

    Set gActiveSpdSheet = frm1.vspdData2
   
    If frm1.vspdData2.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData2
       
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
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData2_GotFocus()
    ggoSpread.Source = Frm1.vspdData2
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
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
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
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
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
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
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

'======================================================================================================
'   Event Name : txtallow_cd_OnChange
'   Event Desc : 수당코드 에러체크 
'=======================================================================================================
Function txtallow_cd_OnChange()
    Dim iDx
    Dim IntRetCd   
        
    IF frm1.txtallow_cd.value = "" THEN
        frm1.txtallow_nm.value = ""
        txtallow_cd_OnChange = true
    ELSE    
        IntRetCd = CommonQueryRs(" allow_nm "," HDA010T "," PAY_CD = " & FilterVar("*", "''", "S") & "  AND CODE_TYPE = " & FilterVar("1", "''", "S") & "  AND ALLOW_CD =  " & FilterVar(frm1.txtallow_cd.value, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        IF IntRetCd = false  Then
            Call DisplayMsgBox("800145","X","X","X")
            frm1.txtallow_nm.value=""
            frm1.txtallow_cd.focus()                
			Set gActiveElement = document.ActiveElement 
'            txtallow_cd_OnChange = false
  	        Exit Function                
        ELSE   '수당코드 
            frm1.txtallow_nm.value=Trim(Replace(lgF0,Chr(11),""))
            txtallow_cd_OnChange = true
        END IF
    END IF  
End Function 
'========================================================================================================
'   Event Name : txtPayCd_Onchange()            
'   Event Desc : 지급구분 에러체크 
'========================================================================================================
    Function txtPayCd_Onchange()
        Dim iDx
        Dim IntRetCd
        
        IF frm1.txtPayCd.value = "" THEN
            frm1.txtPayNm.value = ""
            txtPayCd_Onchange = true
        ELSE
            IntRetCd = CommonQueryRs(" minor_nm "," b_minor "," major_cd = " & FilterVar("H0040", "''", "S") & " and minor_cd =  " & FilterVar(frm1.txtPayCd.value , "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            If IntRetCd = false then
                Call DisplayMsgBox("970000","X","지급구분코드","X")
                frm1.txtPayNm.value = ""
                frm1.txtPayCd.focus()
				Set gActiveElement = document.ActiveElement 
                'txtPayCd_Onchange = false
    	        Exit Function                
            ELSE    
                frm1.txtPayNm.value = Trim(Replace(lgF0,Chr(11),""))   '수당코드 
                txtPayCd_Onchange = true
            END IF
        END IF 
    End Function 

'======================================================================================================
'   Event Name : txtsub_cd2_Onchange
'   Event Desc : 공제코드 에러체크 
'=======================================================================================================
    Function txtsub_cd2_Onchange()
        Dim iDx
        Dim IntRetCd
        
        IF frm1.txtsub_cd2.value = "" THEN
            frm1.txtsub_cd_nm2.value = ""
            txtsub_cd2_Onchange = true
        ELSE    
            IntRetCd = CommonQueryRs(" allow_nm "," HDA010T "," CODE_TYPE=" & FilterVar("2", "''", "S") & " AND allow_cd =  " & FilterVar(frm1.txtsub_cd2.value , "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            IF IntRetCd = false Then
                Call DisplayMsgBox("800176","X","X","X")
                frm1.txtsub_cd_nm2.value = ""
                frm1.txtsub_cd2.focus()   
                Set gActiveElement = document.ActiveElement 
                'txtsub_cd2_Onchange = false
    	        Exit Function                
            ELSE
                frm1.txtsub_cd_nm2.value = Trim(Replace(lgF0,Chr(11),""))
                txtsub_cd2_Onchange = true
            END IF
        
        END IF 
    End Function 
'======================================================================================================
'   Event Name : txtSect_cd_OnChange
'   Event Desc : 근무구역코드 에러체크 
'=======================================================================================================
Function txtSect_cd_OnChange()
    Dim iDx
    Dim IntRetCd
        
    If frm1.txtSect_cd.value = "" Then
        frm1.txtSect_cd_nm.value = ""
        txtSect_cd_OnChange = true
    ELSE
        IntRetCd = CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0035", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtSect_cd.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
        IF IntRetCd = False THEN
            Call DisplayMsgBox("970000","X","근무구역코드","X")
                
	        frm1.txtSect_cd_nm.value = ""
	        Set gActiveElement = document.ActiveElement
	        Exit Function
	    Else
	        frm1.txtSect_cd_nm.value = Trim(Replace(lgF0, Chr(11), ""))
	        txtSect_cd_OnChange = true
	    End If
    End If
    	    
End Function
    
'======================================================================================================
'   Event Name : txtSect_cd_OnChange
'   Event Desc : 근무구역코드 에러체크 
'=======================================================================================================
   Function txtSect_cd2_OnChange()
        Dim iDx
        Dim IntRetCd
        If frm1.txtSect_cd2.value = "" Then
            frm1.txtSect_cd_nm2.value = ""
  	        txtSect_cd2_OnChange = true
        ELSE
            IntRetCd = CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0035", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtSect_cd2.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
            IF IntRetCd = False THEN
                Call DisplayMsgBox("970000","X","근무구역코드","X")
    	        frm1.txtSect_cd_nm2.value = ""
    	        Set gActiveElement = document.ActiveElement
    	        Exit Function
    	    Else
    	        frm1.txtSect_cd_nm2.value = Trim(Replace(lgF0, Chr(11), ""))
    	        txtSect_cd2_OnChange = true
    	    End If
        End If
    	    
    End Function

'========================================================================================================
'   Event Name : txtsub_type2_OnChange()             
'   Event Desc : 공제구분 에러체크 
'========================================================================================================

    Function txtsub_type2_OnChange()
        Dim iDx
        Dim IntRetCd

        IF frm1.txtsub_type2.value = "" THEN
            frm1.txtsub_type_nm2.value = ""
            txtsub_type2_OnChange = true
        ELSE
            IntRetCd = CommonQueryRs(" minor_nm "," b_minor "," major_cd = " & FilterVar("H0040", "''", "S") & " and minor_cd =  " & FilterVar(frm1.txtsub_type2.value , "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            IF IntRetCd = false Then
                Call DisplayMsgBox("970000","X","공제구분코드","X")
               
                frm1.txtsub_type_nm2.value = ""
                Set gActiveElement = document.ActiveElement
                Exit Function              
            ELSE
                frm1.txtsub_type_nm2.value = Trim(Replace(lgF0,Chr(11),""))
                txtsub_type2_OnChange = true
            END IF
        END IF  
        
    End Function 

'========================================================================================================
'   Event Name : txtFr_dept_cd_Onchange
'   Event Desc :
'========================================================================================================
Function txtFr_dept_cd_Onchange()
   
    Dim IntRetCd,Dept_Nm,Internal_cd
    Dim StrDt
    StrDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtTo_dt.Year, Right("0" & frm1.txtTo_dt.month , 2), "01")	   
    If frm1.txtFr_dept_cd.value = "" Then
		frm1.txtFr_dept_nm.value = ""
		frm1.txtFr_internal_cd.value = ""
        txtFr_dept_cd_Onchange = true
    Else
        IntRetCd = FuncDeptName(frm1.txtFr_dept_cd.value , StrDt , lgUsrIntCd,Dept_Nm , Internal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800098","X","X","X")	'부서정보코드에 등록되지 않은 코드입니다.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
		    frm1.txtFr_dept_nm.value = ""
		    frm1.txtFr_internal_cd.value = ""
            frm1.txtFr_dept_cd.focus()
            Set gActiveElement = document.ActiveElement 
            Exit Function      
        Else
			frm1.txtFr_dept_nm.value = Dept_Nm
		    frm1.txtFr_internal_cd.value = Internal_cd
		    
            txtFr_dept_cd_Onchange = true
        End if 
    End if  
    
End Function

'------------------------------------------------------------------------------------------------
Function txtFr_dept_cd2_Onchange()
    
	Dim IntRetCd,Dept_Nm,Internal_cd
    Dim StrDt
    StrDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtTo_dt2.Year, Right("0" & frm1.txtTo_dt2.month , 2), "01")
    If frm1.txtFr_dept_cd2.value = "" Then
		frm1.txtFr_dept_nm2.value = ""
		frm1.txtFr_internal_cd2.value = ""
        txtFr_dept_cd2_Onchange = true
    Else
        IntRetCd = FuncDeptName(frm1.txtFr_dept_cd2.value , StrDt , lgUsrIntCd,Dept_Nm , Internal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800098","X","X","X")	'부서정보코드에 등록되지 않은 코드입니다.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
		    frm1.txtFr_dept_nm2.value = ""
		    frm1.txtFr_internal_cd2.value = ""
            frm1.txtFr_dept_cd2.focus()
            Set gActiveElement = document.ActiveElement 
            Exit Function      
        Else
			frm1.txtFr_dept_nm2.value = Dept_Nm
		    frm1.txtFr_internal_cd2.value = Internal_cd
            txtFr_dept_cd2_Onchange = true
        End if 
    End if  
    
End Function

'========================================================================================================
'   Event Name : txtTo_dept_cd_Onchange
'   Event Desc :
'========================================================================================================
Function txtTo_dept_cd_Onchange()
   
    Dim IntRetCd,Dept_Nm,Internal_cd
    Dim StrDt
    StrDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtTo_dt.Year, Right("0" & frm1.txtTo_dt.month , 2), "01")
    If frm1.txtTo_dept_cd.value = "" Then
		frm1.txtTo_dept_nm.value = ""
		frm1.txtTo_internal_cd.value = ""
        txtTo_dept_cd_Onchange = true
    Else
        IntRetCd = FuncDeptName(frm1.txtTo_dept_cd.value , StrDt , lgUsrIntCd,Dept_Nm , Internal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800098","X","X","X")	'부서정보코드에 등록되지 않은 코드입니다.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
		    frm1.txtTo_dept_nm.value = ""
		    frm1.txtTo_internal_cd.value = ""
            frm1.txtTo_dept_cd.focus()
            Set gActiveElement = document.ActiveElement 
            Exit Function      
        Else
			frm1.txtTo_dept_nm.value = Dept_Nm
		    frm1.txtTo_internal_cd.value = Internal_cd
		    
            txtTo_dept_cd_Onchange = true
        End if 
    End if  
    
End Function
'--------------------------------------------------------------------------------------------
Function txtTo_dept_cd2_Onchange()
   
    Dim IntRetCd,Dept_Nm,Internal_cd
    Dim StrDt
    StrDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtTo_dt2.Year, Right("0" & frm1.txtTo_dt2.month , 2), "01")    

    If frm1.txtTo_dept_cd2.value = "" Then
		frm1.txtTo_dept_nm2.value = ""
		frm1.txtTo_internal_cd2.value = ""
        txtTo_dept_cd2_Onchange = true
    Else
        IntRetCd = FuncDeptName(frm1.txtTo_dept_cd2.value , StrDt , lgUsrIntCd,Dept_Nm , Internal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800098","X","X","X")	'부서정보코드에 등록되지 않은 코드입니다.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
		    frm1.txtTo_dept_nm2.value = ""
		    frm1.txtTo_internal_cd2.value = ""
            frm1.txtTo_dept_cd2.focus()
            Set gActiveElement = document.ActiveElement 
            
            'txtTo_dept_cd2_Onchange = false
            Exit Function      
        Else
			frm1.txtTo_dept_nm2.value = Dept_Nm
		    frm1.txtTo_internal_cd2.value = Internal_cd
		    
            txtTo_dept_cd2_Onchange = true
        End if 
    End if  
    
End Function

'=======================================================================================================
'   Event Name : txtFrom_dt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFrom_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")        
        frm1.txtFrom_dt.Action = 7
        frm1.txtFrom_dt.focus
    End If
End Sub

Sub txtFrom_dt2_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtFrom_dt2.Action = 7
        frm1.txtFrom_dt2.focus
    End If
End Sub
'-----------------------------------   Event Name : txtTo_dt_DblClick(Button)
Sub txtTo_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtTo_dt.Action = 7
        frm1.txtTo_dt.focus
    End If
End Sub

Sub txtTo_dt2_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtTo_dt2.Action = 7
        frm1.txtTo_dt2.focus
    End If
End Sub
'=======================================================================================================
'   Event Name : txtFrom_dt_Keypress(Key)
'   Event Desc : 3rd party control에서 Enter 키를 누르면 조회 실행 
'=======================================================================================================
Sub txtFrom_dt_Keypress(Key)
    If Key = 13 Then
        MAINQUERY()
    End If
End Sub

Sub txtFrom_dt2_Keypress(Key)
    If Key = 13 Then
        MAINQUERY()
    End If
End Sub

Sub txtTo_dt_Keypress(Key)
    If Key = 13 Then
        MAINQUERY()
    End If
End Sub

Sub txtTo_dt2_Keypress(Key)
    If Key = 13 Then
        MAINQUERY()
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>수당내역</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>공제내역</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS="Tab11">
		
		
		
		<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">					
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
        				    <FIELDSET CLASS="CLSFLD">
						  	<TABLE <%=LR_SPACE_TYPE_40%>>
						  		<TR>
						  			<TD CLASS=TD5 NOWRAP>기준연월</TD>
						  			<TD CLASS=TD6 NOWRAP><OBJECT classid=<%=gCLSIDFPDT%> id=txtFrom_dt NAME="txtFrom_dt" CLASS= FPDTYYYYMM title=FPDATETIME  ALT="비교시작기간" tag="12X1" VIEWASTEXT></OBJECT>
						  	             
						  	        <TD CLASS=TD5 NOWRAP>비교연월</TD>
						  			<TD CLASS=TD6 NOWRAP><OBJECT classid=<%=gCLSIDFPDT%> id=txtTo_dt NAME="txtTo_dt" CLASS= FPDTYYYYMM title=FPDATETIME  ALT="비교종료기간" tag="12X1" VIEWASTEXT></OBJECT></TD>
						  		</TR> 
						  		<TR>	
						  			<TD CLASS=TD5 NOWRAP>수당코드</TD>
				 		  			<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtAllow_cd" SIZE="10" MAXLENGTH="3" tag="12XXXU" ALT="수당코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCondAreaPopup(5)"> 
				 		  			                     <INPUT TYPE="Text" NAME="txtAllow_nm" SIZE="20" MAXLENGTH="20" tag="14XXXU" ALT="수당코드"></TD>
						  		
						  		    <TD CLASS=TD5 NOWRAP>직종</TD>
						  		    <TD CLASS=TD6 NOWRAP><SELECT NAME="txtOcpt_type" ALT="직종" STYLE="WIDTH: 100px" TAG="11XXXU"><OPTION VALUE=""></OPTION></SELECT>
						  		</TR>
						  		<TR>
						  		    <TD CLASS=TD5 NOWRAP>지급구분</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayCd" MAXLENGTH="1" SIZE="10" ALT ="지급구분" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCondAreaPopup(9)">
						                                 <INPUT NAME="txtPayNm" MAXLENGTH="20" SIZE="20" ALT ="지급구분명" tag="14XXXU"></TD>
						  		
						  		    <TD CLASS=TD5 NOWRAP>부서코드</TD>
						  			<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFr_dept_cd" ALT="부서코드" TYPE="Text" SiZE="10" MAXLENGTH="10" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(0)">
			                                             <INPUT NAME="txtFr_dept_nm" ALT="부서코드명" TYPE="Text" SiZE="20" MAXLENGTH="40" tag="14XXXU">&nbsp;~
		                                                 <INPUT NAME="txtFr_Internal_cd" ALT="내부부서코드" TYPE="HIDDEN" SiZE="7" MAXLENGTH="7" tag="14XXXU">  
						  		</TR>
						  		<TR>	
						  			<TD CLASS=TD5 NOWRAP>근무구역</TD>
					 	  			<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtSect_cd" SIZE="10" MAXLENGTH="10" tag="11XXXU" ALT="근무구역"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCondAreaPopup(7)">
					 	  			                     <INPUT TYPE="Text" NAME="txtSect_cd_nm" SIZE="20" MAXLENGTH="50" tag="14XXXU" ALT="근무구역"></TD>
						  			
						  			<TD CLASS=TD5 NOWRAP>부서코드</TD>
						  			<TD CLASS=TD6 NOWRAP><INPUT NAME="txtto_dept_cd" MAXLENGTH="10" SIZE="10" ALT ="Order ID" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(1)">
							                             <INPUT NAME="txtto_dept_nm" MAXLENGTH="40" SIZE="20" ALT ="Order ID" tag="14XXXU">
    			                                         <INPUT  NAME="txtTo_Internal_cd" ALT="내부부서코드" TYPE="HIDDEN" SiZE="7" MAXLENGTH="7" tag="14XXXU"></TD>
						  		</TR>
						  		<TR>	
						  			<TD CLASS=TD5 NOWRAP>조회구분</TD>
				        	    	<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rbo_sort" ID="rbo_sort" VALUE="1" CLASS="RADIO" TAG="11" CHECKED><LABEL FOR="rbo_sort1">전체조회</LABEL>&nbsp;
				        	                             <INPUT TYPE="RADIO" NAME="rbo_sort" ID="rbo_sort" VALUE="2" CLASS="RADIO" TAG="11"><LABEL FOR="rbo_sort2">변경사항조회</LABEL></TD>
									<TD CLASS=TD5 NOWRAP></TD>
				        	    	<TD CLASS=TD6 NOWRAP></TD>				        	                             
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
			</TABLE>
        </DIV>    

		<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">					
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
        				    <FIELDSET CLASS="CLSFLD">
						  	<TABLE <%=LR_SPACE_TYPE_40%>>
						  		<TR>
						  			<TD CLASS=TD5 NOWRAP>기준연월</TD>
						  			<TD CLASS=TD6 NOWRAP><OBJECT id=txtFrom_dt2 title=FPDATETIME CLASS= FPDTYYYYMM name=txtFrom_dt2 classid=<%=gCLSIDFPDT%> ALT="비교시작기간" tag="12X1" VIEWASTEXT></OBJECT>
						  	        <TD CLASS=TD5 NOWRAP>비교연월</TD>
						  			<TD CLASS=TD6 NOWRAP><OBJECT id=txtTo_dt2 title=FPDATETIME CLASS= FPDTYYYYMM name=txtTo_dt2 classid=<%=gCLSIDFPDT%>ALT="비교종료기간" tag="12x1" VIEWASTEXT></OBJECT></TD>
						  		</TR> 
						  		<TR>	
						  			<TD CLASS=TD5 NOWRAP>공제코드</TD>
								    <TD CLASS=TD6 NOWRAP><INPUT NAME="txtsub_cd2"  MAXLENGTH="3" SIZE="10" ALT ="공제코드" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCondAreaPopup(6)">
								                         <INPUT NAME="txtsub_cd_nm2"  MAXLENGTH="20" SIZE="20" ALT ="공제코드" tag="14XXXU"></TD>
						  		    <TD CLASS=TD5 NOWRAP>직종</TD>
						  		    <TD CLASS=TD6 NOWRAP><SELECT NAME="txtOcpt_type2" ALT="직종" STYLE="WIDTH: 100px" TAG="11XXXU"><OPTION VALUE=""></OPTION></SELECT>
						  		</TR>
						  		<TR>	
						  			<TD CLASS=TD5 NOWRAP>공제구분</TD>
								    <TD CLASS=TD6 NOWRAP><INPUT NAME="txtsub_type2" MAXLENGTH="1" SIZE="10" ALT ="공제구분" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCondAreaPopup(10)">
								                         <INPUT NAME="txtsub_type_nm2" MAXLENGTH="20" SIZE="20" ALT ="Order ID" tag="14XXXU"></td>
						  		    <TD CLASS=TD5 NOWRAP>부서코드</TD>
						  			<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFr_dept_cd2" ALT="부서코드" TYPE="Text" SiZE="10" MAXLENGTH="10" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(2)">
			                                             <INPUT NAME="txtFr_dept_nm2" ALT="부서코드명" TYPE="Text" SiZE="20" MAXLENGTH="40" tag="14XXXU">&nbsp;~
		                                                 <INPUT NAME="txtFr_Internal_cd2" ALT="내부부서코드" TYPE="HIDDEN" SiZE="7" MAXLENGTH="7" tag="14XXXU">  
				        	    </TR>
						  		<TR>    
				        	        <TD CLASS=TD5 NOWRAP>근무구역</TD>
					 	  			<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtSect_cd2" SIZE="10" MAXLENGTH="10"  tag="11XXXU" ALT="근무구역"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCondAreaPopup(8)">
					 	  			                     <INPUT TYPE="Text" NAME="txtSect_cd_nm2" SIZE=20 MAXLENGTH="50" tag="14XXXU" ALT="근무구역"></TD>                       
						  			
						  			<TD CLASS=TD5 NOWRAP>부서코드</TD>
						  			<TD CLASS=TD6 NOWRAP><INPUT NAME="txtto_dept_cd2" MAXLENGTH="10" SIZE="10" ALT ="Order ID" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(3)">
							                             <INPUT NAME="txtto_dept_nm2" MAXLENGTH="40" SIZE="20" ALT ="Order ID" tag="14XXXU">
    			                                         <INPUT  NAME="txtTo_Internal_cd2" ALT="내부부서코드" TYPE="HIDDEN" SiZE="7" MAXLENGTH="7" tag="14XXXU"></TD>
						  		</TR>
						  		<TR>	
						  			<TD CLASS=TD5 NOWRAP>조회구분</TD>
				        	    	<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rbo_sort2" ID="rbo_sort2" VALUE="1" CLASS="RADIO" TAG="11" CHECKED><LABEL FOR="rbo_sort1">전체조회</LABEL>&nbsp;
				        	                             <INPUT TYPE="RADIO" NAME="rbo_sort2" ID="rbo_sort2" VALUE="2" CLASS="RADIO" TAG="11"><LABEL FOR="rbo_sort2">변경사항조회</LABEL></TD>
									<TD CLASS=TD5 NOWRAP></TD>
				        	    	<TD CLASS=TD6 NOWRAP></TD>				        	                             
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
        </DIV>    

		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
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
