<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          	: Human Resources
*  2. Function Name        	: Multi Sample
*  3. Program ID           	: H6012ma1
*  4. Program Name         	: H6012ma1
*  5. Program Desc         	: 급여조회 
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID       = "h6012mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_ID1      = "h6012mb2.asp"
Const BIZ_PGM_JUMP_ID  = "h6010ma1"
Const BIZ_PGM_JUMP_ID2 = "h7006ma1"
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
Dim lsInternal_cd

Dim C_NAME			'성명												<%'Spread Sheet의 Column별 상수 %>
Dim C_EMP_NO          '사번 
Dim C_DEPT_CD         '부서명 
Dim C_PAY_CD          '급여구분 
Dim C_PROV_TYPE       '지급구분 
Dim C_PROV_TYPE_HIDDEN       '지급구분 
Dim C_PAY_TOT_AMT     '급여총액 
Dim C_INCOME_TAX      '소득세 
Dim C_RES_TAX         '주민세 
Dim C_ANUT            '국민연금 
Dim C_MED_INSUR      '의료보험 
Dim C_EMP_INSUR      '고용보험 
Dim C_SUB_TOT_AMT    '공제총액 
Dim C_REAL_PROV_AMT  '실지급액 

Dim C_NAME2			 '성명												<%'Spread Sheet의 Column별 상수 %>
Dim C_EMP_NO2          '사번 
Dim C_DEPT_CD2         '부서명 
Dim C_PAY_CD2          '급여구분 
Dim C_PROV_TYPE2       '지급구분 
Dim C_PROV_TYPE_HIDDEN2    '지급구분 
Dim C_PAY_TOT_AMT2     '급여총액 
Dim C_INCOME_TAX2      '소득세 
Dim C_RES_TAX2          '주민세 
Dim C_ANUT2             '국민연금 
Dim C_MED_INSUR2       '의료보험 
Dim C_EMP_INSUR2       '고용보험 
Dim C_SUB_TOT_AMT2     '공제총액 
Dim C_REAL_PROV_AMT2   '실지급액 

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  

    C_NAME = 1			  '성명												<%'Spread Sheet의 Column별 상수 %>
    C_EMP_NO = 2          '사번 
    C_DEPT_CD = 3         '부서명 
    C_PAY_CD = 4          '급여구분 
    C_PROV_TYPE = 5       '지급구분 
    C_PROV_TYPE_HIDDEN = 6       '지급구분 
    C_PAY_TOT_AMT = 7     '급여총액 
    C_INCOME_TAX = 8      '소득세 
    C_RES_TAX = 9         '주민세 
    C_ANUT = 10            '국민연금 
    C_MED_INSUR = 11      '의료보험 
    C_EMP_INSUR = 12      '고용보험 
    C_SUB_TOT_AMT = 13    '공제총액 
    C_REAL_PROV_AMT = 14  '실지급액 
    
    C_NAME2 = 1			   '성명												<%'Spread Sheet의 Column별 상수 %>
    C_EMP_NO2 = 2          '사번 
    C_DEPT_CD2 = 3         '부서명 
    C_PAY_CD2 = 4          '급여구분 
    C_PROV_TYPE2 = 5       '지급구분 
    C_PROV_TYPE_HIDDEN2 = 6       '지급구분 
    C_PAY_TOT_AMT2 = 7     '급여총액 
    C_INCOME_TAX2 = 8      '소득세 
    C_RES_TAX2 = 9         '주민세 
    C_ANUT2 = 10            '국민연금 
    C_MED_INSUR2 = 11      '의료보험 
    C_EMP_INSUR2 = 12      '고용보험 
    C_SUB_TOT_AMT2 = 13    '공제총액 
    C_REAL_PROV_AMT2 = 14  '실지급액 

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
	
	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	frm1.txtpay_yymm_dt.focus()	
	frm1.txtpay_yymm_dt.Year = strYear 		 '년월일 default value setting
	frm1.txtpay_yymm_dt.Month = strMonth 
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H", "NOCOOKIE", "MA") %>
End Sub

'------------------------------------------  CookiePage()  --------------------------------------------------
'	Name : CookiePage()
'	Description : Jump시 Condition에서 넘겨오는 값 setting
'---------------------------------------------------------------------------------------------------------
Function CookiePage(Byval Kubun)
    
    Const CookieSplit = 4877	
	
	Dim strTemp, arrVal
	Dim IntRetCD
	
    
	If Kubun = 1 Then   
        
        frm1.vspdData.Row = frm1.vspdData.ActiveRow
	    WriteCookie "PAY_YYMM_DT" , frm1.txtpay_yymm_dt.Text
	    
	    frm1.vspdData.Col = C_EMP_NO               
	    WriteCookie "EMP_NO" , frm1.vspdData.Text
	    
	    frm1.vspdData.Col = C_PROV_TYPE_HIDDEN 
	    WriteCookie "PROV_TYPE_HIDDEN"   , frm1.vspdData.Text

	Else

		strTemp = ReadCookie("PAY_YYMM_DT")                      '           Kubun = 0 일때 수정금지 요망........!    
		If strTemp = "" then Exit Function
        
        frm1.txtpay_yymm_dt.text = strTemp
		FncQuery()              
	
		WriteCookie "PAY_YYMM_DT" , ""
	    WriteCookie "EMP_NO"      , ""
        WriteCookie "PROV_TYPE_HIDDEN"   , ""
        
	End IF
End Function

'--------------------------	Description : 상세조회 클릭시 에러 체크사항 -------------------------------
FUNCTION PgmJumpCheck()         
    If frm1.vspdData.ActiveRow =  0 Then
		Call DisplayMsgBox("800167","X","X","X")
		frm1.txtpay_yymm_dt.focus		
	    Exit Function
	Else
	   
        If 	frm1.vspdData.Text = "1" Or frm1.vspdData.Text = "P" then   
	        PgmJump(BIZ_PGM_JUMP_ID)
	    
	    ElseIf (frm1.vspdData.Text >= "2" And frm1.vspdData.Text <= "9") Or frm1.vspdData.Text = "Q" then  
	        PgmJump(BIZ_PGM_JUMP_ID2)    
	    End if    
	   
	End If	   
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
	Dim strPayYYYYMM
	Dim strPayYYYY
	DIm strPayMM
        Dim Strdt

    strPayYYYY = Frm1.txtpay_yymm_dt.Year
    strPayMM  = Frm1.txtpay_yymm_dt.Month
    
     Strdt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtpay_yymm_dt.Year, Right("0" & frm1.txtpay_yymm_dt.month , 2), "01")
    
    If len(strPayMM) = 1 Then
		strPayMM = "0" & strPayMM
	End if
    strPayYYYYMM = strPayYYYY & strPayMM
    
    lgKeyStream  = strPayYYYYMM & Parent.gColSep
    lgKeyStream  = lgKeyStream & Trim(Frm1.txtemp_no.value) & Parent.gColSep
    lgKeyStream  = lgKeyStream & Trim(Frm1.cboPay_cd.value) & Parent.gColSep
    lgKeyStream  = lgKeyStream & Trim(Frm1.txtprov_cd.value) & Parent.gColSep    
    lgKeyStream  = lgKeyStream & Trim(Frm1.txtfr_internal_cd.value) & Parent.gColSep
    lgKeyStream  = lgKeyStream & Trim(Frm1.txtto_internal_cd.value) & Parent.gColSep
    lgKeyStream  = lgKeyStream & lgUsrIntcd & Parent.gColSep
    lgKeyStream  = lgKeyStream & Strdt & Parent.gColSep    
End Sub        
	
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx

    Call CommonQueryRs("MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0005", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1

    Call SetCombo2(frm1.cboPay_cd,iCodeArr, iNameArr,Chr(11))'    iCodeArr = lgF0
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
        intIndex = ggoSpread.InsertRow
        
        frm1.vspdData2.Col = 0
        frm1.vspdData2.Text = "합계"
        
        frm1.vspdData2.Col = C_PAY_TOT_AMT2  '급여총액 
        frm1.vspdData2.text = FncSumSheet(frm1.vspdData,C_PAY_TOT_AMT, 1, .MaxRows,FALSE , -1, -1, "V")
        
        frm1.vspdData2.Col = C_INCOME_TAX2  '소득세 
        frm1.vspdData2.text = FncSumSheet(frm1.vspdData,C_INCOME_TAX, 1, .MaxRows,FALSE , -1, -1, "V")
        
        frm1.vspdData2.Col = C_RES_TAX2  '주민세 
        frm1.vspdData2.text = FncSumSheet(frm1.vspdData,C_RES_TAX, 1, .MaxRows,FALSE , -1, -1, "V")
        
        frm1.vspdData2.Col = C_ANUT2  '국민연금 
        frm1.vspdData2.text = FncSumSheet(frm1.vspdData,C_ANUT, 1, .MaxRows,FALSE , -1, -1, "V")
        
        frm1.vspdData2.Col = C_MED_INSUR2  '의료보험 
        frm1.vspdData2.text = FncSumSheet(frm1.vspdData,C_MED_INSUR, 1, .MaxRows,FALSE , -1, -1, "V")
        
        frm1.vspdData2.Col = C_EMP_INSUR2  '고용보험 
        frm1.vspdData2.text = FncSumSheet(frm1.vspdData,C_EMP_INSUR, 1, .MaxRows,FALSE , -1, -1, "V")
        
        frm1.vspdData2.Col = C_SUB_TOT_AMT2  '공제총액 
        frm1.vspdData2.text = FncSumSheet(frm1.vspdData,C_SUB_TOT_AMT, 1, .MaxRows,FALSE , -1, -1, "V")
        
        frm1.vspdData2.Col = C_REAL_PROV_AMT2  '실지급액   
        frm1.vspdData2.text = FncSumSheet(frm1.vspdData,C_REAL_PROV_AMT, 1, .MaxRows,FALSE , -1, -1, "V")
     End With

    With frm1
        ggoSpread.Source = frm1.vspdData2
        .vspdData2.ReDraw = False
        ggoSpread.SpreadLock C_NAME2, -1, C_NAME2, -1        '성명	
        ggoSpread.SpreadLock C_EMP_NO2, -1, C_EMP_NO2, -1      '사번 
        ggoSpread.SpreadLock C_DEPT_CD2, -1, C_DEPT_CD2, -1     '부서명     
        ggoSpread.SpreadLock C_PAY_CD2, -1, C_PAY_CD2, -1      '급여구분 
        ggoSpread.SpreadLock C_PROV_TYPE2, -1, C_PROV_TYPE2, -1   '지급구분 
        ggoSpread.SpreadLock C_PAY_TOT_AMT2, -1, C_PAY_TOT_AMT2, -1  '급여총액 
        ggoSpread.SpreadLock C_INCOME_TAX2, -1, C_INCOME_TAX2, -1   '소득세 
        ggoSpread.SpreadLock C_RES_TAX2, -1, C_RES_TAX2, -1      '주민세 
        ggoSpread.SpreadLock C_ANUT2, -1, C_ANUT2, -1         '국민연금 
        ggoSpread.SpreadLock C_MED_INSUR2, -1, C_MED_INSUR2, -1    '의료보험 
        ggoSpread.SpreadLock C_EMP_INSUR2, -1, C_EMP_INSUR2, -1    '고용보험 
        ggoSpread.SpreadLock C_SUB_TOT_AMT2, -1, C_SUB_TOT_AMT2, -1  '공제총액 
        ggoSpread.SpreadLock C_REAL_PROV_AMT2, -1, C_REAL_PROV_AMT2, -1 '실지급액 
        .vspdData2.ReDraw = True
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

           .MaxCols   = C_REAL_PROV_AMT + 1                                                      ' ☜:☜: Add 1 to Maxcols
	                                               ' ☜:☜: Add 1 to Maxcols
	       .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
           .ColHidden = True                                                            ' ☜:☜:

           .MaxRows = 0
            ggoSpread.ClearSpreadData

            Call GetSpreadColumnPos("A") 'sbk

            Call AppendNumberPlace("6","15","0")
            
            ggoSpread.SSSetEdit C_NAME           , "성명"     , 12,,,30,2		'Lock/ Edit
            ggoSpread.SSSetEdit C_EMP_NO         , "사번"     , 12,,,13,2		'Lock/ Edit
            ggoSpread.SSSetEdit C_DEPT_CD        , "부서명"   , 18,,,40,2		'Lock/ Edit
            ggoSpread.SSSetEdit C_PAY_CD         , "급여구분" , 12,,,50,2		'Lock/ Edit
            ggoSpread.SSSetEdit C_PROV_TYPE      , "지급구분" , 12,,,50,2		'Lock/ Edit
            ggoSpread.SSSetEdit C_PROV_TYPE_HIDDEN  , "지급구분code" , 12,,,1,2		'Lock/ Edit
           
            ggoSpread.SSSetFloat C_PAY_TOT_AMT   , "급여총액" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_INCOME_TAX    , "소득세"   ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_RES_TAX       , "주민세"   ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_ANUT          , "국민연금" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_MED_INSUR     , "의료보험" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_EMP_INSUR     , "고용보험" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_SUB_TOT_AMT   , "공제총액" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_REAL_PROV_AMT , "실지급액" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
               
	       .ReDraw = true

            Call ggoSpread.SSSetColHidden(C_PROV_TYPE_HIDDEN,C_PROV_TYPE_HIDDEN,True)
	       
        End With
    End If

    If pvSpdNo = "" OR pvSpdNo = "B" Then
        
        With frm1.vspdData2
	
            ggoSpread.Source = frm1.vspdData2

            ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    'sbk

	       .ReDraw = false
	
           .MaxCols   = C_REAL_PROV_AMT2 + 1                                                      ' ☜:☜: Add 1 to Maxcols
	                                               ' ☜:☜: Add 1 to Maxcols
	       .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
           .ColHidden = True                                                            ' ☜:☜:
            
           .MaxRows = 0
            ggoSpread.ClearSpreadData

           .DisplayColHeaders = False

            Call GetSpreadColumnPos("B") 'sbk

            Call AppendNumberPlace("6","15","0")

            ggoSpread.SSSetEdit C_NAME2           , ""     , 12,,,15,2		'Lock/ Edit
            ggoSpread.SSSetEdit C_EMP_NO2         , ""     , 12,,,15,2		'Lock/ Edit
            ggoSpread.SSSetEdit C_DEPT_CD2        , ""     , 18,,,15,2		'Lock/ Edit
            ggoSpread.SSSetEdit C_PAY_CD2         , ""     , 12,,,15,2		'Lock/ Edit
            ggoSpread.SSSetEdit C_PROV_TYPE2      , ""     , 12,,,15,2		'Lock/ Edit
            ggoSpread.SSSetEdit C_PROV_TYPE_HIDDEN2  , "지급구분code" , 12,,,15,2		'Lock/ Edit
            
            ggoSpread.SSSetFloat C_PAY_TOT_AMT2   , "급여총액" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_INCOME_TAX2    , "소득세"   ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_RES_TAX2       , "주민세"   ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_ANUT2          , "국민연금" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_MED_INSUR2     , "의료보험" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_EMP_INSUR2     , "고용보험" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_SUB_TOT_AMT2   , "공제총액" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_REAL_PROV_AMT2 , "실지급액" ,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
               
	       .ReDraw = true    

            Call ggoSpread.SSSetColHidden(C_PROV_TYPE_HIDDEN2,C_PROV_TYPE_HIDDEN2,True)
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
      
      ggoSpread.Source = frm1.vspdData2
      ggoSpread.SpreadLockWithOddEvenRowColor()
       
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
            
            C_NAME = iCurColumnPos(1) 		      '성명							'Spread Sheet의 Column별 상수 %>
            C_EMP_NO = iCurColumnPos(2)           '사번 
            C_DEPT_CD = iCurColumnPos(3)          '부서명 
            C_PAY_CD = iCurColumnPos(4)           '급여구분 
            C_PROV_TYPE = iCurColumnPos(5)        '지급구분 
            C_PROV_TYPE_HIDDEN = iCurColumnPos(6)        '지급구분 
            C_PAY_TOT_AMT = iCurColumnPos(7)      '급여총액 
            C_INCOME_TAX = iCurColumnPos(8)       '소득세 
            C_RES_TAX = iCurColumnPos(9)          '주민세 
            C_ANUT = iCurColumnPos(10)             '국민연금 
            C_MED_INSUR = iCurColumnPos(11)       '의료보험 
            C_EMP_INSUR = iCurColumnPos(12)       '고용보험 
            C_SUB_TOT_AMT = iCurColumnPos(13)     '공제총액 
            C_REAL_PROV_AMT = iCurColumnPos(14)   '실지급액 
        
       Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

            C_NAME2 = iCurColumnPos(1) 		       '성명						'Spread Sheet의 Column별 상수 %>
            C_EMP_NO2 = iCurColumnPos(2)           '사번 
            C_DEPT_CD2 = iCurColumnPos(3)          '부서명 
            C_PAY_CD2 = iCurColumnPos(4)           '급여구분 
            C_PROV_TYPE2 = iCurColumnPos(5)        '지급구분 
            C_PROV_TYPE_HIDDEN2 = iCurColumnPos(6) '지급구분 
            C_PAY_TOT_AMT2 = iCurColumnPos(7)      '급여총액 
            C_INCOME_TAX2 = iCurColumnPos(8)       '소득세 
            C_RES_TAX2 = iCurColumnPos(9)          '주민세 
            C_ANUT2 = iCurColumnPos(10)             '국민연금 
            C_MED_INSUR2 = iCurColumnPos(11)       '의료보험 
            C_EMP_INSUR2 = iCurColumnPos(12)       '고용보험 
            C_SUB_TOT_AMT2 = iCurColumnPos(13)     '공제총액 
            C_REAL_PROV_AMT2 = iCurColumnPos(14)   '실지급액 
            
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
    
    Call ggoOper.FormatDate(frm1.txtpay_yymm_dt, Parent.gDateFormat, 2)                    '싱글에서 년월말 입력하고 싶은경우 다음 함수를 콜한다.
    
    Call FuncGetAuth(gStrRequestMenuID, Parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
    
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
    Dim strWhere
    Dim StrDt
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

    Dim Fr_dept_cd , To_dept_cd, rFrDept ,rToDept
    
    Fr_dept_cd = frm1.txtFr_internal_cd.value
    To_dept_cd = frm1.txtTo_internal_cd.value
    
    Strdt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtpay_yymm_dt.Year, Right("0" & frm1.txtpay_yymm_dt.month , 2), "01")    
 
    If fr_dept_cd = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,Strdt, rFrDept ,rToDept)
		
		frm1.txtFr_internal_cd.value = rFrDept
		frm1.txtFr_dept_nm.value = ""
	End If	
	
	If to_dept_cd = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,Strdt, rFrDept ,rToDept)
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
    
    Call InitVariables 
    Call MakeKeyStream("X")
    Call DisableToolBar(Parent.TBC_QUERY)
	
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
	IF DBsave =  False Then
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
             .vspdData.Col = C_EMP_NO	 : strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep								
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
	IF DBdelete =  False Then
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
    Dim strVal

    lgIntFlgMode = Parent.OPMD_UMODE    
    ggoSpread.Source       = Frm1.vspdData2
    Frm1.vspdData2.MaxRows = 0
    ggoSpread.ClearSpreadData

    Call MakeKeyStream("X")

    If LayerShowHide(1) = False then
    	Exit Function 
    End if

    strVal = BIZ_PGM_ID1 & "?txtMode="            & Parent.UID_M0001                    '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & 1                             '☜: Max fetched data
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
    Call ggoOper.LockField(Document, "Q")										'⊙: Lock field
	Call SetToolbar("1100000000011111")	 

End Function
'========================================================================================================
' Function Name : DbQueryOk1
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk1()

	lgIntFlgMode      = Parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
	
    Frm1.vspdData2.Col = 0
    Frm1.vspdData2.Text = "합계"
    Frm1.txtpay_yymm_dt.focus 
	Call SetToolbar("1100000000011111")												'⊙: Set ToolBar

	Frm1.vspdData.focus
    Call ggoOper.LockField(Document, "Q")
    
    Set gActiveElement = document.ActiveElement
	frm1.vspdData.focus    
End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    
    Call InitVariables															'⊙: Initializes local global variables
	Call DisableToolBar(Parent.TBC_QUERY)
	IF DBQUERY =  False Then
		Call RestoreToolBar()
		Exit Function
	End If
End Function
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
End Function

'----------------------------------------  OpenEmptName()  ------------------------------------------
'	Name : OpenEmptName()                                                         <==== 성명/사번 팝업 
'	Description : Employee PopUp
'------------------------------------------------------------------------------------------------
Function OpenEmptName(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	Else 'spread
		arrParam(0) = frm1.vspdData.Text			' Code Condition
	End If
	arrParam(1) = ""'frm1.txtName.value			' Name Cindition
	arrParam(2) = lgUsrIntCd
	
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = 0 Then
			frm1.txtEmp_no.focus
		Else
			frm1.vspdData.Col = C_EmpNo
			frm1.vspdData.action =0
		End If	
		Exit Function
	Else
		Call SetEmp(arrRet, iWhere)
	End If	
			
End Function

'------------------------------------------  SetEmp()  ------------------------------------------------
'	Name : SetEmp()
'	Description : Employee Popup에서 Return되는 값 setting
'------------------------------------------------------------------------------------------------------
Function SetEmp(Byval arrRet, Byval iWhere)
		
	With frm1
		If iWhere = 0 Then
			.txtEmp_no.value = arrRet(0)
			.txtName.value = arrRet(1)
			.txtEmp_no.focus
		Else
			.vspdData.Col = C_EmpNm
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_EmpNo
			.vspdData.Text = arrRet(0)
			.vspdData.action =0
		End If
	End With
End Function

'========================================================================================================
' Name : OpenCondAreaPopup()        
' Desc : developer describe this line 
'========================================================================================================

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
	        arrParam(4) = " MAJOR_CD = " & FilterVar("H0040", "''", "S") & " AND MINOR_CD NOT IN (" & FilterVar("B", "''", "S") & " ," & FilterVar("C", "''", "S") & " )"    ' Where Condition							' Where Condition
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
		    Case "2"
		        .txtprov_cd.value = arrRet(0)
		        .txtprov_nm.value = arrRet(1)
		        .txtprov_cd.focus
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
        Dim Strdt

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then
		arrParam(0) = frm1.txtFr_dept_cd.value			            '  Code Condition
	ElseIf iWhere = 1 Then
		arrParam(0) = frm1.txtTo_dept_cd.value			            ' Code Condition
	End If
	
        Strdt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtpay_yymm_dt.Year, Right("0" & frm1.txtpay_yymm_dt.month , 2), "01")    
        
    arrParam(1) = Strdt
	arrParam(2) = lgUsrIntCd                    ' 자료권한 Condition  
	
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
'   Event Name : txtEmp_no_change             '<==인사마스터에 있는 사원인지 확인 
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

    If  frm1.txtEmp_no.value = "" Then
		frm1.txtName.value = ""
    Else
         IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
                
         If  IntRetCd < 0 then
            If  IntRetCd = -1 then
    	Call DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
			frm1.txtName.value = ""
            Frm1.txtEmp_no.focus 
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true
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

'========================================================================================================
'   Event Name : txtFr_dept_cd_Onchange
'   Event Desc :
'========================================================================================================
Function txtFr_dept_cd_Onchange()
    Dim IntRetCd
    Dim strDept_nm
    Dim Strdt    

    Strdt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtpay_yymm_dt.Year, Right("0" & frm1.txtpay_yymm_dt.month , 2), "01")    

    If frm1.txtFr_dept_cd.value = "" Then
		frm1.txtFr_dept_nm.value = ""
		frm1.txtFr_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtFr_dept_cd.value,Strdt,lgUsrIntCd, strDept_nm, lsInternal_cd)

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
    Dim Strdt

    Strdt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtpay_yymm_dt.Year, Right("0" & frm1.txtpay_yymm_dt.month , 2), "01")    
    
    If frm1.txtTo_dept_cd.value = "" Then
		frm1.txtTo_dept_nm.value = ""
		frm1.txtTo_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtTo_dept_cd.value,Strdt,lgUsrIntCd,strDept_nm,lsInternal_cd)
        
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

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
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
'   Event Name : vspdData_Click
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

'=======================================================================================================
'   Event Name : txtYear_DblClick(Button)
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

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST" >
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>급여조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* >&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
		
    <TR HEIGHT=*>
		<TD width=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR>
			        <TD <%=HEIGHT_TYPE_02%>width=100%></TD>
			    </TR>
				<TR>
					<TD HEIGHT=20>
					  <FIELDSET CLASS="CLSFLD">
					   <TABLE <%=LR_SPACE_TYPE_40%>>
						    <TR>
								<TD CLASS=TD5 NOWRAP>급여년월</TD>
								<TD CLASS="TD6" NOWRAP><OBJECT classid=<%=gCLSIDFPDT%> id=txtpay_yymm_dt NAME="txtpay_yymm_dt"  CLASS=FPDTYYYYMM  title=FPDATETIME  ALT="급여년월" tag="12X1" VIEWASTEXT></OBJECT></TD>		
							    <TD CLASS=TD5 NOWRAP>사원</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no" MAXLENGTH="13" SIZE="13" ALT ="사번" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenEmptName(0)">
								                 <INPUT NAME="txtName" MAXLENGTH="30" SIZE="20" ALT ="성명" tag="14XXXU"></TD>
							</TR>
	                        <TR>
	                        	<TD CLASS=TD5 NOWRAP>부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFr_dept_cd" ALT="부서코드" TYPE="Text" SiZE="10" MAXLENGTH="10" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(0)">
			                                       <INPUT NAME="txtFr_dept_nm" ALT="부서코드명" TYPE="Text" SiZE="20" MAXLENGTH="40" tag="14XXXU">&nbsp;~
		                                           <INPUT NAME="txtFr_Internal_cd" ALT="내부부서코드" TYPE="HIDDEN" SiZE="7" MAXLENGTH="7" tag="14XXXU">  
	                            <TD CLASS="TD5" NOWRAP>지급구분</TD>
	                        	<TD CLASS="TD6" NOWRAP><INPUT NAME="txtProv_cd" MAXLENGTH="1" SIZE="10" ALT ="지급구분" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCondAreaPopup(2)">
	                        	                       <INPUT NAME="txtProv_nm" MAXLENGTH="20" SIZE="20" ALT ="지급구분명" tag="14XXXU"></TD>
	                        </TR>
	                        <TR>	
	                        	<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtto_dept_cd" MAXLENGTH="10" SIZE="10" ALT ="Order ID" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(1)">
							                   <INPUT NAME="txtto_dept_nm" MAXLENGTH="40" SIZE="20" ALT ="Order ID" tag="14XXXU">
    			                            <INPUT  NAME="txtTo_Internal_cd" ALT="내부부서코드" TYPE="HIDDEN" SiZE="7" MAXLENGTH="7" tag="14XXXU"></TD>
    			                <TD CLASS="TD5" NOWRAP>급여구분</TD>
	                        	<TD CLASS="TD6" NOWRAP><SELECT Name="cboPay_cd" ALT="급여구분" STYLE="WIDTH: 133px" tag="11"><OPTION Value=""></OPTION></SELECT></TD>
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
							<TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=44 VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT=100%>
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
      <TD <%=HEIGHT_TYPE_01%>></TD>
    </TR>
	<TR HEIGHT=20>
	    <TD>
	        <TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
				<TD WIDTH="*" ALIGN=RIGHT><a href = "VBSCRIPT:PgmJumpCheck()" ONCLICK="VBSCRIPT:CookiePage 1">상세조회</a>
				</TD>
				<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
	    </TD>   
	
	</TR>
	<TR>
		<TD width=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TR>

</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioKind" tag="24">s
<INPUT TYPE=HIDDEN NAME="txtRadioType" tag="24">
<INPUT TYPE=HIDDEN NAME="txtCheck" tag="24">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
