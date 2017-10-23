<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 승급/승격등록 
*  3. Program ID           : H3004ma1
*  4. Program Name         : H3004ma1
*  5. Program Desc         : 근무이력관리/승급/승격등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/23
*  8. Modified date(Last)  : 2003/06/10
*  9. Modifier (First)     : YBI
* 10. Modifier (Last)      : Lee SiNa
* 13. Modified Comment     : 한줄추가하여 사번 선택시 급호와 직위가 조회되지 않는것 수정 
* 14. Comment              :
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID = "H3004mb1.asp"                                      'Biz Logic ASP
Const BIZ_PGM_ID1 = "H3004mb2.asp"
Const BIZ_PGM_JUMP_ID = "H2001ma1" 
Const CookieSplit = 1233

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

Dim C_EMP_NO 
Dim C_EMP_NO_POP
Dim C_NAME
Dim C_DEPT_CD 
Dim C_DEPT_NM
Dim C_PAY_GRD1 
Dim C_PAY_GRD1_NM 
Dim C_PAY_GRD2 
Dim C_ROLL_PSTN 
Dim C_ROLL_PSTN_NM 
Dim C_OCPT_TYPE 
Dim C_OCPT_TYPE_NM 
Dim C_FUNC_CD 
Dim C_FUNC_CD_NM 
Dim C_ROLE_CD
Dim C_ROLE_CD_NM 
Dim C_ENTR_DT 
Dim C_RESENT_PROMOTE_DT 
Dim C_CHNG_DEPT_CD
Dim C_CHNG_DEPT_NM
Dim C_CHNG_DEPT_CD_POP 
Dim C_CHNG_PAY_GRD1 
Dim C_CHNG_PAY_GRD1_NM 
Dim C_CHNG_PAY_GRD1_POP
Dim C_CHNG_PAY_GRD2 
Dim C_CHNG_ROLL_PSTN 
Dim C_CHNG_ROLL_PSTN_NM
Dim C_CHNG_ROLL_PSTN_POP 
Dim C_CHNG_OCPT_TYPE
Dim C_CHNG_OCPT_TYPE_NM 
Dim C_CHNG_OCPT_TYPE_POP
Dim C_CHNG_FUNC_CD
Dim C_CHNG_FUNC_CD_NM 
Dim C_CHNG_FUNC_CD_POP 
Dim C_CHNG_ROLE_CD 
Dim C_CHNG_ROLE_CD_NM 
Dim C_CHNG_ROLE_CD_POP 
Dim C_PROMOTE_DT 
Dim C_CHNG_CD 
Dim C_CHNG_CD_NM
'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
sub InitSpreadPosVariables()

	C_EMP_NO = 1													<%'Spread Sheet의 Column별 상수 %>
	C_EMP_NO_POP = 2	
	C_NAME = 3		
	C_DEPT_CD = 4
	C_DEPT_NM = 5
	C_PAY_GRD1 = 6
	C_PAY_GRD1_NM = 7
	C_PAY_GRD2 = 8
	C_ROLL_PSTN = 9
	C_ROLL_PSTN_NM = 10
	C_OCPT_TYPE = 11
	C_OCPT_TYPE_NM = 12
	C_FUNC_CD = 13
	C_FUNC_CD_NM = 14
	C_ROLE_CD = 15
	C_ROLE_CD_NM = 16
	C_ENTR_DT = 17
	C_RESENT_PROMOTE_DT = 18
	C_CHNG_DEPT_CD = 19
	C_CHNG_DEPT_NM = 20
	C_CHNG_DEPT_CD_POP = 21
	C_CHNG_PAY_GRD1 = 22
	C_CHNG_PAY_GRD1_NM = 23
	C_CHNG_PAY_GRD1_POP = 24
	C_CHNG_PAY_GRD2 = 25
	C_CHNG_ROLL_PSTN = 26
	C_CHNG_ROLL_PSTN_NM = 27
	C_CHNG_ROLL_PSTN_POP = 28
	C_CHNG_OCPT_TYPE = 29
	C_CHNG_OCPT_TYPE_NM = 30
	C_CHNG_OCPT_TYPE_POP = 31
	C_CHNG_FUNC_CD = 32
	C_CHNG_FUNC_CD_NM = 33
	C_CHNG_FUNC_CD_POP = 34
	C_CHNG_ROLE_CD = 35
	C_CHNG_ROLE_CD_NM = 36
	C_CHNG_ROLE_CD_POP = 37
	C_PROMOTE_DT = 38
	C_CHNG_CD = 39
	C_CHNG_CD_NM = 40
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
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay
	Call  ExtractDateFrom("<%=GetsvrDate%>", parent.gServerDateFormat ,  parent.gServerDateType ,strYear,strMonth,strDay)

	frm1.txtChng_cd.value = "99"
	frm1.txtChng_cd_nm.value = "승급/승격"
	
	frm1.txtResent_promote_dt.Year=strYear
	frm1.txtResent_promote_dt.Month=strMonth
	frm1.txtResent_promote_dt.Day=strDay
	
	frm1.txtPro_dt.Year=strYear
	frm1.txtPro_dt.Month=strMonth
	frm1.txtPro_dt.Day=strDay
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
	On Error Resume Next

	Const CookieSplit = 4877						
	Dim strTemp

	If flgs = 1 Then
		 WriteCookie CookieSplit , frm1.txtEmp_no.Value
	ElseIf flgs = 0 Then

		strTemp =  ReadCookie(CookieSplit)
		If strTemp = "" then Exit Function
			
		frm1.txtEmp_no.value =  strTemp

		If Err.number <> 0 Then
			Err.Clear
			 WriteCookie CookieSplit , ""
			Exit Function 
		End If

		 WriteCookie CookieSplit , ""
		
		Call MainQuery()
			
	End If
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)
    lgKeyStream       = Frm1.txtEmp_no.Value & parent.gColSep                                           'You Must append one character( parent.gColSep)
    lgKeyStream = lgKeyStream & Frm1.txtName.Value & parent.gColSep
End Sub        


'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr

    Call  CommonQueryRs(" minor_cd, minor_nm "," b_minor "," major_cd = " & FilterVar("H0001", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    Call  SetCombo2(frm1.txtPay_grd1, iCodeArr, iNameArr, Chr(11))

    Call  CommonQueryRs(" minor_cd, minor_nm "," b_minor "," major_cd = " & FilterVar("H0001", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    Call  SetCombo2(frm1.txtChng_pay_grd1, iCodeArr, iNameArr, Chr(11))

    Call  CommonQueryRs(" minor_cd, minor_nm "," b_minor "," major_cd = " & FilterVar("H0002", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    Call  SetCombo2(frm1.txtRoll_pstn, iCodeArr, iNameArr, Chr(11))

End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
    Call SetSpreadLock
    Call SetSpreadColor(1,Frm1.vspdData.MaxRows)
    ggoSpread.SSSetProtected	C_EMP_NO, 1,Frm1.vspdData.MaxRows
	
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()  
		
	With frm1.vspdData
        ggoSpread.Source = frm1.vspdData
   		ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread    

	    .ReDraw = false
        .MaxCols = C_CHNG_CD + 1	
	    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True
    
        .MaxRows = 0	
		ggoSpread.Source = Frm1.vspdData    
		ggoSpread.ClearSpreadData         
		Call GetSpreadColumnPos("A")

         ggoSpread.SSSetEdit   C_EMP_NO,           "사번",      13,,,13,2
         ggoSpread.SSSetButton C_EMP_NO_POP                  
         ggoSpread.SSSetEdit   C_NAME,             "성명",      08,,,08,2
         ggoSpread.SSSetEdit   C_DEPT_CD,          "부서",      05,,,15,2
         ggoSpread.SSSetEdit   C_DEPT_NM,          "부서",      10,,,15,2
         ggoSpread.SSSetEdit   C_PAY_GRD1,         "급호",      05,,,15,2
         ggoSpread.SSSetEdit   C_PAY_GRD1_NM,      "급호",      10,,,15,2
         ggoSpread.SSSetEdit   C_PAY_GRD2,         "호봉",      06,,,15,2
         ggoSpread.SSSetEdit   C_ROLL_PSTN,        "직위",      05,,,15,2
         ggoSpread.SSSetEdit   C_ROLL_PSTN_NM,     "직위",      10,,,15,2
         ggoSpread.SSSetEdit   C_OCPT_TYPE,        "직종",      05,,,15,2
         ggoSpread.SSSetEdit   C_OCPT_TYPE_NM,     "직종",      10,,,15,2
         ggoSpread.SSSetEdit   C_FUNC_CD,          "직무",      05,,,15,2
         ggoSpread.SSSetEdit   C_FUNC_CD_NM,       "직무",      10,,,15,2
         ggoSpread.SSSetEdit   C_ROLE_CD,          "직책",      05,,,15,2
         ggoSpread.SSSetEdit   C_ROLE_CD_NM,       "직책",      10,,,15,2
         ggoSpread.SSSetDate   C_ENTR_DT,          "입사일",    10,2,  parent.gDateFormat
         ggoSpread.SSSetDate   C_RESENT_PROMOTE_DT,"최근승급일",14,2,  parent.gDateFormat
         ggoSpread.SSSetEdit   C_CHNG_DEPT_CD,     "변동부서",  05,,,20,2
         ggoSpread.SSSetEdit   C_CHNG_DEPT_NM,     "변동부서",  10,,,40,2
         ggoSpread.SSSetButton C_CHNG_DEPT_CD_POP
         ggoSpread.SSSetEdit   C_CHNG_PAY_GRD1,    "변동급호",  05,,,15,2
         ggoSpread.SSSetEdit   C_CHNG_PAY_GRD1_NM, "변동급호",  10,,,40,2
         ggoSpread.SSSetButton C_CHNG_PAY_GRD1_POP
         ggoSpread.SSSetEdit   C_CHNG_PAY_GRD2,    "변동호봉",  10,,,15,2
         ggoSpread.SSSetEdit   C_CHNG_ROLL_PSTN,   "변동직위",  05,,,15,2
         ggoSpread.SSSetEdit   C_CHNG_ROLL_PSTN_NM,"변동직위",  10,,,40,2
         ggoSpread.SSSetButton C_CHNG_ROLL_PSTN_POP
         ggoSpread.SSSetEdit   C_CHNG_OCPT_TYPE,   "직종",      05,,,15,2
         ggoSpread.SSSetEdit   C_CHNG_OCPT_TYPE_NM,"변동직종",  10,,,40,2
         ggoSpread.SSSetButton C_CHNG_OCPT_TYPE_POP
         ggoSpread.SSSetEdit   C_CHNG_FUNC_CD,     "직무",      05,,,15,2
         ggoSpread.SSSetEdit   C_CHNG_FUNC_CD_NM,  "변동직무",  10,,,40,2
         ggoSpread.SSSetButton C_CHNG_FUNC_CD_POP
         ggoSpread.SSSetEdit   C_CHNG_ROLE_CD,     "직책",      05,,,15,2
         ggoSpread.SSSetEdit   C_CHNG_ROLE_CD_NM,  "변동직책",  10,,,40,2
         ggoSpread.SSSetButton C_CHNG_ROLE_CD_POP
         ggoSpread.SSSetEdit   C_PROMOTE_DT,       "승급예정일",14,,,15,2
         ggoSpread.SSSetEdit   C_CHNG_CD,          "변동",  10,,,15,2
         ggoSpread.SSSetEdit   C_CHNG_CD_NM,        "변동사유",  10,,,15,2         

		call ggoSpread.MakePairsColumn(C_NAME,C_EMP_NO_POP)
		call ggoSpread.MakePairsColumn(C_CHNG_DEPT_NM,C_CHNG_DEPT_CD_POP)
		call ggoSpread.MakePairsColumn(C_CHNG_PAY_GRD1_NM,C_CHNG_PAY_GRD1_POP)
		call ggoSpread.MakePairsColumn(C_CHNG_ROLL_PSTN_NM,C_CHNG_ROLL_PSTN_POP)
		call ggoSpread.MakePairsColumn(C_CHNG_OCPT_TYPE_NM,C_CHNG_OCPT_TYPE_POP)
		call ggoSpread.MakePairsColumn(C_CHNG_FUNC_CD_NM,C_CHNG_FUNC_CD_POP)
		call ggoSpread.MakePairsColumn(C_CHNG_ROLE_CD_NM,C_CHNG_ROLE_CD_POP)
                            
        Call ggoSpread.SSSetColHidden(C_DEPT_CD,C_DEPT_CD,True)	
        Call ggoSpread.SSSetColHidden(C_PAY_GRD1,C_PAY_GRD1,True)	
        Call ggoSpread.SSSetColHidden(C_ROLL_PSTN,C_ROLL_PSTN,True)	        
        Call ggoSpread.SSSetColHidden(C_OCPT_TYPE,C_OCPT_TYPE,True)	
        Call ggoSpread.SSSetColHidden(C_FUNC_CD,C_FUNC_CD,True)	
        Call ggoSpread.SSSetColHidden(C_ROLE_CD,C_ROLE_CD,True)	        
        Call ggoSpread.SSSetColHidden(C_CHNG_DEPT_CD,C_CHNG_DEPT_CD,True)	
        Call ggoSpread.SSSetColHidden(C_CHNG_PAY_GRD1,C_CHNG_PAY_GRD1,True)	
        Call ggoSpread.SSSetColHidden(C_CHNG_ROLL_PSTN,C_CHNG_ROLL_PSTN,True)	        
        Call ggoSpread.SSSetColHidden(C_CHNG_OCPT_TYPE,C_CHNG_OCPT_TYPE,True)	
        Call ggoSpread.SSSetColHidden(C_CHNG_FUNC_CD,C_CHNG_FUNC_CD,True)	
        Call ggoSpread.SSSetColHidden(C_CHNG_ROLE_CD,C_CHNG_ROLE_CD,True)	        
        Call ggoSpread.SSSetColHidden(C_CHNG_CD,C_CHNG_CD,True)
	   .ReDraw = true
       Call SetSpreadLock 
    End With
    
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
			C_DEPT_CD = iCurColumnPos(4)
			C_DEPT_NM = iCurColumnPos(5)
			C_PAY_GRD1 = iCurColumnPos(6)
			C_PAY_GRD1_NM = iCurColumnPos(7)
			C_PAY_GRD2 = iCurColumnPos(8)
			C_ROLL_PSTN = iCurColumnPos(9)
			C_ROLL_PSTN_NM = iCurColumnPos(10)
			C_OCPT_TYPE = iCurColumnPos(11)
			C_OCPT_TYPE_NM = iCurColumnPos(12)
			C_FUNC_CD = iCurColumnPos(13)
			C_FUNC_CD_NM = iCurColumnPos(14)
			C_ROLE_CD = iCurColumnPos(15)
			C_ROLE_CD_NM = iCurColumnPos(16)
			C_ENTR_DT = iCurColumnPos(17)
			C_RESENT_PROMOTE_DT = iCurColumnPos(18)
			C_CHNG_DEPT_CD = iCurColumnPos(19)
			C_CHNG_DEPT_NM = iCurColumnPos(20)
			C_CHNG_DEPT_CD_POP = iCurColumnPos(21)
			C_CHNG_PAY_GRD1 = iCurColumnPos(22)
			C_CHNG_PAY_GRD1_NM = iCurColumnPos(23)
			C_CHNG_PAY_GRD1_POP = iCurColumnPos(24)
			C_CHNG_PAY_GRD2 = iCurColumnPos(25)
			C_CHNG_ROLL_PSTN = iCurColumnPos(26)
			C_CHNG_ROLL_PSTN_NM = iCurColumnPos(27)
			C_CHNG_ROLL_PSTN_POP = iCurColumnPos(28)
			C_CHNG_OCPT_TYPE = iCurColumnPos(29)
			C_CHNG_OCPT_TYPE_NM = iCurColumnPos(30)
			C_CHNG_OCPT_TYPE_POP = iCurColumnPos(31)
			C_CHNG_FUNC_CD = iCurColumnPos(32)
			C_CHNG_FUNC_CD_NM = iCurColumnPos(33)
			C_CHNG_FUNC_CD_POP = iCurColumnPos(34)
			C_CHNG_ROLE_CD = iCurColumnPos(35)
			C_CHNG_ROLE_CD_NM = iCurColumnPos(36)
			C_CHNG_ROLE_CD_POP = iCurColumnPos(37)
			C_PROMOTE_DT = iCurColumnPos(38)
			C_CHNG_CD = iCurColumnPos(39)
    End Select    
End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
    .vspdData.ReDraw = False
     ggoSpread.SpreadLock C_NAME, -1,C_NAME, -1
     ggoSpread.SpreadLock C_EMP_NO, -1,C_EMP_NO, -1
     ggoSpread.SpreadLock C_EMP_NO_POP, -1,C_EMP_NO_POP, -1
    
     ggoSpread.SpreadLock C_DEPT_CD, -1,C_DEPT_CD, -1
     ggoSpread.SpreadLock C_PAY_GRD1, -1,C_PAY_GRD1, -1
     ggoSpread.SpreadLock C_PAY_GRD2, -1,C_PAY_GRD2, -1
     ggoSpread.SpreadLock C_ROLL_PSTN, -1,C_ROLL_PSTN, -1
     ggoSpread.SpreadLock C_OCPT_TYPE, -1,C_OCPT_TYPE, -1
     ggoSpread.SpreadLock C_FUNC_CD, -1,C_FUNC_CD, -1
     ggoSpread.SpreadLock C_ROLE_CD, -1,C_ROLE_CD, -1
    
     ggoSpread.SpreadLock C_ENTR_DT, -1,C_ENTR_DT, -1
     ggoSpread.SpreadLock C_RESENT_PROMOTE_DT, -1,C_RESENT_PROMOTE_DT, -1
    
     ggoSpread.SpreadLock C_PROMOTE_DT, -1,C_PROMOTE_DT, -1
     ggoSpread.SpreadLock C_CHNG_CD, -1,C_CHNG_CD, -1
	ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1        
    .vspdData.ReDraw = True
    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
     ggoSpread.SSSetProtected	C_NAME, pvStartRow, pvEndRow
     ggoSpread.SSSetProtected	C_DEPT_CD, pvStartRow, pvEndRow
     ggoSpread.SSSetProtected	C_DEPT_NM, pvStartRow, pvEndRow
     ggoSpread.SSSetProtected	C_PAY_GRD1,pvStartRow, pvEndRow
     ggoSpread.SSSetProtected	C_PAY_GRD1_NM, pvStartRow, pvEndRow
     ggoSpread.SSSetProtected	C_PAY_GRD2, pvStartRow, pvEndRow
     ggoSpread.SSSetProtected	C_ROLL_PSTN, pvStartRow, pvEndRow
     ggoSpread.SSSetProtected	C_ROLL_PSTN_NM, pvStartRow, pvEndRow
     ggoSpread.SSSetProtected	C_OCPT_TYPE, pvStartRow, pvEndRow
     ggoSpread.SSSetProtected	C_OCPT_TYPE_NM, pvStartRow, pvEndRow
     ggoSpread.SSSetProtected	C_FUNC_CD, pvStartRow, pvEndRow
     ggoSpread.SSSetProtected	C_FUNC_CD_NM, pvStartRow, pvEndRow
     ggoSpread.SSSetProtected	C_ROLE_CD, pvStartRow, pvEndRow
     ggoSpread.SSSetProtected	C_ROLE_CD_NM, pvStartRow, pvEndRow
     ggoSpread.SSSetProtected	C_ENTR_DT, lpvStartRow, pvEndRow
     ggoSpread.SSSetProtected	C_RESENT_PROMOTE_DT, pvStartRow, pvEndRow
     ggoSpread.SSSetProtected	C_CHNG_CD_NM, pvStartRow, pvEndRow

     ggoSpread.SSSetProtected	C_PROMOTE_DT, pvStartRow, pvEndRow
   
     ggoSpread.SSSetRequired		C_EMP_NO, pvStartRow, pvEndRow
     ggoSpread.SSSetRequired		C_CHNG_DEPT_NM, pvStartRow, pvEndRow
     ggoSpread.SSSetRequired		C_CHNG_PAY_GRD1_NM, pvStartRow, pvEndRow
     ggoSpread.SSSetRequired		C_CHNG_PAY_GRD2, pvStartRow, pvEndRow
     ggoSpread.SSSetRequired		C_CHNG_ROLL_PSTN_NM, pvStartRow, pvEndRow
     ggoSpread.SSSetRequired		C_CHNG_OCPT_TYPE_NM, pvStartRow, pvEndRow
     ggoSpread.SSSetRequired		C_CHNG_FUNC_CD_NM, pvStartRow, pvEndRow
     ggoSpread.SSSetRequired		C_CHNG_ROLE_CD_NM, pvStartRow, pvEndRow
     ggoSpread.SSSetProtected	C_CHNG_CD, pvStartRow, pvEndRow
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

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear       
                                                                    '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
    
    Call  ggoOper.FormatField(Document, "1", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	Call  ggoOper.LockField(Document, "N")											'⊙: Lock Field
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables

    Call  FuncGetAuth(gStrRequestMenuID,  parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
    Call SetDefaultVal
    Call InitComboBox
    Call SetToolbar("1000111100001111")										        '버튼 툴바 제어 
    frm1.txtDept_cd.Focus

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

    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

     ggoSpread.Source = Frm1.vspdData
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData     															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("X")
    Call  DisableToolBar( parent.TBC_QUERY)

	If DBQuery = False Then
		Call  RestoreToolBar()
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
    Dim lRow
    Dim strAdmi_dt
    Dim strGrudt_dt

    Dim strPay_grd1
    Dim strPay_grd2
    Dim strPromote_dt
    Dim strSQL
    FncSave = False                                                              '☜: Processing is NG
  
    Err.Clear                                                                    '☜: Clear err status
       
     ggoSpread.Source = frm1.vspdData
    If  ggoSpread.SSCheckChange = False Then
        IntRetCD =  DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If

	With Frm1     
			     
		If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
			   Set gActiveElement = document.ActiveElement   
		       Exit Function
		End If

       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0

            if  .vspdData.Text =  ggoSpread.InsertFlag or .vspdData.Text =  ggoSpread.UpdateFlag then
                .vspdData.Row = lRow
                .vspdData.Col = C_EMP_NO
                if Trim(.vspdData.Text)=""	then
					.vspdData.Col = C_NAME
					.vspdData.Text = ""
				end if   

                .vspdData.Col = C_CHNG_PAY_GRD1
                 strPay_grd1 = .vspdData.Text
                .vspdData.Col = C_CHNG_PAY_GRD2
                 strPay_grd2 = .vspdData.Text
  
                .vspdData.Col = C_PROMOTE_DT
             
                 strPromote_dt =  UNIConvDateCompanyToDB(.vspdData.Text,  parent.gDateFormat)
                
                 strSQL =              " pay_grd1 =  " & FilterVar(strPay_grd1 , "''", "S") & ""
                 strSQL = strSQL & " AND pay_grd2 =  " & FilterVar(strPay_grd2 , "''", "S") & ""
                 strSQL = strSQL & " AND apply_strt_dt = "
                 strSQL = strSQL & " (SELECT MAX(apply_strt_dt) FROM hdf010t WHERE apply_strt_dt <=  " & FilterVar(strPromote_dt , "''", "S") & " "
				 strSQL = strSQL & " AND pay_grd1 =  " & FilterVar(strPay_grd1 , "''", "S") & ""
				 strSQL = strSQL & " AND pay_grd2 =  " & FilterVar(strPay_grd2 , "''", "S") & ")"
				 
                .vspdData.Col = C_EMP_NO
				IntRetCD =  CommonQueryRs(" emp_no "," HAA010T "," emp_no= " & FilterVar(.vspdData.Text, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				if  IntRetCD = False   then
					Call  DisplayMsgBox("971001","X","사번","X")
                    .vspdData.Action = 0 ' go to 
                    Exit Function
                end if
                .vspdData.Col = C_CHNG_DEPT_NM
                
				IntRetCD =  CommonQueryRs(" DEPT_CD "," B_ACCT_DEPT "," DEPT_NM= " & FilterVar(.vspdData.Text, "''", "S") & " and ORG_CHANGE_DT = (SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT WHERE ORG_CHANGE_DT < getdate())",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				if  IntRetCD = False   then
					Call  DisplayMsgBox("971001","X","변동부서","X")
                    .vspdData.Col = C_CHNG_DEPT
                    .vspdData.Action = 0 ' go to 
                    Exit Function
                end if 
                
                .vspdData.Col = C_CHNG_PAY_GRD1_NM
				IntRetCD =  CommonQueryRs(" minor_cd, minor_nm "," b_minor "," major_cd = " & FilterVar("H0001", "''", "S") & " and minor_nm= " & FilterVar(.vspdData.Text, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				if  IntRetCD = False   then
					Call  DisplayMsgBox("971001","X","변동급호","X")
                    .vspdData.Col = C_CHNG_PAY_GRD1
                    .vspdData.Action = 0 ' go to 
                    Exit Function
                end if
                 IntRetCD =  CommonQueryRs(" COUNT(*) "," HDF010T ", strSQL,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
                if Replace(lgF0, Chr(11), "") = 0 then
                    Call  DisplayMsgBox("800057","X","X","X")
                    .vspdData.Col = C_CHNG_PAY_GRD2
                    .vspdData.Action = 0 ' go to 
                    Exit Function
                end if
                
                .vspdData.Col = C_CHNG_ROLL_PSTN_NM
				IntRetCD =  CommonQueryRs(" minor_cd, minor_nm "," b_minor "," major_cd = " & FilterVar("H0002", "''", "S") & " and minor_nm= " & FilterVar(.vspdData.Text, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				if  IntRetCD = False   then
					Call  DisplayMsgBox("971001","X","변동직위","X")				
                    .vspdData.Col = C_CHNG_ROLL_PSTN
                    .vspdData.Action = 0 ' go to 
                    Exit Function
                end if

                .vspdData.Col = C_CHNG_OCPT_TYPE_NM
				IntRetCD =  CommonQueryRs(" minor_cd, minor_nm "," b_minor "," major_cd = " & FilterVar("H0003", "''", "S") & " and minor_nm= " & FilterVar(.vspdData.Text, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				if  IntRetCD = False   then
					Call  DisplayMsgBox("971001","X","변동직종","X")				
                    .vspdData.Col = C_CHNG_OCPT_TYPE
                    .vspdData.Action = 0 ' go to 
                    Exit Function
                end if
                
                .vspdData.Col = C_CHNG_FUNC_CD_NM
				IntRetCD =  CommonQueryRs(" minor_cd, minor_nm "," b_minor "," major_cd = " & FilterVar("H0004", "''", "S") & " and minor_nm= " & FilterVar(.vspdData.Text, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				if  IntRetCD = False   then
					Call  DisplayMsgBox("971001","X","변동직무","X")				
                    .vspdData.Col = C_CHNG_FUNC_CD
                    .vspdData.Action = 0 ' go to 
                    Exit Function
                end if

                .vspdData.Col = C_CHNG_ROLE_CD_NM
				IntRetCD =  CommonQueryRs(" minor_cd, minor_nm "," b_minor "," major_cd = " & FilterVar("H0026", "''", "S") & " and minor_nm= " & FilterVar(.vspdData.Text, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				if  IntRetCD = False   then
					Call  DisplayMsgBox("971001","X","변동직책","X")				
                    .vspdData.Col = C_CHNG_ROLE_CD
                    .vspdData.Action = 0 ' go to 
                    Exit Function
                end if
                                                          
           End if
       Next
	End With

    Call  DisableToolBar( parent.TBC_SAVE)
 
    If DbSave = False Then
		Call  RestoreToolBar()
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
			SetSpreadColor .ActiveRow	, .ActiveRow
    
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
     ggoSpread.Source = Frm1.vspdData	
     ggoSpread.EditUndo  
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
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
         ggoSpread.InsertRow ,imRow
        
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.Row = .vspdData.ActiveRow
		.vspdData.Col = C_PROMOTE_DT
		.vspdData.Value = .txtPro_dt.text
  
		.vspdData.Col = C_CHNG_CD
		.vspdData.Text = .txtChng_cd.value
		
		.vspdData.Col = C_CHNG_CD_NM
		.vspdData.text = frm1.txtChng_cd_nm.value        
       .vspdData.ReDraw = True
    End With
    Set gActiveElement = document.ActiveElement   
End Function
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
    Call parent.FncExport( parent.C_MULTI)                                         '☜: 화면 유형 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind( parent.C_MULTI, False)                                    '☜:화면 유형, Tab 유무 
End Function

'========================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()
    Dim IntRetCD
	
	FncExit = False
	
     ggoSpread.Source = frm1.vspdData	
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900016",  parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
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

	If LayerShowHide(1)=False Then
		Exit Function
	End if
	
	Dim strVal
    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
    End With
		
    If lgIntFlgMode =  parent.OPMD_UMODE Then
    Else
    End If

	Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic
    
    DbQuery = True
    
End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
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

	Dim strRes_no

    DbSave = False                                                          
    
	If LayerShowHide(1)=False Then
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
 
               Case  ggoSpread.InsertFlag                                      '☜: Insert
                                                          strVal = strVal & "C" & parent.gColSep
                                                          strVal = strVal & lRow & parent.gColSep
                    .vspdData.Col = C_EMP_NO	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PROMOTE_DT        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_CD           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DEPT_CD	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PAY_GRD1	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PAY_GRD2	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ROLL_PSTN	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_OCPT_TYPE         : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_FUNC_CD           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ROLE_CD           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ENTR_DT           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RESENT_PROMOTE_DT : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_DEPT_CD	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_PAY_GRD1	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_PAY_GRD2	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_ROLL_PSTN	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_OCPT_TYPE    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_FUNC_CD      : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_ROLE_CD      : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                  strVal = strVal & "U" & parent.gColSep
                                                  strVal = strVal & lRow & parent.gColSep
                                                          strVal = strVal & lRow & parent.gColSep
                    .vspdData.Col = C_EMP_NO	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PROMOTE_DT        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_CD           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DEPT_CD	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PAY_GRD1	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PAY_GRD2	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ROLL_PSTN	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_OCPT_TYPE         : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_FUNC_CD           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ROLE_CD           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ENTR_DT           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RESENT_PROMOTE_DT : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_DEPT_CD	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_PAY_GRD1	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_PAY_GRD2	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_ROLL_PSTN	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_OCPT_TYPE    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_FUNC_CD      : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_ROLE_CD      : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.DeleteFlag                                      '☜: Delete
                                                  strDel = strDel & "D" & parent.gColSep
                                                  strDel = strDel & lRow & parent.gColSep
                    .vspdData.Col = C_EMP_NO	: strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PROMOTE_DT: strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_CD   : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	
       .txtMode.value        =  parent.UID_M0002
       .txtUpdtUserId.value  =  parent.gUsrID
       .txtInsrtUserId.value =  parent.gUsrID
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal

	End With
	
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
	
    DbSave = True                                                           
    
End Function

'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
    Dim IntRetCd
    
    FncDelete = False                                                      '⊙: Processing is NG
    
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call  DisplayMsgBox("900002","X","X","X")                                '☆:
        Exit Function
    End If
    
    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If
    
    Call  DisableToolBar( parent.TBC_DELETE)
    If DbDelete= False Then
    	Call  RestoreToolBar()
        Exit Function
    End If
    
    FncDelete = True                                                        '⊙: Processing is OK


End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
	
    lgIntFlgMode =  parent.OPMD_UMODE    
    Call  ggoOper.LockField(Document, "Q")										'⊙: Lock field
    Call InitData()
	Call SetToolbar("100011110011111")									

End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData     
    Call InitVariables															'⊙: Initializes local global variables
    
    Call  DisplayMsgBox("800485","X","X","X")	'결과는 "승급/승격조회"에서 확인하십시요 
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
Function OpenEmpName(Row)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = C_EMP_NO
   
    arrParam(0) = frm1.vspdData.Text
    arrParam(1) = ""
    arrParam(2) = lgUsrIntCd
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
	   	Frm1.vspdData.Col = C_EMP_NO
	   	frm1.vspdData.focus	
		Exit Function
	Else
		Call SetEmpName(arrRet)
	End If	
			
End Function

'======================================================================================================
'	Name : SetEmp()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SetEmpName(arrRet)
	With frm1
	   	frm1.vspdData.Col = C_NAME
		frm1.vspdData.Text = arrRet(1)
	   	Frm1.vspdData.Col = C_EMP_NO
	   	frm1.vspdData.Text = arrRet(0)		
	   	frm1.vspdData.focus
		lgBlnFlgChgValue = False
	End With
End Sub



'========================================================================================================
' Name : OpenDept
' Desc : 부서 POPUP
'========================================================================================================
Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(3)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = frm1.txtDept_cd.value			<%' 조건부에서 누른 경우 Code Condition%>
		arrParam(3) = ""		
	Else 'spread
		arrParam(0) = ""	
		frm1.vspdData.col = C_CHNG_DEPT_NM
		arrParam(3) = frm1.vspdData.text								<%' Name Cindition%>		
	End If
	arrParam(2) = lgUsrIntCd
	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = 0 Then 
			frm1.txtDept_cd.focus
		Else
			frm1.vspdData.Col = C_CHNG_DEPT_NM
			frm1.vspdData.action =0
		End If	
		Exit Function
	Else
		Call SetDept(arrRet, iWhere)
	End If	
			
End Function

'========================================================================================================
' Name : OpenPromoteDt
' Desc : 최근 승급일 POPUP
'========================================================================================================
Function OpenPromoteDt(iWhere)
	Dim arrRet
	Dim arrParam(1)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtResent_promote_dt.text			
	arrParam(1) = 1 'haa010t
	arrRet = window.showModalDialog(HRAskPRAspName("PromoteDtPopup"), Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	if arrRet(0) <> ""	then
		frm1.txtResent_promote_dt.text = arrRet(0)
	end if
	frm1.txtResent_promote_dt.focus
End Function

'------------------------------------------  SetDept()  ------------------------------------------------
'	Name : SetDept()
'	Description : Dept Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetDept(Byval arrRet, Byval iWhere)
		
	With frm1
		If iWhere = 0 Then 'TextBox(Condition)
			.txtDept_cd.value = arrRet(0)
			.txtDept_Nm.value = arrRet(1)
			.txtDept_cd.focus
		Else 'spread
			.vspdData.Col = C_CHNG_DEPT_CD
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_CHNG_DEPT_NM
			.vspdData.Text = arrRet(1)
			.vspdData.action =0
		End If
	End With
End Function

'======================================================================================================
'	Name : OpenCode()
'	Description : Code PopUp
'=======================================================================================================
Function OpenCode(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	Select Case iWhere
	    Case C_CHNG_DEPT_CD_POP
	        arrParam(0) = "전문대학교 팝업"			        <%' 팝업 명칭 %>
	    	arrParam(1) = "B_minor"							    <%' TABLE 명칭 %>
	    	arrParam(2) = strCode                   			<%' Code Condition%>
	    	arrParam(3) = ""									<%' Name Cindition%>
	    	arrParam(4) = "major_cd=" & FilterVar("H0008", "''", "S") & ""			    	<%' Where Condition%>
	    	arrParam(5) = "전문대학교코드" 			        <%' TextBox 명칭 %>
	
	    	arrField(0) = "minor_cd"							<%' Field명(0)%>
	    	arrField(1) = "minor_nm"    						<%' Field명(1)%>
    
	    	arrHeader(0) = "전문대학교코드"	   		    	<%' Header명(0)%>
	    	arrHeader(1) = "학교명"	    		     		<%' Header명(1)%>
	    Case C_CHNG_PAY_GRD1_POP
	        arrParam(0) = "급호코드 팝업"				    <%' 팝업 명칭 %>
	    	arrParam(1) = "B_minor"							    <%' TABLE 명칭 %>

	    	arrParam(2) = ""								  	<%' Code Condition%>
			frm1.vspdData.col = C_CHNG_PAY_GRD1_NM			    		    	
	    	arrParam(3) = frm1.vspdData.text					<%' Name Cindition%>
	    	arrParam(4) = "major_cd=" & FilterVar("H0001", "''", "S") & ""			    	<%' Where Condition%>
	    	arrParam(5) = "급호코드" 				    <%' TextBox 명칭 %>
	
	    	arrField(0) = "minor_cd"							<%' Field명(0)%>
	    	arrField(1) = "minor_nm"    						<%' Field명(1)%>
    
	    	arrHeader(0) = "급호코드"	    	    	<%' Header명(0)%>
	    	arrHeader(1) = "급호명"	    		     		<%' Header명(1)%>
	    Case C_CHNG_ROLL_PSTN_POP
	        arrParam(0) = "직위코드 팝업"				        <%' 팝업 명칭 %>
	    	arrParam(1) = "B_minor"							    <%' TABLE 명칭 %>
	    	arrParam(2) = ""									<%' code Cindition%>	    	
			frm1.vspdData.col = C_CHNG_ROLL_PSTN_NM 
	    	arrParam(3) = frm1.vspdData.text               			<%' NAME Condition%>

	    	arrParam(4) = "major_cd=" & FilterVar("H0002", "''", "S") & ""			    	<%' Where Condition%>
	    	arrParam(5) = "직위코드" 				        <%' TextBox 명칭 %>
	
	    	arrField(0) = "minor_cd"							<%' Field명(0)%>
	    	arrField(1) = "minor_nm"    						<%' Field명(1)%>
    
	    	arrHeader(0) = "직위코드"	       		    	<%' Header명(0)%>
	    	arrHeader(1) = "직위명"	    		     		<%' Header명(1)%>
	    Case C_CHNG_OCPT_TYPE_POP
	        arrParam(0) = "직종코드 팝업"	   			    <%' 팝업 명칭 %>
	    	arrParam(1) = "B_minor"							    <%' TABLE 명칭 %>
			arrParam(2) = ""								<%' Code Condition%>	    	
	    	frm1.vspdData.col = C_CHNG_OCPT_TYPE_NM					<%' Name Cindition%>
	    	arrParam(3) = frm1.vspdData.text                   			
	    	arrParam(4) = "major_cd=" & FilterVar("H0003", "''", "S") & ""			    	<%' Where Condition%>
	    	arrParam(5) = "직종코드" 					    <%' TextBox 명칭 %>
	
	    	arrField(0) = "minor_cd"							<%' Field명(0)%>
	    	arrField(1) = "minor_nm"    						<%' Field명(1)%>
    
	    	arrHeader(0) = "직종코드"	       			<%' Header명(0)%>
	    	arrHeader(1) = "직종명"	        				<%' Header명(1)%>
	    Case C_CHNG_FUNC_CD_POP
	        arrParam(0) = "직무코드 팝업"	    			    <%' 팝업 명칭 %>
	    	arrParam(1) = "B_minor"							    <%' TABLE 명칭 %>
	    	arrParam(2) = ""                   			<%' Code Condition%>
	    	frm1.vspdData.col = C_CHNG_FUNC_CD_NM					<%' Name Cindition%>
	    	arrParam(3) = frm1.vspdData.text
	    	arrParam(4) = "major_cd=" & FilterVar("H0004", "''", "S") & ""			    	<%' Where Condition%>
	    	arrParam(5) = "직무코드" 					    <%' TextBox 명칭 %>
	
	    	arrField(0) = "minor_cd"							<%' Field명(0)%>
	    	arrField(1) = "minor_nm"    						<%' Field명(1)%>
    
	    	arrHeader(0) = "직무코드"		       			<%' Header명(0)%>
	    	arrHeader(1) = "직무명"	        				<%' Header명(1)%>
	    Case C_CHNG_ROLE_CD_POP
	        arrParam(0) = "직책코드 팝업"	    			    <%' 팝업 명칭 %>
	    	arrParam(1) = "B_minor"							    <%' TABLE 명칭 %>
	    	arrParam(2) = ""									<%' Code Cindition%>
			frm1.vspdData.col = C_CHNG_ROLE_CD_NM			    	
	    	arrParam(3) = frm1.vspdData.text                    <%' Name Condition%>
	    	arrParam(4) = "major_cd=" & FilterVar("H0026", "''", "S") & ""			    	<%' Where Condition%>
	    	arrParam(5) = "직책코드" 					    <%' TextBox 명칭 %>
	
	    	arrField(0) = "minor_cd"							<%' Field명(0)%>
	    	arrField(1) = "minor_nm"    						<%' Field명(1)%>
    
	    	arrHeader(0) = "직책코드"		       			<%' Header명(0)%>
	    	arrHeader(1) = "직책명"	        				<%' Header명(1)%>
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
		    Case C_CHNG_DEPT_CD_POP
		    	frm1.vspdData.Col = C_CHNG_DEPT_NM
		    	frm1.vspdData.action =0

		    Case C_CHNG_PAY_GRD1_POP
		    	frm1.vspdData.Col = C_CHNG_PAY_GRD1_NM
				frm1.vspdData.action =0
				
		    Case C_CHNG_ROLL_PSTN_POP
		    	frm1.vspdData.Col = C_CHNG_ROLL_PSTN_NM
		    	frm1.vspdData.action =0

		    Case C_CHNG_OCPT_TYPE_POP
		    	frm1.vspdData.Col = C_CHNG_OCPT_TYPE_NM
		    	frm1.vspdData.action =0

		    Case C_CHNG_FUNC_CD_POP
		    	frm1.vspdData.Col = C_CHNG_FUNC_CD_NM
		    	frm1.vspdData.action =0

		    Case C_CHNG_ROLE_CD_POP
		    	frm1.vspdData.Col = C_CHNG_ROLE_CD_NM
		    	frm1.vspdData.action =0
        End Select	
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
       	 ggoSpread.Source = frm1.vspdData
         ggoSpread.UpdateRow Row
	End If	

End Function

'======================================================================================================
'	Name : SetCode()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)

	With frm1

		Select Case iWhere
		    Case C_CHNG_DEPT_CD_POP
		        .vspdData.Col = C_CHNG_DEPT_CD
		    	.vspdData.text = arrRet(0)
		    	.vspdData.Col = C_CHNG_DEPT_NM
		    	.vspdData.text = arrRet(1)
		    	.vspdData.action =0

		    Case C_CHNG_PAY_GRD1_POP
		        .vspdData.Col = C_CHNG_PAY_GRD1
		    	.vspdData.text = arrRet(0)
		    	.vspdData.Col = C_CHNG_PAY_GRD1_NM
		    	.vspdData.text = arrRet(1)
				.vspdData.action =0
				
		    Case C_CHNG_ROLL_PSTN_POP
		        .vspdData.Col = C_CHNG_ROLL_PSTN
		    	.vspdData.text = arrRet(0)
		    	.vspdData.Col = C_CHNG_ROLL_PSTN_NM
		    	.vspdData.text = arrRet(1)
		    	.vspdData.action =0

		    Case C_CHNG_OCPT_TYPE_POP
		        .vspdData.Col = C_CHNG_OCPT_TYPE
		    	.vspdData.text = arrRet(0)
		    	.vspdData.Col = C_CHNG_OCPT_TYPE_NM
		    	.vspdData.text = arrRet(1)
		    	.vspdData.action =0

		    Case C_CHNG_FUNC_CD_POP
		        .vspdData.Col = C_CHNG_FUNC_CD
		    	.vspdData.text = arrRet(0)
		    	.vspdData.Col = C_CHNG_FUNC_CD_NM
		    	.vspdData.text = arrRet(1)
		    	.vspdData.action =0

		    Case C_CHNG_ROLE_CD_POP
		        .vspdData.Col = C_CHNG_ROLE_CD
		    	.vspdData.text = arrRet(0)
		    	.vspdData.Col = C_CHNG_ROLE_CD_NM
		    	.vspdData.text = arrRet(1)
		    	.vspdData.action =0

        End Select

	End With

End Function

'======================================================================================================
'	Name : SetEmp()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SetEmp(Row)

    Dim IntRetCD 
    Dim strName
    Dim strDept_nm
    Dim strDept_cd    
    Dim strRoll_pstn_nm
    Dim strPay_grd1_nm
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd

    Dim strPay_grd1
    Dim strRoll_pstn
    Dim strOcpt_type
    Dim strFunc_cd
    Dim strRole_cd
    Dim strResent_promote_dt
    Dim strSelect
    Dim strWhere

    Frm1.vspdData.Col = C_EMP_NO

    if  Frm1.vspdData.value = "" then
        exit sub
    end if
    
    '  부서코드를 가져온다.    
    IntRetCD =  CommonQueryRs(" dept_cd "," haa010t "," emp_no= " & FilterVar(Frm1.vspdData.value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If  IntRetCd = true then
        strDept_cd = Replace(lgF0, Chr(11), "")
    end if

	IntRetCd =  FuncGetEmpInf2(frm1.vspdData.value,lgUsrIntCd,strName,strDept_nm,_
	            strRoll_pstn_nm, strPay_grd1_nm, strPay_grd2, strEntr_dt, strInternal_cd)
    if  IntRetCd = 0 then
        Frm1.vspdData.Col = C_NAME
        Frm1.vspdData.value = strName

        Frm1.vspdData.Col = C_DEPT_CD
        Frm1.vspdData.value = strDept_cd

        Frm1.vspdData.Col = C_DEPT_NM
        Frm1.vspdData.value = strDept_nm

        Frm1.vspdData.Col = C_PAY_GRD1_NM
        Frm1.vspdData.value = strPay_grd1_nm

        Frm1.vspdData.Col = C_PAY_GRD2
        Frm1.vspdData.value = strPay_grd2

        Frm1.vspdData.Col = C_ROLL_PSTN_NM
        Frm1.vspdData.value = strRoll_pstn_nm

        Frm1.vspdData.Col = C_ENTR_DT
        Frm1.vspdData.text =  UNIConvDateDBToCompany(strEntr_dt,"")
    else
	    if  IntRetCd = -1 then
    		Call  DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
        else
            Call  DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
        end if
   	    Frm1.vspdData.Col = C_NAME
        Frm1.vspdData.value = ""
    end if

    Frm1.vspdData.Col = C_EMP_NO
    strSelect = " pay_grd1, roll_pstn, ocpt_type, func_cd, role_cd, resent_promote_dt "
    IntRetCD =  CommonQueryRs(strSelect," haa010t "," emp_no= " & FilterVar(Frm1.vspdData.value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If  IntRetCd = true then

        strPay_grd1 = Replace(lgF0, Chr(11), "")
        strRoll_pstn = Replace(lgF1, Chr(11), "")
        strOcpt_type = Replace(lgF2, Chr(11), "")
        strFunc_cd = Replace(lgF3, Chr(11), "")
        strRole_cd = Replace(lgF4, Chr(11), "")
        strResent_promote_dt = Replace(lgF5, Chr(11), "")

       	Frm1.vspdData.Col = C_PAY_GRD1
        Frm1.vspdData.value = strPay_grd1

       	Frm1.vspdData.Col = C_ROLL_PSTN
        Frm1.vspdData.value = strRoll_pstn

       	Frm1.vspdData.Col = C_OCPT_TYPE
        Frm1.vspdData.value = strOcpt_type

       	Frm1.vspdData.Col = C_OCPT_TYPE_NM
        Frm1.vspdData.value =  FuncCodeName(1,"H0003",strOcpt_type)

       	Frm1.vspdData.Col = C_FUNC_CD
        Frm1.vspdData.value = strFunc_cd
    
       	Frm1.vspdData.Col = C_FUNC_CD_NM
        Frm1.vspdData.value =  FuncCodeName(1,"H0004",strFunc_cd)

       	Frm1.vspdData.Col = C_ROLE_CD
        Frm1.vspdData.value = strRole_cd

       	Frm1.vspdData.Col = C_ROLE_CD_NM
        Frm1.vspdData.value =  FuncCodeName(1,"H0026",strRole_cd)

        Frm1.vspdData.Col = C_RESENT_PROMOTE_DT
        Frm1.vspdData.text = UNIConvDateDBToCompany(strResent_promote_dt,"")
	
    end if
		Frm1.vspdData.Col = C_EMP_NO
		Frm1.vspdData.action =0
End Sub

'===========================================================================
' Function Name : OpenSItemDC
' Function Desc : OpenSItemDC Reference Popup
'===========================================================================
Function OpenSItemDC(iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	    Case 4  ' 직무 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtFunc_cd.Value)			' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0004", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "직무"    						    ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "직무코드"			        		' Header명(0)
	    	arrHeader(1) = "직무명"	        					' Header명(1)

	    Case 3  ' 직종 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtOcpt_cd.Value)			' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0003", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "직종"    						    ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "직종코드"			        		' Header명(0)
	    	arrHeader(1) = "직종명"	        					' Header명(1)

	    Case 26  ' 직책 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtRole_cd.Value)			' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0026", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "직책"    						    ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "직책코드"			        		' Header명(0)
	    	arrHeader(1) = "직책명"	        					' Header명(1)

        Case 101  ' 부서코드 
            arrParam(1) = "BCB020T"							    ' TABLE 명칭 
            arrParam(2) = Trim(frm1.txtDept_cd.Value)           ' Code Condition
            arrParam(3) = ""                                    ' Name Cindition
            arrParam(4) = ""		                		    ' Where Condition
            arrParam(5) = "부서코드"						    ' TextBox 명칭 
	
            arrField(0) = "dept_cd"		    				    ' Field명(0)
            arrField(1) = "dept_name"                           ' Field명(1)
    
            arrHeader(0) = "부서코드"                           ' Header명(0)
            arrHeader(1) = "부서명"                             ' Header명(1)
        Case 29  ' 변동사유 
	        arrParam(0) = "변동사유"		        ' 팝업 명칭 %>
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtChng_cd.Value)		' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0029", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "변동사유"    	    			    ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "변동사유코드" 	        			' Header명(0)
	    	arrHeader(1) = "변동사유명"          					' Header명(1)
	End Select

    arrParam(3) = ""	
	arrParam(0) = arrParam(5)								    ' 팝업 명칭 

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Select Case iWhere
		    Case 101 '부서코드 
		    	frm1.txtDept_cd.focus 
		End Select	
		Exit Function
	Else
		Call SetSItemDC(arrRet, iWhere)
	End If	
	
End Function

'------------------------------------------  SetSItemDC()  --------------------------------------------------
'	Name : SetSItemDC()
'	Description : OpenSItemDC Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetSItemDC(Byval arrRet, Byval iWhere)

	With frm1

		Select Case iWhere
		    Case 2
		    	.txtRoll_pstn.value = arrRet(0)
		    	.txtRoll_pstn_nm.value = arrRet(1)  
		    Case 3
		    	.txtOcpt_cd.value = arrRet(0) 
		    	.txtOcpt_cd_nm.value = arrRet(1)   
		    Case 4
		    	.txtFunc_cd.value = arrRet(0) 
		    	.txtFunc_cd_nm.value = arrRet(1)   
		    Case 101 '부서코드 
		    	.txtDept_cd.value = arrRet(0) 
		    	.txtDept_nm.value = arrRet(1)  
		    	.txtDept_cd.focus 
		    Case 26 '직책 
		    	.txtRole_cd.value = arrRet(0) 
		    	.txtRole_cd_nm.value = arrRet(1)   
		    Case 29 '변동사유 
		    	.txtChng_cd.value = arrRet(0) 
		    	.txtChng_cd_nm.value = arrRet(1)   
		End Select

		lgBlnFlgChgValue = True

	End With
	
End Function

'======================================================================================================
'   Event Name : btnAuto_OnClick
'   Event Desc : 자동입력버튼 
'=======================================================================================================
Sub btnAuto_OnClick()

    Dim IntRetCd
    Dim strDept_nm
    Dim strInternal_cd
    Dim strBasDt

	strBasDt = frm1.txtPro_dt.Text 
	
    If  Trim(frm1.txtResent_promote_dt.Text) = "" then
        Call  DisplayMsgBox("970021","X",frm1.txtResent_promote_dt.Alt,"X")
        frm1.txtResent_promote_dt.focus
        Set gActiveElement = document.ActiveElement
        exit sub        
    end if

    If  Trim(frm1.txtPro_dt.Text) = "" then
        Call  DisplayMsgBox("970021","X",frm1.txtPro_dt.Alt,"X")
        frm1.txtPro_dt.focus
        Set gActiveElement = document.ActiveElement
        exit sub        
    end if

    If  Trim(frm1.txtChng_cd.value) = "" then
        Call  DisplayMsgBox("970021","X",frm1.txtChng_cd.Alt,"X")
        frm1.txtChng_cd.focus
        Set gActiveElement = document.ActiveElement
        exit sub        
    end if   

    If  Trim(frm1.txtResent_promote_dt.Text) <> "" and Trim(frm1.txtResent_promote_dt.Text) = Trim(frm1.txtPro_dt.Text) then
        '변동일은 최근승급일보다 커야합니다.
        Call  DisplayMsgBox("800058","X","X","X")
        frm1.txtResent_promote_dt.focus
        Set gActiveElement = document.ActiveElement
        exit sub
    ElseIf   CompareDateByFormat(frm1.txtResent_promote_dt.Text,frm1.txtPro_dt.Text,frm1.txtResent_promote_dt.Alt,frm1.txtPro_dt.Alt,"800058", parent.gDateFormat, parent.gComDateType,True) = False THEN
        frm1.txtResent_promote_dt.focus
        Set gActiveElement = document.ActiveElement
        exit sub
    END IF

    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData 

    if  Trim(frm1.txtDept_cd.value) = "" then
        strInternal_cd = ""
    else
        IntRetCd =  FuncDeptName(frm1.txtDept_cd.value,strBasDt,lgUsrIntCd,strDept_nm,strInternal_cd)
        if  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call  DisplayMsgBox("800012", "x","x","x")   ' 등록되지 않은 부서코드입니다.
            else
                Call  DisplayMsgBox("800455", "x","x","x")   ' 자료권한이 없습니다.
            end if
            exit sub
        end if
    end if

    lgKeyStream = Frm1.txtResent_promote_dt.Text & parent.gColSep          '0
    lgKeyStream = lgKeyStream & Frm1.txtPay_grd1.Value & parent.gColSep    '1
    lgKeyStream = lgKeyStream & Frm1.txtPay_grd2.Value & parent.gColSep    '2
    lgKeyStream = lgKeyStream & Frm1.txtRoll_pstn.Value & parent.gColSep   '3
    if  strInternal_cd = "" then
        lgKeyStream = lgKeyStream & lgUsrIntCd & parent.gColSep
    else
        lgKeyStream = lgKeyStream & strInternal_cd & parent.gColSep        '4
    end if
    lgKeyStream = lgKeyStream & Frm1.txtResent_promote_dt.Text & parent.gColSep    '5
    lgKeyStream = lgKeyStream & Frm1.txtDept_cd.Value & parent.gColSep             '6

    Err.Clear                                                                        '☜: Clear err status

	If LayerShowHide(1)=False Then
		Exit Sub
	End if

	
	Dim strVal
    
    With Frm1
		strVal = BIZ_PGM_ID1 & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
    End With
		
    If lgIntFlgMode =  parent.OPMD_UMODE Then
    Else
    End If

	Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic
    
End Sub

'========================================================================================================
' Function Name : DbAutoQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbAutoQueryOk()													     

    Dim lRow	

    lgIntFlgMode =  parent.OPMD_UMODE    
    Call  ggoOper.LockField(Document, "Q")										'⊙: Lock field
   
     ggoSpread.Source = frm1.vspdData

	frm1.vspddata.ReDraw = false
     ggoSpread.SSSetProtected	C_NAME, -1
     ggoSpread.SSSetProtected	C_EMP_NO, -1
     ggoSpread.SSSetProtected	C_DEPT_NM, -1
     ggoSpread.SSSetProtected	C_PAY_GRD1_NM, -1
     ggoSpread.SSSetProtected	C_PAY_GRD2, -1
     ggoSpread.SSSetProtected	C_ROLL_PSTN_NM, -1
     ggoSpread.SSSetProtected	C_OCPT_TYPE_NM, -1
     ggoSpread.SSSetProtected	C_FUNC_CD_NM, -1
     ggoSpread.SSSetProtected	C_ROLE_CD_NM, -1
     ggoSpread.SSSetProtected	C_CHNG_CD, -1

     ggoSpread.SpreadUnLock C_CHNG_DEPT_CD, -1, C_PROMOTE_DT

     ggoSpread.SSSetProtected	C_PROMOTE_DT, -1

     ggoSpread.SSSetRequired		C_CHNG_DEPT_NM, -1
     ggoSpread.SSSetRequired		C_CHNG_PAY_GRD1_NM, -1
     ggoSpread.SSSetRequired		C_CHNG_PAY_GRD2, -1

     ggoSpread.SSSetRequired		C_CHNG_ROLL_PSTN_NM, -1
     ggoSpread.SSSetRequired		C_CHNG_OCPT_TYPE_NM, -1
     ggoSpread.SSSetRequired		C_CHNG_FUNC_CD_NM, -1
     ggoSpread.SSSetRequired		C_CHNG_ROLE_CD_NM, -1

    frm1.vspdData.Row = -1
    frm1.vspdData.Col = C_PROMOTE_DT
    frm1.vspdData.text = frm1.txtPro_dt.text

    frm1.vspdData.Col = C_CHNG_PAY_GRD1
    frm1.vspdData.text = frm1.txtChng_pay_grd1.value

    frm1.vspdData.Col = C_CHNG_PAY_GRD1_NM
    frm1.vspdData.text =  FuncCodeName(1, "H0001", frm1.txtChng_pay_grd1.value)

    frm1.vspdData.Col = C_CHNG_PAY_GRD2
    frm1.vspdData.text = frm1.txtChng_pay_grd2.value

    frm1.vspdData.Col = C_CHNG_CD
    frm1.vspdData.text = frm1.txtChng_cd.value

    frm1.vspdData.Col = C_CHNG_CD_NM
    frm1.vspdData.text = frm1.txtChng_cd_nm.value
        
    For lRow = 1 To frm1.vspdData.MaxRows
        frm1.vspdData.Row = lRow
        frm1.vspdData.Col = 0
        frm1.vspdData.text =  ggoSpread.InsertFlag
    Next
    
    frm1.vspddata.ReDraw = true
    ggoSpread.ClearSpreadData "T"
	
End Function
'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
    Dim IntRetCD 
    Dim RetStatus
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt

       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
        Case C_EMP_NO
            Call SetEmp(Row)
    End Select    
             
   	If Frm1.vspdData.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData.text) <  UNICDbl(Frm1.vspdData.TypeFloatMin) Then
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

	Call SetPopupMenuItemInf("1101111111")     
    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If
    End If
       frm1.vspdData.Row = Row
  
End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And  gMouseClickStatus = "SPC" Then
        gMouseClickStatus = "SPCR"
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
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
		Exit Sub
	End if
	
	If frm1.vspdData.MaxRows = 0 then
		Exit Sub
	End if
End Sub
'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
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


'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 컬럼버튼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col
	Select Case Col
	    Case C_EMP_NO_POP
            Call OpenEmpName(Row)
            Call SetEmp(Row)
	    Case C_CHNG_DEPT_CD_POP
            Call OpenDept(1)

	    Case C_CHNG_PAY_GRD1_POP
            Call OpenCode("", C_CHNG_PAY_GRD1_POP, Row)

	    Case C_CHNG_ROLL_PSTN_POP
            Call OpenCode("", C_CHNG_ROLL_PSTN_POP, Row)

	    Case C_CHNG_OCPT_TYPE_POP
            Call OpenCode("", C_CHNG_OCPT_TYPE_POP, Row)

	    Case C_CHNG_FUNC_CD_POP
            Call OpenCode("", C_CHNG_FUNC_CD_POP, Row)

	    Case C_CHNG_ROLE_CD_POP
            Call OpenCode("", C_CHNG_ROLE_CD_POP, Row)
    End Select
    
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
' Name : txtPro_dt_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtPro_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtPro_dt.Action = 7 
        frm1.txtPro_dt.focus
    End If
End Sub

'========================================================================================================
' Name : txtResent_promote_dt_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtResent_promote_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")      
        frm1.txtResent_promote_dt.Action = 7 
        frm1.txtResent_promote_dt.focus
    End If
End Sub

Sub txtDept_cd_OnChange()

    Dim IntRetCd
    Dim strDept_nm
    Dim strInternal_cd
    Dim strBasDt

	strBasDt = frm1.txtPro_dt.Text 
	
    if  frm1.txtDept_cd.value <> "" then    
        IntRetCd =  FuncDeptName(frm1.txtDept_cd.value,strBasDt,lgUsrIntCd,strDept_nm,strInternal_cd)
        if  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call  DisplayMsgBox("800012", "x","x","x")   ' 등록되지 않은 부서코드입니다.
            else
                Call  DisplayMsgBox("800455", "x","x","x")   ' 자료권한이 없습니다.
            end if
            frm1.txtDept_nm.value = ""
        else
            frm1.txtDept_nm.value = strDept_nm
        end if
	else
            frm1.txtDept_nm.value = ""	        
    end if
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>승급/승격등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
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
    	            <TD HEIGHT=20 WIDTH=100%>
    	                <FIELDSET CLASS="CLSFLD">
			            <TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>급호</TD>
								<TD CLASS=TD6 NOWRAP><SELECT Name="txtPay_grd1" ALT="급호" CLASS ="cbonormal" tag="11"><OPTION Value=""></OPTION></SELECT>&nbsp;<INPUT NAME="txtPay_grd2" MAXLENGTH="5" SIZE=5 STYLE="TEXT-ALIGN:left" tag="11"></TD>

								<TD CLASS=TD5 NOWRAP>직위</TD>
								<TD CLASS=TD6 NOWRAP><SELECT Name="txtRoll_pstn" ALT="직위" CLASS ="cbonormal" tag="11"><OPTION Value=""></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>변동급호</TD>
								<TD CLASS=TD6 NOWRAP><SELECT Name="txtChng_pay_grd1" ALT="변동급호" CLASS ="cbonormal" tag="11"><OPTION Value=""></OPTION></SELECT>&nbsp;<INPUT NAME="txtChng_pay_grd2" MAXLENGTH="5" SIZE=5 STYLE="TEXT-ALIGN:left" tag="11"></TD>
								<TD CLASS=TD5 NOWRAP>변동사유</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtChng_cd" ALT="변동사유" TYPE="Text" MAXLENGTH=2 SiZE=5 tag=14XXXU>&nbsp;<INPUT NAME="txtChng_cd_nm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>승급예정일</TD>
								<TD CLASS=TD6 NOWRAP><OBJECT classid=<%=gCLSIDFPDT%> id=txtPro_dt name=txtPro_dt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12X1" ALT="승급예정일"></OBJECT></TD>
								<TD CLASS=TD5 NOWRAP>부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDept_cd" ALT="부서" TYPE="Text" MAXLENGTH=13 SiZE=13 tag=11XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenDept(0)">&nbsp;<INPUT NAME="txtDept_nm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>최근승급일</TD>
								<TD CLASS=TD6 NOWRAP><OBJECT classid=<%=gCLSIDFPDT%> id=txtResent_promote_dt name=txtResent_promote_dt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12X1" ALT="최근승급일"></OBJECT>
								<IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPromoteDt" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPromoteDt(0)"></TD>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP></TD>
							</TR>
			            </TABLE>
			    	    </FIELDSET>
			        </TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100% COLSPAN=2></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP COLSPAN=2>
					    <TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"  id=vaSpread1>
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
	<TR HEIGHT="20">
		<TD>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
				    <TD WIDTH=10>&nbsp;</TD>
				    <TD><BUTTON NAME="btnAuto" CLASS="CLSMBTN">자동입력</BUTTON></TD>
				    <TD WIDTH=* Align=RIGHT></TD>
				    <TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hMajorCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

