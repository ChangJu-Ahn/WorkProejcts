<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : Single Sample
*  3. Program ID           : H1004ma1
*  4. Program Name         : H1004ma1
*  5. Program Desc         : 기준정보관리/지급내역코드등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/04
*  8. Modified date(Last)  : 2003/05/15
*  9. Modifier (First)     : YBI
* 10. Modifier (Last)      : Lee Sina
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"> </SCRIPT>

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID = "H1004mb1.asp"                                      'Biz Logic ASP 
Const CookieSplit = 1233
Const TAB1 = 1
Const TAB2 = 2
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
Dim gSelframeFlg			   ' 현재 TAB의 위치를 나타내는 Flag
Dim lgStrPrevKey1

Dim C_PAY_CD
Dim C_PAY_CD_NM
Dim C_ALLOW_CD
Dim C_ALLOW_NM
Dim C_ALLOW_KIND
Dim C_ALLOW_KIND_NM
Dim C_TAX_TYPE
Dim C_TAX_TYPE_NM
Dim C_LIMIT_AMT
Dim C_CALCU_TYPE
Dim C_CRT_STRT_MM
Dim C_CRT_STRT_MM_NM
Dim C_CRT_STRT_DD
Dim C_BAR
Dim C_CRT_END_MM
Dim C_CRT_END_MM_NM
Dim C_CRT_END_DD
Dim C_TEXT1
Dim C_DAY_CALCU
Dim C_DAY_CALCU_NM
Dim C_TEXT2
Dim C_CALCU_BAS_DD
Dim C_CALCU_BAS_DD_NM
Dim C_ALLOW_SEQ

Dim C_ALLOW_CD1
Dim C_ALLOW_NM1
Dim C_TAX_TYPE1
Dim C_TAX_TYPE_NM1
Dim C_CRT_STRT_MM1
Dim C_CRT_STRT_MM_NM1
Dim C_CRT_STRT_DD1
Dim C_BAR1
Dim C_CRT_END_MM1
Dim C_CRT_END_MM_NM1
Dim C_CRT_END_DD1
Dim C_TEXT11
Dim C_DAY_CALCU1
Dim C_DAY_CALCU_NM1
Dim C_TEXT21
Dim C_CALCU_BAS_DD1
Dim C_CALCU_BAS_DD_NM1
Dim C_ALLOW_SEQ1

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize the position
'========================================================================================================
Sub initSpreadPosVariables(ByVal pvSpdNo)  

    If pvSpdNo = "A" Then
         C_PAY_CD			= 1														<%'Spread Sheet의 Column별 상수 %>
		 C_PAY_CD_NM		= 2
		 C_ALLOW_CD			= 3
		 C_ALLOW_NM			= 4
		 C_ALLOW_KIND		= 5
		 C_ALLOW_KIND_NM	= 6
		 C_TAX_TYPE			= 7
		 C_TAX_TYPE_NM		= 8
		 C_LIMIT_AMT		= 9
		 C_CALCU_TYPE		= 10
		 C_CRT_STRT_MM		= 11
		 C_CRT_STRT_MM_NM	= 12
		 C_CRT_STRT_DD		= 13
		 C_BAR				= 14
		 C_CRT_END_MM		= 15
		 C_CRT_END_MM_NM	= 16
		 C_CRT_END_DD		= 17
		 C_TEXT1			= 18
		 C_DAY_CALCU		= 19
		 C_DAY_CALCU_NM		= 20
		 C_TEXT2			= 21
		 C_CALCU_BAS_DD		= 22
		 C_CALCU_BAS_DD_NM	= 23
		 C_ALLOW_SEQ		= 24

    ElseIf pvSpdNo = "B" Then
         C_ALLOW_CD1		= 1															<%'Spread Sheet의 Column별 상수 %>
		 C_ALLOW_NM1		= 2															
		 C_TAX_TYPE1		= 3
		 C_TAX_TYPE_NM1		= 4
		 C_CRT_STRT_MM1		= 5
		 C_CRT_STRT_MM_NM1	= 6
		 C_CRT_STRT_DD1		= 7
		 C_BAR1				= 8
		 C_CRT_END_MM1		= 9
		 C_CRT_END_MM_NM1	= 10
		 C_CRT_END_DD1		= 11
		 C_TEXT11			= 12
		 C_DAY_CALCU1		= 13
		 C_DAY_CALCU_NM1	= 14
		 C_TEXT21			= 15
		 C_CALCU_BAS_DD1	= 16
		 C_CALCU_BAS_DD_NM1 = 17
		 C_ALLOW_SEQ1		= 18
    End If

End Sub
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
Sub MakeKeyStream(pRow)
    lgKeyStream = Frm1.txtallow_cd.Value & parent.gColSep                                           'You Must append one character( parent.gColSep)
    lgKeyStream = lgKeyStream & "*" & parent.gColSep       '급여구분을 "전체"로 한정함 
End Sub        

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox(ByVal pvSpdNo)
    Dim iCodeArr 
    Dim iNameArr
    Dim IDx
    
	If pvSpdNo = "" OR pvSpdNo = "A" Then	
		ggoSpread.Source = frm1.vspdData
	    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," H_PAY_CD "," MINOR_CD = " & FilterVar("*", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	    iCodeArr = lgF0
	    iNameArr = lgF1
	     ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_PAY_CD
	     ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_PAY_CD_NM
	    ' 수당종류 
	    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0087", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	    iCodeArr = lgF0
	    iNameArr = lgF1
	     ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_ALLOW_KIND
	     ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_ALLOW_KIND_NM
	    ' 세액종류 
	    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0039", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	    iCodeArr = lgF0
	    iNameArr = lgF1
	     ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_TAX_TYPE
	     ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_TAX_TYPE_NM
	    ' 계산구분 
	    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("A1020", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	    iCodeArr = lgF0
	     ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_CALCU_TYPE

	    ' 계산구분 
	    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0088", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	    iCodeArr = lgF0
	    iNameArr = lgF1
	     ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_CRT_STRT_MM
	     ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_CRT_STRT_MM_NM

	    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0088", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	    iCodeArr = lgF0
	    iNameArr = lgF1
	     ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_CRT_END_MM
	     ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_CRT_END_MM_NM
	    ' 일할계산방식 
	    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0089", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	    iCodeArr = lgF0
	    iNameArr = lgF1
	     ggoSpread.SetCombo vbtab & Replace(iCodeArr,Chr(11),vbTab), C_DAY_CALCU
	     ggoSpread.SetCombo vbtab & Replace(iNameArr,Chr(11),vbTab), C_DAY_CALCU_NM

	    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0090", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	    iCodeArr = lgF0
	    iNameArr = lgF1
	     ggoSpread.SetCombo vbtab & Replace(iCodeArr,Chr(11),vbTab), C_CALCU_BAS_DD
	     ggoSpread.SetCombo vbtab & Replace(iNameArr,Chr(11),vbTab), C_CALCU_BAS_DD_NM

	End if
	
	If pvSpdNo = "" OR pvSpdNo = "B" Then	
		ggoSpread.Source = frm1.vspdData1
		' 세액종류 
		Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0039", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		iCodeArr = lgF0
		iNameArr = lgF1
		 ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_TAX_TYPE1
		 ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_TAX_TYPE_NM1
		' 계산구분 
		Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0088", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		iCodeArr = lgF0
		iNameArr = lgF1
		 ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_CRT_STRT_MM1
		 ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_CRT_STRT_MM_NM1

		Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0088", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		iCodeArr = lgF0
		iNameArr = lgF1
		 ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_CRT_END_MM1
		 ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_CRT_END_MM_NM1
		' 일할계산방식 
		Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0089", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		iCodeArr = lgF0
		iNameArr = lgF1
		 ggoSpread.SetCombo vbtab & Replace(iCodeArr,Chr(11),vbTab), C_DAY_CALCU1
		 ggoSpread.SetCombo vbtab & Replace(iNameArr,Chr(11),vbTab), C_DAY_CALCU_NM1

		Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0090", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		iCodeArr = lgF0
		iNameArr = lgF1
		 ggoSpread.SetCombo vbtab & Replace(iCodeArr,Chr(11),vbTab), C_CALCU_BAS_DD1
		 ggoSpread.SetCombo vbtab & Replace(iNameArr,Chr(11),vbTab), C_CALCU_BAS_DD_NM1
	End if	
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
	If pvSpdNo = "" OR pvSpdNo = "A" Then

		Call initSpreadPosVariables("A")	
		With frm1.vspdData
 
		    ggoSpread.Source = frm1.vspdData	
		    ggoSpread.Spreadinit "V20021129",,parent.gAllowDragDropSpread    

		    .ReDraw = false
		    .MaxCols = C_ALLOW_SEQ + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
		    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
		    .ColHidden = True
		    
		    .MaxRows = 0
		    ggoSpread.ClearSpreadData
		    
		Call AppendNumberPlace("6","2","0")
		Call GetSpreadColumnPos("A")

		 ggoSpread.SSSetCombo    C_PAY_CD,           "급여구분",         5
		 ggoSpread.SSSetCombo    C_PAY_CD_NM,        "급여구분",         10
		 ggoSpread.SSSetEdit     C_ALLOW_CD,         "코드",             10,,,3,2
		 ggoSpread.SSSetEdit     C_ALLOW_NM,         "코드명" ,          15,,, 20
		 ggoSpread.SSSetCombo    C_ALLOW_KIND,       "수당종류",         5
		 ggoSpread.SSSetCombo    C_ALLOW_KIND_NM,    "수당종류",         10
		 ggoSpread.SSSetCombo    C_TAX_TYPE,         "세액구분",         5
		 ggoSpread.SSSetCombo    C_TAX_TYPE_NM,      "세액구분",         10
		 ggoSpread.SSSetFloat    C_LIMIT_AMT,        "비과세한도금액",   14, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"		 
		 ggoSpread.SSSetCombo    C_CALCU_TYPE,       "계산구분",         10
		 ggoSpread.SSSetCombo    C_CRT_STRT_MM,      "계산기간",         5
		 ggoSpread.SSSetCombo    C_CRT_STRT_MM_NM,   "계산시작월",       10
		 ggoSpread.SSSetMask     C_CRT_STRT_DD,      "일",              05,2, "99"

		 ggoSpread.SSSetEdit     C_BAR,              "" ,                    2,2

		 ggoSpread.SSSetCombo    C_CRT_END_MM,       "",                     10
		 ggoSpread.SSSetCombo    C_CRT_END_MM_NM,    "계산종료월",       10
		 ggoSpread.SSSetMask     C_CRT_END_DD,       "일",              05,2, "99"

		 ggoSpread.SSSetEdit     C_TEXT1,            "일할계산",         10,,, 10    
		 ggoSpread.SSSetCombo    C_DAY_CALCU,        "",                     5
		 ggoSpread.SSSetCombo    C_DAY_CALCU_NM,     "계산단위",         10
		 ggoSpread.SSSetEdit     C_TEXT2,            "",                     2,2    ' '*'
		 ggoSpread.SSSetCombo    C_CALCU_BAS_DD,     "",                     10
		 ggoSpread.SSSetCombo    C_CALCU_BAS_DD_NM,  "일수",             10
		 ggoSpread.SSSetFloat    C_ALLOW_SEQ,        "순번",             10,"6", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,,"0","99"
		 
		 Call ggoSpread.SSSetColHidden(C_PAY_CD		,  C_PAY_CD			, True)
		 Call ggoSpread.SSSetColHidden(C_ALLOW_KIND	,  C_ALLOW_KIND		, True)
		 Call ggoSpread.SSSetColHidden(C_TAX_TYPE	,  C_TAX_TYPE		, True)
		 Call ggoSpread.SSSetColHidden(C_CRT_STRT_MM,  C_CRT_STRT_MM	, True)
		 Call ggoSpread.SSSetColHidden(C_CRT_END_MM	,  C_CRT_END_MM		, True)
		 Call ggoSpread.SSSetColHidden(C_DAY_CALCU	,  C_DAY_CALCU		, True)
		 Call ggoSpread.SSSetColHidden(C_CALCU_BAS_DD, C_CALCU_BAS_DD	, True)
		                           
		.ReDraw = true

		Call SetSpreadLock("A") 
    
		End With
    
    End if
    
    If pvSpdNo = "" OR pvSpdNo = "B" Then		

		Call initSpreadPosVariables("B")	
		With frm1.vspdData1
 
		    ggoSpread.Source = frm1.vspdData1	
		    ggoSpread.Spreadinit "V20021129",,parent.gAllowDragDropSpread    

		    .ReDraw = false
		    .MaxCols = C_ALLOW_SEQ1 + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
		    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
		    .ColHidden = True
		    
		    .MaxRows = 0
		    ggoSpread.ClearSpreadData
	
		 Call GetSpreadColumnPos("B")

		 ggoSpread.SSSetEdit     C_ALLOW_CD1,        "코드",             10,,,3,2
		 ggoSpread.SSSetEdit     C_ALLOW_NM1,        "코드명" ,          15,,, 20

		 ggoSpread.SSSetCombo    C_TAX_TYPE1,        "세액구분",         5
		 ggoSpread.SSSetCombo    C_TAX_TYPE_NM1,     "세액구분",         10

		 ggoSpread.SSSetCombo    C_CRT_STRT_MM1,     "계산기간",         5
		 ggoSpread.SSSetCombo    C_CRT_STRT_MM_NM1,  "계산시작월",         10
		 ggoSpread.SSSetMask     C_CRT_STRT_DD1,     "일",              05,2, "99"

		 ggoSpread.SSSetEdit     C_BAR1,             "" ,					2,2

		 ggoSpread.SSSetCombo    C_CRT_END_MM1,      "",                     10
		 ggoSpread.SSSetCombo    C_CRT_END_MM_NM1,   "계산종료월",                     10
		 ggoSpread.SSSetMask     C_CRT_END_DD1,      "일",              05,2, "99"

		 ggoSpread.SSSetEdit     C_TEXT11,           "일할계산",         10,,, 10
    
		 ggoSpread.SSSetCombo    C_DAY_CALCU1,       "",                     5
		 ggoSpread.SSSetCombo    C_DAY_CALCU_NM1,    "계산단위",             10

		 ggoSpread.SSSetEdit     C_TEXT21,           "",                     2,2    ' '*'
		 ggoSpread.SSSetCombo    C_CALCU_BAS_DD1,    "",                     10
		 ggoSpread.SSSetCombo    C_CALCU_BAS_DD_NM1, "일수",                     10
		 ggoSpread.SSSetFloat    C_ALLOW_SEQ1,       "순번",             10,"6", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,,"0","99"
		 
		 Call ggoSpread.SSSetColHidden(C_TAX_TYPE1		,  C_TAX_TYPE1		, True)
		 Call ggoSpread.SSSetColHidden(C_CRT_STRT_MM1	,  C_CRT_STRT_MM1	, True)
		 Call ggoSpread.SSSetColHidden(C_CRT_END_MM1	,  C_CRT_END_MM1	, True)
		 Call ggoSpread.SSSetColHidden(C_DAY_CALCU1		,  C_DAY_CALCU1		, True)
		 Call ggoSpread.SSSetColHidden(C_CALCU_BAS_DD1	,  C_CALCU_BAS_DD1	, True)

		.ReDraw = true
	
		Call SetSpreadLock("B") 
    
		End With
	End if
    
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)

    If pvSpdNo = "A" Then

        ggoSpread.Source = Frm1.vspdData

        With frm1.vspdData
        	.ReDraw = False
        	
			 ggoSpread.SpreadLock		C_PAY_CD			, -1, C_PAY_CD
			 ggoSpread.SpreadLock		C_PAY_CD_NM			, -1, C_PAY_CD_NM
			 ggoSpread.SpreadLock		C_BAR				, -1, C_BAR
			 ggoSpread.SpreadLock		C_ALLOW_CD			, -1, C_ALLOW_CD
			 ggoSpread.SSSetRequired	C_ALLOW_NM			, -1, C_ALLOW_NM
			 ggoSpread.SSSetProtected	C_ALLOW_KIND		, -1,C_ALLOW_KIND
			 ggoSpread.SSSetRequired	C_ALLOW_KIND_NM		, -1,C_ALLOW_KIND_NM
			 ggoSpread.SSSetProtected	C_TAX_TYPE			, -1,C_TAX_TYPE
			 ggoSpread.SSSetRequired	C_TAX_TYPE_NM		, -1,C_TAX_TYPE_NM
			 ggoSpread.SSSetRequired	C_CALCU_TYPE		, -1,C_CALCU_TYPE
			 ggoSpread.SSSetProtected	C_CRT_STRT_MM		, -1,C_CRT_STRT_MM
			 ggoSpread.SSSetRequired	C_CRT_STRT_MM_NM	, -1,C_CRT_STRT_MM_NM
			 ggoSpread.SSSetRequired	C_CRT_STRT_DD		, -1,C_CRT_STRT_DD
			 ggoSpread.SSSetProtected	C_CRT_END_MM		, -1,C_CRT_END_MM
			 ggoSpread.SSSetRequired	C_CRT_END_MM_NM		, -1,C_CRT_END_MM_NM
			 ggoSpread.SSSetRequired	C_CRT_END_DD		, -1,C_CRT_END_DD        
			 ggoSpread.SpreadLock		C_TEXT1				, -1, C_TEXT1
			 ggoSpread.SSSetRequired	C_ALLOW_SEQ			, -1,C_ALLOW_SEQ
			 ggoSpread.SpreadLock		C_TEXT2				, -1,C_TEXT2
			 ggoSpread.SSSetProtected	.MaxCols			,-1,-1

			.ReDraw = True
        End With
        
 ElseIf pvSpdNo = "B" Then											
        ggoSpread.Source = Frm1.vspdData1

        With frm1.vspdData1
      		.ReDraw = False

			 ggoSpread.SpreadLock		C_ALLOW_CD1			, -1,C_ALLOW_CD1
			 ggoSpread.SSSetRequired	C_ALLOW_NM1			, -1,C_ALLOW_NM1

			 ggoSpread.SSSetProtected	C_TAX_TYPE1			, -1,C_TAX_TYPE1
			 ggoSpread.SSSetRequired	C_TAX_TYPE_NM1		, -1,C_TAX_TYPE_NM1

			 ggoSpread.SSSetProtected	C_CRT_STRT_MM1		, -1,C_CRT_STRT_MM1
			 ggoSpread.SSSetRequired	C_CRT_STRT_MM_NM1	, -1,C_CRT_STRT_MM_NM1
			 ggoSpread.SSSetRequired	C_CRT_STRT_DD1		, -1,C_CRT_STRT_DD1
        
			 ggoSpread.SSSetProtected	C_CRT_END_MM1		, -1,C_CRT_END_MM1
			 ggoSpread.SSSetRequired	C_CRT_END_MM_NM1	, -1,C_CRT_END_MM_NM1
			 ggoSpread.SSSetRequired	C_CRT_END_DD1		, -1,C_CRT_END_DD1
			 ggoSpread.SpreadLock		C_TEXT11			, -1,C_TEXT11
			 ggoSpread.SpreadLock		C_TEXT21			, -1,C_TEXT21
			 ggoSpread.SSSetRequired	C_ALLOW_SEQ1		, -1,C_ALLOW_SEQ1

			 ggoSpread.SpreadLock		C_BAR1				, -1,C_BAR1
			 ggoSpread.SSSetProtected	.MaxCols			,-1,-1
			.ReDraw = True
        End With
    End If
                
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
	If gSelframeFlg = TAB1 Then
		With frm1
    
		.vspdData.ReDraw = False
        
         ggoSpread.SSSetProtected		C_PAY_CD		, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_PAY_CD_NM		, pvStartRow, pvEndRow

         ggoSpread.SSSetRequired		C_ALLOW_CD		, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_ALLOW_NM		, pvStartRow, pvEndRow

         ggoSpread.SSSetProtected		C_ALLOW_KIND	, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_ALLOW_KIND_NM	, pvStartRow, pvEndRow

         ggoSpread.SSSetProtected		C_TAX_TYPE		, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_TAX_TYPE_NM	, pvStartRow, pvEndRow

         ggoSpread.SSSetRequired		C_CALCU_TYPE	, pvStartRow, pvEndRow

         ggoSpread.SSSetProtected		C_CRT_STRT_MM	, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_CRT_STRT_MM_NM, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_CRT_STRT_DD	, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected		C_CRT_END_MM	, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_CRT_END_MM_NM	, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_CRT_END_DD	, pvStartRow, pvEndRow

         ggoSpread.SSSetProtected		C_TEXT1			, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected		C_BAR			, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected		C_TEXT2			, pvStartRow, pvEndRow

         ggoSpread.SSSetRequired		C_ALLOW_SEQ		, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected		.vspdData.MaxCols, pvStartRow, pvEndRow
		.vspdData.ReDraw = True
    
		End With
    Else    
		With frm1
    
		.vspdData1.ReDraw = False        
        
         ggoSpread.SSSetRequired		C_ALLOW_CD1		 , pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_ALLOW_NM1		 , pvStartRow, pvEndRow

         ggoSpread.SSSetRequired		C_TAX_TYPE1		 , pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_TAX_TYPE_NM1	 , pvStartRow, pvEndRow

         ggoSpread.SSSetRequired		C_CRT_STRT_MM1	 , pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_CRT_STRT_MM_NM1, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_CRT_STRT_DD1	 , pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_CRT_END_MM1	 , pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_CRT_END_MM_NM1 , pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_CRT_END_DD1	 , pvStartRow, pvEndRow

         ggoSpread.SSSetProtected		C_TEXT11		 , pvStartRow, pvEndRow
         ggoSpread.SSSetProtected		C_TEXT21	     , pvStartRow, pvEndRow
         ggoSpread.SSSetProtected		C_BAR1			 , pvStartRow, pvEndRow

         ggoSpread.SSSetRequired		C_ALLOW_SEQ1	 , pvStartRow, pvEndRow
         ggoSpread.SSSetProtected		.vspdData1.MaxCols	, pvStartRow, pvEndRow

		.vspdData1.ReDraw = True
    
		End With
	End If

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

'======================================================================================================
' Function Name : GetSpreadColumnPos
' Function Desc : GetSpreadColumnPostion from XML file
'======================================================================================================

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
            C_PAY_CD			= iCurColumnPos(1)														<%'Spread Sheet의 Column별 상수 %>
			C_PAY_CD_NM			= iCurColumnPos(2)
			C_ALLOW_CD			= iCurColumnPos(3)
			C_ALLOW_NM			= iCurColumnPos(4)
			C_ALLOW_KIND		= iCurColumnPos(5)
			C_ALLOW_KIND_NM		= iCurColumnPos(6)
			C_TAX_TYPE			= iCurColumnPos(7)
			C_TAX_TYPE_NM		= iCurColumnPos(8)
			C_LIMIT_AMT			= iCurColumnPos(9)
			C_CALCU_TYPE		= iCurColumnPos(10)
			C_CRT_STRT_MM		= iCurColumnPos(11)
			C_CRT_STRT_MM_NM	= iCurColumnPos(12)
			C_CRT_STRT_DD		= iCurColumnPos(13)
			C_BAR				= iCurColumnPos(14)
			C_CRT_END_MM		= iCurColumnPos(15)
			C_CRT_END_MM_NM		= iCurColumnPos(16)
			C_CRT_END_DD		= iCurColumnPos(17)
			C_TEXT1				= iCurColumnPos(18)
			C_DAY_CALCU			= iCurColumnPos(19)
			C_DAY_CALCU_NM		= iCurColumnPos(20)
			C_TEXT2				= iCurColumnPos(21)
			C_CALCU_BAS_DD		= iCurColumnPos(22)
			C_CALCU_BAS_DD_NM	= iCurColumnPos(23)
			C_ALLOW_SEQ			= iCurColumnPos(24)            
    
       Case "B"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

            C_ALLOW_CD1			= iCurColumnPos(1)															<%'Spread Sheet의 Column별 상수 %>
			C_ALLOW_NM1			= iCurColumnPos(2)															
			C_TAX_TYPE1			= iCurColumnPos(3)
			C_TAX_TYPE_NM1		= iCurColumnPos(4)
			C_CRT_STRT_MM1		= iCurColumnPos(5)
			C_CRT_STRT_MM_NM1	= iCurColumnPos(6)
			C_CRT_STRT_DD1		= iCurColumnPos(7)
			C_BAR1				= iCurColumnPos(8)
			C_CRT_END_MM1		= iCurColumnPos(9)
			C_CRT_END_MM_NM1	= iCurColumnPos(10)
			C_CRT_END_DD1		= iCurColumnPos(11)
			C_TEXT11			= iCurColumnPos(12)
			C_DAY_CALCU1		= iCurColumnPos(13)
			C_DAY_CALCU_NM1		= iCurColumnPos(14)
			C_TEXT21			= iCurColumnPos(15)
			C_CALCU_BAS_DD1		= iCurColumnPos(16)
			C_CALCU_BAS_DD_NM1	= iCurColumnPos(17)
			C_ALLOW_SEQ1		= iCurColumnPos(18)            
            
    End Select    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

	Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
    Call  ggoOper.FormatField(Document, "2", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	Call  ggoOper.LockField(Document, "N")											'⊙: Lock Field
    Call InitSpreadSheet("")                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    Call InitComboBox("")
	gSelframeFlg = TAB1
	Call changeTabs(TAB1)    
    Call SetToolbar("1100110100101111")										        '버튼 툴바 제어 

    frm1.txtAllow_cd.Focus
    gIsTab     = "Y" ' <- "Yes"의 약자 Y(와이) 입니다.[V(브이)아닙니다]
    gTabMaxCnt = 2   ' Tab의 갯수를 적어 주세요    
	lgCurrentSpd = "M"
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

     ggoSpread.Source = Frm1.vspdData
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call  ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field

    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    If txtAllow_cd_Onchange() Then          'enter key 로 조회시 수당코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    Call InitVariables                                                           '⊙: Initializes local global variables

    Call MakeKeyStream("X")    

    lgCurrentSpd = "M"  ' 급여지급내역 
    
	Call  DisableToolBar( parent.TBC_QUERY)
	If DbQuery = False Then
		Call  RestoreToolBar()
		Exit Function
	End If																'☜: Query db data
       
    FncQuery = True                                                              '☜: Processing is OK
    
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD
    Dim lRow
    Dim intStrt_dd
    Dim intEnd_dd
    Dim intStrt_mm
    Dim intEnd_mm
    Dim intAllow_kind_cnt
    Dim len_count
    
    
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    if gSelframeFlg = TAB1 then
         ggoSpread.Source = frm1.vspdData
        If  ggoSpread.SSCheckChange = False Then
            IntRetCD =  DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
            Exit Function
        End If
        lgCurrentSpd = "M"  ' 급여지급내역 
    else
         ggoSpread.Source = frm1.vspdData1
        If  ggoSpread.SSCheckChange = False Then
            IntRetCD =  DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
            Exit Function
        End If
        lgCurrentSpd = "S"  ' 기타지급내역 
    end if
    
    if gSelframeFlg = TAB1 then
		 ggoSpread.Source = frm1.vspdData
		If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
		   Exit Function
		End If
    else
		 ggoSpread.Source = frm1.vspdData1
		If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
		   Exit Function
		End If
	end if

    if  lgCurrentSpd = "M" then  ' 급여지급내역 
         ggoSpread.Source = frm1.vspdData
        intAllow_kind_cnt = 0
	    With Frm1
           For lRow = 1 To .vspdData.MaxRows
               .vspdData.Row = lRow
               .vspdData.Col = 0
               if   .vspdData.Text =  ggoSpread.InsertFlag OR .vspdData.Text =  ggoSpread.UpdateFlag then				
                    .vspdData.Col = C_CRT_STRT_DD					
					
					
					For len_count = 1 to Len(.vspdData.Text)
					     
						If (asc(Mid(.vspdData.Text,len_count,1)) < 48) OR (asc(Mid(.vspdData.Text,len_count,1)) > 58) Then
							call  DisplayMsgBox("126404", "x","x","x")
							.vspdData.Action = 0
							Exit Function						
						End If
					Next					
					    
                    if  Cint(.vspdData.Text) >= 0 AND Cint(.vspdData.Text) <= 9 then
                        .vspdData.Text = "0" & Cstr(Cint(.vspdData.Text))
                    elseif  Cint(.vspdData.Text) >= 10 AND Cint(.vspdData.Text) <= 31 then
                    else
                        call  DisplayMsgBox("800087", "x","x","x")
                        .vspdData.Action = 0 ' go to 
                        exit function
                    end if

                    .vspdData.Col = C_CRT_END_DD
                    
                    For len_count = 1 to Len(.vspdData.Text)					     
						If (asc(Mid(.vspdData.Text,len_count,1)) < 48) OR (asc(Mid(.vspdData.Text,len_count,1)) > 58) Then
							call  DisplayMsgBox("126404", "x","x","x")
							.vspdData.Action = 0
							Exit Function						
						End If
					Next		
                    
                    if  Cint(.vspdData.Text) >= 0 AND Cint(.vspdData.Text) <= 9 then
                        .vspdData.Text = "0" & Cstr(Cint(.vspdData.Text))
                    elseif  Cint(.vspdData.Text) >= 10 AND Cint(.vspdData.Text) <= 31 then
                    else
                        call  DisplayMsgBox("800206", "x","x","x")
                        .vspdData.Action = 0 ' go to 
                        exit function
                    end if

                    .vspdData.Col = C_CRT_STRT_DD
                    intStrt_dd =  UNICDbl(.vspdData.Text)
                    .vspdData.Col = C_CRT_END_DD
                    intEnd_dd =  UNICDbl(.vspdData.Text)

                    .vspdData.Col = C_CRT_STRT_MM
                    intStrt_mm =  UNICDbl(.vspdData.Text)
                    .vspdData.Col = C_CRT_END_MM
                    intEnd_mm =  UNICDbl(.vspdData.Text)

                    If intEnd_dd="00" Or intEnd_dd = "0" Then
                        intEnd_dd = "31"
                    End if
                    
                    if  intStrt_mm > intEnd_mm then
                        call  DisplayMsgBox("800205", "x","x","x")
                        .vspdData.Col = C_CRT_END_MM
                        .vspdData.Action = 0 ' go to
                        Set gActiveElement = document.activeElement
                        exit function
                    elseif  intStrt_mm = intEnd_mm then
                        if  intStrt_dd >= intEnd_dd then
                            call  DisplayMsgBox("800205", "x","x","x")
                            .vspdData.Col = C_CRT_END_DD
                            .vspdData.Action = 0 ' go to
                            Set gActiveElement = document.activeElement
                            exit function
                        end if
                    end if
                end if
                
                .vspdData.Col = C_ALLOW_KIND
                if  .vspdData.Text = "1" then
                    intAllow_kind_cnt = intAllow_kind_cnt + 1
                end if
                
            next
            if  intAllow_kind_cnt > 3 then
                call  DisplayMsgBox("800470", "x","x","x")
                exit function
            end if
        end with
    else
         ggoSpread.Source = frm1.vspdData1

	    With Frm1
           For lRow = 1 To .vspdData1.MaxRows
               .vspdData1.Row = lRow
               .vspdData1.Col = 0
               if   .vspdData1.Text =  ggoSpread.InsertFlag OR .vspdData1.Text =  ggoSpread.UpdateFlag then

                    .vspdData1.Col = C_CRT_STRT_DD1
                    
                    For len_count = 1 to Len(.vspdData1.Text)					     
						If (asc(Mid(.vspdData1.Text,len_count,1)) < 48) OR (asc(Mid(.vspdData1.Text,len_count,1)) > 58) Then
							call  DisplayMsgBox("126404", "x","x","x")
							.vspdData1.Action = 0
							Exit Function						
						End If
					Next		
                    
                    if  Cint(.vspdData1.Text) >= 0 AND Cint(.vspdData1.Text) <= 9 then
                        .vspdData1.Text = "0" & Cstr(Cint(.vspdData1.Text))
                    elseif  Cint(.vspdData1.Text) >= 10 AND Cint(.vspdData1.Text) <= 31 then
                    else
                        call  DisplayMsgBox("800087", "x","x","x")
                        .vspdData1.Action = 0 ' go to 
                        exit function
                    end if

                    .vspdData1.Col = C_CRT_END_DD1
                    
                    For len_count = 1 to Len(.vspdData1.Text)					     
						If (asc(Mid(.vspdData1.Text,len_count,1)) < 48) OR (asc(Mid(.vspdData1.Text,len_count,1)) > 58) Then
							call  DisplayMsgBox("126404", "x","x","x")
							.vspdData1.Action = 0
							Exit Function						
						End If
					Next		
                    
                    if  Cint(.vspdData1.Text) >= 0 AND Cint(.vspdData1.Text) <= 9 then
                        .vspdData1.Text = "0" & Cstr(Cint(.vspdData1.Text))
                    elseif  Cint(.vspdData1.Text) >= 10 AND Cint(.vspdData1.Text) <= 31 then
                    else
                        call  DisplayMsgBox("800206", "x","x","x")
                        .vspdData1.Action = 0 ' go to 
                        exit function
                    end if

                    .vspdData1.Col = C_CRT_STRT_DD1
                    intStrt_dd = Cint(.vspdData1.Text)
                    .vspdData1.Col = C_CRT_END_DD1
                    intEnd_dd = Cint(.vspdData1.Text)

                    .vspdData1.Col = C_CRT_STRT_MM1
                    intStrt_mm = Cint(.vspdData1.Text)
                    .vspdData1.Col = C_CRT_END_MM1
                    intEnd_mm = Cint(.vspdData1.Text)

                    If intEnd_dd="00" Or intEnd_dd = "0" Then
                        intEnd_dd = "31"
                    End if
                    
                    if  intStrt_mm > intEnd_mm then
                        call  DisplayMsgBox("800205", "x","x","x")
                        .vspdData1.Col = C_CRT_END_MM1
                        .vspdData1.Action = 0 ' go to
                        Set gActiveElement = document.activeElement
                        exit function
                    elseif  intStrt_mm = intEnd_mm then
                        if  intStrt_dd >= intEnd_dd then
                            call  DisplayMsgBox("800205", "x","x","x")
                            .vspdData1.Col = C_CRT_END_DD1
                            .vspdData1.Action = 0 ' go to
                            Set gActiveElement = document.activeElement
                            exit function
                        end if
                    end if
                end if
               
            next
        end with
    end if
    
	Call  DisableToolBar( parent.TBC_SAVE)
	If DbSAVE = False Then
		Call  RestoreToolBar()
		Exit Function
	End If																'☜: Query db data
        
    FncSave = True                                                              '☜: Processing is OK
    
End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()
	If gSelframeFlg = TAB1 Then
        lgCurrentSpd = "M"

		If Frm1.vspdData.MaxRows < 1 Then
		   Exit Function
		End If
    
		 ggoSpread.Source = Frm1.vspdData
		With Frm1.VspdData
		     .ReDraw = False
			 If .ActiveRow > 0 Then
		         ggoSpread.CopyRow
				 SetSpreadColor .ActiveRow, .ActiveRow
    
		        .ReDraw = True
			    .Focus
			 End If
		End With
	Else
	    lgCurrentSpd = "S"	
		If Frm1.vspdData1.MaxRows < 1 Then
		   Exit Function
		End If
    
		 ggoSpread.Source = Frm1.vspdData1
		With Frm1.VspdData1
		     .ReDraw = False
			 If .ActiveRow > 0 Then
		         ggoSpread.CopyRow
				 SetSpreadColor .ActiveRow, .ActiveRow
    
		        .ReDraw = True
			    .Focus
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
    if  gSelframeFlg = TAB1 then
         ggoSpread.Source = Frm1.vspdData
         ggoSpread.EditUndo  
    else
         ggoSpread.Source = Frm1.vspdData1
         ggoSpread.EditUndo
    end if
    Call InitData()

End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
	Dim imRow
	Dim iRow

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
        
        For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1 
			
			.vspdData.Row = iRow
			.vspdData.Col  = C_TEXT1
			.vspdData.Text = "수당/"

			.vspdData.Col  = C_TEXT2
			.vspdData.Text = "*"

			.vspdData.Col  = C_BAR
			.vspdData.Text = "~" 
        
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
    If gSelframeFlg = TAB1 Then
	   lgCurrentSpd = "M"
    
		If Frm1.vspdData.MaxRows < 1 then
		   Exit function
		End if	
		With Frm1.vspdData 
			.focus
			 ggoSpread.Source = frm1.vspdData 
			lDelRows =  ggoSpread.DeleteRow
		End With
	Else
	   lgCurrentSpd = "S"	
		If Frm1.vspdData1.MaxRows < 1 then
		   Exit function
		End if	
		With Frm1.vspdData1
			.focus
			 ggoSpread.Source = frm1.vspdData1
			lDelRows =  ggoSpread.DeleteRow
		End With

	End if	   
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
	Call LayerShowHide(1)
	
	Dim strVal
	
	strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         	
	strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
	strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows

	If gSelframeFlg = Tab1 Then    
 	    lgCurrentSpd = "M"	
	    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
	Else
	    lgCurrentSpd = "S"
	    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey1                 '☜: Next key tag
	End If
    strVal = strVal     & "&lgCurrentSpd="       & lgCurrentSpd                      '☜: Next key tag		

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
    Dim intCnt
	
    DbSave = False                                                          
    
    Call LayerShowHide(1)

    strVal = ""
    strDel = ""
    lGrpCnt = 1

    if lgCurrentSpd = "M" then  ' 급여지급내역 
         ggoSpread.Source = frm1.vspdData 
	    With Frm1
           For lRow = 1 To .vspdData.MaxRows
               .vspdData.Row = lRow
               .vspdData.Col = 0
               Select Case .vspdData.Text
                   Case  ggoSpread.InsertFlag                                      '☜: Create
                                                          strVal = strVal & "C" & parent.gColSep
                                                          strVal = strVal & lRow & parent.gColSep
                                                          strVal = strVal & lgCurrentSpd & parent.gColSep
                        .vspdData.Col = C_PAY_CD	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_ALLOW_CD	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_ALLOW_NM	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

                        .vspdData.Col = C_ALLOW_KIND	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_TAX_TYPE	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_LIMIT_AMT	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_CALCU_TYPE	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_CRT_STRT_MM	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_CRT_STRT_DD	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_CRT_END_MM	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_CRT_END_DD	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_DAY_CALCU 	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_CALCU_BAS_DD  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_ALLOW_SEQ     : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                        lGrpCnt = lGrpCnt + 1
                   Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                          strVal = strVal & "U" & parent.gColSep
                                                          strVal = strVal & lRow & parent.gColSep
                                                          strVal = strVal & lgCurrentSpd & parent.gColSep
                        .vspdData.Col = C_PAY_CD	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_ALLOW_CD	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_ALLOW_NM	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

                        .vspdData.Col = C_ALLOW_KIND	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_TAX_TYPE	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_LIMIT_AMT	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_CALCU_TYPE	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_CRT_STRT_MM	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_CRT_STRT_DD	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_CRT_END_MM	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_CRT_END_DD	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_DAY_CALCU 	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_CALCU_BAS_DD  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_ALLOW_SEQ     : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                        lGrpCnt = lGrpCnt + 1


                   Case  ggoSpread.DeleteFlag                                      '☜: Delete
	          .vspdData.Col = C_ALLOW_CD


' 메세지처리 2007.04.20  900020 이 데이타를 참조하고 있는 데이타가 있어서 삭제가 불가능합니다.
     	      If CommonQueryRs(" COUNT(*) "," hdf040t", " allow_cd = " & FilterVar(Trim(.vspdData.Text), "''", "S") & " " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true Then
		intCnt = CInt(Replace(lgF0, Chr(11), "")) 
	      end if 

	      if intCnt > 0  then
	 	Call LayerShowHide(0)
  		Call DisplayMsgbox("900020","X","X","X") 
       		Exit function
   	      end if

   	      If CommonQueryRs(" COUNT(*) "," hdf041t", " allow_cd = " & FilterVar(Trim(.vspdData.Text), "''", "S") & " " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true Then
		intCnt = CInt(Replace(lgF0, Chr(11), "")) 
	      end if 

	      if intCnt > 0  then
	 	Call LayerShowHide(0)
  		Call DisplayMsgbox("900020","X","X","X") 
       		Exit function
   	      end if



                                                          strDel = strDel & "D" & parent.gColSep
                                                          strDel = strDel & lRow & parent.gColSep
                                                          strDel = strDel & lgCurrentSpd & parent.gColSep
                        .vspdData.Col = C_PAY_CD	    : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_ALLOW_CD       : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_ALLOW_NM	    : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                        lGrpCnt = lGrpCnt + 1
               End Select
           Next
           .txtMode.value        =  parent.UID_M0002
           .txtUpdtUserId.value  =  parent.gUsrID
           .txtInsrtUserId.value =  parent.gUsrID
	       .txtMaxRows.value     = lGrpCnt-1	
	       .txtSpread.value      = strDel & strVal
	    End With
    else
         ggoSpread.Source = frm1.vspdData1
	    With Frm1
           For lRow = 1 To .vspdData1.MaxRows
               .vspdData1.Row = lRow
               .vspdData1.Col = 0
               Select Case .vspdData1.Text
                   Case  ggoSpread.InsertFlag                                      '☜: Create
                   Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                          strVal = strVal & "U" & parent.gColSep
                                                          strVal = strVal & lRow & parent.gColSep
                                                          strVal = strVal & lgCurrentSpd & parent.gColSep
                        .vspdData1.Col = C_ALLOW_CD1	: strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
                        .vspdData1.Col = C_ALLOW_NM1	: strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep

                        .vspdData1.Col = C_TAX_TYPE1	: strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
                        .vspdData1.Col = C_CRT_STRT_MM1	: strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
                        .vspdData1.Col = C_CRT_STRT_DD1	: strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
                        .vspdData1.Col = C_CRT_END_MM1	: strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
                        .vspdData1.Col = C_CRT_END_DD1	: strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
                        .vspdData1.Col = C_DAY_CALCU1 	: strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
                        .vspdData1.Col = C_CALCU_BAS_DD1: strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep

                        .vspdData1.Col = C_ALLOW_SEQ1   : strVal = strVal & Trim(.vspdData1.Text) & parent.gRowSep
                        lGrpCnt = lGrpCnt + 1
                   Case  ggoSpread.DeleteFlag                                      '☜: Delete
               End Select
           Next
           .txtMode.value        =  parent.UID_M0002
           .txtUpdtUserId.value  =  parent.gUsrID
           .txtInsrtUserId.value =  parent.gUsrID
	       .txtMaxRows.value     = lGrpCnt-1	
	       .txtSpread.value      = strDel & strVal
	    End With
    end if	

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
	If DbDELETE = False Then
		Call  RestoreToolBar()
		Exit Function
	End If																'☜: Query db data
    
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


    if  frm1.vspdData.MaxRows = 0 and frm1.vspdData1.MaxRows = 0 then
        Call  DisplayMsgBox("900014", "X", "X", "X")
    end if

	if  gSelframeFlg = TAB1 then
        Call SetToolbar("1100111100111111")
        frm1.vspdData.focus
    else
    	Call SetToolbar("1100100100011111")									
        frm1.vspdData1.focus    	
	end if
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    Call  ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field

    
    Call InitVariables															'⊙: Initializes local global variables

    Call MakeKeyStream("X")    
    lgCurrentSpd = "M"

	Call  DisableToolBar( parent.TBC_QUERY)
	If DbQuery = False Then
		Call  RestoreToolBar()
		Exit Function
	End If																'☜: Query db data
    
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
            Call InitComboBox("A")
		Case "vaSpread1"
			Call InitSpreadSheet("B")      		
            Call InitComboBox("B")
	End Select 
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

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
	    Case "1"
	        arrParam(0) = "수당코드팝업"			' 팝업 명칭 
	        arrParam(1) = "HDA010T"				 		' TABLE 명칭 
	        arrParam(2) = frm1.txtAllow_cd.value		    ' Code Condition
	        arrParam(3) = ""'frm1.txtAllow_nm.value		' Name Cindition
	        
	        If  gSelframeFlg = TAB1 Then
	            arrParam(4) = " PAY_CD = " & FilterVar("*", "''", "S") & "  AND CODE_TYPE = " & FilterVar("1", "''", "S") & "  " ' Where Condition
            else
                arrParam(4) = " PAY_CD = " & FilterVar("*", "''", "S") & "  AND CODE_TYPE = " & FilterVar("0", "''", "S") & "  " ' Where Condition
            end if
	        
	        arrParam(5) = "수당코드"			    ' TextBox 명칭 
	
            arrField(0) = "allow_cd"					' Field명(0)
            arrField(1) = "allow_nm"				    ' Field명(1)
    
            arrHeader(0) = "수당코드"				' Header명(0)
            arrHeader(1) = "수당코드명"			    ' Header명(1)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
		
	
	If arrRet(0) = "" Then
		frm1.txtAllow_cd.focus		
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
		    Case "1"
		        .txtAllow_cd.value = arrRet(0)
		        .txtAllow_nm.value = arrRet(1)	
		        .txtAllow_cd.focus	
        End Select
	End With
End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
         Case  C_PAY_CD_NM         ' 급여구분 
                iDx = Frm1.vspdData.value
   	            Frm1.vspdData.Col = C_PAY_CD
                Frm1.vspdData.value = iDx
         Case  C_ALLOW_KIND_NM     ' 세액종류 
                iDx = Frm1.vspdData.value
   	            Frm1.vspdData.Col = C_ALLOW_KIND
                Frm1.vspdData.value = iDx
         Case  C_TAX_TYPE_NM       ' 세액구분 
                iDx = Frm1.vspdData.value
   	            Frm1.vspdData.Col = C_TAX_TYPE
                Frm1.vspdData.value = iDx
         Case  C_CRT_STRT_MM_NM    ' 계산기간 
                iDx = Frm1.vspdData.value
   	            Frm1.vspdData.Col = C_CRT_STRT_MM
                Frm1.vspdData.value = iDx
         Case  C_CRT_END_MM_NM     ' 계산기간 
                iDx = Frm1.vspdData.value
   	            Frm1.vspdData.Col = C_CRT_END_MM
                Frm1.vspdData.value = iDx
         Case  C_DAY_CALCU_NM     ' 계산기간 
                iDx = Frm1.vspdData.value
   	            Frm1.vspdData.Col = C_DAY_CALCU
                Frm1.vspdData.value = iDx
         Case  C_CALCU_BAS_DD_NM     ' 계산기간 
                iDx = Frm1.vspdData.value
   	            Frm1.vspdData.Col = C_CALCU_BAS_DD
                Frm1.vspdData.value = iDx
         Case Else
    End Select    
             
   	If Frm1.vspdData.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData.text) <  UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	 ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row
End Sub

Sub vspdData1_Change(ByVal Col , ByVal Row )
    Dim iDx

   	Frm1.vspdData1.Row = Row
   	Frm1.vspdData1.Col = Col

    Select Case Col
         Case  C_TAX_TYPE_NM1       ' 세액구분 
                iDx = Frm1.vspdData1.value
   	            Frm1.vspdData1.Col = C_TAX_TYPE1
                Frm1.vspdData1.value = iDx
         Case  C_CRT_STRT_MM_NM1    ' 계산기간 
                iDx = Frm1.vspdData1.value
   	            Frm1.vspdData1.Col = C_CRT_STRT_MM1
                Frm1.vspdData1.value = iDx
         Case  C_CRT_END_MM_NM1     ' 계산기간 
                iDx = Frm1.vspdData1.value
   	            Frm1.vspdData1.Col = C_CRT_END_MM1
                Frm1.vspdData1.value = iDx
         Case  C_DAY_CALCU_NM1     ' 계산기간 
                iDx = Frm1.vspdData1.value
   	            Frm1.vspdData1.Col = C_DAY_CALCU1
                Frm1.vspdData1.value = iDx
         Case  C_CALCU_BAS_DD_NM1     ' 계산기간 
                iDx = Frm1.vspdData1.value
   	            Frm1.vspdData1.Col = C_CALCU_BAS_DD1
                Frm1.vspdData1.value = iDx
         Case Else
    End Select    
             
   	If Frm1.vspdData1.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData1.text) <  UNICDbl(Frm1.vspdData1.TypeFloatMin) Then
         Frm1.vspdData1.text = Frm1.vspdData1.TypeFloatMin
      End If
	End If
	
	 ggoSpread.Source = frm1.vspdData1
     ggoSpread.UpdateRow Row
End Sub
'========================================================================================================
'   Event Name : vspdData_MouseDown
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And  gMouseClickStatus = "SPC" Then
        gMouseClickStatus = "SPCR"
     End If
End Sub

Sub vspdData1_MouseDown(Button , Shift , x , y)

    If Button = 2 And  gMouseClickStatus = "SP1C" Then
        gMouseClickStatus = "SP1CR"
     End If
End Sub    
'========================================================================================================
'   Event Name : vspdData_Click
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101111111")
    gMouseClickStatus = "SPC" 
    Set gActiveSpdSheet = frm1.vspdData

	if frm1.vspddata.MaxRows <= 0 then
		exit sub
	end if
	
	if Row <=0 then
		ggoSpread.Source = frm1.vspdData
		if lgSortkey = 1 then
			ggoSpread.SSSort Col
			lgSortKey = 2
		else
			ggoSpread.SSSort Col, lgSortkey
			lgSortKey = 1
		end if
		Exit sub
	end if
	frm1.vspdData.Row = Row     
End Sub

Sub vspdData1_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("0001111111")
    gMouseClickStatus = "SP1C" 
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

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)	
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
    Dim iColumnName
    if Row <= 0 then
		exit sub
	end if
	if Frm1.vspdData.MaxRows = 0 then
		exit sub
	end if
End Sub

Sub vspdData1_DblClick(ByVal Col, ByVal Row)
    Dim iColumnName
    if Row <= 0 then
		exit sub
	end if
	if Frm1.vspdData1.MaxRows = 0 then
		exit sub
	end if
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData1_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
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
'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	If frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then

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
	End If  
End Sub
Function ClickTab1()
	Dim IntRetCD
	If gSelframeFlg = TAB1 Then Exit Function

	ggoSpread.Source = frm1.vspdData1
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X") '☜ 바뀐부분 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	
	Call changeTabs(TAB1)
    frm1.txtAllow_cd.value = ""
    frm1.txtAllow_nm.value = ""
	gSelframeFlg = TAB1
    lgCurrentSpd = "M"          
    gMouseClickStatus = "SPC"             
    if  frm1.vspdData.MaxRows > 0 then
        Call SetToolbar("1100111100111111")
    else
        Call SetToolbar("1100110100111111")
    end if
    Set gActiveSpdSheet = frm1.vspdData
End Function

Function ClickTab2()
	If gSelframeFlg = TAB2 Then Exit Function
	
	Call changeTabs(TAB2)
	frm1.txtAllow_cd.value = ""
    frm1.txtAllow_nm.value = ""
	gSelframeFlg = TAB2
    lgCurrentSpd = "S"        
    Call SetToolbar("1100100100011111")

End Function
'========================================================================================================
'   Event Name : txtAllow_cd_change
'   Event Desc :
'========================================================================================================
Function txtAllow_cd_Onchange()
    Dim IntRetCd , wheretype
    
    If frm1.txtAllow_cd.value = "" Then
		frm1.txtAllow_nm.value = ""
    Else
        
        If  gSelframeFlg = TAB1 Then
	        wheretype = " PAY_CD = " & FilterVar("*", "''", "S") & "  AND CODE_TYPE = " & FilterVar("1", "''", "S") & "  " ' Where Condition
        Else
            wheretype = " PAY_CD = " & FilterVar("*", "''", "S") & "  AND CODE_TYPE = " & FilterVar("0", "''", "S") & "  " ' Where Condition
        End if
   
        IntRetCd =  CommonQueryRs(" ALLOW_NM "," HDA010T ",wheretype & " and ALLOW_CD =  " & FilterVar(frm1.txtAllow_cd.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
   
        If IntRetCd = false then
			 frm1.txtAllow_nm.value = ""
             frm1.txtAllow_cd.focus
            Set gActiveElement = document.ActiveElement
            Exit Function          
        Else
			frm1.txtAllow_nm.value = Trim(Replace(lgF0,Chr(11),""))
        End if 
    End if  
End Function


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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>급여지급내역등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>기타지급내역등록</font></td>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
						<TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
						        <TD CLASS="TD5" NOWRAP>수당코드</TD>
						        <TD CLASS="TD6" NOWRAP><INPUT ID=txtAllow_cd NAME="txtAllow_cd" MAXLENGTH=3 SIZE=10  ALT ="수당코드" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCondAreaPopup('1')">
						                               <INPUT ID=txtAllow_nm NAME="txtAllow_nm" MAXLENGTH=20 SIZE=20  ALT ="수당코드명" tag="14XXXU"></TD>
        						<TD CLASS="TDT" NOWRAP></TD>
	                        	<TD CLASS="TD6" NOWRAP></TD>
							</TR>
						</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
					<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</DIV>
					<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData1 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</DIV>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>  
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
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

