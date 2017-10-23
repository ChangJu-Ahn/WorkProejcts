<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : Multi Sample
*  3. Program ID           : HN101BA1_KO441
*  4. Program Name         : HN101BA1_KO441
*  5. Program Desc         : 급/상여내역UPLOAD
*  6. Comproxy List        :
*  7. Modified date(First) : 2008/01/09
*  8. Modified date(Last)  : 2008/01/09
*  9. Modifier (First)     : 
* 10. Modifier (Last)      : 
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
 
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID      = "HN101BB1_KO441.asp"	
'Const BIZ_PGM_ID1     = "HN101BB1_KO441.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_ID2     = "HN101BB2_KO441.asp"						           '☆: Biz Logic ASP Name
Const CookieSplit = 1233
'Const C_SHEETMAXROWS    =   21	                                      '한 화면에 보여지는 최대갯수*1.5%>
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd
Dim IsOpenPop 
Dim lgStrComDateType		'Company Date Type을 저장(년월 Mask에 사용함.)
Dim lgType
Dim	lgChgStatus	'입력, 수정, 삭제가 발생한 이력이 있는지 확인

'Dim C_DEPT_CD			'부서코드
'Dim C_EMP_NO			'사번
'Dim C_PAY_YYMM			'해당년월	
'Dim C_PROV_TYPE			'지급구분 
'Dim C_PROV_TYPE_HIDDEN	'지급구분
'Dim C_PROV_DT			'지급일			
'Dim	C_PAY_TOT_AMT		'급여총액
'Dim	C_BONUS_TOT_AMT		'상여총액
'Dim C_NONTAX_TOT_AMT	'비과세총액
'Dim C_TAX_AMT			'급여과세총액
'Dim C_BONUS_TAX			'상여과세총액
'Dim C_PROV_TOT_AMT		'지급총액
'Dim C_SUB_TOT_AMT		'제공제계
'Dim C_REAL_PROV_AMT		'실지급액
'Dim C_INCOME_TAX		'소득세	
'Dim C_RES_TAX			'주민세
'Dim C_ANUT				'국민연금
'Dim C_MED_INSURE		'건강보험
'Dim C_EMP_INSURE		'고용보험
''----------------------------------
'Dim C_ALLOW_CD			'수당코드
'Dim C_ALLOW_NM			'수당명
'Dim C_ALLOW				'수당금액
''----------------------------------
'Dim C_SUB_CD			'공제코드
'Dim C_SUB_NM			'공제명
'Dim C_SUB_AMT			'공제금액
'
''========================================================================================================
'' Name : InitSpreadPosVariables()
'' Desc : Initialize value
''========================================================================================================
'Sub InitSpreadPosVariables(ByVal pvSpdNo)	 
'
'	Select Case pvSpdNo
'		   Case "A"
'				C_DEPT_CD			=1			'부서코드
'				C_EMP_NO			=2			'사번
'				C_PAY_YYMM			=3			'해당년월	
'				C_PROV_TYPE			=4			'지급유형 
'				C_PROV_TYPE_HIDDEN	=5			'지급유형코드
'				C_PROV_DT			=6			'지급일			
'				C_PAY_TOT_AMT		=7			'급여총액
'				C_BONUS_TOT_AMT		=8			'상여총액
'				C_NONTAX_TOT_AMT	=9			'비과세총액
'				C_TAX_AMT			=10			'과세총액 (C_TAX_AMT : 급여과세총액, C_BONUS_TAX : 상여과세총액)	
'				C_PROV_TOT_AMT		=11			'지급총액
'				C_SUB_TOT_AMT		=12			'제공제계
'				C_REAL_PROV_AMT		=13			'실지급액
'				C_INCOME_TAX		=14			'소득세	
'				C_RES_TAX			=15			'주민세
'				C_ANUT				=16			'국민연금
'				C_MED_INSURE		=17			'건강보험
'				C_EMP_INSURE		=18			'고용보험
'		   Case "B"				
'				C_PAY_YYMM			=1			'해당년월
'				C_EMP_NO			=2			'사번
'				C_PROV_TYPE			=3			'지급유형
'				C_PROV_TYPE_HIDDEN	=4			'지급유형코드
'				C_ALLOW_CD			=5			'수당코드
'				C_ALLOW_NM			=6			'수당명
'				C_ALLOW				=7			'수당금액
'
'		   Case "C"
'				C_PAY_YYMM			=1			'해당년월
'				C_EMP_NO			=2			'사번
'				C_PROV_TYPE			=3			'지급유형
'				C_PROV_TYPE_HIDDEN	=4			'지급유형코드
'				C_SUB_CD			=5			'공제코드
'				C_SUB_NM			=6			'공제명
'				C_SUB_AMT			=7			'공제금액
'
'	End Select	
'
'End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False											'⊙: Indicates that no value changed
	lgIntGrpCount     = 0												'⊙: Initializes Group View Size
    lgStrPrevKey      = ""												'⊙: initializes Previous Key
    lgSortKey         = 1												'⊙: initializes sort direction
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

	With frm1
		.txtDt.Year = strYear 
		.txtDt.Month = strMonth 
		.txtDt.day = strDay  
				
		.txtYYMM.Year = strYear
		.txtYYMM.Month = strMonth

		.txtFileName2.value = ""
		.txtSpread.value = ""

		Call  ggoOper.FormatDate(.txtDt,  parent.gDateFormat, 1) 
		Call  ggoOper.FormatDate(.txtYYMM,  parent.gDateFormat, 2) 	
		.txtFileName2.focus
	End With

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
 
	With Frm1
		lgKeyStream  = Trim(.txtDt.text) & parent.gColSep
		lgKeyStream  = lgKeyStream & Trim(.txtYYMM.text) & parent.gColSep   
		lgKeyStream  = lgKeyStream & Trim(.txtProv_cd.Value) & parent.gColSep   		
		lgKeyStream  = lgKeyStream & Trim(.hFileName.Value) & parent.gColSep
		
'		If .rdoCase1.Checked = True Then		'파일구분
'			lgKeyStream  = lgKeyStream & "0" & parent.gColSep
'		elseIf .rdoCase2.Checked= True Then
'			lgKeyStream  = lgKeyStream & "1" & parent.gColSep
'		Else
'			lgKeyStream  = lgKeyStream & "2" & parent.gColSep
'		End If
	End With
		
End Sub        



''========================================================================================================
'' Function Name : InitSpreadSheet
'' Function Desc : This method initializes spread sheet column property
''========================================================================================================
'Sub InitSpreadSheet()
'
'	Call initSpreadPosVariables(lgType)  
'
'	With frm1.vspdData
'	
'        ggoSpread.Source = frm1.vspdData
'		ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread    
'	   .ReDraw = false
'		Select Case lgType
'			   Case "A"
'					.MaxCols = C_EMP_INSURE + 1											 ' ☜:☜: Add 1 to Maxcols
'			   Case "B"
'					.MaxCols = C_ALLOW + 1                                                      
'			   Case "C"
'					.MaxCols = C_SUB_AMT + 1     
'		End Select
'
'	   .Col = .MaxCols																	' ☜:☜: Hide maxcols
'       .ColHidden = True																' ☜:☜:
'       .MaxRows = 0
'	   
'		Call GetSpreadColumnPos(lgType)  
'		
'		Select Case lgType
'			   Case "A"	
'					'Call AppendNumberPlace("6","2","0")								
'					                                             
'					ggoSpread.SSSetEdit C_DEPT_CD			, "부서코드"		, 10,,,40,2		'Lock/ Edit		
'					ggoSpread.SSSetEdit C_EMP_NO			, "사번"			, 10,,,13,2		'Lock/ Edit
'					ggoSpread.SSSetEdit C_PAY_YYMM			, "해당년월"		, 12,,,13,2		'Lock/ Edit
'					ggoSpread.SSSetEdit C_PROV_TYPE			, "지급유형"		, 12,,,50,2		'Lock/ Edit
'					ggoSpread.SSSetEdit C_PROV_TYPE_HIDDEN  , "지급유형Code"	, 12,,,1,2		'Lock/ Edit
'					ggoSpread.SSSetEdit C_PROV_DT			, "지급일"			, 12,,,13,2		'Lock/ Edit
'					ggoSpread.SSSetFloat C_PAY_TOT_AMT		, "급여총액"		,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
'					ggoSpread.SSSetFloat C_BONUS_TOT_AMT	, "상여총액"		,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
'					ggoSpread.SSSetFloat C_NONTAX_TOT_AMT	, "비과세총액"		,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
'					ggoSpread.SSSetFloat C_TAX_AMT			, "과세총액"		,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
'					'ggoSpread.SSSetFloat C_BONUS_TAX		, "상여과세총액"	,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
'					ggoSpread.SSSetFloat C_PROV_TOT_AMT		, "지급총액"		,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
'					ggoSpread.SSSetFloat C_SUB_TOT_AMT		, "공제총액"		,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
'					ggoSpread.SSSetFloat C_REAL_PROV_AMT	, "실지급액"		,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
'					
'					ggoSpread.SSSetFloat C_INCOME_TAX		, "소득세"			,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
'					ggoSpread.SSSetFloat C_RES_TAX			, "주민세"			,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
'					ggoSpread.SSSetFloat C_ANUT				, "국민연금"		,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
'					ggoSpread.SSSetFloat C_MED_INSURE		, "건강보험"		,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
'					ggoSpread.SSSetFloat C_EMP_INSURE		, "고용보험"		,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec				
'				
'					Call ggoSpread.SSSetColHidden(C_PROV_TYPE_HIDDEN,C_PROV_TYPE_HIDDEN,True)
'
'			   Case "B"									
'					ggoSpread.SSSetEdit  C_EMP_NO			, "사번"			, 10,,,13,2		'Lock/ Edit
'					ggoSpread.SSSetEdit  C_PAY_YYMM			, "해당년월"		, 12,,,13,2		'Lock/ Edit
'					ggoSpread.SSSetEdit  C_PROV_TYPE		, "지급유형"		, 12,,,50,2		'Lock/ Edit
'					ggoSpread.SSSetEdit  C_PROV_TYPE_HIDDEN , "지급유형Code"	, 12,,,1,2		'Lock/ Edit
'					ggoSpread.SSSetEdit  C_ALLOW_CD			, "수당코드"		, 12,,,50,2		'Lock/ Edit
'					ggoSpread.SSSetEdit  C_ALLOW_NM			, "수당명"			, 12,,,50,2		'Lock/ Edit
'					ggoSpread.SSSetFloat C_ALLOW			, "수당금액"		,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec	
'					
'					Call ggoSpread.SSSetColHidden(C_PROV_TYPE_HIDDEN,C_PROV_TYPE_HIDDEN,True)
'					
'			   Case "C"
'					ggoSpread.SSSetEdit  C_EMP_NO			, "사번"			, 10,,,13,2		'Lock/ Edit
'					ggoSpread.SSSetEdit  C_PAY_YYMM			, "해당년월"		, 12,,,13,2		'Lock/ Edit
'					ggoSpread.SSSetEdit  C_PROV_TYPE		, "지급유형"		, 12,,,50,2		'Lock/ Edit
'					ggoSpread.SSSetEdit  C_PROV_TYPE_HIDDEN , "지급유형Code"	, 12,,,1,2		'Lock/ Edit
'					ggoSpread.SSSetEdit  C_ALLOW_CD			, "공제코드"		, 12,,,50,2		'Lock/ Edit
'					ggoSpread.SSSetEdit  C_ALLOW_NM			, "공제명"			, 12,,,50,2		'Lock/ Edit
'					ggoSpread.SSSetFloat C_ALLOW			, "공제금액"		,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec	
'					
'					Call ggoSpread.SSSetColHidden(C_PROV_TYPE_HIDDEN,C_PROV_TYPE_HIDDEN,True)
'
'		End Select
'		.Redraw = True 
'		 
'        
'    End With
'	
'	ggoSpread.Source = frm1.vspdData
'    ggoSpread.SpreadLockWithOddEvenRowColor()
'	
'End Sub

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
'       For iDx = 1 To  frm1.vspdData.MaxCols - 1
'           Frm1.vspdData.Col = iDx
'           Frm1.vspdData.Row = iRow
'           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <>  parent.UC_PROTECTED Then
'              Frm1.vspdData.Col = iDx
'              Frm1.vspdData.Row = iRow
'              Frm1.vspdData.Action = 0 ' go to 
'              Exit For
'           End If
'           
'       Next
          
    End If   
End Sub

''========================================================================================
'' Function Name : GetSpreadColumnPos
'' Description   : 
''========================================================================================
'Sub GetSpreadColumnPos(ByVal pvSpdNo)
'    
'	Dim iCurColumnPos
'    
'	ggoSpread.Source = frm1.vspdData
'	Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
'	
'    Select Case UCase(pvSpdNo)
'		   Case "A"
'				C_DEPT_CD			=iCurColumnPos(1)			'부서코드
'				C_EMP_NO			=iCurColumnPos(2)			'사번
'				C_PAY_YYMM			=iCurColumnPos(3)			'해당년월	
'				C_PROV_TYPE			=iCurColumnPos(4)			'지급유형 
'				C_PROV_TYPE_HIDDEN	=iCurColumnPos(5)			'지급유형 
'				C_PROV_DT			=iCurColumnPos(6)			'지급일			
'				C_PAY_TOT_AMT		=iCurColumnPos(7)			'급여총액
'				C_BONUS_TOT_AMT		=iCurColumnPos(8)			'상여총액
'				C_NONTAX_TOT_AMT	=iCurColumnPos(9)			'비과세총액
'				C_TAX_AMT			=iCurColumnPos(10)			'과세총액(C_TAX_AMT:급여과세총액, C_BONUS_TAX : 상여과세총액)			
'				C_PROV_TOT_AMT		=iCurColumnPos(11)			'지급총액
'				C_SUB_TOT_AMT		=iCurColumnPos(12)			'제공제계
'				C_REAL_PROV_AMT		=iCurColumnPos(13)			'실지급액
'				C_INCOME_TAX		=iCurColumnPos(14)			'소득세	
'				C_RES_TAX			=iCurColumnPos(15)			'주민세
'				C_ANUT				=iCurColumnPos(16)			'국민연금
'				C_MED_INSURE		=iCurColumnPos(17)			'건강보험
'				C_EMP_INSURE		=iCurColumnPos(18)			'고용보험
'
'		   Case "B"
'				C_PAY_YYMM			=iCurColumnPos(1)			'해당년월
'				C_EMP_NO			=iCurColumnPos(2)			'사번
'				C_PROV_TYPE			=iCurColumnPos(3)			'지급유형
'				C_PROV_TYPE_HIDDEN	=iCurColumnPos(4)			'지급유형코드
'				C_ALLOW_CD			=iCurColumnPos(5)			'수당코드
'				C_ALLOW_NM			=iCurColumnPos(6)			'수당명
'				C_ALLOW				=iCurColumnPos(7)			'수당금액
'
'		   Case "C"
'				C_PAY_YYMM			=iCurColumnPos(1)			'해당년월
'				C_EMP_NO			=iCurColumnPos(2)			'사번
'				C_PROV_TYPE			=iCurColumnPos(3)			'지급유형
'				C_PROV_TYPE_HIDDEN	=iCurColumnPos(4)			'지급유형코드
'				C_SUB_CD			=iCurColumnPos(5)			'공제코드
'				C_SUB_NM			=iCurColumnPos(6)			'공제명
'				C_SUB_AMT			=iCurColumnPos(7)			'공제금액
'
'			
'    End Select    
'End Sub
''======================================================================================================
'' Function Name : vspdData_ScriptLeaveCell
'' Function Desc : 년(YYYY).월(MM) check
''======================================================================================================
'Sub vspdData_ScriptLeaveCell(Col, Row, NewCol, NewRow, Cancel)
'End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear																			'☜: Clear err status
	Call LoadInfTB19029																	'⊙: Load table , B_numeric_format
		
    Call  ggoOper.FormatField(Document, "A", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	Call  ggoOper.LockField(Document, "N")												'⊙: Lock Field
   
	lgType = "A"
    'Call  InitSpreadSheet																'Setup the Spread sheet
	
    Call  InitVariables																	'Initializes local global variables

    Call  FuncGetAuth(gStrRequestMenuID ,  parent.gUsrID, lgUsrIntCd)					' 자료권한:lgUsrIntCd ("%", "1%")

    Call SetDefaultVal
	Call SetToolbar("1100000000001111")													'버튼 툴바 제어

	Call CookiePage (0)
   
End Sub
	
'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

''========================================================================================================
'' Name : FncQuery
'' Desc : developer describe this line Called by MainQuery in Common.vbs
''========================================================================================================
'Function FncQuery()	
'
'    Dim IntRetCD 
'    Dim strwhere
'
'    FncQuery = False															 '☜: Processing is NG
'    Err.Clear                                                                    '☜: Clear err status
'
'    ggoSpread.Source = Frm1.vspdData
'    If  ggoSpread.SSCheckChange = True Then
'		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"x","x")			'☜: Data is changed.  Do you want to display it? 
'		If IntRetCD = vbNo Then
'			Exit Function
'		End If
'    End If
'    
'    ggoSpread.ClearSpreadData
'
'	If  txtprov_cd_Onchange() then
'        Exit Function
'    End If    
'
'
'    Call InitVariables                                                           '⊙: Initializes local global variables
'    Call MakeKeyStream("X")
'	
'    If DbQuery = False Then
'       Exit Function	
'    End If																		 '☜: Query db data
'       
'    FncQuery = True                                                              '☜: Processing is OK
'    
'End Function

'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()

	Dim IntRetCD 
    
    FncDelete = False																'☜: Processing is NG
    
    Err.Clear																		'☜: Clear err status
    
	IF lgType = "A" Then
		IF frm1.txtDT.Text="" then 
			Call  DisplayMsgBox("970021","X","급/상여일자","X")							'급/상여일자를 확인하십시오.    
			frm1.txtDT.focus 
			Exit Function    
		End If
	End If

	IF frm1.txtYYMM.Text = "" then 
	    Call  DisplayMsgBox("970021","X","급/상여년월","X")							'급/상여년월를 확인하십시오.    
	    frm1.txtYYMM.focus 
	    Exit Function    
	End If

	'If lgIntFlgMode <>  parent.OPMD_UMODE Then										 'Check if there is retrived data
    '    Call  DisplayMsgBox("900002","X","X","X")									'☆:
    '    Exit Function
    'End If

    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"X","X")					'⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then															'------ Delete function call area ------ 
		Exit Function	
	End If
       
    Call MakeKeyStream("X")
	
	'Call DisableToolBar( parent.TBC_DELETE)
	
	If DbDelete = False Then
		'Call  RestoreToolBar()
		Exit Function
	End If																		'☜: Query db data
    
    FncDelete = True                                                              '☜: Processing is OK

End Function

''========================================================================================================
'' Name : FncSave
'' Desc : developer describe this line Called by MainSave in Common.vbs
''========================================================================================================
'Function FncSave()
'	
'    Dim IntRetCD 
'    
'    FncSave = False                                                              '☜: Processing is NG
'    
'    Err.Clear                                                                    '☜: Clear err status
'    
'     ggoSpread.Source = frm1.vspdData
'    If  ggoSpread.SSCheckChange = False Then
'        IntRetCD =  DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
'        Exit Function
'    End If
'    
'     ggoSpread.Source = frm1.vspdData
'    If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
'       Exit Function
'    End If
'     	 
'	
'	If lgType = "A" Then
'		IF frm1.txtDT.Text="" then 
'			Call  DisplayMsgBox("970021","X","급/상여일자","X")						'급/상여일자를 확인하십시오.    
'			frm1.txtDT.focus 
'			Exit Function    
'		End If
'	End If
'
'	IF frm1.txtYYMM.Text = "" then 
'	    Call  DisplayMsgBox("970021","X","급/상여년월","X")						'급/상여년월를 확인하십시오.    
'	    frm1.txtYYMM.focus 
'	    Exit Function    
'	End If
'
'	Dim lRow
'
'	'조건부의 급/상여일자와 급/상여년월데이타와 Excel 데이타가 일치하지 않으면 실행되지 않음
'	With Frm1
'		
'        For lRow = 1 To .vspdData.MaxRows
'			
'            .vspdData.Row = lRow
'            .vspdData.Col = 0
'			
'			Select Case lgType
'				   Case "A"
'						.vspdData.Col = C_PAY_YYMM
'						If Trim(Replace(.txtYYMM,"-","")) <> Trim(.vspdData.Text) Then	
'							MsgBox "급/상여년월 데이타가 일치하지 앖습니다.    ", vbExclamation, "uniERPII[Warning]"							
'							Exit Function
'						End If
'					
'						.vspdData.Col = C_PROV_DT
'						If Trim(.txtDt) <> Trim(.vspdData.Text) Then
'							MsgBox "급/상여일자 데이타가 일치하지 앖습니다.    ", vbExclamation, "uniERPII[Warning]"							
'							Exit Function
'						End If
'
'						.vspdData.Col = C_PROV_TYPE
'						If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
'							Call DisplayMsgBox("970000","X","지급유형","X")
'							Exit Function
'						End If
'
'						.vspdData.Col = C_EMP_NO
'						If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
'							Call DisplayMsgBox("970000","X","사원번호","X")
'							Exit Function
'						End If	
'
'				   Case "B"
'						
'						.vspdData.Col = C_PAY_YYMM
'						If Trim(Replace(.txtYYMM,"-","")) <> Trim(.vspdData.Text) Then				
'							MsgBox "급/상여년월 데이타가 일치하지 앖습니다.    ", vbExclamation, "uniERPII[Warning]"							
'							Exit Function
'						End If
'					
'						.vspdData.Col = C_PROV_TYPE
'						If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
'							Call DisplayMsgBox("970000","X","지급유형","X")
'							Exit Function
'						End If
'
'						.vspdData.Col = C_EMP_NO
'						If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
'							Call DisplayMsgBox("970000","X","사원번호","X")
'							Exit Function
'						End If	
'
'						.vspdData.Col = C_ALLOW_CD
'						If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
'							Call DisplayMsgBox("970000","X","수당코드","X")
'							Exit Function
'						End If	
'
'				   Case "C"
'
'						.vspdData.Col = C_PAY_YYMM
'						If Trim(Replace(.txtYYMM,"-","")) <> Trim(.vspdData.Text) Then				
'							MsgBox "급/상여년월 데이타가 일치하지 앖습니다.    ", vbExclamation, "uniERPII[Warning]"							
'							Exit Function
'						End If
'					
'						.vspdData.Col = C_PROV_TYPE
'						If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
'							Call DisplayMsgBox("970000","X","지급유형","X")
'							Exit Function
'						End If
'
'						.vspdData.Col = C_EMP_NO
'						If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
'							Call DisplayMsgBox("970000","X","사원번호","X")
'							Exit Function
'						End If	
'
'						.vspdData.Col = C_SUB_CD
'						If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
'							Call DisplayMsgBox("970000","X","공제코드","X")
'							Exit Function
'						End If	
'
'			End Select
'			
'					            
'        Next
'
'	End With
'       
'    Call MakeKeyStream("X")
'
'	'Call DisableToolBar( parent.TBC_SAVE)
'	
'	If DbSAVE = False Then
'		Call  RestoreToolBar()
'		Exit Function
'	End If																		'☜: Query db data
'    
'    FncSave = True                                                              '☜: Processing is OK
'    
'End Function

''========================================================================================================
'' Function Name : FncCopy
'' Function Desc : This function is related to Copy Button of Main ToolBar
''========================================================================================================
'Function FncCopy()
'
'    If Frm1.vspdData.MaxRows < 1 Then
'       Exit Function
'    End If
'    
'     ggoSpread.Source = Frm1.vspdData
'	With Frm1.VspdData
'         .ReDraw = False
'		 If .ActiveRow > 0 Then
'             ggoSpread.CopyRow
'			 SetSpreadColor .ActiveRow, .ActiveRow
'    
'            .ReDraw = True
'		    .Focus
'		 End If
'	End With
'	
'    Set gActiveElement = document.ActiveElement   
'
'End Function
'
''========================================================================================================
'' Function Name : FncCancel
'' Function Desc : This function is related to Cancel Button of Main ToolBar
''========================================================================================================
'Function FncCancel() 
'     ggoSpread.Source = frm1.vspdData	
'     ggoSpread.EditUndo  
'End Function

''========================================================================================================
'' Function Name : FncInsertRow
'' Function Desc : This function is related to InsertRow Button of Main ToolBar
''========================================================================================================
'Function FncInsertRow(ByVal pvRowCnt)
'	Dim imRow	
'
'    On Error Resume Next                                                          '☜: If process fails
'    Err.Clear                                                                     '☜: Clear error status
' 
'    FncInsertRow = False                                                         '☜: Processing is NG
'
'    If IsNumeric(Trim(pvRowCnt)) Then
'        imRow = CInt(pvRowCnt)
'    Else
'        imRow = AskSpdSheetAddRowCount()
'        If imRow = "" Then
'            Exit Function
'        End If
'    End If
'
'	With frm1
'        .vspdData.ReDraw = False
'        .vspdData.focus
'        ggoSpread.Source = .vspdData
'        ggoSpread.InsertRow .vspdData.ActiveRow, imRow
'        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1         
'        
'       .vspdData.ReDraw = True
'    End With
'
'    If Err.number = 0 Then
'       FncInsertRow = True                                                          '☜: Processing is OK
'    End If   
'    
'    Set gActiveElement = document.ActiveElement   
'End Function
'
''========================================================================================================
'' Function Name : FncDeleteRow
'' Function Desc : This function is related to DeleteRow Button of Main ToolBar
''========================================================================================================
'Function FncDeleteRow() 
'    Dim lDelRows
'    If Frm1.vspdData.MaxRows < 1 then
'       Exit function
'	End if	
'    With Frm1.vspdData 
'    	.focus
'    	 ggoSpread.Source = frm1.vspdData 
'    	lDelRows =  ggoSpread.DeleteRow
'    End With
'    Set gActiveElement = document.ActiveElement   
'End Function

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
	Call Parent.FncExport( parent.C_MULTI)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind( parent.C_MULTI, False)
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
    'Call InitSpreadSheet()          
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
	
'     ggoSpread.Source = frm1.vspdData	
'    If  ggoSpread.SSCheckChange = True Then
'		IntRetCD =  DisplayMsgBox("900016",  parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
'		If IntRetCD = vbNo Then
'			Exit Function
'		End If
'    End If
    FncExit = True
End Function
'
''========================================================================================================
'' Name : DbQuery
'' Desc : This function is called by FncQuery
''========================================================================================================
'Function DbQuery() 	
'
'    DbQuery = False
'    
'    Err.Clear                                                                        '☜: Clear err status
'
'	if LayerShowHide(1) = false then
'		exit Function
'	end if
'
'	Dim strVal	
'   
'    DbQuery = False
'    
'    Err.Clear                                                                        '☜: Clear err status
'
'	
'	if LayerShowHide(1) = false then		
'		exit Function
'	end if
'		    
'    With Frm1
'		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
'        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
'        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
'        strVal = strVal     & "&lgStrPrevKey="		 & lgStrPrevKey                 '☜: Next key tag
'		strVal = strVal     & "&htxtFileGubun="		 & lgType	
'    End With
'		
'    If lgIntFlgMode =  parent.OPMD_UMODE Then
'    Else
'    End If
'	
'	
'	Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic
'    
'    DbQuery = True
'    
'End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave() 
	
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim lStartRow   
    Dim lEndRow     
	Dim strVal, strDel
	
    DbSave = False                                                          
	
	If LayerShowHide(1) = false then		
		exit Function
	End if
	
	With frm1
		.txtMode.value      = parent.UID_M0002                                        '☜: Save
		.txtFlgMode.value   = lgIntFlgMode
	End With

    strVal = ""
    strDel = ""
    'lGrpCnt = 1

	
	With Frm1

       'For lRow = 1 To .vspdData.MaxRows
    
           '.vspdData.Row = lRow
           '.vspdData.Col = 0
     
           'Select Case .vspdData.Text
                  'Case  ggoSpread.InsertFlag																	'☜: Insert                  
						'									strVal = strVal & "C" & parent.gColSep				'array(0)
						'									strVal = strVal & lRow & parent.gColSep
						'Select Case lgType
						'	   Case "A"
'									.vspdData.Col = C_DEPT_CD				: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'									.vspdData.Col = C_EMP_NO				: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'									.vspdData.Col = C_PAY_YYMM				: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'									.vspdData.Col = C_PROV_TYPE_HIDDEN		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'									.vspdData.Col = C_PROV_DT				: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'									.vspdData.Col = C_PAY_TOT_AMT			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'									.vspdData.Col = C_BONUS_TOT_AMT			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'									.vspdData.Col = C_NONTAX_TOT_AMT		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'									.vspdData.Col = C_TAX_AMT				: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'									.vspdData.Col = C_PROV_TOT_AMT			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'									.vspdData.Col = C_SUB_TOT_AMT			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'									.vspdData.Col = C_REAL_PROV_AMT			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'									.vspdData.Col = C_INCOME_TAX			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'									.vspdData.Col = C_RES_TAX				: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'									.vspdData.Col = C_ANUT					: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'									.vspdData.Col = C_MED_INSURE			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'									.vspdData.Col = C_EMP_INSURE			: strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep

						'	   Case "B"
									
'									.vspdData.Col = C_PAY_YYMM				: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'									.vspdData.Col = C_EMP_NO				: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep				
'									.vspdData.Col = C_PROV_TYPE_HIDDEN		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'									.vspdData.Col = C_ALLOW_CD				: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'									.vspdData.Col = C_ALLOW					: strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep									
						'	   Case "C"
'									.vspdData.Col = C_PAY_YYMM				: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'									.vspdData.Col = C_EMP_NO				: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep				
'									.vspdData.Col = C_PROV_TYPE_HIDDEN		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'									.vspdData.Col = C_SUB_CD				: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'									.vspdData.Col = C_SUB_AMT				: strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep									
						'End Select
                   
               
                    'lGrpCnt = lGrpCnt + 1 
			'End Select
       'Next
	   
	   .htxtDt.value		= .txtDt.text
	   .htxtYYMM.value		= .txtYYMM.text
	   .htxtProvCD.value	= .txtProv_CD.Value
	   .htxtFileGubun.value = lgType		
	   
	   '.txtMaxRows.value    = lGrpCnt-1	
	   '.txtSpread.value     = strDel & strVal
'MsgBox .txtSpread.value			
'MsgBox .txtMaxRows.value
	End With
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)		
	
    DbSave = True                                                           
    
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()   
    
	Dim strVal

	DbDelete = False																'⊙: Processing is NG
    
	If LayerShowHide(1) = false then		
		Exit Function
	End if

	With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0003						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                  '☜: Query Key
        strVal = strVal     & "&txtMaxRows=0"         '& .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey="		 & lgStrPrevKey					'☜: Next key tag
		strVal = strVal     & "&htxtFileGubun="		 & lgType	
    End With	

	Call RunMyBizASP(MyBizASP, strVal)
																					'☜: Query db data
    DbDelete = True																	'⊙: Processing is OK

End Function

''========================================================================================================
'' Function Name : DbQueryOk
'' Function Desc : Called by MB Area when query operation is successful
''========================================================================================================
'Function DbQueryOk()
'	
'    lgIntFlgMode =  parent.OPMD_UMODE    
'    Call  ggoOper.LockField(Document, "Q")										'⊙: Lock field
'    Call InitData()
'	Call SetToolbar("110011110011111")	 
'	'Call SetToolbar("1100000000001111")	
'	frm1.vspdData.focus
'
'End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    'ggoSpread.Source = frm1.vspdData	
	'ggoSpread.ClearSpreadData
	'Call RemovedivTextArea	
    Call InitVariables															'⊙: Initializes local global variables
	'ggoSpread.ClearSpreadData
    
    If lgChgStatus > 0 Then
    	Call DisplayMsgBox("183114","X","X","X")
    Else
    	Call DisplayMsgBox("800161","X","X","X")    	
    End If

End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	'ggoSpread.Source = frm1.vspdData	
	'ggoSpread.ClearSpreadData
	'Call RemovedivTextArea	
    Call InitVariables															'⊙: Initializes local global variables
	'ggoSpread.ClearSpreadData
    
    Call DisplayMsgBox("183114","X","X","X")
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
            arrParam(0) = "지급구분 팝업"				' 팝업 명칭 
	        arrParam(1) = "B_MINOR"				 		' TABLE 명칭 
	        arrParam(2) = frm1.txtprov_cd.value		    ' Code Condition
	        arrParam(3) = ""'frm1.txtprov_nm.value		' Name Cindition
	        arrParam(4) = " MAJOR_CD = " & FilterVar("H0040", "''", "S") & " AND MINOR_CD NOT IN (" & FilterVar("B", "''", "S") & " ," & FilterVar("C", "''", "S") & " )"			' Where Condition							' Where Condition
	        arrParam(5) = "지급구분"					' TextBox 명칭 
	
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
' Function Name : Date_DefMask()
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Function Date_DefMask(strMaskYM)
Dim i,j
Dim ArrMask,StrComDateType
	
	Date_DefMask = False
	
	strMaskYM = ""
	
	ArrMask = Split( parent.gDateFormat, parent.gComDateType)
	
	If  parent.gComDateType = "/" Then 
		lgStrComDateType = "/" & parent.gComDateType
	Else
		lgStrComDateType =  parent.gComDateType
	End If
		
	If IsArray(ArrMask) Then
		For i=0 To Ubound(ArrMask)		
			If Instr(UCase(ArrMask(i)),"D") = False Then
				If strMaskYM <> "" Then
					strMaskYM = strMaskYM & lgStrComDateType
				End If
				If Instr(UCase(ArrMask(i)),"M") And Len(ArrMask(i)) >= 3 Then
					strMaskYM = strMaskYM & "U"
					For j=0 To Len(ArrMask(i)) - 2
						strMaskYM = strMaskYM & "L"
					Next
				Else
					strMaskYM = strMaskYM & ArrMask(i)
				End If
			End If
		Next		
	Else
		Date_DefMask = False
		Exit Function
	End If	

	strMaskYM = Replace(UCase(strMaskYM),"Y","9")
	strMaskYM = Replace(UCase(strMaskYM),"M","9")

	Date_DefMask = True 
	
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
			Call DisplayMsgBox("800140","X","X","X")				 '지급구분코드에 등록되지 않은 코드입니다.
            frm1.txtprov_nm.value = ""
            frm1.txtprov_cd.focus
            txtprov_cd_Onchange = true
        ELSE    
            frm1.txtprov_nm.value = Trim(Replace(lgF0,Chr(11),""))   '지급구분코드 
        END IF
    END IF 

End Function


''==========================================================================================
''   Event Name : vspdData_Click
''   Event Desc :
''==========================================================================================
'Sub vspdData_Click(ByVal Col, ByVal Row)
'    Call SetPopupMenuItemInf("1101111111")
'      gMouseClickStatus = "SPC" 
'    Set gActiveSpdSheet = frm1.vspdData
'
'	if frm1.vspddata.MaxRows <= 0 then
'		exit sub
'	end if
'	
'	if Row <=0 then
'		ggoSpread.Source = frm1.vspdData
'		if lgSortkey = 1 then
'			ggoSpread.SSSort Col
'			lgSortKey = 2
'		else
'			ggoSpread.SSSort Col, lgSortkey
'			lgSortKey = 1
'		end if
'		Exit sub
'	end if
'	frm1.vspdData.Row = Row
'     
'End Sub
'
''========================================================================================================
''   Event Name : vspdData_ColWidthChange
''   Event Desc : 
''========================================================================================================
'Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
'    ggoSpread.Source = frm1.vspdData
'    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
'End Sub
''========================================================================================================
''   Event Name : vspdData_DblClick
''   Event Desc : 
''========================================================================================================
'Sub vspdData_DblClick(ByVal Col, ByVal Row)
'    Dim iColumnName
'    if Row <= 0 then
'		exit sub
'	end if
'	if Frm1.vspdData.MaxRows = 0 then
'		exit sub
'	end if
'End Sub
'
''========================================================================================================
''   Event Name : vspdData_ScriptDragDropBlock
''   Event Desc : 
''========================================================================================================
'Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
'
'    ggoSpread.Source = frm1.vspdData
'    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
'    Call GetSpreadColumnPos("A")
'End Sub
''-----------------------------------------
'Sub vspdData_MouseDown(Button , Shift , x , y)
'   If Button = 2 And  gMouseClickStatus = "SPC" Then
'        gMouseClickStatus = "SPCR"
'   End If
'End Sub    
'
''========================================================================================================
''   Event Name : vspdData_ButtonClicked
''   Event Desc : This function is data query with spread sheet scrolling
''========================================================================================================
'Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
'
'   
'    
'End Sub
'
''======================================================================================================
''	Name : DBAutoQueryOk()
''	Description : HN101BB2_KO441.asp 이후 Query OK해 줌
''=======================================================================================================
'Sub DBAutoQueryOk()
'    Dim lRow
'	Dim intIndex
'	Dim daytimeVal 
'	Dim strSub_type 
'    
'    With Frm1
'        .vspdData.ReDraw = false
'         ggoSpread.Source = .vspdData
'   
'       For lRow = 1 To .vspdData.MaxRows
'            .vspdData.Row = lRow
'            .vspdData.Col = 0
'            .vspdData.Text =  ggoSpread.InsertFlag
'       Next
'            .vspdData.ReDraw = TRUE
'        
'    End With 
'    ggoSpread.ClearSpreadData "T"
'     Set gActiveElement = document.ActiveElement   
'End Sub

'=======================================
'   Event Name :txtDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================

Sub txtDt_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("D")    
        frm1.txtDt.Action = 7
        frm1.txtDt.focus
    End If
End Sub

Sub txtYYMM_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("D")    
        frm1.txtYYMM.Action = 7
        frm1.txtYYMM.focus
    End If
End Sub

'==========================================================================================
'   Event Name : txtDt_KeyDown()
'   Event Desc : 조회조건부의 txtDt_KeyDown시 EnterKey일 경우는 Query
'==========================================================================================
Sub txtDt_KeyDown(KeyCode, Shift)
	'If KeyCode = 13	Then Call mainQuery()
End Sub

Sub txtYYMM_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call mainQuery()
End Sub

'==========================================================================================
'   Event Name : rbo_type1_OnClick()
'   Event Desc : radio button Click시 Grid Setting
'==========================================================================================
Sub rdoCase1_OnClick()
    lgType = "A"   
    Call SetDefaultVal
	Call ggoOper.SetReqAttr(frm1.txtDT, "N")
	
End Sub

Sub rdoCase2_OnClick()
    lgType = "B"
    Call SetDefaultVal
	Call ggoOper.SetReqAttr(frm1.txtDT, "Q")
	
End Sub

Sub rdoCase3_OnClick()
    lgType = "C"
    Call SetDefaultVal
	Call ggoOper.SetReqAttr(frm1.txtDT, "Q")
End Sub

'===============================================================================================
'   by Shin hyoung jae 
'	Name : GetOpenFilePath()
'	Description : GetTextFilePath	
'================================================================================================= 
Function GetOpenFilePath()
	Dim dlg
    Dim sPath
 
	On Error Resume Next
	Set dlg = CreateObject("uni2kCM.SaveFile")
	
	If Err.Number <> 0 Then
		Msgbox Err.Number & " : " & Err.Description
	End If
	
    sPath = dlg.GetOpenFilePath()

	If Err.Number <> 0 Then
		Msgbox Err.Number & " : " & Err.Description
	End If

	lgFilePath2 = sPath
	frm1.txtFileName2.Value = ExtractFileName(sPath)

    Set dlg = Nothing
	frm1.hFileName.value = sPath
End Function

Function ExtractFileName(byVal strPath)
	strPath = StrReverse(strPath)
	strPath = Left(strPath, InStr(strPath, "\") - 1)
	ExtractFileName = StrReverse(strPath)
End Function

'======================================================================================================
' Function Name : ExeReflect
' Function Desc : 
'=======================================================================================================
Function ExeReflect() 
    Dim strVal
    Dim IntRetCD
    Dim strSelect
    Dim strFrom
    Dim strWhere
	
	'ggoSpread.ClearSpreadData

	If trim(frm1.txtFileName2.value) = "" Then
		call DisplayMsgBox("970029","X" , frm1.txtFileName2.Alt, "X")
		frm1.txtFileName2.focus 	
		Exit Function
	Else
		
		if (ggoSaveFile.fileExists(frm1.hFileName.value) = 0)  = false  then
			IntRetCD = DisplayMsgBox("115191","x","x","x")                           '☜:There is no picture
			Exit Function
		end if
			
	End If
	
	If trim(Frm1.txtDt.text) = "" Then
		call DisplayMsgBox("970029","X" , frm1.txtDt.Alt, "X")
		Frm1.txtDt.focus 	
		Exit Function
	End If	
	
	If trim(Frm1.txtYYMM.text) = "" Then
		call DisplayMsgBox("970029","X" , frm1.txtYYMM.Alt, "X")
		Frm1.txtYYMM.focus 	
		Exit Function
	End If		
				    
    
    ExeReflect = False                                                              '☜: Processing is NG
	
    Call MakeKeyStream("X")
    
  	Frm1.htxtDt.value		= Frm1.txtDt.text   
  	Frm1.htxtYYMM.value		= Frm1.txtYYMM.text     
  	Frm1.htxtProvCD.value	= Frm1.txtprov_cd.value   
  
	'Call RemovedivTextArea 



	strSelect	= " COUNT(*) "
	
    Select Case lgType
		   Case "A"	
				strFrom		= " H_IF_HDF070T "
		   Case "B"	
				strFrom		= " H_IF_HDF040T "
		   Case "C"	
				strFrom		= " H_IF_HDF060T "								
	End Select	


	strWhere = ""
	Select Case lgType
		   Case "A"
				strWhere = strWhere & " PROV_DT = " & FilterVar(Frm1.htxtDt.value,"''", "S") & " "
				strWhere = strWhere & " AND PAY_YYMM = " & FilterVar(Replace(Frm1.htxtYYMM.value,"-",""),"''", "S") & "  "
				strWhere = strWhere & " AND PROV_TYPE LIKE '" & Trim(Frm1.htxtProvCD.value) & "%' "
		   Case "B"
				strWhere = strWhere & " PAY_YYMM = " & FilterVar(Replace(Frm1.htxtYYMM.value,"-",""),"''", "S") & " "				
				strWhere = strWhere & " AND PROV_TYPE LIKE '" & Trim(Frm1.htxtProvCD.value) & "%' "
		   Case "C"
				strWhere = strWhere & " SUB_YYMM = " & FilterVar(Replace(Frm1.htxtYYMM.value,"-",""),"''", "S") & " "					
				strWhere = strWhere & " AND SUB_TYPE LIKE '" & Trim(Frm1.htxtProvCD.value) & "%' "
	End Select
	
	
    If 	CommonQueryRs(strSelect, strFrom, strWhere, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true Then

        IF Trim(Replace(lgF0, Chr(11), "")) <> 0 THEN        
			IntRetCD = DisplayMsgBox("800397",Parent.VB_YES_NO,"X","X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
        END If
    Else
		Exit Function    	
    End if




    If LayerShowHide(1) = false Then
        Exit Function
    End If
    
    strVal =""
 	   
	strVal = BIZ_PGM_ID2 & "?txtMode="			& Parent.UID_M0001						'☜: Query
	strVal = strVal      & "&txtKeyStream="     & lgKeyStream							'☜: Query Key
	strVal = strVal      & "&lgStrPrevKey="		& lgStrPrevKey							'☜: Next key tag
	strVal = strVal      & "&txtMaxRows=0"       '& Frm1.vspdData.MaxRows					'☜: Max fetched data	
	strVal = strVal      & "&htxtFileGubun="    & lgType
	strVal = strVal      & "&htxtDt="       	& Frm1.htxtDt.value	
	strVal = strVal      & "&htxtYYMM="       	& Frm1.htxtYYMM.value
	strVal = strVal      & "&htxtProvCD="       & Frm1.htxtProvCD.value	
	
'msgbox strVal	
    Call RunMyBizASP(MyBizASP, strVal)													'☜:  Run biz logic
'msgbox 'ExeReflect-end'
    ExeReflect = True																	'☜: Processing is NG		

End Function

''========================================================================================
' ' Function Name : RemovedivTextArea
' ' Function Desc : 저장후, 동적으로 생성된 HTML 객체(TEXTAREA)를 Clear시켜 준다.
''========================================================================================
' Function RemovedivTextArea()
' 
' 	Dim ii
' 		
' 	For ii = 1 To divTextArea.children.length
' 	    divTextArea.removeChild(divTextArea.children(0))
' 	Next
' 
' End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>


<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>급상여내역UPLOAD</font></td>
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
					<TD HEIGHT=20>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>파일명</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ID="txtFileName2" NAME="txtFileName2" SIZE=30 MAXLENGTH=100 STYLE="TEXT-ALIGN: left" ALT="화일명" tag="14X" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenFilePath" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:GetOpenFilePath()"></TD>							
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>파일구분</TD>
								<TD CLASS="TD6">
									<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType" ID="rdoCase1" TAG="1X" checked>
									<LABEL FOR="rdoCase1">급/상여내역</LABEL>
									<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType" ID="rdoCase2" TAG="1X">
									<LABEL FOR="rdoCase2">수당내역</LABEL>
									<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType" ID="rdoCase3" TAG="1X">
									<LABEL FOR="rdoCase3">공제내역</LABEL>
								</TD>
							</TR>
							<TR>
							   	<TD CLASS=TD5 NOWRAP>급/상여일자</TD>
								<TD CLASS=TD6 NOWRAP><OBJECT classid=<%=gCLSIDFPDT%> id=txtDt NAME="txtDt" CLASS=FPDTYYYYMMDD title=FPDATETIME  ALT="급/상여일자" tag="12X1" VIEWASTEXT> </OBJECT></TD>
							</TR>
							<TR>
							   	<TD CLASS=TD5 NOWRAP>급/상여년월</TD>
								<TD CLASS=TD6 NOWRAP><OBJECT classid=<%=gCLSIDFPDT%> id=txtYYMM NAME="txtYYMM" CLASS=FPDTYYYYMM title=FPDATETIME  ALT="급/상여년월" tag="12X1" VIEWASTEXT> </OBJECT></TD>
							</TR>
								<TD CLASS="TD5" NOWRAP>지급구분</TD>
	                        	<TD CLASS="TD6" NOWRAP><INPUT Type="TEXT" NAME="txtProv_cd" MAXLENGTH="1" SIZE="10" ALT ="지급구분" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCondAreaPopup(2)">
	                        	<INPUT Type="TEXT" Name="txtProv_nm" MAXLENGTH="20" SIZE="20" ALT ="지급구분명" tag="14XXXU">
	                        </TD>
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
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnExe"  CLASS="CLSSBTN" onclick="ExeReflect()" Flag=1>실행</BUTTON>&nbsp;
					    <BUTTON NAME="btnExe2" CLASS="CLSSBTN" onclick="FncDelete()"  Flag=1>취소</BUTTON></TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=Bizsize%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=Bizsize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD>
	</TR>

</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" style='display:none'></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioKind" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioType" tag="24">
<INPUT TYPE=HIDDEN NAME="txtCheck" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtDt" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtYYMM" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtProvCD" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtFileGubun" tag="24">
<INPUT TYPE=HIDDEN NAME="hFileName" tag="14" TABINDEX="-1">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
