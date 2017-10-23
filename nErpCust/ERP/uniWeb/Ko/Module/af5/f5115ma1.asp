
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : ACCOUNT
*  2. Function Name        : 
*  3. Program ID           : f5115ma1
*  4. Program Name         : 
*  5. Program Desc         : 자금관리/어음관리/어음일괄등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2002/03/30
*  8. Modified date(Last)  : 
*  9. Modifier (First)     : Oh Soo Min
* 10. Modifier (Last)      : 
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->

<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
	

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID = "f5115mb1.asp"                                      'Biz Logic ASP 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const CookieSplit = 1233

'어음번호, 어음구분, 어음상태, 부서, 거래처, 은행, 발행일, 만기일, 어음금액, 보관장소, 자타수구분, 발행인, 비고 
Dim C_NOTE_TYPE
Dim C_NOTE_NO  
Dim C_NOTE_POP
Dim C_NOTE_STS 
Dim C_DEPT_CD  
Dim C_DEPT_POP 
Dim C_DEPT_NM  
Dim C_BP_CD    
Dim C_BP_POP   
Dim C_BP_NM    
Dim C_BANK_CD  
Dim C_BANK_POP 
Dim C_BANK_NM  
Dim C_ISSUE_DT 
Dim C_DUE_DT   
Dim C_NOTE_AMT
Dim C_CASH_RATE 
Dim C_PLACE    
Dim C_RCPTFG   
Dim C_PUBLISHER
Dim C_DESC     
Dim C_NOTE_TYPECD
Dim C_NOTE_STSCD 
Dim C_PLACE_CD   
Dim C_RCPTFG_CD  
Dim C_COST_CD    
Dim C_BIZ_AREA_CD
Dim C_INTERNAL_CD
Dim C_ORG_CHANGE_ID
Dim C_ISSUE_DT_HH



'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd
Dim IsOpenPop    

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgStrPrevKeyIndex = ""                                      '⊙: initializes Previous Key Index
    lgSortKey         = 1                                       '⊙: initializes sort direction
		
End Sub


Sub initSpreadPosVariables()

	 C_NOTE_TYPE   = 1
	 C_NOTE_NO     = 2
	 C_NOTE_POP    = 3
	 C_NOTE_STS    = 4
	 C_DEPT_CD     = 5
	 C_DEPT_POP    = 6
	 C_DEPT_NM     = 7
	 C_BP_CD       = 8
	 C_BP_POP      = 9
	 C_BP_NM       = 10
	 C_BANK_CD     = 11
	 C_BANK_POP    = 12
	 C_BANK_NM     = 13
	 C_ISSUE_DT    = 14
	 C_DUE_DT      = 15
	 C_NOTE_AMT    = 16
	 C_CASH_RATE   = 17	 
	 C_PLACE       = 18
	 C_RCPTFG      = 19
	 C_PUBLISHER   = 20
	 C_DESC        = 21
	 C_NOTE_TYPECD = 22
	 C_NOTE_STSCD  = 23
	 C_PLACE_CD    = 24
	 C_RCPTFG_CD   = 25
	 C_COST_CD     = 26
	 C_BIZ_AREA_CD = 27
	 C_INTERNAL_CD = 28
	 C_ORG_CHANGE_ID = 29
	 C_ISSUE_DT_HH = 30

End Sub
'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Dim strSvrDate
	DIm strYear, strMonth, strDay
	Dim frDt, toDt
	
	strSvrDate = "<%=GetSvrDate%>"
	Call ExtractDateFrom(strSvrDate, parent.gServerDateFormat, parent.gServerDateType, strYear,strMonth,strDay)
		
	frDt = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")
	toDt = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)
	
	frm1.txtIssueDt1.Text = frDt
	frm1.txtIssueDt2.Text = toDt	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   
    lgKeyStream  = frm1.txtNOTENO.value & parent.gColSep                                               'You Must append one character(parent.gColSep)   
    lgKeyStream  = lgKeyStream & frm1.txtBPCD.value & parent.gColSep
    lgKeyStream  = lgKeyStream & frm1.txtIssueDt1.Text & parent.gColSep
    lgKeyStream  = lgKeyStream & frm1.txtIssueDt2.Text & parent.gColSep
    lgKeyStream  = lgKeyStream & frm1.txtDueDt1.Text & parent.gColSep
    lgKeyStream  = lgKeyStream & frm1.cbonotefg.Value & parent.gColSep
    
        
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        


'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
	Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    
    Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("F1007", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
    Call SetCombo2(frm1.cboNoteFg ,lgF0  ,lgF1  ,Chr(11))
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'어음구분 
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("F1007", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	iCodeArr = lgF0
    iNameArr = lgF1
    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_NOTE_TYPECD
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_NOTE_TYPE
    
	'어음상태 
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("F1008", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	iCodeArr = lgF0
    iNameArr = lgF1
    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_NOTE_STSCD
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_NOTE_STS
    
	'보관장소 
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("F1005", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	iCodeArr = lgF0
    iNameArr = lgF1
    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_PLACE_CD
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_PLACE
    
	'자타수구분 
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("F1009", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	iCodeArr = lgF0
    iNameArr = lgF1
    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_RCPTFG_CD
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_RCPTFG
    
    '------ Developer Coding part (End )   -------------------------------------------------------------- 
   
End Sub

Sub InitCombo()
	Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'어음구분 
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("F1007", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	iCodeArr = lgF0
    iNameArr = lgF1
    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_NOTE_TYPECD
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_NOTE_TYPE
    
	'어음상태 
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("F1008", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	iCodeArr = lgF0
    iNameArr = lgF1
    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_NOTE_STSCD
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_NOTE_STS
    
	'보관장소 
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("F1005", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	iCodeArr = lgF0
    iNameArr = lgF1
    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_PLACE_CD
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_PLACE
    
	'자타수구분 
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("F1009", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	iCodeArr = lgF0
    iNameArr = lgF1
    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_RCPTFG_CD
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_RCPTFG
    
    '------ Developer Coding part (End )   -------------------------------------------------------------- 
   
End Sub

'========================================================================================================
' Name : vspdData_ComboSelChange
' Desc : ComboBox에 값 변경시 처리 
'========================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
	With frm1.vspdData
		
		.Row = Row
    
        Select Case Col
            Case C_NOTE_TYPE
                .Col = Col
                intIndex = .Value        '  COMBO의 VALUE값				
				.Col = C_NOTE_TYPECD      '  CODE값란으로 이동 
				.Value = intIndex        '  CODE란의 값은 COMBO의 VALUE값이된다.				

'		    Case C_NOTE_STS
'                .Col = Col
'                intIndex = .Value        
'				.Col = C_NOTE_STSCD     
'				.Value = intIndex        
				
		    Case C_PLACE
                .Col = Col
                intIndex = .Value   
				.Col = C_PLACE_CD
				.Value = intIndex 
			
			Case C_RCPTFG
                .Col = Col
                intIndex = .Value   
				.Col = C_RCPTFG_CD
				.Value = intIndex 
			  
		End Select
	End With

   	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

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
			.Col = C_NOTE_TYPECD
			intIndex = .value
			.col = C_NOTE_TYPE
			.value = intindex
			
			.Row = intRow
			.Col = C_NOTE_STSCD
			intIndex = .value
			.col = C_NOTE_STS
			.value = intindex
			
			.Row = intRow
			.Col = C_PLACE_CD
			intIndex = .value
			.col = C_PLACE
			.value = intindex	
				
			.Row = intRow
			.Col = C_RCPTFG_CD
			intIndex = .value
			.col = C_RCPTFG
			.value = intindex	
			
		Next	
	End With
	
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

    Call initSpreadPosVariables()

	With frm1.vspdData

       .MaxCols = C_ISSUE_DT_HH + 1
	   .Col = .MaxCols
       .ColHidden = True
       .MaxRows = 0
        ggoSpread.Source = frm1.vspdData

	   .ReDraw = false

        ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread 
        
        Call GetSpreadColumnPos("A")       
                                
        ggoSpread.SSSetCombo   C_NOTE_TYPE,   "어음구분",          10, 0        
        ggoSpread.SSSetEdit    C_NOTE_NO,     "어음번호",          15, 0, , 30, 2
        ggoSpread.SSSetButton  C_NOTE_POP
        ggoSpread.SSSetCombo   C_NOTE_TYPECD, "어음구분CD",         7, 0
        ggoSpread.SSSetCombo   C_NOTE_STS,    "어음상태",          10, 0        
        ggoSpread.SSSetCombo   C_NOTE_STSCD,  "어음상태CD",         7, 0
        ggoSpread.SSSetEdit    C_DEPT_CD,     "부서",               5, 0, , 30, 2
        ggoSpread.SSSetButton  C_DEPT_POP
        ggoSpread.SSSetEdit    C_DEPT_NM,     "부서명",            14, 0, , 30, 2         
        ggoSpread.SSSetEdit    C_BP_CD,       "거래처",            10, 0, , 30, 2
        ggoSpread.SSSetButton  C_BP_POP
        ggoSpread.SSSetEdit    C_BP_NM,       "거래처명",          12, 0, , 30, 2                 
        ggoSpread.SSSetEdit    C_BANK_CD,     "은행",              10, 0, , 30, 2
        ggoSpread.SSSetButton  C_BANK_POP
        ggoSpread.SSSetEdit    C_BANK_NM,     "은행명",            12, 0, , 30, 2                 
        ggoSpread.SSSetDate    C_ISSUE_DT,    "발행일",            12   ,2   ,parent.gDateFormat   ,-1
        ggoSpread.SSSetDate    C_DUE_DT,      "만기일",            12   ,2   ,parent.gDateFormat   ,-1  
        ggoSpread.SSSetFloat   C_NOTE_AMT,    "어음금액",          15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec             
 		ggoSpread.SSSetFloat   C_CASH_RATE,   "현금율",			   15, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetCombo   C_PLACE,       "보관장소",           9, 0
        ggoSpread.SSSetCombo   C_PLACE_CD,    "보관장소CD",         7, 0        
        ggoSpread.SSSetCombo   C_RCPTFG,      "자타수구분",         9, 0
        ggoSpread.SSSetCombo   C_RCPTFG_CD,   "자타수구분CD",       7, 0                
        ggoSpread.SSSetEdit    C_PUBLISHER,   "발행인",            15, 0, , 30, 2        
        ggoSpread.SSSetEdit    C_DESC,        "비고",              20, 0, , 30, 2        
        ggoSpread.SSSetEdit    C_COST_CD,     "코스트센터",        15, 0, , 30, 2
        ggoSpread.SSSetEdit    C_BIZ_AREA_CD, "사업장코드",        15, 0, , 30, 2                
        ggoSpread.SSSetEdit    C_INTERNAL_CD, "내부부서코드",      20, 0, , 30, 2
        ggoSpread.SSSetEdit    C_ORG_CHANGE_ID, "조직변경ID",      20, 0, , 30, 2
        ggoSpread.SSSetDate    C_ISSUE_DT_HH, "발행일",            12   ,2   ,parent.gDateFormat   ,-1
                
	   .ReDraw = true
	
	   call ggoSpread.MakePairsColumn(C_NOTE_NO,C_NOTE_POP)
	   call ggoSpread.MakePairsColumn(C_DEPT_CD,C_DEPT_POP)
       call ggoSpread.MakePairsColumn(C_BP_CD,C_BP_POP)
       call ggoSpread.MakePairsColumn(C_BANK_CD,C_BANK_POP)

       Call ggoSpread.SSSetColHidden(C_NOTE_TYPECD,C_NOTE_TYPECD,True)
       Call ggoSpread.SSSetColHidden(C_NOTE_STSCD,C_NOTE_STSCD,True)
       Call ggoSpread.SSSetColHidden(C_PLACE_CD,C_PLACE_CD,True)
       Call ggoSpread.SSSetColHidden(C_RCPTFG_CD,C_RCPTFG_CD,True)
       Call ggoSpread.SSSetColHidden(C_COST_CD,C_COST_CD,True)
       Call ggoSpread.SSSetColHidden(C_BIZ_AREA_CD,C_BIZ_AREA_CD,True)
       Call ggoSpread.SSSetColHidden(C_INTERNAL_CD,C_INTERNAL_CD,True)
       Call ggoSpread.SSSetColHidden(C_ORG_CHANGE_ID,C_ORG_CHANGE_ID,True)
       Call ggoSpread.SSSetColHidden(C_ISSUE_DT_HH,C_ISSUE_DT_HH,True)
	
       Call SetSpreadLock()   
    
    End With
    
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
	Dim RowCnt
	ggoSpread.Source = frm1.vspdData
	
    With frm1  
      .vspdData.ReDraw = False      
      ggoSpread.SpreadLock		1 ,     -1
      
      For RowCnt = 1 To .vspdData.MaxRows								
			
			.vspdData.Col = C_NOTE_STSCD
			.vspdData.Row = RowCnt			
			If Trim(.vspdData.text) <> "BG" Then
					
				ggoSpread.SpreadLock	C_NOTE_NO		,RowCnt, RowCnt
				ggoSpread.SpreadLock	C_NOTE_TYPE		,RowCnt, RowCnt
				ggoSpread.SpreadLock	C_NOTE_STS		 ,RowCnt, RowCnt
				
				ggoSpread.SpreadLock	C_DEPT_CD		,RowCnt, RowCnt
				ggoSpread.SpreadLock	C_DEPT_POP ,     RowCnt,  RowCnt				
				ggoSpread.SpreadLock    C_DEPT_NM,      RowCnt,   RowCnt
				
				ggoSpread.SpreadLock	C_BP_CD			 ,RowCnt, RowCnt
				ggoSpread.SpreadLock    C_BP_POP,        RowCnt,   RowCnt
				ggoSpread.SpreadLock    C_BP_NM,        RowCnt,   RowCnt
				
				ggoSpread.SpreadLock	C_BANK_CD		 ,RowCnt, RowCnt
				ggoSpread.SpreadLock    C_BANK_POP,        RowCnt,   RowCnt
				ggoSpread.SpreadLock    C_BANK_NM,      RowCnt,   RowCnt	   			  
				
				ggoSpread.SpreadLock	C_ISSUE_DT		 ,RowCnt, RowCnt				
				ggoSpread.SpreadLock	C_DUE_DT		 ,RowCnt, RowCnt
				ggoSpread.SpreadLock	C_NOTE_AMT		 ,RowCnt, RowCnt							
				ggoSpread.SpreadLock	C_PLACE			,RowCnt, RowCnt
				ggoSpread.SpreadLock	C_RCPTFG		 ,RowCnt, RowCnt				
				ggoSpread.SpreadLock	C_PUBLISHER		 ,RowCnt, RowCnt
				ggoSpread.SpreadLock	C_DESC			,RowCnt, RowCnt	
				ggoSpread.SpreadLock	C_CASH_RATE,	RowCnt,		RowCnt				
				
			Else	
											
				ggoSpread.SpreadLock		C_NOTE_NO ,     RowCnt,		RowCnt				
				ggoSpread.SpreadLock		C_NOTE_TYPE,	RowCnt,		RowCnt				
				ggoSpread.SpreadLock		C_NOTE_STS,     RowCnt,		RowCnt

				ggoSpread.SpreadUnLock		C_DEPT_CD,		RowCnt,		RowCnt
				ggoSpread.SSSetRequired     C_DEPT_CD ,     RowCnt,     RowCnt
				ggoSpread.SpreadUnLock		C_DEPT_POP,		RowCnt,		RowCnt
				ggoSpread.SSSetRequired		C_DEPT_POP,		RowCnt,		RowCnt
				.vspdData.Col = C_DEPT_POP
				.vspdData.Row = RowCnt	
				.vspdData.text = "1"
				ggoSpread.SpreadLock        C_DEPT_NM,      RowCnt,		RowCnt

				ggoSpread.SpreadUnLock		C_BP_CD,		RowCnt,		RowCnt
				ggoSpread.SSSetRequired     C_BP_CD ,       RowCnt,     RowCnt
				ggoSpread.SpreadUnLock      C_BP_POP,       RowCnt,		RowCnt
				ggoSpread.SSSetRequired     C_BP_POP,       RowCnt,		RowCnt
				.vspdData.Col = C_BP_POP
				.vspdData.Row = RowCnt	
				.vspdData.text = "1"
				ggoSpread.SpreadLock        C_BP_NM,        RowCnt,		RowCnt

				ggoSpread.SpreadUnLock		C_BANK_CD,		RowCnt,		RowCnt
				ggoSpread.SSSetRequired     C_BANK_CD ,     RowCnt,     RowCnt
				ggoSpread.SpreadUnLock      C_BANK_POP,     RowCnt,		RowCnt
				ggoSpread.SSSetRequired     C_BANK_POP,     RowCnt,		RowCnt
				.vspdData.Col = C_BANK_POP
				.vspdData.Row = RowCnt	
				.vspdData.text = "1"
				ggoSpread.SpreadLock        C_BANK_NM,      RowCnt,		RowCnt	

				ggoSpread.SpreadUnLock		C_ISSUE_DT,		RowCnt,		RowCnt
				ggoSpread.SSSetRequired     C_ISSUE_DT ,    RowCnt,     RowCnt
				
				ggoSpread.SpreadUnLock		C_DUE_DT,		RowCnt,		RowCnt
				ggoSpread.SSSetRequired     C_DUE_DT ,      RowCnt,     RowCnt
				
				ggoSpread.SpreadUnLock		C_NOTE_AMT,		RowCnt,		RowCnt
				ggoSpread.SSSetRequired     C_NOTE_AMT ,    RowCnt,     RowCnt
				
				frm1.vspdData.Row = RowCnt
				frm1.vspdData.Col = C_NOTE_TYPECD				
				Select Case UCase(Trim(frm1.vspdData.text))
					Case "D1"	'받을어음 
			            frm1.vspdData.Col = C_CASH_RATE						
						ggoSpread.SpreadUnLock		C_CASH_RATE,		RowCnt,		RowCnt
						ggoSpread.SSSetRequired     C_CASH_RATE ,		RowCnt,     RowCnt
					Case "D3"	'지급어음 
			            frm1.vspdData.Col = C_CASH_RATE						
						ggoSpread.SpreadLock		C_CASH_RATE,      RowCnt,	 RowCnt
					Case Else
			            frm1.vspdData.Col = C_CASH_RATE						
						ggoSpread.SSSetRequired     C_CASH_RATE ,   RowCnt,     RowCnt
				End Select
				
				ggoSpread.SpreadUnLock		C_PLACE,		RowCnt,		C_PLACE, RowCnt
				ggoSpread.SpreadUnLock		C_RCPTFG,		RowCnt,		C_RCPTFG, RowCnt
				ggoSpread.SpreadUnLock		C_PUBLISHER,	RowCnt,		C_PUBLISHER, RowCnt
				ggoSpread.SpreadUnLock		C_DESC,			RowCnt,		C_DESC, RowCnt
				ggoSpread.SpreadUnLock      C_NOTE_TYPECD,	RowCnt,		C_NOTE_TYPECD, RowCnt
			
			End If 


		Next     
    
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
		
	    ggoSpread.SSSetRequired		C_NOTE_NO ,     pvStartRow,	pvEndRow
		ggoSpread.SSSetRequired		C_NOTE_TYPE,	pvStartRow,	pvEndRow
		ggoSpread.SSSetProtected	C_NOTE_STS,     pvStartRow,	pvEndRow
				
		'ggoSpread.SpreadUnLock		C_DEPT_CD ,     pvStartRow,	pvEndRow
		ggoSpread.SSSetRequired		C_DEPT_CD ,     pvStartRow,	pvEndRow
		'ggoSpread.SpreadUnLock		C_DEPT_POP,		pvStartRow,	pvEndRow		
		ggoSpread.SSSetProtected	C_DEPT_NM,      pvStartRow,	pvEndRow
		
		'ggoSpread.SpreadUnLock		C_BP_CD ,       pvStartRow,	pvEndRow
		ggoSpread.SSSetRequired     C_BP_CD ,       pvStartRow,	pvEndRow
		'ggoSpread.SpreadUnLock     C_BP_POP,       pvStartRow,	pvEndRow		
		ggoSpread.SSSetProtected	C_BP_NM,        pvStartRow,	pvEndRow
		
		'ggoSpread.SpreadUnLock		C_BANK_CD ,     pvStartRow,	pvEndRow
		ggoSpread.SSSetRequired     C_BANK_CD ,     pvStartRow,	pvEndRow
		'ggoSpread.SpreadUnLock     C_BANK_POP,     pvStartRow,	pvEndRow		
		ggoSpread.SSSetProtected	C_BANK_NM,      pvStartRow,	pvEndRow
				
		'ggoSpread.SpreadUnLock		C_ISSUE_DT ,    pvStartRow,	pvEndRow
		ggoSpread.SSSetRequired     C_ISSUE_DT ,    pvStartRow,	pvEndRow
		
		'ggoSpread.SpreadUnLock		C_DUE_DT ,      pvStartRow,	pvEndRow
		ggoSpread.SSSetRequired     C_DUE_DT ,      pvStartRow,	pvEndRow
		
		'ggoSpread.SpreadUnLock		C_NOTE_AMT ,    pvStartRow,	pvEndRow
		ggoSpread.SSSetRequired     C_NOTE_AMT ,    pvStartRow,	pvEndRow
		
		'ggoSpread.SpreadUnLock		C_CASH_RATE ,   pvStartRow,	pvEndRow
		ggoSpread.SSSetRequired     C_CASH_RATE ,   pvStartRow,	pvEndRow
		
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
    iPosArr = Split(iPosArr,parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub
'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

	Call LoadInfTB19029                                                             '⊙: Load table , B_numeric_format
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
        
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    Call InitComboBox       
    
    Call SetDefaultVal    
    Call SetToolbar("1100110100001111")										        '버튼 툴바 제어    
	Call frm1.txtIssueDt1.focus
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	
	
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

	 		C_NOTE_TYPE   = iCurColumnPos(1)
			C_NOTE_NO     = iCurColumnPos(2)
			C_NOTE_POP    = iCurColumnPos(3)
	 		C_NOTE_STS    = iCurColumnPos(4)
 	 		C_DEPT_CD     = iCurColumnPos(5)
 	 		C_DEPT_POP    = iCurColumnPos(6)
 	 		C_DEPT_NM     = iCurColumnPos(7)
 	 		C_BP_CD       = iCurColumnPos(8)
 	 		C_BP_POP      = iCurColumnPos(9)
 	 		C_BP_NM       = iCurColumnPos(10)
 	 		C_BANK_CD     = iCurColumnPos(11)
 	 		C_BANK_POP    = iCurColumnPos(12)
 	 		C_BANK_NM     = iCurColumnPos(13)
 	 		C_ISSUE_DT    = iCurColumnPos(14)
 	 		C_DUE_DT      = iCurColumnPos(15)
 	 		C_NOTE_AMT    = iCurColumnPos(16)
 	 		C_CASH_RATE   = iCurColumnPos(17)
 	 		C_PLACE       = iCurColumnPos(18)
 	 		C_RCPTFG      = iCurColumnPos(19)
 	 		C_PUBLISHER   = iCurColumnPos(20)
 	 		C_DESC        = iCurColumnPos(21)
 	 		C_NOTE_TYPECD = iCurColumnPos(22)
 	 		C_NOTE_STSCD  = iCurColumnPos(23)
 	 		C_PLACE_CD    = iCurColumnPos(24)
 	 		C_RCPTFG_CD   = iCurColumnPos(25)
 	 		C_COST_CD     = iCurColumnPos(26)
 	 		C_BIZ_AREA_CD = iCurColumnPos(27)
 	 		C_INTERNAL_CD = iCurColumnPos(28)
 	 		C_ORG_CHANGE_ID = iCurColumnPos(29)
 	 		C_ISSUE_DT_HH = iCurColumnPos(30)
 	 		
    End Select    
End Sub
'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub


Sub vspdData_Click(ByVal Col, ByVal Row)
	
    Call SetPopupMenuItemInf("1101111111")
    gMouseClickStatus = "SPC"	'Split 상태코드 
    
    Set gActiveSpdSheet = frm1.vspdData
    
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
End Sub

Sub vspdData_DblClick(ByVal Col, ByVal Row)				
     Dim iColumnName
    
    If Row <= 0 Then
      Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
       Exit Sub
    End If     
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
   
	
	
End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
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
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
        
	If (frm1.txtIssueDt1.Text <> "") And (frm1.txtIssueDt2.Text <> "") Then
		If CompareDateByFormat(frm1.txtIssueDt1.Text, frm1.txtIssueDt2.Text, frm1.txtIssueDt1.Alt, frm1.txtIssueDt2.Alt, _
					"970025", frm1.txtIssueDt1.UserDefinedFormat, parent.gComDateType, true) = False Then
			frm1.txtIssueDt1.focus											
			Exit Function
		End if	
	End If
	    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    Call InitVariables                                                           '⊙: Initializes local global variables    
    Call MakeKeyStream("")    
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
    ggospread.ClearSpreadData										 '☜: Clear Contents  Field
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
          
    If DbQuery = False Then
        Exit Function
    End If
              
    FncQuery = True                                                              '☜: Processing is OK
	Set gActiveElement = document.activeElement    
End Function

'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncNew = True																 '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'In Multi, You need not to implement this area
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement       
    FncDelete = True                                                             '☜: Processing is OK
End Function
'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    Dim lRow
    Dim txtIssueDt, txtDueDt, txtCUDFlg, txtIssueDtHH
    
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
    
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	With Frm1
    
		For lRow = 1 To .vspdData.MaxRows
    
			.vspdData.Row = lRow
			.vspdData.Col = C_ISSUE_DT		: txtIssueDt = .vspdData.Text
			.vspdData.Col = C_DUE_DT		: txtDueDt	 = .vspdData.Text
			.vspdData.Col = 0				: txtCUDFlg	 = .vspdData.Text
			.vspdData.Col = C_ISSUE_DT_HH	: txtIssueDtHH = .vspdData.Text
			
			If Trim(txtCUDFlg) <> "" Then
				
				'2003/04/16 지급어음일 경우, 발행은행 체크				
				.vspdData.Col = C_NOTE_TYPECD			
				
				If .vspdData.text = "D3" Then
					.vspdData.Col = C_NOTE_NO				
					Call CommonQueryRs("BANK_CD","F_NOTE_NO","NOTE_NO = " & FilterVar(.vspdData.Text, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					.vspdData.Col = C_BANK_CD
					If UCase(Trim(.vspdData.Text)) <> UCase(Trim(Replace(lgF0, chr(11), ""))) Then
						Call DisplayMsgBox("141423","x","x","x")
						Exit Function
					End If
				End If 

				If (txtIssueDt <> "") And (txtIssueDtHH <> "") Then
					If CompareDateByFormat(txtIssueDtHH, txtIssueDt, "x", "x", _
								"141228", frm1.txtIssueDt1.UserDefinedFormat, parent.gComDateType, true) = False Then
						frm1.vspdData.focus
						Exit Function
					End if	
				End If
				If (txtIssueDt <> "") And (txtDueDt <> "") Then
					If CompareDateByFormat(txtIssueDt, txtDueDt, "발행일", "만기일", _
								"970025", frm1.txtIssueDt1.UserDefinedFormat, parent.gComDateType, true) = False Then
						frm1.vspdData.focus											
						Exit Function
					End if	
				End If 
			End If            


		Next
	End With

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	
    Call MakeKeyStream("X")
       
    
    If DbSave = False Then
        Exit Function
    End If    
    
    FncSave = True                                                              '☜: Processing is OK
    
    Set gActiveElement = document.activeElement
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
			Call SetSpreadColor(Frm1.VspdData.ActiveRow,  frm1.vspdData.ActiveRow)

            .ReDraw = True
		    .Focus
		 End If
	End With
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 	
	
	With Frm1.VspdData
           .Col  = C_NOTE_NO
           .Row  = .ActiveRow
           .Text = ""  
           	 
    End With	
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
		
    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  
    Call initData()
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(Byval pvRowcnt) 
    Dim imRow
    Dim ii
    Dim iCurRowPos

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear   

	FncInsertRow = False                                                         '☜: Processing is NG

    If IsNumeric(Trim(pvRowcnt)) Then 
	   imRow  = Cint(pvRowcnt)
	else
	   imRow = AskSpdSheetAddRowCount()
	   If imRow = "" Then
	      Exit Function
	   End If
	End If                              
		
	With frm1
   		iCurRowPos = frm1.vspdData.ActiveRow
		.vspdData.ReDraw = False
		ggoSpread.Source = frm1.vspdData
		ggoSpread.InsertRow,imRow	
		

		For ii = iCurRowPos + 1 To  iCurRowPos  + imRow 
			.vspdData.Row = ii
			.vspdData.col = C_ISSUE_DT
			.vspdData.Text= UniConvDateAToB("<%=GetSvrDate%>",parent.gServerDateFormat, parent.gDateFormat)	'변경일 default : today		
		Next
		
		Call SetSpreadColor(iCurRowPos + 1, iCurRowPos + imRow)        
		
		.vspdData.ReDraw = True		
     	.vspdData.focus
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
    	lDelRows = ggoSpread.DeleteRow
    End With
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function FncPrint()
    Call parent.FncPrint()
	Set gActiveElement = document.activeElement
End Function

'========================================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================================
Function FncPrev() 
    On Error Resume Next                                                  '☜: Protect system from crashing
End Function

'========================================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================================
Function FncNext() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'========================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)                                         '☜: 화면 유형 
	Set gActiveElement = document.activeElement
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                    '☜:화면 유형, Tab 유무 
	Set gActiveElement = document.activeElement
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
    Call InitCombo()
    Call ggoSpread.ReOrderingSpreadData()
    Call SetSpreadLock()
    Call InitData()
End Sub

'========================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()
    Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

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
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex                 '☜: Next key tag
    End With
	
	
	Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic
    
    DbQuery = True
    
End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    
    Dim lRow        
    Dim lGrpCnt         
	Dim strVal, strDel
		
	On Error Resume Next
    DbSave = False                                                             
    Err.Clear 
    
    Call DisableToolBar(parent.TBC_SAVE)                                                '☜: Disable Save Button Of ToolBar
    Call LayerShowHide(1)

	Frm1.txtMode.value        = parent.UID_M0002                                        '☜: Delete
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	
    ggoSpread.Source = frm1.vspdData

    strVal = ""
    strDel = ""
    lGrpCnt = 1
        
    With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
 
               Case ggoSpread.InsertFlag                                                                '☜: Update추가 
                                                       strVal = strVal & "C" & parent.gColSep 
                                                       strVal = strVal & lRow & parent.gColSep
                    .vspdData.Col = C_NOTE_NO	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep                    
                    .vspdData.Col = C_NOTE_TYPECD	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_NOTE_STSCD	 : strVal = strVal & Trim("BG") & parent.gColSep                    
                    .vspdData.Col = C_DEPT_CD	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep                    
                    .vspdData.Col = C_BP_CD	         : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep                             
                    .vspdData.Col = C_BANK_CD	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep          
                    .vspdData.Col = C_ISSUE_DT	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep                    
                    .vspdData.Col = C_DUE_DT	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep          
                    .vspdData.Col = C_NOTE_AMT	     : strVal = strVal & Trim(.vspdData.Value) & parent.gColSep                    
                    .vspdData.Col = C_PLACE_CD	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep                              
                    .vspdData.Col = C_RCPTFG_CD	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep          
                    .vspdData.Col = C_PUBLISHER	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep                    
                    .vspdData.Col = C_DESC	         : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep         
                    .vspdData.Col = C_COST_CD	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_BIZ_AREA_CD    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_INTERNAL_CD    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ORG_CHANGE_ID  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CASH_RATE      : strVal = strVal & UNICDbl(.vspdData.Text) & parent.gRowSep

                    lGrpCnt = lGrpCnt + 1                                     
                    
               Case ggoSpread.UpdateFlag                                                                 '☜: Update
                                                       strVal = strVal & "U" & parent.gColSep
                                                       strVal = strVal & lRow & parent.gColSep
                    .vspdData.Col = C_NOTE_NO	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep                    
                    .vspdData.Col = C_NOTE_TYPECD	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep          
                    .vspdData.Col = C_NOTE_STSCD	 : strVal = strVal & Trim("BG") & parent.gColSep                    
                    .vspdData.Col = C_DEPT_CD	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep                    
                    .vspdData.Col = C_BP_CD	         : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep                             
                    .vspdData.Col = C_BANK_CD	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep          
                    .vspdData.Col = C_ISSUE_DT	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep                    
                    .vspdData.Col = C_DUE_DT	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep          
                    .vspdData.Col = C_NOTE_AMT	     : strVal = strVal & Trim(.vspdData.Value) & parent.gColSep                    
                    .vspdData.Col = C_PLACE_CD	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep                              
                    .vspdData.Col = C_RCPTFG_CD	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep          
                    .vspdData.Col = C_PUBLISHER	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep                    
                    .vspdData.Col = C_DESC	         : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep         
                    .vspdData.Col = C_COST_CD	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep 
                    .vspdData.Col = C_BIZ_AREA_CD    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep                   
                    .vspdData.Col = C_INTERNAL_CD    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ORG_CHANGE_ID  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep	
                    .vspdData.Col = C_CASH_RATE      : strVal = strVal & UNICDbl(.vspdData.Text) & parent.gRowSep

                    lGrpCnt = lGrpCnt + 1
                    
               Case ggoSpread.DeleteFlag                                                                 '☜: Delete             
                                                       strDel = strDel & "D" & parent.gColSep
                                                       strDel = strDel & lRow & parent.gColSep
                   .vspdData.Col = C_NOTE_NO	     : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep                   
                   lGrpCnt = lGrpCnt + 1
             
           End Select
           
       Next
	
       .txtMode.value        = parent.UID_M0002
       .txtUpdtUserId.value  = parent.gUsrID
       .txtInsrtUserId.value = parent.gUsrID
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
    Err.Clear                                                                    '☜: Clear err status
    DbDelete = False                                                             '☜: Processing is NG
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'In Multi, You need not to implement this area

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    DbDelete = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()												     
    Call InitData()
    '-----------------------
    'Reset variables area
    '-----------------------    
    lgIntFlgMode = parent.OPMD_UMODE    
    Call ggoOper.LockField(Document, "Q")										'⊙: Lock field            
         
    Call SetSpreadLock()		
	
	Call SetToolbar("1100111100111111")		
	Set gActiveElement = document.activeElement 						
	
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
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

Function OpenNoteInfo()

	Dim arrRet
	Dim arrParam(5)	

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
   
	arrRet = window.showModalDialog("../af5/f5101ra1.asp", Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False	

	If arrRet(0) = "" Then	    
		Exit Function
	Else
		frm1.txtNoteNo.value  = arrRet(0)
	End If	

End Function
'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' 채권과 연계(거래처 유무)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "T"							'B :매출 S: 매입 T: 전체 
	arrParam(5) = ""									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EScCode(iwhere)
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
		if iWhere <> 0 then
       	  ggoSpread.Source = frm1.vspdData
           ggoSpread.UpdateRow frm1.vspdData.ActiveRow
        End if
	End If	
End Function
'========================================================================================================
'	Name : OpenPopup()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function OpenPopup(ByVal strCode, Byval iWhere)

	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
    
	Select Case iWhere
	    Case 0, 3	    
	        arrParam(0) = "거래처 팝업"					' 팝업 명칭 
			arrParam(1) = "B_BIZ_PARTNER" 				' TABLE 명칭 
			arrParam(2) = strCode   		            ' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "거래처"						' 조건필드의 라벨 명칭 

			arrField(0) = "BP_CD"						' Field명(0)
			arrField(1) = "BP_NM"						' Field명(1)
    
			arrHeader(0) = "거래처코드"				    ' Header명(0)
			arrHeader(1) = "거래처명"				    ' Header명(1)
			
		Case 1
			arrParam(0) = "은행 팝업"
			arrParam(1) = "B_BANK"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "은행코드"

			arrField(0) = "BANK_CD"
			arrField(1) = "BANK_NM"
    
			arrHeader(0) = "은행코드"
			arrHeader(1) = "은행명"
			
		Case 2		
			arrParam(0) = "부서 팝업"
			arrParam(1) = "B_ACCT_DEPT A, B_COST_CENTER B, B_BIZ_AREA C"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = "A.ORG_CHANGE_ID =  " & FilterVar(parent.gChangeOrgId , "''", "S") & ""
			arrParam(4) = arrParam(4) & " AND A.COST_CD = B.COST_CD "
			arrParam(4) = arrParam(4) & " AND B.BIZ_AREA_CD = C.BIZ_AREA_CD "
			arrParam(5) = "부서"		

			arrField(0) = "A.DEPT_CD"
			arrField(1) = "A.DEPT_NM"
			arrField(2) = "A.COST_CD"			
			arrField(3) = "A.INTERNAL_CD"			
			arrField(4) = "HH" & parent.gColSep & "C.BIZ_AREA_CD"				
    
			arrHeader(0) = "부서코드"
			arrHeader(1) = "부서명"
			arrHeader(2) = "코스트센터"			
			arrHeader(3) = "내부부서코드"
			
		Case 4
			arrParam(0) = "어음팝업"
			arrParam(1) = "F_NOTE_NO A, B_BANK B"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = "A.BANK_CD = B.BANK_CD AND A.STS=" & FilterVar("NP", "''", "S") & " "
			arrParam(5) = "어음번호"

			arrField(0) = "A.NOTE_NO"
			arrField(1) = "A.BANK_CD"
			arrField(2) = "HH" & parent.gColSep & "B.BANK_NM"
			arrField(3) = "DD" & parent.gColSep & "A.ISSUE_DT"

			arrHeader(0) = "어음번호"
			arrHeader(1) = "은행코드"
			arrHeader(2) = "은행명"
			arrHeader(3) = "발행일"
	End Select
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		     "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EScCode(iwhere)
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
		if iWhere <> 0 then
       	  ggoSpread.Source = frm1.vspdData
           ggoSpread.UpdateRow frm1.vspdData.ActiveRow
        End if
	End If	

End Function

'============================================================
'부서코드 팝업 
'============================================================
Function OpenPopupDept(Byval strCode, iWhere)
	Dim arrRet
	Dim arrParam(5)
	Dim intRetCD
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function


	arrParam(0) = frm1.txtIssueDt1.text
   	arrParam(1) = frm1.txtIssueDt2.text
	arrParam(2) = lgUsrIntCd                           ' 자료권한 Condition  
	frm1.vspdData.Col = C_DEPT_CD
	arrParam(3) = frm1.vspdData.text				
	arrParam(4) = "F"										' 결의일자 상태 Condition  


	arrRet = window.showModalDialog("../../comasp/DeptPopupOrg.asp", Array(window.parent,arrParam), _
			"dialogWidth=500px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call SetActiveCell(frm1.vspdData,C_DEPT_CD,frm1.vspdData.ActiveRow ,"M","X","X")	
		Exit Function
	End If

	With frm1
		.vspdData.Col  = C_DEPT_CD
		.vspdData.Text = arrRet(0)
		.vspdData.Col  = C_DEPT_NM
		.vspdData.Text = arrRet(1)
		.vspdData.Col  = C_ISSUE_DT
		.vspdData.Text = arrRet(3)
		Call vspdData_Change(C_DEPT_CD, .vspdData.Row )	
		Call SetActiveCell(.vspdData,C_DEPT_CD,.vspdData.ActiveRow ,"M","X","X")	
	End With
	
End Function
'========================================================================================================
'	Name : EScCode()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function EScCode(Byval iWhere)

	With frm1

		Select Case iWhere
		    Case 0		        
		    	.txtBPCD.focus 		    	
		    Case 1
		    	Call SetActiveCell(.vspdData,C_BANK_CD,.vspdData.ActiveRow ,"M","X","X")  
		    Case 2
		    	Call SetActiveCell(.vspdData,C_DEPT_CD,.vspdData.ActiveRow ,"M","X","X")
		    Case 3
		    	Call SetActiveCell(.vspdData,C_BP_CD,.vspdData.ActiveRow ,"M","X","X")
		    Case 4
				Call SetActiveCell(.vspdData,C_NOTE_NO,.vspdData.ActiveRow ,"M","X","X")
        End Select

	End With

End Function
'========================================================================================================
'	Name : SetCode()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)

	With frm1

		Select Case iWhere
		    Case 0		        
		    	.txtBPCD.value   = arrRet(0) 		    	
		    	.txtBPNm.value   = arrRet(1)
		    	.txtBPCD.focus 		    	
		    Case 1
		        .vspdData.Col    = C_BANK_CD
		    	.vspdData.text   = arrRet(0) 
		    	.vspdData.Col    = C_BANK_NM
		    	.vspdData.text   = arrRet(1) 
		    	Call SetActiveCell(.vspdData,C_BANK_CD,.vspdData.ActiveRow ,"M","X","X")  
		    Case 2
		        .vspdData.Col    = C_DEPT_CD
		    	.vspdData.text   = arrRet(0) 
		    	.vspdData.Col    = C_DEPT_NM
		    	.vspdData.text   = arrRet(1) 
		    	.vspdData.Col    = C_COST_CD
		    	.vspdData.text   = arrRet(2) 	
		    	.vspdData.Col    = C_INTERNAL_CD
		    	.vspdData.text   = arrRet(3) 
		    	.vspdData.Col    = C_BIZ_AREA_CD
		    	.vspdData.text   = arrRet(4)		    	
		    	Call SetActiveCell(.vspdData,C_DEPT_CD,.vspdData.ActiveRow ,"M","X","X")
		    Case 3
		        .vspdData.Col    = C_BP_CD
		    	.vspdData.text   = arrRet(0)
		    	.vspdData.Col    = C_BP_NM
		    	.vspdData.text   = arrRet(1)
		    	Call SetActiveCell(.vspdData,C_BP_CD,.vspdData.ActiveRow ,"M","X","X")
		    Case 4
		        .vspdData.Col    = C_NOTE_NO
		    	.vspdData.text   = arrRet(0)
		    	.vspdData.Col    = C_BANK_CD
		    	.vspdData.text   = arrRet(1)
		    	.vspdData.Col    = C_BANK_NM
		    	.vspdData.text   = arrRet(2)
		    	.vspdData.Col    = C_ISSUE_DT
		    	.vspdData.text   = arrRet(3)
		    	.vspdData.Col    = C_ISSUE_DT_HH
		    	.vspdData.text   = arrRet(3)
				Call SetActiveCell(.vspdData,C_NOTE_NO,.vspdData.ActiveRow ,"M","X","X")
        End Select

	End With

End Function


'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    
    frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col    
    
	 Select Case Col
	 
	    Case C_NOTE_POP
			frm1.vspdData.Col = C_NOTE_TYPECD
			If UCase(Trim(frm1.vspdData.text)) = "D3" Then
				frm1.vspdData.Col = C_NOTE_NO
				Call OpenPopup(frm1.vspdData.Text, 4)
			End If

	    Case C_BP_POP
	         frm1.vspdData.Col = C_BP_CD   
            Call OpenBp(frm1.vspdData.Text, 3)
	 
	    Case C_DEPT_POP
	         frm1.vspdData.Col = C_DEPT_CD
            Call OpenPopupDept(frm1.vspdData.text, 2)
    
       Case C_BANK_POP
	         frm1.vspdData.Col = C_BANK_CD
            Call OpenPopup(frm1.vspdData.Text, 1)    
    
    End Select    
End Sub


'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row )
    Dim iDx
    Dim IntRetCd 
    Dim strSelect, strFrom, strWhere
	Dim arrVal1, arrVal2
	Dim ii
	dim intIndex
      
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
   SELECT CASE Col
   
         CASE C_DEPT_CD, C_ISSUE_DT
         
'              Call CommonQueryRs("a.dept_cd, a.dept_nm, a.cost_cd, a.internal_cd, c.biz_area_cd", "b_acct_dept a, b_cost_center b, b_biz_area c", _ 
'                                 "a.org_change_id = '"&parent.gChangeOrgId & "' AND a.dept_cd='" & frm1.vspdData.Text & "' AND a.cost_cd = b.cost_cd AND b.biz_area_cd = c.biz_area_cd ", _
'                                 lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
'              If (lgF0 = "") OR (lgF1 = "X") then Exit Sub
'				  frm1.vspdData.Col    = C_DEPT_CD
'		    	  frm1.vspdData.text   = Left(lgF0, Len(lgF0)-1)
'		    	  frm1.vspdData.Col    = C_DEPT_NM
'		    	  frm1.vspdData.text   = Left(lgF1, Len(lgF1)-1)
'		    	  frm1.vspdData.Col    = C_COST_CD
'		    	  frm1.vspdData.text   = Left(lgF2, Len(lgF2)-1)		    	  		    	  
'		    	  frm1.vspdData.Col    = C_INTERNAL_CD
'		    	  frm1.vspdData.text   = Left(lgF3, Len(lgF3)-1)
'		    	  frm1.vspdData.Col    = C_BIZ_AREA_CD
'		    	  frm1.vspdData.text   = Left(lgF4, Len(lgF4)-1)

            frm1.vspdData.Col = C_ISSUE_DT
			If Trim(frm1.vspdData.Text = "") Then	Exit sub

			frm1.vspdData.Col = C_DEPT_CD
			If Trim(frm1.vspdData.Text = "") Then	Exit sub

			strSelect	=			 " A.dept_cd, A.org_change_id, A.internal_cd, A.cost_cd, B.Biz_area_cd"    		
			strFrom		=			 " b_acct_dept A(NOLOCK), b_cost_center B "
			strWhere	=			 " A.cost_cd = b.cost_cd "
			strWhere	= strWhere & " and A.dept_Cd = " & FilterVar(frm1.vspdData.Text, "''", "S")
			strWhere	= strWhere & " and A.org_change_id = (select distinct org_change_id "			
			strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"

			frm1.vspdData.row = Row
			frm1.vspdData.Col = C_ISSUE_DT

			strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.vspdData.Text, gDateFormat,""), "''", "S") & "))"
						
			If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
				IntRetCD = DisplayMsgBox("124600","X","X","X")  

				frm1.vspdData.Col = C_DEPT_CD
				frm1.vspdData.Text = ""
				frm1.vspdData.Col = C_DEPT_NM
				frm1.vspdData.Text = ""
				frm1.vspdData.Col = C_ORG_CHANGE_ID
				frm1.vspdData.Text = ""
			Else 
				arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
							
				For ii = 0 to Ubound(arrVal1,1) - 1
					arrVal2 = Split(arrVal1(ii), chr(11))
					frm1.vspdData.Col = C_ORG_CHANGE_ID
					frm1.vspdData.Text = Trim(arrVal2(2))
					frm1.vspdData.Col = C_INTERNAL_CD
					frm1.vspdData.Text = Trim(arrVal2(3))
					frm1.vspdData.Col = C_COST_CD
					frm1.vspdData.Text = Trim(arrVal2(4))
					frm1.vspdData.Col = C_BIZ_AREA_CD
					frm1.vspdData.Text = Trim(arrVal2(5))

				Next	
					
			End If
		    	  
		 CASE C_BP_CD
		    	  Call CommonQueryRs("bp_cd, bp_nm ", "b_biz_partner", "bp_cd= " & FilterVar(frm1.vspdData.Text, "''", "S") & "", _
                                 lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
              if (lgF0 = "") OR (lgF1 = "X") then Exit Sub                            
              frm1.vspdData.Col    = C_BP_CD
		    	  frm1.vspdData.text   = Left(lgF0, Len(lgF0)-1)
		    	  frm1.vspdData.Col    = C_BP_NM
		    	  frm1.vspdData.text   = Left(lgF1, Len(lgF1)-1)
		    	  
		 CASE C_BANK_CD
		    	  Call CommonQueryRs("bank_cd, bank_nm ", "b_bank", "bank_cd= " & FilterVar(frm1.vspdData.Text, "''", "S") & "", _
                                 lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
              if (lgF0 = "") OR (lgF1 = "X") then Exit Sub              
              frm1.vspdData.Col    = C_BANK_CD
		    	  frm1.vspdData.text   = Left(lgF0, Len(lgF0)-1)
		    	  frm1.vspdData.Col    = C_BANK_NM
		    	  frm1.vspdData.text   = Left(lgF1, Len(lgF1)-1)
		    	  
         CASE C_NOTE_TYPE
        
				
			ggoSpread.Source = frm1.vspdData
			frm1.vspdData.row = frm1.vspdData.ActiveRow         
            frm1.vspdData.Col = C_NOTE_TYPECD
				Select Case UCase(Trim(frm1.vspdData.text))
					Case "D1"	'받을어음 
						frm1.vspdData.Col = C_NOTE_POP
						frm1.vspdData.Text = "0"
						ggoSpread.SpreadLock		C_NOTE_POP,      frm1.vspdData.ActiveRow,	C_NOTE_POP, frm1.vspdData.ActiveRow
						
			            frm1.vspdData.Col = C_CASH_RATE
						frm1.vspdData.Text = "0"
						ggoSpread.SpreadUnLock		C_CASH_RATE,      frm1.vspdData.ActiveRow,C_CASH_RATE,	frm1.vspdData.ActiveRow
						ggoSpread.SSSetRequired     C_CASH_RATE ,   frm1.vspdData.ActiveRow,     frm1.vspdData.ActiveRow
					Case "D3"	'지급어음 
						frm1.vspdData.Col = C_NOTE_POP
						frm1.vspdData.Text = "0"
						ggoSpread.SpreadUnLock		C_NOTE_POP,      frm1.vspdData.ActiveRow,	C_NOTE_POP, frm1.vspdData.ActiveRow
						
			            frm1.vspdData.Col = C_CASH_RATE
						frm1.vspdData.Text = "0"
						ggoSpread.SpreadLock		C_CASH_RATE,      frm1.vspdData.ActiveRow,	C_CASH_RATE, frm1.vspdData.ActiveRow
					Case Else
			            frm1.vspdData.Col = C_CASH_RATE
						frm1.vspdData.Text = "0"
						ggoSpread.SSSetRequired     C_CASH_RATE ,   frm1.vspdData.ActiveRow,     frm1.vspdData.ActiveRow
				End Select
				
				
				frm1.vspdData.Col = Col
		        intIndex = frm1.vspdData.Value 
				frm1.vspdData.Col = C_NOTE_TYPECD
				frm1.vspdData.Value = intIndex
				
		 Case C_PLACE
                frm1.vspdData.Col = Col
                intIndex = frm1.vspdData.Value   
				frm1.vspdData.Col = C_PLACE_CD
				frm1.vspdData.Value = intIndex 
			
			Case C_RCPTFG
                frm1.vspdData.Col = Col
                intIndex = frm1.vspdData.Value   
				frm1.vspdData.Col = C_RCPTFG_CD
				frm1.vspdData.Value = intIndex 
			  		
				
	   END SELECT
	    
    '------ Developer Coding part (End   ) -------------------------------------------------------------- 
	
   	If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
    
End Sub

'========================================================================================================
'   Event Name : txtEmp_no_Onchange             
'   Event Desc :
'========================================================================================================
Function txtBPCD_Onchange()
    
    
End Function


'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
     End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
      	   DbQuery
    	End If
    End if
End Sub


'=======================================================================================================
'   Event Name : 
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtIssueDt1_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssueDt1.Action = 7
        Call SetFocusToDocument("M")
		frm1.txtIssueDt1.Focus
    End If
End Sub

Sub txtIssueDt2_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssueDt2.Action = 7
        Call SetFocusToDocument("M")
		frm1.txtIssueDt2.Focus        
    End If
End Sub


Sub txtDueDt1_DblClick(Button)
    If Button = 1 Then
        frm1.txtDueDt1.Action = 7
        Call SetFocusToDocument("M")
		frm1.txtDueDt1.Focus    
    End If
End Sub

'Sub txtDueDt2_DblClick(Button)
'    If Button = 1 Then
'        frm1.txtDueDt2.Action = 7
'    End If
'End Sub


'=======================================================================================================
'   Event Name : txtValidDt_Keypress(Key)
'   Event Desc : 3rd party control에서 Enter 키를 누르면 조회 실행 
'=======================================================================================================
Sub txtIssueDt1_Keypress(Key)
    If Key = 13 Then
		frm1.txtIssueDt2.focus
        Call MainQuery
    End If
End Sub

Sub txtIssueDt2_Keypress(Key)
    If Key = 13 Then
		frm1.txtIssueDt1.focus
        Call MainQuery
    End If
End Sub

Sub txtDueDt1_Keypress(Key)
    If Key = 13 Then
		frm1.txtIssueDt1.focus
        Call MainQuery
    End If
End Sub

'Sub txtDueDt2_Keypress(Key)
'    If Key = 13 Then
'        MainQuery()
'    End If
'End Sub

Sub txtNOTENO_Keypress(Key)
    If Key = 13 Then
        Call MainQuery
    End If
End Sub

Sub txtBPCD_Keypress(Key)
    If Key = 13 Then
        Call MainQuery
    End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<!-- space Area-->

	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>어음정보일괄등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=LIGHT>&nbsp;</TD>
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
							    <TD CLASS=TD5 NOWRAP>발행일</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtIssueDt1 CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="발행일1" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
								                     <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtIssueDt2 CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="발행일2" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
                                <TD CLASS=TD5 NOWRAP>만기일</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime3 name=txtDueDt1 CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="만기일1" tag="11X1" VIEWASTEXT></OBJECT>');</SCRIPT>
								</TD>								                     
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>어음구분</TD>
								<TD CLASS=TD6 NOWRAP><SELECT ID="cboNoteFg" NAME="cboNoteFg" ALT="어음구분" STYLE="WIDTH: 100px" tag="11X"><OPTION></OPTION></SELECT></TD>
								<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
								<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
							</TR>
							<TR>
							    <TD CLASS=TD5 NOWRAP>어음번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtNOTENO" ALT="어음번호" TYPE="Text" MAXLENGTH=30 SIZE=30  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnNOTE_CD" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenNoteInfo()">
								                     </TD>	
								<TD CLASS=TD5 NOWRAP>거래처</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBPCD" ALT="거래처" TYPE="Text" MAXLENGTH=13 SIZE=13  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBP_CD" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBp(frm1.txtBPCD.value, 0)">
								                     <INPUT NAME="txtBPNM" TYPE="Text" MAXLENGTH=30 SIZE=20  tag="14XXXU"></TD>	                                
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
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
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
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioKind" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioType" tag="24">
<INPUT TYPE=HIDDEN NAME="txtCheck" tag="24">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
