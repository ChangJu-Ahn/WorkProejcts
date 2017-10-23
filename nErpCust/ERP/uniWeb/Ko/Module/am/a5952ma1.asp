<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        :
'*  3. Program ID           : A5952MA1
'*  4. Program Name         : 월차 계정 코드 등록 
'*  5. Program Desc         : Single-Multi Sample
'*  6. Component List       :
'*  7. Modified date(First) : 2002/01/15
'*  8. Modified date(Last)  : 2003/05/30
'*  9. Modifier (First)     : song sang min
'* 10. Modifier (Last)      : Jung Sung Ki
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
=======================================================================================================-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"     SRC="../ag/AcctCtrl.vbs">           </SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
	

'========================================================================================================
Const BIZ_PGM_ID = "a5952mb1.asp"                                      'Biz Logic ASP 
'========================================================================================================
Const CookieSplit = 1233

Dim C_REG_CD			'월차코드 
Dim C_REG_CD_PB 		'월차코드버튼 
Dim C_REG_CD_NM 		'월차명 
Dim C_ACCT_CODE			'계정코드 
Dim C_ACCT_PB			'월차코드버튼 
Dim C_ACCT_NM   		'계정명 
Dim C_DR_CR				'차변/대변 
Dim C_ACCT_TYPE			'계정타입 
Dim C_CHOU_HWAN 		'추가/환입 
Dim c_SUN_HU			'선급/후급 

'hidden-------------------------------------------------------
Dim C_DR_CR_H			'차변/대변 HIDDEN
Dim C_ACCT_TYPE_CODE 	'계정타입  HIDDEN
Dim C_ACCT_TYPE_H    
Dim C_CHOU_HWAN_H 		'추가/환입 HIDDEN
Dim c_SUN_HU_H    		'선급/후급 HIDDEN
Dim C_IIK_SON_H   		'이익/손실 HIDDEN
Dim C_SUN_MI_H 	 		'선급/미지급HIDDEN
Dim c_SANG_JONG_H  		'상여종류 HIDDEN
Dim C_ACCT_CODE_H		'계정코드 HIDDEN
Dim C_SANG_JONG_CODE

Dim C_DR_CR_F			'차변/대변 HIDDEN
Dim C_ACCT_TYPE_CODE_F	'계정타입  HIDDEN
Dim C_ACCT_TYPE_F   
Dim C_CHOU_HWAN_F  		'추가/환입 HIDDEN
Dim c_SUN_HU_F     		'선급/후급 HIDDEN
Dim C_IIK_SON_F    		'이익/손실 HIDDEN
Dim C_SUN_MI_F 	 
Dim c_SANG_JONG_F
Dim C_ACCT_CODE_F	 
Dim C_SANG_JONG_CODE_F 
'hidden-------------------------------------------------------

Dim C_IIK_SON     
Dim C_SUN_MI 	
Dim c_SANG_JONG 

Dim C_REG_CD1  	
Dim C_REG_CD_PB1
Dim C_REG_CD_NM1 
'hidden-------------------------------------------------------
Dim C_ACCT_CODE_H1
'hidden-------------------------------------------------------
Dim C_ACCT_CODE1
Dim C_ACCT_PB1	
Dim C_ACCT_NM1  
Dim	C_EVAL_METH_CD		'환평가코드 
Dim C_EVAL_METH_PB		'환평가구분버튼 
Dim	C_EVAL_METH_NM		'환평가이름 
'Dim C_EVAL_METH_H
'Dim C_EVAL_METH
Const TAB1 = 1
Const TAB2 = 2



'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================================================================================
Dim lsConcd
Dim IsOpenPop          

Dim gSelframeFlg			   ' 현재 TAB의 위치를 나타내는 Flag
'========================================================================================================
Sub initSpreadPosVariables()         '1.2 변수에 Constants 값을 할당 
'☜:--------Spreadsheet #1-----------------------------------------------------------------------------   
	C_REG_CD  			= 1		'월차코드 
	C_REG_CD_PB			= 2		'월차코드버튼 
	C_REG_CD_NM			= 3		'월차명 
	C_ACCT_CODE			= 4		'계정코드 
	C_ACCT_PB			= 5		'월차코드버튼 
	C_ACCT_NM			= 6		'계정명 
	C_DR_CR				= 7		'차변/대변 
	C_ACCT_TYPE			= 8		'계정타입 
	C_CHOU_HWAN			= 9		'추가/환입 
	c_SUN_HU			= 10	'선급/후급 

'hidden-------------------------------------------------------
	C_DR_CR_H			= 11	'차변/대변 HIDDEN
	C_ACCT_TYPE_CODE	= 12	'계정타입  HIDDEN
	C_ACCT_TYPE_H		= 13
	C_CHOU_HWAN_H		= 14	'추가/환입 HIDDEN
	c_SUN_HU_H			= 15	'선급/후급 HIDDEN
	C_IIK_SON_H			= 16	'이익/손실 HIDDEN
	C_SUN_MI_H 			= 17	'선급/미지급HIDDEN
	c_SANG_JONG_H		= 18	'상여종류 HIDDEN
	C_ACCT_CODE_H		= 19	'계정코드 HIDDEN
	C_SANG_JONG_CODE	= 20

	C_DR_CR_F			= 21	'차변/대변 HIDDEN
	C_ACCT_TYPE_CODE_F	= 22	'계정타입  HIDDEN
	C_ACCT_TYPE_F		= 23
	C_CHOU_HWAN_F		= 24	'추가/환입 HIDDEN
	c_SUN_HU_F			= 25	'선급/후급 HIDDEN
	C_IIK_SON_F			= 26	'이익/손실 HIDDEN
	C_SUN_MI_F 			= 27	'선급/미지급HIDDEN
	c_SANG_JONG_F		= 28	'상여종류 HIDDEN
	C_ACCT_CODE_F		= 29	'계정코드 HIDDEN
	C_SANG_JONG_CODE_F	= 30
'hidden-------------------------------------------------------

	C_IIK_SON			= 31	'이익/손실 
	C_SUN_MI 			= 32	'선급/미지급 
	c_SANG_JONG			= 33	'상여종류 

	C_REG_CD1  			= 1		'월차코드 
	C_REG_CD_PB1		= 2		'월차코드버튼 
	C_REG_CD_NM1		= 3		'월차명 
'hidden-------------------------------------------------------
	C_ACCT_CODE_H1		= 4		'계정코드 HIDDEN
'hidden-------------------------------------------------------
	C_ACCT_CODE1		= 5		'계정코드 
	C_ACCT_PB1			= 6		'월차코드버튼 
	C_ACCT_NM1			= 7		'계정명 
	C_EVAL_METH_CD		= 8		'환평가 
	C_EVAL_METH_PB		= 9		'환평가 
	C_EVAL_METH_NM		= 10	'환평가명 
End Sub
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE
	lgBlnFlgChgValue  = False
	lgIntGrpCount     = 0
    lgStrPrevKey      = ""
    lgStrPrevKeyIndex = ""
    lgSortKey         = 1
		
End Sub

'========================================================================================================
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("I", "*","NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
Function CookiePage(ByVal flgs)
End Function

'========================================================================================================
Sub MakeKeyStream(pRow)
   
    lgKeyStream = Trim(Frm1.txtRegcd.Value)    & Parent.gColSep '월차코드 
    lgKeyStream = lgKeyStream & "*" & Parent.gColSep
End Sub        


'========================================================================================================
Sub InitComboBox()
   
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Dim iCodeArr
    Dim iNameArr
    Dim iCodeArr1
    Dim iNameArr1
    Dim iDx
 	
 	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0040", "''", "S") & "  and (minor_cd >=" & FilterVar("2", "''", "S") & "  and minor_cd <=" & FilterVar("9", "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	iCodeArr = lgF0
    iNameArr = lgF1
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0071", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr1 = lgF0
    iNameArr1 = lgF1
    
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	ggoSpread.source = frm1.vspdData
	ggoSpread.SetCombo "DR" & vbtab & "CR" , C_DR_CR_H
	ggoSpread.SetCombo "차변" & vbtab & "대변" , C_DR_CR
	
	ggoSpread.SetCombo "*" & vbtab & Replace(iCodeArr1,Chr(11),vbTab), C_ACCT_TYPE_H
	ggoSpread.SetCombo "*" & vbtab & Replace(iNameArr1,Chr(11),vbTab), C_ACCT_TYPE
	
	ggoSpread.SetCombo "*" & vbtab & "1" & vbtab & "2"  , C_CHOU_HWAN_H
	ggoSpread.SetCombo "*" & vbtab & "추가" & vbtab & "환입"  , C_CHOU_HWAN
	
	ggoSpread.SetCombo "*" & vbtab & "1" & vbtab & "2"  , c_SUN_HU_H
	ggoSpread.SetCombo "*" & vbtab & "선급" & vbtab & "후급"  , c_SUN_HU
	
	ggoSpread.SetCombo "*" & vbtab & "1" & vbtab & "2"  , C_IIK_SON_H
	ggoSpread.SetCombo "*" & vbtab & "이익" & vbtab & "손실"  , C_IIK_SON
	
	ggoSpread.SetCombo "*" & vbtab & "1" & vbtab & "2"  , C_SUN_MI_H
	ggoSpread.SetCombo "*" & vbtab & "선급" & vbtab & "미지급" , C_SUN_MI
	
	ggoSpread.SetCombo "*" & vbtab & Replace(iCodeArr,Chr(11),vbTab), c_SANG_JONG_CODE
	ggoSpread.SetCombo "*" & vbtab & Replace(iNameArr,Chr(11),vbTab), c_SANG_JONG

'    ggoSpread.Source = frm1.vspdData1
'    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = 'A1045'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
'    iCodeArr1 = lgF0
'    iNameArr1 = lgF1
'	ggoSpread.SetCombo Replace(iCodeArr1,Chr(11),vbTab), C_EVAL_METH_H
'	ggoSpread.SetCombo Replace(iNameArr1,Chr(11),vbTab), C_EVAL_METH
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
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
			.Col = C_DR_CR_H
			intIndex = .value
			.col = C_DR_CR
			.value = intindex	
			
			.Row = intRow
			.Col = C_ACCT_TYPE_H
			intIndex = .value
			.col = C_ACCT_TYPE
			.value = intindex	
			
			.Row = intRow
			.Col = C_CHOU_HWAN_H
			intIndex = .value
			.col = C_CHOU_HWAN
			.value = intindex	
			
			.Row = intRow
			.Col = c_SUN_HU_H
			intIndex = .value
			.col = c_SUN_HU
			.value = intindex	
			
			.Row = intRow
			.Col = C_IIK_SON_H
			intIndex = .value
			.col = C_IIK_SON
			.value = intindex	
			
			.Row = intRow
			.Col = C_SUN_MI_H
			intIndex = .value
			.col = C_SUN_MI
			.value = intindex	
			
			.Row = intRow
			.Col = c_SANG_JONG_CODE
			intIndex = .value
			.col = c_SANG_JONG
			.value = intindex				
		Next	
	End With
End Sub
'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	With frm1.vspdData

		.Row = Row

        Select Case Col
            Case C_DR_CR
                .Col = Col
                intIndex = .Value
				.Col = C_DR_CR_H
				.Value = intIndex

		    Case C_ACCT_TYPE
                .Col = Col
                intIndex = .Value
		        .Col = C_ACCT_TYPE_H
		        .Value = intIndex

		    Case C_CHOU_HWAN
                .Col = Col
                intIndex = .Value
				.Col = C_CHOU_HWAN_H
				.Value = intIndex

		    Case c_SUN_HU
                .Col = Col
                intIndex = .Value
		        .Col = c_SUN_HU_H
		        .Value = intIndex

		    Case C_IIK_SON
                .Col = Col
                intIndex = .Value
				.Col = C_IIK_SON_H
				.Value = intIndex

		    Case C_SUN_MI
                .Col = Col
                intIndex = .Value
		        .Col = C_SUN_MI_H
		        .Value = intIndex

		    Case c_SANG_JONG
                .Col = Col
                intIndex = .Value
				.Col = c_SANG_JONG_CODE
				.Value = intIndex
		End Select
	End With

   	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub
'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
'Sub vspdData1_ComboSelChange(ByVal Col, ByVal Row)
'	Dim intIndex

'	With frm1.vspdData1

'		.Row = Row

'       Select Case Col
'            Case C_EVAL_METH
'                .Col = Col
'                intIndex = .Value
'				.Col = C_EVAL_METH_H
'				.Value = intIndex

'		End Select
'	End With

'    ggoSpread.Source = frm1.vspdData1
'    ggoSpread.UpdateRow Row

'End Sub

'========================================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()      '1.3 [initSpreadPosVariables] 호출 Logic 추가 

	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021103",,Parent.gAllowDragDropSpread            '2.1 [Spreadinit]에 Source Version No. 와 DragDrop지원여부를 설정 
	
	With frm1.vspdData
	
		.ReDraw = false
	
    	.MaxCols   = C_SANG_JONG + 1                                                 ' ☜:☜: Add 1 to Maxcols
		.Col       = .MaxCols                                                        ' ☜:☜: Hide maxcols
		.ColHidden = True                                                            ' ☜:☜:

		ggoSpread.Source= frm1.vspdData
    ggoSpread.ClearSpreadData


        Call GetSpreadColumnPos("A")
        ggoSpread.SSSetEdit     C_REG_CD         ,     "월차 코드"  ,10,,,5,2
        ggoSpread.SSSetButton   C_REG_CD_PB      
        ggoSpread.SSSetEdit     C_REG_CD_NM      ,     "월차명"   	 ,15,,,50,2
        ggoSpread.SSSetEdit     C_ACCT_CODE      ,     "계정코드"      ,12,,,30,2
        ggoSpread.SSSetButton   C_ACCT_PB        
        ggoSpread.SSSetEdit   	C_ACCT_NM		 ,     "계정명" 		   ,22,,,50 
        ggoSpread.SSSetCombo   	C_DR_CR			 ,     "차변/대변"    ,12,2
        ggoSpread.SSSetCombo    C_ACCT_TYPE      ,     "계정타입" 	  ,12,2
        ggoSpread.SSSetCombo    C_CHOU_HWAN      ,     "추가/환입"    ,12,2
        ggoSpread.SSSetCombo   	C_SUN_HU		 ,     "선급/후급"    ,12,2
        
        'hidden colom--------------------------------------------------------------------------------
        ggoSpread.SSSetCombo     C_DR_CR_H        ,     "차변/대변 HIDDEN"  ,12,2
        ggoSpread.SSSetEdit		 C_ACCT_TYPE_CODE    ,   "계정타입코드  HIDDEN"  ,2,,,10,2 
        ggoSpread.SSSetCombo     C_ACCT_TYPE_H    ,     "계정타입  HIDDEN"  ,12,2
        ggoSpread.SSSetCombo     C_CHOU_HWAN_H    ,     "추가/환입 HIDDEN"  ,12,2
        ggoSpread.SSSetCombo     C_SUN_HU_H       ,     "선급/후급 HIDDEN"  ,12,2
        ggoSpread.SSSetCombo     C_IIK_SON_H      ,     "이익/손실 HIDDEN"  ,12,2
        ggoSpread.SSSetCombo     C_SUN_MI_H       ,     "선급/미지급HIDDEN"  ,12,2
        ggoSpread.SSSetEdit      c_SANG_JONG_H    ,     "상여코드 HIDDEN"  ,12,,,15,2 
        ggoSpread.SSSetEdit      C_ACCT_CODE_H    ,     "계정코드 HIDDEN"  ,10,,,10,2 
        ggoSpread.SSSetCombo     C_SANG_JONG_CODE ,     "상여종류코드"  ,12,2
        
        ggoSpread.SSSetEdit     C_DR_CR_F        ,     "차변/대변 HIDDEN"  ,15,,,10,2
        ggoSpread.SSSetEdit		 C_ACCT_TYPE_CODE_F    ,   "계정타입코드  HIDDEN"  ,15,,,10,2
        ggoSpread.SSSetEdit     C_ACCT_TYPE_F    ,     "계정타입  HIDDEN"  ,15,,,10,2
        ggoSpread.SSSetEdit     C_CHOU_HWAN_F    ,     "추가/환입 HIDDEN"  ,15,,,10,2
        ggoSpread.SSSetEdit     C_SUN_HU_F       ,     "선급/후급 HIDDEN"  ,15,,,10,2
        ggoSpread.SSSetEdit     C_IIK_SON_F     ,     "이익/손실 HIDDEN"  ,15,,,10,2
        ggoSpread.SSSetEdit     C_SUN_MI_F      ,     "선급/미지급HIDDEN" ,15,,,10,2
        ggoSpread.SSSetEdit      c_SANG_JONG_F    ,     "상여코드 HIDDEN"  ,12,,,15,2 
        ggoSpread.SSSetEdit      C_ACCT_CODE_F    ,     "계정코드 HIDDEN"  ,10,,,10,2 
        ggoSpread.SSSetEdit     C_SANG_JONG_CODE_F ,     "상여종류코드"  ,15,,,10,2
        'hidden colom--------------------------------------------------------------------------------
        
        ggoSpread.SSSetCombo    C_IIK_SON        ,     "이익/손실"    ,12,2
        ggoSpread.SSSetCombo    C_SUN_MI         ,     "선급/미지급"   ,12,2
        ggoSpread.SSSetCombo   	C_SANG_JONG		 ,     "상여종류"      ,12,2

		Call ggoSpread.SSSetColHidden(C_ACCT_CODE_H,C_ACCT_CODE_H,True)
		Call ggoSpread.SSSetColHidden(C_SANG_JONG_CODE,C_SANG_JONG_CODE,True)
		Call ggoSpread.SSSetColHidden(C_DR_CR_H,C_DR_CR_H,True)
		Call ggoSpread.SSSetColHidden(C_ACCT_TYPE_H,C_ACCT_TYPE_H,True)
		Call ggoSpread.SSSetColHidden(C_ACCT_TYPE_CODE,C_ACCT_TYPE_CODE,True)
		Call ggoSpread.SSSetColHidden(C_CHOU_HWAN_H,C_CHOU_HWAN_H,True)
		Call ggoSpread.SSSetColHidden(c_SUN_HU_H,c_SUN_HU_H,True)
		Call ggoSpread.SSSetColHidden(C_IIK_SON_H,C_IIK_SON_H,True)
		Call ggoSpread.SSSetColHidden(C_SUN_MI_H,C_SUN_MI_H,True)
		Call ggoSpread.SSSetColHidden(c_SANG_JONG_H,c_SANG_JONG_H,True)
		Call ggoSpread.SSSetColHidden(C_DR_CR_F,C_DR_CR_F,True)
		Call ggoSpread.SSSetColHidden(C_ACCT_TYPE_CODE_F,C_ACCT_TYPE_CODE_F,True)
		Call ggoSpread.SSSetColHidden(C_ACCT_TYPE_F,C_ACCT_TYPE_F,True)
		Call ggoSpread.SSSetColHidden(C_CHOU_HWAN_F,C_CHOU_HWAN_F,True)
		Call ggoSpread.SSSetColHidden(c_SUN_HU_F,c_SUN_HU_F,True)
		Call ggoSpread.SSSetColHidden(C_IIK_SON_F,C_IIK_SON_F,True)
		Call ggoSpread.SSSetColHidden(C_SUN_MI_F,C_SUN_MI_F,True)
		Call ggoSpread.SSSetColHidden(c_SANG_JONG_F,c_SANG_JONG_F,True)
		Call ggoSpread.SSSetColHidden(C_ACCT_CODE_F,C_ACCT_CODE_F,True)
		Call ggoSpread.SSSetColHidden(C_SANG_JONG_CODE_F,C_SANG_JONG_CODE_F,True)
		call ggoSpread.MakePairsColumn(C_REG_CD,C_REG_CD_PB)
		call ggoSpread.MakePairsColumn(C_ACCT_CODE,C_ACCT_PB)
		
		
	.ReDraw = true
    
    End With
    Call SetSpreadLock("A") 

End Sub


Sub InitSpreadSheet1()
	Call initSpreadPosVariables()
    
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.Spreadinit "V20021103",,Parent.gAllowDragDropSpread
        
	With frm1.vspdData1
	
	   .ReDraw = false
	   
       .MaxCols   = C_EVAL_METH_NM + 1
	   .Col       = .MaxCols
       .ColHidden = True

       .MaxRows = 0
    

        Call GetSpreadColumnPos("B")
        ggoSpread.SSSetEdit     C_REG_CD1         ,     "월차 코드"  ,10,,,2,2
        ggoSpread.SSSetButton   C_REG_CD_PB1      
        ggoSpread.SSSetEdit     C_REG_CD_NM1      ,     "월차명"   	 ,20,,,50
        
        'hidden colom--------------------------------------------------------------------------------
        ggoSpread.SSSetEdit     C_ACCT_CODE_H1    ,     "계정코드 HIDDEN"  ,15,,,10,2 
        'hidden colom--------------------------------------------------------------------------------
        ggoSpread.SSSetEdit     C_ACCT_CODE1      ,     "계정코드"      ,15,,,30
        ggoSpread.SSSetButton   C_ACCT_PB1        
        ggoSpread.SSSetEdit   	C_ACCT_NM1		 ,     "계정명" 		   ,20,,,50
        ggoSpread.SSSetEdit     C_EVAL_METH_CD   ,     "환평가"  ,10,,,5,2
		ggoSpread.SSSetButton   C_EVAL_METH_PB      
		ggoSpread.SSSetEdit   	C_EVAL_METH_NM	 ,     "환평가명"  ,15,,,25 
        
        Call ggoSpread.SSSetColHidden(C_ACCT_CODE_H1,C_ACCT_CODE_H1,True)
'       Call ggoSpread.SSSetColHidden(C_EVAL_METH_CD,C_EVAL_METH_PB,True)
		call ggoSpread.MakePairsColumn(C_REG_CD1,C_REG_CD_PB1)
		call ggoSpread.MakePairsColumn(C_ACCT_CODE1,C_ACCT_PB1)
		Call ggoSpread.MakePairsColumn(C_EVAL_METH_CD,C_EVAL_METH_PB,True)
		
	.ReDraw = true

    End With
    
    Call SetSpreadLock("B") 

End Sub

'======================================================================================================
Sub SetSpreadLock(byval iwhere)
	
   With frm1
		Select Case iwhere
		Case "A"
			ggoSpread.Source = frm1.vspdData
			.vspdData.ReDraw = False
			
			ggoSpread.SpreadLock      C_REG_CD, -1, C_REG_CD
			ggoSpread.SpreadLock	  C_REG_Cd_PB, -1, C_REG_Cd_PB
			ggoSpread.SpreadLock      C_REG_CD_NM, -1, C_REG_CD_NM
			ggoSpread.SpreadLock      C_ACCT_NM, -1, C_ACCT_NM
			ggoSpread.SpreadLock	  C_DR_CR, -1, C_DR_CR
			ggoSpread.SpreadLock    C_EVAL_METH_CD, -1, C_EVAL_METH_CD
			ggoSpread.SpreadLock    C_EVAL_METH_PB, -1, C_EVAL_METH_PB
			ggoSpread.SpreadLock	C_EVAL_METH_NM, -1, C_EVAL_METH_NM
			ggoSpread.SpreadLock	C_ACCT_TYPE, -1, C_ACCT_TYPE
			ggoSpread.SpreadLock	C_CHOU_HWAN, -1, C_CHOU_HWAN
			ggoSpread.SpreadLock	C_SUN_HU, -1, C_SUN_HU
			ggoSpread.SpreadLock	C_IIK_SON, -1, C_IIK_SON
			ggoSpread.SpreadLock	C_SUN_MI, -1, C_SUN_MI
			ggoSpread.SpreadLock	C_SANG_JONG, -1,C_SANG_JONG
			ggoSpread.SSSetRequired	C_ACCT_CODE, -1, C_ACCT_CODE
			ggoSpread.SpreadLock	.vspdData.MaxCols, -1,.vspdData.MaxCols

			.vspdData.ReDraw = True
		Case "B"
		
			ggoSpread.Source = frm1.vspdData1
			.vspdData1.ReDraw = False

			ggoSpread.SpreadLock    C_REG_CD1,    -1, C_REG_CD1
			ggoSpread.SpreadLock	C_REG_Cd_PB1, -1, C_REG_Cd_PB1
			ggoSpread.SpreadLock    C_REG_CD_NM1,    -1, C_REG_CD_NM1
			ggoSpread.SpreadLock    C_ACCT_NM1,   -1, C_ACCT_NM1
			ggoSpread.SpreadLock    C_EVAL_METH_CD, -1, C_EVAL_METH_CD
			ggoSpread.SpreadLock    C_EVAL_METH_PB, -1, C_EVAL_METH_PB
			ggoSpread.SpreadLock	C_EVAL_METH_NM, -1, C_EVAL_METH_NM
			ggoSpread.SSSetRequired	C_ACCT_CODE1, -1, C_ACCT_CODE1
			ggoSpread.SpreadLock	.vspdData1.MaxCols, -1,.vspdData1.MaxCols
			.vspdData1.ReDraw = True
		
		End Select
    End With
End Sub

Sub SetSpreadColor1()  
	Dim lRow

    With frm1
     ggoSpread.Source = frm1.vspdData
 	.vspdData.ReDraw = False
	
	For lRow = 1 To .vspdData.MaxRows


		.vspdData.Row = lRow
		.vspdData.Col = C_REG_CD

   If Trim(.vspdData.text) = "09" Then

		ggoSpread.SpreadUnLock 	C_SANG_JONG,lRow,C_SANG_JONG, lRow
		ggoSpread.SSSetRequired	C_SANG_JONG,lRow,lRow
        
   End If

		ggoSpread.SpreadUnLock	   C_DR_CR, lRow,C_DR_CR, lRow
		ggoSpread.SSSetRequired	C_DR_CR, lRow,lRow
		
		ggoSpread.SpreadUnLock	  C_ACCT_TYPE,lRow, C_ACCT_TYPE, lRow
		ggoSpread.SSSetRequired	C_ACCT_TYPE,lRow, lRow

		ggoSpread.SpreadUnLock	C_CHOU_HWAN,lRow, C_CHOU_HWAN, lRow
		ggoSpread.SSSetRequired	C_CHOU_HWAN,lRow, lRow

		ggoSpread.SpreadUnLock	C_SUN_HU,lRow, C_SUN_HU, lRow
		ggoSpread.SSSetRequired	C_SUN_HU,lRow, lRow

		ggoSpread.SpreadUnLock	C_IIK_SON,lRow, C_IIK_SON, lRow
		ggoSpread.SSSetRequired	C_IIK_SON,lRow, lRow

		ggoSpread.SpreadUnLock	C_SUN_MI,lRow, C_SUN_MI, lRow
		ggoSpread.SSSetRequired	C_SUN_MI,lRow, lRow
		
   next  
	.vspdData.ReDraw = True

    End With
End Sub
Sub SetSpreadColor2()  
	Dim lRow

    With frm1

     ggoSpread.Source = frm1.vspdData1
 	.vspdData1.ReDraw = False
	For lRow = 1 To .vspdData1.MaxRows
		.vspdData1.Row = lRow
		.vspdData1.Col = C_REG_CD1
		If Trim(.vspdData1.text) = "07" Then
		ggoSpread.SpreadUnLock	   C_EVAL_METH_CD, lRow,C_EVAL_METH_CD, lRow
        ggoSpread.SpreadUnLock	   C_EVAL_METH_PB, lRow,C_EVAL_METH_PB, lRow
		ggoSpread.SSSetRequired	C_EVAL_METH_CD, lRow,lRow
		End If
	next   
	.vspdData1.ReDraw = True
    End With
End Sub
'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
	if gSelframeFlg = TAB1 then
    With frm1
        ggoSpread.Source = frm1.vspdData
        .vspdData.ReDraw = False
        ggoSpread.SSSetRequired	C_REG_CD,pvStartRow,pvEndRow
     	ggoSpread.SSSetRequired	C_DR_CR,pvStartRow,pvEndRow
        ggoSpread.SSSetRequired	C_ACCT_TYPE,pvStartRow,pvEndRow
        ggoSpread.SSSetRequired	C_CHOU_HWAN,pvStartRow,pvEndRow
        ggoSpread.SSSetRequired	C_SUN_HU,pvStartRow,pvEndRow
        ggoSpread.SSSetRequired	C_IIK_SON,pvStartRow,pvEndRow
        ggoSpread.SSSetRequired	C_SUN_MI,pvStartRow,pvEndRow
'       ggoSpread.SSSetRequired	C_SANG_JONG,pvStartRow,pvEndRow
        ggoSpread.SSSetRequired	C_ACCT_CODE,pvStartRow,pvEndRow
        ggoSpread.SpreadLock	C_SANG_JONG,-1,C_SANG_JONG
        ggoSpread.SpreadLock	C_REG_CD_NM,-1,C_REG_CD_NM
        ggoSpread.SpreadLock	C_ACCT_NM,-1,C_ACCT_NM
	    .vspdData.ReDraw = True    
    End with
   else 
    With frm1
         ggoSpread.Source = frm1.vspdData1
        .vspdData1.ReDraw = False
        ggoSpread.SSSetRequired	C_REG_CD1,pvStartRow,pvEndRow
        ggoSpread.SSSetRequired	C_ACCT_CODE1,pvStartRow,pvEndRow
        ggoSpread.SpreadLock	C_REG_CD_NM1,-1,C_REG_CD_NM1
        ggoSpread.SpreadLock	C_ACCT_NM1,-1,C_ACCT_NM1
'        ggoSpread.SpreadLock	C_EVAL_METH_CD,-1,C_EVAL_METH_CD
'       ggoSpread.SpreadLock	C_EVAL_METH_PB,-1,C_EVAL_METH_PB
        ggoSpread.SpreadLock	C_EVAL_METH_NM,-1,C_EVAL_METH_NM
		'ggoSpread.SpreadLock	  C_EVAL_METH,pvStartRow,C_EVAL_METH,pvEndRow
        .vspdData1.ReDraw = True
    
   End With
  end if 
End Sub

'======================================================================================================
' Function Name : SetSpreadColorCopy
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColorCopy()
	Dim lRow

    With frm1

     ggoSpread.Source = frm1.vspdData1
 	 .vspdData1.ReDraw = False
		.vspdData1.Row = .vspdData1.ActiveRow
		lRow=.vspdData1.Row
		.vspdData1.Col = C_REG_CD1
		If Trim(.vspdData1.text) = "09" Then
			ggoSpread.SpreadUnLock	  c_SANG_JONG,lRow, c_SANG_JONG, lRow
			ggoSpread.SSSetRequired	c_SANG_JONG,lRow, lRow
		Else
			ggoSpread.SpreadLock	  c_SANG_JONG,lRow, c_SANG_JONG, lRow
		End If
	.vspdData1.ReDraw = True
    End With
End Sub

Sub SetSpreadColorCopy1()
	Dim lRow

    With frm1

     ggoSpread.Source = frm1.vspdData1
 	 .vspdData.ReDraw = False
		.vspdData.Row = .vspdData1.ActiveRow
		lRow=.vspdData.Row
		.vspdData.Col = C_REG_CD
		If Trim(.vspdData.text) = "07" Then
			ggoSpread.SpreadUnLock	  C_EVAL_METH_CD,lRow, C_EVAL_METH_CD, lRow
			ggoSpread.SpreadUnLock	  C_EVAL_METH_PB,lRow, C_EVAL_METH_PB, lRow
			ggoSpread.SSSetRequired	C_EVAL_METH_CD,lRow, lRow
		Elseif Trim(.vspdData.text) = "09" Then
			ggoSpread.SpreadUnLock	  c_SANG_JONG,lRow, c_SANG_JONG, lRow
			ggoSpread.SSSetRequired	c_SANG_JONG,lRow, lRow
        	ggoSpread.SpreadLock	  C_EVAL_METH_CD,lRow, C_EVAL_METH_CD, lRow
        Else 	
			ggoSpread.SpreadLock	  C_EVAL_METH_CD,lRow, C_EVAL_METH_CD, lRow
		End If
		
		
	.vspdData1.ReDraw = True
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

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_REG_CD			=  iCurColumnPos(1)
			C_REG_CD_PB			=  iCurColumnPos(2)
			C_REG_CD_NM			=  iCurColumnPos(3)
			C_ACCT_CODE			=  iCurColumnPos(4)
			C_ACCT_PB			=  iCurColumnPos(5)
			C_ACCT_NM			=  iCurColumnPos(6)
			C_DR_CR				=  iCurColumnPos(7)
			C_ACCT_TYPE			=  iCurColumnPos(8)
			C_CHOU_HWAN			=  iCurColumnPos(9)
			c_SUN_HU			=  iCurColumnPos(10)
			C_DR_CR_H			=  iCurColumnPos(11)
			C_ACCT_TYPE_CODE	=  iCurColumnPos(12)
			C_ACCT_TYPE_H		=  iCurColumnPos(13)
			C_CHOU_HWAN_H		=  iCurColumnPos(14)
			c_SUN_HU_H			=  iCurColumnPos(15)
			C_IIK_SON_H			=  iCurColumnPos(16)
			C_SUN_MI_H			=  iCurColumnPos(17)
			c_SANG_JONG_H		=  iCurColumnPos(18)
			C_ACCT_CODE_H		=  iCurColumnPos(19)
			C_SANG_JONG_CODE	=  iCurColumnPos(20)
			C_DR_CR_F			=  iCurColumnPos(21)
			C_ACCT_TYPE_CODE_F	=  iCurColumnPos(22)
			C_ACCT_TYPE_F		=  iCurColumnPos(23)
			C_CHOU_HWAN_F		=  iCurColumnPos(24)
			c_SUN_HU_F			=  iCurColumnPos(25)
			C_IIK_SON_F			=  iCurColumnPos(26)
			C_SUN_MI_F			=  iCurColumnPos(27)
			c_SANG_JONG_F		=  iCurColumnPos(28)
			C_ACCT_CODE_F		=  iCurColumnPos(29)
			C_SANG_JONG_CODE_F	=  iCurColumnPos(30)
			C_IIK_SON			=  iCurColumnPos(31)
			C_SUN_MI			=  iCurColumnPos(32)
			c_SANG_JONG			=  iCurColumnPos(33)
       Case "B"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
			C_REG_CD1			= iCurColumnPos(1)
			C_REG_CD_PB1		= iCurColumnPos(2)
			C_REG_CD_NM1		= iCurColumnPos(3)    
			C_ACCT_CODE_H1		= iCurColumnPos(4)
			C_ACCT_CODE1		= iCurColumnPos(5)
			C_ACCT_PB1			= iCurColumnPos(6)
			C_ACCT_NM1			= iCurColumnPos(7)
			C_EVAL_METH_CD		= iCurColumnPos(8)
			C_EVAL_METH_PB		= iCurColumnPos(9)
			C_EVAL_METH_NM		= iCurColumnPos(10)
    End Select    
End Sub   
'========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Call ggoOper.LockField(Document, "N")
            
    Call InitSpreadSheet
    Call InitSpreadSheet1
    Call InitVariables
    
    Call SetDefaultVal
    Call InitComboBox
'    Call InitData
    
	gSelframeFlg = TAB1

    Call changeTabs(TAB1)    
    Call SetToolbar("1100110100101111")

    gIsTab     = "Y" ' <- "Yes"의 약자 Y(와이) 입니다.[V(브이)아닙니다]
    gTabMaxCnt = 2   ' Tab의 갯수를 적어 주세요    
	frm1.txtRegcd.focus
End Sub

'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub


'========================================================================================================
Function FncQuery()
    Dim IntRetCD 

    FncQuery = False
    Err.Clear
	
    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    
    Call InitVariables

    Call MakeKeyStream("X")    
    lgCurrentSpd = "M"  ' 월차 계정코드 등록(1)
    
	Call DisableToolBar(Parent.TBC_QUERY)
	If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function
	End If
       
    FncQuery = True
    
End Function

'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False
    Err.Clear
    
    FncNew = True
End Function

'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False
    Err.Clear
    
    FncDelete = True
End Function

'========================================================================================================
Function FncSave()
    Dim IntRetCD
    Dim lRow
    Dim intStrt_dd
    Dim intEnd_dd
    Dim intStrt_mm
    Dim intEnd_mm
    Dim intAllow_kind_cnt
    
    
    FncSave = False
    
    Err.Clear
    
   
    if gSelframeFlg = TAB1 then
        ggoSpread.Source = frm1.vspdData
        If ggoSpread.SSCheckChange = False Then
            IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
            Exit Function
        End If
        lgCurrentSpd = "M"  ' 월차 계정코드 등록 (1)
    else
        ggoSpread.Source = frm1.vspdData1
        If ggoSpread.SSCheckChange = False Then                                       '☜: Check contents area
            IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
            Exit Function
        End If
        lgCurrentSpd = "S"  ' 월차 계정코드 등록 (2)
    end if
    

    if gSelframeFlg = TAB1 then
		ggoSpread.Source = frm1.vspdData
		If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
		   Exit Function
		End If
    else
		ggoSpread.Source = frm1.vspdData1
		If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
		   Exit Function
		End If
	end if
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    
	Call DisableToolBar(Parent.TBC_SAVE)
	If DbSAVE = False Then
		Call RestoreToolBar()
		Exit Function
	End If																'☜: Query db data
        
    FncSave = True                                                              '☜: Processing is OK
    
End Function

'========================================================================================================
Function FncCopy()

    if gSelframeFlg = TAB1 then

		If Frm1.vspdData.MaxRows < 1 Then
		   Exit Function
		End If	 
		ggoSpread.Source = Frm1.vspdData
		With Frm1.VspdData
			 .ReDraw = False
			 If .ActiveRow > 0 Then
				ggoSpread.CopyRow
				Call SetSpreadColor(.ActiveRow,.ActiveRow)
				.ReDraw = True
				.Focus
			 End If
		End With
    Elseif gSelframeFlg = TAB2 then
		If Frm1.vspdData1.MaxRows < 1 Then
		   Exit Function
		End If	 
		ggoSpread.Source = Frm1.vspdData1
		With Frm1.vspdData1
			 .ReDraw = False
			 If .ActiveRow > 0 Then
				ggoSpread.CopyRow
				Call SetSpreadColor(.ActiveRow,.ActiveRow)
				Call SetSpreadColorCopy()
				  .Row = .ActiveRow
				  .Col = C_ACCT_CODE_H1
				  .text=""
				  .Row = .ActiveRow
				  .Col = C_ACCT_NM1
				  .text=""
				  .Col = C_ACCT_CODE1
				  .Row = .ActiveRow
				  .text=""
				  .Col = C_EVAL_METH_CD
				  .Row = .ActiveRow
				  .text=""
				  .Col = C_EVAL_METH_NM
				  .Row = .ActiveRow
				  .text=""
				.ReDraw = True
				.Focus
			 End If
		End With
	Else
	        Exit Function
	End If

    Set gActiveElement = document.ActiveElement   

End Function

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
Function  FncInsertRow(ByVal pvRowCnt) 
	Dim imRow
	Dim ii
	Dim iCurRowPos
    Dim lInsertRows
    On Error Resume Next															'☜: If process fails
    Err.Clear																		'☜: Clear error stat	

    
	FncInsertRow = False															'☜: Processing is NG
	
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)
	Else
	    imRow = AskSpdSheetAddRowCount()
    
		If imRow = "" Then
		    Exit Function
		End If
	End If		
	
    if gSelframeFlg = TAB1 then
	    With frm1
          .vspdData.focus
		.vspdData.ReDraw = False

		ggoSpread.Source = .vspdData

		ggoSpread.InsertRow,imRow
		
		'Call SetSpreadLock
		Call SetSpreadColor(.vspdData.ActiveRow,.vspdData.ActiveRow + imRow - 1)
		
		.vspdData.ReDraw = True
		
        End With
    else
    	With frm1
            .vspdData1.focus
		.vspdData1.ReDraw = False

		ggoSpread.Source = .vspdData1

		ggoSpread.InsertRow,imRow
		
		'Call SetSpreadLock
		Call SetSpreadColor(.vspdData1.ActiveRow,.vspdData1.ActiveRow + imRow - 1)
		
		.vspdData1.ReDraw = True
        End With
    end if
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement
    
End Function

'========================================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    
	if gSelframeFlg = TAB1 then
	 If Frm1.vspdData.MaxRows < 1 then
	   Exit function
	 End if
     With Frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
     End With
    ELSE
	 If Frm1.vspdData1.MaxRows < 1 then
	   Exit function
	 End  if   
     With Frm1.vspdData1 
     
    	.focus
    	ggoSpread.Source = frm1.vspdData1 
    	lDelRows = ggoSpread.DeleteRow
     End With
    END IF 
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
Function FncPrint()
    Call parent.FncPrint()
End Function

'========================================================================================================
Function FncPrev() 
    On Error Resume Next
End Function

'========================================================================================================
Function FncNext() 
    On Error Resume Next
End Function

'========================================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)
End Function

'========================================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)
End Function


'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
Sub PopRestoreSpreadColumnInf()
	
	Dim indx

	On Error Resume Next
	Err.Clear 		

	ggoSpread.Source = gActiveSpdSheet
    Select Case Trim(UCase(gActiveSpdSheet.Name))
		Case "VSPDDATA"		
	
'	If gSelframeFlg = TAB1 Then 
		Call ggoSpread.RestoreSpreadInf()
		Call InitSpreadSheet()      
		Call InitComboBox
		
		ggoSpread.Source = gActiveSpdSheet
		Call ggoSpread.ReOrderingSpreadData()
		Call InitData()
		Call SetSpreadColor1
'	else
		Case "VSPDDATA1"			

		Call ggoSpread.RestoreSpreadInf()
		Call InitSpreadSheet1()      
		Call InitComboBox
		
		ggoSpread.Source = gActiveSpdSheet
		Call ggoSpread.ReOrderingSpreadData()
		Call InitData()
		Call SetSpreadColor2

	'end	if
	End Select
	
End Sub



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
Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                                        '☜: Clear err status

	if LayerShowHide(1) = False then
	   Exit Function
	end if
	
	Dim strVal

    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001						         
        strVal = strVal     & "&lgCurrentSpd="       & lgCurrentSpd                      '☜: Next key tag
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex                 '☜: Next key tag
    End With
	
    If lgIntFlgMode = Parent.OPMD_UMODE Then
    Else
    End If
    
	Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic

    DbQuery = True
    
End Function

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
    DIm IntRetCD
 	Dim cr_dr
 	Dim chou_hwan
 	Dim SUN_HU
 	Dim IIK_SON
 	Dim SUN_MI
 	Dim ACCT_TYPE
 	Dim cr_dr_H
 	Dim chou_hwan_H
 	Dim SUN_HU_H
 	Dim IIK_SON_H
 	Dim SUN_MI_H
 	Dim ACCT_TYPE_H
 		
    DbSave = False                                                          
    
    if LayerShowHide(1) = False then
	   Exit Function
	end if


    strVal = ""
    strDel = ""
    lGrpCnt = 1

    if lgCurrentSpd = "M" then  ' 월차 계정코드 등록 (1)
        ggoSpread.Source = frm1.vspdData 
	    With Frm1
           For lRow = 1 To .vspdData.MaxRows

'           
           
               .vspdData.Row = lRow
               .vspdData.Col = 0
               Select Case .vspdData.Text
                   Case ggoSpread.InsertFlag                                      '☜: Create
                                                 		   		strVal = strVal & "C" & Parent.gColSep '0
                                            		       		strVal = strVal & lRow & Parent.gColSep  '1
                     .vspdData.Col = C_REG_CD    	          : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                     .vspdData.Col = C_DR_CR_H       		  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                     .vspdData.Col = C_ACCT_TYPE_H 			  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                     .vspdData.Col = C_CHOU_HWAN_H 	      	  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                     .vspdData.Col = c_SUN_HU_H  			  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                     .vspdData.Col = C_IIK_SON_H  			  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                     .vspdData.Col = C_SUN_MI_H  			  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                     .vspdData.Col = c_SANG_JONG_CODE  		  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep 
                     .vspdData.Col = C_ACCT_CODE  			  : strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep '10
                     
                     
                        lGrpCnt = lGrpCnt + 1
                   Case ggoSpread.UpdateFlag                                      '☜: Update
                    	                                        strVal = strVal & "U" & Parent.gColSep  '0
                    	                                        strVal = strVal & lRow & Parent.gColSep  '1
                     .vspdData.Col = C_REG_CD    	          : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                     .vspdData.Col = C_DR_CR_H       		  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                     .vspdData.Col = C_ACCT_TYPE_H  		  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep '4
                     .vspdData.Col = C_CHOU_HWAN_H 	      	  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                     .vspdData.Col = c_SUN_HU_H  			  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                     .vspdData.Col = C_IIK_SON_H  			  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                     .vspdData.Col = C_SUN_MI_H  			  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                     .vspdData.Col = C_SANG_JONG_CODE 		  : strVal = strVal & Trim(.vspdData.text) & Parent.gColSep '9
                     .vspdData.Col = C_ACCT_CODE  			  : strVal = strVal & Trim(.vspdData.text) & Parent.gColSep '10
                     
                     .vspdData.Col = C_DR_CR_F       	  	  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                     .vspdData.Col = C_ACCT_TYPE_F			  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep '12
                     .vspdData.Col = C_CHOU_HWAN_F 	      	  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                     .vspdData.Col = C_SUN_HU_F  		  	  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                     .vspdData.Col = C_IIK_SON_F  			  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                     .vspdData.Col = C_SUN_MI_F 		  	  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                     .vspdData.Col = C_SANG_JONG_CODE_F		  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep '17
                     .vspdData.Col = C_ACCT_CODE_F	 		  : strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep '18
                    
                        lGrpCnt = lGrpCnt + 1

                   Case ggoSpread.DeleteFlag                                      '☜: Delete
                                                        		strDel = strDel & "D" & Parent.gColSep  '0
                                                  				strDel = strDel & lRow & Parent.gColSep   '1
                     .vspdData.Col = C_REG_CD    	          : strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep '2
                     .vspdData.Col = C_DR_CR_F       		  : strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep '3
                     .vspdData.Col = C_ACCT_TYPE_F  		  : strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep '4
                     .vspdData.Col = C_CHOU_HWAN_F 	      	  : strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep '5
                     .vspdData.Col = C_SUN_HU_F  			  : strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep '6
                     .vspdData.Col = C_IIK_SON_F  			  : strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep '7
                     .vspdData.Col = C_SUN_MI_F  			  : strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep '8
                     .vspdData.Col = C_SANG_JONG_CODE_F  	  : strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep '9
                     .vspdData.Col = C_ACCT_CODE_F  		  : strDel = strDel & Trim(.vspdData.Text) & Parent.gRowSep '10
                     
                     
                        lGrpCnt = lGrpCnt + 1
               End Select
           Next
           .txtMode.value        = Parent.UID_M0002
           .lgCurrentSpd.value   = "M"
           .txtUpdtUserId.value  = Parent.gUsrID
           .txtInsrtUserId.value = Parent.gUsrID
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
                   Case ggoSpread.InsertFlag                                      '☜: Create
                                                 		   		strVal = strVal & "C" & Parent.gColSep '0
                                            		       		strVal = strVal & lRow & Parent.gColSep  '1
                     .vspdData1.Col = C_REG_CD1					: strVal = strVal & Trim(.vspdData1.Text) & Parent.gColSep '2
                     .vspdData1.Col = C_ACCT_CODE1				: strVal = strVal & Trim(.vspdData1.Text) & Parent.gColSep '3
                     .vspdData1.Col = C_EVAL_METH_CD			: strVal = strVal & Trim(.vspdData1.Text) & Parent.gRowSep '4
                        lGrpCnt = lGrpCnt + 1
                   Case ggoSpread.UpdateFlag                                      '☜: Update
                    	                                            strVal = strVal & "U" & Parent.gColSep  '0
                    	                                        strVal = strVal & lRow & Parent.gColSep  '1
					.vspdData1.Col  = C_REG_CD1					: strVal = strVal & Trim(.vspdData1.Text) & Parent.gColSep '2
					.vspdData1.Col = C_ACCT_CODE1				: strVal = strVal & Trim(.vspdData1.Text) & Parent.gColSep '3
					.vspdData1.Col = C_ACCT_CODE_H1				: strVal = strVal & Trim(.vspdData1.Text) & Parent.gColSep '4
					.vspdData1.Col = C_EVAL_METH_CD				: strVal = strVal & Trim(.vspdData1.Text) & Parent.gRowSep '5
                        lGrpCnt = lGrpCnt + 1
                   Case ggoSpread.DeleteFlag                                      '☜: Delete
                                                        		strDel = strDel & "D" & Parent.gColSep  '0
                                                  				strDel = strDel & lRow & Parent.gColSep   '1
                     .vspdData1.Col = C_REG_CD1					: strDel = strDel & Trim(.vspdData1.Text) & Parent.gColSep '2
                     .vspdData1.Col = C_ACCT_CODE1				: strDel = strDel & Trim(.vspdData1.Text) & Parent.gRowSep '10
                        lGrpCnt = lGrpCnt + 1
               End Select
           Next
           .txtMode.value        = Parent.UID_M0002
           .lgCurrentSpd.value   = "S"
           .txtUpdtUserId.value  = Parent.gUsrID
           .txtInsrtUserId.value = Parent.gUsrID
           .txtMaxRows.value     = lGrpCnt-1	
	       .txtSpread.value      = strDel & strVal
	    End With

    end if	
 	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	

    DbSave = True       
   
    Set gActiveElement = document.ActiveElement                                                    
    
End Function

'========================================================================================================
Function DbDelete()
    Dim IntRetCd
    
    FncDelete = False                                                      '⊙: Processing is NG
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                '☆:
        Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If
    
    
	Call DisableToolBar(Parent.TBC_DELETE)
	If DbDELETE = False Then
		Call RestoreToolBar()
		Exit Function
	End If																'☜: Query db data
    
    FncDelete = True                                                        '⊙: Processing is OK


End Function

'========================================================================================================
Function DbQueryOk()													     
 	
 	Dim iRow,intIndex
	Dim varData
 	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE    
    Call ggoOper.LockField(Document, "Q")										'⊙: Lock field

	Set gActiveElement = document.ActiveElement   
    if  frm1.vspdData.MaxRows = 0 and frm1.vspdData1.MaxRows = 0 then
        Call DisplayMsgBox("900014", "X", "X", "X")
    else
        if  frm1.vspdData.MaxRows = 0 and frm1.vspdData1.MaxRows > 0 then
        	Call ClickTab2()
        else
            'Call ClickTab1()
            'Call SetToolbar("1100111100111111")
        end if
    end if

	if  gSelframeFlg = TAB1 then
		frm1.vspdData.focus
        Call SetToolbar("1100111100111111")
    else
		frm1.vspdData1.focus
    	Call SetToolbar("1100111100111111")									
	end if
	
    Call Setspreadcolor1()	
	Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================
Function DbSaveOk()

    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    
    Call InitVariables															'⊙: Initializes local global variables

    Call MakeKeyStream("X")    
    lgCurrentSpd = "M"
    
    Call MainQuery

    
End Function

'========================================================================================================
Function DbDeleteOk()

End Function

'========================================================================================================
'   Event Name : vspdData1_MouseDown
'   Event Desc : 
'========================================================================================================
Sub vspdData1_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC1" Then
       gMouseClickStatus = "SP1CR"
     End If
End Sub    
'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc :
'========================================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
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
'   Event Name : vspdData1_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
End Sub



'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(Col, Row)

	If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else
		Call SetPopupMenuItemInf("1101111111")
	End If
   
    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData.Row = Row
	'	frm1.vspdData.Col = C_MajorCd
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    
    
End Sub

'========================================================================================================
'   Event Name : vspdData1_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData1_Click(Col, Row)
   
   gMouseClickStatus = "SPC1"
	Set gActiveSpdSheet = frm1.vspdData1
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData1
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData1.Row = Row
	'	frm1.vspdData.Col = C_MajorCd
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    Call SetPopupMenuItemInf("1101111111")    

End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
'6.1. SpreadSheet의 이벤트명 ColWidthChange을 추가 
Sub vspdData_ColWidthChange(ByVal Col1, ByVal Col2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(Col1,Col2)

End Sub

'========================================================================================================
'   Event Name : vspdData1_ColWidthChange
'   Event Desc : 
'========================================================================================================
'6.1. SpreadSheet의 이벤트명 ColWidthChange을 추가 
Sub vspdData1_ColWidthChange(ByVal Col1, ByVal Col2)		
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(Col1,Col2)

End Sub


'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(Col , Row)

    Dim iDx
    Dim IntRetCD,EFlag
    Dim REG_CD, REGCD
    DIM ACCT_CD
    DIM SANG_JONG
    Dim ACCT_TYPE
    
    EFlag = False

   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

	
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	
	Select Case Col
		Case C_REG_CD
			REG_CD = Frm1.vspdData.Text
				If REG_CD <>"" Then
				    IntRetCD = CommonQueryRs("a.minor_nm, a.minor_cd","b_minor a, b_configuration b, a_monthly_base c","a.major_cd = " & FilterVar("a1029", "''", "S") & "  and a.minor_type = " & FilterVar("S", "''", "S") & "  and a.major_cd = b.major_cd and a.minor_cd = b.minor_cd and b.seq_no = 1 and b.reference = " & FilterVar("Y", "''", "S") & "  and a.minor_cd = c.reg_cd and c.use_yn = " & FilterVar("Y", "''", "S") & "  AND a.minor_cd =  " & FilterVar(reg_cd , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				    If IntRetCD = False Then
					    Call DisplayMsgBox("800187","X","X","X")
					    Frm1.vspdData.Col = c_reg_Cd
					    Frm1.vspdData.Text = ""
					    Frm1.vspdData.Col = C_REG_CD_NM
					    Frm1.vspdData.Text = ""
					    Frm1.vspdData.Col = Col
					    Frm1.vspdData.Action = 0
					    Set gActiveElement = document.activeElement
					    EFlag = True
				    Else
				      REGCD = Trim(Replace(lgF1,Chr(11),"")) 
				      Select Case  REGCD
						Case "09"
						    ggoSpread.SpreadUnLock	  C_SANG_JONG,Frm1.vspdData.ActiveRow,C_SANG_JONG ,Frm1.vspdData.ActiveRow
							ggoSpread.SSSetRequired	  C_SANG_JONG,Frm1.vspdData.ActiveRow, Frm1.vspdData.ActiveRow
					    Case Else
							ggoSpread.SpreadLock	  C_SANG_JONG,Frm1.vspdData.ActiveRow,C_SANG_JONG ,Frm1.vspdData.ActiveRow
					  End Select	
							Frm1.vspdData.Col = C_REG_CD_NM
							Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
      			    End If
			  End If
	  Case C_ACCT_CODE		
	    ACCT_CD = Frm1.vspdData.Text
				If ACCT_CD <>"" Then
				    IntRetCD = CommonQueryRs("ACCT_NM","A_ACCT","ACCT_CD= " & FilterVar(ACCT_CD, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				    If IntRetCD = False Then
					    Call DisplayMsgBox("110100","x","X","X")
					    Frm1.vspdData.Col = C_ACCT_CODE
					    Frm1.vspdData.Text = ""
					    Frm1.vspdData.Col = C_ACCT_NM
					    Frm1.vspdData.Text = ""
					    Frm1.vspdData.Col = Col
					    Frm1.vspdData.Action = 0
					    Set gActiveElement = document.activeElement
					    EFlag = True
				    Else
					    Frm1.vspdData.Col = C_ACCT_NM
					    Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
			    End If
			End If
	end select 	
	'------ Developer Coding part (End   ) --------------------------------------------------------------

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)

	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

    '------ Developer Coding part (Start ) --------------------------------------------------------------
    '데이터 확인시 틀린데이터에 대해 undo 해준다.
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = 0

    If EFlag And Frm1.vspdData.Text <> ggoSpread.InsertFlag Then
		Call FncCancel()
	End If
	'------ Developer Coding part (End   ) --------------------------------------------------------------
End Sub


'========================================================================================================
'   Event Name : vspdData1_Change
'   Event Desc :
'========================================================================================================
Sub vspdData1_Change(ByVal Col , ByVal Row )
    Dim iDx
    Dim IntRetCD,EFlag
    Dim REG_CD1
    Dim ACCT_CD1

   	Frm1.vspdData1.Row = Row
   	Frm1.vspdData1.Col = Col

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Select Case Col
		Case C_REG_CD1
			REG_CD1 = Frm1.vspdData1.Text
				If REG_CD1 <>"" Then
				    IntRetCD = CommonQueryRs("a.minor_nm,a.minor_cd","b_minor a, b_configuration b, a_monthly_base c","a.major_cd = " & FilterVar("a1029", "''", "S") & "  and a.minor_type = " & FilterVar("S", "''", "S") & "  and  a.major_cd = b.major_cd and a.minor_cd = b.minor_cd and b.seq_no = 2 and b.reference = " & FilterVar("Y", "''", "S") & "  and a.minor_cd = c.reg_cd and c.use_yn = " & FilterVar("Y", "''", "S") & "  AND a.minor_cd =  " & FilterVar(reg_cd1 , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				    If IntRetCD = False Then
					    Call DisplayMsgBox("800187","X","X","X")
					    Frm1.vspdData1.Col = c_reg_Cd
					    Frm1.vspdData1.Text = ""
					    Frm1.vspdData1.Col = C_REG_CD_NM
					    Frm1.vspdData1.Text = ""
					    Frm1.vspdData1.Col = Col
					    Frm1.vspdData1.Action = 0
						ggoSpread.SpreadLock	  C_EVAL_METH,Frm1.vspdData1.ActiveRow,C_EVAL_METH,Frm1.vspdData1.ActiveRow
					    Set gActiveElement = document.activeElement
					    EFlag = True
				    Else

						If Trim(Replace(lgF1,Chr(11),"")) = "07" Then
							ggoSpread.SpreadUnLock	  C_EVAL_METH_CD,Frm1.vspdData1.ActiveRow,C_EVAL_METH_CD ,Frm1.vspdData1.ActiveRow
							ggoSpread.SSSetRequired	C_EVAL_METH_CD,Frm1.vspdData1.ActiveRow, Frm1.vspdData1.ActiveRow
						Else
							ggoSpread.SpreadLock	  C_EVAL_METH_CD,Frm1.vspdData1.ActiveRow,C_EVAL_METH_CD,Frm1.vspdData1.ActiveRow
							Frm1.vspdData1.Col = C_EVAL_METH_CD
							Frm1.vspdData1.Text = ""
							Frm1.vspdData1.Col = C_EVAL_METH_NM
							Frm1.vspdData1.Text = ""
						End If
					    Frm1.vspdData1.Col = C_REG_CD_NM
					    Frm1.vspdData1.Text = Trim(Replace(lgF0,Chr(11),""))
			    End If
			End If
	  Case C_ACCT_CODE1		
	    ACCT_CD1 = Frm1.vspdData1.Text
				If ACCT_CD1 <>"" Then
				    IntRetCD = CommonQueryRs("ACCT_NM","A_ACCT","ACCT_CD= " & FilterVar(ACCT_CD1, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				    If IntRetCD = False Then
					    Call DisplayMsgBox("110100","x","X","X")
					    Frm1.vspdData1.Col = C_ACCT_CODE1
					    Frm1.vspdData1.Text = ""
					    Frm1.vspdData1.Col = C_ACCT_NM1
					    Frm1.vspdData1.Text = ""
					    Frm1.vspdData1.Col = Col
					    Frm1.vspdData1.Action = 0
					    Set gActiveElement = document.activeElement
					    EFlag = True
				    Else
					    Frm1.vspdData1.Col = C_ACCT_NM1
					    Frm1.vspdData1.Text = Trim(Replace(lgF0,Chr(11),""))
			    End If
			End If
	end select 	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
     
    Call CheckMinNumSpread(frm1.vspdData1,Col,Row)
             
   	
	ggoSpread.Source = frm1.vspdData1
    ggoSpread.UpdateRow Row
End Sub



'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
'6.1. SpreadSheet의 이벤트명 ColWidthChange을 추가 
Sub vspdData_ColWidthChange(ByVal Col1, ByVal Col2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(Col1,Col2)

End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ColWidthChange(ByVal Col1, ByVal Col2)		
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(Col1,Col2)

End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
   
    If Row <= 0 Then
       exit sub
    end if
       
    if frm1.vspdData.MaxRows = 0 Then
		exit sub
	end if	
		  
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
  	
End Sub

'========================================================================================================
'   Event Name : vspdData1_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData1_DblClick(ByVal Col, ByVal Row)				
   
    If Row <= 0 Then
       exit sub
    end if
       
    if frm1.vspdData1.MaxRows = 0 Then
		exit sub
	end if	
		  
End Sub


'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then								
           Call DisableToolBar(Parent.TBC_QUERY)
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If
    
End Sub
'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is 
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Select Case Col
	    Case C_REG_CD_PB
	        frm1.vspdData.Col = C_REG_CD
	        Call OpenCost(frm1.vspdData.Text, 1, Row)
	    Case C_ACCT_PB
	        frm1.vspdData.Col = C_ACCT_CODE
	        Call OpenCost(frm1.vspdData.Text, 2, Row)
	End Select
	Call SetActiveCell(frm1.vspdData,Col-1,frm1.vspdData.ActiveRow ,"M","X","X")
End Sub

'========================================================================================================
'   Event Name : vspdData1_ButtonClicked
'   Event Desc : This function is 
'========================================================================================================
Sub vspdData1_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Select Case Col
	    Case C_REG_CD_PB1
	        frm1.vspdData1.Col = C_REG_CD1
	        Call OpenCost1(frm1.vspdData1.Text, 1, Row)
			Call vspdData1_Change(C_REG_CD1 , Row )
	    Case C_ACCT_PB1
	        frm1.vspdData1.Col = C_ACCT_CODE1
	        Call OpenCost1(frm1.vspdData1.Text, 2, Row)
	    Case  C_EVAL_METH_PB
	       frm1.vspdData1.Col = C_EVAL_METH_CD
	       Call OpenCost1(frm1.vspdData1.Text, 3, Row)    
	End Select
	Call SetActiveCell(frm1.vspdData1,Col-1,frm1.vspdData1.ActiveRow ,"M","X","X")
End Sub


'========================================================================================================
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
    	If lgPageNo <> "" Then
           Call DisableToolBar(Parent.TBC_QUERY)
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If

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



Function ClickTab1()
	If gSelframeFlg = TAB1 Then Exit Function
	
	Call changeTabs(TAB1)
  	gSelframeFlg = TAB1
    if  frm1.vspdData.MaxRows > 0 then
        Call SetToolbar("1100111100111111")
    else
        Call SetToolbar("1100110100111111")
    end if
	frm1.txtRegcd.focus
End Function

Function ClickTab2()
	If gSelframeFlg = TAB2 Then Exit Function
	
	Call changeTabs(TAB2)
	gSelframeFlg = TAB2
    if  frm1.vspdData1.MaxRows > 0 then
        Call SetToolbar("1100111100111111")
    else
        Call SetToolbar("1100110100111111")
    end if
	frm1.txtRegcd.focus
End Function



'===========================================================================
' Function Name : OpenCost
' Function Desc : OpenCost Reference Popup
'===========================================================================
Function OpenCost(strCode, iWhere, Row)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)


	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	    Case 1
	    	arrParam(0) = "월차 팝업"										' 팝업 명칭 
	    	arrParam(1) = "b_minor a, b_configuration b, a_monthly_base c"		' TABLE 명칭 
	    	arrParam(2) = Trim(strCode) 	  									' Code Condition
	    	arrParam(3) = "" 													' Name Cindition
	    	arrParam(4) = "a.major_cd = " & FilterVar("a1029", "''", "S") & "  and a.minor_type = " & FilterVar("S", "''", "S") & "  and a.major_cd = b.major_cd and a.minor_cd = b.minor_cd and b.seq_no = 1 and b.reference = " & FilterVar("Y", "''", "S") & "  and a.minor_cd = c.reg_cd and c.use_yn = " & FilterVar("Y", "''", "S") & " " 		' Where Condition
	    	arrParam(5) = "월차코드"										' TextBox 명칭 

	    	arrField(0) = "a.minor_cd"		 									' Field명(0)
	    	arrField(1) = "a.minor_nm"    										' Field명(1)

	    	arrHeader(0) = "월차 코드"										' Header명(0)%>
	    	arrHeader(1) = "월차 코드명"									' Header명(1)%>


	    Case 2
			arrParam(0) = "계정 팝업"										' 팝업 명칭 
			arrParam(1) = "A_ACCT"												<%' TABLE 명칭 %>
			arrParam(2) = Trim(strCode)	     									<%' Code Condition%>
			arrParam(3) = "" 													<%' Name Cindition%>
			arrParam(4) = ""													<%' Where Condition%>
			arrParam(5) = "계정코드"

			arrField(0) = "ACCT_CD"	     	  									<%' Field명(1)%>
			arrField(1) = "ACCT_nm"												<%' Field명(0)%>

			arrHeader(0) = "계정코드"	  									<%' Header명(0)%>
			arrHeader(1) = "계정명"		  									<%' Header명(1)%>
			
			
	End Select

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCost(arrRet, iWhere, Row)
	End If

End Function
'=======================================================================================================

'------------------------------------------  SetCode()  --------------------------------------------------
'	Name : SetCost()
'	Description : OpenCode Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetCost(arrRet, iWhere, Row)

	With frm1
        .vspdData.Row = Row
		Select Case iWhere
		    Case 1
		    	.vspdData.Col = C_REG_CD
		    	.vspdData.text = arrRet(0)
		    	.vspdData.Col = C_REG_CD_NM
		    	.vspdData.text = arrRet(1)
		    	Call vspdData_Change(C_REG_CD, .vspdData.Row)

		    Case 2
			    .vspdData.Col = C_ACCT_CODE
		    	.vspdData.text = arrRet(0)
		    	.vspdData.Col = C_ACCT_NM
		    	.vspdData.text = arrRet(1)
		End Select

		lgBlnFlgChgValue = True
		ggoSpread.Source = frm1.vspdData
		ggoSpread.UpdateRow Row
	End With

End Function


'===========================================================================
' Function Name : OpenCost1
' Function Desc : OpenCost Reference Popup
'===========================================================================
Function OpenCost1(strCode1, iWhere, Row)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)


	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	    Case 1
	    	arrParam(0) = "월차 팝업"										' 팝업 명칭 
	    	arrParam(1) = "b_minor a, b_configuration b, a_monthly_base c"		' TABLE 명칭 
	    	arrParam(2) = Trim(strCode1) 	  									' Code Condition
	    	arrParam(3) = "" 													' Name Cindition
	    	arrParam(4) = "a.major_cd = " & FilterVar("a1029", "''", "S") & "  and a.minor_type = " & FilterVar("S", "''", "S") & "  and  a.major_cd = b.major_cd and a.minor_cd = b.minor_cd and b.seq_no = 2 and b.reference = " & FilterVar("Y", "''", "S") & "  and a.minor_cd = c.reg_cd and c.use_yn = " & FilterVar("Y", "''", "S") & " " 		' Where Condition
	    	arrParam(5) = "월차코드"										' TextBox 명칭 

	    	arrField(0) = "a.minor_cd"		 									' Field명(0)
	    	arrField(1) = "a.minor_nm"    										' Field명(1)%>

	    	arrHeader(0) = "월차 코드"										' Header명(0)%>
	    	arrHeader(1) = "월차 코드명"									' Header명(1)%>

	    Case 2
			frm1.vspdData1.Col = C_REG_CD1
			frm1.vspdData1.Row = frm1.vspdData1.ActiveRow			
			arrParam(0) = "계정 팝업"	
			arrParam(1) = "A_ACCT"
			arrParam(2) = Trim(strCode1)
			arrParam(3) = ""
			If Trim(frm1.vspdData1.text)="07" Then
				arrParam(4) = " FX_EVAL_FG = " & FilterVar("Y", "''", "S") & "  "
			Else
				arrParam(4) = ""
			End If

			arrParam(5) = "계정코드"

			arrField(0) = "ACCT_CD"
			arrField(1) = "ACCT_nm"

			arrHeader(0) = "계정코드"
			arrHeader(1) = "계정명"

		Case 3
		   	arrParam(0) = "환평가구분 팝업"										' 팝업 명칭 
			arrParam(1) = "B_MINOR"												<%' TABLE 명칭 %>
			arrParam(2) = Trim(strCode1)	     									<%' Code Condition%>
			arrParam(3) = "" 													<%' Name Cindition%>
			arrParam(4) = "major_cd = " & FilterVar("a1045", "''", "S") & " "									<%' Where Condition%>
			arrParam(5) = "계정코드"

			arrField(0) = "MINOR_CD"     	  									<%' Field명(1)%>
			arrField(1) = "MINOR_NM"												<%' Field명(0)%>

			arrHeader(0) = "환평가구분코드"	  									<%' Header명(0)%>
			arrHeader(1) = "환평가구분명"		  									<%' Header명(1)%>
		
	End Select

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCost1(arrRet, iWhere, Row)
	End If

End Function


'------------------------------------------  SetCode()  --------------------------------------------------
'	Name : SetCode()
'	Description : OpenCode Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetCost1(arrRet, iWhere, Row)

	With frm1
        .vspdData1.Row = Row
		Select Case iWhere
		    Case 1
		    	.vspdData1.Col = C_REG_CD1
		    	.vspdData1.text = arrRet(0)
		    	.vspdData1.Col = C_REG_CD_NM1
		    	.vspdData1.text = arrRet(1)
		    Case 2
			    .vspdData1.Col = C_ACCT_CODE1
		    	.vspdData1.text = arrRet(0)
		    	.vspdData1.Col = C_ACCT_NM1
		    	.vspdData1.text = arrRet(1)
   		    Case 3
				.vspdData1.Col = C_EVAL_METH_CD
				.vspdData1.text = arrRet(0)
				.vspdData1.Col = C_EVAL_METH_NM
				.vspdData1.text = arrRet(1)	
	
		End Select

		lgBlnFlgChgValue = True
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.UpdateRow Row

	End With

End Function
'======================================================================================================
'	Name : OpenCode()
'	Description : Major PopUp
'=======================================================================================================%>
Function OpenCode()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	Dim aTabPara

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If gSelframeFlg = TAB1 Then
		aTabPara = 1
	Else
		aTabPara = 2
	End If

	arrParam(0) = "월차 팝업"	
	arrParam(1) = "b_minor a, b_configuration b, a_monthly_base c"
	arrParam(2) = frm1.txtRegcd.value
	arrParam(3) = ""
	arrParam(4) = "a.major_cd = " & FilterVar("a1029", "''", "S") & "  and a.minor_type = " & FilterVar("S", "''", "S") & "  and  a.major_cd = b.major_cd and a.minor_cd = b.minor_cd and b.seq_no = " & aTabPara & " and b.reference = " & FilterVar("Y", "''", "S") & "  and a.minor_cd = c.reg_cd and c.use_yn = " & FilterVar("Y", "''", "S") & " "
	arrParam(5) = "월차코드"

   	arrField(0) = "a.minor_cd"
    arrField(1) = "a.minor_nm"


    arrHeader(0) = "월차코드"
    arrHeader(1) = "월차코드명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtRegcd.focus
		Exit Function
	Else
		Call SetCode(arrRet)
	End If

End Function

'======================================================================================================
'	Name : SetCode()
'	Description : 계정코드 Popup에서 Return되는 값 setting
'=======================================================================================================%>
Function SetCode(Byval arrRet)
	With frm1
		.txtRegcd.focus
		.txtRegcd.value = arrRet(0)
		.txtRegnm.value = arrRet(1)
	End With
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>


<BODY TABINDEX="-1" SCROLL="NO">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>월차분개계정등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>월차대상계정등록</font></td>
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
								<TD CLASS=TD5 NOWRAP>월차코드</TD>
								<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE=TEXT NAME="txtRegcd" SIZE=7 MAXLENGTH=5 tag="11XXXU"  ALT="월차코드" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCode1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCode()">
								<INPUT TYPE=TEXT NAME="txtRegnm" SIZE=27 MAXLENGTH=30 tag="14XXXU"  ALT="월차코드명">
								</TD>
								<TD CLASS=TD6 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP>
								</TD>								
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
					<!-- 첫번째 탭 내용 -->
					<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/a5952ma1_vaSpread_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</DIV>
					<!-- 두번째 탭 내용 -->
					<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/a5952ma1_vaSpread_vspdData1.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hMajorCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="lgCurrentSpd" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>

