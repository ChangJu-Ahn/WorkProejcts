<%@ LANGUAGE="VBSCRIPT" %>

<!--
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : 채무관리 
'*  3. Program ID           : a4225ma1
'*  4. Program Name         : 멀티출금반제 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/03/27
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Jeong Yong Kyun
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. History              :
'**********************************************************************************************
 -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<!--'==========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"     SRC="../ag/AcctCtrl.vbs">           </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                                                             '☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_QUERY_ID = "a4125mb1.ASP"
Const BIZ_PGM_SAVE_ID  = "a4125mb2.ASP"
Const BIZ_PGM_DEL_ID   = "a4125mb3.ASP"
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================

Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2
Const TAB3 = 3

' 출금등록 vspdData4
Dim C_AllCTYPE   
Dim C_ALLC_POP   
Dim C_ALLC_NM
Dim C_ETCNO       
Dim C_ETCNO_POP  
Dim C_BPCD
Dim C_BPNM
DIm C_DocCur4
DIm C_DocCurPopup4
DIM C_ExchRate4
DIm C_BALAMT
Dim C_ALLCAMT      
Dim C_ALLCLOCAMT   
Dim C_AllcAcct
Dim C_AllcAcct_POP 
DIm C_AllcAcctNm
Dim C_DESC 
Dim C_BANK_CD
DIm C_BANK_NM
DIm C_BIZCD
Dim C_BIZNM
Dim C_ALLCSEQ      
 
'매출채무 vspdData1
Dim C_APNo  
Dim C_APDUEDT           
Dim C_PAYBPNM
DIm C_DocCur1
DIM C_ExchRate1
Dim C_APBALAMT        
Dim C_APCLSAMT     
Dim C_APCLSLOCAMT  
Dim C_APDCAMT
Dim C_APDCLOCAMT
DIm C_APDcTypeCd
Dim C_APDcTypePopup
Dim C_APDESC       
Dim C_APAMT        
Dim C_APDT 
Dim C_APBALLOCAMT 


'잔액처리 vspdData
Dim  C_ItemSeq		
Dim  C_AcctCd		
Dim  C_AcctPopup	
Dim  C_AcctNm		
Dim  C_DrCrFg		
Dim  C_DrCrNm		
Dim  C_DocCur		
Dim  C_DocCurPopup	
Dim  C_ExchRate	
Dim  C_ItemAmt		
Dim  C_ItemLocAmt	
Dim  C_ItemDesc	
Dim  C_AcctCd2	
	


'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim  strMode
Dim  intItemCnt

 '==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
 '----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim  IsOpenPop          
Dim  lgRetFlag
Dim  gSelframeFlg
Dim  lgCurrRow
Dim lgDBSaveOK


lgDBSaveOK = 0
<%
Dim dtToday
dtToday = GetSvrDate
%>

 '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 

 '#########################################################################################################
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'######################################################################################################### 
'======================================================================================================
' Name : initSpreadPosVariables()
' Description : 그리드(스프래드) 컬럼 관련 변수 초기화 
'=======================================================================================================
Sub initSpreadPosVariables(ByVal pvSpdNo)
	Select Case UCase(Trim(pvSpdNo))
		Case "A"
	' 반제등록 vspdData4
			C_AllCTYPE		= 1 
			C_ALLC_POP		= 2
			C_ALLC_NM		= 3
			C_ETCNO 		= 4
			C_ETCNO_POP		= 5
			C_BPCD			= 6
			C_BPNM			= 7
			C_DocCur4		= 8
			C_DocCurPopup4	= 9
			C_ExchRate4		= 10
			C_BALAMT		= 11
			C_ALLCAMT		= 12    
			C_ALLCLOCAMT	= 13 
			C_AllcAcct		= 14
			C_AllcAcct_POP	= 15
			C_AllcAcctNm	= 16
			C_DESC			= 17
			C_BANK_CD		= 18
			C_BANK_NM		= 19
			C_BIZCD			= 20
			C_BIZNM			= 21
			C_ALLCSEQ		= 22      
		Case "B"
			'매출채무 vspdData1
			 C_APNo			= 1
			 C_APDUEDT      = 2
			 C_PAYBPNM		= 3
			 C_DocCur1		= 4
			 C_ExchRate1	= 5
			 C_APBALAMT		= 6
			 C_APCLSAMT     = 7
			 C_APCLSLOCAMT  = 8
			 C_APDCAMT		= 9
			 C_APDCLOCAMT	= 10	
			 C_APDcTypeCd	= 11
			 C_APDcTypePopup= 12
			 C_APDESC       = 13
			 C_APAMT		= 14
			 C_APDT			= 15
			 C_APBALLOCAMT	= 16
		Case "C"

			C_ItemSeq		= 1 
			C_AcctCd		= 2 
			C_AcctPopup		= 3 
			C_AcctNm		= 4 
			C_DrCrFg		= 5 
			C_DrCrNm		= 6 
			C_DocCur		= 7
			C_DocCurPopup	= 8
			C_ExchRate		= 9
			C_ItemAmt		= 10
			C_ItemLocAmt	= 11
			C_ItemDesc		= 12
			C_AcctCd2		= 13
		End Select	
End Sub 
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub  InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgLngCurRows = 0  
    gSelframeFlg = TAB1
    lgDBSaveOK = 0    
    lgSortKey = 1
End Sub


'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'********************************************************************************************************* 
Sub  SetDefaultVal()	
	
	frm1.txtPAYMDt.text = UniConvDateAToB("<%=dtToday%>", parent.gServerDateFormat,parent.gDateFormat)
	frm1.hOrgChangeId.value = parent.gChangeOrgId	
	frm1.txtTotLocAmt.text		= "0"
	frm1.txtAllcLocAmt.text     = "0"
	frm1.txtApClsLocAmt.text    = "0"
	frm1.txtDrLocAmt.text		= "0"
	frm1.txtCrLocAmt.text		= "0"
		
   lgBlnFlgChgValue = False
End Sub


'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub  LoadInfTB19029()

<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub


'========================================================================================
' Function Name : InitCombo
' Function Desc : This method loads format inf
'======================================================================================== 
Sub InitCombo()
     
    ggoSpread.Source = frm1.vspdData
	
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1012", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_DrCrFg
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_DrCrNm
    
End Sub

'==========================================================================================
'   Event Name : InitData()
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 

	With frm1.vspdData
		For intRow = 1 To .MaxRows
			.Row = intRow

			.col = C_DrCrFg
			intIndex = .value
			.Col = C_DrCrNm
			.value = intindex			
		Next	
	End With

End Sub
'========================================================================================================
' Name : vspdData_ComboSelChange
' Desc : ComboBox에 값 변경시 처리 
'========================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	DIm tmpDrCrFg
	
	With frm1.vspdData
		  .Row = Row
        Select Case Col
            Case C_DrCrNm
       			.Col = Col		       			
				intIndex = .Value
				.Col = C_DrCrFg					
				.Value = intIndex						
				tmpDrCrFg = .text
				Call SetSpread2Color()
				Call DoSum()
		  End Select
		
	End With

   ggoSpread.Source = frm1.vspdData
   ggoSpread.UpdateRow Row

End Sub

'=============================================== 2.2.3 InitSpreadSheet() ================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)
	Call initSpreadPosVariables(pvSpdNo)
	Select Case UCase(Trim(pvSpdNo))
		Case "A"
			With frm1.vspdData4

				.MaxCols = C_ALLCSEQ + 1
				.Col =.MaxCols
				.ColHidden = true
				ggoSpread.Source = frm1.vspdData4
				.Redraw = False	
			
				.MaxRows = 0
				ggoSpread.SpreadInit "V200303220",,parent.gAllowDragDropSpread 

				Call GetSpreadColumnPos(pvSpdNo)
	
				ggoSpread.SSSetEdit		C_AllCTYPE,     "반제유형",     10, 3, , , 2
				ggoSpread.SSSetButton	C_ALLC_POP
				ggoSpread.SSSetEdit		C_ALLC_NM,      "반제유형명",   15, , , 20, 2 
				ggoSpread.SSSetEdit		C_ETCNO ,       "관련번호",     15, 3
				ggoSpread.SSSetButton	C_ETCNO_POP		       
				ggoSpread.SSSetEdit		C_BPCD,         "거래처",        13, 3
				ggoSpread.SSSetEdit     C_BPNM,         "거래처명",     13, 3				
				ggoSpread.SSSetEdit		C_DocCur4,		"거래통화",    9, , , 10, 2
				ggoSpread.SSSetButton	C_DocCurPopup4
				ggoSpread.SSSetFloat	C_ExchRate4,	"환율",         8, Parent.ggExchRateNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
				ggoSpread.SSSetFloat	C_BALAMT,		"잔액(거래)",   15, "A"  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec																 
				ggoSpread.SSSetFloat	C_ALLCAMT,      "반제금액(거래)",15, "A"  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec																 
				ggoSpread.SSSetFloat	C_ALLCLOCAMT,   "반제금액(자국)",15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec        
				ggoSpread.SSSetEdit     C_AllcAcct,		"반제계정코드"  ,12,,,20,2
				ggoSpread.SSSetButton   C_AllcAcct_POP
				ggoSpread.SSSetEdit		C_AllcAcctNm,     "계정코드명", 20, , , 30
				ggoSpread.SSSetEdit		C_DESC,         "비고",			20, 3
				ggoSpread.SSSetEdit		C_BANK_CD,		"은행코드",		15, , ,	30
				ggoSpread.SSSetEdit		C_BANK_NM,		"은행명",		15, , ,	30
				ggoSpread.SSSetEdit		C_BIZCD,		"사업장",       10, 3
				ggoSpread.SSSetEdit		C_BIZNM,		"사업장명",     15, 3
				ggoSpread.SSSetEdit		C_ALLCSEQ,      "순번",			3, 3
				
				Call ggoSpread.MakePairsColumn(C_DocCur4,C_DocCurPopup4)
				Call ggoSpread.MakePairsColumn(C_AllCTYPE,C_ALLC_POP)
				Call ggoSpread.MakePairsColumn(C_AllcAcct,C_AllcAcct_POP)
				Call ggoSpread.SSSetColHidden(C_ALLCSEQ,C_ALLCSEQ,True)				
				Call ggoSpread.SSSetColHidden(C_BPCD,C_BPCD,True)				
				Call ggoSpread.SSSetColHidden(C_BANK_CD,C_BANK_CD,True)	
				
				.Redraw = True
			End With
		
		Case "B"
    
			With frm1.vspdData1
    
				.MaxCols = C_APBALLOCAMT + 1    
				.Col =.MaxCols
				.ColHidden = true
				ggoSpread.Source = frm1.vspdData1
				.Redraw = False	
			
				.MaxRows = 0
				ggoSpread.SpreadInit "V20030323",,parent.gAllowDragDropSpread 

				Call GetSpreadColumnPos(pvSpdNo)
	
				ggoSpread.SSSetEdit  C_APNO,        "채무번호",      18, 2
				ggoSpread.SSSetDate  C_APDUEDT,     "만기일자",      10, 2, parent.gDateFormat  		        
				ggoSpread.SSSetEdit  C_PAYBPNM,     "지급처",          10, 0
				ggoSpread.SSSetEdit  C_DocCur1,     "거래통화",        9, , , 10, 2
				ggoSpread.SSSetFloat C_ExchRate1,   "환율",        10	, Parent.ggExchRateNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
				ggoSpread.SSSetFloat C_APBALAMT,    "채무잔액",      15, "A"  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_APCLSAMT,    "반제금액",      15, "A" ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_APCLSLOCAMT, "반제금액(자국)",15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec																 
				ggoSpread.SSSetFloat C_APDCAMT,		"할인금액",      15, "A" ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_APDCLOCAMT,	"할인금액(자국)",15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec																 
				ggoSpread.SSSetEdit    C_APDcTypeCd,     "할인타입"  ,12,,,20,2
				ggoSpread.SSSetButton  C_APDcTypePopup
				ggoSpread.SSSetEdit  C_APDESC,      "비고",          20, 0
				ggoSpread.SSSetFloat C_APAMT,       "채무액",        15, "A"  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetDate  C_APDT,        "채무일자",      10, 2, parent.gDateFormat
				ggoSpread.SSSetFloat C_APBALLOCAMT, "채무잔액금액(자국)",15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec																 

				Call ggoSpread.MakePairsColumn(C_APDcTypeCd,C_APDcTypePopup)
				Call ggoSpread.SSSetColHidden(C_APBALLOCAMT,C_APBALLOCAMT,True)		
				.Redraw = True

			End With
    
		Case "C"
		
			With frm1.vspdData
				ggoSpread.Source = frm1.vspdData
				ggoSpread.Spreadinit "V20030218",,parent.gAllowDragDropSpread    
				
				.MaxCols = C_AcctCd2 + 1
				.Col = .MaxCols				'☜: 공통콘트롤 사용 Hidden Column
				.ColHidden = True
				.MaxRows = 0
				.ReDraw = False

				Call AppendNumberPlace("6","3","0")
			    Call GetSpreadColumnPos(pvSpdNo)
			   
			    ggoSpread.SSSetEdit   C_ItemSeq,    "순번" , 5, 2, -1, 5
				ggoSpread.SSSetEdit   C_AcctCd,     "계정코드",   15, , , 18
				ggoSpread.SSSetButton C_AcctPopup
				ggoSpread.SSSetEdit   C_AcctNm,     "계정코드명", 20, , , 30
				ggoSpread.SSSetCombo  C_DrCrFg,     "", 8
				ggoSpread.SSSetCombo  C_DrCrNm,     "차대구분",   11
				ggoSpread.SSSetEdit   C_DocCur,     "거래통화",   10, , , 10, 2
			    ggoSpread.SSSetButton C_DocCurPopup
				ggoSpread.SSSetFloat  C_ExchRate,   "환율", 15, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat  C_ItemAmt,    "금액",       15, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat  C_ItemLocAmt, "금액(자국)", 15, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetEdit   C_ItemDesc,   "비  고",     30, , , 128
				ggoSpread.SSSetEdit   C_AcctCd2,    "",     30, , , 128

				
				Call ggoSpread.MakePairsColumn(C_AcctCd,C_AcctPopup)
				Call ggoSpread.MakePairsColumn(C_DrCrFg,C_DrCrNm,"1")
				Call ggoSpread.MakePairsColumn(C_DocCur,C_DocCurPopup)
				
				Call ggoSpread.SSSetColHidden(C_ItemSeq,C_ItemSeq,True)
				Call ggoSpread.SSSetColHidden(C_DrCrFg,C_DrCrFg,True)
				Call ggoSpread.SSSetColHidden(C_AcctCd2,C_AcctCd2,True)

				.ReDraw = True
			                
			End with
			 										'공통콘트롤 사용 Hidden Column
    
	End Select	
	
    intItemCnt = 0
    Call SetSpreadLock(pvSpdNo)
    Call SetSpreadColor(-1, -1)
    
End Sub
'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)

    With frm1
    	Select Case UCase(Trim(pvSpdNo))
			Case "A"
				ggoSpread.source = .vspdData4      
				.vspdData4.ReDraw = False         
				ggoSpread.SpreadLock	C_AllCTYPE,		-1, C_BALAMT, -1
				ggoSpread.SpreadUnLock	C_ALLCAMT,      -1, C_ALLCAMT  , -1 
				ggoSpread.SpreadLock	C_ALLCLOCAMT,   -1, C_ALLCLOCAMT  , -1
				ggoSpread.SpreadUnLock	C_DESC,			-1, C_DESC  , -1 
				ggoSpread.SpreadLock	C_BANK_NM,		-1, C_ALLCSEQ  , -1
				ggoSpread.SpreadLock	C_DocCurPopup4, -1, C_DocCurPopup4  , -1
				ggoSpread.SSSetRequired C_ALLCAMT,		-1, -1
				ggoSpread.SpreadLock	C_AllcAcctNm,	-1, C_AllcAcctNm, -1		
				.vspdData4.ReDraw = True   
			Case "B"         
				ggoSpread.source = .vspdData1      
				ggoSpread.SpreadLock	C_APNo,			-1, C_APBALAMT  , -1      
				ggoSpread.SpreadLock	C_APCLSLOCAMT,	-1, C_APCLSLOCAMT  , -1  
				ggoSpread.SpreadUnLock	C_APDCAMT,      -1, C_APDCAMT  , -1 
				ggoSpread.SpreadLock	C_APDCLOCAMT,   -1, C_APDCLOCAMT  , -1   
				ggoSpread.SSSetRequired C_APCLSAMT,		-1, -1 
				ggoSpread.SpreadLock	C_APAMT,		-1, C_APDT  , -1   
				ggoSpread.SpreadLock	C_APDcTypePopup,  -1, C_APDcTypePopup, -1
				ggoSpread.SpreadLock	C_APDcTypeCd,	-1, C_APDcTypeCd,  -1
				ggoSpread.SpreadunLock	C_APDESC,       -1, C_APDESC  , -1      
				.vspdData1.ReDraw = True   
			 Case "C"
				ggoSpread.source = .vspdData       
				.vspdData.ReDraw = False         
				ggoSpread.SpreadLock	C_ItemSeq       , -1    , C_AcctCd2
				ggoSpread.SpreadunLock	C_AcctCd		, -1    , C_AcctPopup
				ggoSpread.SpreadUnLock	C_DrCrNm		, -1    , C_DrCrNm
				ggoSpread.SpreadUnLock	C_DocCur		, -1    , C_DocCurPopup
				ggoSpread.SpreadUnLock	C_ExchRate		, -1    , C_ExchRate
				ggoSpread.SpreadUnLock	C_ItemAmt		, -1    , C_ItemAmt
				ggoSpread.SpreadLock	C_ItemLocAmt	, -1    , C_ItemLocAmt
				ggoSpread.SpreadUnLock	C_ItemDesc		, -1    , C_ItemDesc
				ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
				.vspdData.ReDraw = True  
         End Select
    End With 
        
End Sub


'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal lRow, ByVal lRow2)
	
	With frm1
		Select Case gSelframeFlg
			Case TAB1
				ggoSpread.Source = .vspdData4	     
				.vspdData4.ReDraw = False 	    
				ggoSpread.SSSetRequired  C_AllCTYPE,     lRow, lRow2    	      
				ggoSpread.SSSetProtected C_ALLC_NM,      lRow, lRow2
				ggoSpread.SSSetRequired  C_AllcAcct,     lRow, lRow2   
				ggoSpread.SSSetProtected  C_AllcAcctNm,  lRow, lRow2  
				ggoSpread.SSSetRequired  C_ETCNO ,		 lRow, lRow2 
				ggoSpread.SSSetProtected C_BPCD,		 lRow, lRow2
				ggoSpread.SSSetProtected C_BPNM,		 lRow, lRow2
				ggoSpread.SSSetProtected C_DocCur4,      lRow, lRow2
				ggoSpread.SSSetProtected C_ExchRate4,    lRow, lRow2
				ggoSpread.SSSetProtected C_DocCurPopup4,    lRow, lRow2
				ggoSpread.SSSetProtected C_BALAMT,       lRow, lRow2
				ggoSpread.SSSetRequired  C_ALLCAMT,      lRow, lRow2
				ggoSpread.SSSetProtected  C_ALLCLOCAMT,  lRow, lRow2
				ggoSpread.SSSetProtected C_BANK_NM,        lRow, lRow2
				ggoSpread.SSSetProtected C_BIZCD,        lRow, lRow2
				ggoSpread.SSSetProtected C_BIZNM,        lRow, lRow2
				.vspdData4.ReDraw = True		
			
			Case TAB2
			
				ggoSpread.Source = .vspdData1	     
				.vspdData1.ReDraw = False 	    
				ggoSpread.SSSetRequired  C_APCLSAMT, lRow, lRow2
				ggoSpread.SpreadLock	C_APDcTypeCd,       -1, C_APDcTypeCd  , -1   
				ggoSpread.SpreadLock	C_APDcTypePopup,       -1, C_APDcTypePopup  , -1      
				
				.vspdData1.ReDraw = True	
			
			Case TAB3	     	     
				 ggoSpread.Source = .vspdData	     
				 .vspdData.ReDraw = False 	  
				ggoSpread.SSSetProtected C_ItemSeq, lRow, lRow2   ' 
				ggoSpread.SSSetProtected C_AcctNm, lRow, lRow2   ' 계정코드명		
				ggoSpread.SSSetRequired C_AcctCd, lRow, lRow2	' 계정코드 
				ggoSpread.SSSetRequired C_DrCrNm, lRow, lRow2	' 차대구분 
				ggoSpread.SSSetRequired C_DocCur, lRow, lRow2	   ' 부서코드 
				ggoSpread.SSSetRequired C_ItemAmt, lRow, lRow2	' 금액 
				ggoSpread.SSSetProtected C_ItemlocAmt, lRow, lRow2  

				 .vspdData.ReDraw = True	          
		End Select				 
	End With	
	
End Sub


'======================================================================================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method call saved columnorder
'=======================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData4

			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)		

			C_AllCTYPE		= iCurColumnPos(1)
			C_ALLC_POP		= iCurColumnPos(2)
			C_ALLC_NM		= iCurColumnPos(3)
			C_ETCNO 		= iCurColumnPos(4)
			C_ETCNO_POP		= iCurColumnPos(5) 
			C_BPCD			= iCurColumnPos(6)
			C_BPNM			= iCurColumnPos(7)
			C_DocCur4		= iCurColumnPos(8)
			C_DocCurPopup4	= iCurColumnPos(9)
			C_ExchRate4		= iCurColumnPos(10)	
			C_BALAMT		= iCurColumnPos(11)
			C_ALLCAMT		= iCurColumnPos(12)
			C_ALLCLOCAMT	= iCurColumnPos(13)
			C_AllcAcct		= iCurColumnPos(14)
			C_AllcAcct_POP	= iCurColumnPos(15)
			C_AllcAcctNM	= iCurColumnPos(16)
			C_DESC			= iCurColumnPos(17)
			C_BANK_CD		= iCurColumnPos(18)
			C_BANK_NM		= iCurColumnPos(19)
			C_BIZCD			= iCurColumnPos(20)
			C_BIZNM			= iCurColumnPos(21)
			C_ALLCSEQ       = iCurColumnPos(22)
		Case "B"
			ggoSpread.Source = frm1.vspdData1

			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)		
						
			C_APNo			= iCurColumnPos(1)
			C_APDUEDT		= iCurColumnPos(2)
			C_PAYBPNM			= iCurColumnPos(3)
			C_DocCur1		= iCurColumnPos(4)
			C_ExchRate1     = iCurColumnPos(5)
			C_APBALAMT		= iCurColumnPos(6)
			C_APCLSAMT		= iCurColumnPos(7)
			C_APCLSLOCAMT	= iCurColumnPos(8)
			C_APDCAMT		= iCurColumnPos(9)
			C_APDCLOCAMT	= iCurColumnPos(10)
			C_APDcTypeCd	= iCurColumnPos(11)
			C_APDcTypePopup	= iCurColumnPos(12)
			C_APDESC		= iCurColumnPos(13)
			C_APAMT			= iCurColumnPos(14)
			C_APDT			= iCurColumnPos(15)
			C_APBALLOCAMT	= iCurColumnPos(16)
		Case "C"
			ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			 C_ItemSeq			= iCurColumnPos(1)
			 C_AcctCd			= iCurColumnPos(2)
			 C_AcctPopup		= iCurColumnPos(3)
			 C_AcctNm			= iCurColumnPos(4)
			 C_DrCrFg			= iCurColumnPos(5)
			 C_DrCrNm			= iCurColumnPos(6)
			 C_DocCur			= iCurColumnPos(7)
			 C_DocCurPopup		= iCurColumnPos(8)
			 C_ExchRate			= iCurColumnPos(9)
			 C_ItemAmt			= iCurColumnPos(10)
			 C_ItemLocAmt		= iCurColumnPos(11)
			 C_ItemDesc			= iCurColumnPos(12)
			 C_AcctCd2			= iCurColumnPos(13)
	End select
End Sub
'======================================================================================================
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'=======================================================================================================
'======================================================================================================
' Function Name : OpenPopupGL
' Function Desc : This method Open The Popup window for GL
'=======================================================================================================
Function OpenPopupGL()
	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName
	iCalledAspName = AskPRAspName("a5120ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function
	
	arrParam(0) = Trim(frm1.txtGlNo.value)							'회계전표번호 
	arrParam(1) = ""												'Reference번호 

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

'======================================================================================================
' Function Name : OpenPopupTempGL
' Function Desc : This method Open The Popup window for TempGL
'=======================================================================================================
Function OpenPopuptempGL()
	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName
	
	iCalledAspName = AskPRAspName("a5130ra1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function
	
	arrParam(0) = Trim(frm1.txtTempGlNo.value)						'회계전표번호 
	arrParam(1) = ""												'Reference번호 

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

'======================================================================================================
'   Function Name : OpenRefPopAP()
'   Function Desc : 채무정보 Popup
'=======================================================================================================
Function OpenRefOpenAP()

	Dim arrRet
	Dim arrParam(8)	
	Dim i
	Dim iCalledAspName
	Dim strAPno

	iCalledAspName = AskPRAspName("a4125ra2")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a4125ra2", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	If gSelframeFlg <> TAB2 Then Exit Function		 		
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtBpCd.value				' 검색조건이 있을경우 파라미터 
	arrParam(1) = frm1.txtBpNm.value			
	arrParam(3) = "M"
	arrParam(6) = ""
   
	strAPno = ""
	
   If frm1.vspdData1.Maxrows > 0 then
		For i = 1 to frm1.vspdData1.maxRows
				frm1.vspdData1.Row = i
				frm1.vspdData1.Col = C_APNO
				IF frm1.vspdData1.text <> "" Then
					strAPno = strAPno & frm1.vspdData1.text & chr(11) 
				End If	
		Next
		
		frm1.vspdData1.Col= C_DocCur1
		arrParam(2) = Trim(frm1.vspdData1.text)
		
	End if
	arrParam(7) = Trim(strAPno)
	
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	If arrRet(0, 0) = "" Then		
		Exit Function
	Else		
		Call SetRefOpenAP(arrRet)
	End If
	
End Function

'======================================================================================================
'   Function Name : SetRefOpenAP(Byval arrRet)
'   Function Desc : 채무정보 Popup 결과 Fresh
'=======================================================================================================
Function SetRefOpenAP(Byval arrRet)
	
	Dim intRtnCnt, strData
	Dim TempRow, I
	Dim j	
	DIM X
	Dim sFindFg
	Dim TempAPClsLocAmt
	
	With frm1
	
		.vspdData1.focus		
		ggoSpread.Source = .vspdData1
		.vspdData1.ReDraw = False	
	
		TempRow = .vspdData1.MaxRows												'☜: 현재까지의 MaxRows

		For I = TempRow to TempRow + Ubound(arrRet, 1) 
			sFindFg	= "N"
			TempAPClsLocAmt= 0
			'//첫번째 spreadsheet에 없는것만 팝업에서 올것이기 때문에 무조건 addrow이다.
			.vspdData1.MaxRows = .vspdData1.MaxRows + 1
			.vspdData1.Row = I + 1				
			.vspdData1.Col = 0
			.vspdData1.Text = ggoSpread.InsertFlag
			
			.vspddata1.Col = C_APNo												
			.vspddata1.text = arrRet(I - TempRow, 0)				
			.vspddata1.Col = C_APDUEDT 												
			.vspddata1.text = arrRet(I - TempRow, 2)				
			.vspddata1.Col = C_PAYBPNM 												
			.vspddata1.text = UCase(arrRet(I - TempRow, 14))
			.vspddata1.Col = C_DocCur1 												
			.vspddata1.text = UCase(arrRet(I - TempRow, 3))
			.vspddata1.Col = C_ExchRate1 												
			.vspddata1.text = arrRet(I - TempRow, 4)				
			.vspddata1.Col = C_APBALAMT
			.vspddata1.text = arrRet(I - TempRow, 7)
			.vspddata1.Col = C_APBALLOCAMT
			.vspddata1.text = arrRet(I - TempRow, 8)	
			.vspddata1.Col = C_APCLSAMT
			
			IF unicdbl(arrRet(I - TempRow, 9)) <> 0 THEN	'반제금액을 입력할때.
				
				.vspddata1.Col = C_APCLSAMT
				.vspddata1.text = arrRet(I - TempRow, 9)

				Call DoMulti(.vspdData1.Row)  ' ApClsLocAmt 
					
			Else
				.vspddata1.Col = C_ApCLSAMT
				.vspddata1.text = arrRet(I - TempRow, 7)
				.vspddata1.Col = C_APCLSLOCAMT
				.vspddata1.text = arrRet(I - TempRow, 8)	
				.vspddata1.Col = C_APDCAMT
				.vspddata1.text="0"
				.vspddata1.Col = C_APDCLOCAMT
				.vspddata1.text="0"					
			End if
			.vspddata1.Col = C_APAMT
			.vspddata1.text = arrRet(I - TempRow, 5)				
			.vspddata1.Col = C_APDT
			.vspddata1.text = arrRet(I - TempRow, 1)	
			
		Next	
		
'		frm1.txtbpCd.Value = arrRet(0, 13)
'		frm1.txtbpNm.Value = arrRet(0, 14)
		
		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData1,-1,-1,C_DocCur1,C_APBALAMT,   "A" ,"I","X","X")
		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData1,-1,-1,C_DocCur1,C_APCLSAMT,   "A" ,"I","X","X")
		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData1,-1,-1,C_DocCur1,C_APAMT,		"A" ,"I","X","X")
		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData1,-1,-1,C_DocCur1,C_APDCAMT,		"A" ,"I","X","X")
		Call DoSum()
		
		Call SetSpreadColor(-1, -1)
		.vspdData1.ReDraw = True
    End With
    
End Function


'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 
'------------------------------------------  OpenDept()  ---------------------------------------
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
	arrParam(4) = "S"							'B :매출 S: 매입 T: 전체 
	arrParam(5) = "PAYTO"									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then	    
		Call EScPopup(iWhere)
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If	
End Function
'======================================================================================================
'   Function Name : OpenPopUp()
'   Function Desc : 
'=======================================================================================================
Function  OpenPopUp(Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
		
	If IsOpenPop = True Then Exit Function
		
	SELECT CASE iWhere
	
	    CASE 1	    
	      arrParam(0)  = "거래처팝업"
			arrParam(1)  = "B_BIZ_PARTNER"
			arrParam(2)  = Trim(frm1.txtBpCd.Value)
			arrParam(3)  = ""
			arrParam(4)  = ""
			arrParam(5)  = "거래처"	
			arrField(0)  = "BP_CD"
			arrField(1)  = "BP_NM"    
			arrHeader(0) = "거래처"
			arrHeader(1) = "거래처명" 			
		CASE 2			'반제유형 카드, 어음일 경우 계정코드팝업 
   			arrParam(0) = "계정코드팝업"							' 팝업 명칭 
			arrParam(1) = "A_Acct A, A_ACCT_GP B, A_JNL_ACCT_ASSN C"    ' TABLE 명칭 
			arrParam(2) = ""											' Code Condition
			arrParam(3) = ""											' Name Condition
		    frm1.vspdData4.Col = C_AllCTYPE
		    frm1.vspdData4.Row = frm1.vspdData4.ActiveRow
			arrParam(4) = "A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A.Acct_cd=C.Acct_CD and C.jnl_cd = " & FilterVar(frm1.vspdData4.Text, "''", "S")         ' Where Condition
			arrParam(4) = arrParam(4) & " and C.trans_type = " & FilterVar("AP020", "''", "S") & " "
			arrParam(5) = "계정코드"								' 조건필드의 라벨 명칭 

			arrField(0) = "A.Acct_CD"									' Field명(0)
			arrField(1) = "A.Acct_NM"									' Field명(1)
			arrField(2) = "B.GP_CD"									' Field명(2)
			arrField(3) = "B.GP_NM"										' Field명(3)
		 
			arrHeader(0) = "계정코드"								' Header명(0)
			arrHeader(1) = "계정코드명"								' Header명(1)
			arrHeader(2) = "그룹코드"								' Header명(2)
			arrHeader(3) = "그룹명"			
		CASE 3	
			frm1.vspdData.Col = C_DocCur
		    frm1.vspdData.Row = frm1.vspdData.ActiveRow
		
			arrParam(0)  = "거래통화팝업"
			arrParam(1)  = "B_CURRENCY"
			arrParam(2)  = Trim(frm1.vspdData.text)
			arrParam(3)  = ""
			arrParam(4)  = ""
			arrParam(5)  = "거래통화"	
			arrField(0)  = "CURRENCY"
			arrField(1)  = "CURRENCY_DESC"    
			arrHeader(0) = "거래통화"
			arrHeader(1) = "거래통화명"
			
		CASE 4	
		    frm1.vspdData.Col = C_ACCTCD
		    frm1.vspdData.Row = frm1.vspdData.ActiveRow
		    		
			arrParam(0) = "계정코드팝업"								' 팝업 명칭 
			arrParam(1) = "A_Acct, A_ACCT_GP" 											' TABLE 명칭 
			arrParam(2) = Trim(frm1.vspdData.Text)								' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = "A_ACCT.GP_CD=A_ACCT_GP.GP_CD AND A_ACCT.DEL_FG <> " & FilterVar("Y", "''", "S") & " "												' Where Condition
			arrParam(5) = "계정코드"									' 조건필드의 라벨 명칭 

			arrField(0) = "A_ACCT.Acct_CD"									' Field명(0)
			arrField(1) = "A_ACCT.Acct_NM"									' Field명(1)
    		arrField(2) = "A_ACCT_GP.GP_CD"									' Field명(2)
			arrField(3) = "A_ACCT_GP.GP_NM"									' Field명(3)
			
			arrHeader(0) = "계정코드"									' Header명(0)
			arrHeader(1) = "계정코드명"									' Header명(1)
			arrHeader(2) = "그룹코드"									' Header명(2)
			arrHeader(3) = "그룹명"										' Header명(3)

			
		CASE 5
		    frm1.vspdData4.Col = C_ETCNO 
		    frm1.vspdData4.Row = frm1.vspdData4.ActiveRow
			
			arrParam(0) = "지불구매카드번호팝업"											' 팝업 명칭 
			arrParam(1) = "f_note a,b_biz_partner b, b_bank c , b_card_co d "					' TABLE 명칭 
			arrParam(2) = Trim( frm1.vspdData4.text)											' Code Condition
			arrParam(3) = ""						
			arrParam(4) = "a.note_sts = " & FilterVar("BG", "''", "S") & "  AND a.note_fg = " & FilterVar("CP", "''", "S") & "  And a.bp_cd = b.bp_cd And a.bank_cd *= c.bank_cd and a.card_co_cd *= d.card_co_cd "		' Where Condition   						
			arrParam(5) = "지불구매카드번호"												' 조건필드의 라벨 명칭						               
						
			arrHeader(0) = "지불구매카드번호"												' Header명(0)
			arrHeader(1) = "금액"															' Header명(1)
			arrHeader(2) = "발행일"															' Header명(1)     
			arrHeader(3) = "거래처"															' Header명(1)
			arrHeader(4) = "거래처명"														' Header명(1)			
			arrHeader(5) = "카드사"															' Header명(1)						
						
			arrField(0) = "Note_no"																' Field명(0)
			arrField(1) =  "F2" & parent.gColSep & "a.Note_amt"									' Field명(1)
			arrField(2) =  "DD" & parent.gColSep & "a.Issue_dt"									' Field명(2)
			arrField(3) = "a.bp_cd"
			arrField(4) = "b.bp_nm"
			arrField(5) = "d.card_co_nm"     
		CASE 6	

		    frm1.vspdData4.Col = C_AllCTYPE
		    frm1.vspdData4.Row = frm1.vspdData4.ActiveRow
		    		     
			arrParam(0) = "반제유형"														' 팝업 명칭						
			arrParam(1) = "(Select  Distinct A.MINOR_CD,A.MINOR_NM,C.REFERENCE From B_MINOR A,B_CONFIGURATION B,B_CONFIGURATION C "
			arrParam(1) = arrParam(1) & "Where  A.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND A.MAJOR_CD = B.MAJOR_CD AND A.MINOR_CD = B.MINOR_CD "   
			arrParam(1) = arrParam(1) & "AND ( (B.SEQ_NO = 3 AND B.REFERENCE = " & FilterVar("PP", "''", "S") & "  )  OR ( B.SEQ_NO = 6 AND B.REFERENCE IN "
			arrParam(1) = arrParam(1) & "(" & FilterVar("CP", "''", "S") & " ," & FilterVar("D3", "''", "S") & " ," & FilterVar("PP", "''", "S") & " ) )) AND  A.MAJOR_CD = C.MAJOR_CD AND A.MINOR_CD =  C.MINOR_CD AND C.SEQ_NO = 6) tmp "
			
			arrParam(2) = Trim(frm1.vspdData4.text)												' Code Condition
			arrParam(3) = ""																	' Name Cindition
			arrParam(5) = "반제유형"														' TextBox 명칭 
		
			arrField(0) = "tmp.MINOR_CD"															' Field명(0)
			arrField(1) = "tmp.MINOR_NM"															'	 Field명(1)
				    
			arrHeader(0) = "반제유형"														' Header명(0)
			arrHeader(1) = "반제유형명"														' Header명(1)		
		
      CASE 7	
		    
			frm1.vspdData4.Col = C_ETCNO 
		    frm1.vspdData4.Row = frm1.vspdData4.ActiveRow
		    
		    arrParam(0)  = "어음번호팝업"
			 arrParam(1)  = "f_note a,b_biz_partner b, b_bank c"
			 arrParam(2)  = Trim(frm1.vspdData4.Text)
			 arrParam(3)  = ""
			 arrParam(4)  = "a.note_sts = " & FilterVar("BG", "''", "S") & "  AND a.note_fg = " & FilterVar("D3", "''", "S") & "  and a.bp_cd = b.bp_cd and a.bank_cd = c.bank_cd"
			 arrParam(5)  = "어음번호"	
			 
			 arrField(0)  = "Note_no"
			 arrField(1)  = "F2" & parent.gColSep  & "CONVERT(VARCHAR(30),a.Note_amt)"
			 arrField(2)  = "DD" & parent.gColSep  & "CONVERT(VARCHAR(40),a.Issue_dt)"
			 arrField(3)  = "a.bp_cd"
			 arrField(4)  = "b.bp_nm"			 
			 arrField(5)  = "c.bank_nm"	
			 
			 arrHeader(0) = "어음번호"
			 arrHeader(1) = "금액"
			 arrHeader(2) = "발행일"
			 arrHeader(3) = "거래처"
			 arrHeader(4) = "거래처명"			 
			 arrHeader(5) = "은행"
		CASE 8	
			frm1.vspdData4.Col = C_DocCur4
		    frm1.vspdData4.Row = frm1.vspdData4.ActiveRow
		
			arrParam(0)  = "거래통화팝업"
			arrParam(1)  = "B_CURRENCY"
			arrParam(2)  = Trim(frm1.vspdData4.text)
			arrParam(3)  = ""
			arrParam(4)  = ""
			arrParam(5)  = "거래통화"	
			arrField(0)  = "CURRENCY"
			arrField(1)  = "CURRENCY_DESC"    
			arrHeader(0) = "거래통화"
			arrHeader(1) = "거래통화명"
		CASE 9
			frm1.vspdData4.Col = C_ETCNO
		    frm1.vspdData4.Row = frm1.vspdData4.ActiveRow
			arrParam(0) = "계좌번호팝업"
			arrParam(1) = "F_DPST, B_BANK"				
			arrParam(2) = Trim(frm1.vspdData4.text)
			arrParam(3) = ""
			
			arrParam(4) = "F_DPST.BANK_CD = B_BANK.BANK_CD "
			arrParam(5) = "계좌번호"			
			
		    arrField(0) = "F_DPST.BANK_ACCT_NO"	
			arrField(1) = "B_BANK.BANK_CD"
			arrField(2) = "B_BANK.BANK_NM"	
		    arrField(3) = "F_DPST.DOC_CUR"	

		    
		    arrHeader(0) = "계좌번호"		
		    arrHeader(1) = "은행코드"	
		    arrHeader(2) = "은행명"	
		    arrHeader(3) = "거래통화"	
		CASE 10			'apDC_ACCT_cD
			frm1.vspdData1.Col = C_ApDcTypeCD
		    frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
   			
   			arrParam(0) = "계정코드팝업"							' 팝업 명칭 
			arrParam(1) = "B_MAJOR A, B_MINOR  B "    ' TABLE 명칭 
			arrParam(2) = Trim(frm1.vspdData1.text)
			arrParam(3) = ""											' Name Condition
	
			arrParam(4) = "A.MAJOR_CD= B.MAJOR_CD AND A.MAJOR_CD=" & FilterVar("A4042", "''", "S") & "  "
			arrParam(5) = "계정코드"								' 조건필드의 라벨 명칭 

			arrField(0) = "B.MINOR_CD"									' Field명(0)
			arrField(1) = "A.MAJOR_NM"									' Field명(1)
		 
			arrHeader(0) = "할인타입"								' Header명(0)
			arrHeader(1) = "할인타입명"								' Header명(1)
	END SELECT
		
	IsOpenPop = True	
	   
	if iWhere = 0 then
		Dim iCalledAspName
		iCalledAspName = AskPRAspName("a4125ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a4125ra1", "X")
			IsOpenPop = False
			Exit Function
		End If
		
		arrRet = window.showModalDialog(iCalledAspName,Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	else
		arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")			
	end if		
		     		
	IsOpenPop = False
		
	If arrRet(0) = "" Then	    
		Call EScPopup(iWhere)
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If

End Function
'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function EScPopup(Byval iWhere)

	With frm1
	
		Select Case iWhere			
		
		   Case 0
				.txtPaymNO.focus					
			Case 1	
				.txtBpCd.focus
			Case 2
				Call SetActiveCell(.vspdData4,C_AllcAcct,.vspdData4.ActiveRow ,"M","X","X")			
			Case 3
				Call SetActiveCell(.vspdData,C_DocCur,.vspdData.ActiveRow ,"M","X","X")							
			Case 4
				Call SetActiveCell(.vspdData,C_ACCTCD,.vspdData.ActiveRow ,"M","X","X")							
			Case 5		'지불구매카드		    
				Call SetActiveCell(.vspdData4,C_ETCNO,.vspdData4.ActiveRow ,"M","X","X")						    	    
			Case 6			    			  
				Call SetActiveCell(.vspdData4,C_AllCTYPE,.vspdData4.ActiveRow ,"M","X","X")							
			Case 7		'어음 
				Call SetActiveCell(.vspdData4,C_ETCNO,.vspdData4.ActiveRow ,"M","X","X")						    
			Case 8
				Call SetActiveCell(.vspdData4,C_DocCur4,.vspdData4.ActiveRow ,"M","X","X")							
			Case 9
				Call SetActiveCell(.vspdData4,C_ETCNO,.vspdData4.ActiveRow ,"M","X","X")					
			Case 10
				Call SetActiveCell(.vspdData1,C_APDcTypeCd,.vspdData1.ActiveRow ,"M","X","X")							
		End Select		
		
	End With
	
End Function
'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetPopup(Byval arrRet,Byval iWhere)

	With frm1
	
		Select Case iWhere			
		
		   Case 0
				.txtPaymNO.value = arrRet(0)
				.txtPaymNO.focus					
			Case 1	
				.txtBpCd.value = arrRet(0)		
				.txtBpNm.value = arrRet(1)
				.txtBpCd.focus
			Case 2
				.vspdData4.Col  = C_AllcAcct
				.vspdData4.Text = arrRet(0)
				.vspdData4.Col  = C_AllcAcctNM
				.vspdData4.Text = arrRet(1)
				Call SetActiveCell(.vspdData4,C_AllcAcct,.vspdData4.ActiveRow ,"M","X","X")			

			Case 3
				.vspdData.Col  = C_DocCur
				.vspdData.Text = arrRet(0)
				call DocCur_OnChange(.vspdData.Text,.vspdData.ActiveRow)	
				Call SetActiveCell(.vspdData,C_DocCur,.vspdData.ActiveRow ,"M","X","X")							
			Case 4
				.vspdData.Col  = C_ACCTCD
				.vspdData.Text = arrRet(0)
				.vspdData.Col  = C_ACCTNM
				.vspdData.Text = arrRet(1)			
				Call vspdData_Change(C_ACCTCD, frm1.vspdData.ActiveRow)
				Call SetActiveCell(.vspdData,C_ACCTCD,.vspdData.ActiveRow ,"M","X","X")							
			Case 5		'지불구매카드		    
			    .vspdData4.Col  = C_ETCNO
			    .vspdData4.Text = arrRet(0)
			    .vspdData4.Col  = C_BALAMT
			    .vspdData4.Text = arrRet(1)
			    .vspdData4.Col  = C_ALLCAMT
			    .vspdData4.Text = arrRet(1)
			    .vspdData4.Col  = C_ALLCLOCAMT
			    .vspdData4.Text = arrRet(1)
			    .vspdData4.Col  = C_BPCD
			    .vspdData4.Text = arrRet(3)
				.vspdData4.Col  = C_BPNM
			    .vspdData4.Text = arrRet(4)				    	
			    .vspdData4.Col  = C_DocCur4
			    .vspdData4.Text = parent.gCurrency
			    .vspdData4.Col  = C_ExchRate4
			    .vspdData4.Text = "1"			    			    
			    Call DoSum()		
				Call SetActiveCell(.vspdData4,C_ETCNO,.vspdData4.ActiveRow ,"M","X","X")						    	    
			Case 6			    			  
			     
			    .vspdData4.Col  = C_AllCTYPE
			    .vspdData4.Text = arrRet(0)
				 .vspdData4.Col  = C_ALLC_NM
				 .vspdData4.Text = arrRet(1)						
				 Call vspdData4_Change(C_AllCTYPE, frm1.vspdData4.ActiveRow)
				Call SetActiveCell(.vspdData4,C_AllCTYPE,.vspdData4.ActiveRow ,"M","X","X")							
			Case 7		'어음 
			    .vspdData4.Col  = C_ETCNO 
			    .vspdData4.Text = arrRet(0)
			    .vspdData4.Col  = C_BPCD
			    .vspdData4.Text = arrRet(3)	
				.vspdData4.Col  = C_BPNM
			    .vspdData4.Text = arrRet(4)		
			    .vspdData4.Col  = C_ALLCAMT
			    .vspdData4.Text = arrRet(1)
			    .vspdData4.Col  = C_ALLCLOCAMT
			    .vspdData4.Text = arrRet(1)
			    .vspdData4.Col  = C_BALAMT
			    .vspdData4.Text = arrRet(1)
			   	.vspdData4.Col  = C_DocCur4
			    .vspdData4.Text = parent.gCurrency
			    .vspdData4.Col  = C_ExchRate4
			    .vspdData4.Text = "1"
			   
			    Call DoSum()			    
				Call SetActiveCell(.vspdData4,C_ETCNO,.vspdData4.ActiveRow ,"M","X","X")						    
			Case 8
				.vspdData4.Col  = C_DocCur4
				.vspdData4.Text = arrRet(0)
				Call DocCur4_OnChange(.vspdData4.Text,.vspdData4.ActiveRow)				
				Call SetActiveCell(.vspdData4,C_DocCur4,.vspdData4.ActiveRow ,"M","X","X")							
			Case 9
				.vspdData4.Col  = C_ETCNO
				.vspdData4.Text = arrRet(0)
				.vspdData4.Col  = C_BANK_CD		
				.vspdData4.Text  = arrRet(1)				
				.vspdData4.Col  = C_BANK_NM		
				.vspdData4.Text  = arrRet(2)				
				.vspdData4.Col  = C_DocCur4		
				.vspdData4.Text  = arrRet(3)

				Call DocCur4_OnChange(.vspdData4.Text,.vspdData4.ActiveRow)			
				Call SetActiveCell(.vspdData4,C_ETCNO,.vspdData4.ActiveRow ,"M","X","X")					
			Case 10
				.vspdData1.Col  = C_APDcTypeCd
				.vspdData1.Text = arrRet(0)
				Call SetActiveCell(.vspdData1,C_APDcTypeCd,.vspdData1.ActiveRow ,"M","X","X")							
		End Select		
		
	End With
	
End Function

'======================================================================================================
'	Name : OpenDept
'	Description : 
'=======================================================================================================%>

Function OpenDept(Byval strCode)
	Dim arrRet
	Dim arrParam(3)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = frm1.txtPaymDT.Text
	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  
	arrParam(3) = "F"									' 결의일자 상태 Condition  
	
	arrRet = window.showModalDialog("../../comasp/DeptPopupDtA2.asp", Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDeptCd.focus
		Exit Function
	Else
		Call SetDept(arrRet)
	End If	
End Function
'------------------------------------------  SetDept()  ------------------------------------------------
'	Name : SetDept()
'	Description : Dept Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetDept(Byval arrRet)
		
	With frm1
		.txtDeptCd.value   = arrRet(0)		
		.txtDeptNm.value   = arrRet(1)
		.horgChangeId.value = arrRet(3)
		.txtPaymDT.text		= arrRet(3)	
		.txtDeptCd.focus		
	End With
	Call txtDeptCD_OnChange()
End Function 

'========================================================================================================= 
'	Name : OpenRefPrePaymNo()
'	Description : Ref 화면을 call한다. 
'=========================================================================================================  
Function OpenRefPrePaymNo(row)
	Dim arrRet
	Dim arrParam(5)
	Dim iCalledAspName
	Dim IntRetCD

	

	iCalledAspName = AskPRAspName("A4107RA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A4107RA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = ""				' 검색조건이 있을경우 파라미터 
	arrParam(1) =	""			
	arrParam(2) = ""				
    arrParam(3) = frm1.txtPaymDT.text			
	arrParam(4) = frm1.txtPaymDT.alt	
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0,0) = "" Then		
		Exit Function
	Else		
		Call SetRefPrePaymNo(arrRet, row)
	End If
	
End Function
 '------------------------------------------  SetRefPrePaymNo()  ---------------------------------------
'	Name : SetRefPrePaymNo()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function  SetRefPrePaymNo(Byval arrRet, byVal Row)
	lgBlnFlgChgValue = True
	With frm1
	
		.vspddata4.Row= Row
		
		.vspddata4.COl	=C_ETCNO 
		.vspddata4.text	= UCase(arrRet(0,0))
		.vspddata4.COl	=C_BPCD
		.vspddata4.text	= UCase(arrRet(0,3))
		.vspddata4.COl	=C_BPNM
		.vspddata4.text	= UCase(arrRet(0,4))
		.vspddata4.COl	=C_DocCur4
		.vspddata4.text	= UCase(arrRet(0,8))
		.vspddata4.COl	=C_ExchRate4
		.vspddata4.text	= arrRet(0,18)
		.vspddata4.COl	=C_BALAMT
		.vspddata4.text	= arrRet(0,11)
		.vspddata4.COl	=C_ALLCAMT
		.vspddata4.text	= arrRet(0,11)
		.vspddata4.COl	=C_ALLCLOCAMT
		.vspddata4.text	= arrRet(0,12)
		.vspddata4.COl	=C_BIZCD
		.vspddata4.text	= UCase(arrRet(0,13))
		.vspddata4.COl	=C_BIZNM
		.vspddata4.text	= UCase(arrRet(0,14))
	 	.vspddata4.COl	=C_AllcAcct
		.vspddata4.text	= UCase(arrRet(0,1))
		.vspddata4.COl	=C_AllcAcctNm
		.vspddata4.text	= UCase(arrRet(0,2))
	End With

	Call Dosum()

End Function
'======================================================================================================
'	기능: Tab Click
'	설명: Tab Click시 필요한 기능을 수행한다.
'=======================================================================================================
Function ClickTab1()
   
	Call changeTabs(TAB1)
	gSelframeFlg = TAB1

	Call MoveJmpClick() 
	Call SetToolbar("1111111100001111") 				 
End Function

Function ClickTab2()
   
	Call changeTabs(TAB2)	 
	gSelframeFlg = TAB2
   Call MoveJmpClick()
   Call SetToolbar("1111100100001111") 
End Function

Function ClickTab3()
	
	Call changeTabs(TAB3)	 
	gSelframeFlg = TAB3
	Call MoveJmpClick()
	Call SetToolbar("1111111100101111")
End Function




'======================================================================================================
'	기능: 
'	설명: Tab Click시 필요한 기능을 수행한다.
'=======================================================================================================
Function MoveJmpClick()

	Select Case gSelframeFlg
		Case TAB1, TAB3
			
			spnArInfo.innerHTML =  "<font color=""#777777"">채무발생정보</font>&nbsp;|&nbsp;"
		Case TAB2
			spnArInfo.innerHTML =  "<a href='vbscript:OpenRefOpenAP()'>채무발생정보</A>&nbsp;|&nbsp;"

	End Select    


   					 
End Function

 '#########################################################################################################
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################
 '******************************************  3.1 Window 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 
'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub  Form_Load()
        
    Call LoadInfTB19029  
        
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, _
							parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)    
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, _
							parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)	'⊙: Load table , B_numeric_format
        
    Call ggoOper.LockField(Document, "N")
    Call InitSpreadSheet("A")	    															'Setup the Spread sheet
    Call InitSpreadSheet("B")	
    Call InitSpreadSheet("C")	
    Call InitCombo()
	Call InitCtrlSpread()
	Call InitCtrlHSpread()
        
    Call InitVariables     
    Call SetDefaultVal
    Call ClickTab1()
	
	 gIsTab     = "Y" 
	 gTabMaxCnt = 3

End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub  Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)
	Call SetPopUpMenuItemInf("0000111111")
    gMouseClickStatus = "SP1C"	'Split 상태코드 
    
    Set gActiveSpdSheet = frm1.vspdData1
	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData1
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If 
		Exit sub   
    End If
End Sub


'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Dim i  
	
	
	Call SetPopUpMenuItemInf("1101111111")
	gMouseClickStatus = "SPC" 'Split 상태코드 
     
    Set gActiveSpdSheet = frm1.vspdData
    If frm1.vspdData.MaxRows = 0 then
	    Exit Sub
	End if
  	
	If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If 
	    Exit Sub
    End If
    
    Call SetSpread2Color
     
    if Col <> C_AcctCd then  Exit Sub
    
	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.row = frm1.vspdData.ActiveRow	
 	
 	frm1.vspdData.Col = C_AcctCd	
   If Len(frm1.vspdData.Text) <= 0 Then	
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData
	End if	
	    
End Sub
'==========================================================================================
'   Event Name : vspdData4_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData4_Click(ByVal Col, ByVal Row)

	Call SetPopUpMenuItemInf("1101111111")

    gMouseClickStatus = "SP5C"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData4
    If frm1.vspdData4.MaxRows = 0 then
	    Exit Sub
	End if
	
    Dim i    
    	
	If Row = 0 Then
        ggoSpread.Source = frm1.vspdData4
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If 
	    Exit Sub
    End If
    

End Sub
'========================================================================================================
' Event Name :vspdData_MouseDown
' Event Desc :Spread Split 상태코드 
'========================================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)

	If Button = 2 and gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub
'========================================================================================================
' Event Name :vspdData1_MouseDown
' Event Desc :Spread Split 상태코드 
'========================================================================================================
Sub vspdData1_MouseDown(Button, Shift, X, Y)

	If Button = 2 and gMouseClickStatus = "SP1C" Then
		gMouseClickStatus = "SP1CR"
	End If

End Sub
'========================================================================================================
' Event Name :vspdData4_MouseDown
' Event Desc :Spread Split 상태코드 
'========================================================================================================
Sub vspdData4_MouseDown(Button, Shift, X, Y)

	If Button = 2 and gMouseClickStatus = "SP5C" Then
		gMouseClickStatus = "SP5CR"
	End If

End Sub

'======================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 상세내역 그리드의 (멀티)컬럼의 너비를 조절하는 경우 
'=======================================================================================================
Sub  vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'======================================================================================================
'   Event Name : vspdData1_ColWidthChange
'   Event Desc : 상세내역 그리드의 (멀티)컬럼의 너비를 조절하는 경우 
'=======================================================================================================
Sub  vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspdData1
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'======================================================================================================
'   Event Name : vspdData1_ColWidthChange
'   Event Desc : 상세내역 그리드의 (멀티)컬럼의 너비를 조절하는 경우 
'=======================================================================================================
Sub  vspdData4_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspdData4
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'======================================================================================================
'   Event Name :vspddata_ScriptDragDropBlock
'   Event Desc :
'=======================================================================================================
Sub  vspddata_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	ggoSpread.Source = frm1.vspdData 
	Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	Call GetSpreadColumnPos("C")
End Sub

'======================================================================================================
'   Event Name :vspddata_ScriptDragDropBlock
'   Event Desc :
'=======================================================================================================
Sub  vspddata1_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	ggoSpread.Source = frm1.vspdData1 
	Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	Call GetSpreadColumnPos("B")
End Sub

'======================================================================================================
'   Event Name :vspddata_ScriptDragDropBlock
'   Event Desc :
'=======================================================================================================
Sub  vspddata4_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	ggoSpread.Source = frm1.vspdData4 
	Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	Call GetSpreadColumnPos("A")
End Sub
 '**************************  3.2 HTML Form Element & Object Event처리  **********************************
'	Document의 TAG에서 발생 하는 Event 처리	
'	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
'	Event간 충돌을 고려하여 작성한다.
'********************************************************************************************************* 

 '******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 

'=======================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc :
'=======================================================================================================

Sub  vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    If Row <> NewRow And NewRow > 0 Then
        
        With frm1       
            .vspdData.Row = NewRow
            .vspdData1.Col = 1            
            .vspdData.Col = C_ITEMSEQ
            .hItemSeq.value = .vspdData.Text
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.ClearSpreadData

        End With
        
        frm1.vspdData.Col = 0
        If frm1.vspdData.Text = ggoSpread.DeleteFlag Then
			Exit Sub
        End if
        
        lgCurrRow = NewRow
        
        Call DbQuery2(lgCurrRow)
        
    End If

End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub  vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    if Row < 0 then Exit Sub
	SELECT CASE Col
           CASE C_AcctPopUp
                 Call OpenPopUp(4)              
	        CASE C_DocCurPopup
                 Call OpenPopUp(3)          
    END SELECT

    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,row,row,C_DocCur, C_ItemAmt,"A" ,"I","X","X")   

End Sub
'======================================================================================================
'   Event Name : vspdData1_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub  vspdData1_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    if Row < 0 then Exit Sub
	SELECT CASE Col
           CASE C_APDcTypePopup
                 Call OpenPopUp(10)              
    END SELECT
    
End Sub
'======================================================================================================
'   Event Name : vspdData4_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub  vspdData4_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    
	if Row < 0 then Exit Sub

	SELECT CASE Col
	    CASE C_DocCurPopup4
             Call OpenPopUp(8)  
		CASE C_ALLC_POP
			Call OpenPopup(6)
		CASE C_AllcAcct_POP                
			Call OpenPopup(2)
		CASE C_ETCNO_POP
			frm1.vspdData4.col=C_AllCTYPE
			frm1.vspdData4.row=row
			Select case UCase(frm1.vspdData4.text)
				Case "PP"'선급금 
					Call OpenRefPrePaymNo(Row)
				Case "NP"	' 지급어음 
					Call OpenPopup(7)
				Case "CP"	'지불구매카드 
					Call OpenPopup(5)
				Case "DP"	'예적금 
					Call OpenPopup(9)
			End Select
		                  
    END SELECT
    
    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData4,row,row,C_DocCur4, C_BALAMT,"A" ,"I","X","X")   
    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData4,row,row,C_DocCur4, C_ALLCAMT,"A" ,"I","X","X")         
	
    
End Sub


'======================================================================================================
'   Event Name :vspdData1_EditChange
'   Event Desc :
'=======================================================================================================
Sub  vspdData1_EditChange(ByVal Col , ByVal Row )
   
	With frm1.vspdData1 
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.UpdateRow Row
    End With
                
End Sub

'======================================================================================================
'   Event Name :vspdData4_EditChange
'   Event Desc :
'=======================================================================================================
Sub  vspdData4_EditChange(ByVal Col , ByVal Row )
   
	With frm1.vspdData4 
		ggoSpread.Source = frm1.vspdData4
		ggoSpread.UpdateRow Row
    End With
                
End Sub

'======================================================================================================
'   Event Name :vspdData_EditChange
'   Event Desc :
'=======================================================================================================
Sub  vspdData_EditChange(ByVal Col , ByVal Row )
   
	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		ggoSpread.UpdateRow Row
    End With
                
End Sub
'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub  vspdData_Change(ByVal Col, ByVal Row )

    lgBlnFlgChgValue = True
    Call CheckMinNumSpread(frm1.vspdData,Col,Row)
    
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
    frm1.vspdData.Row = Row
    frm1.vspdData.Col = 0
    
    SELECT CASE Col
		CASE   C_ACCTCD
			frm1.vspdData.Col = 0
			If frm1.vspdData.Text = ggoSpread.InsertFlag Then
			   frm1.vspdData.Col = C_ITEMSEQ
			   frm1.hItemSeq.value = frm1.vspdData.Text
			   frm1.vspdData.Col = C_ACCTCD

			   If Len(frm1.vspdData.Text) > 0 Then
				  frm1.vspdData.Row = Row
				  frm1.vspdData.Col = C_ItemSeq
				  DeleteHsheet frm1.vspdData.Text
			      Call DbQuery3(Row)
			   End If    
			End If 
			
		CASE C_ItemAmt, C_ExchRate	
			Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_DocCur,C_ItemAmt,  "A" ,"X","X")
			Call DoMulti(Row)	
		    Call DoSum()
		    
		CASE C_DocCur	
			frm1.vspdData.Col= Col
			Call DocCur_OnChange(frm1.vspdData.text, Row )
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,row,row,C_DocCur,C_ItemAmt,"A" ,"I","X","X")  

	END SELECT

End Sub


'======================================================================================================
'   Event Name : vspdData1_Change
'   Event Desc :
'=======================================================================================================
Sub  vspdData1_Change(ByVal Col, ByVal Row )
	 DIm TmpApAmt,TmpApClsAmt
	
	TmpApAmt= 0
	TmpApClsAmt=0
	
    lgBlnFlgChgValue = True
    Call CheckMinNumSpread(frm1.vspdData1,Col,Row)
    
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.UpdateRow Row
 	WIth frm1.vspdData1   
   
		.Row = Row
		.Col = 0             

		SELECT CASE Col
			    CASE C_ApClsAmt
					.Row = Row
					.Col = C_ApAmt
					TmpApAmt = UNICDbl(.Text)
					.Col = C_ApClsAmt
					TmpApClsAmt = UNICDbl(.Text)
					If (TmpApAmt > 0 And TmpApClsAmt < 0) Or (TmpApAmt < 0 And TmpApClsAmt > 0) then
						.Col = C_DocCur1
						TmpApClsAmt= UNIConvNumPCToCompanyByCurrency(TmpApClsAmt * (-1),.text,parent.ggAmtOfMoneyNo, "X", "X")
						.Col = C_ApClsAmt
						.text= TmpApClsAmt
					End If
					
					Call FixDecimalPlaceByCurrency(frm1.vspdData1,Row,C_DocCur1,C_APCLSAMT,  "A" ,"X","X")
					Call DoMulti(Row)	' ApClsLocAmt 
					Call DoSum()
				CASE C_APDCAMT
					
					Call FixDecimalPlaceByCurrency(frm1.vspdData1,Row,C_DocCur1,C_APDCAMT,  "A" ,"X","X")
					Call DoMulti(Row)	' ApClsLocAmt 
					Call DoSum()
					.Row = Row
					.Col = C_APDCAMT
					IF ABS(UNICDbl(.Text)) <> 0  THEN
						ggoSpread.SpreadUnLock C_ApDcTypePopup,    Row, C_ApDcTypePopup, Row
						ggoSpread.SpreadUnLock C_APDcTypeCd,       Row, C_APDcTypeCd, Row
						ggoSpread.SSSetRequired C_APDcTypeCd,       Row, Row	
					ElSE
						.TEXT="0"
						.Col = C_APDcTypeCd
						.text = ""
						ggoSpread.SpreadLock C_APDcTypePopup,       Row, C_APDcTypePopup, Row
						ggoSpread.SpreadLock C_APDcTypeCd,			Row, C_APDcTypeCd,  Row
						ggoSpread.SSSetProtected C_APDcTypeCd,       Row, Row
					End IF 

		END SELECT
	END with
End Sub


'======================================================================================================
'   Event Name : vspdData4_Change
'   Event Desc :
'=======================================================================================================
Sub  vspdData4_Change(ByVal Col, ByVal Row )
    
    Dim IsWhere, lsSelect, IsFrom
    Dim RetFlag
    
    lgBlnFlgChgValue = True
    
    Call CheckMinNumSpread(frm1.vspdData4,Col,Row)
    
    ggoSpread.Source = frm1.vspdData4
    ggoSpread.UpdateRow Row
    
    frm1.vspdData4.Row = Row
    Select Case Col
		Case C_ALLCAMT
			Call FixDecimalPlaceByCurrency(frm1.vspdData4,Row,C_DocCur4,C_ALLCAMT,  "A" ,"X","X")
			Call DoMulti(Row)
			Call DoSum()  
		
		Case C_DocCur4
				frm1.vspdData4.Col  = C_DocCur4
				Call DocCur4_OnChange(frm1.vspdData4.Text,Row)	
				
		Case C_AllCTYPE		
			   frm1.vspdData4.Col = C_AllCTYPE
			   frm1.vspdData4.Row = Row
				   
			   IsFrom = " (Select  Distinct A.MINOR_CD,A.MINOR_NM,C.REFERENCE " & _
						"From   B_MINOR A,B_CONFIGURATION B  ,B_CONFIGURATION C " & _
						"Where  A.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND A.MAJOR_CD = B.MAJOR_CD AND A.MINOR_CD = B.MINOR_CD " & _
						"AND ( (B.SEQ_NO = 3 AND B.REFERENCE = " & FilterVar("PP", "''", "S") & "  )  OR ( B.SEQ_NO = 6 AND B.REFERENCE IN (" & FilterVar("CP", "''", "S") & " ," & FilterVar("D3", "''", "S") & " ," & FilterVar("PP", "''", "S") & " ) )) " & _
						"AND  A.MAJOR_CD = C.MAJOR_CD AND A.MINOR_CD =  C.MINOR_CD AND C.SEQ_NO = 6) tmp "  
			   IsWhere = "tmp.MINOR_CD =  " & FilterVar(frm1.vspdData4.Text , "''", "S") & " "			   			
			   lsSelect = "tmp.MINOR_CD,tmp.MINOR_NM "
			   
			   Call CommonQueryRs( lsSelect, IsFrom , IsWhere, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 	   	   
			   
   				frm1.vspdData4.Col = C_ALLC_NM

			   if (Trim(lgF0) = "X") or (Trim(lgF0) = "") then
			      RetFlag = DisplayMsgBox("140937","X" , "X", "X") 	
			     '// MsgBox "해당하는 입금유형이 존재하지 않습니다."
			      Exit Sub
			   Elseif Trim(frm1.vspdData4.Text)<> Trim(Left(lgF1, Len(lgF1)-1)) then
					frm1.vspdData4.Text =Trim(Left(lgF1, Len(lgF1)-1))
			   End if
				
				frm1.vspdData4.ReDraw = false
				ggoSpread.source = frm1.vspdData4    
				frm1.vspdData4.Col = C_AllCTYPE   

			   SELECT Case UCase(Trim(Left(lgF0, Len(lgF0)-1)))
				   
					Case "CP"'지불카드 
					    ggoSpread.SSSetProtected	C_ALLCAMT,      Row, Row
					    ggoSpread.SpreadUnLock		C_ETCNO,		Row, C_ETCNO, Row
						ggoSpread.SSSetRequired		C_ETCNO,		Row, Row	
						ggoSpread.SpreadUnLock		C_ETCNO_POP,	Row, C_ETCNO_POP, Row	
						ggoSpread.SSSetProtected	C_DocCur4,		Row, Row
						ggoSpread.SpreadLock		C_DocCurPopup4,	Row, C_DocCurPopup4, Row	
						ggoSpread.SSSetProtected	C_ExchRate4,	Row, Row   
						ggoSpread.SSSetRequired		C_AllcAcct,		Row,	Row	
						ggoSpread.SpreadUnLock		C_AllcAcct_POP,	Row,	C_AllcAcct_POP, Row
					Case "NP"' 지급어음 
						ggoSpread.SSSetProtected	C_ALLCAMT,      Row, Row
						ggoSpread.SpreadUnLock		C_ETCNO,		Row, C_ETCNO, Row
						ggoSpread.SSSetRequired		C_ETCNO,		Row, Row	
						ggoSpread.SpreadUnLock		C_ETCNO_POP,    Row, C_ETCNO_POP, Row
						ggoSpread.SSSetProtected	C_DocCur4,		Row, Row
						ggoSpread.SpreadLock		C_DocCurPopup4,	Row, C_DocCurPopup4, Row	
						ggoSpread.SSSetProtected	C_ExchRate4,	Row, Row   	
						ggoSpread.SSSetRequired		C_AllcAcct,		Row,	Row	
						ggoSpread.SpreadUnLock		C_AllcAcct_POP,	Row,	C_AllcAcct_POP, Row
					Case "CS" , "CK"  '현금 
						ggoSpread.SSSetProtected	C_ETCNO,		Row,	Row	
						ggoSpread.SSSetProtected	C_ETCNO_POP,	Row,	Row
						ggoSpread.SpreadUnLock		C_AllcAcct,		Row,	C_AllcAcct, Row
						ggoSpread.SSSetRequired		C_AllcAcct,		Row,	Row	
						ggoSpread.SpreadUnLock		C_AllcAcct_POP,	Row,	C_AllcAcct_POP, Row
						ggoSpread.SpreadUnLock		C_DocCur4,		Row,	C_DocCur4,		Row	
						ggoSpread.SpreadUnLock		C_DocCurPopup4,	Row,	C_DocCurPopup4,  Row	
						ggoSpread.SSSetRequired		C_DocCur4,		Row,	Row	
						ggoSpread.SpreadLock		C_ExchRate4,	Row,	C_ExchRate4,	Row  
						ggoSpread.SpreadUnLock		C_ALLCAMT,		Row,	C_ALLCAMT,  Row
						ggoSpread.SSSetRequired		C_ALLCAMT,      Row, Row
						 
					Case "DP" '예적금 
						ggoSpread.SpreadUnLock		C_ETCNO,		Row, C_ETCNO, Row
						ggoSpread.SSSetRequired		C_ETCNO,		Row, Row	
						ggoSpread.SpreadUnLock		C_ETCNO_POP,	Row, C_ETCNO_POP,	Row	
						ggoSpread.SSSetProtected	C_DocCur4,		Row, Row
						ggoSpread.SpreadLock		C_DocCurPopup4,	Row, C_DocCurPopup4, Row	
						ggoSpread.SSSetProtected	C_ExchRate4,	 Row, Row   
						ggoSpread.SSSetRequired		C_AllcAcct,		Row,	Row	
						ggoSpread.SpreadUnLock		C_AllcAcct_POP,	Row,	C_AllcAcct_POP, Row
						ggoSpread.SpreadUnLock		C_ALLCAMT,		Row,	C_ALLCAMT,  Row
						ggoSpread.SSSetRequired		C_ALLCAMT,      Row, Row
					Case Else 	' 선급 
						ggoSpread.SpreadUnLock		C_ETCNO,		Row, C_ETCNO, Row
						ggoSpread.SSSetRequired		C_ETCNO,		Row, Row	
						ggoSpread.SpreadUnLock		C_ETCNO_POP,	Row, C_ETCNO_POP,	Row	
						ggoSpread.SpreadUnLock		C_ALLCAMT,		Row,	C_ALLCAMT,  Row
						ggoSpread.SSSetRequired		C_ALLCAMT,      Row, Row
						ggoSpread.SSSetProtected	C_AllcAcct,		Row, Row	    
						ggoSpread.SpreadLock		C_AllcAcct_POP,	Row, C_AllcAcct_POP, Row
						ggoSpread.SSSetProtected	C_DocCur4,		Row, Row
						ggoSpread.SpreadLock		C_DocCurPopup4,	Row, C_DocCurPopup4, Row	
						ggoSpread.SSSetProtected	C_ExchRate4,	Row, Row   
		
				END SELECT	
	
				frm1.vspdData4.ReDraw = True
				Call ClearRow(Row)
				Call DoSum()
	
	End Select
End Sub

'======================================================================================================
'   Event Name :vspdData_EditMode
'   Event Desc :
'=======================================================================================================
Sub vspdData_EditMode(ByVal Col, ByVal Row, ByVal Mode, ByVal ChangeMade)
    
    Select Case Col
        Case C_ItemAmt
            Call EditModeCheck(frm1.vspdData, Row,C_DocCur,C_ItemAmt, "A" ,"I", Mode, "X", "X")
    End Select
End Sub

'======================================================================================================
'   Event Name :vspdData1_EditMode
'   Event Desc :
'=======================================================================================================
Sub vspdData1_EditMode(ByVal Col, ByVal Row, ByVal Mode, ByVal ChangeMade)
    Select Case Col
        Case C_ApClsAmt
            Call EditModeCheck(frm1.vspdData1, Row,C_DocCur1,C_APCLSAMT, "A" ,"I", Mode, "X", "X")
    End Select
End Sub

'======================================================================================================
'   Event Name :vspdData4_EditMode
'   Event Desc :
'=======================================================================================================
Sub vspdData4_EditMode(ByVal Col, ByVal Row, ByVal Mode, ByVal ChangeMade)
    Select Case Col
        Case C_ALLCAMT
            Call EditModeCheck(frm1.vspdData4, Row,C_DocCur4,C_ALLCAMT, "A" ,"I", Mode, "X", "X")
    End Select
End Sub
'======================================================================================================
'   Event Name :vspdData1_DblClick
'   Event Desc :
'=======================================================================================================
Sub  vspdData_DblClick( ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    
End Sub

'======================================================================================================
'   Event Name :vspdData1_KeyPress
'   Event Desc :
'=======================================================================================================
Sub  vspdData1_KeyPress(KeyAscii)
     
End Sub

'======================================================================================================
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub  vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim intListGrvCnt 
    Dim LngLastRow    
    Dim LngMaxRow     
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
End Sub


'======================================================================================================
'   Event Name : txtDeptCd_Change()
'   Event Desc : 
'=======================================================================================================
Sub txtDeptCD_OnChange()
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

    lgBlnFlgChgValue = True

	If TRim(frm1.txtDeptCd.value) <>"" Then
		'----------------------------------------------------------------------------------------
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtPaymDT.Text, gDateFormat,""), "''", "S") & "))"			
		
	
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
	
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
			frm1.txtDeptCd.focus
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
			Next	
		End If
	End if		
		'----------------------------------------------------------------------------------------
End Sub


'======================================================================================================
'   Event Name : txtBpCD_OnChange()
'   Event Desc : 
'=======================================================================================================
Sub txtBpCD_OnChange()
    Dim  RetFlag
    
    lgBlnFlgChgValue = True
     Call CommonQueryRs("Bp_cd, Bp_nm", "B_BIZ_PARTNER", _
                        "bp_cd =  " & FilterVar(frm1.txtBpCd.Value, "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
     if (lgF0 = "X") or (lgF0 = "") then 
         RetFlag = DisplayMsgBox("970000","X" , frm1.txtBpCd.Alt, "X")
        '//msgbox "거래처코드가 존재하지 않습니다"
        Exit Sub
     End if
     frm1.txtBpNm.value   = Left(lgF1, Len(lgF1)-1)

End Sub

 '#########################################################################################################
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'######################################################################################################### 


 '#########################################################################################################
'												5. Interface부 
'	기능: Interface
'	설명: 각각의 Toolbar에 대한 처리를 행한다. 
'	      Toolbar의 위치순서대로 기술하는 것으로 한다. 
'	<< 공통변수 정의 부분 >>
' 	공통변수 : Global Variables는 아니지만 각각의 Sub나 Function에서 자주 사용하는 변수로 변수명은 
'				통일하도록 한다.
' 	1. 공통컨트롤을 Call하는 변수 
'    	   ADF (ADS, ADC, ADF는 그대로 사용)
'    	   - ADF는 Set하고 사용한 뒤 바로 Nothing 하도록 한다.
' 	2. 공통컨트롤에서 Return된 값을 받는 변수 
'    		strRetMsg
'######################################################################################################### 
 '*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'********************************************************************************************************* 

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function  FncQuery() 

    Dim IntRetCD 
    Dim Spr1, Spr2, Spr3, Spr4
    
    FncQuery = False
    Err.Clear
    
    ggoSpread.Source = frm1.vspdData1	:    Spr1 = ggoSpread.SSCheckChange
	ggoSpread.Source = frm1.vspdData	:    Spr2 = ggoSpread.SSCheckChange    
    ggoSpread.Source = frm1.vspdData3	:    Spr3 = ggoSpread.SSCheckChange    
    
    If lgBlnFlgChgValue = True Or Spr1 = True Or Spr2 = True Or Spr3 = True Then
       if DisplayMsgBox("900013", parent.VB_YES_NO,"x","x") = vbNO Then	   
			Exit Function
	   End If
    End If        
    															
    If Not chkField(Document, "1") Then
       Exit Function
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	ggoSpread.Source = frm1.vspdData	:	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData1	:	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2	:	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3	:	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData4	:	ggoSpread.ClearSpreadData

    Call InitVariables
    Call ggoOper.ClearField(Document, "2")
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    
    If DbQuery = False Then
        Exit Function
    End If
              
    FncQuery = True              
 
	Set gActiveElement = document.ActiveElement	   
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function  FncNew()              
           
    Dim Spr1, Spr2, Spr3, Spr4
    
    FncNew = False
    
    ggoSpread.Source = frm1.vspdData1	:    Spr1 = ggoSpread.SSCheckChange
	ggoSpread.Source = frm1.vspdData	:    Spr2 = ggoSpread.SSCheckChange    
    ggoSpread.Source = frm1.vspdData3	:    Spr3 = ggoSpread.SSCheckChange    
    
    If lgBlnFlgChgValue = True Or Spr1 = True Or Spr2 = True Or Spr3 = True then
       if DisplayMsgBox("900015", parent.VB_YES_NO,"X","X") = vbNO Then	   
			Exit Function
	   End If
    End If    
    
    Call SetToolbar("1111111100001111") 
    Call ggoOper.ClearField(Document, "1")  
    Call ggoOper.ClearField(Document, "2")
    Call ggoOper.LockField(Document, "N")
        
    ggoSpread.Source = frm1.vspdData	:	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData1	:	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2	:	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3	:	ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData4	:	ggoSpread.ClearSpreadData
    
    Call InitVariables    
    Call ClickTab1()   
    Call SetDefaultVal
        
    FncNew = True
  
	Set gActiveElement = document.ActiveElement	 
End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function  FncDelete()
    FncDelete = False
    
    Dim liRet
        
    Err.Clear
    On Error Resume Next
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then
       liRet = DisplayMsgBox("900002","X","X","X")                                       
       Exit Function
    End If
    
	 If DisplayMsgBox("900003", parent.VB_YES_NO,"X","X") = vbNo Then
	 	 Exit Function
	 End If
    
    If DbDelete = False Then
       Exit Function
    End If       
        
    FncDelete = True

	Set gActiveElement = document.ActiveElement	    
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function  FncSave() 	                
	
    Dim Spr1, Spr2, Spr3, Spr4
	Dim RetFlag
	
	FncSave = False                                                         
    
    Err.Clear 
    
    ggoSpread.Source = frm1.vspdData1	:    Spr1 = ggoSpread.SSCheckChange
	ggoSpread.Source = frm1.vspdData	:    Spr2 = ggoSpread.SSCheckChange    
    ggoSpread.Source = frm1.vspdData3	:    Spr3 = ggoSpread.SSCheckChange   
    ggoSpread.Source = frm1.vspdData4	:    Spr4 = ggoSpread.SSCheckChange   
    
    If lgBlnFlgChgValue = False  Then
       DisplayMsgBox "900001","X","X","X"   
		 Exit Function
	Elseif Spr1 = False And Spr2 = False And Spr3 = False And Spr4 = False And lgIntFlgMode = parent.OPMD_CMODE then
	      DisplayMsgBox "990001","X","X","X"   
		 Exit Function
    End If 
    
	If Not chkField(Document, "2") Then
       Exit Function
    End If    
    
	ggoSpread.Source = frm1.vspdData1
    If Not ggoSpread.SSDefaultCheck Then
		Call ClickTab2()													'⊙: Check contents area
		Exit Function
    End If
	
		ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then
		Call ClickTab3()													'⊙: Check contents area
		Exit Function
    End If
    
	ggoSpread.Source = frm1.vspdData2
    If Not ggoSpread.SSDefaultCheck Then
		Call ClickTab3()													'⊙: Check contents area
		Exit Function
    End If
    
	ggoSpread.Source = frm1.vspdData4
    If Not ggoSpread.SSDefaultCheck Then
		Call ClickTab1()													'⊙: Check contents area
		Exit Function
    End If       
   
    If UniCDbl(frm1.txtTotLocAmt.text) <> 0 then
       RetFlag = DisplayMsgBox("113119","X" , "X", "X") 	
       '//Msgbox " 반제금액이 0일수 없습니다. "
       Exit Function
    End if
     
        
    IF Not CheckDBEtcNo  then 
        Exit Function
    End if
        
    Call DbSave				                                             '☜: Save db data
       
    FncSave = True                                                                '⊙:                                        
	
	Set gActiveElement = document.ActiveElement	
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function  FncCopy() 

	Select Case gSelframeFlg
   
       Case TAB1       
            if frm1.vspdData4.MaxRows < 1 Then Exit Function
            frm1.vspdData4.ReDraw = False
            ggoSpread.Source = frm1.vspdData4	
            ggoSpread.CopyRow        
	         frm1.vspdData4.ReDraw = True
       Case TAB2
            if frm1.vspdData1.MaxRows < 1 Then Exit Function
            frm1.vspdData1.ReDraw = False
            ggoSpread.Source = frm1.vspdData1	
            ggoSpread.CopyRow        
	         frm1.vspdData1.ReDraw = True
       Case TAB3
            if frm1.vspdData.MaxRows < 1 Then Exit Function
            frm1.vspdData.ReDraw = False
            ggoSpread.Source = frm1.vspdData	
            ggoSpread.CopyRow        
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow,C_DocCur,C_ItemAmt,   "A" ,"I","X","X")

			Call SetSpreadColor(frm1.vspdData.ActiveRow,  frm1.vspdData.ActiveRow)
			Call MaxSpreadVal(frm1.vspdData, C_ItemSeq, frm1.vspdData.ActiveRow)
			Call vspdData_Change(C_AcctCd, frm1.vspddata.activerow)
			
			ggoSpread.Source = frm1.vspdData
			frm1.vspddata.Col = C_DocCur
			frm1.vspddata.row= frm1.vspdData.ActiveRow
			if Trim(Frm1.vspdData.Text) = parent.gCurrency then
				ggoSpread.SpreadLock C_ExchRate		, frm1.vspdData.ActiveRow    , C_ExchRate, frm1.vspdData.ActiveRow
			ELse
				ggoSpread.SpreadunLock C_ExchRate		, frm1.vspdData.ActiveRow    , C_ExchRate, frm1.vspdData.ActiveRow
			END if
	        
	         
	         frm1.vspdData.ReDraw = True
	        
   End Select
   	Call Dosum()
   	Set gActiveElement = document.ActiveElement	
	
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function  FncCancel() 
		
   Select Case gSelframeFlg
   
       Case TAB1
       		With frm1.vspdData4 
				if .Maxrows < 1 Then Exit Function
				.Row = .ActiveRow
				.Col = 0            
				ggoSpread.Source = frm1.vspdData4
				ggoSpread.EditUndo
			
				Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData4,.ActiveRow,.ActiveRow,C_DocCur4,C_BALAMT,"A" ,"I","X","X") 			
				Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData4,.ActiveRow,.ActiveRow,C_DocCur4,C_ALLCAMT,"A" ,"I","X","X")
				
		    
				Call DoSum()
				if .Maxrows < 1 Then Exit Function			
				.Row = .ActiveRow
				.Col = 1                        
			ENd With
       
       Case TAB2
			With frm1.vspdData1 
				if .Maxrows < 1 Then Exit Function
				.Row = .ActiveRow
				.Col = 0            
				ggoSpread.Source = frm1.vspdData1
				ggoSpread.EditUndo
		    
				Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData1,.ActiveRow,.ActiveRow,C_DocCur1,C_APBALAMT,   "A" ,"I","X","X")
				Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData1,.ActiveRow,.ActiveRow,C_DocCur1,C_APCLSAMT,   "A" ,"I","X","X")
				Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData1,.ActiveRow,.ActiveRow,C_DocCur1,C_APAMT,		"A" ,"I","X","X")
				Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData1,.ActiveRow,.ActiveRow,C_DocCur1,C_APDCAMT,		"A" ,"I","X","X")
		    
				Call DoSum()
				if .Maxrows < 1 Then Exit Function			
				.Row = .ActiveRow
				.Col = 1               
			End With
      
       Case TAB3
			With frm1.vspdData
				if .Maxrows < 1 Then Exit Function
				.Row = .ActiveRow
				.Col = 0      
				If  .Text = ggoSpread.InsertFlag Then
					 .Col = C_AcctCd
					If Len(Trim( .text)) > 0 Then  
						 .Col = C_ItemSeq          
						Call DeleteHSheet( .Text)
					End If
				End If      
				ggoSpread.Source = frm1.vspdData
				ggoSpread.EditUndo
				Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,.ActiveRow,.ActiveRow,C_DocCur,C_ItemAmt,"A" ,"I","X","X")    

				Call DoSum()
				if .Maxrows < 1 Then Exit Function			
            
				If .Text = ggoSpread.InsertFlag Then
					.Col = C_AcctCd
					If len(Trim(.text)) > 0 Then 
						.Col = C_ItemSeq
						frm1.hItemSeq.value = .Text
						ggoSpread.Source = frm1.vspdData2        				
						ggoSpread.ClearSpreadData						
						Call DbQuery3(.ActiveRow)
						Call SetSpread2Color()
					End IF 
				Else
				    .Col = C_ItemSeq
				    frm1.hItemSeq.value = .Text
					ggoSpread.Source = frm1.vspdData2        				
					ggoSpread.ClearSpreadData						
				    Call DbQuery2(.ActiveRow)
				End if
			End With
                        
   End Select
   	Call InitData()
   
	Set gActiveElement = document.ActiveElement	
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function  FncInsertRow(ByVal pvRowCnt) 
	Dim liTemp
	Dim imRow
	Dim ii
	Dim iCurRowPos
	
	FncInsertRow = False	
	
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)
	Else
	    imRow = AskSpdSheetAddRowCount()
    
		If imRow = "" Then
		    Exit Function
		End If
	End If		 

		 Select Case gSelframeFlg
				Case TAB1  
					With frm1.vspdData4   
						iCurRowPos = .ActiveRow	
						.ReDraw = False		    
						ggoSpread.Source = frm1.vspdData4
				
						ggoSpread.InsertRow ,imRow

						For ii = .ActiveRow To  .ActiveRow + imRow - 1
							Call MaxSpreadVal(frm1.vspdData4, C_ALLCSEQ , ii)
						Next        

						.Col = 1																	' 컬럼의 절대 위치로 이동 
						.Row = ii - 1
						.Action = 0		
						.ReDraw = True

						Call SetSpreadColor(iCurRowPos + 1, iCurRowPos + imRow)   

					    Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData4,iCurRowPos + 1,iCurRowPos + imRow,"X",C_ALLCAMT,"A" ,"I","X","X")

					End With   
				Case TAB3     
					With frm1.vspdData         
						iCurRowPos = .ActiveRow	
						.ReDraw = False		    
						ggoSpread.Source = frm1.vspdData
						ggoSpread.InsertRow ,imRow

						For ii = .ActiveRow To  .ActiveRow + imRow - 1
							Call MaxSpreadVal(frm1.vspdData, C_ItemSeq , ii)
							.col = C_DocCur
							.text	= parent.gCurrency
							.col = C_ExchRate
							.text	= "1"
							ggoSpread.SpreadLock C_ExchRate		, ii    , C_ExchRate, ii  
							
						Next        

						.Col = 1																	' 컬럼의 절대 위치로 이동 
						.Row = ii - 1
						.Action = 0		
						.ReDraw = True

						Call SetSpreadColor(iCurRowPos + 1, iCurRowPos + imRow)        
						Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,iCurRowPos + 1,iCurRowPos + imRow,"X",C_ItemAmt,"A" ,"I","X","X")						  
						If .MaxRows = 1 then	         
						   .Row = .ActiveRow
						   .Col = 1
						   .Text = .MaxRows
						Else
						  .Row = .ActiveRow-1
						   .Col = 1
						   liTemp = CINT(.Text)+1
						   .Row = .ActiveRow
						   .Col = 1
						   .Text = liTemp
						End if   
				End With 						
		End Select

  	
    If Err.number = 0 Then
		FncInsertRow = True                                                          '☜: Processing is OK
    End If  		

   Set gActiveElement = document.ActiveElement	 		
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function  FncDeleteRow() 

  Select Case gSelframeFlg
   
       Case TAB1               
            frm1.vspdData4.ReDraw = False
            ggoSpread.Source = frm1.vspdData4	
            ggoSpread.DeleteRow    		
	        frm1.vspdData4.ReDraw = True	        
	         
       Case TAB2            
            frm1.vspdData1.ReDraw = False
            ggoSpread.Source = frm1.vspdData1	
            ggoSpread.DeleteRow   		
	        frm1.vspdData1.ReDraw = True
	         
       Case TAB3            
            frm1.vspdData.ReDraw = False
            ggoSpread.Source = frm1.vspdData
		    ggoSpread.DeleteRow    		
	        frm1.vspdData.ReDraw = True
  End Select
  
  Call DoSum()
  
  Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function  FncPrint() 
    On Error Resume Next                                                                 '☜: If process fails
    Err.Clear                                                                            '☜: Clear error status

    FncPrint = False                                                                     '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                               '☜: Protect system from crashing

    If Err.number = 0 Then
       FncPrint = True                                                                   '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function  FncPrev() 
    On Error Resume Next                                                                 '☜: If process fails
    Err.Clear                                                                            '☜: Clear error status

    FncPrev = False                                                                      '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    If Err.number = 0 Then
       FncPrev = True                                                                    '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement                                              
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function  FncNext() 
    On Error Resume Next                                                                 '☜: If process fails
    Err.Clear                                                                            '☜: Clear error status

    FncNext = False                                                                      '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    If Err.number = 0 Then
       FncNext = True                                                                    '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement                                              
End Function

'=======================================================================================================
' Function Name : FncFind
' Function Desc : 화면 속성, Tab유무 
'========================================================================================================
Function  FncFind() 
    On Error Resume Next                                                                 '☜: If process fails
    Err.Clear                                                                            '☜: Clear error status

    FncFind = False                                                                      '☜: Processing is NG

	Call Parent.FncFind(Parent.C_MULTI, True)

    If Err.number = 0 Then
       FncFind = True                                                                    '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement                        
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function  FncExcel() 
    On Error Resume Next                                                                 '☜: If process fails
    Err.Clear                                                                            '☜: Clear error status

    FncExcel = False                                                                     '☜: Processing is NG

	Call Parent.FncExport(Parent.C_MULTI)

    If Err.number = 0 Then
       FncExcel = True                                                                   '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement
End Function
'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()

    Dim ACol, ARow, iColumnLimit
	Dim iRet
	
    iColumnLimit = 5
    
    Select Case gSelframeFlg
   
       Case TAB1         
             
			ACol = Frm1.vspddata4.ActiveCol
			ARow = Frm1.vspddata4.ActiveRow

			If ACol > iColumnLimit Then
					Frm1.vspddata4.Col = iColumnLimit : frm1.vspddata4.Row = 0
					iRet = DisplayMsgBox("900030", "X", Trim(frm1.vspddata4.Text), "X")
					Exit Function
			End If

			Frm1.vspddata4.ScrollBars = parent.SS_SCROLLBAR_NONE

			ggoSpread.Source = Frm1.vspddata4
			ggoSpread.SSSetSplit(ACol)

			Frm1.vspddata4.Col = ACol
			Frm1.vspddata4.Row = ARow

			Frm1.vspddata4.Action = Parent.SS_ACTION_ACTIVE_CELL
			Frm1.vspddata4.ScrollBars = parent.SS_SCROLLBAR_BOTH
       
       Case TAB2            
			ACol = Frm1.vspddata1.ActiveCol
			ARow = Frm1.vspddata1.ActiveRow

			If ACol > iColumnLimit Then
					Frm1.vspddata1.Col = iColumnLimit : frm1.vspddata1.Row = 0
					iRet = DisplayMsgBox("900030", "X", Trim(frm1.vspddata1.Text), "X")
					Exit Function
			End If

			Frm1.vspddata1.ScrollBars = parent.SS_SCROLLBAR_NONE

			ggoSpread.Source = Frm1.vspddata1
			ggoSpread.SSSetSplit(ACol)

			Frm1.vspddata1.Col = ACol
			Frm1.vspddata1.Row = ARow

			Frm1.vspddata1.Action = Parent.SS_ACTION_ACTIVE_CELL
			Frm1.vspddata1.ScrollBars = parent.SS_SCROLLBAR_BOTH            
       Case TAB3            
			ACol = Frm1.vspddata.ActiveCol
			ARow = Frm1.vspddata.ActiveRow

			If ACol > iColumnLimit Then
					Frm1.vspddata.Col = iColumnLimit : frm1.vspddata1.Row = 0  	 	 	 	
					iRet = DisplayMsgBox("900030", "X", Trim(frm1.vspddata.Text), "X")
					Exit Function  
			End If   
    
			Frm1.vspddata.ScrollBars = parent.SS_SCROLLBAR_NONE
    
			ggoSpread.Source = Frm1.vspddata
			ggoSpread.SSSetSplit(ACol)    
    
			Frm1.vspddata.Col = ACol
			Frm1.vspddata.Row = ARow
    
			Frm1.vspddata.Action = Parent.SS_ACTION_ACTIVE_CELL    
			Frm1.vspddata.ScrollBars = parent.SS_SCROLLBAR_BOTH            
   End Select
    
End Function


'=======================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function  FncExit()
	Dim Spr1, Spr2, Spr3, Spr4
	
	FncExit = False
   
   ggoSpread.Source = frm1.vspdData1	:   Spr1 = ggoSpread.SSCheckChange
   ggoSpread.Source = frm1.vspdData		:   Spr2 = ggoSpread.SSCheckChange
   ggoSpread.Source = frm1.vspdData3	:   Spr3 = ggoSpread.SSCheckChange
    
   If lgBlnFlgChgValue = True or Spr1 = True or Spr2 = True or Spr3 = True Then
	   If DisplayMsgBox("900016", parent.VB_YES_NO,"X","X") = vbNo Then		
		  Exit Function
	   End If
	End If
        
   FncExit = True   
   	Set gActiveElement = document.ActiveElement   
End Function

 '*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* 

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery() 

   DbQuery = False
	
	On Error Resume Next
	Err.Clear

	Call LayerShowHide(1)
	
	Dim strVal
		
	With frm1
		    strVal = BIZ_PGM_QUERY_ID & "?txtMode=" & parent.UID_M0001
			strVal = strVal & "&txtPaymNO=" & Trim(.txtPaymNO.value)				'조회 조건 데이타 
	End With
	
	Call RunMyBizASP(MyBizASP, strVal)
	  
	DbQuery = True
	Set gActiveElement = document.ActiveElement  
End Function


'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()
	 Dim Row
	 
     With frm1
     	Call SetSpreadLock("A")
		
		For Row= 0 to .vspdData4.MaxRows
				 ggoSpread.source = frm1.vspdData4    
				.vspdData4.row= Row
				.vspdData4.Col =C_AllCTYPE
				ggoSpread.SpreadLock	C_AllcAcct			 , Row    , C_AllcAcct		, Row  
				ggoSpread.SpreadLock	C_AllcAcct_POP       , Row    , C_AllcAcct_POP , Row  

			If UCase(frm1.vspdData4.text)= "CP" or  UCase(frm1.vspdData4.text)= "NP"  then
				ggoSpread.SSSetProtected C_ALLCAMT,    Row, Row
			END if
		Next 
		
		Call SetSpreadLock("B")
		ggoSpread.source = .vspdData1       
		.vspdData.ReDraw = False         
		ggoSpread.SpreadLock	C_APDcTypeCd       , -1    , C_APDcTypeCd
        ggoSpread.SpreadLock	C_APDcTypePopup       , -1    , C_APDcTypePopup

        '-----------------------
        'Reset variables area
        '-----------------------
        If .vspdData.MaxRows > 0 Then
            .vspdData.Row = 1
            .vspdData.Col = C_ItemSeq
            .hItemSeq.Value = .vspdData.Text 
            Call DbQuery2(1)
            
            gSelframeFlg= TAB3
            Call SetSpreadColor(-1, -1)
            ggoSpread.source = .vspdData       
			.vspdData.ReDraw = False         
			ggoSpread.SpreadLock	C_AcctCd		, -1    , C_AcctCd
		    ggoSpread.SpreadLock	C_AcctPopup       , -1    , C_AcctPopup   
			
			For Row= 0 to .vspdData.MaxRows
				.vspdData.Col  = C_DocCur
				.vspdData.row= Row
				if Trim(.vspdData.Text) = parent.gCurrency then
					ggoSpread.SpreadLock C_ExchRate		, Row    , C_ExchRate, Row
				ELse
					ggoSpread.SpreadunLock C_ExchRate	, Row    , C_ExchRate, Row
				END if
			next
			.vspdData.ReDraw = True
			
        End If
    End With    

    lgIntFlgMode = parent.OPMD_UMODE												'Indicates that current mode is Update mode
	 Call ClickTab1()
     Call InitData()
     
     lgBlnFlgChgValue = False
End Function


'========================================================================================
' Function Name : DbQuery2
' Function Desc : This function is data query and display
'========================================================================================
'========================================================================================
' Function Name : DbQuery2
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery2(ByVal Row)
	Dim strVal	
	Dim lngRows
		
	Dim strSelect
	Dim strFrom
	Dim strWhere 	
	
	Dim strTableid
	Dim strColid
	Dim strColNm	
	Dim strMajorCd	
	Dim strNmwhere
	Dim i,Indx1
	Dim arrVal,arrTemp
	
	Err.Clear
	
	With frm1
	    .vspdData.Row = Row
	    .vspdData.Col = C_ItemSeq
	    .hItemSeq.Value = .vspdData.Text
 
	    If Trim(.hItemSeq.Value) = "" Then
	        Exit Function
	    End If
	   
        If CopyFromData(.hItemSeq.Value) = True Then
			Call SetSpread2Color()
            Exit Function
        End If
    	
		Call LayerShowHide(1)
	
		DbQuery2 = False
		
		.vspdData.Row = Row
		.vspdData.Col = C_ItemSeq
		
		strSelect =				" C.DTL_SEQ,  A.CTRL_CD, A.CTRL_NM , C.CTRL_VAL, '',"
		strSelect = strSelect & " CASE  WHEN A.ColM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  END , D.SEQ, LTrim(ISNULL(A.TBL_ID,'')), LTrim(ISNULL(A.DATA_COLM_ID,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')),  LTrim(ISNULL(A.ColM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
		strSelect = strSelect & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  "
		strSelect = strSelect & " END	, " & .hItemSeq.Value & ",  "
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD,'')),CHAR(8) "
    		
		strFrom = " A_CTRL_ITEM	A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_PAYM_DC_DTL C (NOLOCK), A_PAYM_DC D (NOLOCK) "
		
		strWhere =			  " D.PAYM_NO = " & FilterVar(UCase(.txtPAYMNo.value), "''", "S")
		strWhere = strWhere & " AND D.SEQ = " & .hItemSeq.Value & " "
		strWhere = strWhere & " AND D.PAYM_NO  =  C.PAYM_NO  "
		strWhere = strWhere & " AND D.SEQ  =  C.SEQ "
		strWhere = strWhere & "	AND D.ACCT_CD = B.ACCT_CD "
		strWhere = strWhere & " AND C.CTRL_CD = B.CTRL_CD "
		strWhere = strWhere & " AND D.ACCT_CD = B.ACCT_CD "
		strWhere = strWhere & " AND B.CTRL_CD = A.CTRL_CD "
		strWhere = strWhere & " ORDER BY C.DTL_SEQ "
			
		frm1.vspdData2.ReDraw = False
			
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then   
			
			ggoSpread.Source = frm1.vspdData2			
			arrTemp =  Split(lgF2By2,Chr(12))
			For Indx1 = 0 To Ubound(arrTemp) - 1
				arrTemp(indx1) = Replace(arrTemp(indx1), Chr(8), indx1 + 1)
			Next
			
			lgF2By2 = Join(arrTemp,Chr(12))
			ggoSpread.SSShowData lgF2By2
		
			For lngRows = 1 To frm1.vspdData2.MaxRows
		
				frm1.vspddata2.Row = lngRows	
				frm1.vspdData2.Col = C_Tableid 
				If Trim(frm1.vspddata2.text) <> "" Then
					frm1.vspdData2.Col = C_Tableid
					strTableid = frm1.vspddata2.text
					frm1.vspdData2.Col = C_Colid
					strColid = frm1.vspddata2.text
					frm1.vspdData2.Col = C_ColNm
					strColNm = frm1.vspddata2.text	
					frm1.vspdData2.Col = C_MajorCd					
					strMajorCd = frm1.vspddata2.text	
					
					frm1.vspdData2.Col = C_CtrlVal
					
					strNmwhere = strColid & " =   " & FilterVar(frm1.vspddata2.text, "''", "S") & "  " 
					
					If Trim(strMajorCd) <> "" Then
						strNmwhere = strNmwhere & " AND MAJOR_CD =  " & FilterVar(strMajorCd, "''", "S") & "  "
					End If
					
					If CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   
						frm1.vspdData2.Col = C_CtrlValNm
						arrVal = Split(lgF0, Chr(11))  
						frm1.vspddata2.text = arrVal(0)
					End If
				End If								
				
				strVal = strVal & Chr(11) & .hItemSeq.Value 
				frm1.vspdData2.Col = C_DtlSeq  
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_CtrlCd   
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_CtrlNm   
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_CtrlVal 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_CtrlPB   
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_CtrlValNm 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_Seq 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_Tableid 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_Colid 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_ColNm 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_Datatype 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_DataLen 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_DRFg 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_MajorCd 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_MajorCd+1 				
				.vspdData2.Text = lngRows
				strVal = strVal & Chr(11) & .vspddata2.text
				strVal = strVal & Chr(11) & Chr(12)									
			Next
							
			ggoSpread.Source = frm1.vspdData3
			ggoSpread.SSShowData strVal	
		End If 		
		
		Call SetSpread2Color()
	End With
	
	Call LayerShowHide(0)
	
	frm1.vspdData2.ReDraw = True
	
	DbQuery2 = True
	Set gActiveElement = document.ActiveElement
	
End Function
'=======================================================================================================
' Function Name : DbQueryOk2
' Function Desc : DbQuery2가 성공적일 경우 MyBizASP 에서 호출되는 Function
'=======================================================================================================
Function  DbQueryOk2()

	With frm1
	    Call SetSpread2Color
    End With

End Function

'=======================================================================================================
'   Function Name : FindExchRate2
'   Function Desc : 1.날짜, Row를 입력받아 날짜에 해당하는 환율정보를 읽어온다.
'=======================================================================================================

Function FindExchRate2(Byval strDate, Byval FromCurrency,Byval Row )
	Dim strSelect, strFrom, strWhere
	Dim arrTemp
	Dim strExchFg
	Dim strExchRate
	

	strSelect	= "b.minor_cd"
	strFrom		= "b_company a, b_minor b"
	strWhere	= "b.major_cd = " & FilterVar("a1004", "''", "S") & "  and	a.xch_rate_fg = b.minor_cd"
	If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
		arrTemp = Split(lgF0, chr(11))
		strExchFg =  arrTemp(0)
	End If

	If UCase(strExchFg) <> "D" Then 	' Fixed Exchange Rate
		strDate = Mid(strDate, 1, 6)
		strSelect	= "std_rate"
		strFrom		= "b_monthly_exchange_rate (noLock) "
		strWhere	= "from_currency =  " & FilterVar(FromCurrency , "''", "S") & ""
		strWhere	= strWhere & " And to_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
		strWhere	= strWhere & " And apprl_yrmnth  =  " & FilterVar(strDate , "''", "S") & ""

		If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
			arrTemp = Split(lgF0, chr(11))
			strExchRate =  arrTemp(0)
			frm1.vspdData4.row  = Row
			frm1.vspdData4.Col  = C_ExchRate4
			frm1.vspdData4.Text = UNIConvNumPCToCompanyByCurrency(strExchRate, parent.gCurrency, parent.ggExchRateNo, parent.gLocRndPolicyNo, "X")
		Else
			IntRetCD = DisplayMsgBox("121600", "X", "X", "X")
		End If
	Else					' Floating Exchange Rate

		strSelect	= "top 1 std_rate"
		strFrom		= "b_daily_exchange_rate (noLock) "
		strWhere	= "from_currency =  " & FilterVar(FromCurrency , "''", "S") & ""
		strWhere	= strWhere & " And to_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
		strWhere	= strWhere & " And apprl_dt  <= convert(char(21), " & FilterVar(strDate, "''", "S") & ", 20) order by apprl_dt desc"

		If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
			arrTemp = Split(lgF0, chr(11))
			strExchRate =  arrTemp(0)
			frm1.vspdData4.row  = Row
			frm1.vspdData4.Col  = C_ExchRate4
			frm1.vspdData4.Text = UNIConvNumPCToCompanyByCurrency(strExchRate, parent.gCurrency, parent.ggExchRateNo, parent.gLocRndPolicyNo, "X")
		Else
			IntRetCD = DisplayMsgBox("121500", "X", "X", "X")
		End If
	End If
	
End Function    
'=======================================================================================================
' Function Name : Spread의 값을 Return
' Function Desc : 
'=======================================================================================================
Sub GetSpread1()
	Dim IGrpCnt
    Dim strVal
    Dim IRow
    
    '채권내역 
    
    IGrpCnt = 1
	strVal	= ""  
	
	With frm1.vspdData1
		For IRow = 1 To .MaxRows
			.Row = IRow
			.Col = 0
			  		  
			Select Case .Text
				Case ggoSpread.DeleteFlag                                                                '☜: Update추가 
					
				Case Else
					strVal = strVal & "C" & parent.gColSep & IRow & parent.gColSep   
					.Col = C_APNo  
					strVal = strVal & Trim(.text) & parent.gColSep 
					strval = strval & UNIConvDate(Trim(frm1.txtPaymDt.text)) & parent.gColSep 
					.Col = C_DocCur1
					strVal = strVal & Trim(.text) & parent.gColSep 
					.Col = C_APCLSAMT   
					strVal = strVal & Trim(UNIConvNum(.Text,0))  & parent.gColSep 
					.Col = C_APCLSLOCAMT  
					strVal = strVal & Trim(UNIConvNum(.Text,0))  & parent.gColSep 
					.Col = C_APDESC 
					strVal = strVal & Trim(.text) & parent.gColSep 
					.Col = C_ApDcTypeCd 
					strVal = strVal & Trim(.text) & parent.gColSep
					.Col = C_APDCAMT   
					strVal = strVal & Trim(UNIConvNum(.Text,0))  & parent.gColSep 
					.Col = C_APDCLOCAMT  
					strVal = strVal & Trim(UNIConvNum(.Text,0))  & parent.gRowSep 					
					IGrpCnt = IGrpCnt + 1 
			End Select
		 Next		
	End With

	frm1.txtMaxRows1.value = IGrpCnt-1										'Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread1.value  = strVal
End Sub

'=======================================================================================================
' Function Name : Spread의 값을 Return
' Function Desc : 
'=======================================================================================================
Sub GetSpread4()
    Dim IGrpCnt
    Dim strVal
    Dim IRow
        
    '반제 
    IGrpCnt = 1
	strVal  = ""
	With frm1.vspdData4
		For IRow = 1 To .MaxRows
			.Row = IRow
			.Col = 0

			Select Case .Text
         		Case ggoSpread.DeleteFlag                                                                '☜: Update추가 
					
         		Case Else 
					strVal = strVal & "C" & parent.gColSep & IRow & parent.gColSep 
					.Col = C_AllCTYPE  
					strVal = strVal & UCase(Trim(.text))  & parent.gColSep 
					.Col = C_AllcAcct  
					strVal = strVal & Trim(.text) & parent.gColSep
					.Col = C_ETCNO 
					strVal = strVal & Trim(.text) & parent.gColSep
					.col = C_BPCD
					strVal = strVal & Trim(.text) & parent.gColSep
					.Col = C_DocCur4
					strVal = strVal & Trim(.text) & parent.gColSep
					.Col = C_ExchRate4
					strVal = strVal & Trim(UNIConvNum(.text,0))  & parent.gColSep
					.col = C_BALAMT
					strVal = strVal & Trim(UNIConvNum(.text,0))  & parent.gColSep
					.Col = C_ALLCAMT 
					strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep
					.Col = C_ALLCLOCAMT
					strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep
					.Col = C_DESC 
					strVal = strVal & Trim(.text) & parent.gColSep
					.col = C_BIZCD
					strVal = strVal & Trim(.text) & parent.gColSep
					
					strVal = strVal & Trim(frm1.txtPaymNo.value) & parent.gColSep	
					.Col = C_BANK_CD 
					strVal = strVal & Trim(.text) & parent.gRowSep				

                    IGrpCnt = IGrpCnt + 1
			End Select            
		Next		
	End With
	
	frm1.txtMaxRows4.value = IGrpCnt - 1
	frm1.txtSpread4.value  = strVal   
End Sub

'=======================================================================================================
' Function Name : Spread의 값을 Return
' Function Desc : 
'=======================================================================================================
Sub GetSpread()
    Dim IGrpCnt
    Dim strVal
    Dim IRow
    
    IGrpCnt = 1
	strVal = ""  
	 
	With frm1.vspdData
		For IRow = 1 To .MaxRows
			.Row = IRow
			.Col = 0
					  
			Select Case .Text
				Case ggoSpread.DeleteFlag                                                                '☜: Update추가 
					
         		Case Else 
         			strVal = strVal & "C" & parent.gColSep 
					.Col = C_ItemSeq  
					strVal = strVal & UCase(Trim(.Text)) & parent.gColSep 
					.Col = C_AcctCd  
					strVal = strVal & Trim(.text) & parent.gColSep 
					.Col = C_DrCrFg 
					strVal = strVal & Trim(.text) & parent.gColSep 
					.Col = C_DocCur 
					strVal = strVal & Trim(.text)   & parent.gColSep 
					.Col = C_ExchRate
					strVal = strVal & Trim(UNIConvNum(.Text,0))   & parent.gColSep  
					.Col = C_ItemAmt
					strVal = strVal & Trim(UNIConvNum(.Text,0))   & parent.gColSep   
					.Col = C_ItemLocAmt
					strVal = strVal & Trim(UNIConvNum(.Text,0))   & parent.gColSep                
					.Col = C_ItemDesc 
					strVal = strVal & Trim(.text)  & parent.gRowSep		              
                    IGrpCnt = IGrpCnt + 1
			End Select            
		Next		
	End With

	frm1.txtMaxRows.value = IGrpCnt-1
	frm1.txtSpread.value  = strVal     
End Sub

'=======================================================================================================
' Function Name : Spread의 값을 Return
' Function Desc : 
'=======================================================================================================
Sub GetSpread3()
    Dim IGrpCnt
    Dim strVal
    Dim IRow
        
    '잔액내역:관리항목 
    IGrpCnt = 1
	strVal = ""  
	 
	For IRow = 1 To frm1.vspdData3.MaxRows
        frm1.vspdData3.Row = IRow
		frm1.vspdData3.Col = 0
		  		  
        Select Case frm1.vspdData3.Text
				Case ggoSpread.DeleteFlag                                                                '☜: Update추가 
					
         		Case Else 	
         		strVal = strVal & "C" & parent.gColSep & IRow & parent.gColSep 
		        frm1.vspdData3.Col = 1                
		        strVal = strVal & Trim(frm1.vspdData3.Text) & parent.gColSep 
		        frm1.vspdData3.Col = 2                
		        strVal = strVal & Trim(frm1.vspdData3.Text) & parent.gColSep 
		        frm1.vspdData3.Col = 3                
		        strVal = strVal & Trim(frm1.vspdData3.Text) & parent.gColSep 
		        frm1.vspdData3.Col = 5                
		        strVal = strVal & Trim(frm1.vspdData3.Text) & parent.gRowSep
		        IGrpCnt = IGrpCnt + 1      
		End Select   
	Next		

	frm1.txtMaxRows3.value = IGrpCnt-1
	frm1.txtSpread3.value  = strVal     
End Sub
'==========================================================================================
'   Event Name : DocCur_OnChange
'   Event Desc : 
'==========================================================================================
Sub DocCur_OnChange(byVal strDocCur, byVal Row)

    lgBlnFlgChgValue = True

	If Trim(strDocCur) = parent.gCurrency Then
		frm1.vspdData.Col  = C_ExchRate
		frm1.vspdData.Text = "1"
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.SpreadLock C_ExchRate		, Row    , C_ExchRate, Row

	Else
		Call FindExchRate(UniConvDateToYYYYMMDD(frm1.txtPaymDT.text,parent.gDateFormat,""), UCase(Trim(strDocCur)),Row)
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.SpreadunLock C_ExchRate		, Row    , C_ExchRate, Row

	End IF
	
	Call DoMulti(Row)
	Call Dosum()
End Sub
'==========================================================================================
'   Event Name : DocCur4_OnChange
'   Event Desc : 
'==========================================================================================
Sub DocCur4_OnChange(byVal strDocCur, byVal Row)

    lgBlnFlgChgValue = True

	If Trim(strDocCur) = parent.gCurrency Then
		frm1.vspdData4.Col  = C_ExchRate4
		frm1.vspdData4.Text = "1"
		ggoSpread.Source = frm1.vspdData4	
		ggoSpread.SpreadLock	C_ExchRate4		, Row    , C_ExchRate4, Row

	Else
		Call FindExchRate2(UniConvDateToYYYYMMDD(frm1.txtPaymDT.text,parent.gDateFormat,""), UCase(Trim(strDocCur)),Row)
		ggoSpread.Source = frm1.vspdData4	
		ggoSpread.SpreadunLock C_ExchRate4		, Row    , C_ExchRate4, Row

	End IF
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData4,row,row,C_DocCur4,C_ALLCAMT,"A" ,"I","X","X") 
	Call DoMulti(Row)
	Call Dosum()
End Sub
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave()     
    On Error Resume Next                                                   
	Err.Clear  
	Call LayerShowHide(1)
	
    DbSave = False                                                          '⊙: Processing is NG
 
	lgRetFlag = False
	
	With frm1
	     
	    .txtMode.value = lgIntFlgMode	
		 strMode	   = .txtMode.value
		 		 
		 Call GetSpread1()	'채무내역 
		 Call GetSpread4()	'반제 
		 Call GetSpread()	' 입금잔액 
		 Call GetSpread3()	' 관리항목 
		 		 
		 Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)
	
	End With		
	
   DbSave = True                                 
   
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk(ByVal PAYMNo)		
  
    Call LayerShowHide(0)
    
    If lgIntFlgMode = parent.OPMD_CMODE Then
		frm1.txtPAymNO.value = PAYMNo
	End If	      

    Call ggoOper.ClearField(Document, "2")
    
    ggoSpread.Source = frm1.vspddata1	:	Call ggoSpread.ClearSpreadData()
	ggoSpread.Source = frm1.vspddata	:	Call ggoSpread.ClearSpreadData()
	ggoSpread.Source = frm1.vspdData2	:	Call ggoSpread.ClearSpreadData()
	ggoSpread.Source = frm1.vspdData3	:	Call ggoSpread.ClearSpreadData()
 	ggoSpread.Source = frm1.vspdData4	:	Call ggoSpread.ClearSpreadData()
       
    Call InitVariables        
    'Call ClickTab1()  
	If DbQuery = False Then
		Exit Function
	End If
	lgDBSaveOK = 1
    
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function  DbDelete() 
	
	Dim strVal, lsALLCNO 

	Call LayerShowHide(1)
	on Error Resume Next
	Err.Clear
   
   
	DbDelete = False    
   
	frm1.txtMode.value = parent.UID_M0003	
    strMode = frm1.txtMode.value
   
 
	strVal = BIZ_PGM_DEL_ID & "?txtPaymNO=" & Trim(frm1.txtPaymNO.value)
	strVal = strVal & "&txtMode=" & strMode
   
   
	Call RunMyBizASP(MyBizASP, strVal)
    
   DbDelete = True                            
    Set gActiveElement = document.ActiveElement     
End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()		
	 
	 Call LayerShowHide(0)
	 
	Call ggoOper.ClearField(Document, "1")
    Call ggoOper.ClearField(Document, "2")
    Call ggoOper.LockField (Document, "N")
    
    ggoSpread.Source = frm1.vspdData1	:	Call ggoSpread.ClearSpreadData()
    ggoSpread.Source = frm1.vspdData	:	Call ggoSpread.ClearSpreadData()
    ggoSpread.Source = frm1.vspdData2	:	Call ggoSpread.ClearSpreadData()
    ggoSpread.Source = frm1.vspdData3	:	Call ggoSpread.ClearSpreadData()
    ggoSpread.Source = frm1.vspdData4	:	Call ggoSpread.ClearSpreadData()

    Call InitVariables
    Call ClickTab1()
    Call SetDefaultVal
	
End Function


'-------------------- User Function ------------------------------------------------------
'========================================================================================
' Function Name : ClearRow() 
' Function Desc : ClearRow
'========================================================================================
Sub ClearRow(ByVal Row )
	
	With frm1.vspdData4
		.row= row
		.ReDraw =false
		.Col	=C_ETCNO	:		.text	=""
		.Col	=C_BANK_CD	:		.text	=""
		.Col	=C_BANK_NM	:		.text	=""
		.Col	=C_DocCur4	:		.text	=parent.gCurrency
		.Col	=C_ExchRate4:		.text	="1"
		.Col	=C_BPCD		:		.text	=""
		.Col	=C_BPNm		:		.text	=""
		.Col	=C_BALAMT	:		.text	=""
		.Col	=C_ALLCAMT	:		.text	=""
		.Col	=C_ALLCLOCAMT:		.text	=""
		.Col	=C_AllcAcct	:		.text	=""
		.Col	=C_AllcAcctNm:		.text	=""
		.Col	=C_BIZCD	:		.text	=""
		.Col	=C_BIZNM	:		.text	=""
		.ReDraw =true
	ENd with
ENd sub
'========================================================================================
' Function Name : DoSum() 
' Function Desc : 스프레드의 합을 구해 Display한다.
'========================================================================================
Sub DoSum()
	DIm tmpTotLocAmt
	Dim tmpAllcLocSum, tmpAPClsLocSum, tmpDrLocSum,tmpCrLocSum, tmpAPDcLocSum
	DIm row
	Dim liTemp,liTemp2
	
	tmpAllcLocSum	=0
	tmpAPClsLocSum	=0
	tmpDrLocSum		=0
	tmpCrLocSum		=0
	tmpAPDcLocSum	= 0
	
	With frm1
		Select Case gSelframeFlg
			
			Case TAB1 'AllcLocSum
				tmpAllcLocSum = FncSumSheet1(.vspdData4, C_ALLCLOCAMT ,   1, .vspdData4.MaxRows, false, -1, -1, "V")
				tmpAPDcLocSum = FncSumSheet1(.vspdData1, C_APDcLOCAMT ,   1, .vspdData1.MaxRows, false, -1, -1, "V")
				tmpAllcLocSum = unicdbl(tmpAPDcLocSum) + unicdbl(tmpAllcLocSum) 

				.txtAllcLocAmt.text = UNIConvNumPCToCompanyByCurrency(tmpAllcLocSum,parent.gCurrency,parent.ggAmtOfMoneyNo, "X", "X")			
			
			Case TAB2	'APClsLocSum
				tmpAPClsLocSum = FncSumSheet1(.vspdData1, C_ApCLSLOCAMT ,   1, .vspdData1.MaxRows, false, -1, -1, "V")
				tmpAPDcLocSum = FncSumSheet1(.vspdData1, C_ApDcLOCAMT ,   1, .vspdData1.MaxRows, false, -1, -1, "V")
				tmpAPClsLocSum = unicdbl(tmpAPDcLocSum) + unicdbl(tmpAPClsLocSum) 
				
				.txtAPClsLocAmt.text = UNIConvNumPCToCompanyByCurrency(tmpAPClsLocSum,parent.gCurrency,parent.ggAmtOfMoneyNo, "X", "X")			
						
				tmpAllcLocSum = FncSumSheet1(.vspdData4, C_ALLCLOCAMT ,   1, .vspdData4.MaxRows, false, -1, -1, "V")
				tmpAPDcLocSum = FncSumSheet1(.vspdData1, C_APDcLOCAMT ,   1, .vspdData1.MaxRows, false, -1, -1, "V")
				tmpAllcLocSum = unicdbl(tmpAPDcLocSum) + unicdbl(tmpAllcLocSum) 

				.txtAllcLocAmt.text = UNIConvNumPCToCompanyByCurrency(tmpAllcLocSum,parent.gCurrency,parent.ggAmtOfMoneyNo, "X", "X")			
	
			Case TAB3	 'txtDrLocAmt, txtCrLocAmt
				For row = 1 To .vspdData.maxRows
					.vspdData.Col = 0
					.vspdData.Row = row
				
					If .vspdData.Text <> ggoSpread.DeleteFlag Then
						.vspdData.Col = C_DrCrFg
							If .vspdData.text = "DR" Then
								.vspdData.Col = C_ItemLocAmt
								tmpDrLocSum = cdbl(tmpDrLocSum) + unicdbl(.vspdData.text) 
							Else .vspdData.text = "CR" 
								.vspdData.Col = C_ItemLocAmt
								tmpCrLocSum = cdbl(tmpCrLocSum) + unicdbl(.vspdData.text)
							End If
					End If	
				Next
				
				.txtDrLocAmt.text = UNIConvNumPCToCompanyByCurrency(tmpDrLocSum,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")		
				.txtCrLocAmt.text = UNIConvNumPCToCompanyByCurrency(tmpCrLocSum,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")		

		End Select	
			
			tmpTotLocAmt=  (unicdbl(.txtAllcLocAmt.text) + unicdbl(.txtCrLocAmt.text)) - (unicdbl(.txtAPClsLocAmt.text) + unicdbl(.txtDrLocAmt.text))
			.txtTotLocAmt.text = UNIConvNumPCToCompanyByCurrency(tmpTotLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")		
	End With		
End Sub    

 
'========================================================================================
' Function Name : DoMulti()
' Function Desc : 금액과 환율을 곱해서 자국을 구한다.
'========================================================================================
Sub DoMulti(Byval Row)
	Dim TmpXchrate, TmpLocAmt
	DIm TmpApbalAmt, TmpApDCAmt
	
	TmpXchrate=0
	TmpLocAmt=0
	TmpAPbalAmt= 0
	TmpApDCAmt =0
	Select Case gSelframeFlg
		Case TAB1
			With frm1.vspdData4
				.Col=C_ExchRate4
				.Row=Row
				TmpXchrate= UniCDbl(.text)
				.Col= C_ALLCAMT
				TmpLocAmt= UniCDbl(.text) * TmpXchrate
				TmpLocAmt= UNIConvNumPCToCompanyByCurrency(TmpLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, "X", "X")			
				.Col =C_ALLCLOCAMT 
				.text= TmpLocAmt
			End With
		Case TAB2
			With frm1.vspdData1
				.Row=Row
				.Col=C_ExchRate1
				TmpXchrate= UniCDbl(.text)				
				.Col=C_APBALAMT
				TmpAPbalAmt= UniCDbl(.text)
				.Col= C_APCLSAMT
				if UniCDbl(.text) <> TmpAPbalAmt then	' 반제금액과 잔액과 비교하여 같으면 자국금액 셋팅 
					TmpLocAmt= UniCDbl(.text) * TmpXchrate
					TmpLocAmt= UNIConvNumPCToCompanyByCurrency(TmpLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, "X", "X")			
				Else
					.Col =C_APBALLOCAMT 
					TmpLocAmt= .text
				End if
				.Col =C_APCLSLOCAMT 
				.text= TmpLocAmt				
				
				.Col =C_APDCAMT			'DCLOCAMT
				TmpApDCAmt =UniCDbl(.text)
				.Col =C_APDCLOCAMT
				.text= UNIConvNumPCToCompanyByCurrency(TmpApDCAmt* TmpXchrate,parent.gCurrency,parent.ggAmtOfMoneyNo, "X", "X")			
			End With
		
		Case TAB3
			With frm1.vspdData
				.Col=C_ExchRate
				.Row=Row
				TmpXchrate= UniCDbl(.text)
				.Col= C_ItemAmt
				TmpLocAmt= UniCDbl(.text) * TmpXchrate
				TmpLocAmt= UNIConvNumPCToCompanyByCurrency(TmpLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, "X", "X")			
				.Col =C_ItemlocAmt 
				.text= TmpLocAmt
			End With
	End Select 
End Sub  

 
'========================================================================================
' Function Name : CheckDBEtcNo()
' Function Desc : execute storedprocedure to create a_gl
'========================================================================================
Function CheckDBEtcNo()
	Dim IntRetCD
	Dim iRow
	
	CheckDBEtcNo = TRUE
	If lgIntFlgMode <> parent.OPMD_UMODE Then

		With frm1.vspdData1
			For iRow= 1 To .MaxRows
				.Row = iRow
				.Col=C_APNO
				Call CommonQueryRs("BAL_AMT","A_OPEN_AP","AP_NO= " & FilterVar(.text, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					.Col= C_APCLSAMT
					IF ABS(UniCDBl(Trim(.text))) > ABS(Cdbl(Trim(lgF0))) then
						.Col=C_APNO
						IntRetCD = DisplayMsgBox("112919", "X",Trim(.text), "X")	'반제금액이 잔액을 초과할수 없습니다.
				
						CheckDBEtcNo = False
						Exit Function	
					End If
			Next
		ENd With
			
	
		With frm1.vspdData4
			For iRow= 1 To .MaxRows
				.Row = iRow
				.Col=C_AllCTYPE
				
				Select Case UCase(Trim(.text))
					Case "PP" '선급금 

						.Col=C_ETCNO 
						IF CommonQueryRs("BAL_AMT","F_PRPAYM","PRPAYM_NO= " & FilterVar(.text, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
							IntRetCD = DisplayMsgBox("141500", "X", "X", "X")	
							CheckDBEtcNo = False
							Exit Function	
						Else
							.Col= C_ALLCAMT
							IF UniCDBl(Trim(.text)) > Cdbl(Trim(lgF0)) then
								.Col=C_ETCNO 
								IntRetCD = DisplayMsgBox("112919", "X", Trim(.text), "X")	'반제금액이 잔액을 초과할수 없습니다.
								CheckDBEtcNo = False
								Exit Function	
							End If
						End if						


					Case "CP", "NP" '수취구매카드, 어음일때 
							.Col=C_ETCNO 
							IF CommonQueryRs("NOTE_STS","F_NOTE","NOTE_NO= " & FilterVar(.text, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
								IntRetCD = DisplayMsgBox("171910", "X", Trim(.text), "X")	
								CheckDBEtcNo = False
								Exit Function	
							Else
								lgF0 = Split(lgF0, Chr(11)) 
								IF UCase((Trim(lgF0(0)))) <> "BG"then
									IntRetCD = DisplayMsgBox("112919", "X",Trim(.text), "X")	'반제금액이 잔액을 초과할수 없습니다.
									CheckDBEtcNo = False
									Exit Function	
								End If
							ENd if
				End Select
				
			Next
		End With
  	ENd if	

End Function




'=======================================================================================================
'   Event Name : txtPaymDT_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtPaymDT_DblClick(Button)
    If Button = 1 Then
        frm1.txtPaymDT.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtPaymDT.Focus
	End If
End Sub
'=======================================================================================================
'   Event Name : txtPaymDT_onblur()
'   Event Desc : 
'=======================================================================================================
Function txtPaymDT_onblur()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2
 
	lgBlnFlgChgValue = True
	With frm1
	
		If LTrim(RTrim(.txtDeptCd.value)) <> "" and Trim(.txtPaymDT.Text <> "") Then
			'----------------------------------------------------------------------------------------
				strSelect	=			 " Distinct org_change_id "    		
				strFrom		=			 " b_acct_dept(NOLOCK) "		
				strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(.txtDeptCd.value)), "''", "S") 
				strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
				strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
				strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(.txtPaymDT.Text, gDateFormat,""), "''", "S") & "))"			
	
			IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					
			If IntRetCD = False  OR Trim(Replace(lgF0,Chr(11),"")) <> Trim(.hOrgChangeId.value) Then
					.txtDeptCd.value = ""
					.txtDeptNm.value = ""
					.hOrgChangeId.value = ""
					.txtDeptCd.focus
			End if
		End If
	End With
'----------------------------------------------------------------------------------------

End Function

'=======================================================================================================
'   Event Name : deptCheck()
'   Event Desc : 
'=======================================================================================================

Function deptCheck()

	Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim IntRetCD 
	
	With Frm1
		If  .txtDeptCd.value = "" Then
			.txtDeptCd.value = ""
			.txtDeptNm.value=""
			.hCostCd.value=""
			.hInternalCD.value=""
			.hbizcd.value = ""
			.horgChangeId.value = ""
			lgBlnFlgChgValue = True 
			Exit Function
		End If
    
    
		If Trim(.txtPaymDT.Text = "") Then    
			Exit Function
		End If
		

		'----------------------------------------------------------------------------------------
		strSelect	=			 " distinct org_change_id "    		
		strFrom		=			 " b_acct_dept "		
		strWhere	=			 " org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(.txtPaymDT.Text, parent.gDateFormat,""), "''", "S") & "))"			
			
	
		IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 			
	
		If IntRetCD = False or Trim(Replace(lgF0,Chr(11),"")) <> .horgChangeId.value  Then
			.txtDeptCd.value = ""
			.txtDeptNm.value=""
			.hCostCd.value=""
			.hInternalCD.value=""
			.hbizcd.value = ""
			.horgChangeId.value = ""
			lgBlnFlgChgValue = True
			Exit Function
		End If	
	End With
	
End Function
'=======================================================================================================
' Area Name   : 4.5.5 Spread Popup method 
' Description : This part declares spread popup method
'=======================================================================================================
'*******************************************************************************************************
'===================================== PopSaveSpreadColumnInf()  ======================================
' Name : PopSaveSpreadColumnInf()
' Description : 이동한 컬럼의 정보를 저장 
'====================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub
'===================================== PopRestoreSpreadColumnInf()  ======================================
' Name : PopRestoreSpreadColumnInf()
' Description : 컬럼의 순서정보를 복원함 
'====================================================================================================
Sub PopRestoreSpreadColumnInf()
	Dim Row
    ggoSpread.Source = gActiveSpdSheet
	On Error Resume Next
	Err.Clear 	
	
	Select Case UCase(Trim(gActiveSpdSheet.Name))
		Case "VSPDDATA"
			Call PrevspdDataRestore(gActiveSpdSheet)
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("C")
			Call InitCombo()
			Call ggoSpread.ReOrderingSpreadData()
			Call InitData()
			Call SetSpreadColor(-1, -1)
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,-1,-1,C_DocCur,C_ItemAmt,"A" ,"I","X","X")    
			
			IF lgIntFlgMode = parent.OPMD_UMODE Then
				ggoSpread.source = frm1.vspdData       
				frm1.vspdData.ReDraw = False         
				ggoSpread.SpreadLock	C_AcctCd       , -1    , C_AcctCd
				frm1.vspdData.ReDraw = True
			End if
			
		Case "VSPDDATA2"
			Call PrevspdData2Restore(gActiveSpdSheet)   
			Call ggoSpread.RestoreSpreadInf()
			Call InitCtrlSpread()			'관리항목 그리드 초기화 
			Call ggoSpread.ReOrderingSpreadData()
			Call SetSpread2Color()  		
		
		Case "VSPDDATA1"
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("B")
			Call ggoSpread.ReOrderingSpreadData()
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData1,-1,-1,C_DocCur1,C_APBALAMT,   "A" ,"I","X","X")
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData1,-1,-1,C_DocCur1,C_APCLSAMT,   "A" ,"I","X","X")
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData1,-1,-1,C_DocCur1,C_APAMT,		"A" ,"I","X","X")
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData1,-1,-1,C_DocCur1,C_APDCAMT,		"A" ,"I","X","X")
		Case "VSPDDATA4"
			Call PrevspdData5Restore(gActiveSpdSheet)
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("A")
			Call ggoSpread.ReOrderingSpreadData()
			Call SetSpreadColor(-1, -1)
			Call SetSpreadLock("A")
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData4,-1,-1,C_DocCur4,C_BALAMT,"A" ,"I","X","X") 			
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData4,-1,-1,C_DocCur4,C_ALLCAMT,"A" ,"I","X","X")
			
			IF lgIntFlgMode = parent.OPMD_UMODE Then
			
				For Row= 0 to frm1.vspdData4.MaxRows
					ggoSpread.source = frm1.vspdData4    
					frm1.vspdData4.row= Row
					frm1.vspdData4.Col =C_AllCTYPE
					if UCase(frm1.vspdData4.text)= "CP" or  UCase(frm1.vspdData4.text)= "NP"  then
						ggoSpread.SSSetProtected C_ALLCAMT,    Row, Row
						ggoSpread.SSSetRequired C_AllcAcct,    Row, Row	   
					ELse
						ggoSpread.SSSetProtected C_AllcAcct,   Row, Row	    
						ggoSpread.SpreadUnLock	C_AllcAcct_POP,Row, Row
					END if
				Next 
			End if

 	
	End Select
	
	Call DoSum()
	
	If frm1.vspdData2.MaxRows <= 0 Then
		Call DbQuery2(frm1.vspdData.ActiveRow)
	End If
	
End Sub

'===================================== PrevspdDataRestore()  ========================================
' Name : PrevspdDataRestore()
' Description : 그리드 복원시 관리항목 복원 
'====================================================================================================
Sub PrevspdDataRestore(pActiveSheetName)
	Dim indx, indx1

	For indx = 0 To frm1.vspdData.MaxRows
        frm1.vspdData.Row    = indx
        frm1.vspdData.Col    = 0
		
		If frm1.vspdData.Text <> "" Then
			Select Case frm1.vspdData.Text			
				Case ggoSpread.InsertFlag					
					frm1.vspdData.Col = C_ItemSeq					
					Call DeleteHsheet(frm1.vspdData.Text)					
				Case ggoSpread.UpdateFlag		
					For indx1 = 0 To frm1.vspdData3.MaxRows					
						frm1.vspdData3.Row = indx1
						frm1.vspdData3.Col = 0
						Select Case frm1.vspdData3.Text 
							Case ggoSpread.UpdateFlag
								frm1.vspdData.Col = C_ItemSeq
								frm1.vspdData3.Col = 1					
								If UCase(Trim(frm1.vspdData.Text)) = UCase(Trim(frm1.vspdData3.Text)) Then
									Call DeleteHsheet(frm1.vspdData.Text)										
									Call FncRestoreDbQuery2(indx, frm1.vspdData.ActiveRow, frm1.txtPaymNO.Value)
								End If
						End Select
					Next
				Case ggoSpread.DeleteFlag
					Call fncRestoreDbQuery2(indx, frm1.vspdData.ActiveRow, frm1.txtPaymNO.Value)
			End Select
		End If
	Next
	
	ggoSpread.Source = pActiveSheetName
End Sub

'===================================== PrevspdDataRestore()  ========================================
' Name : PrevspdData2Restore()
' Description : 그리드 복원시 관리항목 복원 
'====================================================================================================
Sub PrevspdData2Restore(pActiveSheetName)
	Dim indx, indx1

	For indx = 0 to frm1.vspdData2.MaxRows
        frm1.vspdData2.Row    = indx
        frm1.vspdData2.Col    = 0

		If frm1.vspdData2.Text <> "" Then
			Select Case frm1.vspdData2.Text
				Case ggoSpread.InsertFlag
					frm1.vspdData2.Col = C_HItemSeq
					For indx1 = 0 to frm1.vspdData.MaxRows
						frm1.vspdData.Row = indx1
						frm1.vspdData.Col = C_ItemSeq
						If frm1.vspdData.Text = frm1.vspdData2.Text Then
							Call DeleteHsheet(frm1.vspdData.Text)
							ggoSpread.Source = frm1.vspdData	
					        ggoSpread.EditUndo							
						End If
					Next
				Case ggoSpread.UpdateFlag
					frm1.vspdData2.Col = C_HItemSeq
					For indx1 = 0 to frm1.vspdData.MaxRows
						frm1.vspdData.Row = indx1
						frm1.vspdData.Col = C_ItemSeq
						If frm1.vspdData.Text = frm1.vspdData2.Text Then
							Call DeleteHsheet(frm1.vspdData.Text)
							ggoSpread.Source = frm1.vspdData
							ggoSpread.EditUndo
							Call fncRestoreDbQuery2(indx1, frm1.vspdData.ActiveRow, frm1.txtPaymNO.Value) 
						End If
					Next
				Case ggoSpread.DeleteFlag

			End Select
		End If
	Next
	ggoSpread.Source = pActiveSheetName
End Sub
'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc"  -->	
</HEAD>

<!--'======================================================================================================
'       					6. Tag부 
'	기능: Tag부분 설정 
'======================================================================================================= -->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../image/table/seltab_up_bg.gif"><IMG height=23 src="../../image/table/seltab_up_left.gif" width=9></td>
								<td background="../../image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>반제등록</font></td>
								<td background="../../image/table/seltab_up_bg.gif" align="right"><IMG height=23 src="../../image/table/seltab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../image/table/tab_up_bg.gif"><IMG height=23 src="../../image/table/tab_up_left.gif" width=9></td>
								<td background="../../image/table/tab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>매입채무반제</font></td>
								<td background="../../image/table/tab_up_bg.gif" align="right"><IMG height=23 src="../../image/table/tab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>	
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab3()">
							<TR>
								<td background="../../image/table/tab_up_bg.gif"><IMG height=23 src="../../image/table/tab_up_left.gif" width=9></td>
								<td background="../../image/table/tab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>잔액처리</font></td>
								<td background="../../image/table/tab_up_bg.gif" align="right"><IMG height=23 src="../../image/table/tab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>	
					
					<TD WIDTH=* align=right><span id="spnArInfo"><a href="vbscript:OpenRefOpenAr()">채무발생정보</A>&nbsp;|&nbsp;</span><A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</A></TD>								
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS="Tab11">		
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR HEIGHT=20>
									<TD CLASS="TD5" NOWRAP>출금번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtPaymNO" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="13NXXU" ALT="출금번호"><IMG SRC="../../image/btnPopup.gif" NAME="btnRcptCD" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(0)"></TD>								
									<TD CLASS="TD6" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				
				<TR HEIGHT=120>
				   <TD WIDTH="100%">
				       <TABLE <%=LR_SPACE_TYPE_20%>>								    										
							<TR>
								<TD CLASS=TD5 NOWRAP>출금일</TD>
								<TD CLASS=TD6 NOWRAP>
								          <script language =javascript src='./js/a4125ma1_fpDateTime1_txtPaymDT.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>지급처</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="21NXXU" ALT="지급처"><IMG SRC="../../image/btnPopup.gif" NAME="btnBizCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenBp(frm1.txtBpCd.Value, 1)"> <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24" ALT="수금처명"></TD>								
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDeptCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag=22NXXU" ALT="부서"><IMG SRC="../../image/btnPopup.gif" NAME="btnDeptCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(frm1.txtDeptCd.value)"> <INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="24" ALT="부서명"></TD>								
								<TD class=TD5 NOWRAP>차이금액(자국)</TD>
								<TD class=TD6 NOWRAP>
									      <script language =javascript src='./js/a4125ma1_OBJECT22_txtTotLocAmt.js'></script></TD>
							<TR>
								<TD CLASS=TD5 NOWRAP>채무반제합계(자국)</TD>
								<TD CLASS=TD6 NOWRAP>
								          <script language =javascript src='./js/a4125ma1_OBJECT3_txtAPClsLocAmt.js'></script></TD>								
								<TD CLASS=TD5 NOWRAP>반제합계(자국)</TD>
								<TD CLASS=TD6 NOWRAP>
								          <script language =javascript src='./js/a4125ma1_OBJECT2_txtAllcLocAmt.js'></script></TD>							          								          

							</TR>		
							<TR>
								<TD class=TD5 NOWRAP>차변합계(자국)</TD>
								<TD class=TD6 NOWRAP><script language =javascript src='./js/a4125ma1_OBJECT22_txtDrLocAmt.js'></script></TD></TD>
								<TD class=TD5 NOWRAP>대변합계(자국)</TD>
								<TD class=TD6 NOWRAP><script language =javascript src='./js/a4125ma1_OBJECT22_txtCrLocAmt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>결의전표번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTempGLNo" SIZE=19 MAXLENGTH=18 tag="24" ALT="전표번호"></TD>
								<TD CLASS="TD5" NOWFRAP>전표번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtGLNo" SIZE=19 MAXLENGTH=18 tag="24" ALT="전표번호"></TD>
							</TR>        
					</TABLE>	
					</TD>		
				</TR>

				<TR>		
					<TD WIDTH="100%">
					


					<DIV ID="TabDiv" SCROLL=no>
						<TABLE <%=LR_SPACE_TYPE_20%>>		
						    <TR HEIGHT="100%">
								<TD WIDTH="100%" COLSPAN="4">
									<script language =javascript src='./js/a4125ma1_OBJECT4_vspdData4.js'></script>
								</TD>											
							</TR>												
						</TABLE>		
					</DIV>
					


					<DIV ID="TabDiv"  SCROLL="no">
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR HEIGHT="100%">
								<TD WIDTH="100%" COLSPAN="4">
									<script language =javascript src='./js/a4125ma1_OBJECT5_vspdData1.js'></script>
								</TD>											
							</TR>						
							
						</TABLE>						
					</DIV>
					


					<DIV ID="TabDiv" SCROLL=no>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR HEIGHT="60%">
								<TD WIDTH="100%" COLSPAN="4">
								<script language =javascript src='./js/a4125ma1_OBJECT6_vspdData.js'></script></TD>
							</TR>
							<TR HEIGHT="40%">
								<TD WIDTH="100%" COLSPAN="4">
								<script language =javascript src='./js/a4125ma1_OBJECT7_vspdData2.js'></script></TD>
							</TR>
							
						</TABLE>		
					</DIV>
					
					</TD>
				</TR>									

			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 tabindex=-1></IFRAME>
		</TD>		
	</TR>
</TABLE>

<TEXTAREA Class=hidden name="txtSpread"  tag="24" tabindex=-1></TEXTAREA>
<TEXTAREA Class=hidden name="txtSpread1" tag="24" tabindex=-1></TEXTAREA>
<TEXTAREA Class=hidden name="txtSpread3" tag="24" tabindex=-1></TEXTAREA>
<TEXTAREA Class=hidden name="txtSpread4" tag="24" tabindex=-1></TEXTAREA>

<INPUT TYPE=hidden NAME="txtMaxRows"   tag="24" tabindex=-1>
<INPUT TYPE=hidden NAME="txtMaxRows1"  tag="24" tabindex=-1>
<INPUT TYPE=hidden NAME="txtMaxRows3"  tag="24" tabindex=-1>
<INPUT TYPE=hidden NAME="txtMaxRows4"  tag="24" tabindex=-1>

<INPUT TYPE=hidden NAME="txtMode"      tag="24" tabindex=-1>
<INPUT TYPE=hidden NAME="hItemSeq"     tag="24" tabindex=-1>
<INPUT TYPE=hidden NAME="txtFlgMode"   tag="24" tabindex=-1>


<!-- 부서코드입력시 기본정보  -->

<INPUT TYPE=hidden NAME="horgChangeId" tag="24">


<script language =javascript src='./js/a4125ma1_OBJECT8_vspdData3.js'></script>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
